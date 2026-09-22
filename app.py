#!/usr/bin/env python3
"""Web interface for batch generation of Ozon product cards."""

from __future__ import annotations

import base64
import hmac
import os
import re
import threading
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import BinaryIO

import xlrd
from flask import Flask, abort, jsonify, render_template, request, send_file, send_from_directory, url_for
from openpyxl import load_workbook
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.utils import secure_filename

from generate_card import build_with_name, normalize_article


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("AUTOPARTS_DATA_DIR", BASE_DIR / "data"))
TEMPLATE_PATH = Path(os.environ.get("AUTOPARTS_TEMPLATE", BASE_DIR / "assets" / "Frame 6.png"))
ALLOWED_EXTENSIONS = {".xls", ".xlsx"}
ARTICLE_FILENAME = re.compile(r"[^0-9A-Za-zА-Яа-я._-]+")

app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1)
app.config["MAX_CONTENT_LENGTH"] = 15 * 1024 * 1024
executor = ThreadPoolExecutor(max_workers=int(os.environ.get("AUTOPARTS_WORKERS", "2")))
jobs_lock = threading.Lock()


@dataclass
class Row:
    article: str
    name: str
    status: str = "waiting"
    error: str | None = None
    filename: str | None = None

    def json(self, job_id: str) -> dict[str, str | None]:
        return {
            "article": self.article,
            "name": self.name,
            "status": self.status,
            "error": self.error,
            "file_url": url_for("job_file", job_id=job_id, filename=self.filename)
            if self.filename else None,
        }


@dataclass
class Job:
    id: str
    rows: list[Row]
    output_dir: Path
    created_at: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    status: str = "waiting"

    def json(self) -> dict:
        complete = sum(row.status in {"done", "error"} for row in self.rows)
        return {
            "id": self.id,
            "status": self.status,
            "complete": complete,
            "total": len(self.rows),
            "rows": [row.json(self.id) for row in self.rows],
            "folder_url": url_for("job_files", job_id=self.id),
            "download_url": url_for("download_job", job_id=self.id),
            "status_url": url_for("job_status", job_id=self.id),
            "local_path": str(self.output_dir.resolve()),
        }


jobs: dict[str, Job] = {}


def unauthorized_response():
    return (
        "Требуется авторизация\n",
        401,
        {"WWW-Authenticate": 'Basic realm="Генератор карточек Ozon", charset="UTF-8"'},
    )


@app.before_request
def require_separate_auth():
    """Protect this application with credentials independent of rialsat-admin."""
    if os.environ.get("AUTOPARTS_REQUIRE_AUTH", "0") != "1":
        return None
    expected_login = os.environ.get("AUTOPARTS_LOGIN", "")
    expected_password = os.environ.get("AUTOPARTS_PASSWORD", "")
    if not expected_login or not expected_password:
        return "Авторизация приложения не настроена\n", 503
    header = request.headers.get("Authorization", "")
    scheme, _, token = header.partition(" ")
    if scheme.casefold() != "basic" or not token:
        return unauthorized_response()
    try:
        decoded = base64.b64decode(token, validate=True).decode("utf-8")
    except (ValueError, UnicodeDecodeError):
        return unauthorized_response()
    login, separator, password = decoded.partition(":")
    if not separator:
        return unauthorized_response()
    login_ok = hmac.compare_digest(login, expected_login)
    password_ok = hmac.compare_digest(password, expected_password)
    if not (login_ok and password_ok):
        return unauthorized_response()
    return None


def normalized_header(value: object) -> str:
    return " ".join(str(value or "").strip().casefold().split())


def rows_from_xls(stream: BinaryIO) -> list[tuple[str, str]]:
    book = xlrd.open_workbook(file_contents=stream.read())
    sheet = book.sheet_by_index(0)
    values = [[sheet.cell_value(row, col) for col in range(sheet.ncols)] for row in range(sheet.nrows)]
    return parse_table(values)


def rows_from_xlsx(stream: BinaryIO) -> list[tuple[str, str]]:
    book = load_workbook(stream, read_only=True, data_only=True)
    sheet = book.active
    return parse_table([list(row) for row in sheet.iter_rows(values_only=True)])


def parse_table(values: list[list[object]]) -> list[tuple[str, str]]:
    if not values:
        raise ValueError("Файл пуст")
    headers = [normalized_header(value) for value in values[0]]
    missing = [name for name in ("артикул", "наименование") if name not in headers]
    if missing:
        raise ValueError("Нет обязательных столбцов: " + ", ".join(missing))
    article_col, name_col = headers.index("артикул"), headers.index("наименование")
    result: list[tuple[str, str]] = []
    for source_row in values[1:]:
        article = normalize_article(source_row[article_col] if article_col < len(source_row) else "")
        name = " ".join(str(source_row[name_col] if name_col < len(source_row) else "").split())
        if article and name:
            result.append((article, name))
    if not result:
        raise ValueError("В файле нет строк с заполненными артикулом и наименованием")
    return result


def safe_article_filename(article: str) -> str:
    cleaned = ARTICLE_FILENAME.sub("_", article).strip("._")
    return cleaned or "item"


def process_job(job_id: str) -> None:
    job = jobs[job_id]
    with jobs_lock:
        job.status = "processing"
    for row in job.rows:
        with jobs_lock:
            row.status = "processing"
        try:
            file_article = safe_article_filename(row.article)
            result = build_with_name(
                row.article, row.name, TEMPLATE_PATH, job.output_dir,
                output_stem=file_article,
            )
            expected = job.output_dir / f"{file_article}.jpg"
            if result != expected:
                result.replace(expected)
            with jobs_lock:
                row.filename = expected.name
                row.status = "done"
        except Exception as exc:  # one failed article must not stop the batch
            with jobs_lock:
                row.status = "error"
                row.error = str(exc)
    with jobs_lock:
        job.status = "done"


def get_job(job_id: str) -> Job:
    job = jobs.get(job_id)
    if job is None:
        abort(404)
    return job


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/api/jobs")
def create_job():
    upload = request.files.get("file")
    if upload is None or not upload.filename:
        return jsonify(error="Выберите Excel-файл"), 400
    suffix = Path(secure_filename(upload.filename)).suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        return jsonify(error="Поддерживаются только файлы .xls и .xlsx"), 400
    try:
        parsed = rows_from_xls(upload.stream) if suffix == ".xls" else rows_from_xlsx(upload.stream)
    except (ValueError, xlrd.XLRDError, OSError) as exc:
        return jsonify(error=f"Не удалось прочитать файл: {exc}"), 400

    job_id = uuid.uuid4().hex[:12]
    output_dir = DATA_DIR / "jobs" / job_id / "output"
    job = Job(job_id, [Row(article, name) for article, name in parsed], output_dir)
    with jobs_lock:
        jobs[job_id] = job
    executor.submit(process_job, job_id)
    return jsonify(job.json()), 202


@app.get("/api/jobs/<job_id>")
def job_status(job_id: str):
    with jobs_lock:
        payload = get_job(job_id).json()
    return jsonify(payload)


@app.get("/jobs/<job_id>/files")
def job_files(job_id: str):
    job = get_job(job_id)
    files = sorted(job.output_dir.glob("*.jpg")) if job.output_dir.exists() else []
    return render_template("files.html", job=job, files=files)


@app.get("/jobs/<job_id>/files/<path:filename>")
def job_file(job_id: str, filename: str):
    job = get_job(job_id)
    return send_from_directory(job.output_dir, filename, as_attachment=True)


@app.get("/jobs/<job_id>/download")
def download_job(job_id: str):
    job = get_job(job_id)
    job.output_dir.mkdir(parents=True, exist_ok=True)
    archive = job.output_dir.parent / f"cards-{job_id}.zip"
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as bundle:
        for image in sorted(job.output_dir.glob("*.jpg")):
            bundle.write(image, image.name)
    return send_file(archive, as_attachment=True, download_name=archive.name)


@app.get("/health")
def health():
    return {"status": "ok", "template": TEMPLATE_PATH.exists()}


if __name__ == "__main__":
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")), debug=False)
