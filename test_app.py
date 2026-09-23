import io
import json
import os
from pathlib import Path
import tempfile
import unittest

import app as webapp
from app import app, parse_table


class SpreadsheetValidationTest(unittest.TestCase):
    def tearDown(self):
        for key in ("AUTOPARTS_REQUIRE_AUTH", "AUTOPARTS_LOGIN", "AUTOPARTS_PASSWORD"):
            os.environ.pop(key, None)

    def test_columns_are_case_and_space_insensitive(self):
        rows = parse_table([
            [" Артикул ", "НАИМЕНОВАНИЕ"],
            [636116.0, " Ремкомплект   суппорта "],
        ])
        self.assertEqual(rows, [("636116", "Ремкомплект суппорта")])

    def test_missing_column_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "наименование"):
            parse_table([["артикул", "цена"], ["636116", "100"]])

    def test_non_excel_upload_is_rejected(self):
        client = app.test_client()
        response = client.post(
            "/api/jobs",
            data={"file": (io.BytesIO(bytes("артикул", "utf-8")), "parts.csv")},
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn(".xls", response.get_json()["error"])

    def test_separate_basic_auth(self):
        os.environ.update(
            AUTOPARTS_REQUIRE_AUTH="1",
            AUTOPARTS_LOGIN="parts-user",
            AUTOPARTS_PASSWORD="different-secret",
        )
        client = app.test_client()
        self.assertEqual(client.get("/").status_code, 401)
        token = __import__("base64").b64encode(b"parts-user:different-secret").decode()
        self.assertEqual(client.get("/", headers={"Authorization": f"Basic {token}"}).status_code, 200)

    def test_archives_are_grouped_by_creation_date(self):
        previous = webapp.DATA_DIR
        try:
            with tempfile.TemporaryDirectory() as temporary:
                webapp.DATA_DIR = Path(temporary)
                job_dir = webapp.DATA_DIR / "jobs" / "abcdef123456"
                output = job_dir / "output"
                output.mkdir(parents=True)
                (output / "636116.jpg").write_bytes(b"jpg")
                (output / "636116.json").write_text("{}", encoding="utf-8")
                (job_dir / "job.json").write_text(json.dumps({
                    "id": "abcdef123456",
                    "created_at": "2026-09-23T10:15:00+00:00",
                    "source_filename": "parts.xlsx",
                    "status": "done",
                }), encoding="utf-8")
                groups = webapp.archive_groups()
                self.assertEqual(groups[0]["label"], "23.09.2026")
                self.assertEqual(groups[0]["items"][0]["file_count"], 2)
                self.assertEqual(groups[0]["items"][0]["card_count"], 1)
                response = app.test_client().get("/archives")
                self.assertEqual(response.status_code, 200)
                self.assertIn("23.09.2026", response.get_data(as_text=True))
        finally:
            webapp.DATA_DIR = previous


if __name__ == "__main__":
    unittest.main()
