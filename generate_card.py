#!/usr/bin/env python3
"""Create a 750x1000 marketplace card from WDPYD and a 1C XLS export."""

from __future__ import annotations

import argparse
import io
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, urljoin
from urllib.request import Request, urlopen

import xlrd
from bs4 import BeautifulSoup
from PIL import Image, ImageDraw, ImageFilter, ImageFont


BASE_URL = "http://wdpyd.com"
SEARCH_PATH = "/en/WebSite/ProductList/2.html"
CARD_SIZE = (750, 1000)
GREEN = "#75bd22"
DARK_GREEN = "#4f9519"


@dataclass(frozen=True)
class Application:
    make: str
    summary: str
    details: tuple[str, ...]

    @property
    def model(self) -> str:
        value = re.sub(r"\s+\d{4}(?:-\d{4})?\s*$", "", self.summary).strip()
        return value[len(self.make) :].strip()


@dataclass(frozen=True)
class Product:
    article: str
    image_url: str
    applications: tuple[Application, ...]


def fetch(url: str) -> bytes:
    request = Request(url, headers={"User-Agent": "Mozilla/5.0 AutopartsCardGenerator/1.0"})
    try:
        with urlopen(request, timeout=30) as response:
            return response.read()
    except (HTTPError, URLError) as exc:
        raise RuntimeError(f"Не удалось загрузить {url}: {exc}") from exc


def find_product_url(article: str) -> str:
    url = f"{BASE_URL}{SEARCH_PATH}?{urlencode({'cphm': article})}"
    soup = BeautifulSoup(fetch(url), "html.parser")
    expected = re.compile(rf"WDPYD\s*NO\.\s*:\s*{re.escape(article)}\b", re.I)
    for heading in soup.find_all(["h1", "h2", "h3"]):
        if expected.search(heading.get_text(" ", strip=True)):
            anchor = heading.find_parent("a")
            if anchor and anchor.get("href"):
                return urljoin(BASE_URL, anchor["href"])
    raise RuntimeError(f"Артикул {article} не найден в каталоге WDPYD")


def parse_product(article: str) -> Product:
    detail_url = find_product_url(article)
    soup = BeautifulSoup(fetch(detail_url), "html.parser")
    image = soup.select_one('img[src^="/cptp/"]')
    if image is None:
        raise RuntimeError(f"Для артикула {article} не найдено изображение")

    makes: list[str] = []
    for strong in soup.find_all("strong"):
        if strong.get_text(" ", strip=True).upper().startswith("MAKE"):
            cell = strong.find_parent("td")
            value_cell = cell.find_next_sibling("td") if cell else None
            if value_cell:
                makes = [" ".join(value.split()) for value in value_cell.stripped_strings]
            break

    applications: list[Application] = []
    accordion = soup.select_one(".accordion")
    if accordion:
        for heading in accordion.find_all("h3", recursive=False):
            summary = " ".join(heading.get_text(" ", strip=True).split())
            paragraph = heading.find_next_sibling("p")
            details = ()
            if paragraph:
                details = tuple(
                    " ".join(text.split())
                    for text in paragraph.stripped_strings
                    if text.strip()
                )
            make = next(
                (candidate for candidate in sorted(makes, key=len, reverse=True)
                 if summary.upper().startswith(candidate.upper() + " ")),
                summary.split()[0] if summary else "",
            )
            applications.append(Application(make, summary, details))

    return Product(article, urljoin(BASE_URL, image["src"]), tuple(applications))


def normalize_article(value: object) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def lookup_name(xls_path: Path, article: str) -> str:
    workbook = xlrd.open_workbook(str(xls_path))
    for sheet in workbook.sheets():
        for row in range(1, sheet.nrows):
            if normalize_article(sheet.cell_value(row, 0)).casefold() == article.casefold():
                return " ".join(str(sheet.cell_value(row, 2)).split())
    raise RuntimeError(f"Артикул {article} не найден в XLS {xls_path}")


def unique(items: list[str]) -> list[str]:
    return list(dict.fromkeys(item for item in items if item))


def applicability_labels(applications: tuple[Application, ...]) -> list[str]:
    makes = unique([app.make for app in applications])
    if len(makes) >= 6:
        return makes
    if len(makes) >= 3:
        return unique([f"{app.make} {app.model}".strip() for app in applications])
    labels: list[str] = []
    for app in applications:
        labels.extend(app.details or (app.summary,))
    return unique(labels)


def font(size: int, bold: bool = True, italic: bool = False) -> ImageFont.FreeTypeFont:
    if bold and italic:
        path = "/usr/share/fonts/truetype/liberation2/LiberationSans-BoldItalic.ttf"
    elif bold:
        path = "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf"
    else:
        path = "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf"
    return ImageFont.truetype(path, size)


def draw_green_band(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int]) -> None:
    draw.rectangle(box, fill=GREEN)
    x0, y0, x1, y1 = box
    for x in range(x0 - (y1 - y0), x1 + 20, 16):
        draw.line((x, y1, x + (y1 - y0), y0), fill=DARK_GREEN, width=3)


def fit_line(draw: ImageDraw.ImageDraw, text: str, max_width: int, max_size: int, min_size: int = 14,
             italic: bool = True) -> ImageFont.FreeTypeFont:
    for size in range(max_size, min_size - 1, -1):
        candidate = font(size, bold=True, italic=italic)
        if draw.textbbox((0, 0), text, font=candidate)[2] <= max_width:
            return candidate
    return font(min_size, bold=True, italic=italic)


def split_balanced(labels: list[str]) -> tuple[str, str]:
    if not labels:
        return "", ""
    best = ("", " • ".join(labels))
    best_delta = len(best[1])
    for index in range(1, len(labels)):
        top, bottom = " • ".join(labels[:index]), " • ".join(labels[index:])
        delta = abs(len(top) - len(bottom))
        if delta < best_delta:
            best, best_delta = (top, bottom), delta
    return best


def contain(image: Image.Image, size: tuple[int, int]) -> Image.Image:
    image = image.convert("RGBA")
    image.thumbnail(size, Image.Resampling.LANCZOS)
    return image


def render_card(product: Product, name: str, image_bytes: bytes, template_path: Path | None) -> Image.Image:
    card = Image.new("RGB", CARD_SIZE, "white")
    draw = ImageDraw.Draw(card)
    draw_green_band(draw, (0, 0, 750, 78))
    draw_green_band(draw, (0, 920, 750, 999))

    # Preserve the recognisable logo from the supplied reference template.
    if template_path and template_path.exists():
        template = Image.open(template_path).convert("RGBA")
        logo = template.crop((0, 0, 76, 78))
        card.paste(logo.convert("RGB"), (0, 0))
    else:
        draw.text((12, 22), "WDPYD", font=font(17), fill="white")

    labels = applicability_labels(product.applications)
    top_text, bottom_text = split_balanced(labels)
    for text, y in ((top_text, 39), (bottom_text, 960)):
        if text:
            face = fit_line(draw, text, 650 if y == 39 else 700, 35, 15)
            area_left = 88 if y == 39 else 25
            area_width = 640 if y == 39 else 700
            bbox = draw.textbbox((0, 0), text, font=face)
            x = area_left + (area_width - (bbox[2] - bbox[0])) // 2
            draw.text((x, y), text, font=face, fill="white", anchor="lm")

    # Article badge.
    draw.rounded_rectangle((-8, 80, 161, 190), radius=14, fill="#f7fff0", outline="#83bd37", width=3)
    article_font = fit_line(draw, product.article, 135, 42, 24, italic=False)
    draw.text((76, 135), product.article, font=article_font, fill="black", anchor="mm")

    # Product name: a compact black header plus a green continuation, like the reference.
    words = name.upper().split()
    first = words[0] if words else name.upper()
    rest = " ".join(words[1:])
    first_font = fit_line(draw, first, 390, 43, 25)
    first_width = draw.textbbox((0, 0), first, font=first_font)[2]
    box_left = max(280, 725 - first_width - 38)
    draw.rounded_rectangle((box_left, 86, 725, 176), radius=14, fill="black")
    draw.text(((box_left + 725) // 2, 131), first, font=first_font, fill="white", anchor="mm")
    if rest:
        rest_font = fit_line(draw, rest, 690, 38, 20, italic=False)
        draw.text((375, 218), rest, font=rest_font, fill="#2fa143", anchor="mm")

    part = contain(Image.open(io.BytesIO(image_bytes)), (620, 570))
    px = (750 - part.width) // 2
    py = 285 + (540 - part.height) // 2
    card.paste(part, (px, py), part)

    # Soft shadow below the product.
    shadow = Image.new("RGBA", (440, 55), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow)
    sd.ellipse((35, 15, 405, 42), fill=(0, 0, 0, 140))
    shadow = shadow.filter(ImageFilter.GaussianBlur(13))
    card.paste(shadow, (155, 842), shadow)
    return card


def build(article: str, price: Path, template: Path | None, output_dir: Path) -> Path:
    product = parse_product(article)
    name = lookup_name(price, article)
    image_bytes = fetch(product.image_url)
    output_dir.mkdir(parents=True, exist_ok=True)
    destination = output_dir / f"{article}.jpg"
    render_card(product, name, image_bytes, template).save(destination, "JPEG", quality=95, subsampling=0)
    return destination


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("article", help="артикул WDPYD")
    parser.add_argument("--price", type=Path, required=True, help="XLS-выгрузка из 1С")
    parser.add_argument("--template", type=Path, help="PNG-шаблон/референс с логотипом")
    parser.add_argument("--output-dir", type=Path, default=Path("output"))
    args = parser.parse_args()
    try:
        result = build(args.article.strip(), args.price, args.template, args.output_dir)
    except (RuntimeError, OSError, xlrd.XLRDError) as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1
    print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
