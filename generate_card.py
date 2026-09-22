#!/usr/bin/env python3
"""Create a 750x1000 marketplace card from WDPYD and a 1C XLS export."""

from __future__ import annotations

import argparse
import io
import json
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
    detail_url: str
    image_url: str
    category: str
    part: str
    oe: str
    makes: tuple[str, ...]
    description: str
    brake_system: str
    fitting_position: str
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

    def table_value(label: str) -> tuple[str, ...]:
        expected = label.rstrip(":").strip().casefold()
        for marker in soup.find_all(["strong", "span"]):
            actual = marker.get_text(" ", strip=True).rstrip(":").strip().casefold()
            if actual == expected:
                cell = marker.find_parent("td")
                value_cell = cell.find_next_sibling("td") if cell else None
                if value_cell:
                    return tuple(" ".join(value.split()) for value in value_cell.stripped_strings)
        return ()

    makes = list(table_value("Make"))

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

    category_node = soup.select_one("#oenumber")
    category = " ".join(category_node.get_text(" ", strip=True).split()) if category_node else ""
    return Product(
        article=article,
        detail_url=detail_url,
        image_url=urljoin(BASE_URL, image["src"]),
        category=category,
        part=" ".join(table_value("Part")),
        oe=" ".join(table_value("OE")),
        makes=tuple(makes),
        description=" ".join(table_value("Description")),
        brake_system=" ".join(table_value("BRAKE SYSTEM")),
        fitting_position=" ".join(table_value("FITTING POSITION")),
        applications=tuple(applications),
    )


def product_to_dict(product: Product, name: str) -> dict:
    """Return all catalogue data in a stable, serializable schema."""
    return {
        "schema_version": 1,
        "article": product.article,
        "name": name,
        "source": {
            "catalog": "WDPYD",
            "product_url": product.detail_url,
            "image_url": product.image_url,
        },
        "catalog": {
            "category": product.category,
            "part": product.part,
            "oe": product.oe,
            "makes": list(product.makes),
            "description": product.description,
            "brake_system": product.brake_system,
            "fitting_position": product.fitting_position,
        },
        "applications": [
            {
                "make": application.make,
                "model": application.model,
                "summary": application.summary,
                "details": list(application.details),
            }
            for application in product.applications
        ],
    }


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
        make_models = unique([f"{app.make} {app.model}".strip() for app in applications])
        return makes if len(make_models) >= 6 else make_models
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


def draw_dot_block(draw: ImageDraw.ImageDraw, origin: tuple[int, int]) -> None:
    """Draw the 3x3 orientation motif from the reference layout."""
    start_x, start_y = origin
    for row in range(3):
        for column in range(3):
            x = start_x + column * 18
            y = start_y + row * 22
            draw.ellipse((x, y, x + 4, y + 7), fill="black")


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


def split_balanced_items(labels: list[str]) -> tuple[list[str], list[str]]:
    """Split complete labels into two groups with similar character lengths."""
    if len(labels) < 2:
        return labels, []
    best_index = 1
    best_delta = float("inf")
    for index in range(1, len(labels)):
        left = " • ".join(labels[:index])
        right = " • ".join(labels[index:])
        delta = abs(len(left) - len(right))
        if delta < best_delta:
            best_index, best_delta = index, delta
    return labels[:best_index], labels[best_index:]


def band_lines(labels: list[str], character_limit: int = 50) -> list[str]:
    """Use two lines when a band's content exceeds the requested limit."""
    text = " • ".join(labels)
    if len(text) <= character_limit:
        return [text] if text else []
    left, right = split_balanced_items(labels)
    if right:
        return [" • ".join(left), " • ".join(right)]
    # A single unusually long label still needs to remain fully visible.
    words = text.split()
    if len(words) < 2:
        midpoint = len(text) // 2
        return [text[:midpoint], text[midpoint:]]
    split_at = min(range(1, len(words)), key=lambda i: abs(len(" ".join(words[:i])) - len(" ".join(words[i:]))))
    return [" ".join(words[:split_at]), " ".join(words[split_at:])]


def wrap_by_width(draw: ImageDraw.ImageDraw, text: str, face: ImageFont.FreeTypeFont,
                  max_width: int) -> list[str]:
    lines: list[str] = []
    current = ""
    for word in text.split():
        candidate = f"{current} {word}".strip()
        if current and draw.textbbox((0, 0), candidate, font=face)[2] > max_width:
            lines.append(current)
            current = word
        else:
            current = candidate
    if current:
        lines.append(current)
    return lines


def fit_wrapped_text(draw: ImageDraw.ImageDraw, text: str, max_width: int, max_height: int,
                     max_size: int = 38, min_size: int = 12) -> tuple[ImageFont.FreeTypeFont, list[str], int]:
    """Fit all text into a rectangle, reducing font size before returning."""
    for size in range(max_size, min_size - 1, -1):
        face = font(size, bold=True, italic=False)
        lines = wrap_by_width(draw, text, face, max_width)
        spacing = max(3, size // 5)
        line_height = max(draw.textbbox((0, 0), line, font=face)[3] for line in lines)
        if line_height * len(lines) + spacing * (len(lines) - 1) <= max_height:
            return face, lines, line_height + spacing
    face = font(min_size, bold=True, italic=False)
    return face, wrap_by_width(draw, text, face, max_width), min_size + 3


def contain(image: Image.Image, size: tuple[int, int]) -> Image.Image:
    image = image.convert("RGBA")
    image.thumbnail(size, Image.Resampling.LANCZOS)
    return image


def extract_wheel(template: Image.Image) -> tuple[Image.Image, Image.Image]:
    """Return the wheel crop and an alpha mask without its white/green backdrop.

    The tyre touches the right edge of the reference.  For every scanline we
    detect its low-saturation grey/black pixels, discard unrelated elements
    that are not connected to the right-hand cluster, then retain the complete
    profile from that boundary to the canvas edge.
    """
    box = (620, 640, 750, 1000)
    wheel = template.crop(box).convert("RGB")
    hsv = wheel.convert("HSV")
    mask = Image.new("L", wheel.size, 0)
    mask_pixels = mask.load()
    hsv_pixels = hsv.load()

    for y in range(wheel.height):
        candidates = [
            x for x in range(wheel.width)
            if hsv_pixels[x, y][1] < 105 and hsv_pixels[x, y][2] < 225
        ]
        if not candidates or candidates[-1] < wheel.width - 8:
            continue

        # Work backwards from the rightmost tyre pixel. A sizeable gap marks
        # the end of the tyre and excludes unrelated objects to its left.
        cluster = [candidates[-1]]
        for x in reversed(candidates[:-1]):
            if cluster[-1] - x > 24:
                break
            cluster.append(x)
        boundary = max(0, min(cluster) - 2)
        for x in range(boundary, wheel.width):
            mask_pixels[x, y] = 255

    # A tiny blur preserves the antialiased outer edge without reintroducing
    # a visible rectangle of the original background.
    return wheel, mask.filter(ImageFilter.GaussianBlur(0.7))


def render_card(product: Product, name: str, image_bytes: bytes, template_path: Path | None) -> Image.Image:
    card = Image.new("RGB", CARD_SIZE, "white")
    draw = ImageDraw.Draw(card)
    draw_green_band(draw, (0, 0, 750, 78))
    draw_green_band(draw, (0, 920, 750, 999))

    # Preserve the recognisable logo from the supplied reference template.
    template = None
    if template_path and template_path.exists():
        template = Image.open(template_path).convert("RGBA")
        logo = template.crop((0, 0, 76, 78))
        card.paste(logo.convert("RGB"), (0, 0))
    else:
        draw.text((12, 22), "WDPYD", font=font(17), fill="white")

    labels = applicability_labels(product.applications)
    top_labels, bottom_labels = split_balanced_items(labels)
    for band_labels, center_y, area_left, area_width in (
        (top_labels, 39, 88, 640),
        (bottom_labels, 960, 25, 610),
    ):
        lines = band_lines(band_labels)
        if not lines:
            continue
        max_font_size = 32 if len(lines) == 1 else 23
        line_gap = 29 if len(lines) == 2 else 0
        for index, text in enumerate(lines):
            face = fit_line(draw, text, area_width, max_font_size, 11)
            bbox = draw.textbbox((0, 0), text, font=face)
            x = area_left + (area_width - (bbox[2] - bbox[0])) // 2
            y = center_y + (index - (len(lines) - 1) / 2) * line_gap
            draw.text((x, y), text, font=face, fill="white", anchor="lm")

    # Article badge.
    draw.rounded_rectangle((2, 81, 162, 190), radius=14, fill="#f7fff0", outline="#83bd37", width=3)
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
        rest_font, rest_lines, line_step = fit_wrapped_text(draw, rest, 690, 94, 38, 12)
        block_height = line_step * len(rest_lines)
        start_y = 192 + (94 - block_height) / 2 + line_step / 2
        for index, line in enumerate(rest_lines):
            draw.text((375, start_y + index * line_step), line, font=rest_font,
                      fill="#2fa143", anchor="mm")

    part = contain(Image.open(io.BytesIO(image_bytes)), (620, 570))
    px = (750 - part.width) // 2
    py = 285 + (540 - part.height) // 2
    card.paste(part, (px, py), part)

    # These markers belong to the foreground and must not be covered by a
    # wide source photograph.
    draw_dot_block(draw, (47, 292))
    draw_dot_block(draw, (683, 292))
    draw_dot_block(draw, (47, 780))

    # Soft shadow below the product.
    shadow = Image.new("RGBA", (440, 55), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow)
    sd.ellipse((35, 15, 405, 42), fill=(0, 0, 0, 140))
    shadow = shadow.filter(ImageFilter.GaussianBlur(13))
    card.paste(shadow, (155, 842), shadow)

    # The wheel is a semantic category marker, not just decoration.  It is
    # retained pixel-for-pixel from the supplied template and remains on top
    # of the bottom band, matching the reference composition.
    if template is not None and template.size == CARD_SIZE:
        wheel, wheel_mask = extract_wheel(template)
        card.paste(wheel, (620, 640), wheel_mask)
    return card


def build(article: str, price: Path, template: Path | None, output_dir: Path) -> Path:
    product = parse_product(article)
    name = lookup_name(price, article)
    return build_with_name(article, name, template, output_dir, product)


def build_with_name(article: str, name: str, template: Path | None, output_dir: Path,
                    product: Product | None = None, output_stem: str | None = None) -> Path:
    """Build a card when the name has already been read from an uploaded file."""
    product = product or parse_product(article)
    image_bytes = fetch(product.image_url)
    output_dir.mkdir(parents=True, exist_ok=True)
    destination = output_dir / f"{output_stem or article}.jpg"
    render_card(product, name, image_bytes, template).save(destination, "JPEG", quality=95, subsampling=0)
    json_destination = output_dir / f"{output_stem or article}.json"
    json_destination.write_text(
        json.dumps(product_to_dict(product, name), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
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
