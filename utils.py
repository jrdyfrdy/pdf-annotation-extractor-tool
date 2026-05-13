import re
from datetime import datetime, timezone, timedelta
from typing import Optional, Iterable

import fitz  # PyMuPDF

PDF_DATE_RE = re.compile(
    r"^D:(?P<year>\d{4})"
    r"(?P<month>\d{2})?"
    r"(?P<day>\d{2})?"
    r"(?P<hour>\d{2})?"
    r"(?P<minute>\d{2})?"
    r"(?P<second>\d{2})?"
    r"(?P<offset_sign>[Z+-])?"
    r"(?P<offset_hour>\d{2})?'?"
    r"(?P<offset_minute>\d{2})?'?$"
)


def parse_pdf_date(raw: str) -> str:
    if not raw:
        return ""
    match = PDF_DATE_RE.match(raw)
    if not match:
        return raw
    parts = match.groupdict()
    try:
        year = int(parts["year"])
        month = int(parts["month"] or 1)
        day = int(parts["day"] or 1)
        hour = int(parts["hour"] or 0)
        minute = int(parts["minute"] or 0)
        second = int(parts["second"] or 0)
        tzinfo = None
        sign = parts["offset_sign"]
        if sign == "Z":
            tzinfo = timezone.utc
        elif sign in ("+", "-") and parts["offset_hour"]:
            offset_hours = int(parts["offset_hour"] or 0)
            offset_minutes = int(parts["offset_minute"] or 0)
            delta = timedelta(hours=offset_hours, minutes=offset_minutes)
            if sign == "-":
                delta = -delta
            tzinfo = timezone(delta)
        dt = datetime(year, month, day, hour, minute, second, tzinfo=tzinfo)
        return dt.isoformat()
    except ValueError:
        return raw


def rgb_to_hex(color_tuple: Optional[Iterable[float]]) -> str:
    if not color_tuple:
        return ""
    try:
        r, g, b = color_tuple
        return "#{:02X}{:02X}{:02X}".format(
            int(max(0, min(1, r)) * 255),
            int(max(0, min(1, g)) * 255),
            int(max(0, min(1, b)) * 255),
        )
    except (ValueError, TypeError):
        return ""


def get_annot_color(annot: fitz.Annot) -> str:
    colors = annot.colors or {}
    stroke = colors.get("stroke")
    fill = colors.get("fill")
    return rgb_to_hex(stroke or fill)


def extract_highlighted_text(page: fitz.Page, annot: fitz.Annot) -> str:
    vertices = annot.vertices
    if not vertices:
        return ""
    words = []
    for quad in _iter_quads(vertices):
        clip = quad.rect
        text = page.get_text("text", clip=clip).strip()
        if text:
            words.append(text)
    return " ".join(words).strip()


def _iter_quads(vertices) -> list[fitz.Quad]:
    quads = []
    if not vertices:
        return quads
    for i in range(0, len(vertices), 4):
        quad_points = vertices[i : i + 4]
        if len(quad_points) < 4:
            continue
        quads.append(fitz.Quad(quad_points))
    return quads
