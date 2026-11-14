#!/usr/bin/env python3
"""Lint a .pptx for readability problems.

Flags small type, crowded bullets, extra numbers on non-data slides, long
bullets, boxes outside the slide or overlapping, low contrast, grouped
shapes, Vietnamese text in a font that cannot draw it, and broken links.
"""

from __future__ import annotations

import argparse
import re
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation
from pptx.enum.dml import MSO_FILL
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).resolve().parent))

from common import URL_RE, has_vietnamese, non_vietnamese_font  # noqa: E402

NUMBER_RE = re.compile(r"\d+(?:[.,]\d+)*%?")
GENERIC_NAME = re.compile(r"^(text\s?box( \d+)?|content placeholder \d+)?$", re.IGNORECASE)

NON_BULLET_PREFIXES = (
    "title",
    "subtitle",
    "kicker",
    "big-number",
    "caption",
    "label",
    "source",
    "attribution",
    "footer",
    "slide-number",
    "quote",
    "image-label",
    "column-title",
    "background",
    "panel",
    "rule",
    "closing",
)
NUMBER_SKIP_PREFIXES = ("kicker", "slide-number", "footer", "source")


@dataclass
class Options:
    min_pt: float = 22
    max_bullets: int = 4
    max_numbers: int = 2
    max_words: int = 25
    require_notes: bool = False
    check_links: bool = False
    link_timeout: float = 8


@dataclass
class Issue:
    level: str
    slide: int
    message: str


@dataclass
class ShapeInfo:
    shape: object
    index: int
    name: str
    bounds: tuple
    text: str
    kind: str
    fill: str | None


def snippet(text: str, limit: int = 48) -> str:
    flat = " ".join((text or "").split())
    if len(flat) <= limit:
        return flat
    return flat[: limit - 1] + "…"


def emu_in(value) -> float:
    return float(value) / 914400


def bounds_of(shape) -> tuple[float, float, float, float]:
    left = emu_in(shape.left)
    top = emu_in(shape.top)
    return (left, top, left + emu_in(shape.width), top + emu_in(shape.height))


def solid_hex(shape) -> str | None:
    try:
        fill = shape.fill
        if fill.type != MSO_FILL.SOLID:
            return None
        return str(fill.fore_color.rgb).upper()
    except Exception:
        return None


def slide_bg_hex(slide) -> str | None:
    try:
        fill = slide.background.fill
        if fill.type == MSO_FILL.SOLID:
            return str(fill.fore_color.rgb).upper()
    except Exception:
        return None
    return None


def channel(value: int) -> float:
    c = value / 255
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4


def luminance(hex_color: str) -> float:
    h = hex_color.strip().lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b)


def contrast_ratio(foreground: str, background: str) -> float:
    a, b = luminance(foreground), luminance(background)
    lighter, darker = max(a, b), min(a, b)
    return (lighter + 0.05) / (darker + 0.05)


def min_contrast(size: float, bold: bool) -> float:
    if size >= 24 or (bold and size >= 18):
        return 3.0
    return 4.5


def classify_http(status: int) -> tuple[str, str]:
    if status in (401, 403, 405, 429):
        return "warning", f"link returned {status} (the site may be blocking the checker)"
    if status == 0 or status >= 400:
        return "error", f"broken hyperlink ({status or 'no response'})"
    return "ok", ""


def probe(url: str, timeout: float) -> int:
    headers = {"User-Agent": "OpenSlides-check/1.0"}
    for method in ("HEAD", "GET"):
        request = urllib.request.Request(url, method=method, headers=headers)
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                return int(response.status)
        except urllib.error.HTTPError as error:
            if method == "HEAD" and error.code in (400, 403, 405, 501):
                continue
            return int(error.code)
        except Exception:
            if method == "HEAD":
                continue
            return 0
    return 0


def starts_with_any(name: str, prefixes: tuple[str, ...]) -> bool:
    lowered = (name or "").strip().lower()
    return any(lowered == prefix or lowered.startswith(prefix + "-") or lowered.startswith(prefix + " ") for prefix in prefixes)


def is_bullet_name(name: str) -> bool:
    lowered = (name or "").strip().lower()
    return lowered == "bullet" or lowered.startswith("bullet-") or lowered.startswith("bullet ")


def paragraph_is_bulleted(paragraph) -> bool:
    p_pr = paragraph._p.find(qn("a:pPr"))
    if p_pr is None or p_pr.find(qn("a:buNone")) is not None:
        return False
    return any(p_pr.find(qn(tag)) is not None for tag in ("a:buChar", "a:buAutoNum", "a:buBlip"))


def run_size(run, paragraph) -> float | None:
    if run.font.size is not None:
        return float(run.font.size.pt)
    if paragraph.font.size is not None:
        return float(paragraph.font.size.pt)
    return None


def run_color(run) -> str | None:
    try:
        return str(run.font.color.rgb).upper()
    except Exception:
        return None


def run_bold(run) -> bool:
    return bool(run.font.bold)


def notes_text(slide) -> str:
    if not slide.has_notes_slide:
        return ""
    return (slide.notes_slide.notes_text_frame.text or "").strip()


def slide_title(slide) -> str:
    for shape in slide.shapes:
        name = (getattr(shape, "name", "") or "").lower()
        if name.startswith("title") and getattr(shape, "has_text_frame", False):
            text = " ".join(shape.text_frame.text.split())
            if text:
                return text
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            text = " ".join(shape.text_frame.text.split())
            if text:
                return text
    return ""


def text_does_not_fit(shape) -> bool:
    frame = shape.text_frame
    usable_w = emu_in(shape.width) - emu_in(frame.margin_left) - emu_in(frame.margin_right)
    usable_h = emu_in(shape.height) - emu_in(frame.margin_top) - emu_in(frame.margin_bottom)
    if usable_w <= 0.05 or usable_h <= 0.05:
        return True
    needed = 0.0
    for paragraph in frame.paragraphs:
        raw = paragraph.text or ""
        if not raw.strip():
            continue
        sizes = [run_size(run, paragraph) for run in paragraph.runs if (run.text or "").strip()]
        size = max([s for s in sizes if s] or [18])
        em = size / 72.0
        chars_per_line = max(1, int(usable_w / (em * 0.52)))
        lines = 0
        for piece in raw.splitlines():
            words = piece.split()
            if not words:
                lines += 1
                continue
            lines += 1
            column = 0
            for word in words:
                width = len(word)
                if column == 0:
                    column = width
                    while column > chars_per_line:
                        lines += 1
                        column -= chars_per_line
                elif column + 1 + width > chars_per_line:
                    lines += 1
                    column = width
                else:
                    column += 1 + width
        needed += lines * em * 1.05
    return needed > usable_h + 0.02


def contains(box, x: float, y: float) -> bool:
    return box[0] - 0.01 <= x <= box[2] + 0.01 and box[1] - 0.01 <= y <= box[3] + 0.01


def overlap_amount(a, b, tol: float = 0.03) -> float:
    ix = min(a[2], b[2]) - max(a[0], b[0])
    iy = min(a[3], b[3]) - max(a[1], b[1])
    if ix > tol and iy > tol:
        return ix * iy
    return 0.0


def outside(box, slide_w: float, slide_h: float, tol: float = 0.05) -> bool:
    return box[0] < -tol or box[1] < -tol or box[2] > slide_w + tol or box[3] > slide_h + tol


def collect(slide) -> list[ShapeInfo]:
    infos = []
    for index, shape in enumerate(slide.shapes):
        try:
            shape_type = shape.shape_type
        except Exception:
            shape_type = None
        if shape_type == MSO_SHAPE_TYPE.GROUP:
            continue
        try:
            box = bounds_of(shape)
        except Exception:
            continue
        text = ""
        if getattr(shape, "has_text_frame", False):
            text = " ".join(shape.text_frame.text.split())
        fill = solid_hex(shape)
        if shape_type == MSO_SHAPE_TYPE.PICTURE:
            kind = "picture"
        elif fill:
            kind = "solid"
        else:
            kind = "other"
        infos.append(ShapeInfo(shape, index, shape.name or "", box, text, kind, fill))
    return infos


def background_for(info: ShapeInfo, infos: list[ShapeInfo], slide_bg: str | None) -> tuple[str | None, bool]:
    """Return (hex behind the text, True if a picture is the surface)."""
    if info.fill:
        return info.fill, False
    cx = (info.bounds[0] + info.bounds[2]) / 2
    cy = (info.bounds[1] + info.bounds[3]) / 2
    behind = [item for item in infos if item.index < info.index and contains(item.bounds, cx, cy)]
    behind.sort(key=lambda item: item.index, reverse=True)
    for item in behind:
        if item.kind == "solid" and item.fill:
            return item.fill, False
        if item.kind == "picture":
            return None, True
    return slide_bg, False


def current_size(slide) -> tuple[float, float]:
    presentation = slide.part.package.presentation_part.presentation
    return emu_in(presentation.slide_width), emu_in(presentation.slide_height)


def _check_shapes(slide, number, options, issues, infos, slide_bg, slide_w, slide_h) -> None:
    bullets: list[str] = []
    numbers = 0
    data_slide = any(
        (info.name or "").lower().startswith("big-number")
        or _has_display_size(info, 72)
        for info in infos
    )

    for info in infos:
        if info.name.lower() == "background":
            continue
        if outside(info.bounds, slide_w, slide_h):
            issues.append(Issue("error", number, f"text box \"{info.name or 'shape'}\" is outside the slide"))
        if not info.text or not getattr(info.shape, "has_text_frame", False):
            continue
        if text_does_not_fit(info.shape):
            issues.append(Issue("error", number, f"text does not fit in \"{info.name or 'shape'}\""))
        bg, on_picture = background_for(info, infos, slide_bg)
        if on_picture:
            issues.append(Issue(
                "error", number,
                f"text sits on a picture without a solid panel in \"{info.name or 'shape'}\"",
            ))
        _check_runs(info, number, options, issues, bg)
        bullets.extend(_bullet_texts(info))
        if not data_slide and not starts_with_any(info.name, NUMBER_SKIP_PREFIXES):
            numbers += len(NUMBER_RE.findall(info.text))

    if len(bullets) > options.max_bullets:
        issues.append(Issue("error", number, f"{len(bullets)} bullets (max {options.max_bullets})"))
    for text in bullets:
        count = len(text.split())
        if count > options.max_words:
            issues.append(Issue(
                "error", number,
                f"bullet has {count} words (max {options.max_words}): \"{snippet(text)}\"",
            ))
    if not data_slide and numbers > options.max_numbers:
        issues.append(Issue(
            "error", number,
            f"{numbers} numbers on a non-data slide (max {options.max_numbers})",
        ))

    text_infos = [info for info in infos if info.text]
    for left_i in range(len(text_infos)):
        for right_i in range(left_i + 1, len(text_infos)):
            a, b = text_infos[left_i], text_infos[right_i]
            if overlap_amount(a.bounds, b.bounds):
                issues.append(Issue(
                    "error", number,
                    f"text boxes overlap: \"{a.name}\" and \"{b.name}\"",
                ))

    _check_links(slide, number, options, issues)


def _has_display_size(info: ShapeInfo, minimum: float) -> bool:
    if not info.text or not getattr(info.shape, "has_text_frame", False):
        return False
    for paragraph in info.shape.text_frame.paragraphs:
        for run in paragraph.runs:
            size = run_size(run, paragraph)
            if size is not None and size >= minimum and (run.text or "").strip():
                return True
    return False


def _check_runs(info: ShapeInfo, number: int, options: Options, issues: list[Issue], bg: str | None) -> None:
    for paragraph in info.shape.text_frame.paragraphs:
        for run in paragraph.runs:
            text = run.text or ""
            if not text.strip():
                continue
            size = run_size(run, paragraph)
            if size is None:
                issues.append(Issue("warning", number, f"font size not set explicitly in \"{snippet(text)}\""))
            elif size < options.min_pt:
                issues.append(Issue(
                    "error", number,
                    f"font {size:g}pt is below {options.min_pt:g}pt in \"{snippet(text)}\"",
                ))
            font_name = run.font.name or ""
            if has_vietnamese(text) and non_vietnamese_font(font_name):
                issues.append(Issue(
                    "error", number,
                    f"Vietnamese text is set in {font_name}, which has no diacritics",
                ))
            elif has_vietnamese(text) and not font_name:
                issues.append(Issue("warning", number, f"font not set on Vietnamese text \"{snippet(text)}\""))
            color = run_color(run)
            if color and bg and size is not None:
                ratio = contrast_ratio(color, bg)
                if ratio < min_contrast(size, run_bold(run)):
                    issues.append(Issue(
                        "error", number,
                        f"contrast {ratio:.1f}:1 ({color} on {bg}) in \"{snippet(text)}\"",
                    ))
            elif color and not bg:
                issues.append(Issue(
                    "warning", number,
                    f"contrast not determinable in \"{snippet(text)}\"",
                ))


def _bullet_texts(info: ShapeInfo) -> list[str]:
    paragraphs = [p for p in info.shape.text_frame.paragraphs if (p.text or "").strip()]
    generic = bool(GENERIC_NAME.match((info.name or "").strip()))
    found = []
    for paragraph in paragraphs:
        text = " ".join(paragraph.text.split())
        if is_bullet_name(info.name) or paragraph_is_bulleted(paragraph):
            found.append(text)
        elif generic and len(paragraphs) >= 2 and not starts_with_any(info.name, NON_BULLET_PREFIXES):
            found.append(text)
    return found


def _check_links(slide, number: int, options: Options, issues: list[Issue]) -> None:
    seen = set()
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                try:
                    address = run.hyperlink.address
                except Exception:
                    address = None
                if not address or address in seen:
                    continue
                seen.add(address)
                if not URL_RE.match(address):
                    issues.append(Issue("error", number, f"malformed hyperlink: {address}"))
                    continue
                if not options.check_links:
                    continue
                status = probe(address, options.link_timeout)
                level, message = classify_http(status)
                if level != "ok":
                    issues.append(Issue(level, number, f"{message}: {address}"))


def check_file(path: Path, options: Options) -> tuple[object, list[Issue]]:
    prs = Presentation(str(path))
    issues: list[Issue] = []
    for number, slide in enumerate(prs.slides, start=1):
        if slide.shapes._spTree.find(".//" + qn("p:grpSp")) is not None:
            issues.append(Issue("error", number, "grouped shapes; ungroup so every text box stays editable"))
        if options.require_notes and not notes_text(slide):
            issues.append(Issue("error", number, "missing speaker notes"))
        infos = collect(slide)
        _check_shapes(
            slide, number, options, issues, infos, slide_bg_hex(slide),
            *current_size(slide),
        )
    return prs, issues


def format_report(path: Path, prs, issues: list[Issue], options: Options) -> str:
    lines = [
        "OpenSlides check",
        f"File: {path}",
        f"Slides: {len(prs.slides)}",
        (
            f"Rules: text >= {options.min_pt:g}pt, bullets <= {options.max_bullets}, "
            f"numbers <= {options.max_numbers} outside data slides, "
            f"bullet <= {options.max_words} words"
        ),
        "",
    ]
    grouped: dict[int, list[Issue]] = {}
    for issue in issues:
        grouped.setdefault(issue.slide, []).append(issue)
    for number, slide in enumerate(prs.slides, start=1):
        title = slide_title(slide)
        lines.append(f"Slide {number} — {title}" if title else f"Slide {number}")
        slide_issues = grouped.get(number, [])
        if not slide_issues:
            lines.append("  ok")
        else:
            for issue in slide_issues:
                lines.append(f"  {issue.level.upper()}  {issue.message}")
        lines.append("")
    errors = sum(1 for issue in issues if issue.level == "error")
    warnings = sum(1 for issue in issues if issue.level == "warning")
    status = "pass" if errors == 0 else "fail"
    lines.append(f"Result: {status} ({errors} errors, {warnings} warnings)")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Lint a .pptx for readability problems.")
    parser.add_argument("pptx", help="Path to the .pptx")
    parser.add_argument("--min-pt", type=float, default=22, help="Flag text smaller than this (default 22)")
    parser.add_argument("--max-bullets", type=int, default=4)
    parser.add_argument("--max-numbers", type=int, default=2)
    parser.add_argument("--max-words", type=int, default=25)
    parser.add_argument("--require-notes", action="store_true", help="Fail slides that have no speaker notes")
    parser.add_argument("--check-links", action="store_true", help="Request each http(s) link and flag failures")
    args = parser.parse_args(argv)
    path = Path(args.pptx)
    if not path.is_file():
        print(f"error: file not found: {path}", file=sys.stderr)
        return 2
    options = Options(
        min_pt=args.min_pt,
        max_bullets=args.max_bullets,
        max_numbers=args.max_numbers,
        max_words=args.max_words,
        require_notes=args.require_notes,
        check_links=args.check_links,
    )
    try:
        prs, issues = check_file(path, options)
    except Exception as exc:
        print(f"error: cannot open {path}: {exc}", file=sys.stderr)
        return 2
    print(format_report(path, prs, issues, options))
    return 1 if any(issue.level == "error" for issue in issues) else 0


if __name__ == "__main__":
    sys.exit(main())
