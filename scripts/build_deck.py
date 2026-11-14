#!/usr/bin/env python3
"""Build an editable 16:9 .pptx from a YAML or JSON deck spec.

Every text element is its own text box. Nothing is grouped or flattened.
Slide size is 13.333 x 7.5 in (the usual 1920 x 1080 frame).
"""

from __future__ import annotations

import argparse
import json
import sys
import tempfile
from pathlib import Path

import yaml
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))

from common import (  # noqa: E402
    SLIDE_H_EMU,
    SLIDE_H_IN,
    SLIDE_W_EMU,
    SLIDE_W_IN,
    URL_RE,
    has_vietnamese,
    non_vietnamese_font,
)

ROOT = Path(__file__).resolve().parents[1]
FLOOR_PT = 22.0
CANVA_EXTRA_IN = 30 / 144  # about 30px on a 1920-wide slide

DEFAULT_THEME = {
    "name": "default",
    "fonts": {
        "body": "Be Vietnam Pro SemiBold",
        "body_bold": False,
        "heading": "Be Vietnam Pro",
        "heading_bold": True,
        "display": "Anton",
        "display_bold": False,
        "vietnamese_fallback": "Be Vietnam Pro SemiBold",
        "display_supports_vietnamese": False,
    },
    "colors": {
        "bg": "F4F0E8",
        "ink": "1C1915",
        "accent": "8C3A1E",
        "accent_on_dark": "E7A07A",
        "panel": "FFFCF8",
        "inverse": "1C1915",
        "on_inverse": "F4F0E8",
    },
    "sizes_pt": {
        "cover_title": 48,
        "title": 36,
        "section_title": 40,
        "section_kicker": 56,
        "column_title": 32,
        "body": 28,
        "kicker": 24,
        "big_number": 108,
        "caption": 24,
        "source": 24,
        "quote": 32,
    },
    "spacing_in": {
        "margin_x": 0.7,
        "margin_y": 0.46,
        "bullet_gap": 0.18,
        "canva_extra_in": 0.0,
    },
}

LAYOUTS = (
    "cover",
    "section",
    "bullets",
    "image-left",
    "image-right",
    "big-number",
    "comparison",
    "quote",
    "closing",
    "sources",
)


def die(message: str) -> None:
    raise SystemExit(f"error: {message}")


def text_of(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def deep_merge(base: dict, override: dict) -> dict:
    merged = dict(base)
    for key, value in (override or {}).items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            merged[key] = deep_merge(merged[key], value)
        else:
            merged[key] = value
    return merged


def rgb(value: str) -> RGBColor:
    return RGBColor.from_string(str(value).strip().lstrip("#"))


def pt(theme: dict, key: str) -> float:
    try:
        value = float(theme["sizes_pt"][key])
    except KeyError:
        die(f"theme is missing sizes_pt.{key}")
    except (TypeError, ValueError):
        die(f"theme size {key} is not a number")
    if value < FLOOR_PT:
        print(f"warning: {key} is {value:g}pt; using {FLOOR_PT:g}pt", file=sys.stderr)
        return FLOOR_PT
    return value


def choose_font(theme: dict, role: str, text: str) -> tuple[str, bool]:
    fonts = theme["fonts"]
    name = fonts.get(role) or fonts["body"]
    bold = bool(fonts.get(f"{role}_bold", role == "heading"))
    needs_fallback = has_vietnamese(text) and (
        non_vietnamese_font(name)
        or (role == "display" and not fonts.get("display_supports_vietnamese", False))
    )
    if needs_fallback:
        return fonts.get("vietnamese_fallback") or "Be Vietnam Pro SemiBold", False
    return name, bold


def run_spec(theme, text, role, size, color, bold=None, url=None) -> dict:
    name, role_bold = choose_font(theme, role, text)
    return {
        "text": text,
        "font": name,
        "size": size,
        "bold": role_bold if bold is None else bold,
        "color": color,
        "url": url,
    }


def bullet_runs(theme, label: str, text: str, size: float, color: str) -> list[dict]:
    runs = []
    if label:
        shown = label if label.endswith(":") else f"{label}:"
        runs.append(run_spec(theme, shown + " ", "heading", size, color, bold=True))
    runs.append(run_spec(theme, text, "body", size, color, bold=False))
    return runs


class Geom:
    def __init__(self, theme: dict):
        spacing = theme["spacing_in"]
        self.mx = float(spacing["margin_x"])
        self.my = float(spacing["margin_y"])
        self.gap = float(spacing["bullet_gap"]) + float(spacing.get("canva_extra_in") or 0)
        self.bottom = SLIDE_H_IN - self.my
        self.width = SLIDE_W_IN - 2 * self.mx


class Ctx:
    def __init__(self, theme, spec_path: Path, index: int, max_image_px: int | None):
        self.theme = theme
        self.spec_path = spec_path
        self.index = index
        self.max_image_px = max_image_px
        self.tmpdir = tempfile.TemporaryDirectory(prefix="openslides-")

    def fail(self, message: str) -> None:
        die(f"slide {self.index}: {message}")

    def close(self) -> None:
        self.tmpdir.cleanup()


def stack(ctx: Ctx, top: float, bottom: float, count: int, gap: float, max_h: float, min_h: float):
    if count <= 0:
        return 0.0, []
    usable = bottom - top - gap * (count - 1)
    height = usable / count
    height = min(height, max_h)
    if height < min_h:
        ctx.fail(
            f"{count} text boxes do not fit. Split the slide "
            f"(available height {height:.2f}in, need {min_h:.2f}in)."
        )
    ys = [top + i * (height + gap) for i in range(count)]
    return height, ys


def apply_run(run, spec: dict) -> None:
    run.text = spec["text"]
    lang = "vi-VN" if has_vietnamese(spec["text"]) else "en-US"

    def paint() -> None:
        run.font.name = spec["font"]
        run.font.size = Pt(spec["size"])
        run.font.bold = bool(spec["bold"])
        run.font.italic = False
        run.font.color.rgb = rgb(spec["color"])
        run.font.underline = bool(spec.get("url"))
        r_pr = run._r.get_or_add_rPr()
        r_pr.set("lang", lang)
        for tag in ("latin", "ea", "cs"):
            element = r_pr.find(qn(f"a:{tag}"))
            if element is None:
                element = etree.SubElement(r_pr, qn(f"a:{tag}"))
            element.set("typeface", spec["font"])

    paint()
    if spec.get("url"):
        run.hyperlink.address = spec["url"]
        paint()


def write_paragraph(paragraph, runs: list[dict]) -> None:
    paragraph.alignment = PP_ALIGN.LEFT
    paragraph.level = 0
    paragraph.space_before = Pt(0)
    paragraph.space_after = Pt(0)
    paragraph.line_spacing = 1.0
    for child in list(paragraph._p):
        if child.tag in (qn("a:r"), qn("a:br")):
            paragraph._p.remove(child)
    for spec in runs:
        for index, line in enumerate(spec["text"].split("\n")):
            if index:
                etree.SubElement(paragraph._p, qn("a:br"))
            piece = dict(spec)
            piece["text"] = line
            apply_run(paragraph.add_run(), piece)


def add_text(slide, name: str, x: float, y: float, w: float, h: float, paragraphs: list[list[dict]]):
    if x < -0.02 or y < -0.02 or x + w > SLIDE_W_IN + 0.02 or y + h > SLIDE_H_IN + 0.02:
        die(f"'{name}' is outside the slide ({x:.2f}, {y:.2f}, {w:.2f} x {h:.2f})")
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    shape.name = name
    frame = shape.text_frame
    frame.word_wrap = True
    frame.margin_left = Inches(0.04)
    frame.margin_right = Inches(0.04)
    frame.margin_top = Inches(0.02)
    frame.margin_bottom = Inches(0.02)
    body_pr = frame._txBody.bodyPr
    body_pr.set("anchor", "t")
    for child in list(body_pr):
        if "Fit" in child.tag:
            body_pr.remove(child)
    write_paragraph(frame.paragraphs[0], paragraphs[0])
    for extra in paragraphs[1:]:
        write_paragraph(frame.add_paragraph(), extra)
    return shape


def add_rect(slide, x, y, w, h, fill: str, name: str):
    if name == "background":
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(0), Emu(0), Emu(SLIDE_W_EMU), Emu(SLIDE_H_EMU))
    else:
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.name = name
    shape.line.fill.background()
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    return shape


def paint(slide, color: str) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(color)
    add_rect(slide, 0, 0, SLIDE_W_IN, SLIDE_H_IN, color, "background")


def set_description(shape, text: str) -> None:
    if not text:
        return
    c_nv_pr = shape._element.find(".//" + qn("p:cNvPr"))
    if c_nv_pr is not None:
        c_nv_pr.set("descr", text)


def resolve_image(ctx: Ctx, value: str) -> Path:
    raw = Path(value)
    candidates = [raw] if raw.is_absolute() else [
        Path.cwd() / raw,
        ctx.spec_path.parent / raw,
        ROOT / raw,
    ]
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    ctx.fail(f"image not found: {value}")
    raise AssertionError


def shrink_image(ctx: Ctx, path: Path) -> Path:
    if not ctx.max_image_px:
        return path
    from PIL import Image

    with Image.open(path) as image:
        width, height = image.size
        long_edge = max(width, height)
        if long_edge <= ctx.max_image_px:
            return path
        scale = ctx.max_image_px / long_edge
        resized = image.resize(
            (max(1, round(width * scale)), max(1, round(height * scale))),
            Image.Resampling.LANCZOS,
        )
        dest = Path(ctx.tmpdir.name) / (path.stem + "-lite.png")
        resized.save(dest, "PNG", optimize=True)
        return dest


def add_picture_fitted(slide, path: Path, x, y, w, h, description: str):
    from PIL import Image

    with Image.open(path) as image:
        iw, ih = image.size
    ratio = iw / ih
    box_ratio = w / h
    if ratio > box_ratio:
        dw, dh = w, w / ratio
    else:
        dh, dw = h, h * ratio
    ox = x + (w - dw) / 2
    oy = y + (h - dh) / 2
    picture = slide.shapes.add_picture(str(path), Inches(ox), Inches(oy), Inches(dw), Inches(dh))
    picture.name = "image"
    set_description(picture, description)
    return picture


def as_bullet(ctx: Ctx, item) -> tuple[str, str]:
    if isinstance(item, str):
        text = text_of(item)
        if not text:
            ctx.fail("a bullet is empty")
        return "", text
    if isinstance(item, dict):
        text = text_of(item.get("text") if item.get("text") is not None else item.get("body"))
        if not text:
            ctx.fail("a bullet is missing text")
        return text_of(item.get("label")), text
    ctx.fail("a bullet must be a string or {label, text}")
    raise AssertionError


def require_title(ctx: Ctx, spec: dict) -> str:
    title = text_of(spec.get("title"))
    if not title:
        ctx.fail(f"{spec.get('layout')} layout needs a title")
    return title


def layout_cover(slide, spec, theme, ctx: Ctx, geom: Geom) -> None:
    colors = theme["colors"]
    paint(slide, colors["inverse"])
    y = 1.7
    kicker = text_of(spec.get("kicker"))
    if kicker:
        add_text(
            slide, "kicker", geom.mx, y, geom.width, 0.5,
            [[run_spec(theme, kicker, "heading", pt(theme, "kicker"), colors["accent_on_dark"], bold=True)]],
        )
        y += 0.5 + geom.gap
    add_text(
        slide, "title", geom.mx, y, geom.width, 2.2,
        [[run_spec(theme, require_title(ctx, spec), "heading", pt(theme, "cover_title"), colors["on_inverse"], bold=True)]],
    )
    y += 2.2 + geom.gap
    subtitle = text_of(spec.get("subtitle"))
    if subtitle:
        add_text(
            slide, "subtitle", geom.mx, y, min(geom.width, 11.4), 1.4,
            [[run_spec(theme, subtitle, "body", pt(theme, "body"), colors["on_inverse"], bold=False)]],
        )


def layout_section(slide, spec, theme, ctx: Ctx, geom: Geom) -> None:
    colors = theme["colors"]
    paint(slide, colors["bg"])
    y = 2.1
    kicker = text_of(spec.get("kicker"))
    if kicker:
        if len(kicker) <= 4 and not has_vietnamese(kicker):
            role, size, height = "display", pt(theme, "section_kicker"), 1.25
        else:
            role, size, height = "heading", pt(theme, "kicker"), 0.55
        add_text(
            slide, "kicker", geom.mx, y, geom.width, height,
            [[run_spec(theme, kicker, role, size, colors["accent"])]],
        )
        y += height + geom.gap
    add_text(
        slide, "title", geom.mx, y, geom.width, 1.8,
        [[run_spec(theme, require_title(ctx, spec), "heading", pt(theme, "section_title"), colors["ink"], bold=True)]],
    )


def layout_bullets(slide, spec, theme, ctx: Ctx, geom: Geom) -> None:
    colors = theme["colors"]
    paint(slide, colors["bg"])
    title_h = 1.2
    add_text(
        slide, "title", geom.mx, geom.my, geom.width, title_h,
        [[run_spec(theme, require_title(ctx, spec), "heading", pt(theme, "title"), colors["ink"], bold=True)]],
    )
    bullets = spec.get("bullets") or []
    if not bullets:
        ctx.fail("bullets layout needs at least one bullet")
    top = geom.my + title_h + geom.gap
    height, ys = stack(ctx, top, geom.bottom, len(bullets), geom.gap, max_h=1.4, min_h=0.85)
    body = pt(theme, "body")
    for index, item in enumerate(bullets):
        label, text = as_bullet(ctx, item)
        add_text(
            slide, f"bullet-{index + 1}", geom.mx, ys[index], geom.width, height,
            [bullet_runs(theme, label, text, body, colors["ink"])],
        )


def layout_image(slide, spec, theme, ctx: Ctx, geom: Geom, side: str) -> None:
    colors = theme["colors"]
    paint(slide, colors["bg"])
    panel_w = 6.2
    panel_h = SLIDE_H_IN - 2 * geom.my
    panel_y = geom.my
    if side == "left":
        panel_x = geom.mx
        text_x = panel_x + panel_w + 0.38
        text_w = SLIDE_W_IN - geom.mx - text_x
    else:
        panel_x = SLIDE_W_IN - geom.mx - panel_w
        text_x = geom.mx
        text_w = panel_x - 0.38 - text_x
    add_rect(slide, panel_x, panel_y, panel_w, panel_h, colors["panel"], "panel")
    label = text_of(spec.get("image_label")) or "Replace with a photograph."
    image = text_of(spec.get("image"))
    if image:
        path = shrink_image(ctx, resolve_image(ctx, image))
        add_picture_fitted(
            slide, path, panel_x + 0.18, panel_y + 0.18, panel_w - 0.36, panel_h - 0.36, label,
        )
    else:
        add_text(
            slide, "image-label",
            panel_x + 0.35, panel_y + (panel_h - 1.3) / 2, panel_w - 0.7, 1.3,
            [[run_spec(theme, label, "body", pt(theme, "body"), colors["ink"], bold=False)]],
        )
    title_h = 1.15
    add_text(
        slide, "title", text_x, geom.my, text_w, title_h,
        [[run_spec(theme, require_title(ctx, spec), "heading", pt(theme, "column_title"), colors["ink"], bold=True)]],
    )
    bullets = spec.get("bullets") or []
    top = geom.my + title_h + geom.gap
    height, ys = stack(ctx, top, geom.bottom, len(bullets), geom.gap, max_h=1.6, min_h=0.9)
    body = pt(theme, "body")
    for index, item in enumerate(bullets):
        bullet_label, text = as_bullet(ctx, item)
        add_text(
            slide, f"bullet-{index + 1}", text_x, ys[index], text_w, height,
            [bullet_runs(theme, bullet_label, text, body, colors["ink"])],
        )


def layout_big_number(slide, spec, theme, ctx: Ctx, geom: Geom) -> None:
    colors = theme["colors"]
    paint(slide, colors["bg"])
    number = text_of(spec.get("number"))
    label = text_of(spec.get("label"))
    caption = text_of(spec.get("caption"))
    if not number or not label or not caption:
        ctx.fail("big-number layout needs number, label, and caption")
    size = pt(theme, "big_number")
    role = "display"
    if has_vietnamese(number) or len(number) > 5:
        role, size = "heading", min(size, 72)
    y = 1.35
    add_text(
        slide, "big-number", geom.mx, y, geom.width, 2.15,
        [[run_spec(theme, number, role, size, colors["accent"])]],
    )
    y += 2.15 + geom.gap
    add_text(
        slide, "label", geom.mx, y, geom.width, 0.75,
        [[run_spec(theme, label, "heading", pt(theme, "column_title"), colors["ink"], bold=True)]],
    )
    y += 0.75 + geom.gap
    add_text(
        slide, "caption", geom.mx, y, min(geom.width, 11.2), 1.55,
        [[run_spec(theme, caption, "body", pt(theme, "caption"), colors["ink"], bold=False)]],
    )


def layout_comparison(slide, spec, theme, ctx: Ctx, geom: Geom) -> None:
    colors = theme["colors"]
    paint(slide, colors["bg"])
    title_h = 1.05
    add_text(
        slide, "title", geom.mx, geom.my, geom.width, title_h,
        [[run_spec(theme, require_title(ctx, spec), "heading", pt(theme, "title"), colors["ink"], bold=True)]],
    )
    gutter = 0.48
    col_w = (geom.width - gutter) / 2
    left_x = geom.mx
    right_x = geom.mx + col_w + gutter
    add_rect(slide, geom.mx + col_w + (gutter - 0.03) / 2, geom.my + title_h + 0.15, 0.03, geom.bottom - (geom.my + title_h + 0.3), colors["accent"], "rule")
    body = pt(theme, "body")
    for key, origin in (("left", left_x), ("right", right_x)):
        column = spec.get(key)
        if not isinstance(column, dict):
            ctx.fail(f"comparison layout needs {key}.title and {key}.bullets")
        heading = text_of(column.get("title") or column.get("heading"))
        if not heading:
            ctx.fail(f"comparison {key} column needs a title")
        top = geom.my + title_h + geom.gap
        add_text(
            slide, f"column-title-{key}", origin, top, col_w, 0.58,
            [[run_spec(theme, heading, "heading", pt(theme, "body"), colors["ink"], bold=True)]],
        )
        bullets = column.get("bullets") or []
        if not bullets:
            ctx.fail(f"comparison {key} column needs bullets")
        stack_top = top + 0.58 + geom.gap
        height, ys = stack(ctx, stack_top, geom.bottom, len(bullets), geom.gap, max_h=1.55, min_h=0.9)
        for index, item in enumerate(bullets):
            label, text = as_bullet(ctx, item)
            add_text(
                slide, f"bullet-{key}-{index + 1}", origin, ys[index], col_w, height,
                [bullet_runs(theme, label, text, body, colors["ink"])],
            )


def layout_quote(slide, spec, theme, ctx: Ctx, geom: Geom) -> None:
    colors = theme["colors"]
    paint(slide, colors["bg"])
    quote = text_of(spec.get("quote") or spec.get("text"))
    if not quote:
        ctx.fail("quote layout needs a quote")
    y = 1.55
    add_text(
        slide, "quote", geom.mx, y, geom.width, 3.05,
        [[run_spec(theme, quote, "body", pt(theme, "quote"), colors["ink"], bold=False)]],
    )
    y += 3.05 + geom.gap
    attribution = text_of(spec.get("attribution"))
    if attribution:
        add_text(
            slide, "attribution", geom.mx, y, geom.width, 1.2,
            [[run_spec(theme, attribution, "body", pt(theme, "caption"), colors["ink"], bold=False)]],
        )


def layout_closing(slide, spec, theme, ctx: Ctx, geom: Geom) -> None:
    colors = theme["colors"]
    paint(slide, colors["inverse"])
    y = 2.15
    add_text(
        slide, "title", geom.mx, y, geom.width, 1.8,
        [[run_spec(theme, require_title(ctx, spec), "heading", pt(theme, "section_title"), colors["on_inverse"], bold=True)]],
    )
    y += 1.8 + geom.gap
    subtitle = text_of(spec.get("subtitle") or spec.get("text"))
    if subtitle:
        add_text(
            slide, "subtitle", geom.mx, y, min(geom.width, 11.2), 1.35,
            [[run_spec(theme, subtitle, "body", pt(theme, "body"), colors["on_inverse"], bold=False)]],
        )


def layout_sources(slide, spec, theme, ctx: Ctx, geom: Geom) -> None:
    colors = theme["colors"]
    paint(slide, colors["bg"])
    title = text_of(spec.get("title")) or "Sources"
    title_h = 1.1
    add_text(
        slide, "title", geom.mx, geom.my, geom.width, title_h,
        [[run_spec(theme, title, "heading", pt(theme, "title"), colors["ink"], bold=True)]],
    )
    items = spec.get("sources") or []
    if not items:
        ctx.fail("sources layout needs a list of sources")
    top = geom.my + title_h + geom.gap
    height, ys = stack(ctx, top, geom.bottom, len(items), geom.gap, max_h=1.25, min_h=0.9)
    size = pt(theme, "source")
    for index, item in enumerate(items):
        if not isinstance(item, dict):
            ctx.fail("each source must be {label, url}")
        label = text_of(item.get("label"))
        url = text_of(item.get("url"))
        if not label or not url:
            ctx.fail("each source needs a label and a url")
        if not URL_RE.match(url):
            ctx.fail(f"source url must start with http:// or https:// ({url})")
        add_text(
            slide, f"source-{index + 1}", geom.mx, ys[index], geom.width, height,
            [
                [run_spec(theme, label, "heading", size, colors["accent"], bold=True, url=url)],
                [run_spec(theme, url, "body", size, colors["ink"], bold=False, url=url)],
            ],
        )


def write_notes(slide, spec: dict) -> None:
    chunks = []
    notes = text_of(spec.get("notes"))
    if notes:
        chunks.append(notes)
    if text_of(spec.get("layout")) != "sources":
        lines = []
        for item in spec.get("sources") or []:
            if isinstance(item, dict):
                label = text_of(item.get("label"))
                url = text_of(item.get("url"))
                if label or url:
                    lines.append(f"{label}: {url}".strip())
        if lines:
            chunks.append("Sources:\n" + "\n".join(lines))
    if chunks:
        slide.notes_slide.notes_text_frame.text = "\n\n".join(chunks)


BUILDERS = {
    "cover": layout_cover,
    "section": layout_section,
    "bullets": layout_bullets,
    "image-left": lambda *args: layout_image(*args, "left"),
    "image-right": lambda *args: layout_image(*args, "right"),
    "big-number": layout_big_number,
    "comparison": layout_comparison,
    "quote": layout_quote,
    "closing": layout_closing,
    "sources": layout_sources,
}


def paint_theme_links(prs: Presentation, theme: dict) -> None:
    """Point the theme's hyperlink colors at the accent so links are not default blue."""
    accent = str(theme["colors"]["accent"]).strip().lstrip("#").upper()
    master_part = prs.slide_masters[0].part
    theme_part = None
    for rel in master_part.rels.values():
        if rel.reltype.rsplit("/", 1)[-1] == "theme":
            theme_part = rel.target_part
            break
    if theme_part is None:
        return
    element = etree.fromstring(theme_part.blob)
    scheme = element.find(".//" + qn("a:clrScheme"))
    if scheme is None:
        return
    for tag in ("hlink", "folHlink"):
        node = scheme.find(qn(f"a:{tag}"))
        if node is None:
            node = etree.SubElement(scheme, qn(f"a:{tag}"))
        for child in list(node):
            node.remove(child)
        srgb = etree.SubElement(node, qn("a:srgbClr"))
        srgb.set("val", accent)
    theme_part.blob = etree.tostring(element, xml_declaration=True, encoding="UTF-8", standalone=True)


def blank_layout(prs: Presentation):
    for layout in prs.slide_layouts:
        if layout.name.strip().lower() == "blank":
            return layout
    return prs.slide_layouts[-1]


def load_spec(path: Path) -> dict:
    raw = path.read_text(encoding="utf-8")
    if path.suffix.lower() == ".json":
        data = json.loads(raw)
    else:
        data = yaml.safe_load(raw)
    if not isinstance(data, dict):
        die("deck spec must be a mapping with a slides list")
    return data


def resolve_theme(spec: dict, spec_path: Path, override: str | None) -> dict:
    if override:
        raw = Path(override)
    elif spec.get("theme"):
        raw = Path(str(spec["theme"]))
    else:
        raw = ROOT / "templates" / "theme-default.yaml"
    if raw.is_file():
        path = raw
    else:
        candidates = [
            Path.cwd() / raw,
            spec_path.parent / raw,
            ROOT / "templates" / Path(raw).name,
            ROOT / raw,
        ]
        path = next((candidate for candidate in candidates if candidate.is_file()), None)
    if path is None:
        die(f"theme not found: {raw}")
    loaded = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    if not isinstance(loaded, dict):
        die(f"theme is not a mapping: {path}")
    return deep_merge(DEFAULT_THEME, loaded)


def build(spec: dict, theme: dict, spec_path: Path, max_image_px: int | None) -> Presentation:
    slides = spec.get("slides")
    if not isinstance(slides, list) or not slides:
        die("deck spec needs a non-empty slides list")
    prs = Presentation()
    prs.slide_width = Emu(SLIDE_W_EMU)
    prs.slide_height = Emu(SLIDE_H_EMU)
    prs.core_properties.title = text_of(spec.get("title")) or "Slides"
    prs.core_properties.author = "OpenSlides"
    blank = blank_layout(prs)
    geom = Geom(theme)
    contexts = []
    try:
        for index, slide_spec in enumerate(slides, start=1):
            if not isinstance(slide_spec, dict):
                die(f"slide {index} is not a mapping")
            layout = text_of(slide_spec.get("layout")).lower()
            if layout not in BUILDERS:
                die(f"slide {index}: unknown layout '{layout}'. Known: {', '.join(LAYOUTS)}")
            ctx = Ctx(theme, spec_path, index, max_image_px)
            contexts.append(ctx)
            slide = prs.slides.add_slide(blank)
            BUILDERS[layout](slide, slide_spec, theme, ctx, geom)
            write_notes(slide, slide_spec)
        paint_theme_links(prs, theme)
    finally:
        for ctx in contexts:
            ctx.close()
    return prs


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Build an editable 16:9 pptx from a YAML or JSON deck spec.",
    )
    parser.add_argument("spec", help="Path to the deck YAML or JSON")
    parser.add_argument("-o", "--output", required=True, help="Where to write the .pptx")
    parser.add_argument("--theme", help="Theme YAML. Overrides the theme named in the spec.")
    parser.add_argument(
        "--canva",
        action="store_true",
        help="Add about 30px of extra gap so text still fits after a Canva import.",
    )
    parser.add_argument(
        "--max-image-px",
        type=int,
        default=0,
        help="Downscale embedded images so the long edge is at most this many pixels.",
    )
    args = parser.parse_args(argv)
    spec_path = Path(args.spec)
    if not spec_path.is_file():
        die(f"spec not found: {spec_path}")
    spec = load_spec(spec_path)
    theme = resolve_theme(spec, spec_path, args.theme)
    if args.canva:
        theme["spacing_in"]["canva_extra_in"] = max(
            float(theme["spacing_in"].get("canva_extra_in") or 0),
            CANVA_EXTRA_IN,
        )
    prs = build(spec, theme, spec_path, args.max_image_px or None)
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output)
    print(f"Wrote {output} ({len(spec['slides'])} slides)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
