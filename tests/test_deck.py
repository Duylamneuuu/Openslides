"""Build the example deck, expect a clean check, and expect a bad deck to fail."""

from __future__ import annotations

import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest
from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "tests" / "fixtures"))

import bad_deck  # noqa: E402
import check_deck  # noqa: E402
from common import has_vietnamese  # noqa: E402

PYTHON = sys.executable
VIETNAMESE = "Mỗi ý là một câu đủ nghĩa và giữ đủ dấu tiếng Việt."


def run(args, **kwargs):
    return subprocess.run(
        args, cwd=ROOT, capture_output=True, text=True, check=False, **kwargs,
    )


def build(spec: Path, output: Path, *extra: str):
    result = run([PYTHON, "scripts/build_deck.py", str(spec), "-o", str(output), *extra])
    assert result.returncode == 0, result.stderr + result.stdout
    return result


def check(path: Path, *extra: str):
    return run([PYTHON, "scripts/check_deck.py", str(path), *extra])


def test_example_builds_and_passes(tmp_path: Path):
    output = tmp_path / "example.pptx"
    build(ROOT / "templates" / "example-deck.yaml", output)
    result = check(output, "--require-notes")
    assert result.returncode == 0, result.stdout + result.stderr
    assert "Result: pass" in result.stdout
    assert "0 errors" in result.stdout
    assert "WARNING" not in result.stdout

    prs = Presentation(str(output))
    assert len(prs.slides) == 8
    assert abs(prs.slide_width / 914400 - 13.333333) < 0.01
    assert abs(prs.slide_height / 914400 - 7.5) < 0.01

    quote_font = ""
    number_font = ""
    saw_link = False
    texts = []
    for slide in prs.slides:
        assert slide.notes_slide.notes_text_frame.text.strip()
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            texts.append(shape.text_frame.text)
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if VIETNAMESE[:6] in (run.text or ""):
                        quote_font = run.font.name or ""
                    if shape.name == "big-number":
                        number_font = run.font.name or ""
                    address = run.hyperlink.address or ""
                    if "plainlanguage.gov" in address:
                        saw_link = True
    blob = "\n".join(texts)
    assert VIETNAMESE in blob
    assert "Be Vietnam Pro" in quote_font
    assert "Anton" not in quote_font
    assert number_font == "Anton"
    assert saw_link

    with zipfile.ZipFile(output) as package:
        for name in package.namelist():
            if name.startswith("ppt/slides/slide"):
                xml = package.read(name).decode("utf-8")
                assert "<p:grpSp>" not in xml and "<p:grpSp " not in xml


def test_bad_deck_is_flagged(tmp_path: Path):
    output = tmp_path / "bad.pptx"
    bad_deck.build(output)
    result = check(output)
    assert result.returncode != 0
    report = result.stdout
    for phrase in (
        "below 22pt",
        "bullets",
        "words",
        "numbers",
        "outside the slide",
        "does not fit",
        "overlap",
        "contrast",
        "grouped",
        "hyperlink",
    ):
        assert phrase in report, report

    noted = check(output, "--require-notes")
    assert "missing speaker notes" in noted.stdout


def test_every_layout_and_ink_theme(tmp_path: Path):
    spec = tmp_path / "all.yaml"
    spec.write_text(
        """
title: Every layout
slides:
  - layout: cover
    title: Open with the point
    subtitle: One sentence on what this meeting is for.
    notes: Say the outcome first.
  - layout: section
    kicker: "02"
    title: Move to the next part
    notes: Pause before the new section.
  - layout: bullets
    title: Keep the list short
    bullets:
      - {label: First, text: State the point in one plain sentence.}
      - {label: Second, text: Stop before you add a fifth idea.}
    notes: Two points are enough here.
  - layout: image-right
    title: Leave room for the photo
    image_label: Replace with a photograph.
    bullets:
      - {label: Panel, text: Keep the words on the solid side.}
    notes: The panel is a stand-in.
  - layout: big-number
    number: "4"
    label: points at most
    caption: This is a writing limit, not a measured result.
    notes: Replace the figure when you have a source.
  - layout: comparison
    title: Choose the readable column
    left:
      title: Hard to read
      bullets:
        - {label: Dense, text: Too many lines at once.}
    right:
      title: Easy to read
      bullets:
        - {label: Short, text: One sentence and then stop.}
    notes: Pick the readable column.
  - layout: quote
    quote: "Mỗi ý là một câu đủ nghĩa và giữ đủ dấu tiếng Việt."
    attribution: Replace with a real quotation and its source.
    notes: Read the marks on the rendered slide.
  - layout: closing
    title: Send the editable file
    subtitle: Include the notes and the preview images.
    notes: Do not send a flat picture of the slides.
  - layout: sources
    title: Sources
    sources:
      - {label: Plain language guidelines, url: https://www.plainlanguage.gov/}
    notes: Only keep links the deck uses.
""",
        encoding="utf-8",
    )
    output = tmp_path / "all.pptx"
    build(spec, output)
    result = check(output, "--require-notes")
    assert result.returncode == 0, result.stdout

    ink = tmp_path / "ink.pptx"
    build(ROOT / "templates" / "example-deck.yaml", ink, "--theme", "templates/theme-ink.yaml")
    ink_result = check(ink)
    assert ink_result.returncode == 0, ink_result.stdout


def test_canva_gap_json_and_unknown_layout(tmp_path: Path):
    spec = tmp_path / "gap.yaml"
    spec.write_text(
        """
title: Gap
slides:
  - layout: bullets
    title: Two points
    bullets:
      - {label: One, text: The first point stays short.}
      - {label: Two, text: The second point stays short.}
    notes: Check the gap.
""",
        encoding="utf-8",
    )
    plain = tmp_path / "plain.pptx"
    canva = tmp_path / "canva.pptx"
    build(spec, plain)
    build(spec, canva, "--canva")
    assert check(plain).returncode == 0
    assert check(canva).returncode == 0

    def bullet_y(path: Path) -> float:
        prs = Presentation(str(path))
        for shape in prs.slides[0].shapes:
            if shape.name == "bullet-2":
                return shape.top / 914400
        raise AssertionError("bullet-2 missing")

    assert bullet_y(canva) - bullet_y(plain) > 0.15

    payload = tmp_path / "deck.json"
    payload.write_text(
        '{"title":"Json","slides":[{"layout":"cover","title":"From JSON","notes":"Built from JSON."}]}',
        encoding="utf-8",
    )
    build(payload, tmp_path / "json.pptx")

    bad_spec = tmp_path / "nope.yaml"
    bad_spec.write_text("slides:\n  - layout: poster\n    title: No\n", encoding="utf-8")
    result = run([PYTHON, "scripts/build_deck.py", str(bad_spec), "-o", str(tmp_path / "nope.pptx")])
    assert result.returncode != 0
    assert "unknown layout" in result.stderr


def test_contrast_and_link_helpers():
    assert check_deck.contrast_ratio("000000", "FFFFFF") > 20
    assert check_deck.contrast_ratio("FFFFFF", "FFFFFF") == pytest.approx(1)
    assert check_deck.contrast_ratio("1C1915", "F4F0E8") > 4.5
    assert check_deck.classify_http(200)[0] == "ok"
    assert check_deck.classify_http(404)[0] == "error"
    assert check_deck.classify_http(403)[0] == "warning"
    assert has_vietnamese("Mỗi ý")
    assert not has_vietnamese("Moi y")


@pytest.mark.skipif(
    shutil.which("soffice") is None and shutil.which("libreoffice") is None,
    reason="LibreOffice is not installed",
)
def test_render_preview(tmp_path: Path):
    if shutil.which("pdftoppm") is None:
        pytest.skip("pdftoppm is not installed")
    spec = tmp_path / "one.yaml"
    spec.write_text(
        """
slides:
  - layout: cover
    title: Preview this slide
    subtitle: One slide is enough to prove the render.
    notes: Look at the PNG.
""",
        encoding="utf-8",
    )
    pptx = tmp_path / "one.pptx"
    build(spec, pptx)
    out = tmp_path / "preview"
    result = run(["bash", "scripts/render_preview.sh", str(pptx), str(out)])
    assert result.returncode == 0, result.stderr + result.stdout
    assert (out / "slide-01.png").is_file()
    assert (out / "contact-sheet.png").is_file()
