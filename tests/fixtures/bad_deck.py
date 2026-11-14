"""Build a deck that should fail the checker in obvious ways."""

from __future__ import annotations

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt


def _box(slide, name, x, y, w, h, text, size, color="000000", fill=None, url=None):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    shape.name = name
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(fill)
    frame = shape.text_frame
    frame.word_wrap = True
    frame.margin_left = Inches(0.02)
    frame.margin_right = Inches(0.02)
    frame.margin_top = Inches(0.02)
    frame.margin_bottom = Inches(0.02)
    paragraph = frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    run.font.name = "Calibri"
    if url:
        run.hyperlink.address = url
        run.font.size = Pt(size)
        run.font.color.rgb = RGBColor.from_string(color)
        run.font.name = "Calibri"
    return shape


def _group(slide, shapes):
    left = min(int(shape.left) for shape in shapes)
    top = min(int(shape.top) for shape in shapes)
    right = max(int(shape.left) + int(shape.width) for shape in shapes)
    bottom = max(int(shape.top) + int(shape.height) for shape in shapes)
    grp = etree.Element(qn("p:grpSp"))
    nv = etree.SubElement(grp, qn("p:nvGrpSpPr"))
    c_nv = etree.SubElement(nv, qn("p:cNvPr"))
    c_nv.set("id", "900")
    c_nv.set("name", "Bad group")
    etree.SubElement(nv, qn("p:cNvGrpSpPr"))
    etree.SubElement(nv, qn("p:nvPr"))
    props = etree.SubElement(grp, qn("p:grpSpPr"))
    xfrm = etree.SubElement(props, qn("a:xfrm"))
    off = etree.SubElement(xfrm, qn("a:off"))
    off.set("x", str(left))
    off.set("y", str(top))
    ext = etree.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", str(right - left))
    ext.set("cy", str(bottom - top))
    ch_off = etree.SubElement(xfrm, qn("a:chOff"))
    ch_off.set("x", str(left))
    ch_off.set("y", str(top))
    ch_ext = etree.SubElement(xfrm, qn("a:chExt"))
    ch_ext.set("cx", str(right - left))
    ch_ext.set("cy", str(bottom - top))
    tree = slide.shapes._spTree
    for shape in shapes:
        element = shape._element
        element.getparent().remove(element)
        grp.append(element)
    tree.append(grp)


def build(path) -> None:
    prs = Presentation()
    prs.slide_width = Emu(12_192_000)
    prs.slide_height = Emu(6_858_000)
    blank = prs.slide_layouts[-1]
    for layout in prs.slide_layouts:
        if layout.name.strip().lower() == "blank":
            blank = layout
            break
    slide = prs.slides.add_slide(blank)
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)

    _box(slide, "tiny", 0.35, 0.2, 7.3, 0.45, "Too small to read from the back row.", 12)
    _box(
        slide, "body", 0.35, 0.75, 7.3, 0.9,
        "We tracked 3 goals, 8 teams, 12 weeks, and 40 risks.", 28,
    )
    short_y = 1.8
    for index in range(5):
        _box(slide, f"bullet-{index + 1}", 0.35, short_y, 7.3, 0.55, "State the point.", 28)
        short_y += 0.62
    long = (
        "Please say this with far too many words so that the line becomes a very "
        "long bullet which the checker must flag for length every time"
    )
    _box(slide, "bullet-6", 0.35, short_y + 0.05, 7.3, 2.2, long, 24)

    _box(slide, "overlap-a", 8.15, 0.25, 4.4, 1.05, "First block of overlapping text.", 28)
    _box(slide, "overlap-b", 8.15, 0.25, 4.4, 1.05, "Second block of overlapping text.", 28)
    _box(slide, "hidden", 8.15, 1.5, 4.4, 0.7, "Same color as the fill.", 28, color="FFFFFF", fill="FFFFFF")
    _box(slide, "cramped", 8.15, 2.4, 4.4, 0.24, "This line cannot fit in such a short box.", 28)
    _box(slide, "source-1", 8.15, 2.85, 4.2, 0.55, "Details", 24, url="notaurl")
    _box(slide, "offslide", 12.9, 3.6, 1.5, 0.55, "Past the edge.", 24)

    rect_a = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.15), Inches(4.4), Inches(1.2), Inches(0.7))
    rect_b = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(9.55), Inches(4.4), Inches(1.2), Inches(0.7))
    for rect in (rect_a, rect_b):
        rect.line.fill.background()
        rect.fill.solid()
        rect.fill.fore_color.rgb = RGBColor(0x1C, 0x19, 0x15)
    _group(slide, [rect_a, rect_b])
    prs.save(path)
