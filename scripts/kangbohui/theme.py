"""康博会分论坛文档的共用版式。"""

from __future__ import annotations

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

FONT = "微软雅黑"

INK = RGBColor(0x10, 0x24, 0x33)
TEAL = RGBColor(0x0F, 0x6E, 0x62)
TEAL_DEEP = RGBColor(0x0A, 0x4A, 0x44)
GOLD = RGBColor(0xB8, 0x89, 0x3A)
CREAM = RGBColor(0xF6, 0xF3, 0xEC)
PAPER = RGBColor(0xFF, 0xFF, 0xFF)
SLATE = RGBColor(0x1E, 0x2A, 0x32)
MUTED = RGBColor(0x5C, 0x6B, 0x73)
LINE = RGBColor(0xE3, 0xDC, 0xD0)
MIST = RGBColor(0xF3, 0xF6, 0xF5)
CARD_TEAL = RGBColor(0xE7, 0xF2, 0xF0)
CARD_GOLD = RGBColor(0xF8, 0xF1, 0xE3)
ROSE = RGBColor(0x8C, 0x3E, 0x3E)
ROSE_SOFT = RGBColor(0xF8, 0xEE, 0xEE)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SW = 13.333
SH = 7.5


def hex_of(color: RGBColor) -> str:
    return f"{color[0]:02X}{color[1]:02X}{color[2]:02X}"


def set_run_font(run, size, bold, color, font=FONT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    r_pr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        el = r_pr.find(qn(f"a:{tag}"))
        if el is None:
            el = etree.SubElement(r_pr, qn(f"a:{tag}"))
        el.set("typeface", font)


def no_shadow(shape):
    element = shape._element
    sp_pr = getattr(element, "spPr", None)
    if sp_pr is None:
        return
    effect = sp_pr.find(qn("a:effectLst"))
    if effect is not None:
        sp_pr.remove(effect)
    etree.SubElement(sp_pr, qn("a:effectLst"))


def add_shape(slide, kind, x, y, w, h, fill, line=None, radius=None):
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(0.75)
    if radius is not None and kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            shape.adjustments[0] = radius
        except Exception:
            pass
    no_shadow(shape)
    return shape


def rect(slide, x, y, w, h, fill, line=None):
    return add_shape(slide, MSO_SHAPE.RECTANGLE, x, y, w, h, fill, line)


def rrect(slide, x, y, w, h, fill, line=None, radius=0.08):
    return add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill, line, radius)


def add_text(
    slide,
    text,
    x,
    y,
    w,
    h,
    size=14,
    bold=False,
    color=SLATE,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    spacing=1.0,
):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.vertical_anchor = anchor
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    lines = text.split("\n") if isinstance(text, str) else list(text)
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        if spacing and spacing != 1:
            p.line_spacing = spacing
        run = p.add_run()
        run.text = line
        set_run_font(run, size, bold, color)
    return box


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def footer(slide, page, total, internal=True):
    rect(slide, 0, 7.18, SW, 0.32, CREAM)
    label = "内部执行稿 · 不对外发送" if internal else "对外拟议版 · 权益以正式合同为准"
    add_text(
        slide,
        "2026上海国际智慧康养产业博览会 · 专题分论坛",
        0.4,
        7.2,
        6.3,
        0.26,
        size=10,
        color=MUTED,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    add_text(
        slide,
        label,
        6.7,
        7.2,
        4.35,
        0.26,
        size=10,
        color=MUTED,
        align=PP_ALIGN.RIGHT,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    add_text(
        slide,
        f"{page:02d}  /  {total:02d}",
        11.15,
        7.2,
        1.75,
        0.26,
        size=10,
        bold=True,
        color=TEAL,
        align=PP_ALIGN.RIGHT,
        anchor=MSO_ANCHOR.MIDDLE,
    )


def content_header(slide, kicker, title, page, total, internal=True):
    rect(slide, 0, 0, SW, SH, PAPER)
    rect(slide, 0, 0, SW, 0.08, TEAL)
    add_text(slide, kicker, 0.42, 0.18, 12.4, 0.26, size=12, bold=True, color=TEAL)
    add_text(slide, title, 0.4, 0.42, 12.5, 0.46, size=26, bold=True, color=INK)
    rect(slide, 0.42, 0.96, 1.05, 0.045, GOLD)
    footer(slide, page, total, internal)


def set_cell_fill(cell, color: RGBColor):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    for child in list(tc_pr):
        if child.tag == qn("a:solidFill"):
            tc_pr.remove(child)
    solid = etree.SubElement(tc_pr, qn("a:solidFill"))
    srgb = etree.SubElement(solid, qn("a:srgbClr"))
    srgb.set("val", hex_of(color))


def set_cell_borders(cell, color_hex="E3DCD0"):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    for edge in ("lnL", "lnR", "lnT", "lnB"):
        ln = etree.SubElement(tc_pr, qn(f"a:{edge}"))
        ln.set("w", "6350")
        solid = etree.SubElement(ln, qn("a:solidFill"))
        srgb = etree.SubElement(solid, qn("a:srgbClr"))
        srgb.set("val", color_hex)


def shade_cell_text(cell, text, size, bold, color):
    cell.text = ""
    cell.margin_left = Inches(0.08)
    cell.margin_right = Inches(0.06)
    cell.margin_top = Inches(0.04)
    cell.margin_bottom = Inches(0.03)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = str(text)
    set_run_font(run, size, bold, color)


def add_table(slide, x, y, w, h, data, col_widths, size=12, header=True):
    rows = len(data)
    cols = len(data[0])
    shape = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(h))
    table = shape.table
    for i, cw in enumerate(col_widths):
        table.columns[i].width = Inches(cw)
    for r in range(rows):
        for c in range(cols):
            cell = table.cell(r, c)
            is_header = header and r == 0
            fill = INK if is_header else (PAPER if r % 2 else MIST)
            font_color = WHITE if is_header else SLATE
            shade_cell_text(cell, data[r][c], size, is_header, font_color)
            set_cell_fill(cell, fill)
            set_cell_borders(cell)
    no_shadow(shape)
    return shape


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text
