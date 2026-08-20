# -*- coding: utf-8 -*-
"""16:9 PPT 共用绘制函数。"""
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

PRIMARY = RGBColor(0x0B, 0x3D, 0x5C)
ACCENT = RGBColor(0x1A, 0x7A, 0x6D)
LIGHT = RGBColor(0xE8, 0xF3, 0xF1)
SOFT = RGBColor(0xD0, 0xE8, 0xE4)
GOLD = RGBColor(0xC4, 0xA3, 0x5A)
DARK = RGBColor(0x1A, 0x2A, 0x33)
GREY = RGBColor(0x5A, 0x6A, 0x72)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG_DEEP = RGBColor(0x08, 0x2E, 0x45)
ORANGE = RGBColor(0xC0, 0x6A, 0x2F)


def set_font(run, name="微软雅黑", size=18, bold=False, color=DARK):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("a:ea"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("a:ea"), {"typeface": name})
        rPr.append(rFonts)
    else:
        rFonts.set("typeface", name)


def add_rect(slide, left, top, width, height, fill=PRIMARY, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
    shape.shadow.inherit = False
    return shape


def add_rounded(slide, left, top, width, height, fill=LIGHT):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    shape.shadow.inherit = False
    try:
        shape.adjustments[0] = 0.08
    except Exception:
        pass
    return shape


def add_text(
    slide,
    left,
    top,
    width,
    height,
    text,
    size=18,
    bold=False,
    color=DARK,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    font="微软雅黑",
):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    if isinstance(text, str):
        text = [text]
    for i, t in enumerate(text):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = t
        set_font(run, name=font, size=size, bold=bold, color=color)
    return tb


def add_bullet_list(slide, left, top, width, height, items, size=14, color=DARK, bullet="•"):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.2
        run = p.add_run()
        run.text = f"{bullet}  {it}"
        set_font(run, size=size, color=color)
    return tb


def slide_header(slide, title, subtitle=None, page_no=None, total=None):
    add_rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.42), fill=PRIMARY)
    add_rect(slide, Inches(0), Inches(0.42), Inches(13.333), Inches(0.05), fill=ACCENT)
    add_text(
        slide,
        Inches(0.5),
        Inches(0.55),
        Inches(10.8),
        Inches(0.45),
        title,
        size=22,
        bold=True,
        color=PRIMARY,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    if subtitle:
        add_text(
            slide,
            Inches(0.5),
            Inches(1.0),
            Inches(12.3),
            Inches(0.28),
            subtitle,
            size=12,
            color=GREY,
            anchor=MSO_ANCHOR.MIDDLE,
        )
    if page_no and total:
        add_text(
            slide,
            Inches(11.5),
            Inches(7.05),
            Inches(1.5),
            Inches(0.3),
            f"{page_no} / {total}",
            size=10,
            color=GREY,
            align=PP_ALIGN.RIGHT,
        )
    add_rect(slide, Inches(0), Inches(7.35), Inches(13.333), Inches(0.15), fill=PRIMARY)


def make_card(slide, left, top, width, height, title, body_items, accent=ACCENT, title_size=14, body_size=11):
    add_rounded(slide, left, top, width, height, fill=LIGHT)
    add_rect(slide, left, top, Inches(0.08), height, fill=accent)
    add_text(
        slide,
        left + Inches(0.2),
        top + Inches(0.1),
        width - Inches(0.3),
        Inches(0.32),
        title,
        size=title_size,
        bold=True,
        color=accent,
    )
    add_bullet_list(
        slide,
        left + Inches(0.18),
        top + Inches(0.44),
        width - Inches(0.28),
        height - Inches(0.5),
        body_items,
        size=body_size,
        color=DARK,
    )


def make_table(slide, left, top, width, height, headers, rows, header_color=PRIMARY, font_size=10):
    shape = slide.shapes.add_table(len(rows) + 1, len(headers), left, top, width, height)
    table = shape.table
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_color
        cell.margin_left = Inches(0.05)
        cell.margin_right = Inches(0.05)
        cell.margin_top = Inches(0.03)
        cell.margin_bottom = Inches(0.03)
        tf = cell.text_frame
        tf.text = ""
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h
        set_font(r, size=font_size + 1, bold=True, color=WHITE)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = table.cell(ri + 1, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if ri % 2 == 0 else LIGHT
            cell.margin_left = Inches(0.05)
            cell.margin_right = Inches(0.05)
            cell.margin_top = Inches(0.02)
            cell.margin_bottom = Inches(0.02)
            tf = cell.text_frame
            tf.text = ""
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            r = p.add_run()
            r.text = str(val)
            set_font(r, size=font_size, color=DARK)
    return table
