"""Word 文档共用样式。"""

from __future__ import annotations

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor, Twips

FONT = "微软雅黑"
INK = "102433"
TEAL = "0F6E62"
GOLD = "B8893A"
SLATE = "1E2A32"
MUTED = "5C6B73"
CREAM = "F6F3EC"
MIST = "F3F6F5"
ROSE = "8C3E3E"
WHITE = "FFFFFF"


def _rfonts(r_pr):
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        r_fonts.set(qn(attr), FONT)
    lang = r_pr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        r_pr.append(lang)
    lang.set(qn("w:val"), "zh-CN")
    lang.set(qn("w:eastAsia"), "zh-CN")


def style_run(run, size, bold=False, color=SLATE, font=FONT):
    run.font.size = Pt(size)
    run.bold = bold
    run.font.name = font
    run.font.color.rgb = RGBColor.from_string(color)
    _rfonts(run._element.get_or_add_rPr())


def set_style_font(style, size, bold=False, color=SLATE, before=0, after=8, line=1.15):
    style.font.name = FONT
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.color.rgb = RGBColor.from_string(color)
    _rfonts(style.element.get_or_add_rPr())
    pf = style.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line


def shade_cell(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")


def set_cell_border(cell, color="E3DCD0"):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = tc_pr.find(qn("w:tcBorders"))
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tc_pr.append(tc_borders)
    for edge in ("top", "left", "bottom", "right"):
        el = tc_borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            tc_borders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)


def valign(cell, value="center"):
    tc_pr = cell._tc.get_or_add_tcPr()
    node = tc_pr.find(qn("w:vAlign"))
    if node is None:
        node = OxmlElement("w:vAlign")
        tc_pr.append(node)
    node.set(qn("w:val"), value)


def set_cell_text(cell, text, size=10.5, bold=False, color=SLATE, align="left"):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
    }[align]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.05
    parts = str(text).split("\n")
    for i, part in enumerate(parts):
        if i:
            p.add_run().add_break()
        run = p.add_run(part)
        style_run(run, size, bold, color)
    valign(cell)


def prevent_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def repeat_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:tblHeader")) is None:
        el = OxmlElement("w:tblHeader")
        el.set(qn("w:val"), "true")
        tr_pr.append(el)


def set_table_widths(table, widths):
    table.autofit = False
    table.allow_autofit = False
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    total = sum(int(w.cm * 567) for w in widths)
    tbl_w.set(qn("w:w"), str(total))
    tbl_w.set(qn("w:type"), "dxa")
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for child in list(grid):
            grid.remove(child)
    else:
        grid = OxmlElement("w:tblGrid")
        tbl_pr.addnext(grid)
    for w in widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(w.cm * 567)))
        grid.append(gc)
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            cell.width = widths[i]
            tc = cell._tc
            tc_pr = tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(int(widths[i].cm * 567)))
            tc_w.set(qn("w:type"), "dxa")


def add_page_field(paragraph):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    run._r.append(begin)
    run2 = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    run2._r.append(instr)
    style_run(run2, 9, color=MUTED)
    run3 = paragraph.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run3._r.append(end)


def new_doc(header_left, header_right, footer_left):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(1.8)
    section.header_distance = Cm(0.6)
    section.footer_distance = Cm(0.5)
    section.different_first_page_header_footer = True

    set_style_font(doc.styles["Normal"], 11, color=SLATE, after=8, line=1.15)
    set_style_font(doc.styles["Heading 1"], 16, bold=True, color=INK, before=16, after=8, line=1.1)
    set_style_font(doc.styles["Heading 2"], 13, bold=True, color=TEAL, before=12, after=6, line=1.1)
    set_style_font(doc.styles["Heading 3"], 12, bold=True, color=INK, before=8, after=4, line=1.1)

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    hp.paragraph_format.space_after = Pt(2)
    r = hp.add_run(header_left)
    style_run(r, 9, color=MUTED)
    r2 = hp.add_run("    " + header_right)
    style_run(r2, 9, bold=True, color=TEAL)
    p_pr = hp._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "6")
    bottom.set(qn("w:color"), TEAL)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    fr = fp.add_run(footer_left + "      ")
    style_run(fr, 9, color=MUTED)
    add_page_field(fp)

    core = doc.core_properties
    core.author = "专题分论坛筹备组"
    core.language = "zh-CN"
    return doc


def h1(doc, text):
    p = doc.add_heading(text, level=1)
    for run in p.runs:
        style_run(run, 16, True, INK)
    return p


def h2(doc, text):
    p = doc.add_heading(text, level=2)
    for run in p.runs:
        style_run(run, 13, True, TEAL)
    return p


def h3(doc, text):
    p = doc.add_heading(text, level=3)
    for run in p.runs:
        style_run(run, 12, True, INK)
    return p


def para(doc, text, size=11, bold=False, color=SLATE, after=8, before=0, align="left"):
    p = doc.add_paragraph()
    p.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }[align]
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(text)
    style_run(run, size, bold, color)
    return p


def bullets(doc, items, size=11):
    for item in items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.4)
        p.paragraph_format.first_line_indent = Cm(-0.35)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 1.12
        run = p.add_run("·  " + item)
        style_run(run, size, color=SLATE)
    return


def callout(doc, title, body, fill=CREAM, title_color=TEAL):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_widths(table, [Cm(17.4)])
    cell = table.cell(0, 0)
    shade_cell(cell, fill)
    set_cell_border(cell, "E3DCD0")
    cell.text = ""
    p1 = cell.paragraphs[0]
    p1.paragraph_format.space_after = Pt(2)
    r = p1.add_run(title)
    style_run(r, 11, True, title_color)
    lines = body.split("\n")
    for i, line in enumerate(lines):
        p2 = cell.add_paragraph()
        p2.paragraph_format.space_after = Pt(0 if i == len(lines) - 1 else 2)
        p2.paragraph_format.line_spacing = 1.12
        r2 = p2.add_run(line)
        style_run(r2, 10.5, False, SLATE)
    prevent_split(table.rows[0])
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def add_table(doc, headers, rows, widths, sizes=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    set_table_widths(table, widths)
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_text(cell, header, size=10, bold=True, color=WHITE)
        shade_cell(cell, INK)
        set_cell_border(cell, INK)
    repeat_header(table.rows[0])
    prevent_split(table.rows[0])
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            cell = table.rows[r + 1].cells[c]
            set_cell_text(cell, value, size=9.5 if sizes is None else sizes, color=SLATE)
            shade_cell(cell, WHITE if r % 2 else MIST)
            set_cell_border(cell, "E3DCD0")
        prevent_split(table.rows[r + 1])
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return table


def page_break(doc):
    doc.add_page_break()
