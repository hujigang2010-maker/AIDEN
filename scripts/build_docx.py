# -*- coding: utf-8 -*-
"""把18C估值手册原稿排成可直接阅读的 Word 文档。"""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Cm, Inches, Pt, RGBColor, Twips

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "source" / "HKEX_18C_分行业估值模板_2026-11-06.md"
OUTPUT = ROOT / "output" / "港交所18C分行业估值模板_实务手册.docx"

FONT = "微软雅黑"
INK = RGBColor(0x1C, 0x24, 0x21)
PINE = RGBColor(0x0E, 0x2A, 0x24)
TEAL = RGBColor(0x0F, 0x6E, 0x62)
COPPER = RGBColor(0x8A, 0x5A, 0x32)
MUTED = RGBColor(0x5E, 0x6A, 0x64)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

PAGE_WIDTH = Inches(6.70)


def set_run_font(run, size, color, bold=False, italic=False, ascii_font=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    ascii_name = ascii_font or FONT
    run.font.name = ascii_name
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), ascii_name)
    rFonts.set(qn("w:hAnsi"), ascii_name)
    rFonts.set(qn("w:cs"), ascii_name)
    rFonts.set(qn("w:eastAsia"), FONT)


def shade_run(run, fill):
    rPr = run._r.get_or_add_rPr()
    shd = rPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        rPr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def set_style_font(style, size, color, bold=False):
    style.font.name = FONT
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.color.rgb = color
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), FONT)


def set_paragraph_border(paragraph, color="C6A15B", sz="12"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)


def shade_paragraph(paragraph, fill):
    pPr = paragraph._p.get_or_add_pPr()
    shd = pPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        pPr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def left_bar(paragraph, color="C6A15B"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "16")
    left.set(qn("w:space"), "8")
    left.set(qn("w:color"), color)
    pBdr.append(left)


def set_cell_shading(cell, fill):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=40, bottom=40, left=70, right=70):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.find(qn("w:tcMar"))
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for edge, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = tcMar.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            tcMar.append(node)
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")


def set_cell_border(cell, color="D9D2C5"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge in ("top", "left", "bottom", "right"):
        node = tcBorders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            tcBorders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "4")
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), color)


def prevent_row_split(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    cant = OxmlElement("w:cantSplit")
    trPr.append(cant)


def mark_header_row(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    trPr.append(header)


def set_table_widths(table, widths):
    table.autofit = False
    table.allow_autofit = False
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    total = int(sum(widths))
    tblW.set(qn("w:w"), str(total))
    tblW.set(qn("w:type"), "dxa")
    grid = tbl.find(qn("w:tblGrid"))
    for idx, width in enumerate(widths):
        grid[idx].set(qn("w:w"), str(int(width)))
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            cell.width = Twips(int(widths[idx]))


def add_hyperlink(paragraph, text, url, size=10.5):
    r_id = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), FONT)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0F6E62")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(int(size * 2)))
    rPr.extend([rFonts, color, underline, sz, szCs])
    new_run.append(rPr)
    text_el = OxmlElement("w:t")
    text_el.set(qn("xml:space"), "preserve")
    text_el.text = text
    new_run.append(text_el)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def tokenize(text):
    tokens = []
    i = 0
    n = len(text)
    while i < n:
        if text.startswith("**", i):
            j = text.find("**", i + 2)
            if j != -1:
                tokens.append(("bold", text[i + 2 : j]))
                i = j + 2
                continue
        if text.startswith("`", i):
            j = text.find("`", i + 1)
            if j != -1:
                tokens.append(("code", text[i + 1 : j]))
                i = j + 1
                continue
        if text.startswith("[", i):
            match = re.match(r"\[([^\]]+)\]\(([^)]+)\)", text[i:])
            if match:
                tokens.append(("link", match.group(1), match.group(2)))
                i += match.end()
                continue
        if text.startswith("<http://", i) or text.startswith("<https://", i):
            j = text.find(">", i)
            if j != -1:
                url = text[i + 1 : j]
                tokens.append(("link", url, url))
                i = j + 1
                continue
        if text.startswith("*", i):
            j = text.find("*", i + 1)
            if j != -1 and "\n" not in text[i + 1 : j] and not text.startswith("**", i):
                tokens.append(("italic", text[i + 1 : j]))
                i = j + 1
                continue
        j = i + 1
        while j < n:
            if text.startswith(("**", "`", "[", "<http://", "<https://", "*"), j):
                break
            j += 1
        tokens.append(("text", text[i:j]))
        i = j
    return tokens


def write_inline(paragraph, text, size=10.5, color=INK, bold=False):
    for token in tokenize(text):
        kind = token[0]
        if kind == "text":
            run = paragraph.add_run(token[1])
            set_run_font(run, size, color, bold=bold)
        elif kind == "bold":
            run = paragraph.add_run(token[1])
            set_run_font(run, size, color, bold=True)
        elif kind == "italic":
            run = paragraph.add_run(token[1])
            set_run_font(run, size, color, italic=True)
        elif kind == "code":
            run = paragraph.add_run(token[1])
            set_run_font(run, max(size - 0.5, 8), PINE, ascii_font="Consolas")
            shade_run(run, "F3EFE6")
        elif kind == "link":
            label, url = token[1], token[2]
            add_hyperlink(paragraph, label, url, size=size)


def add_body(doc, text, size=10.5, space_after=6, space_before=0, italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.line_spacing = 1.15
    write_inline(p, text, size=size, bold=False)
    if italic:
        for run in p.runs:
            run.italic = True
    return p


def add_quote(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.left_indent = Cm(0.25)
    p.paragraph_format.right_indent = Cm(0.15)
    shade_paragraph(p, "F7F4EE")
    left_bar(p, "C6A15B")
    write_inline(p, text, size=10.5, color=INK)
    return p


def add_list(doc, text, numbered=False):
    style = "List Number" if numbered else "List Bullet"
    p = doc.add_paragraph(style=style)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.line_spacing = 1.12
    for child in list(p._p):
        if child.tag == qn("w:r"):
            p._p.remove(child)
    write_inline(p, text, size=10.5)
    return p


def add_code(doc, lines):
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    set_cell_shading(cell, "F3EFE6")
    set_cell_border(cell, "E4D8C4")
    set_cell_margins(cell, 60, 60, 100, 100)
    set_table_widths(table, [int(PAGE_WIDTH.twips)])
    cell.text = ""
    for idx, line in enumerate(lines):
        paragraph = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.05
        run = paragraph.add_run(line if line else " ")
        set_run_font(run, 8.5, PINE)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def split_table_row(line):
    raw = line.strip()
    if raw.startswith("|"):
        raw = raw[1:]
    if raw.endswith("|"):
        raw = raw[:-1]
    return [part.strip() for part in raw.split("|")]


def is_separator(cells):
    if not cells:
        return False
    return all(re.fullmatch(r":?-{3,}:?", cell.replace(" ", "")) for cell in cells)


def add_table(doc, rows):
    if not rows:
        return
    cols = max(len(row) for row in rows)
    normalized = [row + [""] * (cols - len(row)) for row in rows]
    table = doc.add_table(rows=len(normalized), cols=cols)
    table.style = "Table Grid"
    usable = int(PAGE_WIDTH.twips)
    if cols >= 8:
        first = int(usable * 0.14)
    elif cols >= 6:
        first = int(usable * 0.16)
    elif cols == 2:
        first = int(usable * 0.28)
    else:
        first = int(usable * 0.22)
    rest = (usable - first) // max(cols - 1, 1)
    widths = [first] + [rest] * (cols - 1)
    widths[-1] = usable - sum(widths[:-1])
    font_size = 8 if cols >= 7 else 8.5 if cols >= 5 else 9
    for r_idx, row_data in enumerate(normalized):
        row = table.rows[r_idx]
        prevent_row_split(row)
        if r_idx == 0:
            mark_header_row(row)
        for c_idx, text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = ""
            set_cell_margins(cell)
            set_cell_border(cell, "E4DCCF" if r_idx else "0E2A24")
            if r_idx == 0:
                set_cell_shading(cell, "0E2A24")
                color = WHITE
                bold = True
            else:
                set_cell_shading(cell, "F7F4EE" if r_idx % 2 == 0 else "FFFFFF")
                color = INK
                bold = c_idx == 0
            paragraph = cell.paragraphs[0]
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.05
            write_inline(paragraph, text, size=font_size, color=color, bold=bold)
    set_table_widths(table, widths)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(6)
    spacer.paragraph_format.space_before = Pt(2)


def parse_blocks(text):
    lines = text.replace("\r\n", "\n").split("\n")
    blocks = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if not stripped:
            i += 1
            continue
        if stripped.startswith("```"):
            i += 1
            code = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code.append(lines[i].rstrip())
                i += 1
            blocks.append(("code", code))
            i += 1
            continue
        if stripped.startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            rows = []
            for raw in table_lines:
                cells = split_table_row(raw)
                if is_separator(cells):
                    continue
                rows.append(cells)
            blocks.append(("table", rows))
            continue
        if stripped.startswith(">"):
            quote = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                quote.append(re.sub(r"^>\s?", "", lines[i].strip()))
                i += 1
            blocks.append(("quote", " ".join(part for part in quote if part)))
            continue
        if stripped in ("---", "***"):
            blocks.append(("rule", None))
            i += 1
            continue
        if stripped.startswith("#"):
            level = len(stripped) - len(stripped.lstrip("#"))
            blocks.append(("heading", level, stripped[level:].strip()))
            i += 1
            continue
        if re.match(r"^[-*]\s+", stripped):
            blocks.append(("bullet", re.sub(r"^[-*]\s+", "", stripped)))
            i += 1
            continue
        if re.match(r"^\d+\.\s+", stripped):
            blocks.append(("number", re.sub(r"^\d+\.\s+", "", stripped)))
            i += 1
            continue
        para = [stripped]
        i += 1
        while i < len(lines):
            nxt = lines[i].strip()
            if not nxt or nxt.startswith(("#", ">", "|", "```", "- ", "* ")) or nxt in ("---", "***"):
                break
            if re.match(r"^\d+\.\s+", nxt):
                break
            para.append(nxt)
            i += 1
        blocks.append(("para", " ".join(para)))
    return blocks


def configure_styles(doc):
    normal = doc.styles["Normal"]
    set_style_font(normal, 10.5, INK)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.15
    pf = normal.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE

    for style_name, size, color, before, after, bold in (
        ("Heading 1", 18, PINE, 18, 6, True),
        ("Heading 2", 14, RGBColor(0x1A, 0x43, 0x3A), 14, 4, True),
        ("Heading 3", 12, COPPER, 12, 3, True),
    ):
        style = doc.styles[style_name]
        set_style_font(style, size, color, bold=bold)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.line_spacing = 1.1

    for style_name in ("List Bullet", "List Number"):
        style = doc.styles[style_name]
        set_style_font(style, 10.5, INK)
        style.paragraph_format.space_after = Pt(2)


def add_page_field(paragraph):
    run = paragraph.add_run()
    set_run_font(run, 8, MUTED)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)


def add_toc(paragraph):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\o "1-2" \\h \\z \\u '
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(separate)
    placeholder = paragraph.add_run("在 Word 中右键目录，选择“更新域”，即可生成页码。")
    set_run_font(placeholder, 10, MUTED, italic=True)
    end_run = paragraph.add_run()
    end_run._r.append(end)


def banner(doc, lines, fill="0E2A24"):
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    set_cell_shading(cell, fill)
    set_cell_margins(cell, 160, 160, 180, 180)
    set_table_widths(table, [int(PAGE_WIDTH.twips)])
    for edge in ("top", "left", "bottom", "right"):
        set_cell_border(cell, fill)
    cell.text = ""
    for idx, (text, size, color, bold) in enumerate(lines):
        paragraph = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(2)
        run = paragraph.add_run(text)
        set_run_font(run, size, color, bold=bold)
    return table


def configure_section(section):
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.9)
    section.bottom_margin = Cm(1.7)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(1.8)
    section.header_distance = Cm(0.6)
    section.footer_distance = Cm(0.5)
    section.different_first_page_header_footer = True

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    hp.paragraph_format.space_after = Pt(2)
    set_paragraph_border(hp, "C6A15B", "8")
    left = hp.add_run("港交所18C  ·  分行业估值模板与模型回调实务手册")
    set_run_font(left, 8, PINE, bold=True)
    right = hp.add_run("          V1.0  ·  基准日 2026-11-06")
    set_run_font(right, 8, MUTED)

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = fp.add_run("虚构演示与方法模板  ·  非投资建议、非独立评估、非法律意见")
    set_run_font(run, 8, MUTED)
    gap = fp.add_run("          ")
    set_run_font(gap, 8, MUTED)
    add_page_field(fp)

    first_footer = section.first_page_footer
    first_footer.is_linked_to_previous = False
    first_footer.paragraphs[0].text = ""


def enable_update_fields(doc):
    settings = doc.settings.element
    node = settings.find(qn("w:updateFields"))
    if node is None:
        node = OxmlElement("w:updateFields")
        settings.append(node)
    node.set(qn("w:val"), "true")


def build():
    blocks = parse_blocks(SOURCE.read_text(encoding="utf-8"))
    title = "港交所18C特专科技企业：分行业估值模板与模型回调实务手册"
    cover_quotes = []
    body = []
    seen_h2 = False
    for block in blocks:
        if not seen_h2 and block[0] == "heading" and block[1] == 1:
            title = block[2]
            continue
        if not seen_h2 and block[0] == "quote":
            cover_quotes.append(block[1])
            continue
        if block[0] == "heading" and block[1] <= 2:
            seen_h2 = True
        if seen_h2:
            body.append(block)

    doc = Document()
    configure_styles(doc)
    configure_section(doc.sections[0])
    enable_update_fields(doc)
    core = doc.core_properties
    core.title = "港交所18C分行业估值模板与模型回调实务手册"
    core.subject = "方法模板与虚构演示，非投资建议"
    core.category = "估值方法手册"
    core.comments = "配套PPT简报与Excel可计算模型。演示数字不是行业基准。"

    banner(
        doc,
        [
            ("HONG KONG EXCHANGES  ·  CHAPTER 18C", 9, RGBColor(0xC6, 0xA1, 0x5B), False),
            ("分行业估值模板", 26, WHITE, True),
            ("与模型回调实务手册", 18, RGBColor(0xF3, 0xEB, 0xD8), False),
        ],
    )
    meta = doc.add_paragraph()
    meta.paragraph_format.space_before = Pt(10)
    meta.paragraph_format.space_after = Pt(4)
    write_inline(meta, "版本 V1.0　　研究基准日 2026-11-06　　默认人民币测算，上市规则测试用港币", size=10.5, color=PINE)

    for quote in cover_quotes:
        add_quote(doc, quote)

    note = add_body(
        doc,
        "本 Word 保留原稿的规则、公式、分行业模块、回调流程和公开来源。配套 PPT 是投委会简报；配套 Excel 把其中可以落地的演示算例做成了可改参数的公式，蓝字是输入，黑字不要手改。",
        size=10.5,
    )
    note.paragraph_format.space_before = Pt(4)

    doc.add_page_break()
    toc_title = doc.add_paragraph(style="Heading 1")
    toc_title.add_run("目录")
    set_paragraph_border(toc_title, "C6A15B", "10")
    toc = doc.add_paragraph()
    add_toc(toc)

    for block in body:
        kind = block[0]
        if kind == "heading":
            level = min(block[1], 3)
            paragraph = doc.add_paragraph(style=f"Heading {level}")
            write_inline(paragraph, block[2], size={1: 18, 2: 14, 3: 12}[level], color={1: PINE, 2: RGBColor(0x1A, 0x43, 0x3A), 3: COPPER}[level], bold=True)
            if level == 1:
                set_paragraph_border(paragraph, "C6A15B", "10")
        elif kind == "para":
            add_body(doc, block[1])
        elif kind == "quote":
            add_quote(doc, block[1])
        elif kind == "bullet":
            add_list(doc, block[1], numbered=False)
        elif kind == "number":
            add_list(doc, block[1], numbered=True)
        elif kind == "code":
            add_code(doc, block[1])
        elif kind == "table":
            add_table(doc, block[1])
        elif kind == "rule":
            spacer = doc.add_paragraph()
            spacer.paragraph_format.space_before = Pt(2)
            spacer.paragraph_format.space_after = Pt(2)

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT)
    return OUTPUT


if __name__ == "__main__":
    path = build()
    print(path)
