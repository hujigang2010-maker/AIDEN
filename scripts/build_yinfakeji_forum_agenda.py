#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成《AI重塑养老未来》分论坛策划议程 Word（V4.0）。

供组委会内部邀约及招商使用。候选名单不代表已经同意出席、主办或赞助。
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# --- 版式 ---

CONTENT_W = 17.6
NAVY = "0C2340"
NAVY_MID = "16324F"
TEAL = "1A6B73"
GOLD = "A68456"
INK = "1C2833"
MUTED = "5E6A72"
LINE = "E4DFD6"
ZEBRA = "F7F5F1"
WHITE = "FFFFFF"
CREAM = "FBF8F3"
WARN_BG = "F8F3E8"
INFO_BG = "F2F6F6"
KEY_BG = "F4F7FB"

ROOT = Path(__file__).resolve().parents[1]
OUT_PATH = ROOT / "output" / "2026银发科技与具身智能产业创新峰会-策划议程-V4.0.docx"


def rgb(value: str) -> RGBColor:
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def _rfonts(rpr, east_asia: str, ascii_font: str) -> None:
    r_fonts = rpr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        rpr.append(r_fonts)
    r_fonts.set(qn("w:ascii"), ascii_font)
    r_fonts.set(qn("w:hAnsi"), ascii_font)
    r_fonts.set(qn("w:cs"), ascii_font)
    r_fonts.set(qn("w:eastAsia"), east_asia)
    r_fonts.set(qn("w:hint"), "eastAsia")


def set_run_font(
    run,
    east_asia: str,
    size_pt: float,
    *,
    bold: bool = False,
    color: str | None = None,
    ascii_font: str = "Calibri",
) -> None:
    run.font.name = ascii_font
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = False
    if color:
        run.font.color.rgb = rgb(color)
    rpr = run._element.get_or_add_rPr()
    _rfonts(rpr, east_asia, ascii_font)
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:eastAsia"), "zh-CN")


def no_snap(paragraph) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    snap = ppr.find(qn("w:snapToGrid"))
    if snap is None:
        snap = OxmlElement("w:snapToGrid")
        ppr.append(snap)
    snap.set(qn("w:val"), "0")


def set_paragraph_spacing(paragraph, before=0, after=6, line=1.15, align="left") -> None:
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    paragraph.alignment = align_map[align]
    no_snap(paragraph)


def add_runs(paragraph, parts, *, size=11, font="宋体", color=INK, bold=False) -> None:
    for part in parts:
        if isinstance(part, str):
            text, opts = part, {}
        else:
            text, opts = part
        run = paragraph.add_run(text)
        set_run_font(
            run,
            opts.get("font", font),
            opts.get("size", size),
            bold=opts.get("bold", bold),
            color=opts.get("color", color),
        )


def add_para(doc, parts, *, size=11, before=0, after=6, line=1.15, align="left", font="宋体", color=INK):
    paragraph = doc.add_paragraph()
    set_paragraph_spacing(paragraph, before, after, line, align)
    paragraph.paragraph_format.widow_control = True
    add_runs(paragraph, parts if isinstance(parts, list) else [parts], size=size, font=font, color=color)
    return paragraph


def add_body(doc, text: str):
    return add_para(doc, text, size=11, after=6, align="justify")


def _decorate_heading(paragraph, *, gold=True) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    existing = ppr.find(qn("w:pBdr"))
    if existing is not None:
        ppr.remove(existing)
    borders = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "16")
    left.set(qn("w:space"), "6")
    left.set(qn("w:color"), GOLD if gold else TEAL)
    borders.append(left)
    ppr.append(borders)


def add_h1(doc, text: str, *, new_page=False):
    paragraph = doc.add_heading(text, level=1)
    set_paragraph_spacing(paragraph, 2 if new_page else 15, 8, 1.08)
    paragraph.paragraph_format.page_break_before = new_page
    paragraph.paragraph_format.keep_with_next = True
    paragraph.paragraph_format.left_indent = Cm(0.22)
    paragraph.paragraph_format.widow_control = True
    for run in paragraph.runs:
        set_run_font(run, "黑体", 16, bold=True, color=NAVY)
    _decorate_heading(paragraph, gold=True)
    return paragraph


def add_h2(doc, text: str):
    paragraph = doc.add_heading(text, level=2)
    set_paragraph_spacing(paragraph, 11, 4, 1.05)
    paragraph.paragraph_format.keep_with_next = True
    paragraph.paragraph_format.widow_control = True
    for run in paragraph.runs:
        set_run_font(run, "微软雅黑", 13, bold=True, color=TEAL)
    return paragraph


def add_h3(doc, text: str):
    paragraph = doc.add_heading(text, level=3)
    set_paragraph_spacing(paragraph, 8, 3, 1.05)
    paragraph.paragraph_format.keep_with_next = True
    for run in paragraph.runs:
        set_run_font(run, "微软雅黑", 12, bold=True, color=NAVY)
    return paragraph


def add_items(doc, items, *, numbered=False, size=11):
    for index, item in enumerate(items, start=1):
        paragraph = doc.add_paragraph()
        set_paragraph_spacing(paragraph, 1, 3, 1.12)
        paragraph.paragraph_format.left_indent = Cm(0.75)
        paragraph.paragraph_format.first_line_indent = Cm(-0.45)
        if numbered:
            marker, marker_color, marker_font = f"{index}. ", TEAL, "Calibri"
        else:
            marker, marker_color, marker_font = "•  ", TEAL, "Calibri"
        run = paragraph.add_run(marker)
        set_run_font(run, "微软雅黑", size, bold=True, color=marker_color, ascii_font=marker_font)
        if isinstance(item, tuple):
            lead, rest = item
            lead_run = paragraph.add_run(lead)
            set_run_font(lead_run, "微软雅黑", size, bold=True, color=NAVY)
            rest_run = paragraph.add_run(rest)
            set_run_font(rest_run, "宋体", size, color=INK)
        else:
            rest_run = paragraph.add_run(item)
            set_run_font(rest_run, "宋体", size, color=INK)


def add_caption(doc, text: str):
    paragraph = add_para(doc, text, size=9, before=8, after=3, font="微软雅黑", color=MUTED)
    paragraph.paragraph_format.keep_with_next = True
    return paragraph


def shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    for old in tc_pr.findall(qn("w:shd")):
        tc_pr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_borders(cell, edges: dict) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    existing = tc_pr.find(qn("w:tcBorders"))
    if existing is not None:
        tc_pr.remove(existing)
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        val, sz, color = edges.get(edge, ("nil", "0", "auto"))
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), val)
        element.set(qn("w:sz"), sz)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)
        borders.append(element)
    tc_pr.append(borders)


def set_cell_margin(cell, top=60, bottom=60, left=100, right=100) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    existing = tc_pr.find(qn("w:tcMar"))
    if existing is not None:
        tc_pr.remove(existing)
    margin = OxmlElement("w:tcMar")
    for name, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        margin.append(node)
    tc_pr.append(margin)


def set_table_widths(table, widths: list[float]) -> None:
    if abs(sum(widths) - CONTENT_W) > 0.08:
        raise ValueError(f"表宽 {sum(widths):.2f} 与版心 {CONTENT_W} 不一致")
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(int(Cm(sum(widths)).twips)))
    tbl_w.set(qn("w:type"), "dxa")

    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")

    tbl_ind = tbl_pr.find(qn("w:tblInd"))
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), "0")
    tbl_ind.set(qn("w:type"), "dxa")

    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), LINE)

    grid = tbl.find(qn("w:tblGrid"))
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(int(Cm(width).twips)))
        grid.append(col)

    for row in table.rows:
        for index, cell in enumerate(row.cells):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(int(Cm(widths[index]).twips)))
            tc_w.set(qn("w:type"), "dxa")


def keep_short_table_together(table) -> None:
    """短表整表换页，避免页尾只剩表头和一行。"""
    last = len(table.rows) - 1
    for row_index, row in enumerate(table.rows):
        if row_index == last:
            break
        seen = set()
        for cell in row.cells:
            marker = id(cell._tc)
            if marker in seen:
                continue
            seen.add(marker)
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.keep_with_next = True


def mark_header_row(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    header = tr_pr.find(qn("w:tblHeader"))
    if header is None:
        header = OxmlElement("w:tblHeader")
        tr_pr.append(header)
    header.set(qn("w:val"), "true")


def prevent_row_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def set_cell_lines(cell, lines: list[dict], *, align="left", center=False) -> None:
    for index, spec in enumerate(lines):
        paragraph = cell.paragraphs[0] if index == 0 else cell.add_paragraph()
        if index == 0:
            paragraph.clear()
        set_paragraph_spacing(
            paragraph,
            spec.get("before", 0),
            spec.get("after", 0),
            spec.get("line", 1.05),
            spec.get("align", align),
        )
        run = paragraph.add_run(spec["text"])
        set_run_font(
            run,
            spec.get("font", "宋体"),
            spec.get("size", 9.5),
            bold=spec.get("bold", False),
            color=spec.get("color", INK),
        )
    cell.vertical_alignment = (
        WD_CELL_VERTICAL_ALIGNMENT.CENTER if center else WD_CELL_VERTICAL_ALIGNMENT.TOP
    )


def add_data_table(
    doc,
    headers: list[str],
    rows: list[list[str]],
    widths: list[float],
    *,
    bold_cols: set[int] | None = None,
    center_cols: set[int] | None = None,
    size: float = 9.5,
    zebra_colors: list[str] | None = None,
):
    bold_cols = bold_cols if bold_cols is not None else {0}
    center_cols = center_cols or set()
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    set_table_widths(table, widths)
    for index, header in enumerate(headers):
        cell = table.rows[0].cells[index]
        set_cell_lines(
            cell,
            [{"text": header, "font": "微软雅黑", "size": 9, "bold": True, "color": WHITE, "align": "center"}],
            align="center",
            center=True,
        )
        shade_cell(cell, NAVY)
        set_cell_margin(cell, 50, 50, 70, 70)
    mark_header_row(table.rows[0])
    prevent_row_split(table.rows[0])

    for r_index, row_data in enumerate(rows):
        fill = ZEBRA if r_index % 2 == 0 else WHITE
        if zebra_colors and r_index < len(zebra_colors) and zebra_colors[r_index]:
            fill = zebra_colors[r_index]
        for c_index, text in enumerate(row_data):
            cell = table.rows[r_index + 1].cells[c_index]
            is_bold = c_index in bold_cols
            align = "center" if c_index in center_cols else "left"
            parts = text.split("\n")
            lines = []
            for part_index, part in enumerate(parts):
                lines.append({
                    "text": part,
                    "font": "微软雅黑" if is_bold else "宋体",
                    "size": size,
                    "bold": is_bold,
                    "color": NAVY if is_bold else INK,
                    "align": align,
                    "after": 1 if part_index < len(parts) - 1 else 0,
                })
            set_cell_lines(cell, lines, align=align, center=True)
            shade_cell(cell, fill)
            set_cell_margin(cell, 46, 46, 70, 70)
        prevent_row_split(table.rows[r_index + 1])
    if len(rows) <= 8:
        keep_short_table_together(table)
    return table


def add_callout(doc, paragraphs: list, *, kind="info", title: str | None = None):
    palette = {
        "info": (INFO_BG, TEAL, TEAL),
        "warn": (WARN_BG, GOLD, "6E5428"),
        "key": (KEY_BG, NAVY, NAVY),
        "navy": (NAVY, GOLD, WHITE),
    }
    fill, border, title_color = palette[kind]
    body_color = WHITE if kind == "navy" else INK
    table = doc.add_table(rows=1, cols=1)
    set_table_widths(table, [CONTENT_W])
    # 外框改成左边强调线，避免和数据表一样的全网格。
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if edge == "left":
            element.set(qn("w:val"), "single")
            element.set(qn("w:sz"), "16")
            element.set(qn("w:color"), border)
        elif edge in ("top", "bottom", "right"):
            element.set(qn("w:val"), "single")
            element.set(qn("w:sz"), "4")
            element.set(qn("w:color"), "E7E1D6" if kind != "navy" else NAVY)
        else:
            element.set(qn("w:val"), "nil")
    cell = table.cell(0, 0)
    shade_cell(cell, fill)
    set_cell_margin(cell, 80, 80, 140, 120)
    lines = []
    if title:
        lines.append({"text": title, "font": "微软雅黑", "size": 11, "bold": True, "color": title_color, "after": 3})
    for index, paragraph in enumerate(paragraphs):
        lines.append({
            "text": paragraph,
            "font": "宋体",
            "size": 10.5,
            "bold": False,
            "color": body_color,
            "after": 2 if index < len(paragraphs) - 1 else 0,
            "before": 1,
        })
    set_cell_lines(cell, lines, center=False)
    prevent_row_split(table.rows[0])
    spacer(doc, 6)
    return table


def spacer(doc, points=6):
    paragraph = doc.add_paragraph()
    set_paragraph_spacing(paragraph, 0, 0, 1.0)
    paragraph.paragraph_format.line_spacing = Pt(points)
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    return paragraph


def add_star_card(doc, guest: dict) -> None:
    table = doc.add_table(rows=1, cols=1)
    set_table_widths(table, [CONTENT_W])
    borders = table._tbl.tblPr.find(qn("w:tblBorders"))
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if edge == "left":
            element.set(qn("w:val"), "single")
            element.set(qn("w:sz"), "18")
            element.set(qn("w:color"), GOLD)
        elif edge in ("top", "bottom", "right"):
            element.set(qn("w:val"), "single")
            element.set(qn("w:sz"), "4")
            element.set(qn("w:color"), "E8E2D8")
        else:
            element.set(qn("w:val"), "nil")
    cell = table.cell(0, 0)
    shade_cell(cell, CREAM)
    set_cell_margin(cell, 70, 70, 140, 120)
    set_cell_lines(
        cell,
        [
            {"text": guest["tag"], "font": "微软雅黑", "size": 9, "bold": True, "color": "8C6239", "after": 1},
            {"text": f"{guest['name']}　　{guest['region']}", "font": "黑体", "size": 14, "bold": True, "color": NAVY, "after": 2},
            {"text": guest["role"], "font": "宋体", "size": 10.5, "color": MUTED, "after": 2},
            {"text": f"建议演讲　《{guest['topic']}》", "font": "微软雅黑", "size": 11, "bold": True, "color": TEAL, "after": 2},
            {"text": guest["note"], "font": "宋体", "size": 10.5, "color": INK, "after": 0},
        ],
    )
    prevent_row_split(table.rows[0])
    spacer(doc, 7)


def configure_styles(doc: Document) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.font.color.rgb = rgb(INK)
    rpr = normal.element.get_or_add_rPr()
    _rfonts(rpr, "宋体", "Calibri")
    pf = normal.paragraph_format
    pf.space_after = Pt(6)
    pf.space_before = Pt(0)
    pf.line_spacing = 1.15

    for style_name, size, east, color in (
        ("Heading 1", 16, "黑体", NAVY),
        ("Heading 2", 13, "微软雅黑", TEAL),
        ("Heading 3", 12, "微软雅黑", NAVY),
    ):
        style = doc.styles[style_name]
        style.font.name = "Calibri"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = rgb(color)
        style.font.italic = False
        rpr = style.element.get_or_add_rPr()
        _rfonts(rpr, east, "Calibri")


def configure_section(section, *, header: bool) -> None:
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.7)
    section.right_margin = Cm(1.7)
    section.top_margin = Cm(1.7)
    section.bottom_margin = Cm(1.55)
    section.header_distance = Cm(0.5)
    section.footer_distance = Cm(0.4)
    sect_pr = section._sectPr
    grid = sect_pr.find(qn("w:docGrid"))
    if grid is None:
        grid = OxmlElement("w:docGrid")
        sect_pr.append(grid)
    grid.set(qn("w:type"), "default")
    grid.set(qn("w:linePitch"), "312")

    if not header:
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False
        section.header.paragraphs[0].text = ""
        section.footer.paragraphs[0].text = ""
        return

    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False
    header_p = section.header.paragraphs[0]
    header_p.clear()
    set_paragraph_spacing(header_p, 0, 2, 1.0)
    header_p.paragraph_format.tab_stops.add_tab_stop(Cm(CONTENT_W), WD_TAB_ALIGNMENT.RIGHT)
    left = header_p.add_run("2026上海国际智慧康养产业博览会")
    set_run_font(left, "微软雅黑", 8.5, color=NAVY)
    right = header_p.add_run("\t内部邀约及招商稿  ·  V4.0")
    set_run_font(right, "微软雅黑", 8.5, color=MUTED)
    ppr = header_p._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), GOLD)
    borders.append(bottom)
    ppr.append(borders)

    footer = section.footer.paragraphs[0]
    footer.clear()
    set_paragraph_spacing(footer, 2, 0, 1.0)
    footer.paragraph_format.tab_stops.add_tab_stop(Cm(CONTENT_W), WD_TAB_ALIGNMENT.RIGHT)
    foot_left = footer.add_run("候选名单，不代表已确认出席、主办或赞助")
    set_run_font(foot_left, "宋体", 8, color=MUTED)
    foot_mid = footer.add_run("\t第 ")
    set_run_font(foot_mid, "宋体", 8, color=MUTED)
    add_field(footer, " PAGE ")
    foot_right = footer.add_run(" 页")
    set_run_font(foot_right, "宋体", 8, color=MUTED)
    ppr = footer._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "6")
    top.set(qn("w:space"), "1")
    top.set(qn("w:color"), LINE)
    borders.append(top)
    ppr.append(borders)


def add_field(paragraph, instruction: str) -> None:
    run = paragraph.add_run()
    set_run_font(run, "宋体", 8, color=MUTED)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(separate)
    run._r.append(text)
    run._r.append(end)


def drop_leading_empty_paragraph(doc: Document) -> None:
    body = doc.element.body
    first = body.find(qn("w:p"))
    if first is not None and not first.findall(".//" + qn("w:t")):
        body.remove(first)


def build_cover(doc: Document) -> None:
    drop_leading_empty_paragraph(doc)
    banner = doc.add_table(rows=1, cols=1)
    set_table_widths(banner, [CONTENT_W])
    borders = banner._tbl.tblPr.find(qn("w:tblBorders"))
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if edge == "bottom":
            element.set(qn("w:val"), "single")
            element.set(qn("w:sz"), "16")
            element.set(qn("w:color"), GOLD)
        else:
            element.set(qn("w:val"), "nil")
    cell = banner.cell(0, 0)
    shade_cell(cell, NAVY)
    set_cell_margin(cell, 180, 160, 180, 160)
    set_cell_lines(
        cell,
        [
            {"text": "2026 上海国际智慧康养产业博览会", "font": "微软雅黑", "size": 11, "color": "E8D7B8", "after": 8},
            {"text": "AI重塑养老未来", "font": "黑体", "size": 28, "bold": True, "color": WHITE, "after": 4},
            {"text": "2026银发科技与具身智能产业创新峰会", "font": "微软雅黑", "size": 16, "bold": True, "color": WHITE, "after": 8},
            {"text": "“十五五”银发经济  ·  人形机器人  ·  AI健康管理  ·  智慧物联网  ·  产业资本", "font": "宋体", "size": 10.5, "color": "F3E6D0", "after": 0},
        ],
    )
    spacer(doc, 14)
    add_para(doc, "拟邀嘉宾、活动议程与产业合作方案", size=16, after=2, font="黑体", color=NAVY, align="left")
    add_para(doc, "V4.0　　2026年11月11日　　组委会内部邀约及招商使用稿", size=11, after=8, font="微软雅黑", color=TEAL)

    add_caption(doc, "基本信息")
    add_data_table(
        doc,
        ["项目", "内容"],
        [
            ["文件性质", "内部工作稿，用于讨论、分头邀约和招商洽谈"],
            ["对外使用", "正式海报、席卡和新闻稿只采用已经确认的名单"],
            ["举办时间", "2026年11月28日（周六）13:00—16:30"],
            ["举办地点", "上海展览中心中央大厅会议室C（以组委会最终确认为准）"],
            ["活动规模", "100—120人"],
            ["活动形式", "主旨演讲、前沿科技展示、产业高峰圆桌、创新路演、供需对接"],
        ],
        [3.3, 14.3],
        bold_cols={0},
        size=10,
    )
    spacer(doc, 10)
    add_callout(
        doc,
        ["全部人选、承办身份和赞助金额均为候选方案，不代表已经同意出席、担任主办单位或提供赞助。"],
        kind="warn",
        title="请先阅读使用须知",
    )


def build_notice(doc: Document) -> None:
    add_h1(doc, "使用须知", new_page=True)
    add_body(doc, "本稿可以发给组委会内部、邀约联系人和招商洽谈对象。向外发布的海报和新闻口径，必须以书面确认结果为准。")
    add_items(
        doc,
        [
            "全部嘉宾、单位和合作身份都属于候选邀约，不代表已经同意出席、担任主办单位或提供赞助。",
            "学术支持、联合承办和产业合作身份，须分别取得有关单位的正式同意。",
            "商业赞助单位不因出资自动取得高校学术身份或会议主办身份。",
            "议程中以斜线分隔的人选是替代关系，每个主旨只安排一名最终确认的主讲人。",
            "招商金额是本场分论坛的建议洽谈报价，不是企业预算，也不是已经取得的赞助意向。未落实的赞助不能作为可用资金，各项报价也不能相加后视为收入。",
            "不承诺政府领导出席、媒体必然报道、投资融资成功、采购成交，或未经授权的高校品牌背书。",
            "被邀请作演讲，不等于已经支付赞助。傅利叶、智元等技术企业宜先落实产业内容合作，再单独谈判赞助。",
            "工商银行上海市分行已列为大会主办单位。邀请其他银行前，先核查大会在金融行业的冠名和赞助排他权益。",
            "本次不安排王红霞上台，不设置甘薯科研成果专题。正式邀请朱震前，核对姓名和职务。",
            "优先以创始人、院长和资深专家发出正式邀请，同时保留副总裁、业务负责人或专业团队负责人的替代方案。",
        ],
        numbered=True,
    )


def build_toc(doc: Document) -> None:
    add_h1(doc, "章节目录")
    add_body(doc, "章节按阅读和执行顺序统一编号，便于分头联系时引用。")
    chapters = [
        "导言    邀约思路与可执行组合",
        "一、建议优先邀请的明星嘉宾",
        "二、活动概况",
        "三、组织架构（拟调整）",
        "四、第一梯队：重点拟邀主旨嘉宾",
        "五、第二梯队：重点拟邀产业领军人物",
        "六、第三梯队：金融资本与产业投资拟邀",
        "七、联合承办方与项目原有核心嘉宾",
        "八、拟邀活动议程",
        "九、高端圆桌嘉宾建议",
        "十、拟邀赞助及商务合作企业",
        "十一、赞助等级及权益",
        "十二、现场展示与产业合作设计",
        "十三、执行目标及时间安排",
        "十四、最终推荐的邀约组合",
        "十五、邀约精力分配",
        "附录    公开资料线索",
    ]
    add_items(doc, chapters, numbered=False, size=11)


def build_intro(doc: Document) -> None:
    add_h1(doc, "导言    邀约思路与可执行组合")
    add_body(
        doc,
        "本方案建议把邀约范围从传统养老行业向外扩展，突出上海及长三角最有影响力的人工智能、具身智能机器人、脑机接口、智慧物联网和数字医疗企业，同时保留复旦学术资源、养老机构和金融投资机构。",
    )
    add_h2(doc, "最理想的组合")
    add_items(
        doc,
        [
            ("复旦学术高度：", "彭希哲、吴玉韶、胡湛等老龄研究专家。"),
            ("机器人产业影响力：", "傅利叶顾捷、智元邓泰华或彭志辉、宇树王兴兴。"),
            ("前沿科技热点：", "强脑科技韩璧丞、商汤徐立，以及智慧物联网和数字医疗企业代表。"),
            ("落地与商业转化：", "胡继刚、养老机构负责人、金融机构和产业投资人。"),
        ],
    )
    add_body(
        doc,
        "复旦老龄研究团队、傅利叶智能及智元机器人的相关负责人已有公开资料可以核实。复旦大学 2025 年发布的《中国银发科技发展报告》特别关注养老机器人、数智养老、主动健康与养老金融，与本次论坛主题高度一致。",
    )
    add_callout(
        doc,
        [
            "不必要求五位明星嘉宾全部到场。能够确认顾捷或智元核心负责人中的一位，再确认彭希哲、吴玉韶或胡湛中的一位，已经足以建立论坛的内容主线。其余嘉宾作为冲刺邀约、圆桌替补或后续产业合作对象。",
        ],
        kind="key",
        title="可执行的最低组合",
    )


def build_stars(doc: Document) -> None:
    add_h1(doc, "一、建议优先邀请的明星嘉宾")
    add_body(doc, "以下五位兼顾行业知名度与上海、长三角科技产业资源，可作为本届最有吸引力的一批邀约。职务和题目在正式发函前再核对一次。")
    guests = [
        {
            "tag": "首要争取",
            "name": "顾捷",
            "region": "上海",
            "role": "傅利叶智能创始人、CEO",
            "topic": "从康复机器人到具身智能：AI如何改变未来养老",
            "note": "上海企业，直接关联康复、机器人及养老应用场景，是最适合担任产业主旨演讲的嘉宾。",
        },
        {
            "tag": "传播亮点",
            "name": "彭志辉",
            "region": "上海",
            "role": "智元机器人联合创始人、总裁、CTO",
            "topic": "具身智能大爆发：人形机器人何时真正进入家庭",
            "note": "科技圈影响力高，适合打造 AI 和人形机器人方面的传播亮点。企业总部位于上海。",
        },
        {
            "tag": "高规格冲刺",
            "name": "王兴兴",
            "region": "杭州",
            "role": "宇树科技创始人、董事长",
            "topic": "机器人从舞台走向千家万户：具身智能的新机遇",
            "note": "杭州企业，具有很强的公众传播力，属于高规格冲刺型嘉宾。",
        },
        {
            "tag": "热点方向",
            "name": "韩璧丞",
            "region": "杭州",
            "role": "强脑科技创始人、CEO",
            "topic": "脑机接口遇见康养机器人：主动康复的新突破",
            "note": "杭州企业，脑机接口与智能康复的结合度高。",
        },
        {
            "tag": "学术高度",
            "name": "彭希哲",
            "region": "上海",
            "role": "复旦大学老龄研究院院长",
            "topic": "“十五五”银发经济：AI时代的长寿社会与产业新机遇",
            "note": "为整场活动提供老龄研究与政策方面的学术高度。",
        },
    ]
    for guest in guests:
        add_star_card(doc, guest)


def build_overview(doc: Document) -> None:
    add_h1(doc, "二、活动概况")
    add_body(
        doc,
        "论坛立足上海、联动长三角，汇聚人口老龄化与银发经济研究专家、人工智能及具身智能企业、康养科技创新机构、养老运营服务商和产业资本，推动新技术与真实养老应用场景深度融合。",
    )
    add_caption(doc, "表 1    活动基本安排")
    add_data_table(
        doc,
        ["项目", "安排"],
        [
            ["举办时间", "2026年11月28日（周六）13:00—16:30"],
            ["举办地点", "上海展览中心中央大厅会议室C（以组委会最终确认为准）"],
            ["活动规模", "100—120人"],
            ["活动形式", "高端主旨演讲、前沿科技展示、产业高峰圆桌、创新项目路演、供需对接"],
            ["场地条件", "场地免费，主办方不提供现金经费"],
            ["建议主标题", "AI重塑养老未来——2026银发科技与具身智能产业创新峰会"],
            ["副标题", "“十五五”银发经济·人形机器人·AI健康管理·智慧物联网·产业资本"],
        ],
        [3.3, 14.3],
        bold_cols={0},
        size=10,
    )
    spacer(doc, 6)
    add_body(doc, "内部现金赞助目标和执行支出见第十三章，未到账金额不计入可用资金。")


def build_org(doc: Document) -> None:
    add_h1(doc, "三、组织架构（拟调整）")
    add_callout(
        doc,
        ["下列架构均为拟议安排。学术支持、承办和产业合作身份须分别获得有关单位正式同意后，才能写入对外材料。"],
        kind="warn",
    )
    add_h2(doc, "大会主办单位")
    add_items(
        doc,
        [
            "上海现代服务业联合会",
            "上海市养老服务和老龄产业协会",
            "上海丰之米会展有限公司",
            "中国工商银行股份有限公司上海市分行",
        ],
    )
    add_h2(doc, "分论坛主办与联合承办")
    add_items(
        doc,
        [
            ("分论坛主办单位：", "上海星智耀科技有限公司（沿用原拟定方案，待组委会确认）。"),
            ("联合承办（拟增补）：", "上海市杨浦区科技企业联合会。"),
            ("联合承办（拟增补）：", "复旦大学住房政策研究中心。"),
        ],
    )
    add_h2(doc, "拟邀学术支持单位")
    add_items(
        doc,
        [
            "复旦大学老龄研究院",
            "复旦大学银龄未来实验室",
            "复旦大学有关AI及计算智能科研团队",
            "复旦大学长寿与衰老研究所",
        ],
    )
    add_h2(doc, "拟邀产业创新合作单位")
    add_items(
        doc,
        [
            "傅利叶智能",
            "智元机器人",
            "宇树科技",
            "强脑科技",
            "商汤科技",
            "璞跃中国",
            "国家技术转移东部中心",
        ],
    )
    add_body(doc, "其他物联网、医疗科技和养老服务机构，按实际合作情况增补。")


def build_tier1(doc: Document) -> None:
    add_h1(doc, "四、第一梯队：重点拟邀主旨嘉宾")
    add_body(doc, "第一梯队采取“同类型择优确认”，不把下表全部安排上台。表中演讲题为建议题目，可按嘉宾本人意见微调。与第一章传播向题目不一致时，以嘉宾确认的题目为准。")
    add_caption(doc, "表 2    第一梯队重点拟邀主旨嘉宾")
    add_data_table(
        doc,
        ["拟邀嘉宾", "单位及公开身份", "地区", "建议演讲主题"],
        [
            ["彭希哲", "复旦大学老龄研究院院长、文科资深教授", "上海", "《“十五五”银发经济与长寿社会发展的新机遇》"],
            ["吴玉韶", "复旦大学老龄研究院副院长、教授", "上海", "《中国养老服务变革与银发科技产业新趋势》"],
            ["胡湛", "复旦大学老龄研究院副院长、\n银龄未来实验室主任", "上海", "《数智养老与AI向善：面向未来的中国式养老》"],
            ["顾捷", "傅利叶智能创始人、CEO", "上海", "《从康复机器人到具身智能：\nAI如何改变未来养老》"],
            ["邓泰华", "智元机器人创始人、董事长兼CEO", "上海", "《具身智能产业化：人形机器人走向现实场景》"],
            ["彭志辉", "智元机器人联合创始人、总裁、CTO", "上海", "《人形机器人进入家庭：\n从技术突破到规模化应用》"],
            ["王兴兴", "宇树科技创始人、董事长", "杭州", "《机器人走入千家万户：具身智能的未来十年》"],
            ["韩璧丞", "强脑科技创始人、CEO", "杭州", "《脑机接口与智能康复：AI赋能主动健康》"],
            ["徐立", "商汤科技董事长兼CEO", "上海", "《大模型与具身智能：AI赋能银发科技新场景》"],
            ["邱锡鹏", "复旦大学教授、上海创智学院全时导师", "上海", "《从大模型到智能体：AI康养的下一代交互技术》"],
        ],
        [2.2, 6.5, 1.6, 7.3],
        bold_cols={0},
        center_cols={2},
        size=9,
    )
    add_h2(doc, "邀请策略")
    add_items(
        doc,
        [
            ("老龄学术：", "彭希哲、吴玉韶、胡湛，优先确认1—2位。"),
            ("具身智能：", "傅利叶、智元、宇树，优先确认1—2家。"),
            ("大模型与脑机接口：", "商汤、复旦AI专家、强脑科技，优先确认1位。"),
        ],
    )
    add_body(doc, "正式邀请优先发给企业创始人、院长和资深专家。若本人无法出席，保留副总裁、业务负责人或专业团队负责人作为替代。")


def build_tier2(doc: Document) -> None:
    add_h1(doc, "五、第二梯队：重点拟邀产业领军人物")
    add_h2(doc, "（一）人形机器人、康复机器人及具身智能")
    add_body(doc, "傅利叶为首选康养技术企业。智元、宇树作为高知名度扩展邀约。擎朗、科沃斯等更贴近日常服务场景。")
    add_caption(doc, "表 3    机器人与具身智能拟邀")
    add_data_table(
        doc,
        ["拟邀嘉宾", "单位", "适合议题"],
        [
            ["朱秋国", "云深处科技创始人", "智能机器人产业化、场景应用"],
            ["钱东奇", "科沃斯集团创始人、董事长", "服务机器人进入家庭与长寿经济"],
            ["钱程", "科沃斯集团副董事长、科沃斯品牌CEO", "家庭服务机器人与智能居家"],
            ["易港", "钛虎机器人创始人、董事长", "机器人核心硬件、灵巧操作及产业链"],
            ["拟邀业务负责人", "擎朗智能", "服务机器人在养老机构的运营场景"],
            ["拟邀技术负责人", "傅利叶智能康复业务团队", "康复设备、养老服务及人机交互"],
            ["拟邀产品负责人", "智元机器人具身智能业务团队", "机器人商业部署与服务应用"],
        ],
        [3.3, 6.5, 7.8],
        size=9.5,
    )
    add_h2(doc, "（二）AI大模型、智慧物联网与数字医疗")
    add_caption(doc, "表 4    大模型、物联网与数字医疗拟邀")
    add_data_table(
        doc,
        ["拟邀嘉宾", "单位及身份", "建议议题"],
        [
            ["肖仰华", "复旦大学教授、上海市数据科学重点实验室主任", "AI智能体、知识图谱与智慧健康服务"],
            ["钱鹏鹤", "移远通信董事长兼CEO", "AIoT、智能感知与智慧养老物联网络"],
            ["王学集", "涂鸦智能创始人兼CEO", "AI Home与未来智慧养老家庭"],
            ["张瑞安", "乐鑫科技创始人、CEO", "AIoT芯片与低成本智能健康设备"],
            ["张强", "联影医疗董事长、联席CEO", "高端医疗设备创新与主动健康"],
            ["沈定刚", "联影智能联合创始人、联席CEO", "医疗AI与智能化健康管理"],
            ["尚海龙", "商汤医疗联合创始人兼首席战略官", "医疗大模型与智慧健康应用"],
            ["拟邀负责人", "萤石网络", "家庭看护、智能居家与物联网终端"],
            ["拟邀负责人", "海康威视智慧养老相关业务", "智能视觉、安全监测、居家感知"],
            ["拟邀负责人", "大华股份智慧物联相关业务", "智慧社区、养老安全及监测设备"],
        ],
        [2.6, 7.4, 7.6],
        size=9,
    )
    spacer(doc, 6)
    add_body(
        doc,
        "邀请方向各有侧重：移远、乐鑫更适合技术产业链；涂鸦、萤石更接近智慧家庭；联影和商汤医疗适合健康管理与医疗AI；科沃斯、擎朗更贴近日常服务机器人。",
    )
    add_h2(doc, "（三）养老运营、智慧健康与应用场景")
    add_caption(doc, "表 5    养老运营与应用场景拟邀")
    add_data_table(
        doc,
        ["拟邀嘉宾", "单位", "建议角色"],
        [
            ["邱建伟", "泰康之家首席执行官", "养老集团战略及智慧康养实践"],
            ["赵一洪", "福寿康联合创始人、副总裁", "AI+居家养老及长期照护"],
            ["康剑", "福寿康集团产品及供给负责人、上海分公司总经理", "智慧护理技术采购与应用"],
            ["瞿文", "上海福苑康养健康管理有限公司联合创始人、董事长", "智慧健康、康复及养老运营"],
            ["陈楠", "博斯腾科技联合创始人、CEO", "AI认知健康与老年认知障碍筛查"],
            ["拟邀负责人", "复星康养、星堡养老", "高品质养老社区科技应用"],
            ["拟邀负责人", "泰康之家·申园", "上海养老机构真实场景"],
            ["拟邀负责人", "上海市第一社会福利院", "养老机构产品试点与采购需求"],
            ["拟邀负责人", "九如城养老集团", "连锁养老机构智慧化改造"],
            ["拟邀负责人", "上海相关康复医院", "康复机器人临床应用与验证"],
        ],
        [2.6, 7.6, 7.4],
        size=9,
    )
    spacer(doc, 6)
    add_callout(
        doc,
        ["最终至少安排一家真正使用设备、承担运营成本或拥有采购决策权限的养老机构上台。避免台上只有技术提供方和投资人，没有真实买方。"],
        kind="key",
        title="买方原则",
    )


def build_tier3(doc: Document) -> None:
    add_h1(doc, "六、第三梯队：金融资本与产业投资拟邀")
    add_h2(doc, "（一）养老金融、产业金融嘉宾")
    add_caption(doc, "表 6    养老金融与科技金融拟邀")
    add_data_table(
        doc,
        ["拟邀对象", "单位", "适合议题"],
        [
            ["封进", "复旦大学特聘教授、老龄经济研究领域专家", "人口老龄化、养老金融与长期护理保险"],
            ["顾建忠", "上海银行董事长", "银发经济及科技金融协同"],
            ["邱建伟", "泰康之家CEO", "养老服务商业模式与金融支付"],
            ["拟邀负责人", "工商银行上海市分行养老金融或科技金融业务", "银发科技融资及金融服务"],
            ["拟邀负责人", "交通银行养老金融业务", "养老金金融及银发产业"],
            ["拟邀负责人", "浦发银行科技金融业务", "机器人及医疗科技企业融资"],
            ["拟邀负责人", "杭州银行上海分行科技金融业务", "长三角科创企业金融支持"],
            ["拟邀负责人", "泰隆银行上海相关分支机构", "科技型中小企业金融服务"],
            ["拟邀负责人", "太平养老、平安养老等机构", "长护险、养老服务支付与保险协同"],
        ],
        [2.6, 7.5, 7.5],
        size=9,
    )
    spacer(doc, 4)
    add_body(doc, "工商银行上海市分行已是大会主办单位。上表若接触其他银行，先完成金融行业排他核查。")
    add_h2(doc, "（二）产业投资机构")
    add_body(doc, "优先邀请具有实际项目投资权限的合伙人、投资总监或业务负责人。不安排纯品牌推广人员代替产业投资嘉宾。")
    add_items(
        doc,
        [
            "优势资本相关投资合伙人",
            "启明创投医疗科技或硬科技投资团队",
            "上海国投相关产业基金",
            "张江科投",
            "浦东创投",
            "临港科创投",
            "复星相关投资团队",
            "机器人及具身智能专项产业基金",
            "长三角医疗科技及智能制造产业基金",
        ],
    )


def build_core_people(doc: Document) -> None:
    add_h1(doc, "七、联合承办方与项目原有核心嘉宾")
    add_h2(doc, "胡继刚")
    add_star_card(
        doc,
        {
            "tag": "核心圆桌",
            "name": "胡继刚",
            "region": "上海",
            "role": "复旦大学住房政策研究中心秘书长；副教授级高级工程师；上海市杨浦区科技企业联合会执行会长",
            "topic": "从高校科技成果到银发产业应用：适老空间、具身智能与资本协同",
            "note": "拟任分论坛产业资源统筹负责人、核心圆桌嘉宾。发言重点围绕高校研究、科技企业、智慧居住场景、产业应用及投资合作。",
        },
    )
    add_h2(doc, "其他原有资源")
    add_items(
        doc,
        [
            ("朱震：", "原方案拟邀的复旦大学科技园及杨浦科企联相关嘉宾，建议参与科技成果转化和区域产业生态环节。正式邀请前核对姓名和职务。"),
            ("赵博誉：", "上海星智耀科技有限公司相关负责人，建议参与技术转移、项目路演和成果跟进。"),
            ("徐洁平：", "璞跃中国，拟邀担任论坛主持人或产业圆桌主持人。"),
            ("上海现代服务业联合会相关负责人：", "建议开幕致辞。"),
            ("上海市养老服务和老龄产业协会相关负责人：", "建议致辞、场景需求发布及采购机构邀约。"),
            ("明确不安排：", "本次不安排王红霞上台，不设置甘薯科研成果专题。"),
        ],
    )


def build_agenda(doc: Document) -> None:
    add_h1(doc, "八、拟邀活动议程", new_page=True)
    add_para(doc, "2026年11月28日（周六）　13:00—16:30　上海展览中心", size=11, before=0, after=4, font="微软雅黑", color=NAVY)
    add_callout(
        doc,
        ["每个主旨只保留一名最终确认的主讲人。单元格里的斜线表示替代人选，不是要求这些人一起上台。"],
        kind="warn",
        title="读表方式",
    )
    add_caption(doc, "表 7    拟邀活动议程")
    rows = [
        ["13:00—13:08", "开幕：AI重塑养老未来\n主持开场及嘉宾介绍", "徐洁平或专业主持人"],
        ["13:08—13:18", "开幕致辞：共建长三角银发科技产业生态", "大会主办方、行业协会代表"],
        ["13:18—13:38", "主旨一：“十五五”银发经济与AI时代的长寿社会", "彭希哲／吴玉韶／胡湛"],
        ["13:38—14:00", "主旨二：具身智能走进养老\n从康复机器人到家庭陪护", "顾捷；备选邓泰华／彭志辉／傅利叶相关高管"],
        ["14:00—14:18", "主旨三：大模型与智能体\n重构未来健康管理与养老服务", "邱锡鹏／肖仰华／徐立／相关业务代表"],
        ["14:18—14:36", "主旨四：万物互联的养老新时代：AIoT与智慧家庭", "钱鹏鹤／王学集／张瑞安／企业代表"],
        ["14:36—14:50", "产业实践：谁来为养老机器人买单？", "邱建伟／赵一洪／康剑／养老机构代表"],
        ["14:50—15:00", "茶歇与机器人展示交流", "参会嘉宾"],
        ["15:00—15:45", "高端圆桌：机器人进入养老院\n技术成熟了，商业模式成熟了吗？", "胡继刚、机器人企业、复旦学者、养老运营方、投资机构"],
        ["15:45—16:15", "创新成果秀：未来养老的AI新产品与应用场景", "傅利叶、智元、强脑及AIoT创新企业中已确认的2—3家"],
        ["16:15—16:25", "科技与康养产业供需对接及合作意向发布", "有实际需求或合作基础的参会单位"],
        ["16:25—16:30", "论坛总结及合影", "主承办方代表"],
    ]
    widths = [3.15, 8.15, 6.3]
    table = doc.add_table(rows=1 + len(rows), cols=3)
    set_table_widths(table, widths)
    for index, header in enumerate(["时间", "环节及题目", "拟邀嘉宾"]):
        cell = table.rows[0].cells[index]
        set_cell_lines(
            cell,
            [{"text": header, "font": "微软雅黑", "size": 9, "bold": True, "color": WHITE, "align": "center"}],
            align="center",
            center=True,
        )
        shade_cell(cell, NAVY)
        set_cell_margin(cell, 50, 50, 70, 70)
    mark_header_row(table.rows[0])
    prevent_row_split(table.rows[0])
    for r_index, (time_text, title, guest) in enumerate(rows):
        is_break = time_text.startswith("14:50")
        fill = "F3E6CC" if is_break else (ZEBRA if r_index % 2 == 0 else WHITE)
        values = [
            (time_text, True, TEAL, "center"),
            (title, True, NAVY, "left"),
            (guest, False, INK, "left"),
        ]
        for c_index, (text, is_bold, color, align) in enumerate(values):
            cell = table.rows[r_index + 1].cells[c_index]
            parts = text.split("\n")
            lines = []
            for part_index, part in enumerate(parts):
                lines.append({
                    "text": part,
                    "font": "微软雅黑" if is_bold else "宋体",
                    "size": 9 if c_index == 0 else 9.5,
                    "bold": is_bold,
                    "color": color,
                    "align": align,
                    "after": 1 if part_index < len(parts) - 1 else 0,
                })
            set_cell_lines(cell, lines, align=align, center=True)
            shade_cell(cell, fill)
            set_cell_margin(cell, 48, 48, 70, 70)
        prevent_row_split(table.rows[r_index + 1])

    add_h2(doc, "议程安排说明")
    add_items(
        doc,
        [
            "每个主旨报告只安排一家最终确认的单位或一名主讲嘉宾。斜线两侧是替代关系。",
            "如确认多位行业顶尖企业创始人，可压缩创新成果秀，增设一场企业家对话。",
            "若机器人实体演示条件不允许，用技术视频和案例分享替代。",
            "圆桌保持真实对话，不做简单的嘉宾轮流演讲。",
            "合作发布必须以实际可确认的项目、场景需求或合作意向为基础。",
        ],
        numbered=True,
    )


def build_roundtable(doc: Document) -> None:
    add_h1(doc, "九、高端圆桌嘉宾建议")
    add_h2(doc, "主题与主持人")
    add_body(doc, "主标题：《养老机器人商业化元年：谁来买单，谁来运营，谁来投资？》")
    add_body(doc, "备选名称：《从机器人到养老院：银发科技跨越商业化最后一公里》。")
    add_body(doc, "建议主持人：徐洁平，璞跃中国（拟邀）。")
    add_h2(doc, "圆桌六个席位")
    add_caption(doc, "表 8    圆桌席位建议")
    add_data_table(
        doc,
        ["席位", "建议人选", "讨论重点"],
        [
            ["第一席", "胡继刚", "高校成果转化、智慧居住、产业场景及资本协同"],
            ["第二席", "顾捷，或傅利叶智能业务负责人", "康复机器人、具身智能和养老应用的技术成熟度"],
            ["第三席", "胡湛，或复旦大学老龄研究专家", "老龄社会真实需求、主动健康及AI向善"],
            ["第四席", "邱建伟、赵一洪、康剑，或养老集团负责人", "机构是否愿意采购、试点及承担运营"],
            ["第五席", "王学集、钱鹏鹤，或AIoT企业负责人", "物联网、智能感知和AI家庭系统如何进入养老场景"],
            ["第六席", "产业投资基金或银行科技金融负责人", "产业融资、设备租赁、采购支付及投资回报"],
        ],
        [2.1, 6.3, 9.2],
        center_cols={0},
        size=9,
    )
    add_h2(doc, "圆桌核心问题")
    add_items(
        doc,
        [
            "人形机器人真正进入养老机构还缺少什么？",
            "未来三年，最有商业价值的养老机器人应用是什么？",
            "养老机构是否具备采购预算？",
            "AIoT与人形机器人如何构成智慧养老服务系统？",
            "养老设备能否通过租赁、订阅和服务付费模式商业化？",
            "科研、企业、机构和资本如何共同形成可复制的落地方案？",
        ],
        numbered=True,
    )


def build_sponsors(doc: Document) -> None:
    add_h1(doc, "十、拟邀赞助及商务合作企业")
    add_callout(
        doc,
        [
            "以下金额只作为本场分论坛的洽谈报价，不是企业预算，也不是已经取得的赞助意向。未签约、未到账的金额不能当作可用资金，不同企业的报价也不能相加后写成收入。",
            "邀请嘉宾上台，不等于对方已经同意赞助。技术企业先谈内容合作，再单独谈赞助。",
        ],
        kind="warn",
        title="金额怎么用",
    )
    add_h2(doc, "A类：优先开发的首席产业合作伙伴")
    add_caption(doc, "表 9    首席产业合作伙伴洽谈建议")
    add_data_table(
        doc,
        ["目标企业或机构", "合作方向", "建议洽谈金额"],
        [
            ["上海银行", "养老金融、科技金融、银发经济", "5万元"],
            ["浦发银行", "科创金融与机器人产业", "3万—5万元"],
            ["杭州银行上海分行", "长三角科创产业金融", "2万—5万元"],
            ["泰康相关机构", "养老金融、养老机构及健康服务", "3万—5万元"],
            ["复星相关康养或保险机构", "康养产业、保险与健康管理", "2万—5万元"],
            ["傅利叶智能及其产业生态伙伴", "具身智能康养应用", "2万—5万元或技术设备支持"],
            ["智元机器人及其生态合作伙伴", "具身智能产业生态", "2万—5万元或技术设备支持"],
        ],
        [5.5, 6.3, 5.8],
        center_cols={2},
        size=9.5,
    )
    spacer(doc, 6)
    add_body(doc, "对傅利叶、智元等技术企业，先落实产业内容合作，再独立谈判赞助。邀请其他银行前，先核查工商银行上海市分行作为大会主办单位所对应的金融行业冠名与赞助排他权益。")
    add_h2(doc, "B类：重点开发的产业创新合作伙伴")
    add_caption(doc, "表 10    产业创新合作伙伴洽谈建议")
    add_data_table(
        doc,
        ["企业", "合作方向", "建议洽谈金额"],
        [
            ["移远通信", "智能终端、物联网模组、端侧AI", "1万—2万元"],
            ["涂鸦智能", "AI Home、智慧养老家庭", "1万—2万元"],
            ["乐鑫科技", "AIoT芯片与智能设备", "1万—2万元"],
            ["科沃斯机器人", "智能家庭与服务机器人", "1万—3万元"],
            ["擎朗智能", "商业服务机器人与养老应用", "1万—2万元"],
            ["强脑科技", "脑机接口与智能康复", "1万—3万元"],
            ["联影相关企业", "AI医疗、数字健康", "1万—3万元"],
            ["商汤医疗", "医疗大模型及智能健康管理", "1万—3万元"],
            ["海康威视及萤石相关业务", "智慧居家及安全监护", "1万—2万元"],
            ["福寿康等养老运营机构", "养老应用场景与服务创新", "项目合作或专项赞助"],
        ],
        [4.8, 7.2, 5.6],
        center_cols={2},
        size=9.5,
    )
    add_h2(doc, "C类：特别支持与资源置换伙伴")
    add_items(
        doc,
        [
            "机器人设备展示伙伴",
            "健康监测和智能感知设备伙伴",
            "茶歇、饮水及接待支持伙伴",
            "摄影摄像与媒体制作伙伴",
            "活动直播和线上传播服务伙伴",
            "项目路演和技术转移服务伙伴",
        ],
    )
    add_body(doc, "对具备实际采购需求的养老机构，优先邀请其作为应用场景合作方，不要求所有机构支付赞助费。")


def build_benefits(doc: Document) -> None:
    add_h1(doc, "十一、赞助等级及权益")
    add_caption(doc, "表 11    赞助等级")
    add_data_table(
        doc,
        ["等级", "赞助金额", "主要权益"],
        [
            ["首席产业合作伙伴", "50,000元", "论坛专项品牌展示、经审核的专业主题报告、\n产品展示及核心产业对接"],
            ["产业创新合作伙伴", "20,000元", "品牌展示、创新案例交流、定向供需对接及嘉宾席位"],
            ["生态合作伙伴", "10,000元", "品牌展示、论坛交流及专项产业合作洽谈"],
            ["特别支持伙伴", "设备、实物或服务", "资源支持、合理品牌鸣谢及相应交流权益"],
        ],
        [3.8, 3.5, 10.3],
        center_cols={1},
        size=9.5,
    )
    add_h2(doc, "首席产业合作伙伴")
    add_items(
        doc,
        [
            "论坛宣传海报及主背景板品牌展示。",
            "6个嘉宾或企业代表席位。",
            "经内容审核的行业主题分享机会。",
            "优先提交产业合作项目及展示申请。",
            "参与论坛组织的产业供需交流。",
            "获得经授权可以交付的活动素材及总结资料。",
        ],
    )
    add_h2(doc, "产业创新合作伙伴")
    add_items(
        doc,
        [
            "论坛宣传与现场品牌展示。",
            "4个嘉宾或企业代表席位。",
            "项目展示申请及专业交流。",
            "参与定向产业对接。",
            "获得活动简讯及经许可的照片资料。",
        ],
    )
    add_h2(doc, "生态合作伙伴")
    add_items(
        doc,
        [
            "论坛品牌鸣谢及展示。",
            "2个参会名额。",
            "参加产业交流及合作需求对接。",
            "提交企业业务和创新项目资料。",
        ],
    )
    add_callout(
        doc,
        ["不承诺政府领导出席、媒体必然报道、投资融资成功、采购成交或未经授权的高校品牌背书。主题分享须经过内容审核。"],
        kind="warn",
        title="权益边界",
    )


def build_onsite(doc: Document) -> None:
    add_h1(doc, "十二、现场展示与产业合作设计")
    add_body(doc, "建议设置三种轻量化互动。")
    add_h3(doc, "1. 机器人及康复技术展示")
    add_body(doc, "由傅利叶或其他已经确认的科技企业提供产品展示、视频演示，或安全可控的现场操作。不要求必须运输大型人形机器人。")
    add_h3(doc, "2. AIoT智慧养老应用展示")
    add_body(doc, "重点展示智能感知、家庭健康管理、安全监测、康复设备互联等可以解释、可以验证的产品应用。")
    add_h3(doc, "3. 科技企业与养老机构需求对接")
    add_body(doc, "会前收集企业解决方案和养老机构的实际问题，按康复护理、智能陪护、居家安全、健康管理等场景匹配。现场安排分组洽谈，会后建立合作台账。")


def build_execution(doc: Document) -> None:
    add_h1(doc, "十三、执行目标及时间安排")
    add_h2(doc, "活动目标")
    add_caption(doc, "表 12    现场目标")
    add_data_table(
        doc,
        ["指标", "目标"],
        [
            ["核心演讲及圆桌嘉宾", "8—10位"],
            ["行业协会、学术及产业代表", "按现场规模安排"],
            ["科技企业", "争取15—20家"],
            ["养老、康复及健康服务机构", "争取10—15家"],
            ["金融及投资机构", "争取5—8家"],
            ["有效产业对接", "争取10—15组"],
            ["可持续跟进的合作项目", "争取2—3项"],
        ],
        [7.2, 10.4],
        size=10.5,
    )
    add_h2(doc, "关键节点")
    add_caption(doc, "表 13    倒排时间")
    add_data_table(
        doc,
        ["时间", "工作"],
        [
            ["11月11—13日", "确定论坛主办及联合承办结构，启动第一梯队邀约。"],
            ["11月13—16日", "集中联系傅利叶、复旦老龄研究专家、智元及行业代表，同时确认首批招商目标企业。"],
            ["11月16—20日", "锁定核心嘉宾、首批赞助合同、实际回款及主要养老机构代表。"],
            ["11月21—25日", "根据正式确认名单发布宣传物料，落实技术展示、参会邀约及执行细节。"],
            ["11月26—27日", "场地技术测试、物料交付、嘉宾接待及流程联排。"],
            ["11月28日", "活动举办。"],
            ["11月29日至12月5日", "赞助权益交付、项目对接跟进和活动总结。"],
        ],
        [4.2, 13.4],
        size=10,
    )
    add_h2(doc, "资金执行原则")
    add_callout(
        doc,
        [
            "场地免费，主办方不提供现金经费。",
            "建议内部现金赞助目标为8万元，基础执行支出控制在约2.3万元以内，具体以实际设备条件和服务报价为准。",
            "未落实的赞助不作为可用资金。主办方及联合承办方应通过书面协议明确招商授权、合同收款、预算审批、服务费用和收益结算。",
        ],
        kind="key",
        title="先有到账，再有支出",
    )


def build_final(doc: Document) -> None:
    add_h1(doc, "十四、最终推荐的邀约组合")
    add_h2(doc, "优先力争")
    add_body(doc, "复旦大学老龄研究专家 1 位，加傅利叶核心负责人 1 位，加 AI 或具身智能头部企业代表 1 位，加养老运营集团负责人 1 位，加产业投资或银行代表 1 位。")
    add_h2(doc, "核心圆桌的固定结构")
    add_body(doc, "胡继刚，加机器人企业负责人，加复旦学术专家，加养老机构负责人，加 AIoT 企业代表，加金融投资代表。")
    add_h2(doc, "招商突破口")
    add_body(doc, "把高水平产业嘉宾、真实养老采购需求、AI 机器人现场展示和科技金融合作打包，形成可洽谈的商业合作价值。")
    add_body(doc, "本届不追求所有知名嘉宾同时出席。议程要有行业吸引力、能够执行，参与者要能形成产业合作。")
    add_callout(
        doc,
        ["打造一场具有复旦学术特色、上海机器人产业影响力、长三角科技创新资源和真实商业价值的银发科技论坛。"],
        kind="navy",
        title="最终目标",
    )


def build_energy(doc: Document) -> None:
    add_h1(doc, "十五、邀约精力分配")
    add_body(doc, "时间紧，不要同时向几十位企业董事长发出没有明确角色的通用邀请函。每个赛道先确定一名首选和两名替补，并写明具体演讲题目或产业合作诉求，争取在11月16日前拿到首批明确回复。")
    add_caption(doc, "表 14    邀约层次与策略")
    add_data_table(
        doc,
        ["邀约层次", "典型对象", "建议策略"],
        [
            ["顶级冲刺", "王兴兴、彭志辉、邓泰华、徐立", "创始人优先，允许推荐高级负责人或视频致辞"],
            ["核心必争取", "顾捷、彭希哲、吴玉韶、胡湛", "优先通过已有高校、科企联及产业关系建立联系"],
            ["技术与商业主力", "钱鹏鹤、王学集、韩璧丞、科沃斯、联影智能", "以具体应用场景及合作需求切入"],
            ["实际落地保障", "福寿康、复星康养、泰康之家上海团队、养老协会", "锁定有采购权和试点意愿的负责人"],
            ["商业赞助主力", "银行科技金融部门、机器人产业链、AIoT企业", "以真实产业对接和赞助权益争取资金"],
        ],
        [3.2, 6.9, 7.5],
        size=9,
    )
    spacer(doc, 6)
    add_h2(doc, "复旦三个新增方向")
    add_body(doc, "复旦方面建议同时推进三条学术线索，均已有高校公开资料支持，发函前仍按使用须知复核职务。")
    add_items(
        doc,
        [
            ("邱锡鹏：", "大模型与智能体。"),
            ("肖仰华：", "认知智能与AI应用。"),
            ("封进：", "养老金融。"),
        ],
    )
    add_h2(doc, "物联网招商不要铺得太开")
    add_body(doc, "移远通信、涂鸦智能、乐鑫科技比泛泛寻找大量智能家居企业更有针对性，分别覆盖设备连接、AI Home 与机器人生态、端侧AI芯片。")
    add_h2(doc, "若只能先推动三条关系")
    add_callout(
        doc,
        [
            "建议优先推动傅利叶智能、复旦大学老龄研究院、上海市养老服务和老龄产业协会。前两者决定论坛的科技含量和学术吸引力，协会决定有没有真正的行业应用场景。招商围绕这三方能够形成的合作价值展开，而不是单纯销售舞台发言时间。",
        ],
        kind="key",
        title="三条主线",
    )


def build_appendix(doc: Document) -> None:
    add_h1(doc, "附录    公开资料线索")
    add_body(doc, "下表只记录本方案编制时参考的公开来源类型，方便邀约前复核身份。它不是出席确认，也不是授权使用对方品牌。")
    add_caption(doc, "表 15    复核时可回溯的公开线索")
    add_data_table(
        doc,
        ["对象", "资料线索"],
        [
            ["复旦老龄研究及\n《中国银发科技发展报告》", "复旦新闻文化网、复旦发展研究院等公开信息"],
            ["顾捷、傅利叶智能", "人民网等公开报道"],
            ["彭志辉、智元机器人", "智元机器人官网（agibot.com.cn）及公开报道"],
            ["王兴兴、宇树科技", "新华网等公开报道"],
            ["韩璧丞、强脑科技", "浙江新闻等公开报道"],
            ["邱锡鹏、肖仰华、封进", "复旦大学计算与智能创新学院等高校公开资料"],
            ["移远通信、涂鸦智能、乐鑫科技", "路演信息网等公开企业资料"],
        ],
        [6.4, 11.2],
        size=9.5,
    )
    spacer(doc, 8)
    add_para(
        doc,
        "— 全文完 · V4.0 · 2026年11月11日 —",
        size=9,
        before=6,
        after=0,
        align="center",
        font="微软雅黑",
        color=MUTED,
    )


def build() -> Document:
    doc = Document()
    configure_styles(doc)
    core = doc.core_properties
    core.title = "AI重塑养老未来——2026银发科技与具身智能产业创新峰会策划议程"
    core.subject = "拟邀嘉宾、活动议程与产业合作方案 V4.0"
    core.category = "组委会内部邀约及招商使用稿"
    core.author = "2026上海国际智慧康养产业博览会组委会内部稿"
    core.comments = "候选名单不代表已经同意出席、担任主办单位或提供赞助。"

    configure_section(doc.sections[0], header=False)
    build_cover(doc)

    body = doc.add_section()
    configure_section(body, header=True)
    # 正文页码从 1 起算，封面不计入。
    pg_num = OxmlElement("w:pgNumType")
    pg_num.set(qn("w:start"), "1")
    body._sectPr.append(pg_num)

    build_notice(doc)
    build_toc(doc)
    build_intro(doc)
    build_stars(doc)
    build_overview(doc)
    build_org(doc)
    build_tier1(doc)
    build_tier2(doc)
    build_tier3(doc)
    build_core_people(doc)
    build_agenda(doc)
    build_roundtable(doc)
    build_sponsors(doc)
    build_benefits(doc)
    build_onsite(doc)
    build_execution(doc)
    build_final(doc)
    build_energy(doc)
    build_appendix(doc)
    return doc


REQUIRED = [
    "顾捷",
    "彭志辉",
    "王兴兴",
    "韩璧丞",
    "彭希哲",
    "吴玉韶",
    "胡湛",
    "邓泰华",
    "徐立",
    "邱锡鹏",
    "朱秋国",
    "钱东奇",
    "钱程",
    "易港",
    "擎朗智能",
    "肖仰华",
    "钱鹏鹤",
    "王学集",
    "张瑞安",
    "张强",
    "沈定刚",
    "尚海龙",
    "萤石网络",
    "海康威视",
    "大华股份",
    "邱建伟",
    "赵一洪",
    "康剑",
    "瞿文",
    "陈楠",
    "复星康养",
    "星堡养老",
    "泰康之家·申园",
    "上海市第一社会福利院",
    "九如城",
    "封进",
    "顾建忠",
    "上海银行",
    "交通银行",
    "浦发银行",
    "杭州银行",
    "泰隆银行",
    "太平养老",
    "平安养老",
    "优势资本",
    "启明创投",
    "上海国投",
    "张江科投",
    "浦东创投",
    "临港科创投",
    "胡继刚",
    "朱震",
    "赵博誉",
    "徐洁平",
    "王红霞",
    "甘薯",
    "璞跃中国",
    "国家技术转移东部中心",
    "上海星智耀科技有限公司",
    "上海丰之米会展有限公司",
    "中国工商银行股份有限公司上海市分行",
    "上海市杨浦区科技企业联合会",
    "复旦大学住房政策研究中心",
    "复旦大学银龄未来实验室",
    "复旦大学长寿与衰老研究所",
    "上海现代服务业联合会",
    "上海市养老服务和老龄产业协会",
    "中国银发科技发展报告",
    "50,000元",
    "8万元",
    "2.3万元",
    "100—120人",
    "13:00—16:30",
    "谁来买单，谁来运营，谁来投资",
    "不安排王红霞上台",
]


def verify(path: Path) -> None:
    document = Document(str(path))
    texts = []
    for paragraph in document.paragraphs:
        texts.append(paragraph.text)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                texts.append(cell.text)
    full = "\n".join(texts)
    missing = [item for item in REQUIRED if item not in full]
    if missing:
        raise SystemExit("文档缺少关键内容：\n" + "\n".join(missing))
    print(f"段落 {len(document.paragraphs)}，表格 {len(document.tables)}，关键内容 {len(REQUIRED)} 项均已写入。")


def main() -> None:
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    document = build()
    document.save(OUT_PATH)
    verify(OUT_PATH)
    print(OUT_PATH)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"生成失败：{exc}", file=sys.stderr)
        raise
