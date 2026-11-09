#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按参考 PDF 的版式重排江苏省专业技术资格评审申报表。

页边、字号和行距取自 PDF 正文（水印不计）。工作经历与工作总结
跟 PDF 同一段：8 条经历、专业技术工作栏留空，总结为三列表加项目分条。
不设页眉。照片、签名和印章仍用原稿。
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets" / "title-application"
OUT = ROOT / "output" / "专业技术资格评审申报表.docx"
# 参考 PDF 表格左缘 54pt、右缘 525.8pt，版心 471.8pt。
CONTENT = 9436


def fit(widths: list[int]) -> list[int]:
    """把一组列宽缩放到当前版心。"""
    scaled = [max(1, int(round(width * CONTENT / sum(widths)))) for width in widths]
    scaled[-1] += CONTENT - sum(scaled)
    return scaled


ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}
VALIGN = {
    "top": WD_CELL_VERTICAL_ALIGNMENT.TOP,
    "center": WD_CELL_VERTICAL_ALIGNMENT.CENTER,
    "bottom": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM,
}


def _rgb(color: str) -> RGBColor:
    return RGBColor(int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16))


def set_run_font(run, east_asia: str, size_pt: float, bold: bool = False, color: str | None = None) -> None:
    ascii_font = "Times New Roman"
    run.bold = bold
    run.italic = False
    run.font.size = Pt(size_pt)
    run.font.name = ascii_font
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    r_fonts.set(qn("w:ascii"), ascii_font)
    r_fonts.set(qn("w:hAnsi"), ascii_font)
    r_fonts.set(qn("w:cs"), ascii_font)
    r_fonts.set(qn("w:eastAsia"), east_asia)
    half = str(int(round(size_pt * 2)))
    for tag in ("w:sz", "w:szCs"):
        node = r_pr.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            r_pr.append(node)
        node.set(qn("w:val"), half)
    if color:
        run.font.color.rgb = _rgb(color)


def style_paragraph(
    paragraph,
    *,
    align: str = "left",
    before: float = 0,
    after: float = 0,
    line_pt: float = 1.0,
    line_rule: str = "auto",
    first_chars: float | None = None,
    hanging_chars: float | None = None,
    font_pt: float = 10.5,
    page_break: bool = False,
    keep_next: bool = False,
) -> None:
    paragraph.alignment = ALIGN[align]
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    p_pr = paragraph._p.get_or_add_pPr()
    for tag in ("w:spacing", "w:ind", "w:pageBreakBefore", "w:keepNext", "w:snapToGrid"):
        node = p_pr.find(qn(tag))
        if node is not None:
            p_pr.remove(node)
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), str(int(round(before * 20))))
    spacing.set(qn("w:after"), str(int(round(after * 20))))
    if line_rule == "exact":
        spacing.set(qn("w:line"), str(int(round(line_pt * 20))))
        spacing.set(qn("w:lineRule"), "exact")
    else:
        spacing.set(qn("w:line"), str(int(round(line_pt * 240))))
        spacing.set(qn("w:lineRule"), "auto")
    p_pr.append(spacing)
    snap = OxmlElement("w:snapToGrid")
    snap.set(qn("w:val"), "0")
    p_pr.append(snap)
    if first_chars or hanging_chars:
        ind = OxmlElement("w:ind")
        if hanging_chars:
            twips = int(round(font_pt * 20 * hanging_chars))
            ind.set(qn("w:left"), str(twips))
            ind.set(qn("w:leftChars"), str(int(hanging_chars * 100)))
            ind.set(qn("w:hanging"), str(twips))
            ind.set(qn("w:hangingChars"), str(int(hanging_chars * 100)))
        elif first_chars:
            twips = int(round(font_pt * 20 * first_chars))
            ind.set(qn("w:firstLine"), str(twips))
            ind.set(qn("w:firstLineChars"), str(int(first_chars * 100)))
        p_pr.append(ind)
    if page_break:
        p_pr.append(OxmlElement("w:pageBreakBefore"))
    if keep_next:
        p_pr.append(OxmlElement("w:keepNext"))


def add_text(paragraph, text: str, font: str, size: float, bold: bool = False, color: str | None = None, underline: bool = False):
    run = paragraph.add_run(text)
    set_run_font(run, font, size, bold, color)
    run.underline = underline
    return run


def para(
    doc,
    text: str = "",
    *,
    font: str = "宋体",
    size: float = 10.5,
    bold: bool = False,
    align: str = "left",
    before: float = 0,
    after: float = 0,
    line_pt: float = 1.0,
    line_rule: str = "auto",
    first_chars: float | None = None,
    hanging_chars: float | None = None,
    page_break: bool = False,
    keep_next: bool = False,
    color: str | None = None,
    underline: bool = False,
):
    paragraph = doc.add_paragraph()
    style_paragraph(
        paragraph,
        align=align,
        before=before,
        after=after,
        line_pt=line_pt,
        line_rule=line_rule,
        first_chars=first_chars,
        hanging_chars=hanging_chars,
        font_pt=size,
        page_break=page_break,
        keep_next=keep_next,
    )
    if text:
        add_text(paragraph, text, font, size, bold, color, underline)
    return paragraph


def _set_tbl_borders(table, sz: str = "6", color: str = "000000", val: str = "single") -> None:
    tbl_pr = table._tbl.tblPr
    old = tbl_pr.find(qn("w:tblBorders"))
    if old is not None:
        tbl_pr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), val)
        node.set(qn("w:sz"), "0" if val == "nil" else sz)
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), "auto" if val == "nil" else color)
        borders.append(node)
    tbl_pr.append(borders)


def _set_cell_margins(table, top=40, bottom=40, left=60, right=60) -> None:
    tbl_pr = table._tbl.tblPr
    old = tbl_pr.find(qn("w:tblCellMar"))
    if old is not None:
        tbl_pr.remove(old)
    mar = OxmlElement("w:tblCellMar")
    for name, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        mar.append(node)
    tbl_pr.append(mar)


def apply_widths(table, widths: list[int]) -> None:
    table.autofit = False
    table.allow_autofit = False
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    total = sum(widths)
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(total))
    tbl_w.set(qn("w:type"), "dxa")
    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")
    grid = tbl.find(qn("w:tblGrid"))
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)
    for tr in tbl.findall(qn("w:tr")):
        index = 0
        for tc in tr.findall(qn("w:tc")):
            tc_pr = tc.find(qn("w:tcPr"))
            if tc_pr is None:
                tc_pr = OxmlElement("w:tcPr")
                tc.insert(0, tc_pr)
            span_node = tc_pr.find(qn("w:gridSpan"))
            span = int(span_node.get(qn("w:val"))) if span_node is not None else 1
            width = sum(widths[index : index + span])
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(width))
            tc_w.set(qn("w:type"), "dxa")
            index += span


def set_tc_borders(cell, **edges) -> None:
    """edges: top/left/bottom/right -> None 表示无边框，或 'single'。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    old = tc_pr.find(qn("w:tcBorders"))
    if old is not None:
        tc_pr.remove(old)
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        node = OxmlElement(f"w:{edge}")
        style = edges.get(edge, "nil")
        if style in (None, "nil"):
            node.set(qn("w:val"), "nil")
            node.set(qn("w:sz"), "0")
            node.set(qn("w:color"), "auto")
        else:
            node.set(qn("w:val"), "single")
            node.set(qn("w:sz"), "8")
            node.set(qn("w:color"), "000000")
        node.set(qn("w:space"), "0")
        borders.append(node)
    tc_pr.append(borders)


def set_row_height(row, twips: int, rule: str = "atLeast") -> None:
    tr_pr = row._tr.get_or_add_trPr()
    node = tr_pr.find(qn("w:trHeight"))
    if node is None:
        node = OxmlElement("w:trHeight")
        tr_pr.append(node)
    node.set(qn("w:val"), str(twips))
    node.set(qn("w:hRule"), rule)


def shade_cell(cell, color: str = "D9D9D9") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    old = tc_pr.find(qn("w:shd"))
    if old is not None:
        tc_pr.remove(old)
    shade = OxmlElement("w:shd")
    shade.set(qn("w:val"), "clear")
    shade.set(qn("w:color"), "auto")
    shade.set(qn("w:fill"), color)
    tc_pr.append(shade)


def mark_header_row(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:tblHeader")) is None:
        tr_pr.append(OxmlElement("w:tblHeader"))


def prevent_row_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def make_table(doc_or_cell, rows: int, cols: int, widths: list[int], *, borders: bool = True, header: bool = False):
    table = doc_or_cell.add_table(rows=rows, cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_tbl_borders(table, val="single" if borders else "nil")
    _set_cell_margins(table, top=28, bottom=28, left=40, right=40)
    apply_widths(table, widths)
    if header:
        mark_header_row(table.rows[0])
    for row in table.rows:
        prevent_row_split(row)
    return table


def fill(
    cell,
    text,
    *,
    font: str = "宋体",
    size: float = 9,
    bold: bool = False,
    align: str = "center",
    va: str = "center",
    line_pt: float | None = None,
    line_rule: str = "auto",
    color: str | None = None,
) -> None:
    parts = text.split("\n") if isinstance(text, str) else list(text)
    if not parts:
        parts = [""]
    cell.vertical_alignment = VALIGN[va]
    extras = cell.paragraphs[1:]
    for extra in extras:
        extra._element.getparent().remove(extra._element)
    rule = line_rule if line_pt is not None else "auto"
    spacing = 1.0 if line_pt is None else line_pt

    def write(paragraph, content: str) -> None:
        for run in list(paragraph.runs):
            run._element.getparent().remove(run._element)
        style_paragraph(paragraph, align=align, before=0, after=0, line_pt=spacing, line_rule=rule, font_pt=size)
        if content:
            add_text(paragraph, content, font, size, bold, color)

    write(cell.paragraphs[0], parts[0])
    for content in parts[1:]:
        write(cell.add_paragraph(), content)


def merge(table, r1: int, c1: int, r2: int, c2: int):
    cell = table.cell(r1, c1).merge(table.cell(r2, c2))
    return cell


def picture(paragraph, name: str, width_cm: float):
    run = paragraph.add_run()
    run.add_picture(str(ASSETS / name), width=Cm(width_cm))
    return run


def stack_vertical(cell, columns: list[str], *, size: float = 8, line_pt: float = 12) -> None:
    """竖排标签。多列从左到右读，每列自上而下。"""
    cell.vertical_alignment = VALIGN["center"]
    if len(columns) == 1:
        fill(cell, list(columns[0]), size=size, align="center", va="center", line_pt=line_pt, line_rule="exact")
        return
    lead = cell.paragraphs[0]
    style_paragraph(lead, align="center", line_pt=1, line_rule="exact")
    widths = [max(420, int(1400 / len(columns)))] * len(columns)
    nested = cell.add_table(rows=1, cols=len(columns))
    nested.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_tbl_borders(nested, val="nil")
    _set_cell_margins(nested, top=0, bottom=0, left=20, right=20)
    apply_widths(nested, widths)
    for index, text in enumerate(columns):
        inner = nested.cell(0, index)
        set_tc_borders(inner)
        fill(inner, list(text), size=size, align="center", va="center", line_pt=line_pt, line_rule="exact")


def _clear_story(story) -> None:
    story.is_linked_to_previous = False
    for paragraph in story.paragraphs:
        for run in list(paragraph.runs):
            run._element.getparent().remove(run._element)
        paragraph.text = ""


def set_page_border(section, enabled: bool) -> None:
    """工作总结各页外框。边线落在正文外侧，不压字。"""
    sect_pr = section._sectPr
    old = sect_pr.find(qn("w:pgBorders"))
    if old is not None:
        sect_pr.remove(old)
    if not enabled:
        return
    borders = OxmlElement("w:pgBorders")
    borders.set(qn("w:offsetFrom"), "page")
    for edge, space in (("top", "26"), ("left", "46"), ("bottom", "20"), ("right", "60")):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "8")
        node.set(qn("w:space"), space)
        node.set(qn("w:color"), "000000")
        borders.append(node)
    sect_pr.append(borders)


def prepare(doc: Document) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(10.5)
    normal.font.bold = False
    r_pr = normal.element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    r_fonts.set(qn("w:ascii"), "Times New Roman")
    r_fonts.set(qn("w:hAnsi"), "Times New Roman")
    r_fonts.set(qn("w:eastAsia"), "宋体")
    r_fonts.set(qn("w:cs"), "Times New Roman")
    lang = r_pr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        r_pr.append(lang)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:eastAsia"), "zh-CN")
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = 1.0
    doc.core_properties.title = "专业技术资格评审申报表"
    doc.core_properties.subject = "胡继刚"
    doc.core_properties.category = "申报表"
    doc.core_properties.comments = "按参考 PDF 的字号、段距和工作栏重排，无页眉"


def configure(section, *, border: bool = False, top: float = 0.72, bottom: float = 1.05) -> None:
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.90)
    section.right_margin = Cm(2.45)
    section.top_margin = Cm(top)
    section.bottom_margin = Cm(bottom)
    section.header_distance = Cm(0.3)
    section.footer_distance = Cm(0.4)
    _clear_story(section.header)
    _clear_story(section.footer)
    set_page_border(section, border)


def section_title(doc, text: str, *, page: bool = False, size: float = 25, after: float = 4, before: float = 0) -> None:
    para(
        doc,
        text,
        font="黑体",
        size=size,
        bold=True,
        align="center",
        before=before,
        after=after,
        line_pt=1.0,
        page_break=page,
    )


def note(doc, text: str, *, before: float = 2, after: float = 0, align: str = "left", size: float = 11.2) -> None:
    para(doc, text, font="宋体", size=size, align=align, before=before, after=after, line_pt=16, line_rule="exact")


def _line_break(paragraph) -> None:
    paragraph.add_run().add_break()


def build_cover(doc: Document) -> None:
    code = make_table(doc, 1, 2, fit([1, 1]), borders=False)
    fill(code.cell(0, 0), "地区代码：320199", size=11.2, align="left")
    fill(code.cell(0, 1), "编号：2506170651940505", size=11.2, align="right")
    set_row_height(code.rows[0], 400)
    para(doc, "专 业 技 术 资 格", font="宋体", size=33.8, bold=False, align="center", before=48, after=0, line_pt=76, line_rule="exact")
    para(doc, "评 审 申 报 表", font="宋体", size=33.8, bold=False, align="center", before=0, after=0, line_pt=76, line_rule="exact")
    fields = [
        ("单　　位", "江苏省建筑装饰设计研究院有限公司"),
        ("姓　　名", "胡继刚"),
        ("评审系列", "建设工程"),
        ("评审专业", "科技管理·工程管理"),
        ("技术资格", "高级工程师"),
    ]
    table = make_table(doc, len(fields), 3, [2018, 1800, 3900], borders=False)
    for index, (label, value) in enumerate(fields):
        fill(table.cell(index, 0), "", size=14)
        fill(table.cell(index, 1), label, font="楷体", size=14, align="right")
        fill(table.cell(index, 2), value, font="宋体", size=12, align="center")
        set_tc_borders(table.cell(index, 0))
        set_tc_borders(table.cell(index, 1))
        set_tc_borders(table.cell(index, 2), bottom="single")
        set_row_height(table.rows[index], 840)
    date = para(doc, "", align="center", before=36, after=0, line_pt=20, line_rule="exact")
    add_text(date, "填表时间：", "楷体", 14)
    add_text(date, "2025", "黑体", 12)
    add_text(date, "  年  ", "楷体", 14)
    add_text(date, "08", "黑体", 12)
    add_text(date, "  月  ", "楷体", 14)
    add_text(date, "12", "黑体", 12)
    add_text(date, "  日", "楷体", 14)
    para(doc, "江苏省专业技术人员职称(职业资格)工作领导小组办公室制", font="宋体", size=14.2, align="center", before=48, after=0, line_pt=22, line_rule="exact")


def build_pledge(doc: Document) -> None:
    section_title(doc, "个 人 承 诺 书", page=True, before=12, after=28, size=25)
    paragraph = para(doc, "", size=14.2, align="left", before=8, after=0, line_pt=32, line_rule="exact")
    lines = [
        [("本人申报", False, 14.2), ("　建设工程　", True, 12.8), ("系列", False, 14.2), ("　科技管理·工程管理　", True, 12.8), ("专业", False, 14.2)],
        [("(学科)", False, 14.2), ("　高级工程师　", True, 12.8), ("资格。本人承诺：所填写的内容及所提", False, 14.2)],
        [("供的参评材料真实准确，如有任何不实或隐瞒，愿按专业技术资格评审的", False, 14.2)],
        [("有关规定接受处理。", False, 14.2)],
    ]
    for index, line in enumerate(lines):
        if index:
            _line_break(paragraph)
        for text, underlined, size in line:
            add_text(paragraph, text, "宋体", size, False, underline=underlined)
    para(doc, "承诺人(签字)：", size=14.2, align="center", before=36, after=0, line_pt=22, line_rule="exact")
    dated = para(doc, "", align="center", before=18, after=0, line_pt=20, line_rule="exact")
    add_text(dated, "日　期：", "宋体", 14.2)
    add_text(dated, "2025 年08 月12日", "宋体", 11.2)
    rule = para(doc, "", before=28, after=8, line_pt=12, line_rule="exact")
    p_pr = rule._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    p_bdr.append(bottom)
    p_pr.append(p_bdr)
    section_title(doc, "填 表 说 明", before=28, after=10, size=25)
    blocks = [
        ["1、本表供评审各级专业技术资格使用,一式2份。"],
        ["2、填表内容应真实、准确、具体，并按表页下“注”的要求填写。", "表内填写不下时，可另加附页，并装订入内。"],
        ["3、本表请从系统中打印，一律用A3纸小册子方式打印，对折后用骑", "马钉方式装订，不得放大或缩小。"],
    ]
    for block in blocks:
        item = para(doc, "", size=14.2, align="left", before=2, after=2, line_pt=28.5, line_rule="exact", first_chars=2)
        for index, text in enumerate(block):
            if index:
                _line_break(item)
            add_text(item, text, "宋体", 14.2)


# 工作经历与参考 PDF 第 5 页一致，专业技术工作栏留空。
JOBS = [
    ("2025年03月01日至", "2025年07月10日", "江苏省建筑装饰设计研究院有限公司", "工程部", "管理岗", "", "副总经理"),
    ("2023年06月07日至", "2025年02月28日", "上海徐汇绿地商业管理有限公司", "工程部", "管理岗", "", "副总经理"),
    ("2022年04月11日至", "2023年05月13日", "上海融之旅文化旅游发展有限公司（融创文旅华东区域）", "工程部", "管理岗", "", "总监"),
    ("2018年12月10日至", "2022年03月18日", "上海万科长宁置业有限公司", "工程部", "管理岗", "", "总监"),
    ("2017年11月06日至", "2018年11月10日", "上海万限实业有限公司【万科（上海）实业发展有限公司】", "工程部", "管理岗", "", "总监"),
    ("2017年06月01日至", "2017年10月26日", "上海金科文化旅游发展集团有限公司", "工程部", "管理岗", "", "资深经理"),
    ("2017年03月20日至", "2017年05月29日", "金科（上海）建筑设计有限公司", "工程部", "管理岗", "", "资深经理"),
    ("2015年06月08日至", "2017年02月24日", "新城控股集团股份有限公司上海第一分公司", "工程部", "管理岗", "", "高级经理"),
]

# 工作总结里的职业履历，三列，与参考 PDF 第 8 页一致。
CAREER = [
    ("2025.03-至今", "江苏省建筑装饰设计研究院有限公司", "工程管理副总经理"),
    ("2023.06-2025.02", "上海徐汇绿地商业管理有限公司", "工程部副总经理"),
    ("2022.04-2023.05", "上海融之旅文化旅游发展有限公司（融创文旅华东区域）", "工程总监"),
    ("2017.11-2022.03", "上海万科企业有限公司", "工程部总监"),
    ("2017.03-2017.10", "金科地产集团有限公司华东区域公司", "工程总监"),
    ("2015.06-2017.02", "新城控股集团上海第一分公司", "工程部高级经理"),
    ("2011.07-2015.05", "江苏中南建设集团股份有限公司", "施工员/资料员"),
]

# 工作总结项目段。每项为 (文本, 首行缩进字数, 段前磅)。
PROJECT_LINES = [
    ("二、核心项目业绩与技术突破", 0, 8),
    ("1. 中南海湾新城·金石国际广场（2011.07-2015.08）", 0, 6),
    ("项目属性：超高层五星级酒店公寓（26万㎡，120m）", 2, 4),
    ("角色：施工技术管理", 2, 0),
    ("技术突破：", 2, 0),
    ("▶ 攻克核心筒框架剪力墙施工难题，通过三维建模预演优化节点处理方案", 1, 0),
    ("▶ 建立实时沉降监测系统，控制结构偏移误差＜3mm", 1, 0),
    ("成果：节约工期45天，获江苏省优质工程奖", 2, 0),
    ("2. 新城香溢紫郡（2015.12-2017.02）", 2, 8),
    ("项目属性：高端住宅（23万㎡）", 2, 0),
    ("角色：工程高级经理", 2, 0),
    ("技术突破：", 2, 0),
    ("▶ 创新桩基切顶高程控制法，降低地基处理成本23%", 1, 0),
    ("▶ 研发混凝土裂缝预警系统，楼板开裂率下降至0.8‰", 1, 0),
    ("成果：项目提前3个月交付，利润率提升12%", 2, 0),
    ("3、合肥金科博翠天辰2017.03-2017.10", 0, 8),
    ("项目属性：高端住宅综合体（总建面约34万㎡，容积率2.5）", 2, 0),
    ("角色：项目技术负责人（隶属金科集团华东区域公司）", 2, 0),
    ("技术突破与创新", 2, 0),
    ("突破一：超高层住宅剪力墙结构抗震优化", 2, 0),
    ("•解决方案与个人贡献：", 2, 0),
    ("本人行动：主导结构模型风洞试验，优化型钢排布方案，使结构刚度分布均匀化，减少扭转效应18%。", 2, 0),
    ("采用BIM预拼装技术解决型钢与钢筋密集碰撞问题，缩短关键路径工期23天。", 2, 0),
    ("突破二：高端住宅绿色科技系统集成", 2, 0),
    ("• 解决方案与个人贡献：", 2, 0),
    ("本人行动：优化管网穿剪力墙节点密封工艺，解决冷桥结露问题，系统能效比提升至4.8（国标≥3.8）。", 2, 0),
    ("▪ 嵌入金科“Life+云平台”，实现户内能耗智能调节，项目获合肥市绿色建筑示范工程认证。", 2, 0),
    ("4.上海万科天空之城（2017.11-2022.02）", 2, 8),
    ("项目属性：商业综合体+住宅（42万㎡）", 2, 0),
    ("角色：工程总监", 2, 0),
    ("技术突破：", 2, 0),
    ("▶ 首创商业-住宅协同施工模型，解决交叉作业冲突", 1, 0),
    ("▶ 应用自愈合混凝土技术于剪力墙节点", 1, 0),
    ("成果：获LEED金级认证，商业区开业首年客流超预期40%", 2, 0),
    ("5.杭州湾融创文旅城水路乐园项目（海盐）（2022.04-2023.05）", 0, 8),
    ("项目属性：高端抗震住宅", 2, 0),
    ("角色：工程总监", 2, 0),
    ("技术突破：", 2, 0),
    ("▶ 优化框架剪力墙抗震参数，地震模拟性能提升35%", 1, 0),
    ("▶ 实施无人机巡检系统，质量验收效率提升60%", 1, 0),
    ("6.苏州绿地江南映月花园（2023.06-2025.03）", 2, 8),
    ("项目属性：现状高端住宅项目", 2, 0),
    ("角色：工程部副总经理", 2, 0),
    ("技术突破：", 2, 0),
    ("▶ 应用BIM+3D打印技术实现异形钢结构精准安装", 1, 0),
    ("▶ 开发碳足迹监测平台，施工碳排放降低28%", 1, 0),
    ("6、项目名称：东台市养老服务中心地下车库装饰工程（2025.3-至今）", 0, 6),
    ("项目属性：适老化公共空间改造", 0, 0),
    ("角色：施工技术负责人", 0, 0),
    ("技术突破与创新", 0, 0),
    ("1. 适老化地下空间无障碍导视系统", 0, 0),
    ("•问题本质：老年群体空间辨识能力弱，传统车库标识系统易引发迷失焦虑。", 1, 0),
    ("• 解决方案与个人贡献：", 1, 0),
    ("▶ 地坪发光标线+AI语音联动系统：", 1, 0),
    ("▶ BIM模拟老人动线：", 1, 0),
    ("2. BIM-无人机协同进度管控", 0, 0),
    ("▶ 建立无人机实景扫描与BIM模型自动比对系统，偏差超5%自动预警；", 1, 0),
    ("▶ 成果：减少返工损失150万元（占预算8.3%），关键节点“零延误”。", 1, 0),
]


def build_basic(doc: Document) -> None:
    section_title(doc, "基  本  情  况", page=True, before=4, after=8, size=25)
    widths = fit([980, 1020, 980, 820, 860, 980, 1140, 980, 1100, 1346])
    table = make_table(doc, 16, 10, widths)
    photo = merge(table, 0, 9, 2, 9)
    merge(table, 0, 1, 0, 3)
    merge(table, 0, 7, 0, 8)
    merge(table, 1, 0, 1, 1)
    merge(table, 1, 2, 1, 8)
    merge(table, 2, 1, 2, 3)
    merge(table, 2, 4, 2, 6)
    merge(table, 2, 7, 2, 8)
    merge(table, 3, 0, 3, 2)
    merge(table, 3, 3, 3, 6)
    merge(table, 3, 7, 3, 8)
    merge(table, 4, 0, 4, 2)
    merge(table, 4, 3, 4, 9)
    merge(table, 5, 0, 5, 2)
    merge(table, 5, 3, 5, 6)
    merge(table, 5, 7, 5, 8)
    merge(table, 6, 0, 6, 2)
    merge(table, 6, 3, 6, 6)
    merge(table, 6, 7, 6, 8)
    merge(table, 7, 0, 10, 0)
    for row in range(7, 11):
        merge(table, row, 1, row, 3)
        merge(table, row, 4, row, 5)
        merge(table, row, 8, row, 9)
    merge(table, 11, 0, 15, 0)
    for row in range(11, 15):
        merge(table, row, 1, row, 5)
        merge(table, row, 6, row, 7)
        merge(table, row, 8, row, 9)
    merge(table, 15, 2, 15, 9)

    pairs = {
        (0, 0): "姓　名",
        (0, 1): "胡继刚",
        (0, 4): "性别",
        (0, 5): "男",
        (0, 6): "民族",
        (0, 7): "汉族",
        (1, 0): "公民身份号码\n（社会保障号）",
        (1, 2): "2 2 0 2 8 3 1 9 8 7 0 4 2 5 0 6 1 4",
        (2, 0): "出生年月",
        (2, 1): "1987年04月25日",
        (2, 4): "参加工作年月",
        (2, 7): "2011年06月27日",
        (3, 0): "取得本行业(专业)准入\n资格、时间及证书号码",
        (3, 3): "",
        (3, 7): "执业\n类别",
        (3, 9): "工程管理",
        (4, 0): "现专业技术资格及取得\n时间、批准单位",
        (4, 3): "上海市住房和城乡建设管理委员会批准取得工程师(20181225)",
        (5, 0): "现任专业技术职务\n及聘任时间",
        (5, 3): "职务：工程-建设工程-工程师\n聘用时间：2018-12-25",
        (5, 7): "申报\n类型",
        (5, 9): "正常申报",
        (6, 0): "现从事专业技术\n工作及年限",
        (6, 3): "现工作：\n专业技术工作年限：14年",
        (6, 7): "拟评审专业\n技术资格",
        (6, 9): "高级工程师",
        (7, 1): "学　校",
        (7, 4): "学习专业",
        (7, 6): "学历",
        (7, 7): "学位",
        (7, 8): "毕(肄、结)业年月",
        (8, 1): "复旦大学",
        (8, 4): "工商管理",
        (8, 6): "硕士研究生",
        (8, 7): "",
        (8, 8): "于2021年01月13日",
        (9, 1): "中国海洋大学青岛学院",
        (9, 4): "土木工程",
        (9, 6): "大学本科",
        (9, 7): "",
        (9, 8): "于2011年06月30日",
        (10, 1): "青岛市城阳一中",
        (10, 4): "理科",
        (10, 6): "高中",
        (10, 7): "其他",
        (10, 8): "于2016年07月04日毕业",
        (11, 1): "荣誉称号、表彰奖励名称",
        (11, 6): "获奖时间",
        (11, 8): "授奖部门",
        (12, 1): "住房政策研究创新贡献奖",
        (12, 6): "2024-12-18",
        (12, 8): "复旦大学住房政策中心专业委员会",
        (13, 1): "城市更新卓越实践奖",
        (13, 6): "2024-04-26",
        (13, 8): "绿地控股集团技术委员会",
        (14, 1): "可持续社区营造标杆项目奖",
        (14, 6): "2019-02-24",
        (14, 8): "上海万科产品创新研究院",
        (15, 1): "处分",
        (15, 2): "无",
    }
    for (row, col), text in pairs.items():
        is_label = text in {
            "姓　名", "性别", "民族", "公民身份号码\n（社会保障号）", "出生年月", "参加工作年月",
            "取得本行业(专业)准入\n资格、时间及证书号码", "执业\n类别", "现专业技术资格及取得\n时间、批准单位",
            "现任专业技术职务\n及聘任时间", "申报\n类型", "现从事专业技术\n工作及年限", "拟评审专业\n技术资格",
            "学　校", "学习专业", "学历", "学位", "毕(肄、结)业年月", "荣誉称号、表彰奖励名称", "获奖时间", "授奖部门", "处分",
        }
        fill(table.cell(row, col), text, size=12 if (row, col) == (1, 2) else (11.2 if is_label else 9.8), align="center")
    stack_vertical(table.cell(7, 0), ["学历情况"], size=9, line_pt=14)
    stack_vertical(table.cell(11, 0), ["奖惩情况"], size=9, line_pt=14)
    photo.vertical_alignment = VALIGN["center"]
    photo_p = photo.paragraphs[0]
    style_paragraph(photo_p, align="center", line_pt=1.0)
    picture(photo_p, "image1.png", 2.05)
    for index, height in enumerate([680, 680, 680, 620, 620, 680, 680, 400, 460, 460, 460, 400, 520, 520, 520, 420]):
        set_row_height(table.rows[index], height)
    apply_widths(table, widths)
    note(doc, "说明：毕(肄、结)业情况应在栏目中注明。", before=4)


def build_education(doc: Document) -> None:
    section_title(doc, "继 续 教 育 情 况", page=True, before=0, after=2, size=21)
    para(doc, "(国内外培训、进修或考察情况)", font="楷体", size=14, align="center", before=2, after=8, line_pt=20, line_rule="exact")
    widths = fit([1900, 2100, 2700, 2600, 906])
    table = make_table(doc, 4, 5, widths, header=True)
    headers = ["起止时间", "专业或主要内容", "组织单位", "学习成绩", "备注"]
    org = "南京市人力资源和社会保\n障局专业技术人员管理处"
    rows = [
        ("2025年06月09日至\n2025年06月21日", "职称继续教育专业课（工程系列）", org, "学时：240；学\n分：240", "无"),
        ("2025年05月05日至\n2025年06月09日", "公需课有效学时累计", org, "学时：152；学\n分：152", "无"),
        ("2025年05月03日至\n2025年06月18日", "知识更新工程高级研修项目", org, "学时：17；学\n分：16.5", "无"),
    ]
    for col, text in enumerate(headers):
        fill(table.cell(0, col), text, size=12)
    for r, row in enumerate(rows, start=1):
        for c, text in enumerate(row):
            fill(table.cell(r, c), text, size=9, line_pt=12, line_rule="exact")
        set_row_height(table.rows[r], 780)
    set_row_height(table.rows[0], 520)
    apply_widths(table, widths)

    section_title(doc, "专业考试成绩情况", before=18, after=8, size=21)
    exam_w = fit([1800, 2200, 2800, 1703, 1703])
    exam = make_table(doc, 2, 5, exam_w, header=True)
    for col, text in enumerate(["时　间", "组 织 单 位", "考试名称", "成　绩", "合格标志"]):
        fill(exam.cell(0, col), text, size=12)
    fill(exam.cell(1, 1), "无", size=9)
    set_row_height(exam.rows[0], 520)
    set_row_height(exam.rows[1], 700)
    apply_widths(exam, exam_w)


def build_jobs(doc: Document) -> None:
    section_title(doc, "工  作  经  历", page=True, before=0, after=2, size=25)
    para(doc, "(专业技术工作经历)", font="楷体", size=14, align="center", before=4, after=6, line_pt=20, line_rule="exact")
    # 列宽对应 PDF 竖线：80.2 / 102 / 68.3 / 68.3 / 84.7 / 68.3 pt。
    widths = [1860, 2040, 1366, 1366, 1694, 1110]
    table = make_table(doc, 1 + len(JOBS), 6, widths, header=True)
    _set_cell_margins(table, top=16, bottom=16, left=28, right=28)
    headers = ["起止时间", "工作单位", "部门", "岗位", "从事的主要\n专业技术工作", "任何职"]
    for col, text in enumerate(headers):
        fill(table.cell(0, col), text, size=12, line_pt=14, line_rule="exact")
    set_row_height(table.rows[0], 568, "exact")
    for index, job in enumerate(JOBS, start=1):
        start, end, company, dept, post, work, title = job
        values = [f"{start}\n{end}", company, dept, post, work, title]
        for col, text in enumerate(values):
            size = 9 if col == 0 else 10
            fill(table.cell(index, col), text, size=size, line_pt=12, line_rule="exact")
        set_row_height(table.rows[index], 800, "atLeast")
    apply_widths(table, widths)


def build_performance(doc: Document) -> None:
    section_title(doc, "任现职后主要专业技术工作业绩登记", page=True, before=2, after=8, size=25)
    widths = fit([1760, 3000, 4200, 1246])
    table = make_table(doc, 4, 4, widths, header=True)
    for col, text in enumerate(["起止时间", "专业技术工作\n项目、课题名称", "获奖励、专利\n及效益情况", "本人起\n何作用"]):
        fill(table.cell(0, col), text, size=12, line_pt=16, line_rule="exact")
    records = [
        (
            "2018年02月\n03日至2023\n年05月11日",
            "专业技术工作项目：杭州湾融创文旅城一水、陆乐园项目\n课题名称：杭州湾融创文旅城",
            "获奖情况：无\n专利情况：无\n效益情况：无",
            "主持",
        ),
        (
            "2019年11月\n15日至2022\n年03月25日",
            "专业技术工作项目：万科青浦区徐泾镇徐盈路西侧地块（01-06F2-04地块一、二期）\n课题名称：万科天空之城",
            "获奖情况：荣获上海万科产品创新研究院颁发的“可持续社区营造标杆项目奖”\n专利情况：专利号：ZL 2020 2 0536480.4；专利号：ZL 2020 2 0557884.1\n效益情况：在上海万科天空之城项目中，主导完成以下创新实践：1. 构建社区级“零碳”能源系统，光伏覆盖率达32%；2. 研发模块化雨水回收装置，节水效率提升45%；3. 搭建社区参与式设计平台，客户满意度达92%。",
            "主持",
        ),
        (
            "2020年10月\n15日至2024\n年01月25日",
            "专业技术工作项目：绿地苏州东太湖旅游度假区WJ-J-2020-013地块商住\n课题名称：绿地苏州东太湖旅游度假区",
            "获奖情况：荣获绿地控股集团技术委员会颁发的“城市更新卓越实践奖”\n专利情况：无\n效益情况：1. 技术体系创新：构建全流程技术标准体系，集成绿色建材与智能建造技术，单位面积改造成本降低18%，获集团技术创新标杆认证。2. 智慧工地研发：开发研发智慧工地管理系统思路，通过物联网与AI算法优化施工调度，降低人力成本15%、材料损耗率23%，并推动BIM技术应用达LOD500标准。3. 经济效益：累计节约项目成本2,800万元，投资回报率提升32%；",
            "主持",
        ),
    ]
    for r, (period, project, benefit, role) in enumerate(records, start=1):
        fill(table.cell(r, 0), period, size=9, va="center", line_pt=12, line_rule="exact")
        fill(table.cell(r, 1), project, size=9.8, align="left", va="top", line_pt=13, line_rule="exact")
        fill(table.cell(r, 2), benefit, size=9.8, align="left", va="top", line_pt=13, line_rule="exact")
        fill(table.cell(r, 3), role, size=9.8, va="center")
    set_row_height(table.rows[0], 560)
    set_row_height(table.rows[1], 1100)
    set_row_height(table.rows[2], 2200)
    set_row_height(table.rows[3], 2600)
    apply_widths(table, widths)
    note(doc, "说明：“本人起何作用”分为主持、参加、独立承担。", before=4)


def build_papers(doc: Document) -> None:
    section_title(doc, "任现职后著作、论文及重要技术报告登记", page=True, before=2, after=8, size=22)
    widths = fit([1500, 2700, 2206, 2100, 1700])
    table = make_table(doc, 4, 5, widths, header=True)
    for col, text in enumerate(["日　期", "论著(报告)标题", "刊物、出版单位\n(学术会议)名称", "主办单位", "本人承担情况"]):
        fill(table.cell(0, col), text, size=12, line_pt=16, line_rule="exact")
    rows = [
        ("2021-07-08", "铝膜板在高层住宅建筑工程中的应用", "中国民营科技促进会，《中国新技术新产品》", "中国民营科技促进会", "合著"),
        ("2021-08-05", "房屋建筑施工中人工智能技术的应用", "深圳住宅与房地产杂志社;中国房地产及住宅研究会，《住宅与房地产》", "深圳住宅与房地产杂志社;中国房地产及住宅研究会", "合著"),
        ("2024-07-10", "基于BIM的装配式钢结构建筑施工新技术研究", "中国建筑金属结构协会，《中国建筑金属结构》", "中国建筑金属结构协会", "独著"),
    ]
    for r, row in enumerate(rows, start=1):
        for c, text in enumerate(row):
            align = "left" if c in (1, 2, 3) else "center"
            fill(table.cell(r, c), text, size=9, align=align, line_pt=12, line_rule="exact")
        set_row_height(table.rows[r], 860)
    set_row_height(table.rows[0], 640)
    apply_widths(table, widths)
    note(doc, "说明：本人承担情况按独著(译)、合著(译)填写。合著(译)应注明本人实际承担的部分。", before=4)


def sline(doc, text: str, *, indent: float = 0, before: float = 0, size: float = 10.5, align: str = "left") -> None:
    """工作总结正文：五号宋体，行距 22 磅，与参考 PDF 一致。"""
    para(
        doc,
        text,
        font="宋体",
        size=size,
        align=align,
        before=before,
        after=0,
        line_pt=22,
        line_rule="exact",
        first_chars=indent or None,
    )


def build_summary(doc: Document) -> None:
    para(doc, "本人任现职以来工作总结", font="黑体", size=20, bold=True, align="center", before=2, after=0, line_pt=28, line_rule="exact")
    para(
        doc,
        "(包括思想品德、职业道德、专业技术能力、工作成绩及履行职责等，乡土人才可提供省三带情况)",
        font="楷体",
        size=10,
        align="center",
        before=6,
        after=0,
        line_pt=16,
        line_rule="exact",
    )
    para(doc, "个人专业技术工作总结", font="宋体", size=18, align="center", before=10, after=0, line_pt=26, line_rule="exact")
    sline(doc, "本文总结我在建筑施工领域的卓越业绩，参与重大项目，解决技术难题，展现专业实力。", indent=2, before=8)
    sline(doc, "提出前瞻性建议，分享学术成果，获同行认可。发表多篇论文，申请专利，为行业发展贡献智慧。以严谨态度与进取精神，积极贡献于行业。")
    sline(doc, "一、工作经历与学术背景", before=8)
    sline(doc, "1.教育背景", indent=2, before=6)
    sline(doc, "本人于2011年6月从中国海洋大学青岛学院土木工程专业毕业，于2021年1月，获得复旦大学工商管理硕士专业学历，进一步丰富了我的知识体系，为我在建筑和金融领域的交叉发展奠定了坚实基础。", indent=2, before=4)
    sline(doc, "2.职业履历", indent=2, before=8)
    widths = [1900, 5000, 2536]
    table = make_table(doc, 1 + len(CAREER), 3, widths, header=True)
    for col, text in enumerate(["时间", "公司名称", "职务"]):
        cell = table.cell(0, col)
        fill(cell, text, size=10.5)
        shade_cell(cell, "D9D9D9")
    for index, row in enumerate(CAREER, start=1):
        for col, text in enumerate(row):
            fill(table.cell(index, col), text, size=10.5, align="center", line_pt=16, line_rule="exact")
        set_row_height(table.rows[index], 520)
    set_row_height(table.rows[0], 420)
    apply_widths(table, widths)
    sline(doc, "3.社会职务", indent=2, before=8)
    sline(doc, "复旦大学住房政策研究中心秘书长（研究员）", indent=2)
    sline(doc, "上海市杨浦区科技企业联合会执行会长", indent=2)
    sline(doc, "以上丰富的工作经历与学术背景使我具备了深厚的专业素养和广泛的行业影响力，我将继续以敬业的态度和专业的精神，为建筑和金融领域的发展贡献自己的力量。", indent=2, before=4)
    for text, indent, before in PROJECT_LINES:
        sline(doc, text, indent=indent, before=before)
    sline(doc, "三、技术创新与学术贡献", before=8)
    sline(doc, "1.知识产权", indent=2, before=6)
    sline(doc, "a发明专利《建筑工程管理用功能脚架》（ZL202020536480.4），解决高空作业安全防护难题，已在12个项目推广；", indent=2)
    sline(doc, "b实用新型专利《建筑工程防护围栏》（ZL202020557884.1），模块化设计使安装效率提升200%", indent=2)
    sline(doc, "2.学术论著", indent=2, before=8)
    paper_w = [1400, 5200, 2836]
    papers = make_table(doc, 4, 3, paper_w, header=True)
    for col, text in enumerate(["时间", "论文标题", "期刊"]):
        fill(papers.cell(0, col), text, size=12)
    paper_rows = [
        ("2021.07", "《铝模板在高层住宅建筑工程中的应用》", "《中国新技术新产品》"),
        ("2021.08", "《人工智能技术在房屋建筑工程中的应用》", "《住宅与房地产》"),
        ("2024.07", "《BIM装配式钢结构建筑施工技术与管理研究》", "《中国建筑金属结构》"),
    ]
    for r, row in enumerate(paper_rows, start=1):
        for c, text in enumerate(row):
            fill(papers.cell(r, c), text, size=10.5, align="center" if c == 0 else "left", line_pt=16, line_rule="exact")
        set_row_height(papers.rows[r], 560)
    set_row_height(papers.rows[0], 400)
    apply_widths(papers, paper_w)
    sline(doc, "学术交流", indent=2, before=8)
    sline(doc, "a2022.07 上海市建筑施工技术交流会", indent=2)
    sline(doc, "主题演讲：《绿色建筑施工技术创新路径》", indent=1)
    sline(doc, "b2022.11 江苏省建筑技术研讨会", indent=2, before=2)
    sline(doc, "发表演讲：《3D打印技术在预制建筑中的实践突破》", indent=1)
    sline(doc, "四、专业能力总结", before=8)
    for line in [
        "1.技术攻坚：在超高层结构、抗震优化、绿色建造等领域拥有9项核心技术",
        "2.管理创新：独创“动态节点管理法”实现多业态项目协同开发",
        "3.行业贡献：推动BIM技术与ESG理念在长三角项目落地，累计节约工程造价超1.2亿元",
        "4.发展愿景：聚焦智能建造与低碳技术融合，助力建筑业数字化转型",
    ]:
        sline(doc, line, indent=2)
    sline(doc, "本人签名 _____________________", align="right", size=12, before=16)
    sline(doc, "年　　月　　日", align="right", size=12, before=2)


def build_assessment(doc: Document) -> None:
    section_title(doc, "任现职以来考核情况", page=False, before=2, after=8, size=25)
    widths = fit([900, 1500, 3600, 1500, 1700, 1006])
    table = make_table(doc, 4, 6, widths, header=True)
    merge(table, 0, 0, 3, 0)
    stack_vertical(table.cell(0, 0), ["年度（任期）考核情况"], size=9, line_pt=13)
    for col, text in enumerate(["年　份", "受聘专业技术职务(岗位)", "考核等次", "考核单位", "备注"], start=1):
        fill(table.cell(0, col), text, size=9, bold=True)
    fill(table.cell(1, 2), "无", size=10.5)
    for row in range(1, 4):
        set_row_height(table.rows[row], 900)
    set_row_height(table.rows[0], 700)
    apply_widths(table, widths)


def build_verify(doc: Document) -> None:
    section_title(doc, "工作单位或行业协会（学会）或人事档案管理单位核实情况", page=True, before=4, after=8, size=15)
    widths = fit([1700, 4253, 4253])
    table = make_table(doc, 3, 3, widths)
    merge(table, 0, 0, 2, 0)
    merge(table, 0, 1, 0, 2)
    stack_vertical(table.cell(0, 0), ["单位或行业协会（学会）或", "人事档案管理单位核实情况"], size=8, line_pt=12)
    fill(table.cell(0, 1), "经核查，申报人所填内容及提交的材料属实。", font="黑体", size=14, bold=True)
    fill(table.cell(1, 1), "核 实 人(签字)", size=10.5, bold=True)
    fill(table.cell(1, 2), "单位(印章)", size=10.5, bold=True)
    fill(table.cell(2, 1), "日　期", size=10.5)
    fill(table.cell(2, 2), "日　期", size=10.5)
    set_row_height(table.rows[0], 1100)
    set_row_height(table.rows[1], 640)
    set_row_height(table.rows[2], 1500)
    apply_widths(table, widths)

    section_title(doc, "组  织  意  见", before=16, after=8, size=21)
    org_w = [980, 4123, 980, 4123]
    org = make_table(doc, 1, 4, org_w)
    stack_vertical(org.cell(0, 0), ["县（市）、区", "职称部门或市主管部门意见"], size=8, line_pt=12)
    stack_vertical(org.cell(0, 2), ["市、厅（局）人事职称部门意见"], size=8, line_pt=12)
    fill(
        org.cell(0, 1),
        "区县行政主管章\n区县行业主管章\n区县职称办盖章\n市行政主管盖章\n申报点盖章\n\n公　章\n\n年　　月　　日",
        size=10.5,
        va="center",
    )
    right = org.cell(0, 3)
    right.vertical_alignment = VALIGN["center"]
    lead = right.paragraphs[0]
    style_paragraph(lead, align="center", before=2, after=2, line_pt=1.0)
    add_text(lead, "市职称办盖章", "宋体", 10.5, True)
    pic_p = right.add_paragraph()
    style_paragraph(pic_p, align="center", before=2, after=2, line_pt=1.0)
    picture(pic_p, "image2.png", 2.55)
    sign_p = right.add_paragraph()
    style_paragraph(sign_p, align="center", before=2, after=0, line_pt=1.0)
    add_text(sign_p, "负责人(签字)　　公　章", "宋体", 10.5)
    date_p = right.add_paragraph()
    style_paragraph(date_p, align="center", before=2, after=2, line_pt=1.0)
    add_text(date_p, "2025年08月12日", "宋体", 10.5)
    set_row_height(org.rows[0], 4600)
    apply_widths(org, org_w)


def _vote_table(cell, attend: str, present: str, yes: str, no: str, abstain: str):
    lead = cell.paragraphs[0]
    style_paragraph(lead, align="center", line_pt=2, line_rule="exact")
    widths = fit([1200, 560, 1200, 560, 1100, 560, 1100, 560, 1100, 860])
    table = cell.add_table(1, 10)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_tbl_borders(table, sz="4")
    _set_cell_margins(table, top=30, bottom=30, left=40, right=40)
    apply_widths(table, widths)
    values = ["应到人数", attend, "实到人数", present, "赞成票", yes, "反对票", no, "弃权票", abstain]
    for index, text in enumerate(values):
        bold = index % 2 == 0
        fill(table.cell(0, index), text, size=9, bold=bold)
    apply_widths(table, widths)
    return table


def _plain_paragraph(cell, text: str = "", *, size=10.5, bold=False, align="center", before=2, after=2, font="宋体"):
    paragraph = cell.add_paragraph()
    style_paragraph(paragraph, align=align, before=before, after=after, line_pt=1.15, font_pt=size)
    if text:
        add_text(paragraph, text, font, size, bold)
    return paragraph


def build_review(doc: Document) -> None:
    section_title(doc, "评审、登记备案情况", page=True, before=2, after=8, size=21)
    widths = fit([1000, 4603, 4603])
    table = make_table(doc, 6, 3, widths)
    merge(table, 0, 1, 0, 2)
    merge(table, 1, 1, 1, 2)
    merge(table, 2, 1, 2, 2)
    merge(table, 3, 0, 5, 0)
    merge(table, 5, 1, 5, 2)
    stack_vertical(table.cell(0, 0), ["专业（学科）评议组意见"], size=8, line_pt=12)
    stack_vertical(table.cell(1, 0), ["评审委员会评审意见"], size=8, line_pt=13)
    stack_vertical(table.cell(2, 0), ["省、市、县或相应", "职称主管部门登记备案情况"], size=8, line_pt=12)
    stack_vertical(table.cell(3, 0), ["备注"], size=9, line_pt=14)

    group = table.cell(0, 1)
    group.vertical_alignment = VALIGN["top"]
    _vote_table(group, "3", "3", "3", "0", "0")
    sign = group.add_paragraph()
    style_paragraph(sign, align="left", before=8, after=0, line_pt=1.0)
    picture(sign, "image3.png", 2.3)
    add_text(sign, "　　评议组组长(签字)", "宋体", 10.5)
    date = group.add_paragraph()
    style_paragraph(date, align="right", before=2, after=4, line_pt=1.0)
    add_text(date, "2025年09月21日", "宋体", 10.5)

    committee = table.cell(1, 1)
    committee.vertical_alignment = VALIGN["top"]
    _vote_table(committee, "27", "27", "27", "0", "0")
    opinion = _plain_paragraph(
        committee,
        "经评审，胡继刚（220283198704250614）已具备建设工程·科技管理·工程管理专业（学科）高级工程师资格。",
        size=12,
        align="left",
        before=8,
        after=6,
    )
    opinion.alignment = ALIGN["left"]
    style_paragraph(opinion, align="left", before=8, after=4, line_pt=1.3, first_chars=2, font_pt=12)
    line = committee.add_paragraph()
    style_paragraph(line, align="left", before=4, after=0, line_pt=1.0)
    picture(line, "image4.png", 2.4)
    add_text(line, "　主任委员(签字)", "宋体", 10.5)
    picture(line, "image5.png", 2.15)
    add_text(line, "　评委会盖章　　公　章", "宋体", 10.5)
    cdate = committee.add_paragraph()
    style_paragraph(cdate, align="right", before=2, after=2, line_pt=1.0)
    add_text(cdate, "2025年10月19日", "宋体", 10.5)

    filing = table.cell(2, 1)
    filing.vertical_alignment = VALIGN["top"]
    lead = filing.paragraphs[0]
    style_paragraph(lead, align="center", line_pt=2, line_rule="exact")
    stamps = filing.add_table(2, 2)
    stamps.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_tbl_borders(stamps, val="nil")
    _set_cell_margins(stamps, top=20, bottom=20, left=40, right=40)
    apply_widths(stamps, [4500, 4500])
    fill(stamps.cell(0, 0), "签发单位盖章", size=10.5, bold=True)
    fill(stamps.cell(0, 1), "备案部门盖章", size=10.5, bold=True)
    for col, image, width in ((0, "image6.png", 2.35), (1, "image7.png", 2.35)):
        cell = stamps.cell(1, col)
        set_tc_borders(cell)
        cell.vertical_alignment = VALIGN["center"]
        p = cell.paragraphs[0]
        style_paragraph(p, align="center", line_pt=1.0)
        picture(p, image, width)
    seal = filing.add_paragraph()
    style_paragraph(seal, align="center", before=2, after=0, line_pt=1.0)
    add_text(seal, "公　章", "宋体", 10.5)
    fdate = filing.add_paragraph()
    style_paragraph(fdate, align="right", before=2, after=2, line_pt=1.0)
    add_text(fdate, "2025年12月09日", "宋体", 10.5)

    fill(table.cell(3, 1), "人员姓名：胡继刚", size=10.5, align="left")
    fill(table.cell(3, 2), "证件号码：220283198704250614", size=10.5, align="left")
    fill(table.cell(4, 1), "联系电话：18678408669", size=10.5, align="left")
    fill(table.cell(4, 2), "流程编号：2506170651940505", size=10.5, align="left")
    fill(table.cell(5, 1), "工作单位：江苏省建筑装饰设计研究院有限公司", size=10.5, align="left")
    for row, height in ((0, 2500), (1, 3200), (2, 3000), (3, 420), (4, 420), (5, 420)):
        set_row_height(table.rows[row], height)
    apply_widths(table, widths)
    note(doc, "说明：“评审委员会评审意见”栏，应在“具备”前写明“已”或“不”。", before=4)


def build() -> Path:
    doc = Document()
    prepare(doc)
    configure(doc.sections[0])
    build_cover(doc)
    build_pledge(doc)
    build_basic(doc)
    build_education(doc)
    build_jobs(doc)
    build_performance(doc)
    build_papers(doc)
    summary = doc.add_section(WD_SECTION.NEW_PAGE)
    configure(summary, border=True, top=1.15, bottom=1.10)
    build_summary(doc)
    rest = doc.add_section(WD_SECTION.NEW_PAGE)
    configure(rest)
    build_assessment(doc)
    build_verify(doc)
    build_review(doc)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUT)
    return OUT


def self_check(path: Path) -> None:
    doc = Document(path)
    chunks: list[str] = []
    bold_intro = None
    intro_size = None
    for paragraph in doc.paragraphs:
        chunks.append(paragraph.text)
        if paragraph.text.startswith("本文总结") and paragraph.runs:
            bold_intro = paragraph.runs[0].bold
            intro_size = paragraph.runs[0].font.size.pt
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                chunks.append(cell.text)
                for nested in cell.tables:
                    for nrow in nested.rows:
                        for ncell in nrow.cells:
                            chunks.append(ncell.text)
    text = "\n".join(chunks)
    required = [
        "胡继刚",
        "2 2 0 2 8 3 1 9 8 7 0 4 2 5 0 6 1 4",
        "220283198704250614",
        "江苏省建筑装饰设计研究院有限公司",
        "2025年07月10日",
        "上海徐汇绿地商业管理有限公司",
        "上海万限实业有限公司",
        "资深经理",
        "工程管理副总经理",
        "江苏省优质工程奖",
        "高级工程师",
        "复旦大学",
        "中国海洋大学青岛学院",
        "职称继续教育专业课（工程系列）",
        "专业技术人员管理处",
        "智慧工地研发",
        "2018年02月",
        "ZL 2020 2 0536480.4",
        "铝模板在高层住宅建筑工程中的应用",
        "房屋建筑施工中人工智能技术的应用",
        "独著",
        "18678408669",
        "2506170651940505",
        "应到人数",
        "已具备",
        "个 人 承 诺 书",
        "东台市养老服务中心",
        "赞成票",
        "从事的主要",
    ]
    missing = [item for item in required if item not in text]
    if missing:
        raise SystemExit("缺少内容：" + "、".join(missing))
    if "上海惟嘉" in text or "工程技术及项目管理" in text:
        raise SystemExit("工作经历仍是旧稿")
    if "认 管 证" in text or "职称继 续" in text:
        raise SystemExit("仍有断行或粘连错误")
    for section in doc.sections:
        header = "".join(paragraph.text for paragraph in section.header.paragraphs).strip()
        if header:
            raise SystemExit("页眉未去掉：" + header)
    if bold_intro is True:
        raise SystemExit("工作总结正文仍为粗体")
    if intro_size != 10.5:
        raise SystemExit(f"工作总结字号异常：{intro_size}")
    images = [rel for rel in doc.part.rels.values() if "image" in rel.reltype]
    if len(images) < 7:
        raise SystemExit(f"图片数量不足：{len(images)}")
    print(f"内容检查通过，图片 {len(images)} 张，文本约 {len(text)} 字")


def main() -> None:
    path = build()
    self_check(path)
    print(path)


if __name__ == "__main__":
    main()
