#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成可签署的《劳动合同书》：上海贝迪创建材科技有限公司 — 胡继刚。"""

from __future__ import annotations

import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "output"
DOCX_NAME = "劳动合同书_上海贝迪创建材科技有限公司_胡继刚.docx"

SONG = "宋体"
HEI = "黑体"
TNR = "Times New Roman"
INK = "000000"
RULE = "333333"

COMPANY = "上海贝迪创建材科技有限公司"
CREDIT = "91310113MAERPR7M1D"
LEGAL_REP = "刘李杨"
REG_ADDR = "上海市宝山区蕰川路5475号4幢部分"
OFFICE_ADDR = "上海市浦东新区康桥东路111号13号楼一层"
CONTACT = "段鹏涛"
CONTACT_PHONE = "13916901963"

EMPLOYEE = "胡继刚"
BIRTH = "1987年4月"
ID_NO = "220283198704250614"
GENDER = "男"
HUKOU = "山东省青岛市李沧区黑龙江中路629号2-2-2502"
HOME = "上海市杨浦区爱国路389号1-1-401"
PHONE = "13262607888"
EMAIL = "hujigang2010@gmail.com"

DEPT = "工程管理部"
POST = "高级工程师"
RANK = "副总经理"
WAGE = "28,000"
WAGE_DX = "人民币贰万捌仟元整"
CONTRACT_NO = "BDCC-2026-001"


def set_run_font(run, east: str, size: float, bold: bool = False) -> None:
    run.bold = bold
    run.italic = False
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(INK)
    run.font.name = TNR
    rpr = run._element.get_or_add_rPr()
    r_fonts = rpr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        rpr.append(r_fonts)
    r_fonts.set(qn("w:ascii"), TNR)
    r_fonts.set(qn("w:hAnsi"), TNR)
    r_fonts.set(qn("w:cs"), TNR)
    r_fonts.set(qn("w:eastAsia"), east)
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:eastAsia"), "zh-CN")


def set_paragraph(
    paragraph,
    *,
    before: float = 0,
    after: float = 0,
    line: float = 1.15,
    align: str = "left",
    first_indent: float = 0,
    left_indent: float = 0,
    keep_next: bool = False,
    keep_lines: bool = False,
    page_break: bool = False,
) -> None:
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.line_spacing = line
    fmt.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    fmt.first_line_indent = Cm(first_indent) if first_indent else Pt(0)
    fmt.left_indent = Cm(left_indent) if left_indent else Pt(0)
    fmt.widow_control = True
    fmt.page_break_before = page_break
    paragraph.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }[align]
    ppr = paragraph._p.get_or_add_pPr()
    if keep_next and ppr.find(qn("w:keepNext")) is None:
        ppr.append(OxmlElement("w:keepNext"))
    if keep_lines and ppr.find(qn("w:keepLines")) is None:
        ppr.append(OxmlElement("w:keepLines"))


def add_text(paragraph, text: str, east: str, size: float, bold: bool = False):
    run = paragraph.add_run(text)
    set_run_font(run, east, size, bold)
    return run


def add_bottom_border(paragraph, sz: str = "8") -> None:
    ppr = paragraph._p.get_or_add_pPr()
    borders = ppr.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        ppr.append(borders)
    line = OxmlElement("w:bottom")
    line.set(qn("w:val"), "single")
    line.set(qn("w:sz"), sz)
    line.set(qn("w:space"), "1")
    line.set(qn("w:color"), "000000")
    borders.append(line)


def add_top_border(paragraph, sz: str = "6") -> None:
    ppr = paragraph._p.get_or_add_pPr()
    borders = ppr.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        ppr.append(borders)
    line = OxmlElement("w:top")
    line.set(qn("w:val"), "single")
    line.set(qn("w:sz"), sz)
    line.set(qn("w:space"), "1")
    line.set(qn("w:color"), "000000")
    borders.append(line)


def add_field(paragraph, instruction: str) -> None:
    def fld(kind: str):
        run = paragraph.add_run()
        set_run_font(run, SONG, 9)
        node = OxmlElement("w:fldChar")
        node.set(qn("w:fldCharType"), kind)
        run._r.append(node)

    fld("begin")
    run = paragraph.add_run()
    set_run_font(run, SONG, 9)
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruction} "
    run._r.append(instr)
    fld("separate")
    add_text(paragraph, "1", SONG, 9)
    fld("end")


def shade(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def cell_margins(cell, top=60, bottom=60, left=80, right=80) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.find(qn("w:tcMar"))
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for edge, value in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        node = tc_mar.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_fixed(table, widths_cm: list[float]) -> None:
    widths = [int(round(w * 567)) for w in widths_cm]
    table.autofit = False
    table.allow_autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(sum(widths)))
    tbl_w.set(qn("w:type"), "dxa")
    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")
    jc = tbl_pr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        tbl_pr.append(jc)
    jc.set(qn("w:val"), "center")
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tbl_pr.addnext(grid)
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)
    for row in table.rows:
        tr = row._tr
        tr_pr = tr.get_or_add_trPr()
        if tr_pr.find(qn("w:cantSplit")) is None:
            tr_pr.append(OxmlElement("w:cantSplit"))
        for index, cell in enumerate(row.cells):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(widths[index]))
            tc_w.set(qn("w:type"), "dxa")


def set_borders(table, color: str = RULE, sz: str = "6") -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), sz)
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), color)


def clear_borders(table) -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "nil")


def set_row_height(row, height_cm: float, rule: str = "atLeast") -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tr_height = tr_pr.find(qn("w:trHeight"))
    if tr_height is None:
        tr_height = OxmlElement("w:trHeight")
        tr_pr.append(tr_height)
    tr_height.set(qn("w:val"), str(int(height_cm * 567)))
    tr_height.set(qn("w:hRule"), rule)


def write_cell(cell, text: str, *, bold: bool = False, size: float = 10.5, align: str = "left", fill: str | None = None) -> None:
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    cell_margins(cell)
    if fill:
        shade(cell, fill)
    paragraph = cell.paragraphs[0]
    set_paragraph(paragraph, before=0, after=0, line=1.0, align=align)
    if paragraph.runs:
        paragraph.runs[0].text = ""
    add_text(paragraph, text, SONG, size, bold)


def party_table(doc, rows: list[tuple[str, str]], widths: list[float]) -> None:
    table = doc.add_table(rows=len(rows), cols=2)
    set_table_fixed(table, widths)
    set_borders(table)
    for i, (label, value) in enumerate(rows):
        set_row_height(table.rows[i], 0.78)
        write_cell(table.cell(i, 0), label, bold=True, align="center", fill="F4F4F4")
        write_cell(table.cell(i, 1), value, align="left")
    spacer(doc, 8)


def employee_table(doc) -> None:
    widths = [3.3, 4.3, 3.4, 5.2]
    data = [
        ["姓名", EMPLOYEE, "性别", GENDER],
        ["出生年月", BIRTH, "身份证号码", ID_NO],
        ["户籍地址", HUKOU, "", ""],
        ["现居住地址", HOME, "", ""],
        ["联系电话", PHONE, "电子邮箱", EMAIL],
    ]
    table = doc.add_table(rows=len(data), cols=4)
    set_table_fixed(table, widths)
    set_borders(table)
    for r, row_data in enumerate(data):
        set_row_height(table.rows[r], 0.78)
        if r in (2, 3):
            table.cell(r, 1).merge(table.cell(r, 3))
            write_cell(table.cell(r, 0), row_data[0], bold=True, align="center", fill="F4F4F4")
            write_cell(table.cell(r, 1), row_data[1], align="left")
        else:
            for c, text in enumerate(row_data):
                write_cell(
                    table.cell(r, c),
                    text,
                    bold=c % 2 == 0,
                    align="center" if c % 2 == 0 else "left",
                    fill="F4F4F4" if c % 2 == 0 else None,
                )
    spacer(doc, 6)


def spacer(doc, after: float = 6) -> None:
    paragraph = doc.add_paragraph()
    set_paragraph(paragraph, before=0, after=after, line=1.0)
    add_text(paragraph, "", SONG, 6)


def heading_party(doc, text: str) -> None:
    paragraph = doc.add_paragraph()
    set_paragraph(paragraph, before=2, after=4, line=1.0, keep_next=True)
    add_text(paragraph, text, HEI, 12, True)


def chapter(doc, text: str) -> None:
    paragraph = doc.add_paragraph()
    set_paragraph(paragraph, before=11, after=3, line=1.15, keep_next=True, keep_lines=True)
    add_text(paragraph, text, HEI, 14, True)


def article(doc, text: str, keep_next: bool = False) -> None:
    paragraph = doc.add_paragraph()
    set_paragraph(
        paragraph,
        before=1,
        after=1,
        line=1.35,
        align="justify",
        first_indent=0.74,
        keep_next=keep_next,
        keep_lines=True,
    )
    match = re.match(r"(第[一二三四五六七八九十百零]+条)(.*)", text)
    if match:
        add_text(paragraph, match.group(1), SONG, 12, True)
        add_text(paragraph, match.group(2), SONG, 12, False)
    else:
        add_text(paragraph, text, SONG, 12, False)


def item(doc, text: str, last: bool = False) -> None:
    paragraph = doc.add_paragraph()
    set_paragraph(
        paragraph,
        before=0,
        after=1 if last else 0,
        line=1.35,
        align="justify",
        left_indent=0.74,
        keep_next=not last,
        keep_lines=True,
    )
    add_text(paragraph, text, SONG, 12, False)


def configure(doc: Document) -> None:
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.4)
    section.right_margin = Cm(2.4)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(1.8)
    section.header_distance = Cm(0.6)
    section.footer_distance = Cm(0.5)
    section.different_first_page_header_footer = True

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    set_paragraph(hp, align="left", before=0, after=2, line=1.0)
    ppr = hp._p.get_or_add_pPr()
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:pos"), str(int(16.2 * 567)))
    tabs.append(tab)
    ppr.append(tabs)
    add_text(hp, COMPANY, SONG, 9)
    add_text(hp, "\t劳动合同书", SONG, 9)
    add_bottom_border(hp, "8")

    first_header = section.first_page_header
    first_header.is_linked_to_previous = False
    fp_h = first_header.paragraphs[0]
    set_paragraph(fp_h, before=0, after=0, line=1.0)
    add_text(fp_h, "", SONG, 9)

    def fill_footer(footer) -> None:
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0]
        set_paragraph(fp, align="center", before=2, after=0, line=1.0)
        add_top_border(fp, "6")
        add_text(fp, "第 ", SONG, 9)
        add_field(fp, "PAGE")
        add_text(fp, " 页  共 ", SONG, 9)
        add_field(fp, "NUMPAGES")
        add_text(fp, " 页", SONG, 9)

    fill_footer(section.footer)
    fill_footer(section.first_page_footer)

    normal = doc.styles["Normal"]
    normal.font.name = TNR
    normal.font.size = Pt(12)
    normal.font.color.rgb = RGBColor.from_string(INK)
    rpr = normal.element.get_or_add_rPr()
    r_fonts = rpr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        rpr.append(r_fonts)
    r_fonts.set(qn("w:ascii"), TNR)
    r_fonts.set(qn("w:hAnsi"), TNR)
    r_fonts.set(qn("w:cs"), TNR)
    r_fonts.set(qn("w:eastAsia"), SONG)

    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        node = OxmlElement("w:updateFields")
        node.set(qn("w:val"), "true")
        settings.append(node)

    core = doc.core_properties
    core.title = "劳动合同书"
    core.subject = f"{COMPANY} {EMPLOYEE}"
    core.author = COMPANY
    core.category = "劳动合同"


def signature_page(doc) -> None:
    title = doc.add_paragraph()
    set_paragraph(title, before=0, after=8, line=1.0, align="center", page_break=True)
    add_text(title, "（以下无正文）", SONG, 12)

    intro = doc.add_paragraph()
    set_paragraph(intro, before=0, after=8, line=1.15, align="justify", first_indent=0.74, keep_next=True)
    add_text(intro, "双方确认已阅读本合同全部条款。甲方盖章并由法定代表人或授权代表签字，乙方本人签字后，本合同生效。", SONG, 12)

    table = doc.add_table(rows=1, cols=2)
    set_table_fixed(table, [8.1, 8.1])
    clear_borders(table)
    set_row_height(table.rows[0], 7.2)
    left = table.cell(0, 0)
    right = table.cell(0, 1)
    left.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    cell_margins(left, top=40, bottom=40, left=0, right=120)
    cell_margins(right, top=40, bottom=40, left=120, right=0)

    def sign_line(cell) -> None:
        paragraph = cell.add_paragraph()
        set_paragraph(paragraph, before=1, after=2, line=1.0)
        add_bottom_border(paragraph, "8")
        add_text(paragraph, " ", SONG, 14)

    def block(cell, heading: str, name: str, sign_label: str) -> None:
        first = cell.paragraphs[0]
        set_paragraph(first, before=0, after=2, line=1.15)
        add_text(first, heading, SONG, 12, True)
        name_p = cell.add_paragraph()
        set_paragraph(name_p, before=2, after=0, line=1.15)
        add_text(name_p, name, SONG, 12, False)
        label = cell.add_paragraph()
        set_paragraph(label, before=46, after=0, line=1.15)
        add_text(label, sign_label, SONG, 12, False)
        sign_line(cell)
        date = cell.add_paragraph()
        set_paragraph(date, before=10, after=0, line=1.15)
        add_text(date, "日期：________年____月____日", SONG, 12, False)

    block(left, "甲方（盖章）", COMPANY, "法定代表人或授权代表（签字）")
    block(right, "乙方（签字）", EMPLOYEE, "本人签字")

    place = doc.add_paragraph()
    set_paragraph(place, before=10, after=2, line=1.15, align="left")
    add_text(place, "签署地点：上海市浦东新区", SONG, 12)

    line = doc.add_paragraph()
    set_paragraph(line, before=8, after=4, line=1.0)
    add_bottom_border(line, "6")

    receipt_title = doc.add_paragraph()
    set_paragraph(receipt_title, before=6, after=2, line=1.15)
    add_text(receipt_title, "合同文本领取确认", HEI, 12, True)

    receipt = doc.add_paragraph()
    set_paragraph(receipt, before=2, after=8, line=1.35, align="justify", first_indent=0.74)
    add_text(receipt, "乙方确认：已领取双方签署的本合同文本一份。", SONG, 12)

    sign = doc.add_paragraph()
    set_paragraph(sign, before=4, after=0, line=1.5)
    add_text(sign, "乙方签字：________________    日期：________年____月____日", SONG, 12)


def build() -> Document:
    doc = Document()
    configure(doc)

    title = doc.add_paragraph()
    set_paragraph(title, before=0, after=2, line=1.0, align="center")
    add_text(title, "劳  动  合  同  书", HEI, 22, True)

    sub = doc.add_paragraph()
    set_paragraph(sub, before=0, after=2, line=1.0, align="center")
    add_text(sub, "（固定期限）", SONG, 12)

    number = doc.add_paragraph()
    set_paragraph(number, before=0, after=8, line=1.0, align="center")
    add_bottom_border(number, "12")
    add_text(number, f"合同编号：{CONTRACT_NO}", SONG, 10.5)

    heading_party(doc, "甲方（用人单位）")
    party_table(
        doc,
        [
            ("名称", COMPANY),
            ("统一社会信用代码", CREDIT),
            ("法定代表人", LEGAL_REP),
            ("注册地址", REG_ADDR),
            ("办公地址", OFFICE_ADDR),
            ("联系人", f"{CONTACT}　　联系电话：{CONTACT_PHONE}"),
        ],
        [4.2, 12.0],
    )

    heading_party(doc, "乙方（劳动者）")
    employee_table(doc)

    lead = doc.add_paragraph()
    set_paragraph(lead, before=4, after=2, line=1.35, align="justify", first_indent=0.74)
    add_text(
        lead,
        "根据《中华人民共和国劳动法》《中华人民共和国劳动合同法》及其他有关法律、法规和上海市现行规定，甲乙双方在平等自愿、协商一致的基础上订立本合同，共同遵守。",
        SONG,
        12,
    )

    chapter(doc, "第一章　劳动合同期限和试用期")
    article(doc, "第一条　本合同为固定期限劳动合同，期限自2026年5月1日起至2029年4月30日止，共3年。劳动关系自实际用工之日起建立。")
    article(doc, "第二条　试用期自2026年5月1日起至2026年7月31日止，共三个月，包含在上述劳动合同期限内。试用期届满，甲方未依法解除本合同的，乙方即为试用合格。同一用人单位与同一劳动者只能约定一次试用期。")
    article(doc, "第三条　甲方应事先向乙方明确告知录用条件及试用期考核要求。甲方在试用期内解除本合同，应当说明理由，并证明乙方不符合录用条件。乙方在试用期内提前3日通知甲方，可以解除本合同。")

    chapter(doc, "第二章　工作内容和工作地点")
    article(doc, f"第四条　乙方工作部门为{DEPT}，岗位（专业技术职务）为{POST}，职级为{RANK}。乙方主要从事建筑工程技术及项目管理工作，按照岗位职责和甲方的工作安排履行职责。")
    article(doc, "第五条　乙方主要工作职责包括：")
    item(doc, "（一）工程项目技术管理：组织或参与技术交底、技术审查，协调解决实施过程中的工程技术问题；")
    item(doc, "（二）施工组织协调：参与施工组织设计和进度安排，协调设计、施工、监理等相关单位的技术配合；")
    item(doc, "（三）质量安全控制：落实工程质量及安全技术要求，开展检查、问题整改与风险防范；")
    item(doc, "（四）工程技术方案管理：组织或参与技术方案的编制、论证、优化、实施检查与技术资料管理；")
    item(doc, "（五）项目全过程管理：参与项目策划、建设实施、验收移交等阶段的技术和项目管理工作。", last=True)
    article(doc, f"第六条　乙方工作地点为上海市，主要工作场所为{OFFICE_ADDR}。甲方因项目需要在上海市内调整具体工作场所的，应当提前告知乙方。安排乙方到上海市以外临时出差的，甲方承担差旅费用。长期变更工作城市的，由双方书面协商确认。变更工作部门、岗位、职级或劳动报酬的，由双方协商一致并签订书面协议。")

    chapter(doc, "第三章　工作时间和休息休假")
    article(doc, "第七条　乙方实行标准工时制，每日工作8小时、每周工作40小时，每周至少休息1日。作息安排由甲方确定并告知乙方。")
    article(doc, "第八条　甲方安排乙方休息日、法定节假日、带薪年休假及依法享有的其他假期。甲方安排加班应当与乙方协商，并依法支付加班工资或安排补休。")

    chapter(doc, "第四章　劳动报酬")
    article(doc, f"第九条　乙方转正后正常出勤月工资为人民币{WAGE}元（税前，大写：{WAGE_DX}）。上述月工资对应法定工作时间内的正常劳动。奖金、津贴、补贴按甲方依法制定并告知乙方的薪酬制度执行，不计入上述月工资。加班工资另计。")
    article(doc, f"第十条　乙方试用期月工资为人民币{WAGE}元（税前，大写：{WAGE_DX}），与转正后工资标准相同，不降低。乙方提供正常劳动的，试用期工资不得低于上海市当期最低工资标准。")
    article(doc, "第十一条　甲方每月15日以银行转账方式，将上一自然月工资支付至乙方本人银行账户，并向乙方提供工资清单。发薪日遇休息日或法定节假日的，提前至最近一个工作日支付。")
    article(doc, "第十二条　甲方依法代扣代缴个人所得税，以及乙方个人应当缴纳的社会保险费和住房公积金。乙方提供正常劳动的，工资不得低于上海市当期最低工资标准。")
    article(doc, "第十三条　甲方安排乙方在工作日延长工作时间的，支付不低于工资150%的加班工资；安排乙方在休息日工作又不能安排补休的，支付不低于工资200%的加班工资；安排乙方在法定节假日工作的，支付不低于工资300%的加班工资。加班工资以本合同约定的月工资为计算基数。日工资按月工资除以21.75计算，小时工资按日工资除以8小时计算。")
    article(doc, "第十四条　乙方依法休假、患病或非因工负伤，以及甲方停工停产期间的工资待遇，按照国家和上海市规定执行。调整本合同约定的工资标准，由双方协商一致并采用书面形式。")

    chapter(doc, "第五章　社会保险和住房公积金")
    article(doc, "第十五条　甲方自实际用工之日起，为乙方办理社会保险登记并按时足额缴纳职工基本养老保险、职工基本医疗保险、失业保险和工伤保险。乙方按照上海市规定享受生育保险待遇。试用期计入缴费期间。乙方个人应当缴纳的部分，由甲方代扣代缴。")
    article(doc, "第十六条　甲方自实际用工之日起，为乙方办理住房公积金缴存登记并按时足额缴存。缴费基数、缴存基数和缴存比例按照乙方工资及上海市规定执行。甲方不以现金或其他补贴代替上述法定缴费、缴存。")

    chapter(doc, "第六章　劳动保护和劳动条件")
    article(doc, "第十七条　甲方提供符合国家规定的劳动条件、劳动防护用品及必要的工作条件，对乙方进行安全生产和岗位培训，并告知工作中可能存在的职业危害及防护措施。乙方发生工伤或患职业病的，甲方依法办理认定手续并落实相应待遇。")
    article(doc, "第十八条　乙方遵守安全操作要求，正确使用防护用品，及时报告安全隐患。乙方有权拒绝违章指挥和强令冒险作业。")

    chapter(doc, "第七章　保密与知识产权")
    article(doc, "第十九条　乙方对工作中知悉的甲方商业秘密承担保密义务，不得擅自披露、使用或允许他人使用。已经合法公开的信息，乙方合法独立取得的信息，以及法律法规要求披露的信息，不属于本条保密范围。依法继续有效的保密义务，不因本合同解除或终止而消失。")
    article(doc, "第二十条　乙方履行职务形成的技术方案、图纸、工程资料、软件和发明创造，其知识产权依照法律规定确定。双方另有书面约定的，从其约定。乙方离职时交还工作资料、设备和物品，并办理工作交接。")
    article(doc, "第二十一条　本合同不约定竞业限制。确需约定的，双方另行签订书面协议，明确范围、地域、期限和按月支付的经济补偿，期限不得超过2年。除法定服务期和依法约定的竞业限制外，本合同不约定由乙方承担违约金。")

    chapter(doc, "第八章　规章制度")
    article(doc, "第二十二条　乙方遵守甲方依法制定并向其公示的规章制度。规章制度与法律或本合同不一致的，按照法律和本合同执行。")
    article(doc, "第二十三条　甲方不得扣押乙方的居民身份证、资格证书或其他证件原件，不得收取押金或要求乙方提供担保。乙方保证向甲方提供的身份、学历、职称等信息真实。")

    chapter(doc, "第九章　劳动合同的变更、解除和终止")
    article(doc, "第二十四条　变更本合同，由双方协商一致并采用书面形式。甲方变更名称、法定代表人、主要负责人或投资人的，不影响本合同履行。甲方合并或分立的，本合同由承继其权利义务的单位继续履行。")
    article(doc, "第二十五条　双方协商一致，可以解除本合同。乙方提前30日以书面形式通知甲方，可以解除本合同。甲方未及时足额支付劳动报酬、未依法缴纳社会保险，或有《中华人民共和国劳动合同法》第三十八条规定的其他情形的，乙方可以依法解除本合同。")
    article(doc, "第二十六条　甲方单方解除本合同，应当符合《中华人民共和国劳动合同法》第三十九条、第四十条、第四十一条的规定。依据第四十条解除的，应当提前30日书面通知乙方，或额外支付乙方1个月工资。乙方具有该法第四十二条规定情形的，甲方不得依照第四十条、第四十一条解除本合同。")
    article(doc, "第二十七条　本合同期满或出现其他法定终止情形的，本合同终止。依照法律规定应当续延的，延续至相应情形消失时终止。乙方符合订立无固定期限劳动合同的条件并提出订立的，甲方应当订立无固定期限劳动合同。")
    article(doc, "第二十八条　解除或终止本合同时，甲方出具解除或终止劳动合同的证明，并在15日内为乙方办理档案和社会保险关系转移手续，结清工资及其他应当支付的费用。乙方办理工作交接。")

    chapter(doc, "第十章　经济补偿")
    article(doc, "第二十九条　解除或终止本合同，符合《中华人民共和国劳动合同法》第四十六条规定的，甲方依法向乙方支付经济补偿。固定期限劳动合同期满终止的，除甲方维持或提高劳动合同约定条件续订、乙方不同意续订的以外，甲方依法支付经济补偿。")
    article(doc, "第三十条　经济补偿按乙方在甲方的工作年限计算，每满1年支付1个月工资；6个月以上不满1年的，按1年计算；不满6个月的，支付半个月工资。月工资按劳动合同解除或终止前12个月的平均工资计算；工作不满12个月的，按实际工作月数计算。月工资高于本市上年度职工月平均工资三倍的，按照法律规定的上限执行。经济补偿在乙方办理工作交接时支付。")
    article(doc, "第三十一条　甲方违法解除或终止本合同，乙方要求继续履行的，甲方应当继续履行；乙方不要求继续履行或本合同已经不能继续履行的，甲方依照《中华人民共和国劳动合同法》第八十七条支付赔偿金。")

    chapter(doc, "第十一章　争议处理和其他")
    article(doc, "第三十二条　因本合同发生争议的，双方可以协商解决，也可以向劳动合同履行地或甲方所在地有管辖权的劳动人事争议仲裁委员会申请仲裁。")
    article(doc, "第三十三条　双方通讯地址和联系电话以本合同载明的内容为准。任何一方变更的，应当自变更之日起7日内书面通知对方。")
    article(doc, "第三十四条　本合同未尽事宜，按照国家和上海市现行规定执行。本合同与强制性规定不一致的，按照强制性规定执行。双方另订的书面补充协议与本合同具有同等效力。双方无其他补充约定。")
    article(doc, "第三十五条　本合同一式两份，甲乙双方各执一份，具有同等法律效力。")

    signature_page(doc)
    return doc


def plain_text(doc: Document) -> str:
    parts: list[str] = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            seen: set[int] = set()
            for cell in row.cells:
                key = id(cell._tc)
                if key in seen:
                    continue
                seen.add(key)
                parts.append(cell.text)
    return "\n".join(parts)


def verify(doc: Document) -> None:
    text = plain_text(doc)
    required = [
        COMPANY,
        CREDIT,
        LEGAL_REP,
        REG_ADDR,
        OFFICE_ADDR,
        CONTACT,
        CONTACT_PHONE,
        EMPLOYEE,
        BIRTH,
        ID_NO,
        GENDER,
        HUKOU,
        HOME,
        PHONE,
        EMAIL,
        DEPT,
        POST,
        RANK,
        f"{WAGE}元",
        WAGE_DX,
        "每月15日",
        "三个月",
        "2026年5月1日",
        "2026年7月31日",
        "2029年4月30日",
        "不降低",
        "银行转账",
        CONTRACT_NO,
        "专业技术职务",
        "第三十五条",
    ]
    missing = [item for item in required if item not in text]
    banned = ["工资表", "勾稽", "划线注销", "请核对已填写", "gamil", "空白事项", "据实填写", "人工智能", "仅供参考", "高工资人员"]
    hit = [item for item in banned if item in text]
    if missing or hit:
        raise SystemExit(f"合同校验未通过。缺少：{missing}；不应出现：{hit}")
    if "＿" in text:
        raise SystemExit("合同中仍有下划线空栏")


def main() -> None:
    doc = build()
    verify(doc)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUT_DIR / DOCX_NAME
    doc.save(path)
    print(path)


if __name__ == "__main__":
    sys.exit(main())
