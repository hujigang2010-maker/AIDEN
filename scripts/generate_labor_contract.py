#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成可直接签署的《劳动合同书》。

用人单位按登记名称书写。正文只写双方确认的约定，不写填表说明。
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

SONG = "宋体"
HEI = "黑体"
INK = "222222"
MUTED = "555555"
LINE = "222222"
LABEL_FILL = "F3F3F3"
HEAD_FILL = "E8E8E8"

COMPANY = "上海贝迪创建科技有限公司"
USCC = "91310113MAERPR7M1D"
LEGAL_REP = "刘李杨"
REG_ADDRESS = "上海市宝山区蕰川路5475号4幢部分"
OFFICE = "上海市浦东新区康桥东路111号13号楼一层"
CONTACT = "段鹏涛"
CONTACT_PHONE = "13916901963"

EMPLOYEE = "胡继刚"
GENDER = "男"
BIRTH = "1987年4月25日"
ID_NO = "220283198704250614"
HUKOU = "青岛市李沧区黑龙江中路629号2-2-2502"
RESIDENCE = "上海市杨浦区爱国路389号1-1-401"
MOBILE = "13262607888"
EMAIL = "hujigang2010@gmail.com"

DEPARTMENT = "工程管理部"
POSITION = "高级工程师"
RANK = "副总经理"
WAGE = 28000
WAGE_TEXT = "28,000"
WAGE_DX = "人民币贰万捌仟元整"
TERM_START = "2026年5月1日"
TERM_END = "2029年4月30日"
PROBATION_END = "2026年7月31日"
PAYDAY = "15"

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "output"
DOCX_NAME = f"劳动合同书_{COMPANY}_{EMPLOYEE}.docx"


def _id_ok(number: str) -> bool:
    weights = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
    checks = "10X98765432"
    total = sum(int(number[i]) * weights[i] for i in range(17))
    return len(number) == 18 and checks[total % 11] == number[17]


if not _id_ok(ID_NO):
    raise SystemExit("身份证号码校验未通过")
if ID_NO[6:14] != "19870425" or int(ID_NO[16]) % 2 != 1:
    raise SystemExit("身份证号码与出生日期或性别不一致")


def set_run_font(run, name: str, size: float, bold: bool = False, color: str = INK) -> None:
    run.bold = bold
    run.italic = False
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    run.font.name = "Times New Roman"
    rpr = run._element.get_or_add_rPr()
    r_fonts = rpr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        rpr.append(r_fonts)
    r_fonts.set(qn("w:ascii"), "Times New Roman")
    r_fonts.set(qn("w:hAnsi"), "Times New Roman")
    r_fonts.set(qn("w:eastAsia"), name)
    r_fonts.set(qn("w:cs"), name)


def set_paragraph_format(
    paragraph,
    *,
    before: float = 0,
    after: float = 0,
    line: float = 1.15,
    indent_cm: float = 0,
    align: str = "left",
    keep_next: bool = False,
) -> None:
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.line_spacing = line
    fmt.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    fmt.first_line_indent = Cm(indent_cm)
    fmt.widow_control = True
    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    paragraph.alignment = align_map[align]
    ppr = paragraph._p.get_or_add_pPr()
    if keep_next and ppr.find(qn("w:keepNext")) is None:
        ppr.append(OxmlElement("w:keepNext"))
    if ppr.find(qn("w:keepLines")) is None:
        ppr.append(OxmlElement("w:keepLines"))


def add_bottom_border(paragraph, sz: str = "12", color: str = "1A1A1A") -> None:
    ppr = paragraph._p.get_or_add_pPr()
    borders = ppr.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        ppr.append(borders)
    edge = OxmlElement("w:bottom")
    edge.set(qn("w:val"), "single")
    edge.set(qn("w:sz"), sz)
    edge.set(qn("w:space"), "1")
    edge.set(qn("w:color"), color)
    borders.append(edge)


def add_top_border(paragraph, sz: str = "6", color: str = "1A1A1A") -> None:
    ppr = paragraph._p.get_or_add_pPr()
    borders = ppr.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        ppr.append(borders)
    edge = OxmlElement("w:top")
    edge.set(qn("w:val"), "single")
    edge.set(qn("w:sz"), sz)
    edge.set(qn("w:space"), "1")
    edge.set(qn("w:color"), color)
    borders.append(edge)


def shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top: int = 50, bottom: int = 50, left: int = 80, right: int = 80) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.find(qn("w:tcMar"))
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for edge, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
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
        for index, cell in enumerate(row.cells):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(widths[index]))
            tc_w.set(qn("w:type"), "dxa")


def set_table_borders(table, color: str = LINE, sz: str = "6") -> None:
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


def set_row_cant_split(row, height_cm: float | None = None) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    if height_cm is not None:
        tr_height = tr_pr.find(qn("w:trHeight"))
        if tr_height is None:
            tr_height = OxmlElement("w:trHeight")
            tr_pr.append(tr_height)
        tr_height.set(qn("w:val"), str(int(height_cm * 567)))
        tr_height.set(qn("w:hRule"), "atLeast")
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def write_cell(
    cell,
    text: str,
    *,
    bold: bool = False,
    size: float = 10.5,
    align: str = "left",
    fill: str | None = None,
    font: str = SONG,
    color: str = INK,
) -> None:
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_margins(cell)
    if fill:
        shade_cell(cell, fill)
    paragraph = cell.paragraphs[0]
    set_paragraph_format(paragraph, before=0, after=0, line=1.05, align=align)
    for run in list(paragraph.runs):
        run._element.getparent().remove(run._element)
    run = paragraph.add_run(text)
    set_run_font(run, font, size, bold=bold, color=color)


def add_table(doc, rows: int, widths: list[float]):
    table = doc.add_table(rows=rows, cols=len(widths))
    set_table_fixed(table, widths)
    set_table_borders(table)
    for row in table.rows:
        set_row_cant_split(row, 0.72)
    return table


def add_field(paragraph, instruction: str) -> None:
    def fld(char_type: str | None = None, text: str | None = None, instr: str | None = None):
        run = paragraph.add_run()
        set_run_font(run, SONG, 9, color=MUTED)
        if char_type:
            node = OxmlElement("w:fldChar")
            node.set(qn("w:fldCharType"), char_type)
            run._r.append(node)
        if instr is not None:
            node = OxmlElement("w:instrText")
            node.set(qn("xml:space"), "preserve")
            node.text = f" {instr} "
            run._r.append(node)
        if text is not None:
            run.text = text

    fld(char_type="begin")
    fld(instr=instruction)
    fld(char_type="separate")
    fld(text="1")
    fld(char_type="end")


def configure_header_footer(section) -> None:
    section.different_first_page_header_footer = True

    first_header = section.first_page_header
    first_header.is_linked_to_previous = False
    hp0 = first_header.paragraphs[0]
    set_paragraph_format(hp0, before=0, after=0, line=1.0)
    hp0.add_run("")

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    set_paragraph_format(hp, before=0, after=2, line=1.0, align="left")
    left = hp.add_run(COMPANY)
    set_run_font(left, SONG, 9, color=MUTED)
    right = hp.add_run("　劳动合同书")
    set_run_font(right, SONG, 9, color=MUTED)
    add_bottom_border(hp, sz="8")

    def fill_footer(paragraph) -> None:
        set_paragraph_format(paragraph, before=2, after=0, line=1.0, align="center")
        add_top_border(paragraph, sz="6")
        lead = paragraph.add_run("第 ")
        set_run_font(lead, SONG, 9, color=MUTED)
        add_field(paragraph, "PAGE")
        mid = paragraph.add_run(" 页")
        set_run_font(mid, SONG, 9, color=MUTED)

    first_footer = section.first_page_footer
    first_footer.is_linked_to_previous = False
    fill_footer(first_footer.paragraphs[0])
    footer = section.footer
    footer.is_linked_to_previous = False
    fill_footer(footer.paragraphs[0])


def configure_styles(doc: Document) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.font.color.rgb = RGBColor.from_string(INK)
    rpr = normal.element.get_or_add_rPr()
    r_fonts = rpr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        rpr.append(r_fonts)
    r_fonts.set(qn("w:ascii"), "Times New Roman")
    r_fonts.set(qn("w:hAnsi"), "Times New Roman")
    r_fonts.set(qn("w:eastAsia"), SONG)
    r_fonts.set(qn("w:cs"), SONG)
    pf = normal.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.15

    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        update = OxmlElement("w:updateFields")
        update.set(qn("w:val"), "true")
        settings.append(update)
    theme_lang = settings.find(qn("w:themeFontLang"))
    if theme_lang is None:
        theme_lang = OxmlElement("w:themeFontLang")
        settings.append(theme_lang)
    theme_lang.set(qn("w:val"), "en-US")
    theme_lang.set(qn("w:eastAsia"), "zh-CN")


def add_text(
    doc,
    text: str,
    *,
    size: float = 12,
    font: str = SONG,
    bold: bool = False,
    before: float = 0,
    after: float = 0,
    line: float = 1.15,
    indent_cm: float = 0,
    align: str = "left",
    color: str = INK,
    keep_next: bool = False,
    page_break: bool = False,
):
    paragraph = doc.add_paragraph()
    set_paragraph_format(
        paragraph,
        before=before,
        after=after,
        line=line,
        indent_cm=indent_cm,
        align=align,
        keep_next=keep_next,
    )
    if page_break:
        paragraph.paragraph_format.page_break_before = True
    run = paragraph.add_run(text)
    set_run_font(run, font, size, bold=bold, color=color)
    return paragraph


def add_article(doc, number: str, paragraphs: list[str]) -> None:
    for index, text in enumerate(paragraphs):
        paragraph = doc.add_paragraph()
        set_paragraph_format(
            paragraph,
            before=8 if index == 0 else 1,
            after=1,
            line=1.2,
            indent_cm=0.74,
            align="justify",
        )
        if index == 0:
            lead = paragraph.add_run(f"{number}　")
            set_run_font(lead, SONG, 12, bold=True)
        body = paragraph.add_run(text)
        set_run_font(body, SONG, 12)


def add_party_table(doc, title: str, rows: list[tuple[str, str]]) -> None:
    table = add_table(doc, 1 + len(rows), [3.7, 12.9])
    table.cell(0, 0).merge(table.cell(0, 1))
    write_cell(table.cell(0, 0), title, bold=True, size=11, font=HEI, fill=HEAD_FILL, align="left")
    set_row_cant_split(table.rows[0], 0.78)
    for index, (label, value) in enumerate(rows, start=1):
        write_cell(table.cell(index, 0), label, bold=True, fill=LABEL_FILL, align="center")
        write_cell(table.cell(index, 1), value, bold=(index == 1))
        set_row_cant_split(table.rows[index], 0.7)


def build() -> Document:
    doc = Document()
    configure_styles(doc)
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(1.7)
    section.header_distance = Cm(0.6)
    section.footer_distance = Cm(0.45)
    configure_header_footer(section)

    sect_pr = section._sectPr
    grid = sect_pr.find(qn("w:docGrid"))
    if grid is not None:
        sect_pr.remove(grid)

    if doc.paragraphs and not doc.paragraphs[0].text:
        doc.element.body.remove(doc.paragraphs[0]._element)

    core = doc.core_properties
    core.title = "劳动合同书"
    core.subject = f"{COMPANY}与{EMPLOYEE}"
    core.category = "劳动合同"
    core.last_modified_by = COMPANY
    core.comments = ""
    for child in core._element:
        if child.tag.endswith("creator"):
            child.text = COMPANY

    title = add_text(
        doc,
        "劳 动 合 同 书",
        size=22,
        font=HEI,
        bold=True,
        align="center",
        before=0,
        after=2,
        line=1.0,
    )
    subtitle = add_text(
        doc,
        "（固定期限）",
        size=12,
        font=HEI,
        align="center",
        before=0,
        after=6,
        line=1.0,
        color=MUTED,
    )
    add_bottom_border(subtitle, sz="16", color="1A1A1A")

    add_party_table(
        doc,
        "甲方（用人单位）",
        [
            ("名称", COMPANY),
            ("统一社会信用代码", USCC),
            ("法定代表人", LEGAL_REP),
            ("住所", REG_ADDRESS),
            ("实际办公地址", OFFICE),
            ("联系人", f"{CONTACT}　　联系电话：{CONTACT_PHONE}"),
        ],
    )
    add_text(doc, "", size=6, after=2, line=1.0)
    add_party_table(
        doc,
        "乙方（劳动者）",
        [
            ("姓名", f"{EMPLOYEE}　　性别：{GENDER}　　出生日期：{BIRTH}"),
            ("居民身份证号码", ID_NO),
            ("户籍地址", HUKOU),
            ("现居住地址", RESIDENCE),
            ("联系电话", MOBILE),
            ("电子邮箱", EMAIL),
        ],
    )

    add_text(
        doc,
        "甲乙双方根据《中华人民共和国劳动法》《中华人民共和国劳动合同法》及有关规定，经平等自愿、协商一致，订立本合同。",
        before=10,
        after=6,
        indent_cm=0.74,
        align="justify",
    )

    summary = add_table(doc, 10, [3.7, 12.9])
    summary.cell(0, 0).merge(summary.cell(0, 1))
    write_cell(summary.cell(0, 0), "合同主要事项", bold=True, size=11, font=HEI, fill=HEAD_FILL)
    items = [
        ("合同期限", f"{TERM_START}起至{TERM_END}止"),
        ("试用期", f"三个月，{TERM_START}起至{PROBATION_END}止"),
        ("工作部门", DEPARTMENT),
        ("岗位", POSITION),
        ("职级", RANK),
        ("工作地点", OFFICE),
        ("月工资（税前）", f"{WAGE_TEXT}元（大写：{WAGE_DX}）"),
        ("试用期工资", f"{WAGE_TEXT}元（大写：{WAGE_DX}），与月工资相同，不予降低"),
        ("发薪日", f"每月{PAYDAY}日，银行转账"),
    ]
    for index, (label, value) in enumerate(items, start=1):
        write_cell(summary.cell(index, 0), label, bold=True, fill=LABEL_FILL, align="center")
        write_cell(summary.cell(index, 1), value)
        set_row_cant_split(summary.rows[index], 0.66)
    set_row_cant_split(summary.rows[0], 0.74)

    add_article(
        doc,
        "第一条",
        [f"本合同为固定期限劳动合同，期限自{TERM_START}起至{TERM_END}止。劳动关系自实际用工之日起建立。"],
    )
    add_article(
        doc,
        "第二条",
        [
            f"试用期为三个月，自{TERM_START}起至{PROBATION_END}止，包含在第一条的劳动合同期限内。",
            "试用期月工资与第五条约定的月工资相同，按全额计发，不予降低。",
            "试用期内，乙方提前三日通知甲方，可以解除本合同。甲方在试用期内解除本合同的，应当说明理由，并证明乙方不符合录用条件。同一用人单位与同一劳动者只能约定一次试用期。",
        ],
    )
    add_article(
        doc,
        "第三条",
        [
            f"乙方的工作部门为{DEPARTMENT}，岗位为{POSITION}，职级为{RANK}。",
            "乙方从事工程技术及项目管理工作，包括工程技术管理、施工组织协调、质量与安全技术管理、工程技术方案管理、项目过程管理，以及与上述岗位和职级相应的其他工作。",
            "甲方调整乙方的工作部门、岗位、职级或者劳动报酬的，应当与乙方协商一致，并采用书面形式。",
        ],
    )
    add_article(
        doc,
        "第四条",
        [
            f"乙方的工作地点为{OFFICE}。",
            "因项目需要在上海市范围内变更具体工作场所的，甲方提前告知乙方。甲方安排乙方到上海市以外地区临时出差的，承担差旅费用。长期变更工作城市的，双方另行书面约定。",
        ],
    )
    add_article(
        doc,
        "第五条",
        [
            "乙方实行标准工时制度，每日工作时间不超过八小时，每周工作时间不超过四十小时。具体作息时间由甲方确定并告知乙方。",
            "乙方依法享有休息日、法定节假日、带薪年休假及其他法定假期。甲方安排加班的，依法支付加班工资或者安排补休。",
        ],
    )
    add_article(
        doc,
        "第六条",
        [
            f"乙方月工资为人民币{WAGE_TEXT}元（大写：{WAGE_DX}），税前。该工资为乙方在法定工作时间内提供正常劳动的工资。",
            f"试用期月工资为人民币{WAGE_TEXT}元（大写：{WAGE_DX}），税前。该工资与前款月工资相同，按全额计发，不予降低。",
            f"甲方于每月{PAYDAY}日以银行转账方式，向乙方支付上一个自然月的工资。{PAYDAY}日为休息日或法定节假日的，提前至最近一个工作日支付。",
            "甲方支付工资时向乙方提供工资清单。乙方个人缴纳的社会保险费、住房公积金，以及依法应扣缴的个人所得税，由甲方代扣代缴。",
            "加班工资以本合同约定的月工资为基数，按法定标准计发。日工资按月工资除以21.75计算，小时工资按日工资除以8小时计算。",
        ],
    )
    add_article(
        doc,
        "第七条",
        [
            "甲方自实际用工之日起三十日内，为乙方办理社会保险登记和住房公积金缴存登记，并按时足额缴纳、缴存。试用期计入缴费、缴存期间。",
            "社会保险费和住房公积金的缴费基数、缴存基数，按乙方工资及上海市规定确定。甲方不得以现金或其他补贴代替法定缴纳、缴存。",
        ],
    )
    add_article(
        doc,
        "第八条",
        [
            "甲方为乙方提供符合国家规定的劳动安全卫生条件和必要的劳动防护用品，对乙方进行安全生产教育，并告知本岗位可能存在的职业危害及防护措施。",
            "乙方应当遵守安全操作要求。乙方有权拒绝违章指挥和强令冒险作业。乙方发生工伤的，甲方依法申请工伤认定并落实相应待遇。",
        ],
    )
    add_article(
        doc,
        "第九条",
        [
            "乙方对在工作中知悉的甲方商业秘密承担保密义务，不得擅自披露、使用或者允许他人使用。该义务不因本合同解除或终止而免除。已经合法公开的信息、乙方有证据证明合法独立取得的信息，以及依照法律应当向有关机关披露的信息，不在此限。",
            "乙方履行职务形成的工作成果，其知识产权依照法律规定确定。双方另有书面约定的，从其约定。",
            "本合同不约定竞业限制，也不约定由乙方承担违约金。双方另行书面约定竞业限制的，应当明确范围和期限，由甲方在解除或终止劳动合同后按月支付经济补偿，期限不超过二年。",
        ],
    )
    add_article(
        doc,
        "第十条",
        [
            "乙方应当遵守甲方依法制定并向乙方公示的规章制度。规章制度与本合同不一致的，按本合同执行。甲方不得以规章制度免除自身法定责任、加重乙方责任或者排除乙方主要权利。",
            "甲方不得扣押乙方的居民身份证、职称证书、资格证书或其他证件原件，不得要求乙方提供担保或者以其他名义向乙方收取财物。",
        ],
    )
    add_article(
        doc,
        "第十一条",
        [
            "变更本合同，应当经双方协商一致，并采用书面形式。甲方变更名称、法定代表人、主要负责人或者投资人，不影响本合同履行。",
            "乙方提前三十日以书面形式通知甲方，可以解除本合同。甲方未按本合同约定支付劳动报酬、未依法为乙方缴纳社会保险费，或者有法律规定的其他情形的，乙方可以依法解除本合同。",
            "甲方解除本合同，应当符合法律规定的条件。乙方在规定的医疗期内，或者有法律规定不得解除劳动合同的情形的，甲方不得违法解除。",
            "解除或终止本合同时，甲方应当出具解除或终止劳动合同的证明，在十五日内为乙方办理档案和社会保险关系转移手续，并结清工资及其他应当支付的费用。乙方应当办理工作交接。符合法律规定应当支付经济补偿或赔偿金的，甲方依法支付。",
        ],
    )
    add_article(
        doc,
        "第十二条",
        [
            "因本合同发生的争议，双方可以协商解决，也可以依法向劳动人事争议仲裁委员会申请仲裁。协商不是申请仲裁的必经程序。",
        ],
    )
    add_article(
        doc,
        "第十三条",
        [
            f"甲方的送达地址为{OFFICE}，联系人{CONTACT}，联系电话{CONTACT_PHONE}。乙方的送达地址为{RESIDENCE}，联系电话{MOBILE}，电子邮箱{EMAIL}。任何一方变更送达地址或联系方式的，应当在变更后七日内书面通知对方。",
            "本合同未尽事宜，按国家和上海市有关规定执行。本合同的约定与法律、法规的强制性规定不一致的，按强制性规定执行。",
            "本合同一式两份，甲乙双方各执一份，具有同等效力。本合同自甲方盖章、法定代表人或授权代表签字，且乙方本人签字之日起生效。",
        ],
    )

    add_text(
        doc,
        "签署页",
        size=16,
        font=HEI,
        bold=True,
        align="center",
        before=14,
        after=2,
        line=1.0,
    )
    add_text(
        doc,
        "（以下无正文）",
        size=10.5,
        align="center",
        before=0,
        after=8,
        line=1.0,
        color=MUTED,
    )
    add_text(
        doc,
        f"签署地点：{OFFICE}",
        size=12,
        before=0,
        after=8,
        align="left",
    )

    sign = add_table(doc, 5, [8.3, 8.3])
    write_cell(sign.cell(0, 0), "甲方（盖章）", bold=True, font=HEI, size=11, fill=HEAD_FILL, align="center")
    write_cell(sign.cell(0, 1), "乙方（签字）", bold=True, font=HEI, size=11, fill=HEAD_FILL, align="center")
    write_cell(sign.cell(1, 0), COMPANY, bold=True, align="center")
    write_cell(sign.cell(1, 1), EMPLOYEE, bold=True, align="center")
    write_cell(sign.cell(2, 0), "", align="center")
    write_cell(sign.cell(2, 1), "", align="center")
    write_cell(sign.cell(3, 0), "法定代表人或授权代表签字：", align="left")
    write_cell(sign.cell(3, 1), "", align="left")
    write_cell(sign.cell(4, 0), "日期：　　　　年　　月　　日", align="left")
    write_cell(sign.cell(4, 1), "日期：　　　　年　　月　　日", align="left")
    for index, height in ((0, 0.8), (1, 0.85), (2, 3.8), (3, 1.2), (4, 0.85)):
        set_row_cant_split(sign.rows[index], height)

    add_text(
        doc,
        "乙方确认：已收到本合同文本一份。",
        before=14,
        after=8,
        indent_cm=0,
        align="left",
    )
    receipt = add_table(doc, 1, [8.3, 8.3])
    write_cell(receipt.cell(0, 0), "乙方签字：", align="left")
    write_cell(receipt.cell(0, 1), "日期：　　　　年　　月　　日", align="left")
    set_row_cant_split(receipt.rows[0], 1.35)

    return doc


def document_text(doc: Document) -> str:
    parts = [paragraph.text for paragraph in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            parts.extend(cell.text for cell in row.cells)
    return "\n".join(parts)


def assert_document(doc: Document) -> None:
    text = document_text(doc)
    required = [
        COMPANY,
        USCC,
        LEGAL_REP,
        REG_ADDRESS,
        OFFICE,
        CONTACT,
        CONTACT_PHONE,
        EMPLOYEE,
        GENDER,
        BIRTH,
        ID_NO,
        HUKOU,
        RESIDENCE,
        MOBILE,
        EMAIL,
        DEPARTMENT,
        POSITION,
        RANK,
        WAGE_TEXT,
        WAGE_DX,
        TERM_START,
        TERM_END,
        PROBATION_END,
        "三个月",
        "每月15日",
        "银行转账",
        "不予降低",
        "全额计发",
        "标准工时",
        "社会保险",
        "住房公积金",
        "商业秘密",
        "竞业限制",
        "经济补偿",
        "劳动人事争议仲裁",
        "签署页",
        "已收到本合同文本一份",
    ]
    missing = [item for item in required if item not in text]
    if missing:
        raise SystemExit("合同缺少内容：" + "、".join(missing))
    forbidden = [
        "创建材",
        "gamil",
        "勾稽",
        "工资表",
        "总裁办",
        "下划线",
        "空白",
        "待填",
        "百分之八十",
        "22,400",
        "本附件",
        "作为参考",
        "AI",
    ]
    hit = [item for item in forbidden if item in text]
    if hit:
        raise SystemExit("合同含有不应当出现的文字：" + "、".join(hit))
    for number in (
        "第一条",
        "第二条",
        "第三条",
        "第四条",
        "第五条",
        "第六条",
        "第七条",
        "第八条",
        "第九条",
        "第十条",
        "第十一条",
        "第十二条",
        "第十三条",
    ):
        if number not in text:
            raise SystemExit(f"缺少{number}")
    if "第十四条" in text:
        raise SystemExit("条款超出约定范围")


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc = build()
    assert_document(doc)
    out = OUT_DIR / DOCX_NAME
    doc.save(out)
    print(f"已生成：{out}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"生成失败：{exc}", file=sys.stderr)
        raise
