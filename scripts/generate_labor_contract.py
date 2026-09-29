#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成可直接签署的固定期限劳动合同（上海贝迪创建科技有限公司 / 胡继刚）。"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "output" / "劳动合同书_上海贝迪创建科技有限公司_胡继刚.docx"


def _keep_together(text: str) -> str:
    """在连字符两侧插入词连接符，避免门牌号在换行时被拆开。"""
    return text.replace("-", "\u2060-\u2060")


CONTRACT_NO = "BDCC-LD-2026-001"
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
HUKOU = _keep_together("山东省青岛市李沧区黑龙江中路629号2-2-2502")
RESIDENCE = _keep_together("上海市杨浦区爱国路389号1-1-401")
PHONE = "13262607888"
EMAIL = "hujigang2010@gmail.com"

DEPT = "工程管理部"
POSITION = "高级工程师"
RANK = "副总经理"
WAGE = "28,000"
WAGE_UPPER = "人民币贰万捌仟元整"
TERM_START = "2026年5月1日"
TERM_END = "2029年4月30日"
PROB_START = "2026年5月1日"
PROB_END = "2026年7月31日"
PAY_DAY = "15"

PAGE_WIDTH_CM = 16.0
BLACK = RGBColor(0, 0, 0)


def cm_to_dxa(value: float) -> int:
    return int(round(value * 1440 / 2.54))


def set_run_font(run, east_asia: str, size: float, bold: bool = False) -> None:
    run.bold = bold
    run.italic = False
    run.font.size = Pt(size)
    run.font.color.rgb = BLACK
    run.font.name = "Times New Roman"
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    r_fonts.set(qn("w:ascii"), "Times New Roman")
    r_fonts.set(qn("w:hAnsi"), "Times New Roman")
    r_fonts.set(qn("w:cs"), "Times New Roman")
    r_fonts.set(qn("w:eastAsia"), east_asia)
    lang = r_pr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        r_pr.append(lang)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:eastAsia"), "zh-CN")
    lang.set(qn("w:bidi"), "ar-SA")


def disable_snap(paragraph) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    snap = p_pr.find(qn("w:snapToGrid"))
    if snap is None:
        snap = OxmlElement("w:snapToGrid")
        p_pr.append(snap)
    snap.set(qn("w:val"), "0")


def set_paragraph_border(paragraph, edge: str = "bottom", sz: str = "12") -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    element = OxmlElement(f"w:{edge}")
    element.set(qn("w:val"), "single")
    element.set(qn("w:sz"), sz)
    element.set(qn("w:space"), "1")
    element.set(qn("w:color"), "000000")
    p_bdr.append(element)


def add_text_paragraph(
    document,
    text: str,
    *,
    font: str = "宋体",
    size: float = 12,
    bold: bool = False,
    align: str = "left",
    first_indent: float | None = 24,
    before: float = 0,
    after: float = 2,
    line: float = 1.3,
    keep_next: bool = False,
    page_break: bool = False,
):
    paragraph = document.add_paragraph()
    if page_break:
        paragraph.runs  # touch
        paragraph.paragraph_format.page_break_before = True
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = line
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    if first_indent is None:
        paragraph.paragraph_format.first_line_indent = Pt(0)
    else:
        paragraph.paragraph_format.first_line_indent = Pt(first_indent)
    paragraph.paragraph_format.keep_with_next = keep_next
    alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }[align]
    paragraph.alignment = alignment
    disable_snap(paragraph)
    if text:
        run = paragraph.add_run(text)
        set_run_font(run, font, size, bold)
    return paragraph


def add_mixed_paragraph(
    document,
    parts: list[tuple[str, str, float, bool]],
    *,
    align: str = "justify",
    first_indent: float | None = 24,
    before: float = 3,
    after: float = 2,
    line: float = 1.3,
    keep_next: bool = False,
):
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = line
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    paragraph.paragraph_format.first_line_indent = Pt(0 if first_indent is None else first_indent)
    paragraph.paragraph_format.keep_with_next = keep_next
    paragraph.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }[align]
    disable_snap(paragraph)
    for text, font, size, bold in parts:
        run = paragraph.add_run(text)
        set_run_font(run, font, size, bold)
    return paragraph


def shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    existing = tc_pr.find(qn("w:shd"))
    if existing is not None:
        tc_pr.remove(existing)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_margins(cell, top=40, bottom=40, left=90, right=90) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.find(qn("w:tcMar"))
    if tc_mar is not None:
        tc_pr.remove(tc_mar)
    tc_mar = OxmlElement("w:tcMar")
    for name, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        tc_mar.append(node)
    tc_pr.append(tc_mar)


def prevent_row_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def set_row_height(row, cm: float) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    existing = tr_pr.find(qn("w:trHeight"))
    if existing is not None:
        tr_pr.remove(existing)
    height = OxmlElement("w:trHeight")
    height.set(qn("w:val"), str(cm_to_dxa(cm)))
    height.set(qn("w:hRule"), "atLeast")
    tr_pr.append(height)


def set_table_widths(table, widths: list[float]) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.allow_autofit = False
    total = sum(widths)
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(cm_to_dxa(total)))
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
    if borders is not None:
        tbl_pr.remove(borders)
    borders = OxmlElement("w:tblBorders")
    for edge, sz in (
        ("top", "8"),
        ("left", "8"),
        ("bottom", "8"),
        ("right", "8"),
        ("insideH", "4"),
        ("insideV", "4"),
    ):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), sz)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "000000")
        borders.append(element)
    tbl_pr.append(borders)

    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        cols = grid.findall(qn("w:gridCol"))
        for index, col in enumerate(cols):
            col.set(qn("w:w"), str(cm_to_dxa(widths[index])))

    for row in table.rows:
        prevent_row_split(row)
        for index, cell in enumerate(row.cells):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(cm_to_dxa(widths[index])))
            tc_w.set(qn("w:type"), "dxa")


def write_cell(cell, text: str, *, font: str, size: float = 10.5, bold: bool = False, align: str = "left") -> None:
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    paragraph = cell.paragraphs[0]
    paragraph.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
    }[align]
    paragraph.paragraph_format.space_before = Pt(1)
    paragraph.paragraph_format.space_after = Pt(1)
    paragraph.paragraph_format.line_spacing = 1.08
    paragraph.paragraph_format.first_line_indent = Pt(0)
    disable_snap(paragraph)
    for run in list(paragraph.runs):
        run._element.getparent().remove(run._element)
    run = paragraph.add_run(text)
    set_run_font(run, font, size, bold)
    set_cell_margins(cell)


def add_kv_table(document, rows: list[tuple[str, str]], label_cm: float = 3.7) -> None:
    table = document.add_table(rows=len(rows), cols=2)
    set_table_widths(table, [label_cm, PAGE_WIDTH_CM - label_cm])
    for index, (label, value) in enumerate(rows):
        label_cell = table.rows[index].cells[0]
        value_cell = table.rows[index].cells[1]
        shade_cell(label_cell, "F3F3F3")
        write_cell(label_cell, label, font="黑体", bold=True)
        write_cell(value_cell, value, font="宋体")
        set_row_height(table.rows[index], 0.62)


def configure_document(document: Document) -> None:
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(1.7)
    section.header_distance = Cm(0.85)
    section.footer_distance = Cm(0.6)
    section.different_first_page_header_footer = True

    normal = document.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.font.color.rgb = BLACK
    r_pr = normal.element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    r_fonts.set(qn("w:ascii"), "Times New Roman")
    r_fonts.set(qn("w:hAnsi"), "Times New Roman")
    r_fonts.set(qn("w:eastAsia"), "宋体")
    r_fonts.set(qn("w:cs"), "Times New Roman")
    pf = normal.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.3

    settings = document.settings.element
    if settings.find(qn("w:updateFields")) is None:
        update = OxmlElement("w:updateFields")
        update.set(qn("w:val"), "true")
        settings.append(update)

    document.core_properties.title = "劳动合同书"
    document.core_properties.subject = f"{COMPANY}与{EMPLOYEE}固定期限劳动合同"
    document.core_properties.author = COMPANY
    document.core_properties.category = "劳动合同"
    document.core_properties.comments = ""


def add_page_field(paragraph, instruction: str, placeholder: str) -> None:
    run_begin = paragraph.add_run()
    set_run_font(run_begin, "宋体", 9)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run_begin._r.append(fld_begin)

    run_instr = paragraph.add_run()
    set_run_font(run_instr, "宋体", 9)
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruction} "
    run_instr._r.append(instr)

    run_sep = paragraph.add_run()
    set_run_font(run_sep, "宋体", 9)
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run_sep._r.append(fld_sep)

    run_text = paragraph.add_run(placeholder)
    set_run_font(run_text, "宋体", 9)

    run_end = paragraph.add_run()
    set_run_font(run_end, "宋体", 9)
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run_end._r.append(fld_end)


def fill_footer(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(2)
    paragraph.paragraph_format.space_after = Pt(0)
    disable_snap(paragraph)
    set_paragraph_border(paragraph, "top", "6")
    run_left = paragraph.add_run("第 ")
    set_run_font(run_left, "宋体", 9)
    add_page_field(paragraph, "PAGE", "1")
    run_mid = paragraph.add_run(" 页  共 ")
    set_run_font(run_mid, "宋体", 9)
    add_page_field(paragraph, "NUMPAGES", "6")
    run_right = paragraph.add_run(" 页")
    set_run_font(run_right, "宋体", 9)


def fill_header(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(2)
    disable_snap(paragraph)
    tab_stops = paragraph.paragraph_format.tab_stops
    tab_stops.add_tab_stop(Cm(PAGE_WIDTH_CM), WD_TAB_ALIGNMENT.RIGHT)
    set_paragraph_border(paragraph, "bottom", "8")
    left = paragraph.add_run("劳动合同书")
    set_run_font(left, "宋体", 9)
    tab = paragraph.add_run("\t")
    set_run_font(tab, "宋体", 9)
    right = paragraph.add_run(f"合同编号：{CONTRACT_NO}")
    set_run_font(right, "宋体", 9)


def add_article(document, number: str, paragraphs: list[str], *, together: bool = False) -> None:
    for index, text in enumerate(paragraphs):
        if index == 0:
            paragraph = add_mixed_paragraph(
                document,
                [(f"{number}　", "黑体", 12, True), (text, "宋体", 12, False)],
                before=6,
                after=1,
                keep_next=len(paragraphs) > 1 or together,
            )
        else:
            paragraph = add_text_paragraph(document, text, align="justify", before=1, after=1)
        if together:
            paragraph.paragraph_format.keep_together = True
            paragraph.paragraph_format.keep_with_next = index < len(paragraphs) - 1


def add_items(document, items: list[str]) -> None:
    for item in items:
        add_text_paragraph(document, item, align="justify", before=0, after=0, first_indent=24)


def build_articles(document: Document) -> None:
    add_text_paragraph(
        document,
        "一、劳动合同期限",
        font="黑体",
        size=14,
        bold=True,
        first_indent=None,
        before=0,
        after=2,
        line=1.15,
        keep_next=True,
        page_break=True,
    )
    add_article(
        document,
        "第一条",
        [
            f"本合同为固定期限劳动合同，自{TERM_START}起至{TERM_END}止。劳动关系自甲方实际用工之日起建立。"
        ],
    )
    add_article(
        document,
        "第二条",
        [
            (
                f"试用期为三个月，自{PROB_START}起至{PROB_END}止，包含在第一条的合同期限内。"
                f"试用期工资为人民币{WAGE}元（大写：{WAGE_UPPER}，税前），与本合同月工资相同，按全额发放，不降低、不打折。"
            ),
            "试用期届满甲方未解除本合同的，乙方试用合格。试用期内乙方提前三日通知甲方，可以解除本合同。甲方在试用期内解除本合同的，应当说明理由，并证明乙方不符合录用条件。甲方与乙方只能约定一次试用期。",
        ],
    )

    add_text_paragraph(
        document,
        "二、工作内容和工作地点",
        font="黑体",
        size=14,
        bold=True,
        first_indent=None,
        before=12,
        after=2,
        line=1.15,
        keep_next=True,
    )
    add_article(
        document,
        "第三条",
        [f"乙方的工作部门为{DEPT}，岗位为{POSITION}，职级为{RANK}。乙方的工作职责包括："],
    )
    add_items(
        document,
        [
            "（一）组织或参与工程技术交底和技术审查，协调解决实施过程中的技术问题；",
            "（二）参与施工组织设计和进度安排，协调设计、施工、监理等相关单位；",
            "（三）落实工程质量和安全技术要求，组织检查、整改和风险防范；",
            "（四）组织或参与工程技术方案的编制、论证、优化、实施检查和技术资料管理；",
            "（五）参与项目策划、建设实施、验收和移交，以及与上述岗位、职级相应的其他工作。",
        ],
    )
    add_text_paragraph(
        document,
        "甲方变更乙方的工作部门、岗位、职级或劳动报酬，应当与乙方协商一致并采用书面形式。甲方可以根据生产经营需要调整乙方具体负责的项目，不因此变更乙方的岗位和劳动报酬。",
        align="justify",
        before=1,
        after=1,
    )
    add_article(
        document,
        "第四条",
        [
            (
                f"乙方的工作地点为{OFFICE}。因项目需要在上海市范围内调整具体工作场所的，甲方应当提前告知乙方。"
                "甲方安排乙方到上海市以外地区出差的，承担差旅费用。变更工作城市的，由双方另行书面约定。"
            )
        ],
    )

    add_text_paragraph(
        document,
        "三、工作时间和休息休假",
        font="黑体",
        size=14,
        bold=True,
        first_indent=None,
        before=12,
        after=2,
        line=1.15,
        keep_next=True,
    )
    add_article(
        document,
        "第五条",
        [
            "乙方实行标准工时制度，每日工作时间不超过八小时，每周工作时间不超过四十小时，每周至少休息一日。具体作息时间由甲方安排并告知乙方。",
            "乙方依法享有休息日、法定节假日和带薪年休假，以及婚假、丧假、产假、陪产假、育儿假、病假等假期。",
            "甲方安排加班的，应当与乙方协商，并依法支付加班工资或安排补休。",
        ],
    )

    add_text_paragraph(
        document,
        "四、劳动报酬",
        font="黑体",
        size=14,
        bold=True,
        first_indent=None,
        before=12,
        after=2,
        line=1.15,
        keep_next=True,
    )
    add_article(
        document,
        "第六条",
        [
            (
                f"乙方月工资为人民币{WAGE}元（大写：{WAGE_UPPER}，税前）。"
                "该工资为乙方在法定工作时间内正常出勤的工资。试用期工资按第二条约定执行。"
            ),
            (
                f"甲方于每月{PAY_DAY}日以银行转账方式，将上一自然月工资支付至乙方本人银行账户。"
                f"{PAY_DAY}日为休息日或法定节假日的，提前至最近一个工作日支付。"
                "甲方支付工资时向乙方提供工资清单，列明应发项目、代扣项目和实发金额。"
            ),
            "乙方个人缴纳的社会保险费、住房公积金及个人所得税，由甲方代扣代缴。甲方不得克扣或者无故拖欠乙方工资，不得以罚款形式扣发工资。",
            (
                f"加班工资以本合同约定的月工资为基数，按月工资除以21.75计算日工资，按日工资除以8小时计算小时工资。"
                "工作日延长工作时间的，支付不低于工资150%的加班工资；休息日工作又不能安排补休的，支付不低于工资200%的加班工资；法定节假日工作的，支付不低于工资300%的加班工资。"
            ),
            "奖金、津贴和其他不固定项目，按照甲方公示并告知乙方的制度发放，不计入上述月工资，也不冲抵加班工资。乙方请假、缺勤以及依法休假、医疗期、停工停产期间的待遇，按照国家和上海市规定执行。调整月工资的，双方另行书面约定。",
        ],
    )

    add_text_paragraph(
        document,
        "五、社会保险和住房公积金",
        font="黑体",
        size=14,
        bold=True,
        first_indent=None,
        before=12,
        after=2,
        line=1.15,
        keep_next=True,
    )
    add_article(
        document,
        "第七条",
        [
            "甲方自实际用工之日起三十日内，为乙方办理社会保险登记和住房公积金缴存登记，并按时足额缴纳、缴存。试用期计入缴费和缴存期间。",
            "社会保险包括职工基本养老保险、职工基本医疗保险、失业保险、工伤保险，生育保险按照上海市规定执行。缴费基数、缴存基数和缴存比例按照乙方工资及上海市规定确定。甲方不以现金、补贴或其他形式代替上述缴纳、缴存。",
        ],
    )

    add_text_paragraph(
        document,
        "六、劳动保护、劳动条件和职业危害防护",
        font="黑体",
        size=14,
        bold=True,
        first_indent=None,
        before=12,
        after=2,
        line=1.15,
        keep_next=True,
    )
    add_article(
        document,
        "第八条",
        [
            "甲方为乙方提供符合国家规定的劳动安全卫生条件、必要的劳动防护用品和工作条件，对乙方进行安全生产教育，并告知工程现场及本岗位可能存在的职业危害、防护措施和应急方法。甲方依法安排职业健康检查。",
            "乙方应当遵守安全操作规程，正确使用劳动防护用品，发现事故隐患及时报告。乙方有权拒绝违章指挥和强令冒险作业，甲方不得因此处分乙方。",
            "乙方发生工伤或者患职业病的，甲方依法申请工伤认定、协助劳动能力鉴定，并落实相应待遇。",
        ],
    )

    add_text_paragraph(
        document,
        "七、保密和知识产权",
        font="黑体",
        size=14,
        bold=True,
        first_indent=None,
        before=12,
        after=2,
        line=1.15,
        keep_next=True,
    )
    add_article(
        document,
        "第九条",
        [
            "乙方对工作中知悉的甲方商业秘密承担保密义务，不得擅自披露、使用或者允许他人使用。已经合法公开的信息、乙方有证据证明系合法独立取得的信息，以及依照法律规定应当向有关机关提供的信息，不在保密范围内。上述保密义务不因本合同解除或终止而免除。",
            "乙方履行职务形成的技术方案、图纸、工程资料、软件和其他工作成果，其知识产权依照法律规定确定。双方另有书面约定的，按照书面约定执行。乙方离职时应当交还工作资料、设备和物品，并办理工作交接。",
            "本合同不约定竞业限制。除专项技术培训的服务期协议外，本合同不约定由乙方承担违约金。双方另行约定竞业限制的，应当签订书面协议，明确范围、地域和期限，由甲方在解除或终止劳动合同后按月支付经济补偿，期限不超过二年。",
        ],
    )

    add_text_paragraph(
        document,
        "八、规章制度",
        font="黑体",
        size=14,
        bold=True,
        first_indent=None,
        before=12,
        after=2,
        line=1.15,
        keep_next=True,
    )
    add_article(
        document,
        "第十条",
        [
            "甲方制定、修改直接涉及乙方切身利益的规章制度，应当依法听取意见，并向乙方公示或者告知。乙方应当遵守甲方依法制定并已公示的规章制度。规章制度与本合同不一致的，按照本合同执行。",
            "甲方不得扣押乙方的居民身份证、职称证书、资格证书或其他证件原件，不得要求乙方提供担保或者以其他名义向乙方收取财物。乙方应当如实提供与劳动合同直接相关的身份、学历、职称等情况。",
        ],
    )

    add_text_paragraph(
        document,
        "九、劳动合同的变更、解除和终止",
        font="黑体",
        size=14,
        bold=True,
        first_indent=None,
        before=12,
        after=2,
        line=1.15,
        keep_next=True,
    )
    add_article(
        document,
        "第十一条",
        [
            "变更本合同，应当经甲乙双方协商一致，并采用书面形式。甲方变更名称、法定代表人、主要负责人或者投资人的，不影响本合同履行。甲方合并或者分立的，本合同由承继其权利和义务的单位继续履行。",
            "双方协商一致，可以解除本合同。乙方提前三十日以书面形式通知甲方，可以解除本合同。",
            "甲方未按本合同约定支付劳动报酬、未依法为乙方缴纳社会保险费，或者有《中华人民共和国劳动合同法》第三十八条规定的其他情形的，乙方可以解除本合同。",
            "甲方单方解除本合同，应当符合《中华人民共和国劳动合同法》第三十九条、第四十条、第四十一条的规定。依照第四十条解除的，应当提前三十日以书面形式通知乙方，或者额外支付乙方一个月工资。乙方有该法第四十二条规定情形的，甲方不得依照第四十条、第四十一条解除本合同。",
            "本合同期满或者出现法律规定的终止情形时，本合同终止。依照法律规定应当续延的，续延至相应情形消失时终止。乙方符合订立无固定期限劳动合同的条件并要求订立的，甲方应当与乙方订立无固定期限劳动合同。",
            "解除或者终止本合同时，甲方应当出具解除或者终止劳动合同的证明，并在十五日内为乙方办理档案和社会保险关系转移手续，同时结清工资和其他应当支付的费用。乙方应当办理工作交接。",
        ],
    )

    add_text_paragraph(
        document,
        "十、经济补偿",
        font="黑体",
        size=14,
        bold=True,
        first_indent=None,
        before=12,
        after=2,
        line=1.15,
        keep_next=True,
    )
    add_article(
        document,
        "第十二条",
        [
            "解除或者终止本合同，符合《中华人民共和国劳动合同法》第四十六条规定的，甲方应当向乙方支付经济补偿。固定期限劳动合同期满终止，甲方维持或者提高劳动合同约定条件续订劳动合同，乙方不同意续订的，甲方不支付经济补偿；除此以外，甲方依法支付经济补偿。",
            "经济补偿按乙方在甲方工作的年限，每满一年支付一个月工资；六个月以上不满一年的，按一年计算；不满六个月的，支付半个月工资。月工资按劳动合同解除或者终止前十二个月的平均工资计算；工作不满十二个月的，按照实际月数计算。月工资高于本市上年度职工月平均工资三倍的，经济补偿按法律规定的上限支付。",
            "甲方应当在乙方办理工作交接时支付经济补偿。甲方违法解除或者终止本合同，乙方要求继续履行的，甲方应当继续履行；乙方不要求继续履行或者劳动合同已经不能继续履行的，甲方依照《中华人民共和国劳动合同法》第八十七条支付赔偿金。",
        ],
    )

    add_text_paragraph(
        document,
        "十一、争议处理",
        font="黑体",
        size=14,
        bold=True,
        first_indent=None,
        before=12,
        after=2,
        line=1.15,
        keep_next=True,
    )
    add_article(
        document,
        "第十三条",
        [
            "因本合同发生的争议，双方可以协商解决，也可以向劳动合同履行地或者甲方所在地有管辖权的劳动人事争议仲裁委员会申请仲裁。对仲裁裁决不服的，依法向人民法院提起诉讼。"
        ],
    )

    add_text_paragraph(
        document,
        "十二、其他",
        font="黑体",
        size=14,
        bold=True,
        first_indent=None,
        before=12,
        after=2,
        line=1.15,
        keep_next=True,
    )
    add_article(
        document,
        "第十四条",
        [
            (
                f"甲方送达地址：{OFFICE}。联系人：{CONTACT}。联系电话：{CONTACT_PHONE}。"
                f"乙方送达地址：{RESIDENCE}。联系电话：{PHONE}。电子邮箱：{EMAIL}。"
                "任何一方变更地址或联系方式的，应当在变更后七日内书面通知对方。书面通知可以当面交付、邮寄，或者发送至本合同载明的电子邮箱。"
            )
        ],
        together=True,
    )
    add_article(
        document,
        "第十五条",
        [
            "本合同未约定事项，按照国家和上海市有关规定执行。本合同与法律、法规强制性规定不一致的，按照强制性规定执行。双方就专项技术培训约定服务期的，另行签订书面协议。",
            "本合同一式两份，甲乙双方各执一份，具有同等法律效力。本合同自甲方盖章、法定代表人或者授权代表签字、乙方本人签字之日起生效。",
        ],
    )


def add_signature_page(document: Document) -> None:
    add_text_paragraph(
        document,
        "（以下无正文）",
        font="宋体",
        size=12,
        align="center",
        first_indent=None,
        before=12,
        after=6,
        line=1.15,
        page_break=False,
        keep_next=True,
    )
    add_text_paragraph(
        document,
        f"签署地点：{OFFICE}",
        font="宋体",
        size=12,
        align="left",
        first_indent=None,
        before=0,
        after=8,
        line=1.15,
        keep_next=True,
    )

    table = document.add_table(rows=1, cols=2)
    set_table_widths(table, [PAGE_WIDTH_CM / 2, PAGE_WIDTH_CM / 2])
    set_row_height(table.rows[0], 7.8)
    left, right = table.rows[0].cells
    for cell in (left, right):
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
        paragraph = cell.paragraphs[0]
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        disable_snap(paragraph)

    def fill_sign_cell(cell, lines: list[tuple[str, str, float, bool, float]]) -> None:
        first = cell.paragraphs[0]
        first.alignment = WD_ALIGN_PARAGRAPH.LEFT
        first.paragraph_format.first_line_indent = Pt(0)
        first.paragraph_format.space_before = Pt(2)
        first.paragraph_format.space_after = Pt(lines[0][4])
        first.paragraph_format.line_spacing = 1.15
        disable_snap(first)
        text, font, size, bold, _after = lines[0]
        run = first.add_run(text)
        set_run_font(run, font, size, bold)
        for text, font, size, bold, after in lines[1:]:
            paragraph = cell.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.first_line_indent = Pt(0)
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(after)
            paragraph.paragraph_format.line_spacing = 1.15
            disable_snap(paragraph)
            run = paragraph.add_run(text)
            set_run_font(run, font, size, bold)

    fill_sign_cell(
        left,
        [
            ("甲方（盖章）", "黑体", 12, True, 10),
            (COMPANY, "宋体", 12, False, 72),
            ("法定代表人或授权代表（签字）：", "宋体", 12, False, 22),
            ("日期：　　　　年　　　月　　　日", "宋体", 12, False, 2),
        ],
    )
    fill_sign_cell(
        right,
        [
            ("乙方（签字）", "黑体", 12, True, 10),
            (EMPLOYEE, "宋体", 12, False, 72),
            ("本人签字：", "宋体", 12, False, 22),
            ("日期：　　　　年　　　月　　　日", "宋体", 12, False, 2),
        ],
    )

    add_text_paragraph(
        document,
        "合同文本领取确认",
        font="黑体",
        size=12,
        bold=True,
        first_indent=None,
        before=16,
        after=4,
        line=1.15,
        keep_next=True,
    )
    add_text_paragraph(
        document,
        "乙方确认：已领取本合同文本一份。",
        first_indent=None,
        before=0,
        after=10,
        line=1.15,
        keep_next=True,
    )
    add_text_paragraph(
        document,
        "乙方签字：　　　　　　　　　　　　日期：　　　　年　　　月　　　日",
        first_indent=None,
        before=6,
        after=0,
        line=1.15,
    )


def document_text(document: Document) -> str:
    chunks: list[str] = []
    for paragraph in document.paragraphs:
        if paragraph.text.strip():
            chunks.append(paragraph.text.strip())
    for table in document.tables:
        for row in table.rows:
            seen = []
            for cell in row.cells:
                value = cell.text.strip()
                if value and value not in seen:
                    seen.append(value)
            if seen:
                chunks.append(" ".join(seen))
    return "\n".join(chunks)


def validate(document: Document) -> None:
    text = document_text(document)
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
        PHONE,
        EMAIL,
        DEPT,
        POSITION,
        RANK,
        WAGE,
        WAGE_UPPER,
        TERM_START,
        TERM_END,
        PROB_END,
        "三个月",
        "不降低、不打折",
        f"每月{PAY_DAY}日",
        "银行转账",
        "标准工时",
        "（以下无正文）",
        "已领取本合同文本一份",
    ]
    missing = [item for item in required if item not in text]
    forbidden = [
        "建材",
        "gamil",
        "工资表",
        "勾稽",
        "下划线",
        "空白处",
        "据实填写",
        "总裁办",
        "已经高于",
        "2026年8月",
        "2026年10月",
        "22,400",
        "作为该月",
        "未填写",
        "本合同期限为3年，上述试用期不超过",
    ]
    hit = [item for item in forbidden if item in text]
    if missing or hit:
        raise SystemExit(f"合同校验未通过。缺少：{missing}；不应出现：{hit}")
    if "六个月" in text and "不满六个月" not in text:
        raise SystemExit("试用期被写成六个月")
    if text.count("不降低、不打折") < 2:
        raise SystemExit("试用期工资全额发放未在首页和正文同时写明")


def generate() -> Path:
    document = Document()
    configure_document(document)
    section = document.sections[0]
    fill_header(section.header.paragraphs[0])
    fill_footer(section.footer.paragraphs[0])
    fill_footer(section.first_page_footer.paragraphs[0])
    section.first_page_header.paragraphs[0].text = ""

    add_text_paragraph(
        document,
        COMPANY,
        font="黑体",
        size=12,
        bold=True,
        align="center",
        first_indent=None,
        before=0,
        after=2,
        line=1.0,
    )
    add_text_paragraph(
        document,
        "劳　动　合　同　书",
        font="黑体",
        size=26,
        bold=True,
        align="center",
        first_indent=None,
        before=2,
        after=0,
        line=1.0,
    )
    add_text_paragraph(
        document,
        "（固定期限）",
        font="宋体",
        size=12,
        align="center",
        first_indent=None,
        before=2,
        after=2,
        line=1.0,
    )
    number = add_text_paragraph(
        document,
        f"合同编号：{CONTRACT_NO}",
        font="宋体",
        size=10.5,
        align="center",
        first_indent=None,
        before=0,
        after=3,
        line=1.0,
    )
    set_paragraph_border(number, "bottom", "12")

    add_text_paragraph(
        document,
        "甲方（用人单位）",
        font="黑体",
        size=12,
        bold=True,
        first_indent=None,
        before=6,
        after=2,
        line=1.0,
        keep_next=True,
    )
    add_kv_table(
        document,
        [
            ("名　　称", COMPANY),
            ("统一社会信用代码", USCC),
            ("法定代表人", LEGAL_REP),
            ("注册地址", REG_ADDRESS),
            ("实际办公地址", OFFICE),
            ("联系人及电话", f"{CONTACT}　　{CONTACT_PHONE}"),
        ],
    )
    add_text_paragraph(
        document,
        "乙方（劳动者）",
        font="黑体",
        size=12,
        bold=True,
        first_indent=None,
        before=6,
        after=2,
        line=1.0,
        keep_next=True,
    )
    add_kv_table(
        document,
        [
            ("姓　　名", f"{EMPLOYEE}　　性别：{GENDER}　　出生日期：{BIRTH}"),
            ("居民身份证号码", ID_NO),
            ("户籍地址", HUKOU),
            ("现居住地址", RESIDENCE),
            ("联系电话", PHONE),
            ("电子邮箱", EMAIL),
        ],
    )
    add_text_paragraph(
        document,
        "甲乙双方根据《中华人民共和国劳动法》《中华人民共和国劳动合同法》及有关规定，经平等自愿、协商一致，订立本合同。双方约定的主要内容如下：",
        align="justify",
        first_indent=24,
        before=2,
        after=2,
        keep_next=True,
    )
    add_kv_table(
        document,
        [
            ("合同类型", "固定期限劳动合同"),
            ("合同期限", f"{TERM_START}起至{TERM_END}止"),
            ("试用期", f"三个月，{PROB_START}起至{PROB_END}止"),
            ("工作部门", DEPT),
            ("岗　　位", POSITION),
            ("职　　级", RANK),
            ("工作地点", OFFICE),
            ("月工资（税前）", f"人民币{WAGE}元（大写：{WAGE_UPPER}）"),
            ("试用期工资", f"人民币{WAGE}元，与月工资相同，按全额发放，不降低、不打折"),
            ("发　薪　日", f"每月{PAY_DAY}日，以银行转账支付上一自然月工资"),
        ],
    )

    build_articles(document)
    add_signature_page(document)
    validate(document)
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    document.save(OUTPUT)
    return OUTPUT


def main() -> int:
    path = generate()
    print(path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
