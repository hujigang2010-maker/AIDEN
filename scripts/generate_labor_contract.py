#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成《劳动合同书》：上海贝迪创建科技有限公司 — 胡继刚。

工资、社保和公积金数字只来自 2026 年 8 月工资表，并在脚本内复核勾稽。
法定代表人、注册地址、统一社会信用代码、身份证号、发薪日等未知事项留空，不编造。
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor, Twips

SERIF = "Noto Serif CJK SC"
SANS = "Noto Sans CJK SC"
INK = "222222"
MUTED = "444444"
LINE = "666666"
LABEL_FILL = "F3F4F6"
HEAD_FILL = "E6E8EB"
TOTAL_FILL = "F7F7F7"
CONTENT_CM = 17.2

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "output"
DOCX_NAME = "劳动合同书_上海贝迪创建科技有限公司_胡继刚.docx"

COMPANY = "上海贝迪创建科技有限公司"
EMPLOYEE = "胡继刚"

# 2026年8月工资表。比例与金额均可整除，脚本启动时复核。
BASE = 28000
PENSION_CO, PENSION_EE = 0.16, 0.08
MEDICAL_CO, MEDICAL_EE = 0.085, 0.02
LOCAL_MED_CO = 0.005
UNEMP_CO, UNEMP_EE = 0.005, 0.005
INJURY_CO = 0.004
HF_CO, HF_EE = 0.05, 0.05
IIT_AUG = 1866


def money(rate: float) -> int:
    value = BASE * rate
    if abs(value - round(value)) > 1e-6:
        raise SystemExit(f"缴费金额不是整数：基数 {BASE} × {rate} = {value}")
    return int(round(value))


PENSION_CO_AMT = money(PENSION_CO)
PENSION_EE_AMT = money(PENSION_EE)
MEDICAL_CO_AMT = money(MEDICAL_CO)
MEDICAL_EE_AMT = money(MEDICAL_EE)
LOCAL_MED_AMT = money(LOCAL_MED_CO)
UNEMP_CO_AMT = money(UNEMP_CO)
UNEMP_EE_AMT = money(UNEMP_EE)
INJURY_AMT = money(INJURY_CO)
HF_CO_AMT = money(HF_CO)
HF_EE_AMT = money(HF_EE)

ER_SI = PENSION_CO_AMT + MEDICAL_CO_AMT + LOCAL_MED_AMT + UNEMP_CO_AMT + INJURY_AMT
EE_SI = PENSION_EE_AMT + MEDICAL_EE_AMT + UNEMP_EE_AMT
SI_TOTAL = ER_SI + EE_SI
HF_TOTAL = HF_CO_AMT + HF_EE_AMT
SI_HF_TOTAL = SI_TOTAL + HF_TOTAL
NET_AUG = BASE - EE_SI - HF_EE_AMT - IIT_AUG
PROBATION_FLOOR = int(BASE * 0.8)

assert (ER_SI, EE_SI, SI_TOTAL, HF_TOTAL, SI_HF_TOTAL, NET_AUG, PROBATION_FLOOR) == (
    7252,
    2940,
    10192,
    2800,
    12992,
    21794,
    22400,
)


def rmb_daxie(num: int) -> str:
    """整数金额转为人民币大写，以“元整”结尾。"""
    if num < 0:
        raise ValueError(num)
    if num == 0:
        return "零元整"
    chars = "零壹贰叁肆伍陆柒捌玖"
    small = ["", "拾", "佰", "仟"]
    big = ["", "万", "亿", "兆"]
    groups: list[int] = []
    n = num
    while n:
        groups.append(n % 10000)
        n //= 10000
    pieces: list[str] = []
    for gi in range(len(groups) - 1, -1, -1):
        group = groups[gi]
        if group == 0:
            continue
        part = ""
        need_zero = False
        for i in range(3, -1, -1):
            digit = group // (10**i) % 10
            if digit == 0:
                if part:
                    need_zero = True
            else:
                if need_zero:
                    part += "零"
                    need_zero = False
                part += chars[digit] + small[i]
        if pieces and group < 1000:
            part = "零" + part
        pieces.append(part + big[gi])
    return "".join(pieces) + "元整"


def _self_check_daxie() -> None:
    expected = {
        28000: "贰万捌仟元整",
        22400: "贰万贰仟肆佰元整",
        21794: "贰万壹仟柒佰玖拾肆元整",
        10192: "壹万零壹佰玖拾贰元整",
        12992: "壹万贰仟玖佰玖拾贰元整",
        1866: "壹仟捌佰陆拾陆元整",
        800: "捌佰元整",
        10: "壹拾元整",
        10001: "壹万零壹元整",
        20000: "贰万元整",
        10100: "壹万零壹佰元整",
        110: "壹佰壹拾元整",
    }
    for amount, text in expected.items():
        got = rmb_daxie(amount)
        if got != text:
            raise SystemExit(f"大写金额错误：{amount} -> {got}，期望 {text}")


_self_check_daxie()

WAGE_DX = rmb_daxie(BASE)
FLOOR_DX = rmb_daxie(PROBATION_FLOOR)


def cn_num(n: int) -> str:
    digits = "零一二三四五六七八九"
    if n <= 0 or n >= 100:
        raise ValueError(n)
    if n < 10:
        return digits[n]
    if n == 10:
        return "十"
    if n < 20:
        return "十" + digits[n - 10]
    tens, ones = divmod(n, 10)
    return digits[tens] + "十" + (digits[ones] if ones else "")


def fmt(amount: int, decimals: bool = True) -> str:
    if decimals:
        return f"{amount:,.2f}"
    return f"{amount:,}"


def pct(rate: float) -> str:
    text = f"{rate * 100:.1f}".rstrip("0").rstrip(".")
    return text + "%"


def set_run_font(run, name: str, size: float, bold: bool = False, color: str = INK) -> None:
    run.bold = bold
    run.italic = False
    run.font.size = Pt(size)
    run.font.name = name
    run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._element.get_or_add_rPr()
    r_fonts = rpr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        rpr.append(r_fonts)
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        r_fonts.set(qn(f"w:{attr}"), name)


def set_paragraph_format(
    paragraph,
    *,
    before: float = 0,
    after: float = 0,
    line: float = 1.15,
    indent: bool = False,
    align: str = "left",
    keep_next: bool = False,
    keep_lines: bool = False,
) -> None:
    fmt_p = paragraph.paragraph_format
    fmt_p.space_before = Pt(before)
    fmt_p.space_after = Pt(after)
    fmt_p.line_spacing = line
    fmt_p.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    fmt_p.first_line_indent = Pt(20) if indent else Pt(0)
    fmt_p.widow_control = True
    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    paragraph.alignment = align_map[align]
    ppr = paragraph._p.get_or_add_pPr()
    if keep_next:
        if ppr.find(qn("w:keepNext")) is None:
            ppr.append(OxmlElement("w:keepNext"))
    if keep_lines:
        if ppr.find(qn("w:keepLines")) is None:
            ppr.append(OxmlElement("w:keepLines"))


def add_marked_runs(paragraph, text: str, size: float, font: str = SERIF, color: str = INK) -> None:
    parts = text.split("**")
    if len(parts) % 2 == 0:
        raise ValueError(f"加粗标记不成对：{text}")
    for index, part in enumerate(parts):
        if not part:
            continue
        run = paragraph.add_run(part)
        set_run_font(run, font, size, bold=(index % 2 == 1), color=color)


def add_text(
    doc,
    text: str,
    *,
    size: float = 10,
    font: str = SERIF,
    bold: bool = False,
    before: float = 0,
    after: float = 2,
    line: float = 1.0,
    indent: bool = False,
    align: str = "justify",
    color: str = INK,
    keep_next: bool = False,
    page_break: bool = False,
) -> None:
    paragraph = doc.add_paragraph()
    set_paragraph_format(
        paragraph,
        before=before,
        after=after,
        line=line,
        indent=indent,
        align=align,
        keep_next=keep_next,
    )
    if page_break:
        paragraph.paragraph_format.page_break_before = True
    if bold:
        text = f"**{text}**"
    add_marked_runs(paragraph, text, size, font, color)
    return paragraph


def add_paragraph_border(paragraph, *, edge: str = "bottom", sz: str = "12", color: str = "1A1A1A", space: str = "1") -> None:
    ppr = paragraph._p.get_or_add_pPr()
    borders = ppr.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        ppr.append(borders)
    line = OxmlElement(f"w:{edge}")
    line.set(qn("w:val"), "single")
    line.set(qn("w:sz"), sz)
    line.set(qn("w:space"), space)
    line.set(qn("w:color"), color)
    borders.append(line)


def shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, margin: int = 40) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.find(qn("w:tcMar"))
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for edge in ("top", "left", "bottom", "right"):
        node = tc_mar.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(margin))
        node.set(qn("w:type"), "dxa")


def set_table_widths(table, widths_cm: list[float]) -> None:
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


def set_table_borders(table, color: str = LINE, sz: str = "4") -> None:
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


def set_row_height(row, height_cm: float) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tr_height = tr_pr.find(qn("w:trHeight"))
    if tr_height is None:
        tr_height = OxmlElement("w:trHeight")
        tr_pr.append(tr_height)
    tr_height.set(qn("w:val"), str(int(height_cm * 567)))
    tr_height.set(qn("w:hRule"), "atLeast")
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def prevent_row_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def mark_header_row(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:tblHeader")) is None:
        tr_pr.append(OxmlElement("w:tblHeader"))


def write_cell(
    cell,
    text: str,
    *,
    bold: bool = False,
    size: float = 10.5,
    align: str = "left",
    fill: str | None = None,
    font: str = SERIF,
    color: str = INK,
) -> None:
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_margins(cell)
    if fill:
        shade_cell(cell, fill)
    paragraph = cell.paragraphs[0]
    set_paragraph_format(paragraph, before=0, after=0, line=1.0, align=align)
    # 清空默认空 run
    if paragraph.runs:
        paragraph.runs[0].text = ""
    add_marked_runs(paragraph, f"**{text}**" if bold else text, size, font, color)


def add_table(doc, rows: list[list[dict]], widths: list[float], header: bool = False, min_height: float = 0.72):
    table = doc.add_table(rows=len(rows), cols=len(widths))
    set_table_widths(table, widths)
    set_table_borders(table)
    for r_index, row_data in enumerate(rows):
        row = table.rows[r_index]
        set_row_height(row, min_height)
        if header and r_index == 0:
            mark_header_row(row)
        for c_index, spec in enumerate(row_data):
            write_cell(table.cell(r_index, c_index), **spec)
    return table


def label(text: str, align: str = "center") -> dict:
    return {"text": text, "bold": True, "align": align, "fill": LABEL_FILL, "size": 10.5}


def value(text: str, *, bold: bool = False, align: str = "left") -> dict:
    return {"text": text, "bold": bold, "align": align, "size": 10.5}


def head(text: str, align: str = "center") -> dict:
    return {"text": text, "bold": True, "align": align, "fill": HEAD_FILL, "size": 9.5}


def num(amount: int, *, bold: bool = False, fill: str | None = None, empty: str | None = None) -> dict:
    text = empty if empty is not None else fmt(amount)
    spec = {"text": text, "bold": bold, "align": "right", "size": 10}
    if fill:
        spec["fill"] = fill
    return spec


def blank(width: int = 18) -> str:
    return "＿" * width


def add_field(paragraph, instruction: str, placeholder: str = "1") -> None:
    run_begin = paragraph.add_run()
    set_run_font(run_begin, SERIF, 9, color=MUTED)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    run_begin._r.append(begin)

    run_instr = paragraph.add_run()
    set_run_font(run_instr, SERIF, 9, color=MUTED)
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruction} "
    run_instr._r.append(instr)

    run_sep = paragraph.add_run()
    set_run_font(run_sep, SERIF, 9, color=MUTED)
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    run_sep._r.append(separate)

    run_text = paragraph.add_run(placeholder)
    set_run_font(run_text, SERIF, 9, color=MUTED)

    run_end = paragraph.add_run()
    set_run_font(run_end, SERIF, 9, color=MUTED)
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run_end._r.append(end)


def configure_section(section) -> None:
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.9)
    section.right_margin = Cm(1.9)
    section.top_margin = Cm(1.28)
    section.bottom_margin = Cm(1.08)
    section.header_distance = Cm(0.35)
    section.footer_distance = Cm(0.28)

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    set_paragraph_format(hp, align="center", before=0, after=0, line=1.0)
    run = hp.add_run(f"{COMPANY}  ·  劳动合同书")
    set_run_font(run, SANS, 8, color=MUTED)
    add_paragraph_border(hp, edge="bottom", sz="6", color="1A1A1A", space="1")

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    set_paragraph_format(fp, align="center", before=0, after=0, line=1.0)
    add_paragraph_border(fp, edge="top", sz="6", color="1A1A1A", space="1")
    left = fp.add_run("第 ")
    set_run_font(left, SERIF, 8, color=MUTED)
    add_field(fp, "PAGE")
    mid = fp.add_run(" 页 / 共 ")
    set_run_font(mid, SERIF, 8, color=MUTED)
    add_field(fp, "NUMPAGES")
    right = fp.add_run(" 页")
    set_run_font(right, SERIF, 8, color=MUTED)

    sect_pr = section._sectPr
    grid = sect_pr.find(qn("w:docGrid"))
    if grid is not None:
        sect_pr.remove(grid)


def configure_styles(doc: Document) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = SERIF
    normal.font.size = Pt(10)
    normal.font.color.rgb = RGBColor.from_string(INK)
    rpr = normal.element.get_or_add_rPr()
    r_fonts = rpr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        rpr.append(r_fonts)
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        r_fonts.set(qn(f"w:{attr}"), SERIF)
    pf = normal.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0

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


class ContractBuilder:
    def __init__(self, doc: Document) -> None:
        self.doc = doc
        self.article_no = 0

    def chapter(self, text: str) -> None:
        add_text(
            self.doc,
            text,
            size=12,
            font=SANS,
            bold=True,
            before=5,
            after=1,
            line=1.0,
            align="left",
            keep_next=True,
        )

    def article(self, *paragraphs: str) -> None:
        self.article_no += 1
        number = f"第{cn_num(self.article_no)}条"
        for index, text in enumerate(paragraphs):
            paragraph = self.doc.add_paragraph()
            set_paragraph_format(
                paragraph,
                before=0,
                after=1 if index == len(paragraphs) - 1 else 0,
                line=1.0,
                indent=True,
                align="justify",
            )
            if index == 0:
                run = paragraph.add_run(f"{number}  ")
                set_run_font(run, SERIF, 10, bold=True)
            add_marked_runs(paragraph, text, 10)

    def body(self, text: str, **kwargs) -> None:
        add_text(self.doc, text, indent=True, align="justify", **kwargs)


def build() -> Document:
    doc = Document()
    configure_styles(doc)
    configure_section(doc.sections[0])
    core = doc.core_properties
    core.title = "劳动合同书"
    core.subject = f"{COMPANY}与{EMPLOYEE}固定期限劳动合同"
    core.category = "劳动合同"
    core.keywords = "劳动合同;固定期限;胡继刚;高级工程师;工程管理部"
    core.comments = "报酬及社保公积金数据依据2026年8月工资表填写；未知主体信息留空。"

    b = ContractBuilder(doc)

    add_text(
        doc,
        f"合同编号：{blank(12)}",
        size=9,
        align="right",
        before=0,
        after=0,
        line=1.0,
    )
    add_text(
        doc,
        "劳 动 合 同 书",
        size=18,
        font=SANS,
        bold=True,
        align="center",
        before=2,
        after=0,
        line=1.0,
    )
    subtitle = add_text(
        doc,
        "（固定期限）",
        size=11,
        font=SANS,
        align="center",
        before=0,
        after=2,
        line=1.0,
        color=MUTED,
    )
    add_paragraph_border(subtitle, edge="bottom", sz="12", color="1A1A1A", space="1")

    add_text(doc, "甲方（用人单位）", size=11, font=SANS, bold=True, before=4, after=1, align="left")
    add_table(
        doc,
        [
            [label("名        称"), value(COMPANY, bold=True)],
            [label("统一社会信用代码"), value(blank(18))],
            [label("法定代表人或主要负责人"), value(blank(16))],
            [label("住        所"), value("注册地址：" + blank(16))],
            [label("实际办公地址"), value(blank(18))],
            [label("联  系  人"), value("联系人：" + blank(6) + "    联系电话：" + blank(8))],
        ],
        [5.5, 11.7],
        min_height=0.52,
    )

    add_text(doc, "乙方（劳动者）", size=11, font=SANS, bold=True, before=4, after=1, align="left")
    employee_table = add_table(
        doc,
        [
            [
                label("姓        名"),
                value(EMPLOYEE, bold=True, align="center"),
                label("出 生 年 月"),
                value("1987年4月", bold=True, align="center"),
            ],
            [
                label("居民身份证号码"),
                value(blank(10)),
                label("性        别"),
                value(blank(6)),
            ],
            [label("户籍地址"), value(blank(20)), value(""), value("")],
            [label("现居住及通讯地址"), value(blank(20)), value(""), value("")],
            [
                label("联系方式"),
                value("联系电话：" + blank(6) + "    电子邮箱：" + blank(10)),
                value(""),
                value(""),
            ],
        ],
        [4.0, 4.9, 3.4, 4.9],
        min_height=0.52,
    )
    for row_index in (2, 3, 4):
        employee_table.cell(row_index, 1).merge(employee_table.cell(row_index, 3))

    b.body(
        "根据《中华人民共和国劳动法》《中华人民共和国劳动合同法》《上海市劳动合同条例》及其他有关法律、法规、规章，甲乙双方在平等自愿、协商一致、诚实信用的基础上订立本合同，共同遵守。",
        before=4,
        after=1,
    )

    add_text(
        doc,
        "合同主要事项",
        size=11,
        font=SANS,
        bold=True,
        before=4,
        after=1,
        align="left",
        keep_next=True,
    )
    add_table(
        doc,
        [
            [head("事项"), head("约定")],
            [label("合同类型"), value("固定期限劳动合同", bold=True)],
            [
                label("合同期限"),
                value("自2026年5月1日起至2029年4月30日止，共3年", bold=True),
            ],
            [label("劳动关系建立"), value("自实际用工之日起建立", bold=True)],
            [
                label("试用期"),
                value("自2026年5月1日起至2026年8月1日止，共3个月", bold=True),
            ],
            [
                label("试用期满后月工资"),
                value(
                    "乙方在正常出勤并提供正常劳动时，月工资为人民币28,000元（大写：人民币贰万捌仟元整，税前）",
                    bold=True,
                ),
            ],
        ],
        [4.9, 12.3],
        header=True,
        min_height=0.5,
    )
    add_text(
        doc,
        "上表与第一章、第四章相应条款一致。试用期包含在劳动合同期限内。",
        size=9,
        before=2,
        after=1,
        indent=False,
        align="left",
        color=MUTED,
    )

    b.chapter("第一章  劳动合同期限和试用期")
    b.article(
        "本合同为固定期限劳动合同。期限自**2026年5月1日**起至**2029年4月30日**止，共3年。劳动关系自实际用工之日起建立。"
    )
    b.article(
        "试用期自**2026年5月1日**起至**2026年8月1日**止，共3个月，包含在第一条的劳动合同期限内。本合同期限为3年，上述试用期不超过《中华人民共和国劳动合同法》第十九条规定的六个月上限。",
        "同一用人单位与同一劳动者只能约定一次试用期。双方按照实际用工及已经成立的约定履行，不得追溯补设、重复约定或重新起算试用期，也不得任意延长试用期。",
    )
    b.article(
        "甲方在试用期开始前，应向乙方明确告知合法、合理的录用条件及考核要求。甲方在试用期内解除本合同的，应当有证据证明乙方不符合录用条件，向乙方说明理由，并依法办理。试用期内，乙方提前三日通知甲方，可以解除本合同。"
    )

    b.chapter("第二章  工作内容和工作地点")
    b.article(
        f"乙方的工作部门为**工程管理部**，岗位（专业技术职务）为**高级工程师**。乙方主要从事建筑工程技术及项目管理，包括工程项目技术管理、施工组织协调、质量安全控制、工程技术方案管理及项目全过程管理，并按照岗位职责和甲方依法告知的工作安排履行职责。",
        "甲方2026年8月工资表的职位栏为空白；该表“社保、公积金”部分记载的所属部门为“**总裁办**”。双方确认，乙方的工作部门和岗位以本条前款为准。甲方办理工资发放、社会保险和住房公积金登记时，所使用的部门、岗位信息应当与本条约定以及乙方实际从事的工作一致。",
    )
    b.article(
        "乙方的主要工作职责包括下列事项：",
    )
    for item in (
        "（一）工程项目技术管理：组织或参与技术交底、技术审查，协调解决实施过程中的工程技术问题。",
        "（二）施工组织协调：参与施工组织设计和进度安排，协调设计、施工、监理等相关单位的技术配合。",
        "（三）质量安全控制：落实工程质量及安全技术要求，开展检查、问题整改和风险防范。",
        "（四）工程技术方案管理：组织或参与技术方案的编制、论证、优化、实施检查和技术资料管理。",
        "（五）项目全过程管理：参与项目策划、建设实施、验收移交等阶段的技术和项目管理工作。",
    ):
        b.body(item, before=0, after=1)
    b.article(
        "甲方可以根据生产经营需要，在不改变乙方岗位性质和本合同约定的劳动报酬的前提下，合理调整乙方承担的具体项目。变更工作部门、岗位、劳动报酬或其他重要内容的，应当由双方协商一致，并采用书面形式。"
    )
    b.article(
        f"乙方的工作地点为**上海市**。具体办公或项目地址：{blank(8)}。",
        "因项目需要在上海市范围内变更具体工作场所的，甲方应当提前告知，该等变更不视为变更工作城市。甲方安排乙方到上海市以外临时出差的，应当合理安排行程，并按规定承担差旅费用。长期变更工作城市的，应当协商一致并书面确认。",
    )

    b.chapter("第三章  工作时间和休息休假")
    b.article(
        "乙方实行标准工时制度，每日工作时间不超过八小时，每周工作时间不超过四十小时，每周至少休息一日。具体作息时间由甲方依法确定并告知乙方。实行特殊工时制度的，甲方应当依法办理审批手续，并向乙方书面说明适用安排。"
    )
    b.article(
        "甲方依法保障乙方的休息日、法定节假日和带薪年休假，以及乙方依法享有的婚假、丧假、产假、育儿假、陪产假、病假和其他假期。",
        "甲方安排加班的，应当与乙方协商，遵守法定时限，并依法支付加班工资或者安排补休。甲方不得强迫或者变相强迫乙方加班。",
    )

    b.chapter("第四章  劳动报酬")
    b.article(
        f"乙方试用期满后，在正常出勤并提供正常劳动时，月工资为人民币**{fmt(BASE, False)}元**（大写：**人民币{WAGE_DX}**，税前）。",
        "工资构成如下：基本工资人民币**28,000元**；岗位补贴无固定金额；其他补贴无固定金额；提成无固定金额。上述月工资对应法定标准工时内的正常劳动，与甲方2026年8月工资表记载的应发合计一致。",
        "奖金、一次性补贴或其他不固定项目，只在甲方依法制定并向乙方公示或者告知后，按照该制度发放，不计入本条固定月工资。加班工资依法另计，不包含在上述月工资之内。",
    )
    b.article(
        f"试用期月工资为人民币**{fmt(BASE, False)}元**（大写：**人民币{WAGE_DX}**，税前）。",
        f"乙方在试用期提供正常劳动的，试用期工资不得低于本合同约定的转正后月工资的百分之八十，即不得低于人民币**{fmt(PROBATION_FLOOR, False)}元**（大写：**人民币{FLOOR_DX}**），不得低于本单位相同岗位最低档工资，并不得低于上海市当期最低工资标准。本条约定的试用期工资为转正月工资的百分之百，已经高于上述百分之八十的下限。国家和本市另有更高强制性标准的，从其规定。",
        "甲方2026年8月工资表记载，乙方该月应发工资为人民币28,000元。试用期至2026年8月1日止，2026年8月已经适用试用期满后的月工资标准，与前款约定一致。试用期月工资不另降低，与试用期满后的月工资相同。",
    )
    b.article(
        f"甲方按月以人民币支付工资，每月{blank(4)}日支付上一自然月工资。支付日遇休息日或法定节假日的，提前至最近的工作日支付。工资发放方式：{blank(8)}。",
        "甲方每次支付工资时，应当向乙方提供工资清单，列明应发项目、代扣的社会保险费和个人住房公积金、代扣个人所得税以及实发金额。",
    )
    b.article(
        "甲方依法代扣代缴个人所得税，以及乙方个人应当缴纳的社会保险费和住房公积金。乙方提供正常劳动的，其工资待遇不得违反上海市最低工资规定。最低工资的构成项目和扣除口径，按国家和上海市的规定执行。",
        "乙方缺勤或请假的，甲方按照国家和上海市规定以及依法公示的规章制度核算工资，但不得以罚款形式克扣工资。",
    )
    b.article(
        "在标准工时制度下，甲方安排乙方在工作日延长工作时间的，支付不低于工资百分之一百五十的加班工资；安排乙方在休息日工作、又不能安排补休的，支付不低于工资百分之二百的加班工资；安排乙方在法定节假日工作的，支付不低于工资百分之三百的加班工资。",
        "加班工资计发基数不得低于本合同约定的转正后月工资。日工资按照月工资除以月计薪天数21.75计算，小时工资按照日工资除以8小时计算。国家和上海市另有强制性规定的，从其规定。",
    )
    b.article(
        "乙方依法享受年休假、婚丧假、产假和其他假期，以及在规定的医疗期内停工治疗，或者甲方停工停产的，相关工资待遇按照国家和上海市规定执行。变更本合同约定的工资标准，应当由双方协商一致，并采用书面形式。"
    )

    b.chapter("第五章  社会保险和住房公积金")
    b.article(
        "甲乙双方按照国家和上海市规定，参加职工基本养老保险、职工基本医疗保险、失业保险、工伤保险，并依法享受生育保险相关待遇。生育保险的办理口径以上海市现行规定为准。",
        "甲方自实际用工之日起三十日内，为乙方办理社会保险登记，并按时足额缴费。试用期计入参保缴费期间。乙方个人应当缴纳的部分，由甲方从工资中代扣代缴。甲方不得以现金或其他补贴代替法定缴费。",
        "根据甲方2026年8月工资及社会保险缴纳记录，该月各项社会保险的缴费基数均为人民币**28,000元**，与本合同约定的月工资一致。该月实际缴费比例如下，只作为该月的执行口径。缴费比例或者缴费基数上下限依法调整的，按照调整后的规定执行：",
    )
    for item in (
        "（一）养老保险：单位缴纳16%，个人缴纳8%；",
        "（二）医疗保险：单位缴纳8.5%，个人缴纳2%；",
        "（三）地方附加医疗保险：单位缴纳0.5%；",
        "（四）失业保险：单位缴纳0.5%，个人缴纳0.5%；",
        "（五）工伤保险：单位缴纳0.4%，个人不缴纳。具体费率以社会保险经办机构核定的为准。",
    ):
        b.body(item, before=0, after=1)
    b.body(
        "甲方应当按照乙方的实际工资和上海市规定申报缴费基数，不得违反规定瞒报、漏报或者少缴。该月单位和个人的具体金额见本合同附件。",
        before=1,
        after=3,
    )
    b.article(
        "甲方自实际用工之日起，为乙方办理住房公积金缴存登记，并按时足额缴存。根据2026年8月缴存记录，缴存基数为人民币**28,000元**，单位和个人的缴存比例**各为5%**。该比例处于上海市规定的缴存比例区间内。",
        "以后年度调整缴存基数，或者在法定区间内调整缴存比例的，甲方应当按照规定执行并告知乙方。甲方承担单位缴存部分，乙方承担个人缴存部分。甲方不得以现金补贴代替住房公积金缴存。",
    )

    b.chapter("第六章  劳动保护、劳动条件和职业危害防护")
    b.article(
        "甲方应当提供符合国家规定的劳动安全卫生条件和必要的劳动防护用品、工作设备，对乙方进行安全生产、职业卫生和岗位技能培训，并如实告知工程现场及本岗位可能存在的职业危害、防护措施和应急方法。甲方依法安排职业健康检查。"
    )
    b.article(
        "乙方应当遵守安全技术操作规程，正确使用防护用品，发现事故隐患或者其他不安全因素时及时报告。乙方有权拒绝违章指挥和强令冒险作业，不因拒绝上述行为而受到不利处分。",
        "乙方发生工伤、患职业病或者被诊断、鉴定为疑似职业病的，甲方应当依法申请工伤认定、协助进行劳动能力鉴定，并落实相应待遇。",
    )

    b.chapter("第七章  保密与知识产权")
    b.article(
        "乙方对在工作中知悉的、依法属于甲方商业秘密的信息，以及与甲方知识产权有关且尚未公开的技术信息和经营信息，负有保密义务，不得擅自披露、使用或者允许他人使用。",
        "保密义务不适用于已经合法公开的信息、乙方有证据证明合法独立取得的信息，以及法律法规要求向行政机关、司法机关披露的情形。乙方依法举报、控告违法行为，不受本条限制。",
    )
    b.article(
        "乙方履行职务形成的技术方案、图纸、工程资料、计算机软件、发明创造和其他工作成果，其知识产权归属、署名、奖励报酬和使用权限，依照法律规定执行；双方另有合法书面约定的，从其约定。",
        "乙方离职时，应当按照甲方的合理要求交接工作资料和物品。依法继续有效的保密义务，不因劳动合同解除或终止而消失。",
    )
    b.article(
        "本合同不直接约定竞业限制。确需约定竞业限制的，双方应当另行签订书面协议，明确限制的人员范围、业务范围、地域、期限，以及解除或终止劳动合同后按月支付的经济补偿。竞业限制期限不得超过二年。",
        "甲方未按约定支付经济补偿的，乙方可以依法解除竞业限制约定。除法定服务期和依法成立的竞业限制以外，本合同不约定由乙方承担违约金。",
    )

    b.chapter("第八章  规章制度")
    b.article(
        "甲方制定、修改或者决定有关劳动报酬、工作时间、休息休假、劳动安全卫生、保险福利、职工培训、劳动纪律以及劳动定额管理等直接涉及乙方切身利益的规章制度或者重大事项，应当依法履行讨论、协商、公示或者告知程序。",
        "乙方应当遵守已经依法制定并向其公示或者告知的规章制度和职业规范。规章制度与本合同不一致，或者违反法律、法规的，按照法律、法规和本合同执行。甲方不得以规章制度免除自身法定责任、加重乙方责任或者排除乙方的主要权利。",
    )
    b.article(
        "甲方不得扣押乙方的居民身份证、职称证书、资格证书或者其他证件原件，不得要求乙方提供担保或者以其他名义向乙方收取财物。",
        "乙方应当如实说明与劳动合同直接相关的基本情况，诚信履行岗位职责。",
    )

    b.chapter("第九章  劳动合同的变更、解除和终止")
    b.article(
        "变更本合同，应当由双方协商一致，并采用书面形式。变更文本由双方各执一份。",
        "甲方变更名称、法定代表人、主要负责人或者投资人等事项，不影响本合同履行。甲方发生合并或者分立的，本合同由承继其权利和义务的用人单位继续履行。",
    )
    b.article(
        "双方协商一致，可以解除本合同。",
        "乙方提前三十日以书面形式通知甲方，可以解除本合同。乙方在试用期内提前三日通知甲方，可以解除本合同。",
        "甲方未按照本合同约定支付劳动报酬、未依法为乙方缴纳社会保险费，或者存在《中华人民共和国劳动合同法》第三十八条所列其他情形的，乙方可以依法解除本合同；符合可以立即解除条件的，依法执行。",
    )
    b.article(
        "甲方单方解除本合同，应当符合《中华人民共和国劳动合同法》第三十九条、第四十条、第四十一条规定的条件，并依法履行说明理由、通知和工会程序。",
        "甲方依照《中华人民共和国劳动合同法》第四十条解除本合同的，应当提前三十日以书面形式通知乙方，或者额外支付乙方一个月工资。本合同不得另行扩大甲方单方解除劳动合同的法定范围。",
    )
    b.article(
        "乙方患病或者非因工负伤，在规定的医疗期内的，或者具有《中华人民共和国劳动合同法》第四十二条所列其他情形的，甲方不得依照该法第四十条、第四十一条解除本合同。",
        "劳动合同期满，或者出现其他法定终止情形的，本合同依法终止。依照法律规定应当续延的，延续至相应情形消失时终止。乙方符合订立无固定期限劳动合同法定条件并提出订立的，甲方应当订立无固定期限劳动合同。",
    )
    b.article(
        "解除或者终止本合同时，甲方应当出具解除或者终止劳动合同的证明，并在十五日内为乙方办理档案和社会保险关系转移手续，依法结清工资及其他应当支付的费用。",
        "乙方应当按照合法约定办理工作交接。甲方不得以工作交接为由扣押乙方证件，也不得拒绝出具法定证明。",
    )

    b.chapter("第十章  经济补偿")
    b.article(
        "解除或者终止本合同，符合《中华人民共和国劳动合同法》第四十六条规定的，甲方应当向乙方支付经济补偿。",
        "固定期限劳动合同期满终止的，除甲方维持或者提高劳动合同约定条件续订劳动合同、乙方不同意续订的以外，甲方依法支付经济补偿。",
    )
    b.article(
        "经济补偿按乙方在甲方的工作年限计算，每满一年支付一个月工资。六个月以上不满一年的，按一年计算；不满六个月的，支付半个月工资。",
        "月工资按乙方在劳动合同解除或者终止前十二个月的平均工资计算；工作不满十二个月的，按实际工作的月数计算平均工资。月工资高于本市上年度职工月平均工资三倍的，经济补偿的标准和年限按照法律规定的上限执行。",
    )
    b.article(
        "甲方应当在乙方办理完毕工作交接时，向乙方支付依法应当支付的经济补偿。",
        "甲方违法解除或者终止本合同，乙方要求继续履行劳动合同的，甲方应当继续履行；乙方不要求继续履行，或者本合同已经不能继续履行的，甲方应当依照《中华人民共和国劳动合同法》第八十七条，按照经济补偿标准的二倍向乙方支付赔偿金。",
    )

    b.chapter("第十一章  争议处理")
    b.article(
        "因本合同的订立、履行、变更、解除或终止发生争议的，双方可以协商解决，也可以向有权调解组织申请调解。协商、调解不是申请仲裁的必经程序。",
        "双方也可以依法向劳动合同履行地或者甲方所在地有管辖权的劳动人事争议仲裁委员会申请仲裁。对仲裁裁决不服、依法可以提起诉讼或者申请撤销的，按照法定程序办理。",
    )
    b.article(
        "申请劳动争议仲裁的时效期间为一年。时效的起算、中断、中止，以及劳动关系存续期间拖欠劳动报酬发生争议的特殊时效，按照法律规定执行。乙方依法投诉、举报或者向有关部门主张权利，不因本合同的任何约定受到限制。"
    )

    b.chapter("第十二章  其他约定")
    b.article(
        "本合同附件《劳动报酬、社会保险及住房公积金确认表（2026年8月）》是本合同的组成部分，与正文具有同等效力。",
        "附件记载的是2026年8月实际发生的工资发放和缴纳数据，用来确认月工资标准以及该月的缴费、缴存口径。以后各月的个人所得税按照税法重新计算，不以附件所载2026年8月的税额作为固定扣款。",
    )
    b.article(
        "双方确认，本合同记载的用人单位名称、劳动者姓名、合同期限、工作部门、岗位、工作地点和月工资，应当与实际用工、工资发放、社会保险缴纳和住房公积金缴存情况一致。甲方向乙方或者有关主管部门提供劳动关系、岗位和工资证明时，应当与本合同以及实际履行情况相符。"
    )
    b.article(
        "本合同首部和正文中留有下划线的事项，包括合同编号、统一社会信用代码、法定代表人或主要负责人、注册地址、实际办公地址、联系人、联系电话、乙方居民身份证号码、性别、户籍地址、现居住及通讯地址、电子邮箱、具体办公或项目地址、每月发薪日和工资发放方式，在签署时尚未填写的，由双方据实填写。",
        "下划线空白未经双方书面确认，不视为已经作出约定。填写错误需要修改的，修改处由双方签字或者盖章确认。",
    )
    b.article(
        "双方的通讯地址和联系电话，以本合同填写的内容为准。任何一方变更的，应当在变更后七日内书面通知对方。书面通知可以当面交付、邮寄，或者采用双方确认并且能够留存记录的电子方式送达。受送达方拒绝签收或者无法送达的，按照法律规定和实际送达证据认定效力。"
    )
    b.article(
        "本合同未尽事宜，按照国家和上海市现行法律、法规、规章执行。本合同的约定与强制性规定不一致的，按照强制性规定执行。双方就专项技术培训依法约定服务期的，另行签订书面协议。合法有效的书面补充协议是本合同的组成部分。"
    )

    # 附件
    add_text(
        doc,
        "附件  劳动报酬、社会保险及住房公积金确认表",
        size=12,
        font=SANS,
        bold=True,
        align="center",
        before=8,
        after=0,
        line=1.0,
    )
    add_text(
        doc,
        "（依据甲方2026年8月计薪工资表）",
        size=9,
        font=SANS,
        align="center",
        before=0,
        after=2,
        line=1.0,
        color=MUTED,
    )
    b.body(
        f"本附件确认甲方向乙方**{EMPLOYEE}**发放2026年8月工资，以及该月社会保险、住房公积金的缴纳情况。本附件是本合同组成部分。表内金额单位为人民币元。",
        before=1,
        after=1,
    )

    add_text(doc, "一、应发工资", size=10.5, font=SANS, bold=True, before=3, after=1, align="left", keep_next=True)
    add_table(
        doc,
        [
            [head("项目"), head("金额"), head("说明")],
            [value("基本工资", bold=True), num(BASE, bold=True), value("月薪，对应正常出勤")],
            [value("岗位补贴"), num(0, empty="—"), value("无固定金额")],
            [value("其他补贴"), num(0, empty="—"), value("无固定金额")],
            [value("提成"), num(0, empty="—"), value("无固定金额")],
            [value("扣款"), num(0, empty="—"), value("该月工资表记载为无")],
            [
                {"text": "应发合计", "bold": True, "align": "left", "fill": TOTAL_FILL, "size": 10.5},
                num(BASE, bold=True, fill=TOTAL_FILL),
                {"text": "与本合同约定的月工资一致", "bold": False, "align": "left", "fill": TOTAL_FILL, "size": 10.5},
            ],
        ],
        [3.8, 3.6, 9.8],
        header=True,
        min_height=0.42,
    )
    add_text(
        doc,
        "工资表“计薪天数”“职位”两栏未填写。本合同月工资按月薪28,000元约定，不按日工资倒推。",
        size=9,
        before=1,
        after=1,
        indent=False,
        align="left",
        color=MUTED,
    )

    add_text(doc, "二、社会保险", size=10.5, font=SANS, bold=True, before=4, after=1, align="left", keep_next=True)
    si_rows = [
        [
            head("险种"),
            head("缴费基数"),
            head("单位比例"),
            head("单位金额"),
            head("个人比例"),
            head("个人金额"),
        ],
        si_line("养老保险", PENSION_CO, PENSION_CO_AMT, PENSION_EE, PENSION_EE_AMT),
        si_line("医疗保险", MEDICAL_CO, MEDICAL_CO_AMT, MEDICAL_EE, MEDICAL_EE_AMT),
        si_line("地方附加医疗保险", LOCAL_MED_CO, LOCAL_MED_AMT, None, None),
        si_line("失业保险", UNEMP_CO, UNEMP_CO_AMT, UNEMP_EE, UNEMP_EE_AMT),
        si_line("工伤保险", INJURY_CO, INJURY_AMT, None, None),
        [
            {"text": "合计", "bold": True, "align": "center", "fill": TOTAL_FILL, "size": 10},
            {"text": fmt(BASE), "bold": True, "align": "right", "fill": TOTAL_FILL, "size": 10},
            {"text": "—", "bold": True, "align": "center", "fill": TOTAL_FILL, "size": 10},
            num(ER_SI, bold=True, fill=TOTAL_FILL),
            {"text": "—", "bold": True, "align": "center", "fill": TOTAL_FILL, "size": 10},
            num(EE_SI, bold=True, fill=TOTAL_FILL),
        ],
    ]
    add_table(doc, si_rows, [3.7, 2.7, 2.3, 2.8, 2.3, 3.4], header=True, min_height=0.4)
    add_text(
        doc,
        f"社会保险单位与个人合计 {fmt(SI_TOTAL)} 元。生育保险以上海市现行与职工基本医疗保险的衔接规定为准，2026年8月工资表未单列生育保险缴费。",
        size=9,
        before=1,
        after=1,
        indent=False,
        align="left",
        color=MUTED,
    )

    add_text(doc, "三、住房公积金", size=10.5, font=SANS, bold=True, before=4, after=1, align="left", keep_next=True)
    add_table(
        doc,
        [
            [head("项目"), head("缴存基数"), head("比例"), head("金额")],
            [value("单位缴存"), num(BASE), value(pct(HF_CO), align="center"), num(HF_CO_AMT)],
            [value("个人缴存"), num(BASE), value(pct(HF_EE), align="center"), num(HF_EE_AMT)],
            [
                {"text": "合计", "bold": True, "align": "left", "fill": TOTAL_FILL, "size": 10.5},
                {"text": "—", "bold": True, "align": "right", "fill": TOTAL_FILL, "size": 10},
                {"text": "—", "bold": True, "align": "center", "fill": TOTAL_FILL, "size": 10},
                num(HF_TOTAL, bold=True, fill=TOTAL_FILL),
            ],
        ],
        [4.5, 4.2, 3.8, 4.7],
        header=True,
        min_height=0.4,
    )

    add_text(doc, "四、2026年8月代扣与实发", size=10.5, font=SANS, bold=True, before=4, after=1, align="left", keep_next=True)
    add_table(
        doc,
        [
            [head("项目"), head("金额"), head("备注")],
            [value("应发合计"), num(BASE, bold=True), value("　基本工资")],
            [value("减：个人养老保险"), num(PENSION_EE_AMT), value("　8%")],
            [value("减：个人医疗保险"), num(MEDICAL_EE_AMT), value("　2%")],
            [value("减：个人失业保险"), num(UNEMP_EE_AMT), value("　0.5%")],
            [value("减：个人住房公积金"), num(HF_EE_AMT), value("　5%")],
            [value("减：个人所得税"), num(IIT_AUG), value("　仅为本月累计预扣数")],
            [
                {"text": "本月实发", "bold": True, "align": "left", "fill": TOTAL_FILL, "size": 10.5},
                num(NET_AUG, bold=True, fill=TOTAL_FILL),
                {"text": "　应发减去上述各项代扣", "bold": False, "align": "left", "fill": TOTAL_FILL, "size": 10.5},
            ],
        ],
        [5.5, 3.8, 7.9],
        header=True,
        min_height=0.4,
    )
    add_text(
        doc,
        "个人所得税1,866.00元是2026年8月按照累计预扣法代扣的金额，随累计收入、专项扣除和专项附加扣除变化，不作为本合同的固定扣款。",
        size=9,
        before=1,
        after=1,
        indent=False,
        align="left",
        color=MUTED,
    )

    add_text(doc, "五、与工资表合计数的勾稽", size=10.5, font=SANS, bold=True, before=4, after=1, align="left", keep_next=True)
    add_table(
        doc,
        [
            [head("勾稽项目"), head("金额")],
            [value("社会保险单位缴纳合计"), num(ER_SI)],
            [value("社会保险个人缴纳合计"), num(EE_SI)],
            [value("社会保险合计"), num(SI_TOTAL, bold=True)],
            [value("住房公积金单位与个人合计"), num(HF_TOTAL, bold=True)],
            [
                {"text": "社会保险与住房公积金总计", "bold": True, "align": "left", "fill": TOTAL_FILL, "size": 10.5},
                num(SI_HF_TOTAL, bold=True, fill=TOTAL_FILL),
            ],
        ],
        [10.8, 6.4],
        header=True,
        min_height=0.4,
    )
    add_text(
        doc,
        "工资表右下角合计数12,992.00元，等于社会保险合计10,192.00元加上住房公积金合计2,800.00元，与上表一致。该合计数不是从乙方工资中重复扣除的项目。",
        size=9,
        before=1,
        after=1,
        indent=False,
        align="left",
        color=MUTED,
    )

    # 签署页
    add_text(
        doc,
        "签署页",
        size=12,
        font=SANS,
        bold=True,
        align="center",
        before=4,
        after=0,
        line=1.0,
    )
    add_text(
        doc,
        "（正文及附件完）",
        size=9,
        align="center",
        before=0,
        after=1,
        line=1.0,
        color=MUTED,
    )
    b.article(
        "本合同经甲方盖章，并由甲方的法定代表人或者授权代表签字，同时由乙方本人签字后生效。劳动合同期限以及劳动关系的建立时间，按照第一条执行。",
        "本合同一式两份，甲乙双方各执一份，具有同等法律效力。双方已经阅读并理解全部条款。签署前应当核对已经填写的内容，补齐必要的空白事项；不需要约定的空白处注明“无”或者划线注销。",
    )

    sign = add_table(
        doc,
        [
            [
                label("甲方（用人单位）"),
                value(""),
                label("乙方（劳动者）"),
                value(""),
            ],
            [
                label("单位名称"),
                value(COMPANY, bold=True),
                label("劳动者"),
                value(EMPLOYEE, bold=True),
            ],
            [
                label("授权代表签字"),
                value(blank(8)),
                label("乙方本人签字"),
                value(blank(8)),
            ],
            [
                label("甲方盖章"),
                value("（盖章处）"),
                label("签字"),
                value("（签字处）"),
            ],
            [
                label("签署日期"),
                value("____年__月__日"),
                label("签署日期"),
                value("____年__月__日"),
            ],
            [
                label("签署地点"),
                value(blank(6)),
                label("签署地点"),
                value(blank(6)),
            ],
            [
                {"text": "合同文本领取确认：乙方已领取本合同文本一份。", "bold": False, "align": "left", "size": 9},
                value(""),
                label("乙方签字"),
                value("____年__月__日"),
            ],
        ],
        [2.8, 5.8, 2.8, 5.8],
        min_height=0.42,
    )
    sign.cell(0, 0).merge(sign.cell(0, 1))
    sign.cell(0, 2).merge(sign.cell(0, 3))
    sign.cell(6, 0).merge(sign.cell(6, 1))
    set_row_height(sign.rows[2], 1.05)
    set_row_height(sign.rows[3], 1.45)

    return doc


def si_line(name: str, co_rate: float, co_amt: int, ee_rate: float | None, ee_amt: int | None) -> list[dict]:
    return [
        value(name),
        num(BASE),
        value(pct(co_rate), align="center"),
        num(co_amt),
        value("—" if ee_rate is None else pct(ee_rate), align="center"),
        num(0, empty="—") if ee_amt is None else num(ee_amt),
    ]


def assert_document(doc: Document) -> None:
    import re

    texts: list[str] = []
    own_articles: list[str] = []
    for paragraph in doc.paragraphs:
        texts.append(paragraph.text)
        match = re.match(r"^(第[零一二三四五六七八九十]+条)", paragraph.text)
        if match:
            own_articles.append(match.group(1))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                texts.append(cell.text)
    full = "\n".join(texts)
    expected_articles = [f"第{cn_num(i)}条" for i in range(1, 41)]
    if own_articles != expected_articles:
        raise SystemExit(
            "条款编号不连续。实际为："
            + "、".join(own_articles)
            + "；期望共40条。"
        )
    required = [
        COMPANY,
        EMPLOYEE,
        "1987年4月",
        "2026年5月1日",
        "2029年4月30日",
        "2026年8月1日",
        "共3年",
        "共3个月",
        "合同主要事项",
        "固定期限劳动合同",
        "工程管理部",
        "高级工程师",
        "建筑工程技术及项目管理",
        "工程项目技术管理",
        "施工组织协调",
        "质量安全控制",
        "工程技术方案管理",
        "项目全过程管理",
        "上海市",
        "28,000",
        WAGE_DX,
        "22,400",
        FLOOR_DX,
        "百分之八十",
        "最低工资",
        "总裁办",
        "16%",
        "8.5%",
        "0.5%",
        "0.4%",
        "5%",
        "21,794.00",
        "1,866.00",
        "12,992.00",
        "10,192.00",
        "7,252.00",
        "2,940.00",
        "标准工时",
        "经济补偿",
        "劳动人事争议仲裁",
        "保密",
        "知识产权",
        "竞业限制",
        "签署页",
        "合同文本领取确认",
    ]
    missing = [item for item in required if item not in full]
    if missing:
        raise SystemExit("合同缺少必要内容：" + "、".join(missing))
    forbidden = [
        "上海贝迪创建材科技有限公司",
        "刘李杨",
        "蕰川路",
        "91310000",
        "2026年10月31日",
        "共六个月",
        "共三年",
        "处于第二条约定的试用期内",
    ]
    hit = [item for item in forbidden if item in full]
    if hit:
        raise SystemExit("合同写入了不应出现的内容：" + "、".join(hit))
    if own_articles[-1] != "第四十条":
        raise SystemExit("签署页条款编号异常。")


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc = build()
    assert_document(doc)
    out = OUT_DIR / DOCX_NAME
    doc.save(out)
    print(f"已生成：{out}")
    print(f"条款数检查通过。转正月工资 {fmt(BASE)} 元，试用期工资 {fmt(BASE)} 元，80%下限 {fmt(PROBATION_FLOOR)} 元。")
    print(f"2026年8月实发 {fmt(NET_AUG)} 元；社保公积金总计 {fmt(SI_HF_TOTAL)} 元。")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"生成失败：{exc}", file=sys.stderr)
        raise
