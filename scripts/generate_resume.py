#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成胡继刚履历：Word + HTML + PDF。

版式：两页 A4 专家履历，午夜蓝 + 香槟金。
口径按来稿：现任只写复旦中心秘书长；一线投资计至 2021；
2026 四项按职责动词落笔，不列引进企业名单。
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
import time
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output"
TPL = Path(__file__).resolve().parent / "templates"
FONT_DIR = Path(__file__).resolve().parent / "fonts"

INK = RGBColor(0x0B, 0x1F, 0x3A)
INK_HEX = "0B1F3A"
GOLD = RGBColor(0xC5, 0xA5, 0x72)
GOLD_DEEP = RGBColor(0x9A, 0x7B, 0x4F)
GOLD_HEX = "C5A572"
TEXT = RGBColor(0x24, 0x30, 0x40)
MUTED = RGBColor(0x5E, 0x6A, 0x78)
LINE = "E6DCCB"
SOFT = "F7F3EB"
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CREAM = RGBColor(0xF3, 0xE6, 0xC8)


def set_run_font(run, *, name="微软雅黑", size=10.5, bold=False, color=None, east_asia=None):
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = name
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    ea = east_asia or name
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:eastAsia"), ea)
    if color is not None:
        run.font.color.rgb = color


def shade_cell(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_borders(cell, color="FFFFFF", sz="0"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    old = tcPr.find(qn("w:tcBorders"))
    if old is not None:
        tcPr.remove(old)
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:color"), color)
        el.set(qn("w:space"), "0")
        borders.append(el)
    tcPr.append(borders)


def set_cell_margin(cell, top=40, bottom=40, left=80, right=80):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for name, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def v_align(cell, val="center"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    va = OxmlElement("w:vAlign")
    va.set(qn("w:val"), val)
    tcPr.append(va)


def set_table_fixed(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), "5000")
    tblW.set(qn("w:type"), "pct")


def prevent_row_split(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    cant = OxmlElement("w:cantSplit")
    trPr.append(cant)


def disable_autofit(table):
    table.autofit = False
    table.allow_autofit = False
    set_table_fixed(table)


def set_col_widths(table, widths_cm):
    disable_autofit(table)
    for row in table.rows:
        prevent_row_split(row)
        for i, w in enumerate(widths_cm):
            if i < len(row.cells):
                row.cells[i].width = Cm(w)


def clear_cell(cell):
    cell.text = ""
    for p in cell.paragraphs[1:]:
        p._element.getparent().remove(p._element)


def para_in(
    cell,
    text,
    *,
    size=10,
    bold=False,
    color=None,
    align=WD_ALIGN_PARAGRAPH.LEFT,
    space_after=0,
    space_before=0,
):
    p = cell.paragraphs[0] if not cell.paragraphs[0].text and not cell.paragraphs[0].runs else cell.add_paragraph()
    if cell.paragraphs[0].text == "" and not cell.paragraphs[0].runs:
        p = cell.paragraphs[0]
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = 1.08
    run = p.add_run(text)
    set_run_font(run, name="微软雅黑", size=size, bold=bold, color=color, east_asia="微软雅黑")
    return p


def disable_auto_space(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    for tag in ("autoSpaceDE", "autoSpaceDN"):
        el = pPr.find(qn(f"w:{tag}"))
        if el is None:
            el = OxmlElement(f"w:{tag}")
            pPr.append(el)
        el.set(qn("w:val"), "0")
    adj = pPr.find(qn("w:adjustRightInd"))
    if adj is None:
        adj = OxmlElement("w:adjustRightInd")
        pPr.append(adj)
    adj.set(qn("w:val"), "0")


def finalize_docx(doc):
    for p in doc.paragraphs:
        disable_auto_space(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    disable_auto_space(p)
    # 取消东亚文档网格，避免 LibreOffice / Word 按网格撑开行距
    sectPr = doc.sections[0]._sectPr
    docGrid = sectPr.find(qn("w:docGrid"))
    if docGrid is None:
        docGrid = OxmlElement("w:docGrid")
        sectPr.append(docGrid)
    docGrid.set(qn("w:type"), "default")
    docGrid.set(qn("w:linePitch"), "240")
    # 兼容设置
    settings = doc.settings._element
    compat = settings.find(qn("w:compat"))
    if compat is None:
        compat = OxmlElement("w:compat")
        settings.append(compat)
    for tag in ("doNotExpandShiftReturn", "balanceSingleByteDoubleByteWidth"):
        el = compat.find(qn(f"w:{tag}"))
        if el is None:
            el = OxmlElement(f"w:{tag}")
            compat.append(el)


def add_p(doc, text, *, size=10.5, bold=False, color=None, space_before=0, space_after=4, align=None, line=1.15):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color, east_asia="微软雅黑")
    return p


def add_hr(doc, color=GOLD_HEX, sz="12"):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def section_title(doc, text):
    p = add_p(doc, text, size=11, bold=True, color=INK, space_before=4, space_after=1)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "18")
    left.set(qn("w:space"), "4")
    left.set(qn("w:color"), GOLD_HEX)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), LINE)
    pBdr.append(left)
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def bullet(doc, text, *, size=9):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.28)
    pf.first_line_indent = Cm(-0.28)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0.8)
    pf.line_spacing = 1.08
    run = p.add_run("●  " + text)
    set_run_font(run, size=size, color=TEXT, east_asia="微软雅黑")
    return p


def job_header(doc, company, title, dates):
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(table, [15.2, 3.3])
    left, right = table.rows[0].cells
    clear_cell(left)
    clear_cell(right)
    set_cell_borders(left)
    set_cell_borders(right)
    set_cell_margin(left, 20, 8, 0, 40)
    set_cell_margin(right, 20, 8, 40, 0)
    p = left.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    r1 = p.add_run(company)
    set_run_font(r1, size=10.5, bold=True, color=INK, east_asia="微软雅黑")
    r2 = p.add_run("  ·  " + title)
    set_run_font(r2, size=9.5, bold=False, color=MUTED, east_asia="微软雅黑")
    p2 = right.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.paragraph_format.space_after = Pt(0)
    r3 = p2.add_run(dates)
    set_run_font(r3, size=9.5, bold=True, color=GOLD_DEEP, east_asia="微软雅黑")
    return table


def styled_table(doc, rows, widths, *, first_col_bold=True):
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(table, widths)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.rows[r].cells[c]
            clear_cell(cell)
            set_cell_borders(cell, LINE, "4")
            set_cell_margin(cell, 18, 18, 40, 40)
            v_align(cell, "center")
            if r == 0:
                shade_cell(cell, INK_HEX)
                para_in(cell, val, size=8.5, bold=True, color=WHITE)
            else:
                if r % 2 == 0:
                    shade_cell(cell, SOFT)
                para_in(
                    cell,
                    val,
                    size=8.5,
                    bold=(first_col_bold and c == 0),
                    color=INK if c == 0 else TEXT,
                )
    return table


def add_footer(section):
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.space_before = Pt(2)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "6")
    top.set(qn("w:space"), "4")
    top.set(qn("w:color"), LINE)
    pBdr.append(top)
    pPr.append(pBdr)
    r = p.add_run("产业招商 · 城市更新 · 智能制造")
    set_run_font(r, size=8, color=MUTED, east_asia="微软雅黑")
    r2 = p.add_run("　　胡继刚  ·  ")
    set_run_font(r2, size=8, color=MUTED, east_asia="微软雅黑")

    def add_field(paragraph, instr):
        run = paragraph.add_run()
        set_run_font(run, size=8, color=MUTED, east_asia="微软雅黑")
        r = run._r
        fld1 = OxmlElement("w:fldChar")
        fld1.set(qn("w:fldCharType"), "begin")
        it = OxmlElement("w:instrText")
        it.set(qn("xml:space"), "preserve")
        it.text = instr
        fld2 = OxmlElement("w:fldChar")
        fld2.set(qn("w:fldCharType"), "end")
        r.append(fld1)
        r.append(it)
        r.append(fld2)

    add_field(p, " PAGE ")
    r4 = p.add_run(" / ")
    set_run_font(r4, size=8, color=MUTED, east_asia="微软雅黑")
    add_field(p, " NUMPAGES ")


def set_table_edges(table, *, top=None, bottom=None):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        if edge == "top" and top:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), top[1])
            el.set(qn("w:color"), top[0])
        elif edge == "bottom" and bottom:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), bottom[1])
            el.set(qn("w:color"), bottom[0])
        else:
            el.set(qn("w:val"), "nil")
        el.set(qn("w:space"), "0")
        borders.append(el)
    tblPr.append(borders)


def build_mast(doc):
    head = doc.add_table(rows=1, cols=2)
    head.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(head, [13.2, 5.3])
    set_table_edges(head, top=(GOLD_HEX, "16"), bottom=(GOLD_HEX, "12"))
    left, right = head.rows[0].cells
    for cell in (left, right):
        clear_cell(cell)
        shade_cell(cell, INK_HEX)
        set_cell_borders(cell, INK_HEX, "0")
        v_align(cell, "center")
    set_cell_margin(left, 50, 50, 100, 50)
    set_cell_margin(right, 50, 50, 20, 100)

    p = left.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("胡继刚")
    set_run_font(r, size=20, bold=True, color=WHITE, east_asia="微软雅黑")
    r2 = p.add_run("   HU JIGANG")
    set_run_font(r2, size=9, color=GOLD, east_asia="微软雅黑")
    p2 = left.add_paragraph()
    p2.paragraph_format.space_before = Pt(1)
    p2.paragraph_format.space_after = Pt(0)
    r = p2.add_run("产业招商 · 城市更新 · 智能制造")
    set_run_font(r, size=10, bold=True, color=CREAM, east_asia="微软雅黑")
    r = p2.add_run("   副教授级高级工程师 · 复旦大学硕士研究生")
    set_run_font(r, size=8.5, color=RGBColor(0xD5, 0xDE, 0xE8), east_asia="微软雅黑")
    p3 = left.add_paragraph()
    p3.paragraph_format.space_after = Pt(0)
    r = p3.add_run("现任复旦大学住房政策研究中心秘书长　·　1987 年生 · 九三学社 · 上海 · IELTS 6.0")
    set_run_font(r, size=8, color=RGBColor(0xC5, 0xD0, 0xDC), east_asia="微软雅黑")
    p6 = left.add_paragraph()
    p6.paragraph_format.space_before = Pt(2)
    p6.paragraph_format.space_after = Pt(0)
    r = p6.add_run("求职方向  产业招商 / 城市更新 / 存量空间 / 国资平台  ·  上海")
    set_run_font(r, size=8.5, bold=True, color=GOLD, east_asia="微软雅黑")

    p = right.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("18678408669")
    set_run_font(r, size=9, color=WHITE, east_asia="微软雅黑")
    for line in ("262782809@qq.com", "微信 hu262782809"):
        pp = right.add_paragraph()
        pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pp.paragraph_format.space_after = Pt(0)
        rr = pp.add_run(line)
        set_run_font(rr, size=8.5, color=RGBColor(0xD9, 0xE2, 0xEC), east_asia="微软雅黑")


def build_docx() -> Path:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(0)
    section.bottom_margin = Cm(1.0)
    section.left_margin = Cm(1.25)
    section.right_margin = Cm(1.25)
    section.header_distance = Cm(0)
    section.footer_distance = Cm(0.4)
    add_footer(section)

    # 默认段落更紧凑
    styles = doc.styles["Normal"]
    styles.font.name = "微软雅黑"
    styles.font.size = Pt(10.5)
    rPr = styles.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), "微软雅黑")
    styles.paragraph_format.space_after = Pt(0)
    styles.paragraph_format.line_spacing = 1.08

    build_mast(doc)

    stats = doc.add_table(rows=1, cols=4)
    stats.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(stats, [4.7, 4.7, 4.7, 4.7])
    items = [
        ("2026 已签约产业合作", "4 项", "定位招商 · 委托协同"),
        ("一线投拓 2011–2021", "10 年", "中南 · 新城 · 万科"),
        ("靖江 + 金坛", "43 亿", "勾地 / 收并购代表作"),
        ("华为汽车两园区", "3.5 万㎡", "辅助全流程落地"),
    ]
    for i, (k, v, sub) in enumerate(items):
        c = stats.rows[0].cells[i]
        clear_cell(c)
        shade_cell(c, SOFT)
        set_cell_borders(c, LINE, "4")
        # 金色顶边：用上边框加粗金色
        tc = c._tc
        tcPr = tc.get_or_add_tcPr()
        borders = tcPr.find(qn("w:tcBorders"))
        top = borders.find(qn("w:top"))
        top.set(qn("w:color"), GOLD_HEX)
        top.set(qn("w:sz"), "16")
        set_cell_margin(c, 28, 28, 50, 50)
        para_in(c, k, size=8, color=MUTED, space_after=0)
        p = c.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(v)
        set_run_font(run, size=14, bold=True, color=INK, east_asia="微软雅黑")
        p2 = c.add_paragraph()
        p2.paragraph_format.space_before = Pt(1)
        p2.paragraph_format.space_after = Pt(0)
        run = p2.add_run(sub)
        set_run_font(run, size=8, color=GOLD_DEEP, east_asia="微软雅黑")

    add_p(
        doc,
        "2011–2021 年一线投拓，历任中南、新城、万科至上海区域投资副总经理。"
        "2021 年起任复旦大学住房政策研究中心秘书长，兼杨浦区科技企业联合会执行会长。"
        "2026 年已签约创智汇、东方枢纽、华为汽车两园区、森马产业园等产业合作，"
        "覆盖定位招商、联动对接与全流程落地辅助。一线投资只计至 2021 年。"
        "能够连接政府、高校、企业与空间载体，推动城市更新、产业导入和存量盘活。",
        size=9.5,
        color=TEXT,
        space_before=5,
        space_after=1,
        line=1.16,
    )

    section_title(doc, "现任")
    job_header(doc, "复旦大学住房政策研究中心", "秘书长 · 上海", "2021.06 – 至今")
    add_p(
        doc,
        "副教授级高级工程师为职称，不是本职务行政级别。",
        size=9,
        color=GOLD_DEEP,
        space_before=2,
        space_after=2,
    )
    for x in [
        "牵头住房保障、租赁与城市更新课题及产学研合作；兼杨浦区科技企业联合会执行会长。2026 年已签约四项产业合作，见下表。",
        "2025.05「全球新经济增长引擎峰会」人民网、中新社报道：以秘书长身份作为主办方代表发言，并做房地产圆桌。",
        "2026.03 北欧创新国际会客厅落户杨浦；区政府官网点名执行会长。宝龙商办为共同发起单位，不是任职。",
    ]:
        bullet(doc, x)

    section_title(doc, "2026 已签约产业合作")
    deals = [
        ("项目", "规模", "职责"),
        ("创智汇", "约 6600 ㎡", "负责人工智能与 IP 定位招商，辅助向智能制造、出海及人工智能转型"),
        ("东方枢纽", "约 143 万㎡", "推动与产业方对接，促成联动招商"),
        ("华为汽车园区", "静安 1.5 万㎡ + 浦东 2 万㎡", "辅助拿地、规划国土、定位至招商落地全流程"),
        ("闵行森马产业园", "约 2 万㎡", "工业空间定位与招商方案拟定"),
    ]
    styled_table(doc, deals, [3.6, 5.2, 9.4])
    add_p(doc, "均为已签约合作，职责按合同。表内不列引进企业名单。", size=8.5, color=MUTED, space_before=3, space_after=2)

    section_title(doc, "社会职务")
    add_p(
        doc,
        "上海市杨浦区科技企业联合会执行会长（区政府官网）　·　复旦大学不动产资产管理协会创始人　·　上海山东省商会理事　·　复旦研究生管理联考面试官",
        size=9,
        color=TEXT,
        space_before=1,
        space_after=2,
        line=1.12,
    )

    section_title(doc, "专业经历")
    job_header(doc, "万科企业股份有限公司 · 上海区域江浙事业部", "投资副总经理", "2017.03 – 2021.06")
    add_p(doc, "上海 · 汇报投资合作开发部总经理 · 团队 8 人", size=9, color=MUTED, space_before=2, space_after=2)
    for x in [
        "深耕江浙未进入城市，负责策略勾地、收并购、债务重组与不良资产处置。",
        "政府 / 国企合作：海宁代建 18.53 亿、约 23 万㎡；镇江大港与大港集团各半，约 199 亩 / 32 万方。",
        "存量与轻资产：龙湖滟澜山法拍，测算 IRR 约 16.16%；靖江印象城勾地 + 印力委管，约 15 亿 / 556 亩 / 7 万方购物中心。",
        "产业载体：主导获取并落地万纬平湖、奉贤南桥冷链园；金坛理想城约 28.05 亿 / 561 亩底价获取。",
        "以收并购、策略勾地、不良处置等方式新增 4 宗土地，进入盐城、台州、常州金坛。",
    ]:
        bullet(doc, x)

    section_title(doc, "专业经历（续）")
    job_header(doc, "新城控股集团股份有限公司 · 总部战略投资中心", "资深专业经理 · 上海", "2015.11 – 2017.02")
    for x in [
        "负责山东、江苏等区域战略布局与土地获取。",
        "临沂吾悦：对接市级重点招商引资政策，推动商业综合体落地，系集团首进山东的重资产项目。",
        "青岛吾悦：轻资产品牌输出 + 委管，约 13.68 万㎡。",
    ]:
        bullet(doc, x)

    job_header(doc, "中南建设集团股份有限公司 · 总部新项目发展中心", "投资拓展高级经理 · 上海 / 南通", "2011.06 – 2015.10")
    add_p(
        doc,
        "2011.06–2012.10 总裁办战略企划；其后新项目发展中心。从集团战略到项目投资的完整历练。",
        size=9,
        color=MUTED,
        space_before=2,
        space_after=2,
    )
    for x in [
        "获取如皋中南世纪城、苏州中南锦苑；江阴白鹭湾与碧桂园同股同权合作。",
        "推进美国养老、日本永旺等国际合作对接（推进中，非已签约）。2014 年底派驻四川任西南拓展负责人。",
    ]:
        bullet(doc, x)

    section_title(doc, "社会职务（工作要点）")
    for x in [
        "杨浦科企联：主持日常运作，策划科创项目评审、投融资对接与政策宣讲。2026.03 区政府官网点名执行会长，参与北欧创新国际会客厅落户杨浦宝龙旭辉广场环创中心。",
        "复旦不动产资管协会：2018 年起创立面向校友的资管交流平台，组织行业分享、项目考察与资源对接。",
    ]:
        bullet(doc, x)

    section_title(doc, "代表项目")
    add_p(
        doc,
        "金额与面积证明投资判断与政府 / 国企协同；引进企业、签约入驻需另补经营台账。",
        size=8.5,
        color=MUTED,
        space_before=1,
        space_after=3,
    )
    rows = [
        ("项目", "角色", "规模", "方式"),
        ("靖江印象城", "主导获取", "15 亿 / 556 亩", "勾地 + 轻资产委管"),
        ("嘉兴海宁商住", "主导拓展", "18.53 亿 / 23 万㎡", "政府代建"),
        ("金坛理想城", "主导获取", "28.05 亿 / 561 亩", "收并购 + 债务重组"),
        ("镇江大港金域蓝湾", "主导", "199 亩 / 32 万方", "收并购 + 国企各半"),
        ("龙湖滟澜山", "主导", "IRR ≈ 16.16%", "法拍不良资产"),
        ("南桥 / 平湖冷链园", "主导落地", "—", "产业勾地 / 收购改造"),
        ("临沂吾悦广场", "主导勾地", "—", "市级招商政策协同"),
        ("青岛吾悦广场", "获取", "13.68 万㎡", "轻资产输出 + 委管"),
    ]
    styled_table(doc, rows, [4.4, 2.6, 4.6, 6.6])

    section_title(doc, "教育背景")
    job_header(doc, "复旦大学", "工商管理硕士 · 财务金融方向", "2018.09 – 2021.06")
    add_p(doc, "在职攻读，与万科上海区域任职并行。", size=9, color=MUTED, space_before=1, space_after=4)
    job_header(doc, "中国海洋大学", "土木工程学士", "2007.09 – 2011.07")

    section_title(doc, "资质与研究")
    add_p(
        doc,
        "职称与证书：副教授级高级工程师；会计师中级；PMP；住建委项目管理工程师；IELTS 6.0；普通话一级甲等；BIM；上海市人才引进认证资质。",
        size=9,
        color=TEXT,
        space_before=1,
        space_after=1,
        line=1.12,
    )
    add_p(
        doc,
        "论文、专利与研修：《万科物流地产平台业务发展战略研究》（知网）；《房屋建筑工程人工智能技术的应用》（《住宅与房地产》）。"
        "实用新型专利 2 项（功能脚架、防护围栏）；交大城市治理数字化转型高研班。",
        size=9,
        color=TEXT,
        space_before=0,
        space_after=1,
        line=1.12,
    )

    tags = (
        "产业招商　定位招商　智能制造　城市更新　存量空间　国资合作　"
        "轻资产委管　产业园区　政企协同　企业出海　住房保障　募投管退"
    )
    add_p(
        doc,
        "专业方向：" + tags,
        size=9,
        color=INK,
        space_before=4,
        space_after=0,
    )

    finalize_docx(doc)
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "胡继刚-简历.docx"
    doc.save(path)
    return path


FONT_URLS = {
    "NotoSansSC.ttf": "https://github.com/google/fonts/raw/main/ofl/notosanssc/NotoSansSC%5Bwght%5D.ttf",
    "NotoSerifSC.ttf": "https://github.com/google/fonts/raw/main/ofl/notoserifsc/NotoSerifSC%5Bwght%5D.ttf",
}


def ensure_fonts() -> None:
    FONT_DIR.mkdir(parents=True, exist_ok=True)
    for name, url in FONT_URLS.items():
        path = FONT_DIR / name
        if path.exists() and path.stat().st_size > 1_000_000:
            continue
        print(f"下载字体 {name} …")
        subprocess.run(["curl", "-fsSL", "-o", str(path), url], check=True)


def copy_fonts() -> None:
    ensure_fonts()
    dest = OUT / "fonts"
    dest.mkdir(parents=True, exist_ok=True)
    for name in FONT_URLS:
        src = FONT_DIR / name
        if src.exists():
            shutil.copy2(src, dest / name)


def write_html() -> Path:
    OUT.mkdir(parents=True, exist_ok=True)
    copy_fonts()
    html = (TPL / "resume.html").read_text(encoding="utf-8")
    path = OUT / "胡继刚-简历.html"
    path.write_text(html, encoding="utf-8")
    return path


def write_notes() -> Path:
    text = """# 胡继刚履历 · 排版说明

两页 A4，午夜蓝 + 香槟金。现任只写复旦大学住房政策研究中心秘书长。一线投资计至 2021。2026 四项按职责动词落笔，表内不列引进企业名单。

## 来稿取舍

- 社会职务与复旦中心拆开；科企联执行会长见杨浦区政府官网。
- 工商联房地产商会不写入现任（公开秘书长为李祥）。
- 资管协会写创始人，不写现任秘书长。
- 「上海市城市更新研究会等机构评审专家」来稿标注自述待核，未上印刷面。
- 印刷面不出现「待核 / 待补 / 引进华为 / 创智天地」。

## 文件

- `output/胡继刚-简历.docx`　Word
- `output/胡继刚-简历.pdf`　PDF（由 HTML 按 A4 导出，版式与预览一致）
- `output/胡继刚-简历.html`　浏览器预览，可再「导出 PDF」

重新生成：`python3 scripts/generate_resume.py`
"""
    path = OUT / "排版说明.md"
    path.write_text(text, encoding="utf-8")
    return path


def export_pdf(html_path: Path, pdf_path: Path) -> None:
    html_uri = html_path.resolve().as_uri()
    if pdf_path.exists():
        pdf_path.unlink()
    with tempfile.TemporaryDirectory(prefix="chrome-resume-") as tmp:
        cmd = [
            "google-chrome",
            "--headless=new",
            "--no-sandbox",
            "--disable-gpu",
            "--disable-dev-shm-usage",
            "--disable-extensions",
            "--disable-background-networking",
            f"--user-data-dir={tmp}",
            "--no-first-run",
            "--no-pdf-header-footer",
            "--virtual-time-budget=4000",
            f"--print-to-pdf={pdf_path}",
            html_uri,
        ]
        try:
            subprocess.run(cmd, check=True, timeout=12)
        except subprocess.TimeoutExpired:
            time.sleep(0.5)
            if not pdf_path.exists() or pdf_path.stat().st_size < 1000:
                raise
            print(f"chrome 打印后未退出，但已生成：{pdf_path.name}")


def main():
    docx_path = build_docx()
    html_path = write_html()
    notes = write_notes()
    pdf_path = OUT / "胡继刚-简历.pdf"
    export_pdf(html_path, pdf_path)
    print(f"docx: {docx_path}")
    print(f"html: {html_path}")
    print(f"pdf: {pdf_path}")
    print(f"notes: {notes}")


if __name__ == "__main__":
    main()
