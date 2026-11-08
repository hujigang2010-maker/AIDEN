#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把《专业技术资格评审申报表》PDF 转成内容一致的 Word。

版式按原表分页：封面、承诺与填表说明、基本情况、继续教育、工作经历、
业绩、论文、工作总结（4 页）、考核、核实与组织意见、评审备案。
证件照、手写签名和红色印章从 PDF 中裁出后嵌入。
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor, Twips

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"
OUTPUT = ROOT / "output" / "专业技术资格评审申报表.docx"

PAGE_W = 180.0  # 表格内容宽度，毫米


def mm_to_twip(mm: float) -> int:
    return int(round(mm * 56.6929133858))


def set_run_font(run, size=10.5, bold=False, font="宋体", color=None, underline=False):
    run.bold = bold
    run.underline = underline
    run.font.size = Pt(size)
    run.font.name = "Times New Roman"
    if color is not None:
        run.font.color.rgb = RGBColor(*color)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), "Times New Roman")
    rfonts.set(qn("w:hAnsi"), "Times New Roman")
    rfonts.set(qn("w:cs"), "Times New Roman")
    rfonts.set(qn("w:eastAsia"), font)


def set_style_font(style, size=10.5, font="宋体"):
    style.font.name = "Times New Roman"
    style.font.size = Pt(size)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), "Times New Roman")
    rfonts.set(qn("w:hAnsi"), "Times New Roman")
    rfonts.set(qn("w:eastAsia"), font)


def shade_cell(cell, fill: str):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = tcpr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcpr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=30, bottom=30, left=50, right=50):
    tcpr = cell._tc.get_or_add_tcPr()
    tcmar = tcpr.find(qn("w:tcMar"))
    if tcmar is None:
        tcmar = OxmlElement("w:tcMar")
        tcpr.append(tcmar)
    else:
        for child in list(tcmar):
            tcmar.remove(child)
    for edge, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcmar.append(node)


def set_cell_borders(cell, sz="8", color="000000"):
    tcpr = cell._tc.get_or_add_tcPr()
    borders = tcpr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcpr.append(borders)
    else:
        for child in list(borders):
            borders.remove(child)
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)


def set_row_height(row, mm, rule="atLeast"):
    trpr = row._tr.get_or_add_trPr()
    trh = trpr.find(qn("w:trHeight"))
    if trh is None:
        trh = OxmlElement("w:trHeight")
        trpr.append(trh)
    trh.set(qn("w:val"), str(mm_to_twip(mm)))
    trh.set(qn("w:hRule"), rule)


def prevent_row_split(row):
    trpr = row._tr.get_or_add_trPr()
    if trpr.find(qn("w:cantSplit")) is None:
        trpr.append(OxmlElement("w:cantSplit"))


def apply_widths(table, widths):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.allow_autofit = False
    tbl = table._tbl
    tblpr = tbl.tblPr
    tblw = tblpr.find(qn("w:tblW"))
    if tblw is None:
        tblw = OxmlElement("w:tblW")
        tblpr.append(tblw)
    tblw.set(qn("w:type"), "dxa")
    tblw.set(qn("w:w"), str(mm_to_twip(sum(widths))))
    layout = tblpr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblpr.append(layout)
    layout.set(qn("w:type"), "fixed")
    grid = tbl.find(qn("w:tblGrid"))
    cols = list(grid)
    for i, gc in enumerate(cols):
        if i < len(widths):
            gc.set(qn("w:w"), str(mm_to_twip(widths[i])))
    for tr in tbl.tr_lst:
        idx = 0
        for tc in tr.tc_lst:
            tcpr = tc.get_or_add_tcPr()
            gs = tcpr.find(qn("w:gridSpan"))
            span = int(gs.get(qn("w:val"))) if gs is not None else 1
            width = sum(widths[idx:idx + span])
            tcw = tcpr.find(qn("w:tcW"))
            if tcw is None:
                tcw = OxmlElement("w:tcW")
                tcpr.append(tcw)
            tcw.set(qn("w:w"), str(mm_to_twip(width)))
            tcw.set(qn("w:type"), "dxa")
            idx += span


def stack(text: str) -> str:
    return "\n".join(text)


def new_table(doc, rows, widths, heights=None, header_rows=None, sz="8"):
    table = doc.add_table(rows=rows, cols=len(widths))
    table.style = "Table Grid"
    apply_widths(table, widths)
    header_rows = header_rows or set()
    for r, row in enumerate(table.rows):
        prevent_row_split(row)
        if heights and r < len(heights) and heights[r]:
            set_row_height(row, heights[r])
        for cell in row.cells:
            set_cell_borders(cell, sz=sz)
            set_cell_margins(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if r in header_rows:
                shade_cell(cell, "D0D0D0")
    return table


def merge(table, r1, c1, r2, c2):
    if (r1, c1) != (r2, c2):
        table.cell(r1, c1).merge(table.cell(r2, c2))


def write(table, r, c, text, size=9, bold=False, font="宋体", align="center",
          valign="center", color=None, fill=None):
    cell = table.cell(r, c)
    if valign == "top":
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    elif valign == "bottom":
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
    else:
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    if fill:
        shade_cell(cell, fill)
    lines = text if isinstance(text, (list, tuple)) else str(text).split("\n")
    tc = cell._tc
    for extra in tc.xpath("./w:p")[1:]:
        tc.remove(extra)
    # 清掉合并时带进来的嵌套内容以外的文字
    p0 = cell.paragraphs[0]
    for run in list(p0.runs):
        run._element.getparent().remove(run._element)
    align_map = {
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    for i, line in enumerate(lines):
        p = p0 if i == 0 else cell.add_paragraph()
        p.alignment = align_map[align]
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        run = p.add_run(line)
        set_run_font(run, size=size, bold=bold, font=font, color=color)
    return cell


def add_picture(cell, path, width_cm):
    tc = cell._tc
    for extra in tc.xpath("./w:p")[1:]:
        tc.remove(extra)
    p = cell.paragraphs[0]
    for run in list(p.runs):
        run._element.getparent().remove(run._element)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run()
    run.add_picture(str(path), width=Cm(width_cm))
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def para(doc, text="", size=10.5, bold=False, font="宋体", align="left",
         before=0, after=0, indent=False, underline=False, color=None):
    p = doc.add_paragraph()
    p.alignment = {
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }[align]
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = 1.08
    if indent:
        pf.first_line_indent = Pt(21)
    if text:
        run = p.add_run(text)
        set_run_font(run, size=size, bold=bold, font=font, color=color, underline=underline)
    return p


def add_runs(p, parts):
    """parts: (text, size, font, bold, underline)"""
    for text, size, font, bold, underline in parts:
        run = p.add_run(text)
        set_run_font(run, size=size, bold=bold, font=font, underline=underline)
    return p


def page_break(doc):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    run = p.add_run()
    set_run_font(run, size=1)
    run.add_break(WD_BREAK.PAGE)


def add_watermark(section, text="江苏省一体化平台"):
    """页眉里放一条浅色斜向水印，对应原件上的“江苏省一体化平台”。"""
    from lxml import etree

    v_ns = "urn:schemas-microsoft-com:vml"
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    header = section.header
    header.is_linked_to_previous = False
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run()
    pict = etree.Element("{%s}pict" % w_ns)
    shape = etree.SubElement(
        pict,
        "{%s}shape" % v_ns,
        nsmap={"v": v_ns},
    )
    shape.set("id", "Watermark")
    shape.set(
        "style",
        "position:absolute;margin-left:0;margin-top:0;width:460pt;height:90pt;"
        "rotation:315;z-index:-251658240;mso-position-horizontal:center;"
        "mso-position-horizontal-relative:margin;mso-position-vertical:center;"
        "mso-position-vertical-relative:margin",
    )
    shape.set("fillcolor", "#c8c8c8")
    shape.set("stroked", "f")
    fill = etree.SubElement(shape, "{%s}fill" % v_ns)
    fill.set("opacity", ".45")
    path = etree.SubElement(shape, "{%s}textpath" % v_ns)
    path.set("style", 'font-family:"宋体";font-size:28pt')
    path.set("string", text)
    path.set("on", "t")
    run._r.append(pict)
    # 正文页眉再放一行浅灰色文字，方便不支持 VML 的软件也能看到水印字样
    mark = p.add_run("江苏省一体化平台")
    set_run_font(mark, size=9, font="宋体", color=(210, 210, 210))


def setup_document():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Mm(15)
    section.right_margin = Mm(15)
    section.top_margin = Mm(14)
    section.bottom_margin = Mm(14)
    section.header_distance = Mm(6)
    section.footer_distance = Mm(6)
    normal = doc.styles["Normal"]
    set_style_font(normal, 10.5, "宋体")
    pf = normal.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.08
    add_watermark(section)
    core = doc.core_properties
    core.title = "专业技术资格评审申报表"
    core.subject = "胡继刚"
    core.category = "专业技术资格评审申报表"
    return doc


def scale_widths(deltas, total=PAGE_W):
    s = sum(deltas)
    return [d / s * total for d in deltas]


# ---------- 第 1 页 封面 ----------

def build_cover_page(doc):
    meta = new_table(doc, 1, [90, 90], heights=[8], sz="2")
    # 无边框
    for cell in meta.rows[0].cells:
        set_cell_borders(cell, sz="0", color="FFFFFF")
    write(meta, 0, 0, "地区代码：  320199", size=11, align="left")
    write(meta, 0, 1, "编号：  2506170651940505", size=11, align="right")

    para(doc, "专 业 技 术 资 格", size=28, font="宋体", align="center", before=28, after=6)
    para(doc, "评 审 申 报 表", size=28, font="宋体", align="center", before=6, after=26)

    fields = [
        ("单    位", "江苏省建筑装饰设计研究院有限公司"),
        ("姓    名", "胡继刚"),
        ("评审系列", "建设工程"),
        ("评审专业", "科技管理·工程管理"),
        ("技术资格", "高级工程师"),
    ]
    table = new_table(doc, len(fields), [36, 108], heights=[11] * len(fields), sz="2")
    for cell in table._cells:
        set_cell_borders(cell, sz="0", color="FFFFFF")
    for i, (label, value) in enumerate(fields):
        write(table, i, 0, label, size=14, font="楷体", align="center")
        cell = write(table, i, 1, value, size=12, align="center")
        # 下划线画在段落上，避免和下一行边框互相抵消
        p = cell.paragraphs[0]
        ppr = p._p.get_or_add_pPr()
        pbdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "12")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "000000")
        pbdr.append(bottom)
        ppr.append(pbdr)

    para(doc, "填表时间：    2025    年    08    月    12    日",
         size=14, font="楷体", align="center", before=28, after=22)
    para(doc, "江苏省专业技术人员职称(职业资格)工作领导小组办公室制",
         size=12, align="center", before=8, after=0)


# ---------- 第 2 页 ----------

def build_pledge(doc):
    para(doc, "个 人 承 诺 书", size=18, bold=True, font="黑体", align="center", before=6, after=14)
    p = para(doc, align="left", before=2, after=0)
    add_runs(p, [
        ("本人申报    ", 12, "宋体", False, False),
        ("建设工程", 12, "宋体", False, True),
        ("    系列    ", 12, "宋体", False, False),
        ("科技管理·工程管理", 12, "宋体", False, True),
        ("    专业", 12, "宋体", False, False),
    ])
    p = para(doc, align="left", before=2, after=0)
    add_runs(p, [
        ("(学科)    ", 12, "宋体", False, False),
        ("高级工程师", 12, "宋体", False, True),
        ("    资格。本人承诺：所填写的内容及所提", 12, "宋体", False, False),
    ])
    para(doc, "供的参评材料真实准确，如有任何不实或隐瞒，愿按专业技术资格评审的", size=12, before=2, after=0)
    para(doc, "有关规定接受处理。", size=12, before=2, after=16)
    para(doc, "承诺人(签字)：", size=12, align="center", before=10, after=8)
    para(doc, "日 期：    2025 年08 月12日", size=12, align="center", before=4, after=8)
    para(doc, "_" * 78, size=10, align="center", before=8, after=12)
    para(doc, "填 表 说 明", size=18, bold=True, font="黑体", align="center", before=10, after=12)
    para(doc, "1、本表供评审各级专业技术资格使用,一式2份。", size=12, before=2, after=4)
    para(doc, "2、填表内容应真实、准确、具体，并按表页下“注”的要求填写。", size=12, before=2, after=0)
    para(doc, "表内填写不下时，可另加附页，并装订入内。", size=12, before=0, after=4)
    para(doc, "3、本表请从系统中打印，一律用A3纸小册子方式打印，对折后用骑", size=12, before=2, after=0)
    para(doc, "马钉方式装订，不得放大或缩小。", size=12, before=0, after=0)


# ---------- 第 3 页 基本情况 ----------

def build_basic(doc):
    para(doc, "基  本  情  况", size=18, bold=True, font="黑体", align="center", before=0, after=6)
    deltas = [36.8, 43.4, 34.6, 16.4, 17.3, 17.3, 16.4, 34.6, 16.4, 34.6, 16.4, 34.6, 16.4, 17.3, 17.3, 16.4, 85.6]
    widths = scale_widths(deltas)
    heights = [8.2, 8.2, 8.2, 8.6, 8.4, 8.6, 9.2, 6.2, 7.2, 7.2, 7.4, 6.0, 7.2, 7.2, 7.2, 6.4]
    table = new_table(doc, 16, widths, heights=heights, sz="6")

    # 照片占右上三行
    merge(table, 0, 16, 2, 16)
    # 姓名行
    merge(table, 0, 0, 0, 1)
    merge(table, 0, 2, 0, 5)
    merge(table, 0, 6, 0, 7)
    merge(table, 0, 8, 0, 9)
    merge(table, 0, 10, 0, 11)
    merge(table, 0, 12, 0, 15)
    # 身份证行
    merge(table, 1, 0, 1, 1)
    merge(table, 1, 2, 1, 15)
    # 出生行
    merge(table, 2, 0, 2, 1)
    merge(table, 2, 2, 2, 6)
    merge(table, 2, 7, 2, 10)
    merge(table, 2, 11, 2, 15)
    # 准入
    merge(table, 3, 0, 3, 3)
    merge(table, 3, 4, 3, 11)
    merge(table, 3, 12, 3, 13)
    merge(table, 3, 14, 3, 16)
    # 现资格
    merge(table, 4, 0, 4, 2)
    merge(table, 4, 3, 4, 16)
    # 现职务
    merge(table, 5, 0, 5, 2)
    merge(table, 5, 3, 5, 12)
    merge(table, 5, 13, 5, 14)
    merge(table, 5, 15, 5, 16)
    # 现从事
    merge(table, 6, 0, 6, 2)
    merge(table, 6, 3, 6, 8)
    merge(table, 6, 9, 6, 12)
    merge(table, 6, 13, 6, 16)
    # 学历
    merge(table, 7, 0, 10, 0)
    for r in range(7, 11):
        merge(table, r, 1, r, 4)
        merge(table, r, 5, r, 8)
        merge(table, r, 9, r, 10)
        merge(table, r, 11, r, 12)
        merge(table, r, 13, r, 16)
    # 奖惩
    merge(table, 11, 0, 15, 0)
    for r in range(11, 15):
        merge(table, r, 1, r, 9)
        merge(table, r, 10, r, 12)
        merge(table, r, 13, r, 16)
    merge(table, 15, 1, 15, 1)
    merge(table, 15, 2, 15, 16)

    apply_widths(table, widths)

    write(table, 0, 0, "姓  名", size=9)
    write(table, 0, 2, "胡继刚", size=9)
    write(table, 0, 6, "性别", size=9)
    write(table, 0, 8, "男", size=9)
    write(table, 0, 10, "民族", size=9)
    write(table, 0, 12, "汉族", size=9)
    add_picture(table.cell(0, 16), ASSETS / "photo.png", 2.15)

    write(table, 1, 0, ["公民身份号码", "(社会保障号）"], size=8)
    write(table, 1, 2, "2 2 0 2 8 3 1 9 8 7 0 4 2 5 0 6 1 4", size=10)

    write(table, 2, 0, "出生年月", size=9)
    write(table, 2, 2, "1987年04月25日", size=9)
    write(table, 2, 7, "参加工作年月", size=9)
    write(table, 2, 11, "2011年06月27日", size=9)

    write(table, 3, 0, ["取得本行业(专业)准入", "资格、时间及证书号码"], size=7.5)
    write(table, 3, 4, "", size=8)
    write(table, 3, 12, ["执业", "类别"], size=8)
    write(table, 3, 14, "工程管理", size=9)

    write(table, 4, 0, ["现专业技术资格及取得", "时间、批准单位"], size=7.5)
    write(table, 4, 3, "上海市住房和城乡建设管理委员会批准取得工程师(20181225)", size=8)

    write(table, 5, 0, ["现任专业技术职务", "及聘任时间"], size=8)
    write(table, 5, 3, ["职务：工程-建设工程-工程师", "聘用时间：2018-12-25"], size=8)
    write(table, 5, 13, ["申报", "类型"], size=8)
    write(table, 5, 15, "正常申报", size=8)

    write(table, 6, 0, ["现从事专业技术工", "作及年限"], size=8)
    write(table, 6, 3, ["现工作：", "专业技术工作年限：14年"], size=7.5)
    write(table, 6, 9, ["拟评审专业技术", "资  格"], size=8)
    write(table, 6, 13, "高级工程师", size=9)

    write(table, 7, 0, stack("学历情况"), size=8)
    write(table, 7, 1, "学  校", size=8)
    write(table, 7, 5, "学习专业", size=8)
    write(table, 7, 9, "学历", size=8)
    write(table, 7, 11, "学位", size=8)
    write(table, 7, 13, "毕(肄、结)业年月", size=8)

    write(table, 8, 1, "复旦大学", size=8)
    write(table, 8, 5, "工商管理", size=8)
    write(table, 8, 9, "硕士研究生", size=8)
    write(table, 8, 11, "", size=8)
    write(table, 8, 13, "于2021年01月13日", size=8)

    write(table, 9, 1, "中国海洋大学青岛学院", size=7.5)
    write(table, 9, 5, "土木工程", size=8)
    write(table, 9, 9, "大学本科", size=8)
    write(table, 9, 11, "", size=8)
    write(table, 9, 13, "于2011年06月30日", size=8)

    write(table, 10, 1, "青岛市城阳一中", size=8)
    write(table, 10, 5, "理科", size=8)
    write(table, 10, 9, "高中", size=8)
    write(table, 10, 11, "其他", size=8)
    write(table, 10, 13, "于2016年07月04日毕业", size=8)

    write(table, 11, 0, stack("奖惩情况"), size=8)
    write(table, 11, 1, "荣誉称号、表彰奖励名称", size=8)
    write(table, 11, 10, "获奖时间", size=8)
    write(table, 11, 13, "授奖部门", size=8)

    write(table, 12, 1, "住房政策研究创新贡献奖", size=8)
    write(table, 12, 10, "2024-12-18", size=8)
    write(table, 12, 13, ["复旦大学住房政策中心专业委", "员会"], size=7.5)

    write(table, 13, 1, "城市更新卓越实践奖", size=8)
    write(table, 13, 10, "2024-04-26", size=8)
    write(table, 13, 13, "绿地控股集团技术委员会", size=7.5)

    write(table, 14, 1, "可持续社区营造标杆项目奖", size=8)
    write(table, 14, 10, "2019-02-24", size=8)
    write(table, 14, 13, "上海万科产品创新研究院", size=7.5)

    write(table, 15, 1, "处分", size=8)
    write(table, 15, 2, "无", size=8)

    para(doc, "说明：毕(肄、结)业情况应在栏目中注明。", size=9, before=3, after=0)


# ---------- 第 4 页 ----------

def build_training(doc):
    para(doc, "继 续 教 育 情 况", size=16, bold=True, font="黑体", align="center", before=0, after=1)
    para(doc, "(国内外培训、进修或考察情况)", size=12, align="center", before=0, after=6)
    widths = [38, 48, 46, 32, 16]
    rows = [
        ["起止时间", "专业或主要内容", "组织单位", "学习成绩", "备注"],
        ["2025年06月09日至\n2025年06月21日", "职称继 续教育专业课（工程系列）",
         "南京市人力资源和社会保障局专业技术人员管理处", "学时：240；学分：240", "无"],
        ["2025年05月05日至\n2025年06月09日", "公需课有效学时累计",
         "南京市人力资源和社会保障局专业技术人员管理处", "学时：152；学分：152", "无"],
        ["2025年05月03日至\n2025年06月18日", "知识更新工程高级研修项目",
         "南京市人力资源和社会保障局专业技术人员管理处", "学时：17；学分：16.5", "无"],
    ]
    table = new_table(doc, 4, widths, heights=[8, 14, 14, 14], sz="6")
    for r, row in enumerate(rows):
        for c, text in enumerate(row):
            write(table, r, c, text, size=8, align="center", valign="center")

    para(doc, "专业考试成绩情况", size=16, bold=True, font="黑体", align="center", before=12, after=6)
    widths2 = [28, 40, 48, 32, 32]
    table2 = new_table(doc, 2, widths2, heights=[8, 14], sz="6")
    headers = ["时  间", "组 织 单 位", "考试名称", "成  绩", "合格标志"]
    for c, text in enumerate(headers):
        write(table2, 0, c, text, size=9)
    write(table2, 1, 0, "", size=9)
    write(table2, 1, 1, "无", size=9)
    for c in range(2, 5):
        write(table2, 1, c, "", size=9)


# ---------- 第 5 页 ----------

def build_career(doc):
    para(doc, "工  作  经  历", size=18, bold=True, font="黑体", align="center", before=0, after=1)
    para(doc, "(专业技术工作经历)", size=12, align="center", before=0, after=4)
    widths = [36, 48, 18, 18, 36, 24]
    headers = ["起止时间", "工作单位", "部门", "岗位", "从事的主要\n专业技术工作", "任何职"]
    data = [
        ["2025年03月01日至\n2025年07月10日", "江苏省建筑装饰设计研究院有限公司", "工程部", "管理岗", "", "副总经理"],
        ["2023年06月07日至\n2025年02月28日", "上海徐汇绿地商业管理有限公司", "工程部", "管理岗", "", "副总经理"],
        ["2022年04月11日至\n2023年05月13日", "上海融之旅文化旅游发展有限公司（融创文旅华东区域）", "工程部", "管理岗", "", "总监"],
        ["2018年12月10日至\n2022年03月18日", "上海万科长宁置业有限公司", "工程部", "管理岗", "", "总监"],
        ["2017年11月06日至\n2018年11月10日", "上海万限实业有限公司 【万科（上海）实业发展有限公司】", "工程部", "管理岗", "", "总监"],
        ["2017年06月01日至\n2017年10月26日", "上海金科文化旅游发展集团有限公司", "工程部", "管理岗", "", "资深经理"],
        ["2017年03月20日至\n2017年05月29日", "金科（上海）建筑设计有限公司", "工程部", "管理岗", "", "资深经理"],
        ["2015年06月08日至\n2017年02月24日", "新城控股集团股份有限公司上海第一分公司", "工程部", "管理岗", "", "高级经理"],
    ]
    table = new_table(doc, 1 + len(data), widths, heights=[10] + [13] * len(data), sz="6")
    for c, text in enumerate(headers):
        write(table, 0, c, text, size=8)
    for r, row in enumerate(data, start=1):
        for c, text in enumerate(row):
            write(table, r, c, text, size=8)


# ---------- 第 6 页 ----------

def build_performance(doc):
    para(doc, "任现职后主要专业技术工作业绩登记", size=16, bold=True, font="黑体", align="center", before=0, after=6)
    widths = [32, 52, 74, 22]
    headers = ["起止时间", "专业技术工作\n项目、课题名称", "获奖励、专利\n及效益情况", "本人起\n何作用"]
    data = [
        [
            "2018年02月03日至2023年05月11日",
            "专业技术工作项目：杭州湾融创文旅城一水、陆乐园项目\n课题名称：杭州湾融创文旅城",
            "获奖情况：无\n专利情况：无\n效益情况：无",
            "主持",
        ],
        [
            "2019年11月15日至2022年03月25日",
            "专业技术工作项目：万科青浦区徐泾镇徐盈路西侧地块（01-06F2-04 地块一、二期）\n课题名称：万科天空之城",
            "获奖情况：荣获上海万科产品创新研究院颁发的“可持续社区营造标杆项目奖”\n"
            "专利情况：专利号：ZL 2020 2 0536480.4 专利号：ZL 2020 2 0557884.1\n"
            "效益情况：在上海万科天空之城项目中，主导完成以下创新实践： 1. 构建社区级\"零碳\"能源系统，光伏覆盖率达32%； "
            "2. 研发模块化雨水回收装置，节水效率提升45%； 3. 搭建社区参与式设计平台，客户满意度达92%。",
            "主持",
        ],
        [
            "2020年10月15日至2024年01月25日",
            "专业技术工作项目：绿地苏州东太湖旅游度假区WJ-J-2020-013 地块商住\n课题名称：绿地苏州东太湖旅游度假区",
            "获奖情况：荣获绿地控股集团技术委员会颁发的“城市更新卓越实践奖”\n专利情况：无\n"
            "效益情况：1. 技术体系创新：构建全流程技术标准体系，集成绿色建材与智能建造技术，单位面积改造成本降低18%，获集团技术创新标杆认证。 "
            "2. 智慧工地研发：开发研发智慧工地管理系统思路，通过物联网与AI算法优化施工调度，降低人力成本15%、材料损耗率23%，并推动BIM技术应用达LOD500标准。 "
            "3. 经济效益：累计节约项目成本2,800万元，投资回报率提升32%；",
            "主持",
        ],
    ]
    table = new_table(doc, 4, widths, heights=[10, 22, 42, 48], sz="6")
    for c, text in enumerate(headers):
        write(table, 0, c, text, size=8)
    for r, row in enumerate(data, start=1):
        write(table, r, 0, row[0], size=8, align="center")
        write(table, r, 1, row[1], size=8, align="left", valign="center")
        write(table, r, 2, row[2], size=7.5, align="left", valign="center")
        write(table, r, 3, row[3], size=9, align="center")
    para(doc, "说明：“本人起何作用”分为主持、参加、独立承担。", size=9, before=3, after=0)


# ---------- 第 7 页 ----------

def build_papers(doc):
    para(doc, "任现职后著作、论文及重要技术报告登记", size=15, bold=True, font="黑体", align="center", before=0, after=6)
    widths = [28, 52, 42, 36, 22]
    headers = ["日  期", "论著(报告)标题", "刊物、出版单位\n(学术会议)名称", "主办单位", "本人承担情况"]
    data = [
        ["2021-07-08", "铝膜板在高层住宅建筑工程中的应用",
         "中国民营科技促进会，《中国新技术新产品》", "中国民营科技促进会", "合著"],
        ["2021-08-05", "房屋建筑施工中人工智能技术的应用",
         "深圳住宅与房地产杂志社;中国房地产及住宅研究会，《住宅与房地产》",
         "深圳住宅与房地产杂志社;中国房地产及住宅研究会", "合著"],
        ["2024-07-10", "基于BIM的装配式钢结构建筑施工新技术研究",
         "中国建筑金属结构协会，《中国建筑金属结构》", "中国建筑金属结构协会", "独著"],
    ]
    table = new_table(doc, 4, widths, heights=[12, 16, 18, 16], sz="6")
    for c, text in enumerate(headers):
        write(table, 0, c, text, size=8)
    for r, row in enumerate(data, start=1):
        for c, text in enumerate(row):
            align = "left" if c in (1, 2, 3) else "center"
            write(table, r, c, text, size=8, align=align)
    para(doc, "说明：本人承担情况按独著(译)、合著(译)填写。合著(译)应注明本人实际承担的部分。",
         size=9, before=3, after=0)


# ---------- 第 8–11 页 工作总结 ----------

def build_summary(doc):
    para(doc, "本人任现职以来工作总结", size=16, bold=True, font="黑体", align="center", before=0, after=2)
    para(doc, "(包括思想品德、职业道德、专业技术能力、工作成绩及履行职责等，乡土人才可提供省三带情况)",
         size=9, align="center", before=0, after=8)
    para(doc, "个人专业技术工作总结", size=14, align="center", before=4, after=8)
    para(doc, "本文总结我在建筑施工领域的卓越业绩，参与重大项目，解决技术难题，展现专业实力。",
         size=10.5, indent=True, before=1, after=1)
    para(doc, "提出前瞻性建议，分享学术成果，获同行认可。发表多篇论文，申请专利，为行业发展贡献智慧。以严谨态度与进取精神，积极贡献于行业。",
         size=10.5, indent=True, before=0, after=6)
    para(doc, "一、工作经历与学术背景", size=12, bold=True, before=4, after=3)
    para(doc, "1.教育背景", size=10.5, bold=True, before=2, after=2)
    para(doc, "本人于2011年6月从中国海洋大学青岛学院土木工程专业毕业，于2021年1月，获得复旦大学工商管理硕士专业学历，进一步丰富了我的知识体系，为我在建筑和金融领域的交叉发展奠定了坚实基础。",
         size=10.5, indent=True, before=0, after=4)
    para(doc, "2.职业履历", size=10.5, bold=True, before=2, after=3)

    widths = [32, 100, 48]
    headers = ["时间", "公司名称", "职务"]
    data = [
        ["2025.03-至今", "江苏省建筑装饰设计研究院有限公司", "工程管理副总经理"],
        ["2023.06-2025.02", "上海徐汇绿地商业管理有限公司", "工程部副总经理"],
        ["2022.04-2023.05", "上海融之旅文化旅游发展有限公司（融创文旅华东区域）", "工程总监"],
        ["2017.11-2022.03", "上海万科企业有限公司", "工程部总监"],
        ["2017.03-2017.10", "金科地产集团有限公司华东区域公司", "工程总监"],
        ["2015.06-2017.02", "新城控股集团上海第一分公司", "工程部高级经理"],
        ["2011.07-2015.05", "江苏中南建设集团股份有限公司", "施工员/资料员"],
    ]
    table = new_table(doc, 8, widths, heights=[7] + [7.2] * 7, header_rows={0}, sz="6")
    for c, text in enumerate(headers):
        write(table, 0, c, text, size=10.5, bold=True)
    for r, row in enumerate(data, start=1):
        for c, text in enumerate(row):
            write(table, r, c, text, size=10.5)

    page_break(doc)
    para(doc, "3.社会职务", size=10.5, bold=True, before=2, after=2)
    para(doc, "复旦大学住房政策研究中心秘书长（研究员）", size=10.5, before=1, after=1)
    para(doc, "上海市杨浦区科技企业联合会执行会长", size=10.5, before=1, after=3)
    para(doc, "以上丰富的工作经历与学术背景使我具备了深厚的专业素养和广泛的行业影响力，我将继续以敬业的态度和专业的精神，为建筑和金融领域的发展贡献自己的力量。",
         size=10.5, indent=True, before=2, after=6)
    para(doc, "二、核心项目业绩与技术突破", size=12, bold=True, before=4, after=4)

    blocks = [
        ("1. 中南海湾新城·金石国际广场（2011.07-2015.08）", True),
        ("项目属性：超高层五星级酒店公寓（26万㎡，120m）", False),
        ("角色：施工技术管理", False),
        ("技术突破：", False),
        ("▶ 攻克核心筒框架剪力墙施工难题，通过三维建模预演优化节点处理方案", False),
        ("▶ 建立实时沉降监测系统，控制结构偏移误差＜3mm", False),
        ("成果：节约工期45天，获江苏省优质工程奖", False),
        ("2. 新城香溢紫郡（2015.12-2017.02）", True),
        ("项目属性：高端住宅（23万㎡）", False),
        ("角色：工程高级经理", False),
        ("技术突破：", False),
        ("▶ 创新桩基切顶高程控制法，降低地基处理成本23%", False),
        ("▶ 研发混凝土裂缝预警系统，楼板开裂率下降至0.8‰", False),
        ("成果：项目提前3个月交付，利润率提升12%", False),
        ("3、合肥金科博翠天辰2017.03-2017.10", True),
        ("项目属性：高端住宅综合体（总建面约34万㎡，容积率2.5）", False),
        ("角色：项目技术负责人（隶属金科集团华东区域公司）", False),
        ("技术突破与创新", False),
        ("突破一：超高层住宅剪力墙结构抗震优化", False),
        ("•解决方案与个人贡献：", False),
        ("本人行动：主导结构模型风洞试验，优化型钢排布方案，使结构刚度分布均匀化，减少扭转效应18%。", False),
        ("采用BIM预拼装技术解决型钢与钢筋密集碰撞问题，缩短关键路径工期23天。", False),
        ("突破二：高端住宅绿色科技系统集成", False),
        ("• 解决方案与个人贡献：", False),
        ("本人行动：优化管网穿剪力墙节点密封工艺，解决冷桥结露问题，系统能效比提升至4.8（国标≥3.8）。", False),
        ("▪ 嵌入金科“Life+云平台”，实现户内能耗智能调节，项目获合肥市绿色建筑示范工程", False),
    ]
    for text, head in blocks:
        para(doc, text, size=10.5, bold=head, before=1, after=1)

    page_break(doc)
    para(doc, "认证。", size=10.5, before=0, after=3)
    more = [
        ("4.上海万科天空之城（2017.11-2022.02）", True),
        ("l项目属性：商业综合体+住宅（42万㎡）", False),
        ("l角色：工程总监", False),
        ("l技术突破：", False),
        ("▶ 首创商业-住宅协同施工模型，解决交叉作业冲突", False),
        ("▶ 应用自愈合混凝土技术于剪力墙节点", False),
        ("l成果：获LEED金级认证，商业区开业首年客流超预期40%", False),
        ("5.杭州湾融创文旅城水路乐园项目（海盐）（2022.04-2023.05）", True),
        ("l项目属性：高端抗震住宅", False),
        ("l角色：工程总监", False),
        ("l技术突破：", False),
        ("▶ 优化框架剪力墙抗震参数，地震模拟性能提升35%", False),
        ("▶ 实施无人机巡检系统，质量验收效率提升60%", False),
        ("6.苏州绿地江南映月花园（2023.06-2025.03）", True),
        ("l项目属性：现状高端住宅项目", False),
        ("l角色：工程部副总经理", False),
        ("l技术突破：", False),
        ("▶ 应用BIM+3D打印技术实现异形钢结构精准安装", False),
        ("▶ 开发碳足迹监测平台，施工碳排放降低28%", False),
        ("6、项目名称：东台市养老服务中心地下车库装饰工程（2025.3-至今）", True),
        ("项目属性：适老化公共空间改造", False),
        ("角色：施工技术负责人", False),
        ("技术突破与创新", False),
        ("1. 适老化地下空间无障碍导视系统", False),
        ("•问题本质：老年群体空间辨识能力弱，传统车库标识系统易引发迷失焦虑。", False),
        ("• 解决方案与个人贡献：", False),
        ("▶ 地坪发光标线+AI语音联动系统：", False),
        ("▶ BIM模拟老人动线：", False),
        ("2. BIM-无人机协同进度管控", False),
        ("▶ 建立无人机实景扫描与BIM模型自动比对系统，偏差超5%自动预警；", False),
        ("▶ 成果：减少返工损失150万元（占预算8.3%），关键节点“零延误”。", False),
    ]
    for text, head in more:
        para(doc, text, size=10.5, bold=head, before=1, after=1)

    page_break(doc)
    para(doc, "三、技术创新与学术贡献", size=12, bold=True, before=2, after=3)
    para(doc, "1.知识产权", size=10.5, bold=True, before=2, after=2)
    para(doc, "a发明专利《建筑工程管理用功能脚架》（ZL202020536480.4），解决高空作业安全防护难题，已在12个项目推广；",
         size=10.5, before=1, after=1)
    para(doc, "b实用新型专利《建筑工程防护围栏》（ZL202020557884.1），模块化设计使安装效率提升200%",
         size=10.5, before=1, after=3)
    para(doc, "2.学术论著", size=10.5, bold=True, before=2, after=3)

    widths = [24, 78, 58, 20]
    table = new_table(doc, 4, widths, heights=[7, 12, 12, 12], header_rows={0}, sz="6")
    for c, text in enumerate(["时间", "论文标题", "期刊", "色"]):
        write(table, 0, c, text, size=10.5, bold=True)
    papers = [
        ["2021.07", "《铝模板在高层住宅建筑工程中的应用》", "《中国新技术新产品》", "作"],
        ["2021.08", "《人工智能技术在房屋建筑工程中的应用》", "《住宅与房地产》", "作"],
        ["2024.07", "《BIM装配式钢结构建筑施工技术与管理研究》", "《中国建筑金属结构》", "作"],
    ]
    for r, row in enumerate(papers, start=1):
        for c, text in enumerate(row):
            write(table, r, c, text, size=10.5, align="center" if c in (0, 3) else "left")

    para(doc, "2.学术论著", size=10.5, bold=True, before=8, after=2)
    para(doc, "学术交流", size=10.5, bold=True, before=2, after=2)
    para(doc, "a2022.07 上海市建筑施工技术交流会", size=10.5, before=1, after=0)
    para(doc, "主题演讲：《绿色建筑施工技术创新路径》", size=10.5, before=0, after=2)
    para(doc, "b2022.11 江苏省建筑技术研讨会", size=10.5, before=1, after=0)
    para(doc, "发表演讲：《3D打印技术在预制建筑中的实践突破》", size=10.5, before=0, after=4)
    para(doc, "四、专业能力总结", size=12, bold=True, before=4, after=3)
    for line in [
        "1.技术攻坚：在超高层结构、抗震优化、绿色建造等领域拥有9项核心技术",
        "2.管理创新：独创“动态节点管理法”实现多业态项目协同开发",
        "3.行业贡献：推动BIM技术与ESG理念在长三角项目落地，累计节约工程造价超1.2亿元",
        "4.发展愿景：聚焦智能建造与低碳技术融合，助力建筑业数字化转型",
    ]:
        para(doc, line, size=10.5, before=1, after=1)
    para(doc, "本人签名 _____________________", size=12, align="right", before=16, after=4)
    para(doc, "年      月      日", size=12, align="right", before=2, after=0)


# ---------- 第 12 页 ----------

def build_assessment(doc):
    para(doc, "任现职以来考核情况", size=18, bold=True, font="黑体", align="center", before=0, after=8)
    widths = [16, 24, 62, 28, 32, 18]
    table = new_table(doc, 4, widths, heights=[12, 14, 14, 14], sz="6")
    merge(table, 0, 0, 3, 0)
    apply_widths(table, widths)
    write(table, 0, 0, stack("年度（任期）考核情况"), size=8)
    for c, text in enumerate(["年  份", "受聘专业技术职务(岗位)", "考核等次", "考核单位", "备注"], start=1):
        write(table, 0, c, text, size=8)
    write(table, 1, 2, "无", size=10)
    for r in range(1, 4):
        for c in range(1, 6):
            if not (r == 1 and c == 2):
                write(table, r, c, "", size=9)


# ---------- 第 13 页 ----------

def build_org(doc):
    para(doc, "工作单位或行业协会（学会）或人事档案管理单位核实情况",
         size=13, bold=True, font="黑体", align="center", before=0, after=6)
    widths = [28, 76, 76]
    table = new_table(doc, 3, widths, heights=[16, 12, 10], sz="6")
    merge(table, 0, 1, 0, 2)
    apply_widths(table, widths)
    # 侧栏两列竖排，与原件左右两列读法一致
    side = table.cell(0, 0)
    # 清掉默认段落后放一个无边框内表
    write(table, 0, 0, "", size=8)
    inner = side.add_table(rows=1, cols=2)
    iw = [13, 13]
    apply_widths(inner, iw)
    for cell in inner.rows[0].cells:
        set_cell_borders(cell, sz="0", color="FFFFFF")
        set_cell_margins(cell, 20, 20, 20, 20)
    write(inner, 0, 0, stack("单位或行业协会（学会）或"), size=8)
    write(inner, 0, 1, stack("人事档案管理单位核实情况"), size=8)
    # 内表会落在文字段落后，把空段高度压掉
    side.paragraphs[0].paragraph_format.space_before = Pt(0)
    side.paragraphs[0].paragraph_format.space_after = Pt(0)

    write(table, 0, 1, "经核查，申报人所填内容及提交的材料属实。", size=14, bold=True, font="黑体")
    write(table, 1, 0, "", size=8)
    write(table, 1, 1, "核 实 人(签字)", size=10, align="center")
    write(table, 1, 2, "单位(印章)", size=10, align="center")
    write(table, 2, 0, "", size=8)
    write(table, 2, 1, "日  期", size=10)
    write(table, 2, 2, "日  期", size=10)
    # 左侧竖排跨三行
    merge(table, 0, 0, 2, 0)
    apply_widths(table, widths)

    para(doc, "组  织  意  见", size=16, bold=True, font="黑体", align="center", before=14, after=6)
    widths2 = [18, 72, 18, 72]
    table2 = new_table(doc, 1, widths2, heights=[78], sz="6")
    write(table2, 0, 0, stack("县（市）﹑区职称部门或市主管部门意见"), size=8)
    write(table2, 0, 2, stack("市﹑厅（局）人事职称部门意见"), size=8)

    left = table2.cell(0, 1)
    write(table2, 0, 1, "", size=8)
    left.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    gray = (128, 128, 128)
    lines_left = [
        (["区县行政主管章", "区县行业主管章"], 8),
        (["区县职称办盖章", "市行政主管盖章"], 8),
        (["申报点盖章"], 8),
        (["公  章"], 11),
        (["年      月      日"], 11),
    ]
    # 用已有空段开始写
    p = left.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # 上面 write 已经写了空 run，清掉
    for run in list(p.runs):
        run._element.getparent().remove(run._element)
    first = True
    for texts, size in lines_left:
        for t in texts:
            if first:
                pp = p
                first = False
            else:
                pp = left.add_paragraph()
            pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pp.paragraph_format.space_before = Pt(2)
            pp.paragraph_format.space_after = Pt(2)
            color = None if t in ("公  章", "年      月      日") else gray
            font = "宋体"
            run = pp.add_run(t)
            set_run_font(run, size=size, font=font, color=color)

    right = table2.cell(0, 3)
    write(table2, 0, 3, "", size=8)
    right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p = right.paragraphs[0]
    for run in list(p.runs):
        run._element.getparent().remove(run._element)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("市职称办盖章")
    set_run_font(run, size=8, color=gray)
    p2 = right.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(2)
    p2.paragraph_format.space_after = Pt(0)
    run = p2.add_run()
    run.add_picture(str(ASSETS / "seal_hr.png"), width=Cm(3.1))
    p3 = right.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_before = Pt(1)
    p3.paragraph_format.space_after = Pt(0)
    run = p3.add_run("负责人(签字)          公  章")
    set_run_font(run, size=10)
    p4 = right.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.paragraph_format.space_before = Pt(2)
    p4.paragraph_format.space_after = Pt(0)
    run = p4.add_run("2025 年08 月12 日")
    set_run_font(run, size=11)


# ---------- 第 14 页 ----------

def build_review(doc):
    para(doc, "评审、登记备案情况", size=16, bold=True, font="黑体", align="center", before=0, after=4)

    # 用 11 列方便票数小格和备注两列对齐
    # 0 侧栏 | 1-10 内容
    deltas = [16, 22, 12, 22, 12, 18, 12, 18, 12, 16, 12]
    widths = scale_widths(deltas, PAGE_W)
    # 行：票数、签字、票数、意见、主任签字、备案、备注三行
    table = new_table(doc, 9, widths, heights=[8, 18, 8, 22, 28, 32, 6.5, 6.5, 6.5], sz="6")

    # 评议组侧栏
    merge(table, 0, 0, 1, 0)
    # 第一行票数：标签+数字交替，从 col 1
    # 应到人数 | 3 | 实到人数 | 3 | 赞成票 | 3 | 反对票 | 0 | 弃权票 | 0
    # 10 content cols: we'll merge some so five pairs fit.
    # cols 1-2 应到人数, 3 数字, 4-5 实到人数, 6 数字, 7 赞成票, 8 数字, 9 反对票? not enough.
    # Redo with a nested approach: side cell + one content cell containing inner tables.
    # The outer 11-col plan is awkward. Rebuild this page with 2-col outer tables stacked.

    # 放弃当前表，改用分段表。先把刚加的空表删掉。
    tbl = table._tbl
    tbl.getparent().remove(tbl)
    # 删掉表前可能多余的空段？ add_table 不额外加段。直接重建。

    def side_block(title, inner_builder, side_w=16, height=28):
        widths_b = [side_w, PAGE_W - side_w]
        block = new_table(doc, 1, widths_b, heights=[height], sz="6")
        write(block, 0, 0, stack(title), size=8)
        cell = block.cell(0, 1)
        write(block, 0, 1, "", size=8)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        # 去掉空段文字
        p = cell.paragraphs[0]
        for run in list(p.runs):
            run._element.getparent().remove(run._element)
        inner_builder(cell)
        return block

    def votes(parent, pairs):
        # pairs: [(label, value), ...]
        p = parent.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        # 用内表
        inner = parent.add_table(rows=1, cols=len(pairs) * 2)
        each = (PAGE_W - 16 - 2) / (len(pairs) * 2)
        # 数字格窄一点
        ws = []
        for _ in pairs:
            ws.extend([each * 1.35, each * 0.65])
        # 归一
        scale = (PAGE_W - 18) / sum(ws)
        ws = [w * scale for w in ws]
        apply_widths(inner, ws)
        for cell in inner.rows[0].cells:
            set_cell_borders(cell, sz="6")
            set_cell_margins(cell, 30, 30, 30, 30)
        for i, (label, value) in enumerate(pairs):
            write(inner, 0, i * 2, label, size=8)
            write(inner, 0, i * 2 + 1, value, size=9)

    def group_inner(cell):
        votes(cell, [("应到人数", "3"), ("实到人数", "3"), ("赞成票", "3"), ("反对票", "0"), ("弃权票", "0")])
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run("评议组组长(签字)")
        set_run_font(run, size=10)
        run = p.add_run("    ")
        set_run_font(run, size=10)
        run = p.add_run()
        run.add_picture(str(ASSETS / "sign_group.png"), width=Cm(2.6))
        run = p.add_run("          2025 年09 月21 日")
        set_run_font(run, size=11)

    side_block("专业（学科）评议组意见", group_inner, height=32)

    def committee_inner(cell):
        votes(cell, [("应到人数", "27"), ("实到人数", "27"), ("赞成票", "27"), ("反对票", "0"), ("弃权票", "0")])
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run("经评审，胡继刚（220283198704250614） 已 具备 建设工程·科技管理·工程管理 专业（学科）高级工程师 资格。")
        set_run_font(run, size=11)
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p2.paragraph_format.space_before = Pt(4)
        p2.paragraph_format.space_after = Pt(0)
        run = p2.add_run("主任委员(签字)  ")
        set_run_font(run, size=10)
        run = p2.add_run()
        run.add_picture(str(ASSETS / "sign_chair.png"), width=Cm(2.4))
        run = p2.add_run("   ")
        set_run_font(run, size=10)
        run = p2.add_run("评委会盖章")
        set_run_font(run, size=9, color=(80, 80, 80))
        run = p2.add_run()
        run.add_picture(str(ASSETS / "seal_committee.png"), width=Cm(2.5))
        run = p2.add_run("  公  章")
        set_run_font(run, size=10)
        p3 = cell.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p3.paragraph_format.space_before = Pt(1)
        p3.paragraph_format.space_after = Pt(0)
        run = p3.add_run("2025 年10 月19 日")
        set_run_font(run, size=11)

    side_block("评审委员会评审意见", committee_inner, height=48)

    def record_inner(cell):
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        run = p.add_run("签发单位盖章")
        set_run_font(run, size=9, color=(80, 80, 80))
        run = p.add_run("                    ")
        set_run_font(run, size=9)
        run = p.add_run("备案部门盖章")
        set_run_font(run, size=9, color=(80, 80, 80))
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        run = p2.add_run()
        run.add_picture(str(ASSETS / "seal_issue.png"), width=Cm(2.7))
        run = p2.add_run("     ")
        set_run_font(run, size=10)
        run = p2.add_run()
        run.add_picture(str(ASSETS / "seal_archive.png"), width=Cm(2.7))
        p3 = cell.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p3.paragraph_format.space_before = Pt(0)
        run = p3.add_run("公  章")
        set_run_font(run, size=10)
        p4 = cell.add_paragraph()
        p4.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p4.paragraph_format.space_before = Pt(1)
        run = p4.add_run("2025 年12 月09 日")
        set_run_font(run, size=11)

    side_block("省、市、县或相应职称主管部门登记备案情况", record_inner, height=42)

    # 备注
    widths_n = [16, 82, 82]
    note = new_table(doc, 3, widths_n, heights=[6.2, 6.2, 6.2], sz="6")
    merge(note, 0, 0, 2, 0)
    apply_widths(note, widths_n)
    write(note, 0, 0, stack("备注"), size=9)
    write(note, 0, 1, "人员姓名：胡继刚", size=9, align="left")
    write(note, 0, 2, "证件号码：220283198704250614", size=9, align="left")
    write(note, 1, 1, "联系电话：18678408669", size=9, align="left")
    write(note, 1, 2, "流程编号：2506170651940505", size=9, align="left")
    merge(note, 2, 1, 2, 2)
    apply_widths(note, widths_n)
    write(note, 2, 1, "工作单位：江苏省建筑装饰设计研究院有限公司", size=9, align="left")

    para(doc, "说明：“评审委员会评审意见”栏，应在“具备”前写明“已”或“不”。", size=9, before=3, after=0)


def build():
    doc = setup_document()
    build_cover_page(doc)
    page_break(doc)
    build_pledge(doc)
    page_break(doc)
    build_basic(doc)
    page_break(doc)
    build_training(doc)
    page_break(doc)
    build_career(doc)
    page_break(doc)
    build_performance(doc)
    page_break(doc)
    build_papers(doc)
    page_break(doc)
    build_summary(doc)
    page_break(doc)
    build_assessment(doc)
    page_break(doc)
    build_org(doc)
    page_break(doc)
    build_review(doc)
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT)
    return OUTPUT


# 必须出现在 Word 里的原文片段。用来防止漏字。
REQUIRED = [
    "320199",
    "2506170651940505",
    "专 业 技 术 资 格",
    "评 审 申 报 表",
    "江苏省建筑装饰设计研究院有限公司",
    "胡继刚",
    "建设工程",
    "科技管理·工程管理",
    "高级工程师",
    "江苏省专业技术人员职称(职业资格)工作领导小组办公室制",
    "个 人 承 诺 书",
    "所填写的内容及所提供的参评材料真实准确",
    "填 表 说 明",
    "一式2份",
    "不得放大或缩小",
    "2 2 0 2 8 3 1 9 8 7 0 4 2 5 0 6 1 4",
    "1987年04月25日",
    "2011年06月27日",
    "工程管理",
    "上海市住房和城乡建设管理委员会批准取得工程师(20181225)",
    "工程-建设工程-工程师",
    "2018-12-25",
    "正常申报",
    "专业技术工作年限：14年",
    "复旦大学",
    "工商管理",
    "硕士研究生",
    "于2021年01月13日",
    "中国海洋大学青岛学院",
    "土木工程",
    "大学本科",
    "于2011年06月30日",
    "青岛市城阳一中",
    "于2016年07月04日毕业",
    "住房政策研究创新贡献奖",
    "2024-12-18",
    "复旦大学住房政策中心专业委员会",
    "城市更新卓越实践奖",
    "2024-04-26",
    "绿地控股集团技术委员会",
    "可持续社区营造标杆项目奖",
    "2019-02-24",
    "上海万科产品创新研究院",
    "职称继 续教育专业课（工程系列）",
    "学时：240；学分：240",
    "公需课有效学时累计",
    "学时：152；学分：152",
    "知识更新工程高级研修项目",
    "学时：17；学分：16.5",
    "南京市人力资源和社会保障局专业技术人员管理处",
    "上海徐汇绿地商业管理有限公司",
    "上海融之旅文化旅游发展有限公司（融创文旅华东区域）",
    "上海万科长宁置业有限公司",
    "上海万限实业有限公司 【万科（上海）实业发展有限公司】",
    "上海金科文化旅游发展集团有限公司",
    "金科（上海）建筑设计有限公司",
    "新城控股集团股份有限公司上海第一分公司",
    "副总经理",
    "资深经理",
    "杭州湾融创文旅城一水、陆乐园项目",
    "ZL 2020 2 0536480.4",
    "ZL 2020 2 0557884.1",
    "光伏覆盖率达32%",
    "节水效率提升45%",
    "客户满意度达92%",
    "WJ-J-2020-013",
    "单位面积改造成本降低18%",
    "材料损耗率23%",
    "LOD500",
    "2,800万元",
    "投资回报率提升32%",
    "铝膜板在高层住宅建筑工程中的应用",
    "房屋建筑施工中人工智能技术的应用",
    "基于BIM的装配式钢结构建筑施工新技术研究",
    "中国新技术新产品",
    "住宅与房地产",
    "中国建筑金属结构",
    "合著",
    "独著",
    "2021-07-08",
    "2021-08-05",
    "2024-07-10",
    "个人专业技术工作总结",
    "中国海洋大学青岛学院土木工程专业毕业",
    "复旦大学住房政策研究中心秘书长（研究员）",
    "上海市杨浦区科技企业联合会执行会长",
    "中南海湾新城·金石国际广场",
    "新城香溢紫郡",
    "合肥金科博翠天辰",
    "扭转效应18%",
    "Life+云平台",
    "上海万科天空之城",
    "LEED金级认证",
    "杭州湾融创文旅城水路乐园项目（海盐）",
    "苏州绿地江南映月花园",
    "东台市养老服务中心地下车库装饰工程",
    "零延误",
    "建筑工程管理用功能脚架",
    "ZL202020536480.4",
    "建筑工程防护围栏",
    "ZL202020557884.1",
    "《铝模板在高层住宅建筑工程中的应用》",
    "《人工智能技术在房屋建筑工程中的应用》",
    "《BIM装配式钢结构建筑施工技术与管理研究》",
    "绿色建筑施工技术创新路径",
    "3D打印技术在预制建筑中的实践突破",
    "动态节点管理法",
    "1.2亿元",
    "经核查，申报人所填内容及提交的材料属实。",
    "2025 年08 月12 日",
    "220283198704250614",
    "18678408669",
    "评委会盖章",
    "已 具备",
    "2025 年09 月21 日",
    "2025 年10 月19 日",
    "2025 年12 月09 日",
    "应在“具备”前写明“已”或“不”",
]


def document_text(path: Path) -> str:
    doc = Document(str(path))
    chunks = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                chunks.append(cell.text)
    return "\n".join(chunks)


def check(path: Path) -> list[str]:
    text = document_text(path).replace("\n", "")
    missing = []
    for item in REQUIRED:
        if item.replace("\n", "") not in text:
            missing.append(item)
    return missing


def main():
    path = build()
    missing = check(path)
    print(f"已生成 {path}")
    if missing:
        print("缺少以下原文：")
        for item in missing:
            print(" -", item)
        sys.exit(1)
    print(f"原文核对通过，共 {len(REQUIRED)} 条。")


if __name__ == "__main__":
    main()
