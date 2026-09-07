#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成胡继刚优化版简历（Word + 可预览 HTML）。

主线：国资城市更新 / 租赁住房 / 存量收并购。
不写入身份证、住址、配偶等隐私；不把未核验职务写成现任；不编造业绩数字。
一线投资只计 2011–2021；复旦 MBA 标明在职攻读。
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output"

NAVY = RGBColor(0x14, 0x2B, 0x4A)
NAVY_HEX = "142B4A"
GOLD = RGBColor(0xC4, 0x8A, 0x2B)
GOLD_HEX = "C48A2B"
GRAY = RGBColor(0x2C, 0x33, 0x3A)
MUTED = RGBColor(0x5C, 0x66, 0x70)
LINE = "D5DDE8"
SOFT = "F3F6FA"


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


def clear_cell(cell):
    cell.text = ""
    for p in cell.paragraphs[1:]:
        p._element.getparent().remove(p._element)


def para_in(cell, text, *, size=10, bold=False, color=None, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=0):
    p = cell.paragraphs[0] if not cell.paragraphs[0].text and not cell.paragraphs[0].runs else cell.add_paragraph()
    if cell.paragraphs[0].text == "" and not cell.paragraphs[0].runs:
        p = cell.paragraphs[0]
    p.alignment = align
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.line_spacing = 1.08
    run = p.add_run(text)
    set_run_font(run, name="微软雅黑", size=size, bold=bold, color=color, east_asia="微软雅黑")
    return p


def add_p(doc, text, *, size=10.5, bold=False, color=None, space_before=0, space_after=4, align=None, line=1.12):
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
    p.paragraph_format.space_after = Pt(4)
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
    add_p(doc, text, size=12, bold=True, color=NAVY, space_before=8, space_after=1)
    add_hr(doc, GOLD_HEX, "14")


def bullet(doc, text, *, size=9.5):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.32)
    pf.first_line_indent = Cm(-0.32)
    pf.space_before = Pt(0)
    pf.space_after = Pt(1.5)
    pf.line_spacing = 1.12
    run = p.add_run("•  " + text)
    set_run_font(run, size=size, color=GRAY, east_asia="微软雅黑")
    return p


def set_col_widths(table, widths_cm):
    table.autofit = False
    table.allow_autofit = False
    for row in table.rows:
        for i, w in enumerate(widths_cm):
            if i < len(row.cells):
                row.cells[i].width = Cm(w)


def job_header(doc, company, title, dates):
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(table, [12.6, 5.0])
    left, right = table.rows[0].cells
    clear_cell(left)
    clear_cell(right)
    set_cell_borders(left)
    set_cell_borders(right)
    set_cell_margin(left, 20, 20, 0, 40)
    set_cell_margin(right, 20, 20, 40, 0)
    p = left.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    r1 = p.add_run(company)
    set_run_font(r1, size=10.5, bold=True, color=NAVY, east_asia="微软雅黑")
    r2 = p.add_run("  |  " + title)
    set_run_font(r2, size=9.5, bold=False, color=MUTED, east_asia="微软雅黑")
    p2 = right.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.paragraph_format.space_after = Pt(0)
    r3 = p2.add_run(dates)
    set_run_font(r3, size=9.5, bold=True, color=GOLD, east_asia="微软雅黑")
    return table


def build_docx() -> Path:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.15)
    section.bottom_margin = Cm(1.15)
    section.left_margin = Cm(1.4)
    section.right_margin = Cm(1.4)

    head = doc.add_table(rows=1, cols=1)
    head.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(head, [18.2])
    cell = head.rows[0].cells[0]
    clear_cell(cell)
    shade_cell(cell, NAVY_HEX)
    set_cell_borders(cell, NAVY_HEX, "0")
    set_cell_margin(cell, 90, 90, 140, 140)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("胡继刚")
    set_run_font(r, size=22, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), east_asia="微软雅黑")
    r2 = p.add_run("    HU Jigang")
    set_run_font(r2, size=11, color=RGBColor(0xE8, 0xD5, 0xA3), east_asia="微软雅黑")
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(2)
    p2.paragraph_format.space_after = Pt(2)
    r = p2.add_run("万科上海区域投资副总经理（2017–2021）")
    set_run_font(r, size=11, bold=True, color=RGBColor(0xE8, 0xD5, 0xA3), east_asia="微软雅黑")
    p2b = cell.add_paragraph()
    p2b.paragraph_format.space_after = Pt(2)
    r = p2b.add_run("现任复旦大学住房政策研究中心秘书长  ·  城市更新 / 租赁住房")
    set_run_font(r, size=10, color=RGBColor(0xD5, 0xDE, 0xE8), east_asia="微软雅黑")
    p3 = cell.add_paragraph()
    p3.paragraph_format.space_after = Pt(0)
    r = p3.add_run("1987 年生  ｜  九三学社  ｜  上海  ｜  副教授级高级工程师（职称，见资质栏）")
    set_run_font(r, size=9, color=RGBColor(0xC5, 0xD0, 0xDC), east_asia="微软雅黑")
    p5 = cell.add_paragraph()
    p5.paragraph_format.space_after = Pt(0)
    r = p5.add_run("18678408669  ｜  262782809@qq.com  ｜  微信 hu262782809")
    set_run_font(r, size=9, color=RGBColor(0xC5, 0xD0, 0xDC), east_asia="微软雅黑")

    stats = doc.add_table(rows=1, cols=4)
    stats.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(stats, [4.55, 4.55, 4.55, 4.55])
    for i, (k, v) in enumerate(
        [("一线投资", "2011–2021"), ("投拓年限", "约 10 年"), ("靖江 + 金坛", "约 43 亿"), ("新进城市", "盐城 · 台州 · 金坛")]
    ):
        c = stats.rows[0].cells[i]
        clear_cell(c)
        shade_cell(c, SOFT)
        set_cell_borders(c, LINE, "4")
        set_cell_margin(c, 60, 60, 70, 70)
        para_in(c, k, size=8.5, color=MUTED, space_after=0)
        p = c.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(v)
        set_run_font(run, size=12, bold=True, color=NAVY, east_asia="微软雅黑")

    add_p(
        doc,
        "万科上海区域投资副总经理。2011–2021 年一线投拓，做过收并购、策略勾地、不良处置、物流园落地与国资合作；"
        "2021 年起转向住房保障、租赁市场与城市更新研究。一线投资只计至 2021 年，不把此后研究年限算进投资经验。",
        size=9.5,
        color=GRAY,
        space_before=8,
        space_after=2,
        line=1.18,
    )

    section_title(doc, "现任")
    job_header(doc, "复旦大学住房政策研究中心", "秘书长", "2021.06 – 至今")
    add_p(
        doc,
        "上海。副教授级高级工程师为职称，不是本职务的行政级别。",
        size=9,
        color=MUTED,
        space_before=2,
        space_after=2,
    )
    for t in [
        "牵头住房保障体系、租赁市场发展、城市更新等领域课题，撰写研究报告，为政府决策提供学术支持（原稿口径；具体课题名称与采纳单位待补）。",
        "组织论坛与闭门研讨，推动学界、企业与政府部门对话。",
        "可核对的公开活动：2024.06 主持锦天城「他山之石——中国投资者的海外不动产战略布局」；2025.05 参加中心联合主办的「全球新经济增长引擎峰会」房地产×资本×产业圆桌；2025.05 以中心秘书长身份出席上海市工商联房地产商会资管分会成立并登台战略合作。",
    ]:
        bullet(doc, t)

    section_title(doc, "社会职务")
    grid = doc.add_table(rows=2, cols=3)
    set_col_widths(grid, [6.05, 6.05, 6.1])
    items = [
        ("杨浦区科技企业联合会", "执行会长（公开报道可核验）"),
        ("复旦不动产资管协会", "创始参与（2018/2019）"),
        ("工商联房地产商会", "中心合作与活动出席"),
        ("上海山东省商会", "理事（原稿自述，待核）"),
        ("复旦研究生管理联考", "面试官（原稿自述，待核）"),
        ("评审专家", "城市更新研究会、节能减排工程技术协会等"),
    ]
    for idx, (title, body) in enumerate(items):
        r, c = divmod(idx, 3)
        cell = grid.rows[r].cells[c]
        clear_cell(cell)
        shade_cell(cell, SOFT)
        set_cell_borders(cell, LINE, "4")
        set_cell_margin(cell, 60, 60, 70, 70)
        v_align(cell, "top")
        para_in(cell, title, size=9.5, bold=True, color=NAVY, space_after=1)
        p = cell.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(body)
        set_run_font(run, size=8, color=MUTED, east_asia="微软雅黑")
    add_p(
        doc,
        "社会职务与复旦中心拆开书写。工商联房地产商会原稿标题写「秘书长」、正文或作「副秘书长」；"
        "官网现任秘书长为李祥，本版不写入现任。复旦 MBA 不动产资管协会现任秘书长公开为武金城。",
        size=8,
        color=MUTED,
        space_before=4,
        space_after=2,
    )

    section_title(doc, "专业经历")
    job_header(doc, "万科企业股份有限公司 · 上海区域", "江浙事业部投资副总经理", "2017.03 – 2021.06")
    add_p(
        doc,
        "上海。汇报投资合作开发部总经理。直接下属约 8 人。",
        size=9,
        color=MUTED,
        space_before=2,
        space_after=2,
    )
    for t in [
        "负责长三角江浙沪重大项目投资拓展，覆盖收并购、策略勾地、不良处置与国资合作。",
        "靖江印象城（2017.10）：勾地 + 印力委管，约 15 亿元 / 556 亩，自持购物中心约 7 万方。",
        "常州金坛理想城：约 28.05 亿元 / 561 亩，底价获取。",
        "镇江大港金域蓝湾：收并购 + 与大港集团各半合作，约 199 亩 / 32 万方。",
        "嘉兴海宁商住代建：18.53 亿元 / 约 23 万㎡；上海龙湖滟澜山法拍，测算 IRR 约 16.16%。",
        "主导获取并落地万纬嘉兴平湖、上海奉贤南桥冷链物流园；推动进入盐城、台州、常州金坛。",
    ]:
        bullet(doc, t)

    doc.add_page_break()
    section_title(doc, "专业经历（续）")
    job_header(doc, "新城控股集团股份有限公司 · 总部战略投资中心", "投资拓展资深专业经理", "2015.11 – 2017.02")
    for t in [
        "负责山东、江苏等区域战略布局与土地获取。",
        "主导临沂新城吾悦广场勾地，对接市级重点招商政策，系集团首进山东的重资产项目。",
        "以轻资产模式获取青岛新城吾悦广场（品牌输出 + 委管，约 13.68 万㎡）。",
    ]:
        bullet(doc, t)

    job_header(doc, "中南建设集团 / 中南控股集团", "投资拓展经理 / 战略企划管理主管", "2011.06 – 2015.10")
    add_p(doc, "上海 / 南通。2011.06–2012.10 总裁办战略企划；其后新项目发展中心。", size=9, color=MUTED, space_before=2, space_after=2)
    for t in [
        "从集团战略企划到项目投资拓展的完整历练。",
        "获取如皋中南世纪城、苏州中南锦苑；江阴白鹭湾与碧桂园同股同权合作。",
        "推进美国养老、日本永旺等国际合作对接（推进中，非已签约落地）。2014 年底派驻四川任西南拓展负责人。",
    ]:
        bullet(doc, t)

    section_title(doc, "社会职务说明")
    for t in [
        "杨浦科企联：主持日常运作与战略发展，策划科创项目评审、投融资对接与政策宣讲；2026.03 以执行会长身份参与「北欧创新国际会客厅」落户杨浦宝龙旭辉广场环创中心（区政府官网）。",
        "工商联房地产商会：原稿写会员发展、行业培训、产业论坛与白皮书。公开活动以住房政策中心秘书长出席，不把商会秘书长写成现任。",
        "复旦不动产资管协会：创立并运营面向校友的交流平台。现任秘书长请以院友会公布为准。",
    ]:
        bullet(doc, t)

    section_title(doc, "代表项目")
    proj = doc.add_table(rows=10, cols=4)
    proj.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(proj, [4.4, 2.6, 5.6, 5.6])
    proj_rows = [
        ("项目", "角色", "规模", "方式"),
        ("靖江印象城", "主导获取", "约 15 亿 / 556 亩 / 购物中心约 7 万方", "勾地 + 印力委管"),
        ("金坛理想城", "主导获取", "约 28.05 亿 / 561 亩", "收并购 + 债务重组，底价"),
        ("镇江大港金域蓝湾", "主导", "约 199 亩 / 32 万方", "收并购 + 国资各半合作"),
        ("嘉兴海宁商住", "拓展", "18.53 亿 / 约 23 万㎡", "政府代建"),
        ("上海龙湖滟澜山", "收并购", "测算 IRR 约 16.16%", "法拍不良资产"),
        ("南桥 / 平湖冷链园", "主导获取落地", "原稿未列总价", "物流产业园勾地 / 收购改造"),
        ("临沂新城吾悦广场", "主导勾地", "原稿未列总价", "集团首进山东 · 重资产"),
        ("青岛新城吾悦广场", "获取", "约 13.68 万㎡", "轻资产委管"),
        ("如皋中南世纪城 / 苏州中南锦苑", "获取", "原稿未列总价", "早期投资拓展"),
    ]
    for r, row in enumerate(proj_rows):
        for c, val in enumerate(row):
            cell = proj.rows[r].cells[c]
            clear_cell(cell)
            set_cell_borders(cell, LINE, "6")
            set_cell_margin(cell, 35, 35, 50, 50)
            if r == 0:
                shade_cell(cell, NAVY_HEX)
                para_in(cell, val, size=8.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
            else:
                if r % 2 == 0:
                    shade_cell(cell, SOFT)
                para_in(cell, val, size=8, bold=(c == 0), color=NAVY if c == 0 else GRAY)

    section_title(doc, "教育背景")
    job_header(doc, "复旦大学", "工商管理硕士 · 财务金融方向", "2018.09 – 2021.06")
    add_p(doc, "在职攻读，与万科上海区域任职并行。", size=9, color=MUTED, space_before=1, space_after=3)
    job_header(doc, "中国海洋大学", "土木工程 · 工学学士", "2007.09 – 2011.07")

    section_title(doc, "资质与研究")
    add_p(doc, "副教授级高级工程师；会计师中级；PMP；住建委项目管理工程师；IELTS 6.0；普通话一级甲等；BIM；上海市人才引进认证资质。", size=9.5, color=GRAY, space_after=2)
    add_p(
        doc,
        "行业核心期刊论文《万科物流地产平台业务发展战略研究》等。实用新型专利 2 项（建筑工程管理用功能脚架、防护围栏；专利号原稿脱敏，不写残号）。"
        "上海交通大学城市治理数字化转型高级研修班；复旦大学–花旗银行实践课程。",
        size=9.5,
        color=GRAY,
        space_after=2,
    )
    add_p(
        doc,
        "专业方向：投资拓展、收并购、策略勾地、不良资产、物流地产、产业园区、住房保障、租赁住房、城市更新、国资合作、募投管退。"
        "标签只标已做过或研究过的方向，不把未操盘的 REITs、保租房收储写成个人业绩。",
        size=9,
        color=MUTED,
        space_before=4,
        space_after=2,
    )
    add_p(
        doc,
        "投递方向：城市更新 / 租赁住房与存量收并购 / 产业不动产（国资平台、城投、住房与更新机构）。身份证、住址、配偶信息未写入。",
        size=8,
        color=MUTED,
        space_after=0,
    )

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "胡继刚-简历-优化版.docx"
    doc.save(path)
    return path


def write_html() -> tuple[Path, Path]:
    OUT.mkdir(parents=True, exist_ok=True)
    tpl = Path(__file__).resolve().parent / "templates"
    full = OUT / "胡继刚-简历-优化版.html"
    brief = OUT / "胡继刚-简历-一页精华.html"
    full.write_text((tpl / "resume-full.html").read_text(encoding="utf-8"), encoding="utf-8")
    brief.write_text((tpl / "resume-brief.html").read_text(encoding="utf-8"), encoding="utf-8")
    return full, brief


def write_notes() -> Path:
    text = """# 胡继刚简历优化说明（第三版）

按「国资城市更新 / 租赁住房 / 存量收并购」主投方向重排。人名统一为 **胡继刚**。事实来自原稿与已核对的公开报道，没有添造业绩数字。

打开 `胡继刚-简历-优化版.html` 可预览两页 A4、切换中英、查看本说明，并在打印框里选「另存为 PDF」。

## 结构

1. **现任与兼职拆开。** 复旦住房政策研究中心单独成条；社会职务改成网格。
2. **职称归位。** 「副教授级高级工程师」从任职标题剥离，放入资质栏。
3. **两页收口。** 第一页：现在的身份 + 万科证明。第二页：新城 / 中南、项目表、教育与资质。抬头先写万科投资副总，避免被读成「办论坛的人」。
4. **可预览、可切换中英、可导出 PDF。** 纸张选 A4。

## 内容

1. **删空话，留证据。** 去掉「复合型专家」「业绩卓著」。改用可核对的项目、金额、城市和职责。
2. **商会职务不强行写成现任秘书长。** 原稿标题「秘书长」、正文或作「副秘书长」。工商联官网现任秘书长为李祥；公开报道里胡继刚是以住房政策中心秘书长出席资管分会。本版 **不写入现任**。若实际是副职，按事实改回即可。
3. **不动产协会不写成确定现任。** 2018/2019 创立可写；管院校友页现任秘书长为武金城。
4. **教育变成加分项。** 复旦 MBA 标明「在职攻读」，与万科任职并行。
5. **诚实的时间口径。** 一线投资写 2011–2021（约 10 年），政策研究写 2021 至今。
6. **代表项目收成一表。** 靖江、金坛、冷链园、吾悦广场及海宁、镇江、法拍等原稿数字收入同一张表。原稿没有总价的写明「原稿未列总价」。靖江 + 金坛约 43 亿是 15 亿与 28 亿的合计。

## 公开职务核验

| 职务 | 本版写法 | 依据 |
| --- | --- | --- |
| 复旦大学住房政策研究中心秘书长 | 现任 | 央广网、人民网、杨浦区政府、工商联、锦天城 |
| 杨浦区科技企业联合会执行会长 | 现任 | 杨浦区政府、央广网、上观 |
| 工商联房地产商会秘书长 | **不写入现任** | 官网现任为李祥 |
| 复旦 MBA 不动产资管协会秘书长 / 理事长 | 创始参与，不作为确定现任 | 2024.06 报道曾用秘书长；院友会现公布武金城 |
| 山东商会理事、联考面试官、评审专家 | 原稿自述，待核 | 未见权威名单页 |

## 就业建议（不写进对外简历正文）

传统住宅投拓结构上已收缩。这份履历要卖的是：会拿地、会收并购、懂住房与更新政策的人，转去盘存量、做租赁与城市更新。

1. **主攻：** 上海地产集团及城市更新 / 住发 / 城方、杨浦城投与区属保障房公司、华润置地与保利发展上海的更新 / 产城条线。岗位名搜城市更新投资、租赁住房投资、存量收购、资产管理，不要海投普通「投资经理」。
2. **辅线：** AMC / 物流与产业园。弹药是不良处置、冷链园、收并购。未操盘过的 REITs 退出不要写成业绩。
3. **创收：** 把网络产品化——区属国企来沪的更新 / 租赁路径顾问、片区人才安居对接、课题与内训。
4. **不要做：** 海投民营房企投拓；不要为转行去考与履历错配的证书班；协会头衔投市场化基金时再删薄。

九三学社：投国企、区属平台、智库保留；投市场化基金可拿掉。

待补后简历会更硬：课题名称与采纳单位、复旦中心是否全职带薪、商会到底是秘书长还是副秘书长、专利号全文、2021 年后是否有更新 / 租赁项目交付。

## 文件

- `output/胡继刚-简历-优化版.html` / `.pdf` / `.docx`：两页投递稿（HTML 含中英与优化说明）
- `output/胡继刚-简历-一页精华.html` / `.pdf`：沟通用
"""
    path = OUT / "简历优化说明.md"
    path.write_text(text, encoding="utf-8")
    return path


def export_pdf(html_path: Path, pdf_path: Path) -> None:
    import subprocess
    import tempfile
    import time

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
            "--virtual-time-budget=8000",
            f"--print-to-pdf={pdf_path}",
            html_uri,
        ]
        try:
            subprocess.run(cmd, check=True, timeout=20)
        except subprocess.TimeoutExpired:
            time.sleep(0.5)
            if not pdf_path.exists() or pdf_path.stat().st_size < 1000:
                raise
            print(f"chrome 打印后未退出，但已生成：{pdf_path.name}")


def main():
    docx_path = build_docx()
    html_full, html_brief = write_html()
    notes = write_notes()
    pdf_full = OUT / "胡继刚-简历-优化版.pdf"
    pdf_brief = OUT / "胡继刚-简历-一页精华.pdf"
    export_pdf(html_full, pdf_full)
    export_pdf(html_brief, pdf_brief)
    print(f"docx: {docx_path}")
    print(f"html: {html_full}")
    print(f"html_brief: {html_brief}")
    print(f"pdf: {pdf_full}")
    print(f"pdf_brief: {pdf_brief}")
    print(f"notes: {notes}")


if __name__ == "__main__":
    main()
