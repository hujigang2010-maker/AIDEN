#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成胡继刚优化版简历（Word + 可预览 HTML）。

主线：国资平台 · 城市更新 / 存量收并购 / 产业不动产。
现任只写复旦大学住房政策研究中心秘书长。不写房商会现任秘书长。
一线投资只计 2011–2021。印刷面不出现待补/待核/原稿口径。
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
    r = p2.add_run("国资平台 · 城市更新 / 存量收并购 / 产业不动产")
    set_run_font(r, size=11, bold=True, color=RGBColor(0xE8, 0xD5, 0xA3), east_asia="微软雅黑")
    p2b = cell.add_paragraph()
    p2b.paragraph_format.space_after = Pt(2)
    r = p2b.add_run("副教授级高级工程师 ｜ 复旦大学硕士研究生 ｜ 原万科上海区域投资副总经理")
    set_run_font(r, size=10, color=RGBColor(0xD5, 0xDE, 0xE8), east_asia="微软雅黑")
    p3 = cell.add_paragraph()
    p3.paragraph_format.space_after = Pt(0)
    r = p3.add_run("现任复旦大学住房政策研究中心秘书长  ｜  杨浦区科技企业联合会执行会长")
    set_run_font(r, size=9, color=RGBColor(0xC5, 0xD0, 0xDC), east_asia="微软雅黑")
    p4 = cell.add_paragraph()
    p4.paragraph_format.space_after = Pt(0)
    r = p4.add_run("1987 年生  ｜  九三学社  ｜  上海")
    set_run_font(r, size=9, color=RGBColor(0xC5, 0xD0, 0xDC), east_asia="微软雅黑")
    p5 = cell.add_paragraph()
    p5.paragraph_format.space_after = Pt(0)
    r = p5.add_run("18678408669  ｜  262782809@qq.com  ｜  微信 hu262782809")
    set_run_font(r, size=9, color=RGBColor(0xC5, 0xD0, 0xDC), east_asia="微软雅黑")

    stats = doc.add_table(rows=1, cols=4)
    stats.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(stats, [4.55, 4.55, 4.55, 4.55])
    for i, (k, v) in enumerate(
        [("一线投拓", "10 年"), ("万科团队", "约 8 人"), ("靖江 + 金坛", "43 亿"), ("新进城市", "盐城 / 台州 / 金坛")]
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
        "万科上海区域投资副总经理。2011–2021 年一线投拓，做成国资合作、政府代建、不良盘活、轻资产委管与物流园落地。"
        "2021 年起任复旦大学住房政策研究中心秘书长，研究城市更新与存量资产；租赁住房为研究方向，不是操盘业绩。一线投资计至 2021 年。",
        size=9.5,
        color=GRAY,
        space_before=8,
        space_after=2,
        line=1.18,
    )

    section_title(doc, "现任")
    job_header(doc, "复旦大学住房政策研究中心", "秘书长", "2021.06 – 至今")
    add_p(doc, "上海", size=9, color=MUTED, space_before=2, space_after=2)
    for x in [
        "围绕城市更新、存量资产开展政策与产业研究，为政府和企业提供决策支持。租赁住房为研究方向。",
        "2025.05「全球新经济增长引擎峰会」：主办方代表发言，并与竺劲、汪毅、王维军圆桌（人民网上海、中新社上海点名）。",
        "2025.05.28 以中心秘书长出席上海市工商联房地产商会资管分会成立（上海德必）。",
        "2024.06 主持锦天城「他山之石——中国投资者的海外不动产战略布局」。",
    ]:
        bullet(doc, x)

    section_title(doc, "代表案例 · 存量商业空间产业导入")
    job_header(doc, "北欧创新国际会客厅", "共同发起 / 落地杨浦", "2026.03")
    add_p(doc, "杨浦宝龙旭辉广场环创中心。宝龙商办为共同发起单位，不是雇主。", size=9, color=MUTED, space_before=2, space_after=2)
    for x in [
        "复旦大学住房政策研究中心、杨浦区科技企业联合会与宝龙商办等共同发起，将国际创新合作平台导入存量商业空间。",
        "公开报道点名会长王晓明、执行会长胡继刚（杨浦区政府官网等）。",
    ]:
        bullet(doc, x)

    section_title(doc, "社会职务")
    grid = doc.add_table(rows=1, cols=2)
    set_col_widths(grid, [9.1, 9.1])
    for i, (title, body) in enumerate(
        [("杨浦区科技企业联合会", "执行会长"), ("复旦不动产资管协会", "创始人")]
    ):
        cell = grid.rows[0].cells[i]
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

    section_title(doc, "专业经历")
    job_header(doc, "万科企业股份有限公司 · 上海区域", "江浙事业部投资副总经理", "2017.03 – 2021.06")
    add_p(doc, "上海。汇报投资合作开发部总经理。团队约 8 人。", size=9, color=MUTED, space_before=2, space_after=2)
    for x in [
        "镇江大港金域蓝湾：收并购，与大港集团各半合作（约 199 亩 / 32 万方）。",
        "嘉兴海宁商住：政府代建落地（18.53 亿元 / 约 23 万㎡）。",
        "上海龙湖滟澜山：法拍不良资产盘活（测算 IRR 约 16.16%）。",
        "靖江印象城：勾地后引入印力轻资产委管（约 15 亿元 / 556 亩 / 自持购物中心约 7 万方）；金坛理想城底价收并购（约 28.05 亿元 / 561 亩）。",
        "主导获取并落地万纬嘉兴平湖、上海奉贤南桥冷链物流园；推动进入盐城、台州、常州金坛。",
    ]:
        bullet(doc, x)

    doc.add_page_break()
    section_title(doc, "专业经历（续）")
    job_header(doc, "新城控股集团股份有限公司 · 总部战略投资中心", "投资拓展资深专业经理", "2015.11 – 2017.02")
    for x in [
        "负责山东、江苏等区域战略布局与土地获取。",
        "主导临沂新城吾悦广场勾地，对接市级重点招商政策，系集团首进山东的重资产项目。",
        "以轻资产模式获取青岛新城吾悦广场（品牌输出 + 委管，约 13.68 万㎡）。",
    ]:
        bullet(doc, x)

    job_header(doc, "中南建设集团 / 中南控股集团", "投资拓展经理 / 战略企划管理主管", "2011.06 – 2015.10")
    add_p(doc, "上海 / 南通。2011.06–2012.10 总裁办战略企划；其后新项目发展中心。", size=9, color=MUTED, space_before=2, space_after=2)
    for x in [
        "从集团战略企划到项目投资拓展的完整历练；获取如皋中南世纪城、苏州中南锦苑。",
        "江阴白鹭湾与碧桂园同股同权合作。2014 年底派驻四川任西南拓展负责人。",
        "推进美国养老、日本永旺等国际合作对接（推进中，非已签约落地）。",
    ]:
        bullet(doc, x)

    section_title(doc, "代表项目")
    rows = [
        ("项目", "类型", "做法", "规模"),
        ("镇江大港金域蓝湾", "国资合作", "收并购，与大港集团各半", "约 199 亩 / 32 万方"),
        ("嘉兴海宁商住", "政府代建", "代建落地", "18.53 亿 / 约 23 万㎡"),
        ("上海龙湖滟澜山", "不良盘活", "法拍收并购", "测算 IRR 约 16.16%"),
        ("靖江印象城", "轻资产委管", "勾地 + 印力委管", "约 15 亿 / 556 亩 / 购物中心约 7 万方"),
        ("南桥 / 平湖冷链园", "产业园", "获取落地", "物流园"),
        ("金坛理想城", "收并购", "底价获取", "约 28.05 亿 / 561 亩"),
        ("临沂新城吾悦广场", "政府招商", "对接市级重点招商政策", "集团首进山东"),
        ("青岛新城吾悦广场", "轻资产", "品牌输出 + 委管", "约 13.68 万㎡"),
    ]
    proj = doc.add_table(rows=len(rows), cols=4)
    proj.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(proj, [4.6, 2.8, 5.4, 5.4])
    for r, row in enumerate(rows):
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
    job_header(doc, "复旦大学", "硕士研究生 · 工商管理（财务金融方向）", "2018.09 – 2021.06")
    add_p(doc, "在职攻读，与万科上海区域任职并行。", size=9, color=MUTED, space_before=1, space_after=3)
    job_header(doc, "中国海洋大学", "土木工程 · 工学学士", "2007.09 – 2011.07")

    section_title(doc, "资质")
    add_p(doc, "副教授级高级工程师；会计师中级；IELTS 6.0；上海市人才引进认证资质。", size=9.5, color=GRAY, space_after=2)
    add_p(
        doc,
        "论文《万科物流地产平台业务发展战略研究》。实用新型专利 2 项（建筑工程管理用功能脚架、防护围栏）。",
        size=9.5,
        color=GRAY,
        space_after=2,
    )
    add_p(
        doc,
        "投递方向：国资平台 · 城市更新 / 存量收并购 / 产业不动产。能力：国资合作、政府代建、不良盘活、轻资产委管、产业园落地。"
        "租赁住房为研究方向。",
        size=9,
        color=MUTED,
        space_before=4,
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
    text = """# 胡继刚简历优化说明（第四版 · 国资投递稿）

印刷面已清掉「待补 / 待核 / 原稿口径 / 原稿未列总价」。现任只写 **复旦大学住房政策研究中心秘书长**。宝龙商办是北欧会客厅共同发起单位，不是雇主。

## 必须改过的

1. 页眉求职方向：国资平台 · 城市更新 / 存量收并购 / 产业不动产。租赁住房只作研究方向。
2. 2025.05 峰会按人民网上海、中新社上海点名：主办方代表发言，并与竺劲、汪毅、王维军圆桌。
3. 北欧会客厅提到第 1 页，写成存量商业空间产业导入；公开报道点名会长王晓明、执行会长胡继刚。
4. 万科按国资语言重排：镇江国企各半 → 海宁政府代建 → 滟澜山不良盘活 → 印力轻资产委管 → 冷链园。
5. 社会职务只留科企联执行会长、资管协会创始人。
6. 资管分会（2025.05.28 德必）只写以中心秘书长出席。

## 没有改的

一线投资计至 2021；副教授级高工归职称；房商会不写现任秘书长（官网为李祥）；不写身份证住址配偶；不把 REITs / 保租房收储写成业绩；中南保留江阴白鹭湾同股同权与 2014 西南拓展。

## 文件

- `output/胡继刚-简历-优化版.html` / `.pdf` / `.docx`
- `output/胡继刚-简历-一页精华.html` / `.pdf`
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
