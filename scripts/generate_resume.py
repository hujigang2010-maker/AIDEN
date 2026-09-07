#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成胡继刚优化版简历（Word + 供打印的 HTML）。

依据 2026-04-27 原简历，并补入 2024–2026 可公开核验的 AI、出海、招商经历。
不写入身份证、家庭住址、配偶等隐私信息；不编造未核验业绩。
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
TEAL_HEX = "0F6B63"
BLUE_HEX = "1D4F91"
AMBER_HEX = "9A4A12"
GRAY = RGBColor(0x44, 0x4C, 0x56)
MUTED = RGBColor(0x5C, 0x67, 0x73)
LINE = "D7DEE8"
SOFT = "F4F7FB"


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


def para_in(cell, text, *, size=10, bold=False, color=None, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=0, space_before=0):
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


def add_mixed(doc, segments, *, space_before=0, space_after=3, line=1.12):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line
    for text, kwargs in segments:
        run = p.add_run(text)
        set_run_font(run, east_asia="微软雅黑", **kwargs)
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
    p = add_p(doc, text, size=12, bold=True, color=NAVY, space_before=8, space_after=1)
    add_hr(doc, GOLD_HEX, "14")
    return p


def bullet(doc, text, *, size=9.5):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.32)
    pf.first_line_indent = Cm(-0.32)
    pf.space_before = Pt(0)
    pf.space_after = Pt(1.6)
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
    set_col_widths(table, [12.4, 5.2])
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

    # 页眉条
    head = doc.add_table(rows=1, cols=1)
    head.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(head, [18.2])
    cell = head.rows[0].cells[0]
    clear_cell(cell)
    shade_cell(cell, NAVY_HEX)
    set_cell_borders(cell, NAVY_HEX, "0")
    set_cell_margin(cell, 90, 90, 140, 140)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("胡继刚")
    set_run_font(r, size=22, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), east_asia="微软雅黑")
    r2 = p.add_run("    HU Jigang")
    set_run_font(r2, size=11, color=RGBColor(0xE8, 0xD5, 0xA3), east_asia="微软雅黑")
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(2)
    p2.paragraph_format.space_after = Pt(4)
    r = p2.add_run("产业招商  ·  AI 生态运营  ·  国际化资源对接")
    set_run_font(r, size=12, bold=True, color=RGBColor(0xE8, 0xD5, 0xA3), east_asia="微软雅黑")
    p3 = cell.add_paragraph()
    p3.paragraph_format.space_after = Pt(1)
    r = p3.add_run("复旦大学住房政策研究中心秘书长  ｜  上海市杨浦区科技企业联合会执行会长")
    set_run_font(r, size=9.5, color=RGBColor(0xF0, 0xF4, 0xF8), east_asia="微软雅黑")
    p4 = cell.add_paragraph()
    p4.paragraph_format.space_after = Pt(0)
    r = p4.add_run("男 ｜ 1987.04 ｜ 上海 ｜ 硕士 ｜ 副教授级高级工程师 ｜ 九三学社 ｜ IELTS 6.0")
    set_run_font(r, size=9, color=RGBColor(0xC5, 0xD0, 0xDC), east_asia="微软雅黑")
    p5 = cell.add_paragraph()
    p5.paragraph_format.space_after = Pt(0)
    r = p5.add_run("18678408669  ｜  262782809@qq.com  ｜  微信 hu262782809")
    set_run_font(r, size=9, color=RGBColor(0xC5, 0xD0, 0xDC), east_asia="微软雅黑")

    add_mixed(
        doc,
        [
            ("求职意向  ", {"size": 9, "bold": True, "color": GOLD}),
            (
                "产业招商 / 政府关系 / 国际化业务与出海对接 / AI 与科创生态运营    上海    薪资面议",
                {"size": 9.5, "color": NAVY},
            ),
        ],
        space_before=8,
        space_after=2,
    )

    section_title(doc, "职业摘要")
    add_p(
        doc,
        "土木本科、复旦 MBA，15 年不动产投资拓展与政府资源整合经验，历任中南、新城、万科投资岗至区域投资副总经理，"
        "熟悉勾地、轻资产输出、代建、收并购与不良资产「募投管退」。2021 年起任职高校智库秘书长，并以杨浦科企联执行会长身份搭建政产学研平台。"
        "近两年公开主持/主办人工智能商业化峰会、中国投资者海外不动产主题活动、北欧创新国际会客厅等，形成「AI 产业化落地 + 企业出海对接 + 政府招商落地」复合能力。",
        size=9.5,
        color=GRAY,
        space_after=2,
        line=1.18,
    )

    section_title(doc, "核心优势：AI × 出海 × 招商")
    cards = doc.add_table(rows=1, cols=3)
    cards.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(cards, [6.05, 6.05, 6.1])
    pillars = [
        (
            TEAL_HEX,
            "AI 产业化与生态运营",
            [
                "2026.05 作为主办方代表参与「人工智能商业化落地与硬核投资破局峰会」（约 600 人），主持「从算力引擎到新质资产」圆桌，提出将峰会升级为陪伴科创企业的超级孵化器。",
                "2025.05 联合主办全球新经济增长引擎峰会，议程覆盖 AI 产业链重构与中东/亚太出海投资。",
                "发表《房屋建筑工程人工智能技术的应用》；完成上海交大城市治理数字化转型高研班。",
                "以科企联执行会长身份服务科技企业评审、对接与成长。",
            ],
        ),
        (
            BLUE_HEX,
            "出海与国际化对接",
            [
                "2024.06 主持「他山之石——中国投资者的海外不动产战略布局」（锦天城 × 复旦住房政策研究中心）。",
                "2026.03 参与「北欧创新国际会客厅」揭牌，围绕中欧创新合作、科创生态联动与企业国际化交流。",
                "英语商务洽谈 / IELTS 6.0；早期推进美国养老、日本永旺等国际合作对接。",
                "能把高校智库、区级科创组织与涉外法务/产业顾问连成可落地的出海对接场。",
            ],
        ),
        (
            AMBER_HEX,
            "招商与政府资源落地",
            [
                "14 年投资拓展：重资产勾地、轻资产品牌输出、委托代建、股权收并购、法拍与不良资产。",
                "临沂吾悦广场对接市级重点招商政策；万科阶段落地靖江印象城、海宁、镇江大港、金坛及冷链园等。",
                "深耕江浙沪、皖鲁闽政府与国企关系，能独立完成可行性研究与落地协同。",
                "现职连接区科企联、高校智库、工商联房地产商会，具备「政策—空间—企业—资本」闭环。",
            ],
        ),
    ]
    for i, (accent, title, items) in enumerate(pillars):
        c = cards.rows[0].cells[i]
        clear_cell(c)
        shade_cell(c, SOFT)
        set_cell_borders(c, LINE, "4")
        set_cell_margin(c, 70, 70, 90, 90)
        v_align(c, "top")
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(3)
        bar = p.add_run("▍")
        rgb = RGBColor(int(accent[0:2], 16), int(accent[2:4], 16), int(accent[4:6], 16))
        set_run_font(bar, size=11, bold=True, color=rgb, east_asia="微软雅黑")
        t = p.add_run(title)
        set_run_font(t, size=10, bold=True, color=NAVY, east_asia="微软雅黑")
        for item in items:
            bp = c.add_paragraph()
            bp.paragraph_format.space_after = Pt(2)
            bp.paragraph_format.line_spacing = 1.08
            run = bp.add_run("• " + item)
            set_run_font(run, size=8, color=GRAY, east_asia="微软雅黑")

    section_title(doc, "工作经历")

    job_header(doc, "复旦大学住房政策研究中心", "秘书长 / 副教授级高级工程师", "2021.06 – 至今")
    add_p(
        doc,
        "兼杨浦区科技企业联合会执行会长、复旦 MBA 不动产资产管理协会秘书长（创始理事长）。牵头住房保障、租赁市场、城市更新等课题，组织高峰论坛、闭门研讨与战略合作。",
        size=9,
        color=MUTED,
        space_before=2,
        space_after=2,
        line=1.12,
    )
    for t in [
        "AI：2026.05 主办方代表，人工智能商业化落地与硬核投资破局峰会（上海北外滩，600+ 嘉宾）；主持「从算力引擎到新质资产（AI 全产业链商业化实战）」圆桌，公开提出做实产学研融协同、服务科创企业。",
        "出海：2026.03「北欧创新国际会客厅」揭牌（科企联、复旦住房政策研究中心等共同发起），交流中欧创新机制与企业国际化；2024.06 主持中国投资者海外不动产战略布局主题活动。",
        "招商/平台：2025.05 联合主办全球新经济增长引擎峰会并参加房地产×资本×产业圆桌；出席上海市工商联房地产商会资产管理分会成立大会并参与战略合作仪式。",
        "对外解读住房政策、接受媒体采访，持续把智库研究转化为政企对话与产业对接场景。",
    ]:
        bullet(doc, t)

    job_header(doc, "万科企业集团股份有限公司 上海区域", "江浙事业部 投资副总经理", "2017.03 – 2021.06")
    add_p(
        doc,
        "汇报投资合作开发部总经理，团队 8 人。深耕江浙未进入城市，负责策略勾地、收并购、债务重组与不良资产处置，覆盖商业综合体、产城、物流冷链。",
        size=9,
        color=MUTED,
        space_before=2,
        space_after=2,
        line=1.12,
    )
    for t in [
        "靖江印象城（2017.10）：重资产勾地 + 印力轻资产委管，约 15 亿获取 556 亩及约 7 万方自持购物中心，苏中县市首家印象城。",
        "政府/国企合作：嘉兴海宁商住代建拓展（18.53 亿 / 约 23 万㎡）；镇江大港金域蓝湾与大港集团各半合作；常州金坛理想城约 28.05 亿、561 亩底价获取。",
        "特资与产业：上海龙湖滟澜山法拍收并购（IRR 约 16.16%）；万纬嘉兴平湖、上海南桥冷链园政府勾地与收购改造（含智能化园区）。",
        "协助临沂河东鲁商新都会（国企基金 + 勾地 + 代建），并推动盐城、台州、金坛等城市进入。",
    ]:
        bullet(doc, t)

    doc.add_page_break()
    section_title(doc, "工作经历（续）")

    job_header(doc, "新城控股集团股份有限公司", "战略投资中心 投资拓展资深专业经理", "2015.11 – 2017.02")
    for t in [
        "临沂吾悦广场：对接市级重点招商引资政策，推动商业综合体落地；齐鲁吾悦、青岛吾悦（轻资产品牌输出 + 委管，约 13.68 万㎡）。",
        "维护华东及多区域政府关系，支持全国化拓展中的信息研判、谈判与落地协同。",
    ]:
        bullet(doc, t)

    job_header(doc, "中南建设集团 / 中南控股集团", "投资拓展经理 / 战略企划管理主管", "2011.06 – 2015.10")
    for t in [
        "如皋世纪城等勾地综合体、苏州中南锦苑、江阴白鹭湾（与碧桂园同股同权合作）等项目拓展；主导多宗公开竞拍。",
        "推进美国养老、日本永旺、平安不动产等合作对接；2014 年底晋升并派驻四川任西南拓展负责人。前期在总裁办负责战略企划与制度建设。",
    ]:
        bullet(doc, t)

    section_title(doc, "代表项目（招商 / 产业落地）")
    proj = doc.add_table(rows=9, cols=3)
    proj.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(proj, [4.6, 4.4, 9.2])
    proj_rows = [
        ("项目", "模式", "要点"),
        ("万科靖江印象城", "勾地 + 轻资产委管", "约 15 亿获取 556 亩及约 7 万方自持购物中心，苏中县市首家印象城"),
        ("嘉兴海宁商住", "政府代建拓展", "18.53 亿两宗商住地，建面约 23 万㎡，配建安置房"),
        ("镇江大港金域蓝湾", "收并购 + 国企合作", "199 亩 / 约 32 万方，与镇江大港集团各半合作"),
        ("常州金坛理想城", "收并购 + 债务重组", "约 28.05 亿、561 亩，底价获取"),
        ("上海龙湖滟澜山", "法拍不良资产", "金融测算 IRR 约 16.16%，特资「募投管退」样本"),
        ("万纬平湖 / 南桥冷链", "产业勾地 / 收购改造", "嘉兴首个高标准自动化冷链园；南桥智能化园区"),
        ("临沂吾悦广场", "市级重点招商", "对接市重点招商政策，推动商业综合体落地"),
        ("青岛吾悦广场", "轻资产品牌输出", "品牌输出 + 委管，约 13.68 万㎡"),
    ]
    for r, row in enumerate(proj_rows):
        for c, val in enumerate(row):
            cell = proj.rows[r].cells[c]
            clear_cell(cell)
            set_cell_borders(cell, LINE, "6")
            set_cell_margin(cell, 40, 40, 60, 60)
            if r == 0:
                shade_cell(cell, NAVY_HEX)
                para_in(cell, val, size=8.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), align=WD_ALIGN_PARAGRAPH.CENTER)
            else:
                if r % 2 == 0:
                    shade_cell(cell, SOFT)
                para_in(cell, val, size=8, bold=(c == 0), color=NAVY if c == 0 else GRAY)

    section_title(doc, "教育、资质与成果")
    edu = doc.add_table(rows=2, cols=2)
    set_col_widths(edu, [9.1, 9.1])
    blocks = [
        (
            "教育与培训",
            "2018.09–2021.06  复旦大学  工商管理硕士（财务金融）\n"
            "2007.09–2011.07  中国海洋大学  土木工程  工学学士\n"
            "上海交大 · 城市治理数字化转型高研班；复旦·怡安人力与风险管理；复旦·花旗银行实践课程",
        ),
        (
            "职称证书",
            "副教授级高级工程师；会计师中级；PMP；上海市住建委项目管理工程师；江苏省机械管理师\n"
            "IELTS 6.0；普通话一级甲等；BIM；上海市人才引进认证",
        ),
        (
            "论文与专利",
            "《万科物流地产平台业务发展战略研究》（知网）\n"
            "《房屋建筑工程人工智能技术的应用》（《住宅与房地产》）\n"
            "《铝模板在高层住宅建筑工程中的应用》《基于 BIM 的装配式钢结构建筑施工新技术与管理研究》\n"
            "实用新型专利 2 项：建筑工程管理用功能脚架；建筑工程用防护围栏",
        ),
        (
            "社会职务与专家",
            "杨浦区科技企业联合会执行会长；上海市工商联房地产商会秘书长\n"
            "复旦管院不动产资产管理协会创始理事长；上海山东省商会理事\n"
            "城市更新 / 碳中和节能 / 虹口科技企业评审专家；复旦研究生管理联考面试官",
        ),
    ]
    for idx, (title, body) in enumerate(blocks):
        r, c = divmod(idx, 2)
        cell = edu.rows[r].cells[c]
        clear_cell(cell)
        shade_cell(cell, SOFT)
        set_cell_borders(cell, LINE, "4")
        set_cell_margin(cell, 60, 60, 80, 80)
        v_align(cell, "top")
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(title)
        set_run_font(run, size=9.5, bold=True, color=NAVY, east_asia="微软雅黑")
        for line in body.split("\n"):
            bp = cell.add_paragraph()
            bp.paragraph_format.space_after = Pt(1)
            bp.paragraph_format.line_spacing = 1.08
            run = bp.add_run(line)
            set_run_font(run, size=8, color=GRAY, east_asia="微软雅黑")

    add_p(
        doc,
        "说明：公开活动职务与角色均可对照央广网、人民网、上观新闻、杨浦区政府官网、锦天城官网等报道核验。原 Boss 直聘稿中的身份证、家庭住址、配偶信息已从对外简历中移除。",
        size=8,
        color=MUTED,
        space_before=8,
        space_after=0,
    )

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "胡继刚-简历-优化版.docx"
    doc.save(path)
    return path


def write_html() -> tuple[Path, Path]:
    """复制精排 HTML 模板到 output，供 Chrome 打印 PDF。"""
    OUT.mkdir(parents=True, exist_ok=True)
    tpl = Path(__file__).resolve().parent / "templates"
    full = OUT / "胡继刚-简历-优化版.html"
    brief = OUT / "胡继刚-简历-一页精华.html"
    full.write_text((tpl / "resume-full.html").read_text(encoding="utf-8"), encoding="utf-8")
    brief.write_text((tpl / "resume-brief.html").read_text(encoding="utf-8"), encoding="utf-8")
    return full, brief


def write_notes() -> Path:
    text = """# 胡继刚简历优化说明

依据 `胡先生-简历0427.pdf`（Boss 直聘导出，抬头为「汤先生」）重写，并补入 2024–2026 可公开核验经历。人名统一为 **胡继刚**。

## 这次改了什么

1. **定位从「会务/政府关系」升级为三条主线并行**
   - **AI**：主办/主持人工智能商业化峰会、建筑 AI 论文、城市数字化培训、科企联服务科技企业。
   - **出海**：主持中国投资者海外不动产活动、北欧创新国际会客厅、英语商务洽谈、早期国际合作对接。
   - **招商**：14 年投资拓展与勾地落地，叠加现职政产学研平台。
2. **求职意向**改为：产业招商 / 政府关系 / 国际化业务与出海对接 / AI 与科创生态运营。薪资改为面议（原稿为会务/政府关系 30–60K×12，可按投递对象改回）。
3. **现职经历补全到 2026 年 5 月**，不再停在「组织论坛、解读政策」的笼统表述。
4. **工作业绩改为成果导向**：保留可核对的项目名、金额、模式（勾地/代建/收并购/法拍），删掉空泛的自我评价套话。
5. **对外稿去掉隐私**：身份证号、家庭住址、配偶姓名电话、身高体重不写入。需要填应聘登记表时再用原件。

## 三条主线怎么写进简历

| 主线 | 写进简历的事实（均可核验或来自原简历） | 没有写成的内容 |
| --- | --- | --- |
| AI | 2026.05 峰会主办方代表 + 主持圆桌（央广网/上观/中新网）；2025 峰会议程含 AI 产业链；建筑 AI 论文；交大数字化高研班 | 未把白皮书撰写、内部活动方案写成「WAIC 主办人」等无法核验头衔 |
| 出海 | 2024.06 海外不动产主题主持（锦天城官网）；2026.03 北欧会客厅（杨浦区政府/界面）；IELTS 6.0；中南阶段美国养老/日本永旺**对接**（原简历为推进中，未写成已签约） | 未把迪拜沙龙材料、跨境电商白皮书等内部文稿写成个人任职 |
| 招商 | 万科/新城/中南勾地与收并购业绩；临沂吾悦「市级重点招商」；现职连接科企联+智库+商会 | 未写入百科稿刻意排除的招商面积/签约额等经营数据 |

## 相对原稿的删减

- 大段性格描述、法律财税自我评价（改为项目事实）。
- 新城阶段过长的「到访城市清单」（改为「华东及多区域政府关系」）。
- 证书中育婴师、红十字救护员、中级厨师等与目标岗位无关项（需要生活技能证明时可补回）。
- 专利号仍为原稿脱敏号，故只写名称、不写残号。

## 投递时建议微调

- **投政府/园区招商**：把核心优势第三列和万科勾地项目前置，意向写「产业招商 / 政府关系」。
- **投 AI 机构/活动/生态运营**：强调 2026 峰会主办方角色、科企联、超级孵化器表述。
- **投国际化/律所/家办/出海服务**：强调 2024 海外不动产主持、北欧会客厅、英语。
- 若对方仍要会展执行岗，可把意向加回「会务会展」，峰会主办经验本身就是会展能力证明。

## 文件

- `output/胡继刚-简历-优化版.docx`：可编辑 Word
- `output/胡继刚-简历-优化版.pdf`：对外投递用（两页）
- `output/胡继刚-简历-一页精华.pdf`：活动/微信沟通用
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
