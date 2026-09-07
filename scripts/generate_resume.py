#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成胡继刚优化版简历（Word + 供打印的 HTML）。

主线：产业合作与招商。AI 为新业务能力，出海为中期方向。
不写入身份证、家庭住址、配偶等隐私；不把未核验职务写成现任；不编造入驻/营收数字。
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
    r = p2.add_run("产业发展  ·  产业合作  ·  招商负责人")
    set_run_font(r, size=12, bold=True, color=RGBColor(0xE8, 0xD5, 0xA3), east_asia="微软雅黑")
    p2b = cell.add_paragraph()
    p2b.paragraph_format.space_before = Pt(0)
    p2b.paragraph_format.space_after = Pt(3)
    r = p2b.add_run("AI 商业化 = 新业务能力　　企业出海 / 海外园区合作 = 中期方向")
    set_run_font(r, size=9, color=RGBColor(0xD5, 0xDE, 0xE8), east_asia="微软雅黑")
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
                "产业合作 / 招商负责人（科创园区、产业集团、企业服务平台）    上海    薪资面议",
                {"size": 9.5, "color": NAVY},
            ),
        ],
        space_before=8,
        space_after=2,
    )

    section_title(doc, "职业摘要")
    add_p(
        doc,
        "土木本科、复旦 MBA。万科、新城经历支撑项目判断、投资拓展与合作落地；研究中心、杨浦科企联经历支撑政策理解、企业连接与跨机构协调。"
        "两者结合，是转向产业发展岗位的基础。职业主线放在产业合作与招商；AI 作为新业务能力，从园区、资管和企业服务场景切入；"
        "出海作为中期延伸，先做国内客户开发与项目组织。2021 年后公开成果以平台搭建和活动组织为主，付费项目与入驻签约数字待补。",
        size=9.5,
        color=GRAY,
        space_after=2,
        line=1.18,
    )

    section_title(doc, "方向优先级")
    cards = doc.add_table(rows=1, cols=3)
    cards.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(cards, [6.05, 6.05, 6.1])
    pillars = [
        (
            AMBER_HEX,
            "近期主线｜产业发展 / 招商",
            [
                "科创园区、产业集团、企业服务平台：项目拓展、合作谈判与落地。",
                "证据：拿地、收并购、国企合作、产业载体导入。",
                "公开报道尚无引进企业签约入驻的经营数字。",
            ],
        ),
        (
            TEAL_HEX,
            "重点突破｜AI 行业合作",
            [
                "从园区、资管、企业服务场景联合技术团队做客户项目。",
                "已有：峰会主办方代表与圆桌主持、建筑 AI 论文。",
                "待补：付费试点与客户项目。",
            ],
        ),
        (
            BLUE_HEX,
            "中期延伸｜企业出海服务",
            [
                "国内客户开发、需求梳理和项目组织，与当地机构共同交付。",
                "已有：海外不动产主题主持、北欧会客厅发起。",
                "待补：出海成交项目。",
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

    section_title(doc, "三个代表案例")
    for title, body in [
        (
            "1. 投资拓展落地｜万科靖江印象城（2017.10）",
            "职责：江浙事业部投资副总经理。动作：重资产勾地并引入印力轻资产委管。结果：约 15 亿获取 556 亩及约 7 万方自持购物中心，苏中县市首家印象城。说明：证明项目判断与合作落地；属拿地能力，不等于后续招商入驻经营数据。",
        ),
        (
            "2. 近期产业合作｜北欧创新国际会客厅（2026.03）",
            "职责：杨浦科企联执行会长；中心为共同发起方。动作：与宝龙商办、复旦国家大学科技园、峰唯峰等共同发起并参与揭牌交流。结果：会客厅落户杨浦宝龙旭辉广场环创中心（区政府官网）。说明：证明跨机构协调与载体导入；公开报道未披露入驻企业数或收费。",
        ),
        (
            "3. AI 新业务能力｜人工智能商业化峰会（2026.05）",
            "职责：主办方代表，主持「从算力引擎到新质资产」圆桌。动作：对接腾讯云、火山引擎、达观数据等嘉宾，提出将峰会升级为科创企业超级孵化器。结果：约 600 人，央广网/杨浦区政府/上观报道。说明：证明 AI 行业合作组织能力，尚不是付费客户项目。",
        ),
    ]:
        add_p(doc, title, size=10, bold=True, color=NAVY, space_before=2, space_after=1)
        add_p(doc, body, size=8.5, color=GRAY, space_after=2, line=1.12)

    section_title(doc, "工作经历")

    job_header(doc, "复旦大学住房政策研究中心", "秘书长 / 副教授级高级工程师", "2021.06 – 至今")
    add_p(
        doc,
        "兼杨浦区科技企业联合会执行会长。住房保障、租赁、城市更新课题 + 高峰论坛与战略合作。2021 年后公开成果以平台搭建和活动组织为主，营收/客户价值数字待补。",
        size=9,
        color=MUTED,
        space_before=2,
        space_after=2,
        line=1.12,
    )
    for t in [
        "产业合作：2026.03 北欧创新国际会客厅共同发起并落户杨浦园区；2025.05 代表中心出席工商联房地产商会资管分会成立大会并参与战略合作（中心为合作方；商会现任秘书长公开为李祥）。",
        "AI 能力：2026.05 人工智能商业化峰会主办方代表并主持圆桌；2025.05 中心联合主办全球新经济增长引擎峰会，本人参加房地产×资本×产业圆桌。",
        "出海铺垫：2024.06 主持「他山之石——中国投资者的海外不动产战略布局」。",
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

    section_title(doc, "投资拓展项目（拿地 / 合作落地，不等于企业招商入驻）")
    add_p(
        doc,
        "下列金额与面积来自原简历，证明投资判断与政府/国企协同；引进企业、签约入驻、收入贡献需另补经营台账。",
        size=8.5,
        color=MUTED,
        space_after=3,
    )
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
            "已核验 / 待核验职务",
            "已核验：住房政策研究中心秘书长；杨浦科企联执行会长\n"
            "2024.06 报道曾写复旦 MBA 不动产资管协会秘书长；院友会现公布为武金城，现任请核对\n"
            "工商联房地产商会：出席并代表中心参与战略合作，公开秘书长为李祥，不写入现任\n"
            "简历自述待核：城市更新 / 碳中和 / 虹口科企评审专家；复旦联考面试官；山东商会理事",
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
        "说明：公开活动可对照央广网、人民网、上观新闻、杨浦区政府官网、锦天城官网、工商联官网、复旦管院校友页核验。"
        "原 Boss 直聘稿中的身份证、家庭住址、配偶信息已移除。2021 年后若补上团队规模、付费试点、入驻签约，可将案例 2、3 改为经营成果表达。",
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
    text = """# 胡继刚简历优化说明（第二版）

按「产业合作与招商为主线，AI 为新业务能力，出海为中期方向」重排。人名统一为 **胡继刚**。

## 本版相对上一版的关键调整

1. **求职意向收窄**：产业合作 / 招商负责人（科创园区、产业集团、企业服务平台）。不再把 AI、出海、政府关系、会务并列成四个求职方向。
2. **能力不再三线平铺**：近期主线是产业发展与招商；AI 是重点突破的新能力；出海是中期延伸。
3. **补上三个可核验案例**（投资项目 / 近期产业合作 / AI 活动），写清职责、动作、结果，并标明还缺什么经营数字。
4. **区分两类证据**：万科、新城的拿地、收并购、物流园获取 = 项目拓展；真正的招商成绩还需要引进企业、签约入驻。公开报道里没有后者，所以没有编造。
5. **核验社会职务后降级或删除无法交叉印证的现任头衔。**

## 公开信息核验（2026-09-07）

| 职务 | 结论 | 依据 |
| --- | --- | --- |
| 复旦大学住房政策研究中心秘书长 | 保留现任 | 央广网、人民网、杨浦区政府、工商联、锦天城 |
| 杨浦区科技企业联合会执行会长 | 保留现任 | 杨浦区政府、央广网、上观 |
| 工商联房地产商会秘书长 | **不写入现任** | 工商联官网现任秘书长为李祥；胡继刚以中心秘书长身份出席资管分会并登台参与战略合作 |
| 复旦 MBA 不动产资管协会秘书长 / 创始理事长 | **不作为现任** | 2024.06 锦天城报道曾如此称呼；管院校友页现公布秘书长为武金城 |
| 上海山东省商会理事 | **待核** | 原简历自述，公开理事会名单未交叉印证 |
| 三类评审专家、复旦联考面试官 | **待核** | 来自原简历，未见权威名单页 |

## 三个代表案例怎么读

| 案例 | 能证明什么 | 还不能写成什么 |
| --- | --- | --- |
| 靖江印象城 | 项目判断、勾地、轻资产合作落地 | 购物中心后续招商去化、租金或利润 |
| 北欧创新国际会客厅 | 跨机构协调、国际化载体导入杨浦园区 | 入驻企业数、收费、持续运营 KPI |
| 2026 AI 商业化峰会 | AI 行业合作组织、圆桌主持、主办方代表发声 | 付费客户、商业化合同、孵化营收 |

杨浦《深入推进「人工智能+」行动方案（2026–2028）》强调政校企联动的应用场景和产业载体；上海市商务委出海服务体系强调与商协会、专业机构联动。这两条是岗位切入口，不是个人已经完成的招商或出海业绩。

## 未来 90 天建议（写在说明里，不写进对外简历）

1. 把三个案例补成经营台账：团队规模、关键动作、可出示结果。
2. 从现有企业联系中访谈 10–15 家，找一个有预算的问题，与技术或出海合作伙伴做一次付费试点。
3. 求职优先有明确业务、团队和预算的产业平台负责人岗位；创业则先验证一个服务能否持续成交。

补上付费试点或入驻签约后，再把案例 2、3 从活动表达改成业绩表达。

## 文件

- `output/胡继刚-简历-优化版.docx` / `.pdf`：两页投递稿
- `output/胡继刚-简历-一页精华.pdf`：沟通用
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
