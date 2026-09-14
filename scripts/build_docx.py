# -*- coding: utf-8 -*-
"""生成 Word：东昇聚变政府事务与上海产业落地 90 天工作设想。"""
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

import content as C

OUT = Path(__file__).resolve().parent.parent / "deliverables"
OUT.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT / "东昇聚变_政府事务与上海产业落地_90天工作设想.docx"

NAVY = (0x07, 0x1A, 0x2B)
CYAN = (0x1F, 0xA8, 0xB4)
EMBER = (0xE3, 0x6B, 0x2C)
INK = (0x1A, 0x24, 0x33)
GREY = (0x5B, 0x67, 0x75)


def set_run_font(run, size=11, bold=False, color=None, name="微软雅黑"):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)


def set_cell_bg(cell, color_hex: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)


def set_cell_margins(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for key, val in kwargs.items():
        node = OxmlElement(f"w:{key}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def add_p(doc, text, bold=False, size=11, color=None, align=None, space_after=8, space_before=0):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing = 1.35
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color or INK)
    return p


def add_h(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        set_run_font(run, size={0: 22, 1: 16, 2: 13, 3: 12}.get(level, 12), bold=True, color=NAVY)
    h.paragraph_format.space_before = Pt(14 if level == 1 else 10)
    h.paragraph_format.space_after = Pt(8)
    return h


def add_bullets(doc, items):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        run = p.add_run(it)
        set_run_font(run, size=11, color=INK)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.3


def add_table(doc, headers, rows, header_fill="071A2B"):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.autofit = True
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        r = p.add_run(h)
        set_run_font(r, size=10, bold=True, color=(255, 255, 255))
        set_cell_bg(cell, header_fill)
        set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = table.rows[ri + 1].cells[ci]
            cell.text = ""
            p = cell.paragraphs[0]
            r = p.add_run(str(val))
            set_run_font(r, size=9.5, color=INK)
            set_cell_bg(cell, "EEF3F6" if ri % 2 else "FFFFFF")
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
    doc.add_paragraph()
    return table


def build():
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "微软雅黑"
    style.font.size = Pt(11)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    for section in doc.sections:
        section.top_margin = Cm(2.2)
        section.bottom_margin = Cm(2.2)
        section.left_margin = Cm(2.4)
        section.right_margin = Cm(2.4)
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)

    add_p(doc, C.CONFIDENTIAL, bold=True, size=11, color=CYAN, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_p(doc, C.COMPANY, bold=True, size=14, color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_p(doc, C.DOC_TITLE, bold=True, size=22, color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    add_p(doc, C.TAGLINE, size=12, color=EMBER, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    add_p(
        doc,
        f"提交人：{C.AUTHOR}　　{C.DATE_STR}　　{C.VERSION}　　{C.HORIZON}",
        size=11,
        color=GREY,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=18,
    )

    add_h(doc, "一、为什么写这份设想", 1)
    add_p(
        doc,
        "东昇聚变是重投资、大装置型企业。政府事务岗如果按“多跑部门、多吃饭局”来理解，"
        "成本和招一个人相比并不划算；如果按专职拿地来理解，又会与公司主业错位——"
        "拿地、报建、施工完全可以交给专业合作方。",
    )
    add_p(
        doc,
        "因此，这份 90 天工作设想把岗位收束为高管参谋：把市—区关系、产业协同、落地风控"
        "做成可调用的三张地图，服务已经在推进的浦东落地沟通，同时把杨浦做成高校与人才窗口。"
        "政策写成路径和窗口，不写成收入或拿地承诺。",
    )
    add_p(doc, C.ROLE_ONE_LINER, bold=True)

    add_h(doc, "二、岗位理解：三张地图，一条边界", 1)
    for m in C.MAPS:
        add_h(doc, f"{m['name']}——{m['goal']}", 2)
        add_bullets(doc, m["points"])
    add_h(doc, "一条边界", 2)
    add_bullets(doc, C.BOUNDARY)

    add_h(doc, "三、上海产业落地的基本判断", 1)
    add_p(doc, C.LANDING_JUDGEMENT)
    add_h(doc, "选址比选", 2)
    add_table(doc, C.SITE_ROWS[0], C.SITE_ROWS[1:])
    add_h(doc, "选址原则", 2)
    add_bullets(doc, C.SITE_PRINCIPLES)

    add_h(doc, "四、90 天总目标与三件套", 1)
    add_p(doc, C.NINETY_GOAL, bold=True)
    for i, d in enumerate(C.DELIVERABLES, 1):
        add_p(doc, f"{i}. {d['name']}", bold=True, space_after=2)
        add_p(doc, f"内容：{d['what']}", space_after=2)
        add_p(doc, f"审定：{d['owner']}", color=GREY)

    add_h(doc, "五、三阶段安排", 1)
    for ph in C.PHASES:
        add_h(doc, f"{ph['key']}  {ph['span']}  {ph['name']}", 2)
        add_p(doc, ph["one"], bold=True)
        add_bullets(doc, ph["outcomes"])

    add_h(doc, "5.1 第一阶段明细（入局 · 对齐）", 2)
    for tr in C.PHASE1_TRACKS:
        add_h(doc, tr["title"], 3)
        add_bullets(doc, tr["items"])

    add_h(doc, "5.2 第二阶段部门动作（开渠 · 比选）", 2)
    add_table(doc, C.PHASE2_GOV[0], C.PHASE2_GOV[1:])
    add_h(doc, "现场核验只做六件事", 3)
    add_bullets(doc, C.PHASE2_SITE)

    add_h(doc, "5.3 第三阶段交卷（闭环）", 2)
    add_bullets(doc, C.PHASE3_ITEMS)

    add_h(doc, "六、用地、对赌与政策包", 1)
    add_p(
        doc,
        "工业用地长期低价，本质是政府用土地换企业后续税收。对东昇这种税收爬坡需要时间的装置企业，"
        "地拿得便宜不等于划算。本岗在 90 天内把指标翻译清楚，供决策使用，不代替合作方去谈地、报建、施工。",
    )
    add_table(doc, C.LAND_TYPES[0], C.LAND_TYPES[1:])
    add_bullets(doc, C.TAX_NOTES)
    add_h(doc, "政策包：能做与不能做", 2)
    add_table(doc, C.POLICY_PACK[0], C.POLICY_PACK[1:])

    add_h(doc, "七、融资协同", 1)
    add_p(doc, C.FINANCE_NOTE)
    add_bullets(doc, C.FINANCE_ACTIONS)

    add_h(doc, "八、周节奏与产出物", 1)
    add_table(doc, C.CADENCE[0], C.CADENCE[1:])
    add_h(doc, "分周计划", 2)
    add_table(doc, C.WEEKLY_PLAN[0], C.WEEKLY_PLAN[1:])
    add_h(doc, "产出物时间表", 2)
    add_table(doc, C.OUTPUT_LIST[0], C.OUTPUT_LIST[1:])

    add_h(doc, "九、政府关系作战图（90 天口径）", 1)
    add_table(doc, C.GOV_CONTACTS[0], C.GOV_CONTACTS[1:])

    add_h(doc, "十、风险与需要公司支持的事项", 1)
    add_table(doc, C.RISKS[0], C.RISKS[1:])
    add_h(doc, "需要公司打开的门", 2)
    add_bullets(doc, C.SUPPORT_NEEDED)

    add_h(doc, "十一、90 天结束时的交卷标准", 1)
    add_p(
        doc,
        "90 天不是把关系“搞熟”的口号期，而是让高管手里留下三样能继续往下干的东西：",
    )
    add_bullets(
        doc,
        [
            "谁已对接、谁该补、谁不要再去——写在《政府关系作战图》上。",
            "重资产主方案在浦东，杨浦做窗口——写在《落地比选与政策包》上。",
            "1–2 条可启动的政策路径、15–20 家机构短名单、下一季度节奏——写在《事项闭环清单》上。",
        ],
    )
    add_p(
        doc,
        "先懂装置，再出门；补位不抢位；政策不口头化。这是本岗对东昇聚变最有用的 90 天。",
        bold=True,
        space_before=8,
    )
    add_p(
        doc,
        f"{C.AUTHOR}　{C.DATE_STR}　{C.VERSION}　{C.CONFIDENTIAL}",
        size=10,
        color=GREY,
        align=WD_ALIGN_PARAGRAPH.RIGHT,
        space_before=18,
    )

    doc.save(OUT_FILE)
    print(f"已生成 {OUT_FILE}")
    return OUT_FILE


if __name__ == "__main__":
    build()
