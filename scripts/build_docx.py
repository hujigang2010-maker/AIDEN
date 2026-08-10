# -*- coding: utf-8 -*-
"""生成 Word：上海人工智能产业展馆 · 对外汇报版。"""
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

import content as C

OUT = Path(__file__).resolve().parent.parent / "deliverables"
OUT.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT / "上海人工智能产业展馆_对外汇报方案.docx"

NAVY = (0x0B, 0x3D, 0x5C)
TEAL = (0x0F, 0x7A, 0x6E)
INK = (0x1A, 0x2B, 0x2E)
GREY = (0x5C, 0x6B, 0x70)


def set_run_font(run, size=10.5, bold=False, color=None, name="微软雅黑"):
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


def add_p(doc, text, bold=False, size=10.5, color=None, align=None, space_after=6):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    return p


def add_h(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        set_run_font(
            run,
            size={0: 22, 1: 16, 2: 13, 3: 11}.get(level, 11),
            bold=True,
            color=NAVY,
        )
    return h


def add_bullets(doc, items):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        run = p.add_run(it)
        set_run_font(run, size=10.5, color=INK)
        p.paragraph_format.space_after = Pt(3)


def add_table(doc, headers, rows, header_fill="0B3D5C"):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        r = p.add_run(h)
        set_run_font(r, size=9.5, bold=True, color=(255, 255, 255))
        set_cell_bg(cell, header_fill)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = table.rows[ri + 1].cells[ci]
            cell.text = ""
            p = cell.paragraphs[0]
            r = p.add_run(str(val))
            set_run_font(r, size=9, color=INK)
            if ri % 2 == 1:
                set_cell_bg(cell, "D7EBE6")
    doc.add_paragraph()
    return table


def build():
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "微软雅黑"
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.3)
        section.right_margin = Cm(2.3)

    # 封面
    for _ in range(2):
        add_p(doc, "")
    add_p(doc, "对外汇报方案", bold=True, size=12, color=TEAL,
          align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p(doc, C.PROJECT, bold=True, size=20, color=NAVY,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    add_p(doc, f"样板对标：{C.BENCHMARK}", bold=True, size=12, color=TEAL,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)
    add_p(doc, "租金与装修产业生态 · 多元盈利 · 支持单位协同", size=12,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=20)
    add_p(doc, f"版本：{C.VERSION}　　{C.DATE_STR}", size=11, color=GREY,
          align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p(doc, "汇报团队：陈晟 · 胡继刚 · 陈红苗", size=11, color=GREY,
          align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p(doc, "支持单位：" + "、".join(C.SUPPORT_ORGS), size=10, color=GREY,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    doc.add_page_break()

    add_h(doc, "目录", level=1)
    for t in [
        "一、项目概要",
        "二、对标加州计算机历史博物馆（一对一深化）",
        "三、租金的产业生态逻辑（重点）",
        "四、装修费的分层共担（重点）",
        "五、投资与赞助结构",
        "六、支持单位",
        "七、多元盈利模式",
        "八、政策与资金匹配",
        "九、选址、内容与推进",
        "十、风险与下一步",
    ]:
        add_p(doc, t, size=11, space_after=4)
    doc.add_page_break()

    add_h(doc, "一、项目概要", level=1)
    add_p(doc, C.EXEC_SUMMARY)
    add_p(doc, "核心议题：", bold=True)
    add_bullets(doc, C.CORE_CONCERNS)
    add_p(doc, "共识要点：", bold=True)
    for kp in C.KEY_POINTS:
        add_p(doc, kp["议题"], bold=True, color=TEAL, space_after=2)
        add_p(doc, kp["结论"], space_after=8)

    add_h(doc, "二、对标加州计算机历史博物馆（一对一深化）", level=1)
    add_p(doc, "样板馆基本画像：", bold=True)
    add_table(doc, ["维度", "要点"], [[k, v] for k, v in C.CHM_PROFILE.items()])
    add_p(doc, "一对一对照：", bold=True)
    add_table(doc, C.CHM_COMPARE[0], C.CHM_COMPARE[1:])
    add_p(doc, "上海本土化原则：", bold=True)
    add_bullets(doc, C.CHM_LOCALIZE)
    add_p(doc, "差异化：", bold=True)
    add_bullets(doc, C.DIFFERENTIATION)

    add_h(doc, "三、租金的产业生态逻辑（重点）", level=1)
    add_p(doc, C.RENT_ECOLOGY)
    add_table(
        doc,
        ["层级", "逻辑", "谁承担", "谈判点"],
        [[x["层级"], x["逻辑"], x["谁承担"], x["谈判点"]] for x in C.RENT_LAYERS],
    )

    add_h(doc, "四、装修费的分层共担（重点）", level=1)
    add_p(doc, C.FITOUT_ECOLOGY)
    add_table(
        doc,
        ["科目", "典型内容", "建议承担", "回收方式"],
        [[x["科目"], x["典型内容"], x["建议承担"], x["回收方式"]] for x in C.FITOUT_LAYERS],
    )
    add_p(doc, "落地步骤：", bold=True)
    add_table(doc, C.RENT_FITOUT_PLAYBOOK[0], C.RENT_FITOUT_PLAYBOOK[1:])
    add_p(doc, "经济假设（讨论稿）：", bold=True)
    add_table(doc, C.UNIT_ECONOMICS[0], C.UNIT_ECONOMICS[1:])

    add_h(doc, "五、投资与赞助结构", level=1)
    add_p(doc, C.INVEST_PRINCIPLE)
    add_table(
        doc,
        ["路径", "出资方式", "出资方画像", "回报机制", "优先级"],
        [
            [it["路径"], it["出资方式"], it["出资方画像"], it["回报机制"], it["优先级"]]
            for it in C.INVEST_PATHS
        ],
    )
    add_p(doc, "赞助层级：", bold=True)
    add_table(doc, C.SPONSOR_TIERS[0], C.SPONSOR_TIERS[1:])

    add_h(doc, "六、支持单位", level=1)
    add_table(doc, C.SUPPORT_ORG_ROLES[0], C.SUPPORT_ORG_ROLES[1:])
    add_p(
        doc,
        "拟通过合作备忘录明确联合课题、会员企业对接、活动联办与品牌互挂；"
        "支持单位重在生态赋能，不替代运营主体的启动金与经营责任。",
    )

    add_h(doc, "七、多元盈利模式", level=1)
    add_p(doc, C.PROFIT_LOGIC)
    add_table(
        doc,
        ["收入线", "描述", "延伸玩法", "占比", "里程碑"],
        [
            [r["收入线"], r["描述"], r["想象力延伸"], r["占比"], r["里程碑"]]
            for r in C.REVENUE_STREAMS
        ],
    )

    add_h(doc, "八、政策与资金匹配", level=1)
    add_table(
        doc,
        ["政策/抓手", "要点", "对项目价值", "落地步骤"],
        [
            [x["政策/抓手"], x["要点"], x["对项目价值"], x["落地步骤"]]
            for x in C.POLICY_SUPPORT
        ],
    )
    add_table(doc, C.POLICY_FUND_MATCH[0], C.POLICY_FUND_MATCH[1:])

    add_h(doc, "九、选址、内容与推进", level=1)
    add_table(doc, C.SITE_CANDIDATES[0], C.SITE_CANDIDATES[1:])
    add_table(doc, C.CONTENT_LAYERS[0], C.CONTENT_LAYERS[1:])
    add_table(doc, C.ORG_ROLES[0], C.ORG_ROLES[1:])
    add_table(doc, C.ROADMAP[0], C.ROADMAP[1:])

    add_h(doc, "十、风险与下一步", level=1)
    add_table(doc, C.RISKS[0], C.RISKS[1:])
    add_p(doc, "建议下一步：", bold=True)
    add_bullets(doc, C.NEXT_STEPS)

    add_p(doc, "")
    add_p(doc, "— 汇报材料结束 —", align=WD_ALIGN_PARAGRAPH.CENTER, color=GREY, size=10)
    add_p(
        doc,
        "配套：上海人工智能产业展馆_对外汇报方案.pptx　｜　上海人工智能产业展馆_投资盈利政策落地表.xlsx",
        align=WD_ALIGN_PARAGRAPH.CENTER, color=GREY, size=9,
    )

    doc.save(OUT_FILE)
    print(f"已生成: {OUT_FILE}")


if __name__ == "__main__":
    build()
