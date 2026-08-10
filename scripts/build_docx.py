# -*- coding: utf-8 -*-
"""生成 Word：会议纪要 + 文字版可执行落地方案。"""
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

import content as C

OUT = Path(__file__).resolve().parent.parent / "deliverables"
OUT.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT / "上海AI博物馆_会议纪要与可执行方案.docx"

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

    # ===== 封面 =====
    for _ in range(2):
        add_p(doc, "")
    add_p(doc, "工作会议材料", bold=True, size=12, color=TEAL,
          align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p(doc, C.PROJECT, bold=True, size=22, color=NAVY,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    add_p(doc, "会议纪要与可执行落地方案", bold=True, size=16, color=NAVY,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    add_p(
        doc,
        "核心聚焦：投资方怎么投 · 项目怎么赚 · 政策支持怎么落地",
        bold=True, size=12, color=TEAL,
        align=WD_ALIGN_PARAGRAPH.CENTER, space_after=20,
    )
    add_p(doc, f"版本：{C.VERSION}　　日期：{C.DATE_STR}", size=11,
          color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p(doc, "资料来源：得到大脑分享笔记（两场）", size=10, color=GREY,
          align=WD_ALIGN_PARAGRAPH.CENTER)
    for u in C.SOURCE_NOTES:
        add_p(doc, u, size=9, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    doc.add_page_break()

    # ===== 目录 =====
    add_h(doc, "目录", level=1)
    toc = [
        "第一部分　双场会议纪要",
        "　　一、会议概要",
        "　　二、分议题结论",
        "　　三、待办事项",
        "第二部分　可执行落地方案（文字版）",
        "　　一、项目定位与差异化",
        "　　二、投资方如何投资（重点）",
        "　　三、盈利模式如何赚钱（重点）",
        "　　四、政策性支持与扶持基金（重点）",
        "　　五、场地与空间经济",
        "　　六、展陈与资源",
        "　　七、组织分工与推进节奏",
        "　　八、风险与下一步",
    ]
    for t in toc:
        add_p(doc, t, size=11, space_after=4)
    doc.add_page_break()

    # ===== 第一部分 纪要 =====
    add_h(doc, "第一部分　双场会议纪要", level=1)
    add_h(doc, "一、会议概要", level=2)
    add_p(doc, C.MEETING_SUMMARY)
    add_p(doc, "会议元信息：", bold=True)
    add_table(
        doc,
        ["场次", "时间", "时长", "人数", "类型"],
        [[m["场次"], m["时间"], m["时长"], m["人数"], m["类型"]] for m in C.MEETING_META],
    )
    add_p(doc, "本方案提炼时特别关注的三点：", bold=True)
    add_bullets(doc, C.CORE_CONCERNS)

    add_h(doc, "二、分议题结论", level=2)
    for kp in C.MEETING_KEY_POINTS:
        add_p(doc, kp["议题"], bold=True, color=TEAL, space_after=2)
        add_p(doc, kp["结论"], space_after=8)

    add_h(doc, "三、待办事项", level=2)
    add_table(
        doc,
        ["责任人", "事项", "优先级", "时间要求"],
        C.TODOS,
    )
    add_p(
        doc,
        "说明：场次一后半段为地铁出行与私人事务，与项目无关内容未写入本纪要。",
        size=9, color=GREY,
    )
    doc.add_page_break()

    # ===== 第二部分 方案 =====
    add_h(doc, "第二部分　可执行落地方案（文字版）", level=1)
    add_p(
        doc,
        "本部分把会议共识整理为可路演、可谈判、可排期的执行文本。"
        "阅读顺序建议：先看投资与盈利、再看政策匹配，最后落到场地谈判与90天节奏。",
    )

    add_h(doc, "一、项目定位与差异化", level=2)
    add_p(
        doc,
        "定位：对标美国计算机历史博物馆，打造世界人工智能大会闭幕后的365天常态化"
        "AI展示与孵化平台；覆盖芯片算力、技术应用、具身/移动AI全产业链；"
        "抢占「AI博物馆」名称IP，并形成可向其他城市输出的标准模板。",
    )
    add_p(doc, "差异化要点：", bold=True)
    add_bullets(doc, C.DIFFERENTIATION)

    add_h(doc, "二、投资方如何投资（重点）", level=2)
    add_p(doc, C.INVEST_PRINCIPLE)
    add_p(
        doc,
        "实务上，把「投资」拆成可组合的七条路径：场地方出地出装修设备、主赞助出核心费用、"
        "金牌赞助出展席与捐赠、政府出补贴与扶持基金、科协出名义与专项、孵化跟投出股权资金、"
        "海外伙伴出内容授权。P0是前四条中的场地、主赞助、政府基金、科协专项——没有它们，"
        "项目无法开工。",
    )
    add_table(
        doc,
        ["路径", "出资方式", "出资方画像", "回报机制", "优先级"],
        [
            [it["路径"], it["出资方式"], it["出资方画像"], it["回报机制"], it["优先级"]]
            for it in C.INVEST_PATHS
        ],
    )
    add_p(doc, "赞助层级（产品化，便于企业立项）：", bold=True)
    add_table(doc, C.SPONSOR_TIERS[0], C.SPONSOR_TIERS[1:])
    add_p(doc, "关键落地动作：", bold=True)
    add_bullets(doc, [
        "出具《场地与出资需求书》（免租10–20年 + 约5000万设备装修 + 分成公式）。",
        "华润等主赞助：一页纸权益包 + 对齐对方科创KPI的立项叙述。",
        "头部AI企业：捐赠协议 + 「AI行业历史席位」权益清单 + 内容更新SLA。",
        "科协：正式合作/请示文件，锁定名称与主管关系。",
        "政府：纳入区重点项目库，匹配租金补贴与专项申报窗口。",
    ])

    add_h(doc, "三、盈利模式如何赚钱（重点）", level=2)
    add_p(doc, C.PROFIT_LOGIC)
    add_p(
        doc,
        "建设期尽量做到项目方零出资本金；开馆后利润不依赖门票，而依赖经营性现金流与"
        "资本性收益。六条收入线可并行，但启动顺序建议为："
        "租金运营 → 研学接待 → 认证服务 → 孵化跟投 → 模板输出 → 海外内容。",
    )
    add_table(
        doc,
        ["收入线", "描述", "稳态占比", "里程碑"],
        [
            [r["收入线"], r["描述"], r["目标占比(稳态)"], r["里程碑"]]
            for r in C.REVENUE_STREAMS
        ],
    )
    add_p(doc, "空间与经济假设（讨论稿，待预算细化）：", bold=True)
    add_table(doc, C.UNIT_ECONOMICS[0], C.UNIT_ECONOMICS[1:])

    add_h(doc, "四、政策性支持与扶持基金（重点）", level=2)
    add_p(
        doc,
        "政策不是「背景介绍」，而是可转化的客流、场地与资金。优先级最高的三件事："
        "①把教育部研学/教师培训要求变成进馆采购；"
        "②用科协牵头拿下名称与合法性；"
        "③用地方政府盘活闲置物业政策谈下免租和配套资金，再叠加各类专项基金。",
    )
    add_table(
        doc,
        ["政策/抓手", "要点", "对项目价值", "落地步骤"],
        [
            [x["政策/抓手"], x["要点"], x["对项目价值"], x["落地步骤"]]
            for x in C.POLICY_SUPPORT
        ],
    )
    add_p(doc, "扶持资金类型与用途匹配：", bold=True)
    add_table(doc, C.POLICY_FUND_MATCH[0], C.POLICY_FUND_MATCH[1:])
    add_p(
        doc,
        "执行口诀：先签场地框架（免租+配套资金）→ 同步科协名义 → 再按窗口填报专项；"
        "现金资助是加法，场地与设备封闭才是开工充要条件。",
        bold=True, color=TEAL,
    )

    add_h(doc, "五、场地与空间经济", level=2)
    add_p(
        doc,
        "理想总面积约6000–10000㎡：1–2层约2000㎡作博物馆展陈，其余用于AI企业孵化办公与研学。"
        "对场地方的硬条件是：10–20年免费使用权，并出资约5000万用于装修与设备。",
    )
    add_table(doc, C.SITE_CANDIDATES[0], C.SITE_CANDIDATES[1:])

    add_h(doc, "六、展陈与资源", level=2)
    add_table(doc, C.EXHIBIT_LAYERS[0], C.EXHIBIT_LAYERS[1:])
    add_bullets(doc, [
        "美国侧已筛选约前20项核心展品，侧重基础理论；中国侧优先头部有收入企业，侧重应用。",
        "海外可对接谷歌等企业，并探索与李飞飞等人物的直播连线；需单独评估中美内容合规与接受度。",
        "国内头部大模型企业多数在沪，适合集中拜访、批量谈展席捐赠。",
    ])

    add_h(doc, "七、组织分工与推进节奏", level=2)
    add_table(doc, C.ORG_ROLES[0], C.ORG_ROLES[1:])
    add_p(doc, "阶段节奏：", bold=True)
    add_table(doc, C.ROADMAP[0], C.ROADMAP[1:])
    add_p(
        doc,
        "方法论约束：先完成大框架与清晰预算，再进入动线与施工图；"
        "避免在资金未封闭时陷入详细设计空转。",
    )

    add_h(doc, "八、风险与下一步", level=2)
    add_table(doc, C.RISKS[0], C.RISKS[1:])
    add_p(doc, "建议立即执行的下一步：", bold=True)
    add_bullets(doc, [
        "陈院长发送极简文字材料，并持续补充美国侧资源清单。",
        "胡老师以本文档 / PPT / Excel 为底稿形成可外发路演包。",
        "国内团队并行启动：创智汇与复兴岛场地谈判、华润主赞助沟通、科协正式文件。",
        "邀请曾达明确生态合作角色；同步头部企业展席意向。",
        "指定政策专员做扶持基金窗口表，按申报期倒排材料。",
    ])

    add_p(doc, "")
    add_p(
        doc,
        "— 文档结束 —",
        align=WD_ALIGN_PARAGRAPH.CENTER, color=GREY, size=10,
    )
    add_p(
        doc,
        "配套文件：上海AI博物馆_可执行落地方案.pptx　｜　上海AI博物馆_投资盈利政策落地表.xlsx",
        align=WD_ALIGN_PARAGRAPH.CENTER, color=GREY, size=9,
    )

    doc.save(OUT_FILE)
    print(f"已生成: {OUT_FILE}")


if __name__ == "__main__":
    build()
