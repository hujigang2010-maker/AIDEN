# -*- coding: utf-8 -*-
"""生成《复兴岛全球创客岛收官答卷大会》策划方案 Word 文档。"""
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

import content as C

OUT = Path(__file__).resolve().parent.parent / "deliverables"
OUT.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT / "复兴岛_全球创客岛_科创出海与具身智能国际大会_策划方案.docx"

NAVY = (0x0B, 0x3D, 0x91)
INK = (0x1B, 0x2A, 0x44)
GREY = (0x66, 0x66, 0x66)
ACCENT = (0x0A, 0x6E, 0x6A)  # 青绿，呼应江岛科创


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
        set_run_font(run, size={0: 22, 1: 16, 2: 13, 3: 11}.get(level, 11), bold=True, color=NAVY)
    return h


def add_bullets(doc, items):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        run = p.add_run(it)
        set_run_font(run, size=10.5, color=INK)
        p.paragraph_format.space_after = Pt(3)


def add_table(doc, headers, rows, header_fill="0B3D91"):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        r = p.add_run(h)
        set_run_font(r, size=10, bold=True, color=(255, 255, 255))
        set_cell_bg(cell, header_fill)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = table.rows[ri + 1].cells[ci]
            cell.text = ""
            p = cell.paragraphs[0]
            r = p.add_run(str(val))
            set_run_font(r, size=9.5, color=INK)
            if ri % 2 == 1:
                set_cell_bg(cell, "F2F6FA")
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
    add_p(doc, "杨浦 · 复兴岛", bold=True, size=14, color=ACCENT, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p(
        doc,
        C.PROJECT_NAME,
        bold=True,
        size=22,
        color=NAVY,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=8,
    )
    add_p(
        doc,
        C.PROJECT_FULL,
        bold=True,
        size=16,
        color=NAVY,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=8,
    )
    add_p(
        doc,
        C.PROJECT_SUBTITLE,
        size=12,
        color=GREY,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=8,
    )
    add_p(
        doc,
        "Slogan：百年复兴，智造再启航",
        bold=True,
        size=12,
        color=ACCENT,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=16,
    )
    add_p(
        doc,
        f"建议举办日期：{C.EVENT_DATE}",
        bold=True,
        size=12,
        color=ACCENT,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    add_p(
        doc,
        "（9 月 15 日前最靠近收官节点的开市 / 挂匾黄道吉日）",
        size=10.5,
        color=GREY,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=20,
    )
    add_p(doc, C.ONE_LINER, size=10.5, color=INK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=24)
    for _ in range(4):
        add_p(doc, "")
    add_p(doc, C.ORG_LINE, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p(doc, C.CO_ORG, size=10.5, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p(
        doc,
        f"版本：{C.VERSION}　　编制：活动策划组",
        size=10,
        color=GREY,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    doc.add_page_break()

    # ===== 目录 =====
    add_h(doc, "目录", level=1)
    toc = [
        "一、活动背景与战略意义",
        "二、历史文脉回顾与议题衍生",
        "三、市、区领导寄托与活动回应",
        "四、大上海 / 杨浦规划对齐与六大课题",
        "五、「船台·智能体」黑客松专章",
        "六、择日专章：为何选定 2026 年 9 月 12 日",
        "七、活动定位、目标与主题板块",
        "八、活动概览与规模结构",
        "九、场地建议（复兴岛）",
        "十、嘉宾阵容与邀约策略",
        "十一、国际会议厅 / 会客厅揭牌落位计划",
        "十二、详细议程安排",
        "十三、组织架构与执行倒排期",
        "十四、预算测算（示意）",
        "十五、预期成效与 KPI",
        "十六、风险预案与下一步行动",
        "附录：拟邀「一带一路」国家参考池",
    ]
    for t in toc:
        add_p(doc, t, size=11, color=INK, space_after=4)
    doc.add_page_break()

    # ===== 一、背景 =====
    add_h(doc, "一、活动背景与战略意义", level=1)
    add_p(
        doc,
        "本场活动不是一般性行业论坛，而是服务复兴岛领导政绩呈现、对接市级「量子城市」政策、"
        "完成「全球创客岛」阶段性收官叙事的重大型政治—产业综合活动。",
        color=INK,
    )
    for title, body in C.BACKGROUND:
        add_h(doc, title, level=2)
        add_p(doc, body, color=INK)
    add_h(doc, "政策迎合要点", level=2)
    add_bullets(doc, C.POLICY_HOOKS)
    add_h(doc, "本场活动要回答的四个问题", level=2)
    add_bullets(
        doc,
        [
            "百年「复兴」文脉如何转译为今日创客岛的主题与空间体验？",
            "自 2025 年 12 月启动以来，全球创客岛交出了什么阶段性答卷？",
            "杨浦如何把「量子城市」与市、区领导寄托做成可汇报实践？",
            "国际化与创客生态是否同时可见：领事、会客厅、黑客松都到位了吗？",
        ],
    )

    # ===== 二、历史文脉 =====
    add_h(doc, "二、历史文脉回顾与议题衍生", level=1)
    add_p(
        doc,
        "复兴岛是黄浦江内唯一封闭式内陆岛。由周家嘴浅滩经浚浦局吹填成岛，"
        "历经周家嘴岛、定海岛（一度昭和岛），抗战胜利后定名「复兴岛」，寓国家复兴。"
        "造船、仓储、渔业等工业文脉深厚；船台、塔吊、定海桥成为今日活化载体。",
        color=INK,
    )
    add_h(doc, "时间轴", level=2)
    add_table(doc, ["时期", "节点", "要点"], list(C.HISTORY_TIMELINE))
    add_h(doc, "关键文脉载体", level=2)
    add_table(doc, ["载体", "说明"], list(C.HISTORY_HERITAGE))
    add_h(doc, "衍生议题（回顾→当代转译）", level=2)
    for d in C.HISTORY_DERIVATIVES:
        add_h(doc, d["name"], level=3)
        add_bullets(doc, d["points"])

    # ===== 三、领导寄托 =====
    add_h(doc, "三、市、区领导寄托与活动回应", level=1)
    for block in C.LEADERSHIP:
        add_h(doc, block["who"], level=2)
        add_bullets(doc, block["points"])
    add_h(doc, "本场活动如何接住寄托", level=2)
    add_bullets(doc, C.LEADERSHIP_RESPONSE)

    # ===== 四、规划课题 =====
    add_h(doc, "四、大上海 / 杨浦规划对齐与六大课题", level=1)
    add_table(doc, ["规划维度", "对齐要点"], list(C.PLANNING_ALIGN))
    add_h(doc, "六大课题设计", level=2)
    add_table(doc, ["课题", "题目", "聚焦"], list(C.TOPIC_AGENDA))

    # ===== 五、黑客松 =====
    add_h(doc, "五、「船台·智能体」黑客松专章", level=1)
    add_p(doc, f"{C.HACKATHON['name']}　｜　{C.HACKATHON['slogan']}", bold=True, color=ACCENT)
    add_p(doc, f"窗口：{C.HACKATHON['window']}　｜　规模：{C.HACKATHON['scale']}", color=INK)
    add_p(doc, f"场地：{C.HACKATHON['venue']}", color=INK)
    add_h(doc, "为何必须做", level=2)
    add_bullets(doc, C.HACKATHON["why"])
    add_h(doc, "四大赛道", level=2)
    add_table(doc, ["赛道", "课题方向"], list(C.HACKATHON["tracks"]))
    add_h(doc, "规则要点", level=2)
    add_bullets(doc, C.HACKATHON["rules"])
    add_h(doc, "赛程", level=2)
    add_table(doc, ["时间", "环节"], list(C.HACKATHON["schedule"]))
    add_h(doc, "黑客松 KPI", level=2)
    add_bullets(doc, C.HACKATHON["kpis"])

    # ===== 六、择日 =====
    add_h(doc, "六、择日专章：为何选定 2026 年 9 月 12 日", level=1)
    add_p(
        doc,
        "收官节点要求活动不晚于 2026 年 9 月 15 日。"
        "同时需满足：宜开市（大会启动）、宜挂匾（揭牌落位）、宜立券/交易（签约）、宜出行（国际嘉宾到访）。",
        color=INK,
    )
    add_h(doc, "首选日期", level=2)
    add_table(
        doc,
        ["项目", "内容"],
        [
            ["公历 / 星期", C.HUANGLI["date"]],
            ["干支 / 值日", f"{C.HUANGLI['ganzhi']} · {C.HUANGLI['zhiri']}"],
            ["冲煞", C.HUANGLI["chong"]],
            ["宜", C.HUANGLI["yi"]],
            ["忌", C.HUANGLI["ji"]],
            ["择日理由", C.HUANGLI["why"]],
        ],
        header_fill="0A6E6A",
    )
    add_h(doc, "备选与规避", level=2)
    for b in C.DATE_BACKUP:
        add_p(doc, f"· {b['date']}：{b['note']}", color=INK, space_after=4)
    for a in C.DATE_AVOID:
        add_p(doc, f"· 规避：{a}", color=GREY, space_after=4)
    add_p(
        doc,
        "建议决策口径：主推 9 月 12 日（最靠近 9/15）；若区主要领导周六难以出席，则改 9 月 9 日（周三）。",
        bold=True,
        color=ACCENT,
    )

    # ===== 七、定位 =====
    add_h(doc, "七、活动定位、目标与主题板块", level=1)
    add_h(doc, "一句话定位", level=2)
    add_p(doc, C.ONE_LINER, color=INK)
    add_h(doc, "五大目标", level=2)
    add_bullets(doc, C.OBJECTIVES)
    add_h(doc, "四大主题板块（协同而非割裂）", level=2)
    for pillar in C.THEME_PILLARS:
        add_h(doc, f"{pillar['name']}（{pillar['tag']}）", level=3)
        add_bullets(doc, pillar["points"])

    # ===== 八、概览 =====
    add_h(doc, "八、活动概览与规模结构", level=1)
    add_table(doc, ["项目", "内容"], C.OVERVIEW)
    add_h(doc, "席位结构（200–300 人）", level=2)
    add_table(doc, ["席别", "人数", "组成"], C.SEAT_PLAN)

    # ===== 九、场地 =====
    add_h(doc, "九、场地建议（复兴岛）", level=1)
    add_p(doc, "必须在岛上办——这是政治象征与叙事完整性的底线。场地遴选标准：", color=INK)
    add_bullets(doc, C.VENUE_CRITERIA)
    for v in C.VENUE_OPTIONS:
        add_h(doc, v["name"], level=2)
        add_p(doc, v["desc"], color=INK)
        add_p(doc, f"容量：{v['capacity']}　　优势：{v['pros']}", color=GREY)

    # ===== 十、嘉宾 =====
    add_h(doc, "十、嘉宾阵容与邀约策略", level=1)
    add_p(
        doc,
        "政治规格与国际规格必须同时在场：区委区政府主要领导 + 东亚/东南亚驻沪总领事"
        "（暂按全部出席准备）亲自到场，才能形成「政绩答卷」所需的高位画面。",
        color=INK,
    )
    for g in C.GUEST_TIERS:
        add_h(doc, g["tier"], level=2)
        add_bullets(doc, g["targets"])
        add_p(doc, f"目标：{g['goal']}", color=GREY, space_after=8)
    add_h(doc, "东亚 + 东南亚驻沪总领事名单（暂按全部参加）", level=2)
    add_p(doc, C.CONSUL_INVITE_NOTE, color=GREY, space_after=8)
    add_table(
        doc,
        ["片区", "国家", "总领事", "英文名", "合作侧重", "现场安排"],
        [
            [
                c["region"],
                c["country"],
                c["name_zh"],
                c["name_en"],
                c["focus"],
                c["role"],
            ]
            for c in C.EAST_SE_ASIA_CONSULS
        ],
    )
    add_h(doc, "暂未在沪设总领事馆的国家/地区", level=2)
    add_bullets(doc, [f"{a}：{b}" for a, b in C.CONSUL_NO_SHANGHAI_CG])
    add_h(doc, "邀约策略要点", level=2)
    add_bullets(
        doc,
        [
            "外事路径：经区外办规范邀约；11 国全名单并联推进，揭牌国必须本人出席。",
            "领导路径：以「全球创客岛收官答卷 + 量子城市岛上实践」请示件报区主要领导。",
            "产业路径：以场景合作与入岛政策吸引头部企业负责人，而非仅市场部门出席。",
            "规模控制：贵宾席按全员预留；核心席质量优先，未确认到场可改书面致辞+副总领事。",
        ],
    )

    # ===== 十一、揭牌 =====
    add_h(doc, "十一、国际会议厅 / 会客厅揭牌落位计划", level=1)
    add_p(
        doc,
        "揭牌是本场活动区别于普通峰会的「硬成果」。建议采用「国家厅 + 片区厅 + 产业平台」组合，"
        "确保领导讲话与新闻通稿都有可指认的空间落位。",
        color=INK,
    )
    rows = [[u["name"], u["count"], u["form"], u["value"]] for u in C.UNVEILING]
    add_table(doc, ["落位类型", "数量建议", "形式", "价值"], rows)
    add_h(doc, "落位原则", level=2)
    add_bullets(doc, C.UNVEILING_PRINCIPLES)

    # ===== 十二、议程 =====
    add_h(doc, "十二、详细议程安排", level=1)
    add_p(
        doc,
        f"日程以 {C.EVENT_DATE} 全天为基准设计；上午聚焦「政治规格 + 揭牌签约」，"
        "下午聚焦「产业深度 + 国际对接」。",
        color=INK,
    )
    add_table(doc, ["时间", "环节", "内容要点", "责任"], C.AGENDA)

    # ===== 十三、组织与倒排 =====
    add_h(doc, "十三、组织架构与执行倒排期", level=1)
    add_h(doc, "组织架构", level=2)
    add_table(doc, ["组别", "职责"], C.ORG_STRUCTURE)
    add_h(doc, "倒排期（自即日起）", level=2)
    add_table(doc, ["节点", "关键任务"], C.TIMELINE)

    # ===== 十四、预算 =====
    add_h(doc, "十四、预算测算（示意）", level=1)
    add_table(doc, ["成本项", "预算(万元)", "说明"], C.BUDGET)
    add_p(doc, C.BUDGET_NOTE, color=GREY)

    # ===== 十五、KPI =====
    add_h(doc, "十五、预期成效与 KPI", level=1)
    add_table(doc, ["维度", "量化目标"], C.KPIS)

    # ===== 十六、风险 =====
    add_h(doc, "十六、风险预案与下一步行动", level=1)
    add_h(doc, "主要风险与对策", level=2)
    add_table(doc, ["风险", "对策", "等级"], C.RISKS)
    add_h(doc, "下一步行动（建议立即执行）", level=2)
    add_bullets(doc, C.NEXT_STEPS)

    # ===== 附录 =====
    add_h(doc, "附录：拟邀「一带一路」延伸国家参考池", level=1)
    add_p(
        doc,
        "东亚与东南亚驻沪总领事已按全部出席单独建档（见第十章）。"
        "下表为中东/中亚/中东欧等延伸参考池，视揭牌谈判与档期追加邀约。",
        color=INK,
    )
    add_table(doc, ["国家", "合作侧重"], C.BRI_COUNTRY_POOL)

    add_p(doc, "")
    add_p(
        doc,
        "（本稿为策划建议稿，黄历信息供择日参考；最终以区领导决策、外事报批与现场执行方案为准。）",
        size=9.5,
        color=GREY,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )

    doc.save(OUT_FILE)
    print(f"已生成：{OUT_FILE}")


if __name__ == "__main__":
    build()
