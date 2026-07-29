# -*- coding: utf-8 -*-
"""生成《复旦杨浦科创生态共建计划》路演 PPT。

视觉：复旦深蓝底 + 金色点缀；16:9。
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

OUT = Path(__file__).resolve().parents[1] / "deliverables" / "复旦杨浦科创生态共建计划_路演PPT.pptx"

NAVY = RGBColor(0x0F, 0x2E, 0x5C)
NAVY_DEEP = RGBColor(0x08, 0x1C, 0x3A)
NAVY_MID = RGBColor(0x16, 0x45, 0x7F)
GOLD = RGBColor(0xC8, 0x96, 0x3E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT = RGBColor(0xF5, 0xF7, 0xFA)
GREY = RGBColor(0xC3, 0xCB, 0xD4)
INK = RGBColor(0x1A, 0x1D, 0x21)
MUTED = RGBColor(0x5C, 0x65, 0x70)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
FONT = "微软雅黑"

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
BLANK = prs.slide_layouts[6]


def add_slide(bg=NAVY_DEEP):
    s = prs.slides.add_slide(BLANK)
    fill = s.background.fill
    fill.solid()
    fill.fore_color.rgb = bg
    return s


def rect(slide, x, y, w, h, fill):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


def round_rect(slide, x, y, w, h, fill, radius=0.08):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    sh.adjustments[0] = radius
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


def textbox(slide, x, y, w, h, lines, *, size=18, color=WHITE, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.2):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    if isinstance(lines, str):
        lines = [lines]
    first = True
    for line in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = spacing
        if isinstance(line, tuple):
            t, opts = line
        else:
            t, opts = line, {}
        r = p.add_run()
        r.text = t
        r.font.name = FONT
        r.font.size = Pt(opts.get("size", size))
        r.font.bold = opts.get("bold", bold)
        r.font.color.rgb = opts.get("color", color)
    return tb


def footer(slide, page: str):
    textbox(slide, Inches(0.5), Inches(7.05), Inches(10), Inches(0.3), "复旦杨浦科创生态共建计划", size=11, color=GREY)
    textbox(slide, Inches(11.8), Inches(7.05), Inches(1), Inches(0.3), page, size=11, color=GOLD, align=PP_ALIGN.RIGHT)


def gold_bar(slide, x, y, w=Inches(0.55), h=Inches(0.06)):
    return rect(slide, x, y, w, h, GOLD)


# ---------- slides ----------
def slide_cover():
    s = add_slide(NAVY_DEEP)
    rect(s, Inches(0), Inches(0), Inches(0.18), SLIDE_H, GOLD)
    textbox(
        s,
        Inches(0.8),
        Inches(1.2),
        Inches(11),
        Inches(0.4),
        "联合发起　·　面向杨浦　·　链接复旦",
        size=14,
        color=GOLD,
        bold=True,
    )
    textbox(
        s,
        Inches(0.8),
        Inches(1.9),
        Inches(11),
        Inches(1.6),
        [
            ("复旦杨浦", {"size": 44, "bold": True, "color": WHITE}),
            ("科创生态共建计划", {"size": 44, "bold": True, "color": WHITE}),
        ],
        size=44,
        bold=True,
        spacing=1.15,
    )
    gold_bar(s, Inches(0.8), Inches(3.75))
    textbox(
        s,
        Inches(0.8),
        Inches(4.0),
        Inches(10),
        Inches(0.6),
        "以住房政策研究为纽带，连接科技企业、产业资源与高端客户群",
        size=16,
        color=GREY,
    )
    textbox(
        s,
        Inches(0.8),
        Inches(5.0),
        Inches(10),
        Inches(0.8),
        [
            ("联合组织单位", {"size": 12, "color": GOLD, "bold": True}),
            ("复旦大学住房政策研究中心　×　杨浦区科技企业联合会", {"size": 16, "color": WHITE, "bold": True}),
            ("秘书长 / 会长　联合发起", {"size": 12, "color": GREY}),
        ],
        spacing=1.35,
    )
    # keywords
    for i, kw in enumerate(["促招商", "强链条", "聚人群"]):
        x = Inches(0.8) + Inches(i * 1.7)
        round_rect(s, x, Inches(6.35), Inches(1.5), Inches(0.42), NAVY)
        textbox(s, x, Inches(6.38), Inches(1.5), Inches(0.4), kw, size=13, color=WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def slide_toc():
    s = add_slide(NAVY)
    textbox(s, Inches(0.7), Inches(0.45), Inches(6), Inches(0.4), "CONTENTS　本次汇报", size=14, color=GOLD, bold=True)
    textbox(s, Inches(0.7), Inches(0.95), Inches(8), Inches(0.6), "一场活动，三重价值", size=28, bold=True, color=WHITE)
    gold_bar(s, Inches(0.7), Inches(1.65))

    items = [
        ("01", "为什么是现在", "把「资源相遇」变成「生态共建」"),
        ("02", "三大目标", "促招商 · 强链条 · 聚人群"),
        ("03", "活动怎么做", "认识 → 对话 → 匹配 → 跟进"),
        ("04", "谁应该来到这里", "让有合作可能的人坐到同一张桌旁"),
        ("05", "活动信息与合作", "报名、对接与后续跟进"),
    ]
    for i, (n, t, d) in enumerate(items):
        y = Inches(2.05) + Inches(i * 0.85)
        round_rect(s, Inches(0.7), y, Inches(11.8), Inches(0.72), NAVY_DEEP)
        textbox(s, Inches(0.95), y + Inches(0.12), Inches(1), Inches(0.5), n, size=22, color=GOLD, bold=True)
        textbox(s, Inches(2.1), y + Inches(0.08), Inches(5), Inches(0.35), t, size=18, color=WHITE, bold=True)
        textbox(s, Inches(2.1), y + Inches(0.38), Inches(8), Inches(0.3), d, size=13, color=GREY)
    footer(s, "01")


def slide_why():
    s = add_slide(NAVY)
    textbox(s, Inches(0.7), Inches(0.4), Inches(8), Inches(0.35), "01　为什么是现在", size=13, color=GOLD, bold=True)
    textbox(s, Inches(0.7), Inches(0.8), Inches(11), Inches(0.55), "把「资源相遇」，变成「生态共建」", size=26, bold=True, color=WHITE)
    gold_bar(s, Inches(0.7), Inches(1.45))
    textbox(
        s,
        Inches(0.7),
        Inches(1.7),
        Inches(12),
        Inches(1.2),
        [
            "一个区域的科创活力，从来不是靠企业数量堆出来的。真正决定高度的，是连接的密度。",
            "杨浦拥有深厚的科创基因与产业腹地；复旦汇聚政策研究能力、学术智力与高层次人才网络。",
        ],
        size=15,
        color=GREY,
        spacing=1.4,
    )

    round_rect(s, Inches(0.7), Inches(3.15), Inches(12), Inches(1.15), NAVY_DEEP)
    textbox(s, Inches(1.0), Inches(3.3), Inches(11), Inches(0.3), "核心判断", size=12, color=GOLD, bold=True)
    textbox(
        s,
        Inches(1.0),
        Inches(3.65),
        Inches(11.2),
        Inches(0.5),
        "招商不是一次性的「引进」，而是围绕企业全生命周期，持续链接空间、政策、资本、技术、人才与客户。",
        size=15,
        color=WHITE,
    )

    cards = [
        ("落地连接", "企业与杨浦产业空间、政策服务、应用场景建立明确联系。"),
        ("合作线索", "企业与高校院所、投资机构、上下游伙伴形成可跟进的合作线索。"),
        ("长期圈层", "从「来参加一次活动」，走向「进入一个长期有价值的圈层」。"),
    ]
    for i, (t, d) in enumerate(cards):
        x = Inches(0.7) + Inches(i * 4.05)
        round_rect(s, x, Inches(4.6), Inches(3.85), Inches(1.85), NAVY_DEEP)
        textbox(s, x + Inches(0.25), Inches(4.8), Inches(3.3), Inches(0.35), f"0{i+1}　{t}", size=15, color=GOLD, bold=True)
        textbox(s, x + Inches(0.25), Inches(5.3), Inches(3.3), Inches(0.9), d, size=13, color=GREY)
    footer(s, "02")


def slide_goals():
    s = add_slide(NAVY)
    textbox(s, Inches(0.7), Inches(0.4), Inches(8), Inches(0.35), "02　三大目标", size=13, color=GOLD, bold=True)
    textbox(s, Inches(0.7), Inches(0.8), Inches(11), Inches(0.5), "我们只做三件事", size=26, bold=True, color=WHITE)
    textbox(s, Inches(0.7), Inches(1.4), Inches(11), Inches(0.35), "三件事，一条逻辑：让有价值的连接，反复发生。", size=14, color=GREY)
    gold_bar(s, Inches(0.7), Inches(1.85))

    goals = [
        ("01", "促进招商", "引资源", "让优质科技企业更了解杨浦、走进杨浦、落地杨浦。把空间、政策、场景、客户一次讲清楚。"),
        ("02", "促进生态链达成", "强协同", "让企业与产业伙伴、科研机构、资本方及应用场景高效对接，形成可跟进、可落地的合作线索。"),
        ("03", "促进高端客户群形成", "聚人群", "以高质量活动聚合高质量人群，沉淀长期互信与合作关系。"),
    ]
    for i, (n, t, tag, d) in enumerate(goals):
        y = Inches(2.2) + Inches(i * 1.45)
        round_rect(s, Inches(0.7), y, Inches(12), Inches(1.3), NAVY_DEEP)
        textbox(s, Inches(1.0), y + Inches(0.25), Inches(1), Inches(0.4), n, size=22, color=GOLD, bold=True)
        textbox(s, Inches(2.2), y + Inches(0.22), Inches(7), Inches(0.4), t, size=20, color=WHITE, bold=True)
        round_rect(s, Inches(10.6), y + Inches(0.25), Inches(1.6), Inches(0.38), NAVY_MID, radius=0.2)
        textbox(s, Inches(10.6), y + Inches(0.28), Inches(1.6), Inches(0.35), tag, size=12, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
        textbox(s, Inches(2.2), y + Inches(0.7), Inches(9.8), Inches(0.45), d, size=14, color=GREY)
    footer(s, "03")


def slide_goal_invest():
    s = add_slide(NAVY)
    textbox(s, Inches(0.7), Inches(0.4), Inches(10), Inches(0.35), "02.1　目标一", size=13, color=GOLD, bold=True)
    textbox(s, Inches(0.7), Inches(0.8), Inches(12), Inches(0.5), "促进招商：把「引进来」做成「留得住」", size=24, bold=True, color=WHITE)
    textbox(s, Inches(0.7), Inches(1.4), Inches(12), Inches(0.4), "不做单向招引。把空间、政策、场景、客户一次讲清楚——企业自己会算这笔账。", size=14, color=GREY)
    gold_bar(s, Inches(0.7), Inches(1.9))

    items = [
        ("空间", "产业载体、办公与中试场地匹配"),
        ("政策", "专项扶持、人才政策与申报辅导"),
        ("场景", "区域应用场景与首试首用机会"),
        ("客户", "本地及长三角高质量客户资源"),
    ]
    for i, (t, d) in enumerate(items):
        x = Inches(0.7) + Inches(i * 3.1)
        round_rect(s, x, Inches(2.4), Inches(2.9), Inches(2.8), NAVY_DEEP)
        textbox(s, x + Inches(0.2), Inches(2.7), Inches(2.5), Inches(0.5), f"0{i+1}", size=18, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
        textbox(s, x + Inches(0.2), Inches(3.3), Inches(2.5), Inches(0.5), t, size=22, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        textbox(s, x + Inches(0.2), Inches(4.0), Inches(2.5), Inches(0.9), d, size=13, color=GREY, align=PP_ALIGN.CENTER)

    round_rect(s, Inches(0.7), Inches(5.55), Inches(12), Inches(1.0), NAVY_DEEP)
    textbox(
        s,
        Inches(1.0),
        Inches(5.75),
        Inches(11.4),
        Inches(0.6),
        "招商的本质是让企业算清一笔账：在这里，成本、效率、客户与成长速度，是不是更优。",
        size=15,
        color=WHITE,
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    footer(s, "04")


def slide_goal_chain():
    s = add_slide(NAVY)
    textbox(s, Inches(0.7), Inches(0.4), Inches(10), Inches(0.35), "02.2　目标二", size=13, color=GOLD, bold=True)
    textbox(s, Inches(0.7), Inches(0.8), Inches(12), Inches(0.5), "促进生态链达成：让对接真正落到项目上", size=24, bold=True, color=WHITE)
    gold_bar(s, Inches(0.7), Inches(1.45))

    chain = ["科技企业", "产业伙伴", "科研机构", "资本方", "应用场景"]
    for i, name in enumerate(chain):
        x = Inches(0.7) + Inches(i * 2.45)
        round_rect(s, x, Inches(1.8), Inches(2.15), Inches(0.7), NAVY_DEEP)
        textbox(s, x, Inches(1.9), Inches(2.15), Inches(0.5), name, size=14, color=WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        if i < 4:
            textbox(s, x + Inches(2.0), Inches(1.9), Inches(0.4), Inches(0.5), "›", size=20, color=GOLD, align=PP_ALIGN.CENTER)

    steps = [
        ("01", "需求收集", "会前征集企业真实诉求，形成结构化需求清单。"),
        ("02", "定向匹配", "按行业、阶段与诉求做点对点安排，而非自由社交。"),
        ("03", "现场对接", "一对一洽谈 + 小范围闭门交流，当场推进到下一步。"),
        ("04", "会后跟进", "建立跟进台账与回访机制，让线索转化可追踪。"),
    ]
    for i, (n, t, d) in enumerate(steps):
        x = Inches(0.7) + Inches((i % 2) * 6.2)
        y = Inches(3.0) + Inches((i // 2) * 1.7)
        round_rect(s, x, y, Inches(5.9), Inches(1.5), NAVY_DEEP)
        textbox(s, x + Inches(0.3), y + Inches(0.25), Inches(1), Inches(0.4), n, size=18, color=GOLD, bold=True)
        textbox(s, x + Inches(1.3), y + Inches(0.28), Inches(4), Inches(0.35), t, size=18, color=WHITE, bold=True)
        textbox(s, x + Inches(1.3), y + Inches(0.75), Inches(4.2), Inches(0.55), d, size=13, color=GREY)
    footer(s, "05")


def slide_goal_circle():
    s = add_slide(NAVY)
    textbox(s, Inches(0.7), Inches(0.4), Inches(10), Inches(0.35), "02.3　目标三", size=13, color=GOLD, bold=True)
    textbox(s, Inches(0.7), Inches(0.8), Inches(12), Inches(0.5), "促进高端客户群形成：把一次见面变成长期关系", size=22, bold=True, color=WHITE)
    textbox(
        s,
        Inches(0.7),
        Inches(1.45),
        Inches(12),
        Inches(0.7),
        "圈层不是被组织出来的，是被反复见面沉淀出来的。以小而精控制规模与质量：宁可人少一点，也要让每一次见面都值得。",
        size=14,
        color=GREY,
    )
    gold_bar(s, Inches(0.7), Inches(2.25))

    kws = [
        ("专业", "有判断力的内容，而非泛泛的致辞"),
        ("开放", "愿意共享资源、认真对话的参与者"),
        ("精准", "按诉求定向匹配，不做无效社交"),
        ("持续", "形成年度节奏，长期滚动运营"),
    ]
    for i, (t, d) in enumerate(kws):
        x = Inches(0.7) + Inches(i * 3.1)
        round_rect(s, x, Inches(2.7), Inches(2.9), Inches(2.6), NAVY_DEEP)
        textbox(s, x + Inches(0.2), Inches(3.1), Inches(2.5), Inches(0.5), t, size=22, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
        textbox(s, x + Inches(0.25), Inches(3.8), Inches(2.4), Inches(1.1), d, size=14, color=GREY, align=PP_ALIGN.CENTER)

    round_rect(s, Inches(0.7), Inches(5.6), Inches(12), Inches(0.95), NAVY_DEEP)
    textbox(
        s,
        Inches(1.0),
        Inches(5.8),
        Inches(11.4),
        Inches(0.55),
        "高质量的人群，是高质量活动反复筛选出来的结果——这需要时间，也值得时间。",
        size=15,
        color=WHITE,
        align=PP_ALIGN.CENTER,
    )
    footer(s, "06")


def slide_how():
    s = add_slide(NAVY)
    textbox(s, Inches(0.7), Inches(0.4), Inches(10), Inches(0.35), "03　活动怎么做", size=13, color=GOLD, bold=True)
    textbox(s, Inches(0.7), Inches(0.8), Inches(12), Inches(0.5), "从一次相聚，到一条可持续的生态链", size=24, bold=True, color=WHITE)
    textbox(s, Inches(0.7), Inches(1.4), Inches(12), Inches(0.4), "强调「认识 → 对话 → 匹配 → 跟进」的完整闭环。", size=14, color=GREY)
    gold_bar(s, Inches(0.7), Inches(1.9))

    steps = [
        ("1", "主题引导", "围绕住房政策、科技产业、企业发展与区域机会，输出有价值的趋势判断。"),
        ("2", "企业呈现", "精选科技企业与产业项目，讲清技术、产品、场景与合作诉求。"),
        ("3", "资源对接", "组织企业、投资、科研、园区、服务机构与高端客户进行定向交流。"),
        ("4", "持续跟进", "沉淀需求清单、合作线索与后续拜访机制，让活动成果可追踪。"),
    ]
    for i, (n, t, d) in enumerate(steps):
        y = Inches(2.25) + Inches(i * 1.05)
        round_rect(s, Inches(0.7), y, Inches(12), Inches(0.9), NAVY_DEEP)
        textbox(s, Inches(1.0), y + Inches(0.2), Inches(0.7), Inches(0.5), n, size=22, color=GOLD, bold=True)
        textbox(s, Inches(1.9), y + Inches(0.15), Inches(3), Inches(0.35), t, size=18, color=WHITE, bold=True)
        textbox(s, Inches(1.9), y + Inches(0.5), Inches(10.2), Inches(0.35), d, size=13, color=GREY)
    footer(s, "07")


def slide_units():
    s = add_slide(NAVY)
    textbox(s, Inches(0.7), Inches(0.4), Inches(10), Inches(0.35), "03.1　活动单元", size=13, color=GOLD, bold=True)
    textbox(s, Inches(0.7), Inches(0.8), Inches(12), Inches(0.5), "四个单元，各司其职", size=26, bold=True, color=WHITE)
    gold_bar(s, Inches(0.7), Inches(1.45))

    units = [
        ("01", "主旨分享", "复旦智力与杨浦实践的交汇", "40 分钟 · 1–2 位主讲"),
        ("02", "圆桌对话", "科技企业如何在区域中找到成长坐标", "45 分钟 · 4–5 位嘉宾"),
        ("03", "项目路演", "企业把需求说清楚，把机会讲明白", "每家 8 分钟 · 5–6 家"),
        ("04", "闭门交流", "面向重点企业与高端客户的定向洽谈", "限额邀请制"),
    ]
    for i, (n, t, d, meta) in enumerate(units):
        x = Inches(0.7) + Inches((i % 2) * 6.2)
        y = Inches(1.9) + Inches((i // 2) * 2.3)
        round_rect(s, x, y, Inches(5.9), Inches(2.05), NAVY_DEEP)
        textbox(s, x + Inches(0.35), y + Inches(0.3), Inches(1), Inches(0.4), n, size=20, color=GOLD, bold=True)
        textbox(s, x + Inches(1.4), y + Inches(0.32), Inches(4), Inches(0.4), t, size=20, color=WHITE, bold=True)
        textbox(s, x + Inches(0.35), y + Inches(0.95), Inches(5.2), Inches(0.4), d, size=14, color=GREY)
        textbox(s, x + Inches(0.35), y + Inches(1.4), Inches(5.2), Inches(0.35), meta, size=13, color=GOLD)
    footer(s, "08")


def slide_who():
    s = add_slide(NAVY)
    textbox(s, Inches(0.7), Inches(0.35), Inches(10), Inches(0.3), "04　谁应该来到这里", size=13, color=GOLD, bold=True)
    textbox(s, Inches(0.7), Inches(0.7), Inches(12), Inches(0.45), "让真正有合作可能的人，坐到同一张桌旁", size=24, bold=True, color=WHITE)
    gold_bar(s, Inches(0.7), Inches(1.25))

    audiences = [
        ("科技企业", "寻找落地空间、产业协同、政策服务、投融资与高端客户。"),
        ("投资与金融机构", "发现具有成长潜力的项目，建立长期观察与服务关系。"),
        ("高校院所与科研团队", "推动技术成果、人才资源与企业应用场景有效连接。"),
        ("产业链伙伴", "围绕供应链、客户、渠道、技术和服务展开合作。"),
        ("园区与专业服务机构", "提供空间、人才、法务、财税、品牌等综合支持。"),
        ("高端客户群", "以专业判断与产业视野，参与更有价值的生态网络。"),
    ]
    for i, (t, d) in enumerate(audiences):
        x = Inches(0.7) + Inches((i % 3) * 4.1)
        y = Inches(1.55) + Inches((i // 3) * 2.15)
        round_rect(s, x, y, Inches(3.9), Inches(1.95), NAVY_DEEP)
        textbox(s, x + Inches(0.25), y + Inches(0.3), Inches(3.4), Inches(0.4), t, size=16, color=GOLD, bold=True)
        textbox(s, x + Inches(0.25), y + Inches(0.85), Inches(3.4), Inches(0.9), d, size=13, color=GREY)

    textbox(
        s,
        Inches(0.7),
        Inches(6.0),
        Inches(12),
        Inches(0.6),
        "我们寻找的伙伴：愿意开放资源、认真对话、长期合作——把「单点机会」变成「生态增量」。",
        size=14,
        color=WHITE,
    )
    footer(s, "09")


def slide_gains():
    s = add_slide(NAVY)
    textbox(s, Inches(0.7), Inches(0.4), Inches(10), Inches(0.35), "04.1　参与收获", size=13, color=GOLD, bold=True)
    textbox(s, Inches(0.7), Inches(0.8), Inches(12), Inches(0.5), "参与者将获得什么", size=26, bold=True, color=WHITE)
    gold_bar(s, Inches(0.7), Inches(1.45))

    gains = [
        ("01", "区域理解", "获得区域机会与政策环境的第一手理解，判断更准，决策更快。"),
        ("02", "合作机会", "与目标伙伴深入交流、验证需求，形成可执行的合作方案。"),
        ("03", "长期平台", "进入由复旦智力、杨浦产业与企业家网络共同支撑的连接平台。"),
    ]
    for i, (n, t, d) in enumerate(gains):
        x = Inches(0.7) + Inches(i * 4.1)
        round_rect(s, x, Inches(1.9), Inches(3.9), Inches(3.3), NAVY_DEEP)
        textbox(s, x + Inches(0.3), Inches(2.3), Inches(3.3), Inches(0.4), n, size=20, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
        textbox(s, x + Inches(0.3), Inches(2.9), Inches(3.3), Inches(0.5), t, size=22, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        textbox(s, x + Inches(0.35), Inches(3.7), Inches(3.2), Inches(1.2), d, size=14, color=GREY, align=PP_ALIGN.CENTER)

    round_rect(s, Inches(0.7), Inches(5.55), Inches(12), Inches(1.0), NAVY_DEEP)
    textbox(
        s,
        Inches(1.0),
        Inches(5.75),
        Inches(11.4),
        Inches(0.6),
        "我们不做「来了、听了、散了」的活动；我们做「认识了、对上了、跟进了」的连接。",
        size=15,
        color=WHITE,
        align=PP_ALIGN.CENTER,
    )
    footer(s, "10")


def slide_info():
    s = add_slide(NAVY)
    textbox(s, Inches(0.7), Inches(0.4), Inches(10), Inches(0.35), "05　活动信息", size=13, color=GOLD, bold=True)
    textbox(s, Inches(0.7), Inches(0.8), Inches(12), Inches(0.5), "报名与合作咨询", size=26, bold=True, color=WHITE)
    gold_bar(s, Inches(0.7), Inches(1.45))

    left = [
        ("活动时间", "【待定】"),
        ("活动地点", "【待定】"),
        ("活动形式", "主旨分享 · 圆桌对话 · 项目路演 · 闭门交流"),
        ("参与规模", "【待定】人（限额，需审核）"),
    ]
    for i, (k, v) in enumerate(left):
        y = Inches(1.8) + Inches(i * 0.95)
        round_rect(s, Inches(0.7), y, Inches(7.3), Inches(0.85), NAVY_DEEP)
        textbox(s, Inches(1.0), y + Inches(0.12), Inches(2.2), Inches(0.3), k, size=13, color=GOLD, bold=True)
        textbox(s, Inches(1.0), y + Inches(0.42), Inches(6.7), Inches(0.35), v, size=15, color=WHITE)

    round_rect(s, Inches(8.3), Inches(1.8), Inches(4.3), Inches(3.7), NAVY_DEEP)
    textbox(s, Inches(8.6), Inches(2.1), Inches(3.7), Inches(0.35), "联系方式", size=14, color=GOLD, bold=True)
    contacts = [
        ("联系人", "【姓名】"),
        ("电话", "【电话】"),
        ("微信", "【微信】"),
        ("公众号", "复联会"),
    ]
    for i, (k, v) in enumerate(contacts):
        y = Inches(2.65) + Inches(i * 0.65)
        textbox(s, Inches(8.6), y, Inches(1.5), Inches(0.3), k, size=13, color=GREY)
        textbox(s, Inches(10.1), y, Inches(2.2), Inches(0.3), v, size=15, color=WHITE, bold=True)

    textbox(
        s,
        Inches(0.7),
        Inches(5.85),
        Inches(12),
        Inches(0.6),
        "参与对象：科技企业｜投资机构｜高校院所｜产业伙伴｜园区及专业服务机构｜高端客户",
        size=14,
        color=GREY,
    )
    footer(s, "11")


def slide_end():
    s = add_slide(NAVY_DEEP)
    rect(s, Inches(0), Inches(0), Inches(0.18), SLIDE_H, GOLD)
    textbox(
        s,
        Inches(0.9),
        Inches(2.0),
        Inches(11.5),
        Inches(1.6),
        [
            ("资源不会自己相遇，", {"size": 28, "bold": True, "color": WHITE}),
            ("合作不会自动发生。", {"size": 28, "bold": True, "color": WHITE}),
        ],
        spacing=1.3,
    )
    textbox(
        s,
        Inches(0.9),
        Inches(3.7),
        Inches(11),
        Inches(0.6),
        "但只要把对的人放在一起，剩下的事情，会比想象中快。",
        size=16,
        color=GREY,
    )
    gold_bar(s, Inches(0.9), Inches(4.5))
    textbox(s, Inches(0.9), Inches(4.85), Inches(11), Inches(0.45), "复旦杨浦科创生态共建计划", size=20, bold=True, color=WHITE)
    textbox(
        s,
        Inches(0.9),
        Inches(5.45),
        Inches(11),
        Inches(0.7),
        [
            "复旦大学住房政策研究中心　｜　杨浦区科技企业联合会",
            "关注「复联会」，获取活动信息与合作机会",
        ],
        size=14,
        color=GOLD,
        spacing=1.4,
    )


def build():
    slide_cover()
    slide_toc()
    slide_why()
    slide_goals()
    slide_goal_invest()
    slide_goal_chain()
    slide_goal_circle()
    slide_how()
    slide_units()
    slide_who()
    slide_gains()
    slide_info()
    slide_end()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUT)
    print(f"已生成：{OUT}")


if __name__ == "__main__":
    build()
