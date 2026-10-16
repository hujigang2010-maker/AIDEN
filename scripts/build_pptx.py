# -*- coding: utf-8 -*-
"""生成《潮起潮落 TIDELINE × 北外滩尚9·一滴水》综合合作方案 PPT。

用法：
    python3 scripts/build_pptx.py [输出路径.pptx]
"""

import os
import sys

from pptx import Presentation
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

import proposal_data as D
from ppt_kit import (
    FONT, GOLD, GOLD_DEEP, ICE, INK, LIGHT_GOLD, LINE, MIST, MUTED, NAVY,
    NAVY_MID, SAND, SLIDE_H, SLIDE_W, TEAL, WHITE,
    add_lines, add_page_slot, add_rect, add_table, add_text, blank, bullets,
    callout, card, footer, header, note, place_grid, rgb_of, stamp_pages,
)

PALE = RGBColor(0xD7, 0xE0, 0xE8)
PALE_GOLD = RGBColor(0xE6, 0xD3, 0xA1)


def content(prs, kicker, title):
    slide = blank(prs)
    header(slide, kicker, title)
    return slide


def hbar(slide, x, y, w, h, num, title, body, fill=WHITE):
    add_rect(slide, x, y, w, h, fill, LINE, 0.75)
    add_rect(slide, x, y, 0.07, h, GOLD)
    add_text(slide, x + 0.18, y, 0.62, h, num, 16, GOLD_DEEP, True, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, x + 0.86, y + 0.08, w - 1.05, 0.32, title, 15, NAVY, True)
    add_text(slide, x + 0.86, y + 0.40, w - 1.05, h - 0.50, body, 13, INK)


def section(prs, no, title, subtitle, focus=False):
    slide = blank(prs)
    add_rect(slide, 0, 0, 13.333, 7.5, NAVY)
    add_rect(slide, 0, 0, 0.12, 7.5, GOLD)
    add_text(slide, 0.7, 1.85, 11.5, 0.35, "第 " + no + " 部分", 16, GOLD, True)
    add_text(slide, 0.7, 2.35, 11.8, 1.1, title, 40, WHITE, True)
    add_rect(slide, 0.72, 3.6, 1.8, 0.035, GOLD)
    add_text(slide, 0.7, 3.85, 11.2, 1.1, subtitle, 18, PALE)
    if focus:
        add_rect(slide, 0.7, 5.25, 1.7, 0.42, GOLD)
        add_text(slide, 0.7, 5.25, 1.7, 0.42, "汇报重点", 14, NAVY, True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, 0.7, 6.7, 10, 0.3, "潮起潮落 TIDELINE  ×  北外滩尚9·一滴水", 12, PALE_GOLD)
    add_page_slot(slide, PALE_GOLD)
    return slide


def slide_cover(prs):
    s = blank(prs)
    add_rect(s, 0, 0, 13.333, 7.5, NAVY)
    add_rect(s, 0, 0, 13.333, 0.08, GOLD)
    add_text(s, 0.62, 0.38, 12, 0.28, "北外滩尚9 · 一滴水", 15, GOLD, True)
    add_text(s, 0.62, 1.15, 12, 0.32, "TIDELINE", 16, PALE_GOLD, True)
    add_text(s, 0.62, 1.55, 12, 1.05, "潮起潮落", 60, WHITE, True)
    add_text(s, 0.62, 2.75, 12, 0.7, D.SUBTITLE, 22, WHITE, True)
    add_rect(s, 0.64, 3.55, 2.0, 0.03, GOLD)
    add_text(s, 0.62, 3.75, 12, 0.7, D.POSITIONING, 16, PALE)
    facts = [
        ("拟合作场地", "北外滩尚9·一滴水"),
        ("拟举办时间", "2026年11月下旬"),
        ("项目类型", "主题展与商业合作"),
        ("方案性质", "研讨稿，附工作表"),
    ]
    for i, (k, v) in enumerate(facts):
        x = 0.62 + i * 3.12
        add_rect(s, x, 4.75, 2.98, 1.15, NAVY_MID)
        add_text(s, x + 0.16, 4.86, 2.66, 0.28, k, 12, GOLD, True)
        add_text(s, x + 0.16, 5.18, 2.66, 0.52, v, 15, WHITE, True)
    add_text(
        s, 0.62, 6.15, 12, 0.55,
        "客流、曝光、合作资源和经营收益，均为计划目标或合作设想，不代表已经兑现的结果。赞助金额为原稿建议档位。",
        13, PALE,
    )
    add_text(
        s, 0.62, 6.75, 12, 0.3,
        "配套表格：赞助权益、空间日程、团队案例、职责评估、待确认事项，以及一滴水收益测算表。",
        12, PALE_GOLD,
    )
    add_page_slot(s, PALE_GOLD)
    note(s, "开场先说明这是研讨稿。数字都是计划口径。重点章节是商业价值、长期业务和合作模式。")


def slide_three(prs):
    s = content(prs, "开篇  ·  给经营者", "先看清三件事")
    items = [
        ("参与的理由", "年轻文化消费正在进入婚礼、宴会和主题社交。一滴水可以用一次江景主题活动，接触未来的婚庆和聚会客户，并测试非核心档期怎么经营。"),
        ("能够得到什么", "本次争取场地使用和餐饮收入；中期积累年轻客群认知和可跟进的咨询；长期再验证主题婚礼、茶会、生日会、摄影和品牌包场。"),
        ("汇报时的口径", "客流、曝光、赞助金额、创作者资源和经营收益，都是计划目标或合作设想。审批、主办主体和活动日期尚未闭合。"),
    ]
    boxes = place_grid(items, 3, 1.18, 2.85)
    for x, y, w, h, (title, body) in boxes:
        card(s, x, y, w, h, title, body, title_size=18, body_size=15)
    callout(
        s, 0.42, 4.25, 12.50, 2.45,
        "建议把汇报时间主要留给三部分：核心商业价值、长期业务方向、收入与合作模式。\n"
        "传播、场地、团队和执行细节放在后面，供追问时展开。\n"
        "签约前请逐项闭合文末八件待确认事项。日期、主办主体、审批和头部资源，目前都还没有闭合。",
        size=16,
    )
    note(s, "这三件事回答老板最关心的问题：为什么听、能得到什么、哪些话现在不能说死。")


def slide_toc(prs):
    s = content(prs, "目录", "这份方案的九个部分")
    boxes = place_grid(D.TOC, 3, 1.12, 1.78, gap_y=0.12)
    for x, y, w, h, (no, title, desc) in boxes:
        focus = no in ("05", "06", "07")
        add_rect(s, x, y, w, h, SAND if focus else WHITE, LINE, 0.75)
        add_rect(s, x, y, w, 0.045, GOLD)
        add_text(s, x + 0.16, y + 0.16, 1.1, 0.36, no, 18, GOLD_DEEP, True)
        if focus:
            add_text(s, x + w - 1.15, y + 0.18, 0.95, 0.30, "重点", 12, GOLD_DEEP, True, align=PP_ALIGN.RIGHT)
        add_text(s, x + 0.16, y + 0.58, w - 0.32, 0.36, title, 16, NAVY, True)
        add_text(s, x + 0.16, y + 1.05, w - 0.32, 0.58, desc, 12, INK)


def slide_market(prs):
    s = content(prs, "01  ·  合作契机", "年轻消费正在重塑文化消费")
    add_text(
        s, 0.42, 1.08, 12.5, 0.72,
        "兴趣、文化、社交和情绪价值成为更多年轻人的消费理由。二次元、游戏、角色扮演、沉浸式体验、主题摄影和文化社交，正在从垂直圈层走向可经营的商业场景。",
        15, INK,
    )
    items = [
        ("兴趣社群", "消费者围绕特定作品、游戏角色、创作者或文化主题聚集，兴趣标签清楚，社群有凝聚力。"),
        ("线下体验", "他们重视参与感和沉浸式体验，愿意为情感认同、审美特色和社交价值支付费用。"),
        ("二次传播", "参与者会在小红书、B站、微博、抖音发布照片、摄影作品、短视频和活动攻略。"),
    ]
    boxes = place_grid(items, 3, 1.90, 3.15)
    for x, y, w, h, (title, body) in boxes:
        card(s, x, y, w, h, title, body, title_size=18, body_size=15)
    callout(
        s, 0.42, 5.25, 12.50, 1.55,
        "项目材料把杨浦复兴岛青年文化活动、静安大悦城二次元消费场景作为市场观察方向，认为这类活动具有客流吸引、消费组织和社交传播的可能。具体客流规模与商业效果尚需进一步核实。",
        size=14,
    )


def slide_traits(prs):
    s = content(prs, "01  ·  合作契机", "这一客群的四个特点")
    boxes = place_grid(D.TRAITS, 2, 1.15, 2.35, gap_y=0.14)
    for x, y, w, h, (title, body) in boxes:
        card(s, x, y, w, h, title, body, title_size=18, body_size=15)
    callout(
        s, 0.42, 6.15, 12.50, 0.82,
        "对高品质宴会场地来说，值得研究的是：怎样把这些年轻消费需求，与自身的空间、餐饮和婚庆服务接在一起。",
        size=14,
    )


def slide_pressure(prs):
    s = content(prs, "01  ·  合作契机", "传统宴会场地的新课题")
    add_rect(s, 0.42, 1.15, 6.05, 5.65, WHITE, LINE, 0.75)
    add_rect(s, 0.42, 1.15, 6.05, 0.045, GOLD)
    add_text(s, 0.62, 1.35, 5.7, 0.4, "传统业务仍然是基本盘", 18, NAVY, True)
    bullets(s, 0.62, 1.95, 5.6, 4.5, [
        "婚宴、商务会议、企业年会和各类庆典，一直是高品质宴会场地的重要业务来源。",
        "客户对空间品质、活动体验、主题设计和服务个性化程度，提出了更高要求。",
        "婚宴和商务会议存在季节性及档期差异。提高非核心档期的利用率、增加配套消费，是经营上的重要课题。",
        "年轻群体的婚庆观念正走向多元和个性化，婚礼被期待能体现兴趣、审美、情感和独特体验。",
    ], size=14, spacing=12)
    add_rect(s, 6.65, 1.15, 6.27, 5.65, WHITE, LINE, 0.75)
    add_rect(s, 6.65, 1.15, 6.27, 0.045, TEAL)
    add_text(s, 6.85, 1.35, 5.9, 0.4, "可以同时研究的新场景", 18, NAVY, True)
    bullets(s, 6.85, 1.95, 5.85, 4.5, D.VENUE_SCENES, size=15, spacing=10)


def slide_assets(prs):
    s = content(prs, "01  ·  合作契机", "一滴水已经具备的空间基础")
    boxes = place_grid(D.VENUE_ASSETS, 2, 1.12, 2.20, gap_y=0.12)
    for x, y, w, h, (title, body) in boxes:
        card(s, x, y, w, h, title, body, title_size=16, body_size=13)
    callout(
        s, 0.42, 5.80, 12.50, 1.15,
        "因此，本次合作的定位，是一次探索青年文化消费场景和新型宴会业务的商业合作。场地租赁是执行形式，合作内容还包括内容、客群、餐饮、婚庆和后续运营。",
        size=15,
    )


def slide_goals(prs):
    s = content(prs, "02  ·  定位与路径", "合作定位与五个目标")
    add_rect(s, 0.42, 1.10, 12.50, 0.62, NAVY)
    add_text(
        s, 0.58, 1.10, 12.2, 0.62,
        "高品质场地  ＋  青年文化内容  ＋  兴趣社群  ＋  特色餐饮  ＋  婚庆消费  ＋  长期运营",
        15, WHITE, True, anchor=MSO_ANCHOR.MIDDLE,
    )
    for i, (num, title, body) in enumerate(D.GOALS):
        hbar(s, 0.42, 1.88 + i * 0.96, 12.50, 0.88, num, title, body)


def slide_phases(prs):
    s = content(prs, "02  ·  定位与路径", "分三阶段推进，先试点再扩展")
    boxes = place_grid(D.PHASES, 3, 1.15, 4.15)
    for x, y, w, h, (stage, title, body) in boxes:
        card(s, x, y, w, h, title, body, kicker=stage, title_size=18, body_size=14)
    callout(
        s, 0.42, 5.50, 12.50, 1.30,
        "进入下一阶段的前提，是试点取得可接受的效果。首期任务是验证组织、客流、承载、传播和实际收益。缺少测算时，不在试点之前承诺长期免费使用场地。",
        size=15,
    )


def slide_event(prs):
    s = content(prs, "03  ·  首期活动", "首期活动名片")
    boxes = place_grid(D.EVENT_METRICS, 4, 1.10, 1.28, gap_x=0.12)
    for x, y, w, h, (num, label, note_text) in boxes:
        add_rect(s, x, y, w, h, WHITE, LINE, 0.75)
        add_rect(s, x, y, w, 0.045, GOLD)
        add_text(s, x + 0.12, y + 0.12, w - 0.24, 0.40, num, 16, NAVY, True)
        add_text(s, x + 0.12, y + 0.52, w - 0.24, 0.28, label, 12, GOLD_DEEP, True)
        add_text(s, x + 0.12, y + 0.82, w - 0.24, 0.36, note_text, 11, MUTED)
    rows = [("项目", "内容")] + [(k, v) for k, v in D.EVENT_FACTS]
    add_table(s, 0.42, 2.55, [2.3, 10.2], rows, font=12, row_h=0.38, header_h=0.36)
    callout(
        s, 0.42, 6.40, 12.50, 0.62,
        "原方案采用单一 IP 同人 Only 展，强调核心玩家的社群黏性。日期、人数和主办主体都要在场地条件确认后闭合。",
        size=13,
    )


def slide_contents(prs):
    s = content(prs, "03  ·  首期活动", "现场的八项核心内容")
    boxes = place_grid(D.CONTENTS, 4, 1.12, 2.75, gap_y=0.14)
    for x, y, w, h, (num, title, body) in boxes:
        card(s, x, y, w, h, title, body, kicker=num, title_size=15, body_size=12)


def slide_highlights(prs):
    s = content(prs, "03  ·  首期活动", "活动的六个亮点")
    boxes = place_grid(D.HIGHLIGHTS, 3, 1.15, 2.70, gap_y=0.14)
    for x, y, w, h, (title, body) in boxes:
        card(s, x, y, w, h, title, body, title_size=16, body_size=13)


def slide_audiences(prs):
    s = content(prs, "03  ·  首期活动与客群", "谁会来，以及未来可能消费什么")
    boxes = place_grid(D.AUDIENCES, 4, 1.10, 2.55)
    for x, y, w, h, (title, body) in boxes:
        card(s, x, y, w, h, title, body, title_size=15, body_size=12)
    add_rect(s, 0.42, 3.82, 12.50, 2.95, SAND)
    add_rect(s, 0.42, 3.82, 0.07, 2.95, GOLD)
    add_text(s, 0.68, 3.94, 12, 0.34, "对一滴水而言，价值不按这一次门票计算", 15, NAVY, True)
    bullets(s, 0.68, 4.35, 6.0, 1.7, D.FUTURE_CLIENTS[:4], size=13, spacing=3)
    bullets(s, 6.9, 4.35, 5.7, 1.7, D.FUTURE_CLIENTS[4:], size=13, spacing=3)
    add_text(
        s, 0.68, 6.15, 12.0, 0.48,
        "以上是策划中的目标客群判断。消费能力、男女比例和品类需求，需用调研和运营数据确认。单笔更高的可能，来自特色婚庆、宴会、团体包场和品牌活动。",
        12, INK,
    )


def slide_zones(prs):
    s = content(prs, "04  ·  主题场景", "三大主题区与色彩")
    boxes = place_grid(D.ZONES, 3, 1.10, 4.85)
    for x, y, w, h, zone in boxes:
        add_rect(s, x, y, w, h, WHITE, LINE, 0.75)
        add_rect(s, x, y, w, 0.045, GOLD)
        add_text(s, x + 0.16, y + 0.14, w - 0.32, 0.28, zone["name"], 12, GOLD_DEEP, True)
        add_text(s, x + 0.16, y + 0.42, w - 0.32, 0.36, zone["place"], 18, NAVY, True)
        add_text(s, x + 0.16, y + 0.82, w - 0.32, 0.28, zone["mood"], 12, MUTED)
        n = len(zone["colors"])
        chip_w = (w - 0.32) / n
        for i, (hex_color, name) in enumerate(zone["colors"]):
            cx = x + 0.16 + i * chip_w
            add_rect(s, cx, y + 1.22, 0.36, 0.28, rgb_of(hex_color), LINE, 0.5)
            add_text(s, cx, y + 1.52, chip_w - 0.04, 0.40, name, 10, INK)
        add_text(s, x + 0.16, y + 2.05, w - 0.32, 1.2, zone["set"], 13, INK)
        add_text(s, x + 0.16, y + 3.30, w - 0.32, 1.3, zone["use"], 13, NAVY, True)
    callout(
        s, 0.42, 6.12, 12.50, 0.85,
        "以上是原稿的创意设想。角色形象、标识、官方素材和衍生展示，需在完成知识产权合规审核后使用。",
        size=13,
    )


def slide_interactions(prs):
    s = content(prs, "04  ·  主题场景", "八个主题互动")
    boxes = place_grid(D.INTERACTIONS, 2, 1.12, 1.35, gap_y=0.10)
    for x, y, w, h, (num, title, body) in boxes:
        add_rect(s, x, y, w, h, WHITE, LINE, 0.75)
        add_rect(s, x, y, 0.62, h, NAVY)
        add_text(s, x, y, 0.62, h, num, 14, GOLD, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.76, y + 0.10, w - 0.90, 0.32, title, 15, NAVY, True)
        add_text(s, x + 0.76, y + 0.46, w - 0.90, 0.78, body, 12, INK)


def slide_values(prs):
    s = content(prs, "05  ·  核心商业价值  ·  汇报重点", "一滴水的八项商业价值")
    boxes = place_grid(D.VALUES, 4, 1.12, 2.75, gap_y=0.14)
    for x, y, w, h, (num, title, body) in boxes:
        card(s, x, y, w, h, title, body, kicker=num, title_size=16, body_size=13)
    note(s, "这一页是全篇地图。后面四页分别展开婚庆、档期产品、获客，以及餐饮、夜间和品牌资产。")


def slide_wedding(prs):
    s = content(prs, "05  ·  核心商业价值  ·  汇报重点", "提前进入新一代婚庆市场")
    add_text(
        s, 0.42, 1.08, 12.5, 0.78,
        "新一代消费者进入婚恋和家庭生活阶段后，婚礼偏好更个性化、主题化和情感化。传统婚礼围绕规模、流程和档次展开；其中一部分年轻人，同时看重兴趣、共同记忆和独特体验。",
        14, INK,
    )
    add_rect(s, 0.42, 1.95, 7.55, 3.35, WHITE, LINE, 0.75)
    add_text(s, 0.58, 2.08, 7.2, 0.32, "可能出现的婚礼表达", 15, NAVY, True)
    bullets(s, 0.58, 2.48, 7.2, 2.65, D.WEDDING_EXAMPLES, size=13, spacing=4)
    add_rect(s, 8.15, 1.95, 4.77, 3.35, SAND)
    add_rect(s, 8.15, 1.95, 4.77, 0.045, GOLD)
    add_text(s, 8.32, 2.15, 4.45, 0.7, "高端婚礼定位保持不变", 16, NAVY, True)
    add_text(
        s, 8.32, 2.95, 4.45, 2.1,
        "本合作是在传统婚礼产品之外，增加一个新的细分市场。一滴水已有高品质宴会空间、餐饮服务和江景，适合用来研究主题婚礼和沉浸式庆典。",
        14, INK,
    )
    callout(
        s, 0.42, 5.48, 12.50, 1.32,
        "本次活动用来集中接触这一年轻文化消费群体，并了解他们的真实需求。建议现场适度设置特色婚礼展示、江景宴会介绍和官方咨询入口，让到场体验留下可以跟进的业务线索。",
        size=14,
    )
    note(s, "老板若担心活动拉低婚宴调性，用右侧这句话回应：原有高端婚礼继续做，这里是新增细分。")


def slide_offpeak(prs):
    s = content(prs, "05  ·  核心商业价值  ·  汇报重点", "九类可以研究的档期产品")
    rows = [("业务类型", "主要消费内容", "对场地方的价值")] + [row for row in D.OFFPEAK]
    add_table(s, 0.42, 1.12, [3.1, 5.2, 4.2], rows, font=13, row_h=0.46, header_h=0.42)
    callout(
        s, 0.42, 6.20, 12.50, 0.78,
        "工作日、晚间和非热门婚期，都有机会用适当的主题活动提高利用率。产品按经营数据再做定价、标准化和套餐化。表内是方向，不是已经在售的清单。",
        size=13,
    )


def slide_acquisition(prs):
    s = content(prs, "05  ·  核心商业价值  ·  汇报重点", "多一条精准到店的获客路径")
    add_text(
        s, 0.42, 1.08, 12.5, 0.85,
        "婚宴客户通常来自婚庆公司、广告平台、转介绍和其他商业渠道。二次元活动可以再增加一条路径：专业团队用社群，把有明确兴趣的年轻人带进场地。他们看过建筑和空间之后，未来办婚礼、生日、纪念、聚会或摄影时，一滴水可以进入选择范围。",
        14, INK,
    )
    for i, (num, title, body) in enumerate(D.ACQUISITION):
        y = 2.05 + i * 0.78
        add_rect(s, 0.42, y, 12.50, 0.70, WHITE, LINE, 0.75)
        add_rect(s, 0.42, y, 0.70, 0.70, NAVY)
        add_text(s, 0.42, y, 0.70, 0.70, num, 14, GOLD, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 1.28, y, 2.5, 0.70, title, 15, NAVY, True, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 3.85, y, 8.85, 0.70, body, 14, INK, anchor=MSO_ANCHOR.MIDDLE)
    callout(
        s, 0.42, 6.05, 12.50, 0.92,
        "咨询登记需要符合个人信息保护要求，并形成可以持续跟进的潜在业务线索。特色婚庆、宴会、团体包场和品牌活动，可能带来较高的单笔经营收入。",
        size=14,
    )


def slide_ops(prs):
    s = content(prs, "05  ·  核心商业价值  ·  汇报重点", "餐饮、夜间与北外滩联动")
    blocks = [
        ("餐饮收入", "原则上优先发挥一滴水现有餐饮能力。具体内容和收入分配写入正式合同。", D.CATERING),
        ("夜间场景", "原稿把部分体验延伸到晚间 20:30，包括 VIP 江景餐饮、主题灯光、乐队和全场合影。这是一次夜间运营测试。效果良好时，可再研究江景音乐、主题晚宴、夜间沙龙和小型 Livehouse 类活动。前提是场地许可、噪声、消防和经营条件。", None),
        ("区域联动", "活动预计吸引上海本地及周边的年轻消费者。交通、排队和现场消费安排合理时，可与周边餐饮、零售做专属优惠、商户推荐、餐饮套餐和摄影路线，把单场地活动延伸成区域消费。", None),
    ]
    boxes = place_grid(blocks, 3, 1.12, 5.65)
    for x, y, w, h, (title, lead, items) in boxes:
        add_rect(s, x, y, w, h, WHITE, LINE, 0.75)
        add_rect(s, x, y, w, 0.045, GOLD)
        add_text(s, x + 0.16, y + 0.16, w - 0.32, 0.36, title, 16, NAVY, True)
        add_text(s, x + 0.16, y + 0.58, w - 0.32, 1.7 if items else 4.6, lead, 13, INK)
        if items:
            bullets(s, x + 0.16, y + 2.30, w - 0.32, 3.1, items, size=13, spacing=3)


def slide_brand(prs):
    s = content(prs, "05  ·  核心商业价值  ·  汇报重点", "年轻化认知与可复用内容")
    add_rect(s, 0.42, 1.12, 12.50, 1.35, NAVY)
    add_text(s, 0.62, 1.28, 12.1, 0.32, "传播中建议同时出现的关键词", 13, GOLD, True)
    add_text(s, 0.62, 1.68, 12.1, 0.55, "北外滩江景    特色玻璃建筑    高品质空间    沉浸式活动    年轻化文化体验", 18, WHITE, True)
    add_text(
        s, 0.42, 2.62, 12.5, 0.6,
        "一滴水的空间品质已经成立。部分年轻消费者还不熟悉具体产品和服务。到场之后的照片、视频和用户分享，可以继续扩大认知，并留下可重复使用的内容。",
        14, INK,
    )
    boxes = place_grid(D.BRAND_ASSETS, 4, 3.30, 1.55)
    for x, y, w, h, text in boxes:
        add_rect(s, x, y, w, h, WHITE, LINE, 0.75)
        add_rect(s, x, y, 0.07, h, GOLD)
        add_text(s, x + 0.2, y, w - 0.32, h, text, 14, NAVY, True, anchor=MSO_ANCHOR.MIDDLE)
    callout(
        s, 0.42, 6.15, 12.50, 0.82,
        "上述资料在取得使用授权后，可用于后续品牌宣传、客户推广和招商合作。影像素材的商业使用，需另行取得授权。",
        size=13,
    )


def slide_longterm(prs):
    s = content(prs, "06  ·  长期业务方向  ·  汇报重点", "长期合作的八个业务方向")
    boxes = place_grid(D.LONGTERM, 4, 1.12, 2.75, gap_y=0.14)
    for x, y, w, h, (num, title, body) in boxes:
        card(s, x, y, w, h, title, body, kicker=num, title_size=15, body_size=12)


def slide_wedding_layers(prs):
    s = content(prs, "06  ·  长期业务方向  ·  汇报重点", "主题婚礼分成三个产品层次")
    boxes = place_grid(D.WEDDING_LAYERS, 3, 1.15, 4.20)
    for x, y, w, h, (level, title, body) in boxes:
        card(s, x, y, w, h, title, body, kicker=level, title_size=18, body_size=14)
    callout(
        s, 0.42, 5.55, 12.50, 1.25,
        "未来可以把主题婚礼、宴会餐饮、婚礼摄影和场地租赁组合成特色产品。虚拟角色庆典和任何使用特定商业 IP 的产品，都先明确授权，再进入销售。",
        size=14,
    )


def slide_social(prs):
    s = content(prs, "06  ·  长期业务方向  ·  汇报重点", "主题社交与创作者活动")
    add_rect(s, 0.42, 1.12, 6.15, 5.70, WHITE, LINE, 0.75)
    add_rect(s, 0.42, 1.12, 6.15, 0.045, GOLD)
    add_text(s, 0.58, 1.28, 5.8, 0.36, "主题茶话会与社交", 16, NAVY, True)
    bullets(s, 0.58, 1.75, 5.8, 3.5, D.SOCIAL_FORMS, size=13, spacing=4)
    add_text(s, 0.58, 5.55, 5.8, 1.05, "适合预约制、小规模和主题化经营。优先考虑工作日下午、部分晚间及其他非核心时段。", 13, MUTED)
    add_rect(s, 6.75, 1.12, 6.17, 5.70, WHITE, LINE, 0.75)
    add_rect(s, 6.75, 1.12, 6.17, 0.045, TEAL)
    add_text(s, 6.91, 1.28, 5.85, 0.36, "创作者签售与见面", 16, NAVY, True)
    bullets(s, 6.91, 1.75, 5.8, 3.3, D.SIGNING_FORMS, size=13, spacing=4)
    add_text(s, 6.91, 5.20, 5.8, 1.4, "专业团队负责嘉宾邀请、粉丝组织、票务和传播。一滴水提供符合活动要求的场地及配套服务。", 13, MUTED)


def slide_more_products(prs):
    s = content(prs, "06  ·  长期业务方向  ·  汇报重点", "庆典、摄影、餐饮、快闪与夜间")
    blocks = [
        ("角色生日与纪念", "围绕年轻消费者喜爱的文化主题，开发纪念日和庆典。可组合美陈、影像、互动、餐饮和摄影，并逐步做成标准化包场。"),
        ("主题摄影与商拍", "依托全玻璃建筑和江景，与专业摄影机构合作，推出不同档次的场地和摄影套餐。"),
        ("主题餐饮与宴会", "江景套餐、下午茶、创作者晚宴、小型主题宴会、品牌晚宴，以及生日纪念餐饮。使用特定商业 IP 前先取得授权。"),
        ("快闪、音乐与夜间", "研究潮玩、饮品、数码、美妆和文创快闪。夜间可研究音乐沙龙、小型乐队、艺术交流和主题派对，以许可、消防、噪声和演出管理为前提。"),
    ]
    boxes = place_grid(blocks, 2, 1.15, 2.70, gap_y=0.14)
    for x, y, w, h, (title, body) in boxes:
        card(s, x, y, w, h, title, body, title_size=16, body_size=14)
    add_text(s, 0.42, 6.80, 12.5, 0.28, "摄影方向还包括：Coser、角色扮演、年轻婚礼、个人纪念、品牌宣传和青年文化主题摄影。", 12, MUTED)


def slide_sponsors(prs):
    s = content(prs, "07  ·  收入与合作模式  ·  汇报重点", "三档赞助权益")
    boxes = place_grid(D.SPONSOR_TIERS, 3, 1.08, 4.95)
    for x, y, w, h, tier in boxes:
        fill = SAND if tier["name"].startswith("钻石") else WHITE
        add_rect(s, x, y, w, h, fill, LINE, 0.75)
        add_rect(s, x, y, w, 0.045, GOLD)
        add_text(s, x + 0.16, y + 0.10, w - 0.32, 0.28, tier["name"], 13, GOLD_DEEP, True)
        add_text(s, x + 0.16, y + 0.38, w - 0.32, 0.40, tier["price"], 22, NAVY, True)
        add_text(s, x + 0.16, y + 0.84, w - 0.32, 0.55, tier["fit"], 12, MUTED)
        bullets(s, x + 0.12, y + 1.45, w - 0.24, 3.35, tier["items"], size=12, spacing=2)
    callout(
        s, 0.42, 6.18, 12.50, 0.80,
        "以上为原稿建议档位，并非已经确认的赞助订单。高一档包含前一档权益。结案报告拟含现场照片、触达数据、推文阅读量、品牌互动和效果总结。完整对照见附表。",
        size=13,
    )


def slide_barter(prs):
    s = content(prs, "07  ·  收入与合作模式  ·  汇报重点", "置换合作与一滴水自有权益")
    add_text(s, 0.42, 1.08, 6.1, 0.34, "现金之外的四种合作", 16, NAVY, True)
    add_text(s, 6.75, 1.08, 6.1, 0.34, "建议预留给一滴水的权益", 16, NAVY, True)
    for i, (title, body) in enumerate(D.BARTER):
        y = 1.52 + i * 1.05
        card(s, 0.42, y, 6.15, 0.96, title, body, title_size=14, body_size=12)
    for i, (title, body) in enumerate(D.VENUE_RIGHTS):
        y = 1.52 + i * 0.84
        add_rect(s, 6.75, y, 6.17, 0.76, WHITE, LINE, 0.75)
        add_rect(s, 6.75, y, 0.07, 0.76, GOLD)
        add_text(s, 6.95, y + 0.06, 5.8, 0.26, title, 13, NAVY, True)
        add_text(s, 6.95, y + 0.32, 5.8, 0.40, body, 12, INK)
    callout(
        s, 0.42, 6.20, 12.50, 0.78,
        "场地方权益在不侵犯品牌赞助商权益的前提下，由双方共同确定。目标品牌包括潮玩文创、饮料食品、数码 3C、美妆个护、摄影器材、服饰及其他适合青年文化场景的品牌。",
        size=13,
    )


def slide_income(prs):
    s = content(prs, "07  ·  收入与合作模式  ·  汇报重点", "收入分开列，避免重复计算")
    add_rect(s, 0.42, 1.12, 6.15, 4.85, WHITE, LINE, 0.75)
    add_rect(s, 0.42, 1.12, 6.15, 0.50, NAVY)
    add_text(s, 0.58, 1.12, 5.85, 0.50, "活动主办方的四类收入", 15, WHITE, True, anchor=MSO_ANCHOR.MIDDLE)
    for i, (title, body) in enumerate(D.HOST_INCOME):
        y = 1.78 + i * 0.98
        add_text(s, 0.62, y, 5.75, 0.28, title, 14, NAVY, True)
        add_text(s, 0.62, y + 0.30, 5.75, 0.58, body, 12, INK)
    add_rect(s, 6.75, 1.12, 6.17, 4.85, WHITE, LINE, 0.75)
    add_rect(s, 6.75, 1.12, 6.17, 0.50, TEAL)
    add_text(s, 6.91, 1.12, 5.85, 0.50, "建议单独列示的一滴水收入", 15, WHITE, True, anchor=MSO_ANCHOR.MIDDLE)
    for i, (title, body) in enumerate(D.VENUE_INCOME):
        y = 1.74 + i * 0.80
        add_text(s, 6.95, y, 5.8, 0.24, title, 13, NAVY, True)
        add_text(s, 6.95, y + 0.26, 5.8, 0.48, body, 12, INK)
    callout(
        s, 0.42, 6.12, 12.50, 0.85,
        "场地、餐饮、设备配套计入本次直接收入。后续婚宴和包场订单单独列示，不并入本次活动利润。联合招商的分成按合同约定，避免和场地费重复计算。",
        size=13,
    )


def slide_costs(prs):
    s = content(prs, "07  ·  收入与合作模式  ·  汇报重点", "成本结构与备用金")
    left = D.COSTS[:6]
    right = D.COSTS[6:]
    for col, items in enumerate((left, right)):
        for i, text in enumerate(items):
            x = 0.42 + col * 6.35
            y = 1.15 + i * 0.72
            add_rect(s, x, y, 6.15, 0.62, WHITE, LINE, 0.75)
            add_rect(s, x, y, 0.08, 0.62, GOLD)
            add_text(s, x + 0.24, y, 5.75, 0.62, text, 15, NAVY, True, anchor=MSO_ANCHOR.MIDDLE)
    callout(
        s, 0.42, 5.60, 12.50, 1.20,
        "原稿建议预留总预算的 5%—10% 作为备用金。请活动团队形成完整预算表，并向一滴水提供与场地合作直接相关的预算、保障安排和资金承担主体。成本清单见附表，金额待填。",
        size=14,
    )


def slide_modes(prs):
    s = content(prs, "07  ·  收入与合作模式  ·  汇报重点", "三种合作方式与建议")
    boxes = place_grid(D.MODES, 3, 1.08, 4.00)
    for x, y, w, h, (name, title, what, strength, risk) in boxes:
        add_rect(s, x, y, w, h, WHITE, LINE, 0.75)
        add_rect(s, x, y, w, 0.045, GOLD)
        add_text(s, x + 0.14, y + 0.12, w - 0.28, 0.24, name, 12, GOLD_DEEP, True)
        add_text(s, x + 0.14, y + 0.38, w - 0.28, 0.40, title, 16, NAVY, True)
        add_text(s, x + 0.14, y + 0.88, w - 0.28, 1.15, what, 12, INK)
        add_text(s, x + 0.14, y + 2.10, w - 0.28, 0.24, "适合的情况", 12, TEAL, True)
        add_text(s, x + 0.14, y + 2.34, w - 0.28, 0.70, strength, 12, INK)
        add_text(s, x + 0.14, y + 3.05, w - 0.28, 0.22, "需要写清的条件", 12, GOLD_DEEP, True)
        add_text(s, x + 0.14, y + 3.28, w - 0.28, 0.65, risk, 12, INK)
    add_rect(s, 0.42, 5.24, 12.50, 1.55, NAVY)
    add_text(s, 0.62, 5.36, 12.1, 0.30, "首期建议优先研究", 13, GOLD, True)
    add_text(s, 0.62, 5.68, 12.1, 0.40, "优惠场租  ＋  餐饮最低消费  ＋  场地品牌传播  ＋  后续业务合作权", 16, WHITE, True)
    add_text(s, 0.62, 6.15, 12.1, 0.48, "具体条件按商业原则协商。首期以验证模式为核心。缺少成本和收益测算时，不承诺长期免费使用场地。", 13, PALE)
    note(s, "若对方希望免费或纯置换，回到这页：置换可以谈，但要核算场地成本和机会成本，并写明未交付权益怎么处理。")


def slide_channels(prs):
    s = content(prs, "08  ·  传播、场地与执行", "传播渠道与计划目标")
    add_text(s, 0.42, 1.05, 12.5, 0.36, "传播方向：潮起潮落 TIDELINE  ×  北外滩江景  ×  青年文化  ×  沉浸式体验  ×  尚9·一滴水", 13, NAVY, True)
    rows = [("渠道", "计划覆盖或曝光", "主要内容")] + [row for row in D.CHANNELS]
    add_table(s, 0.42, 1.48, [2.5, 2.5, 7.5], rows, font=12, row_h=0.40, header_h=0.38,
              aligns=[PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.LEFT])
    extra = place_grid(D.EXTRA_CHANNELS, 2, 4.85, 0.85)
    for x, y, w, h, (name, num, body) in extra:
        add_rect(s, x, y, w, h, ICE, LINE, 0.6)
        add_text(s, x + 0.14, y + 0.08, w - 0.28, 0.26, f"{name}    {num}", 13, NAVY, True)
        add_text(s, x + 0.14, y + 0.38, w - 0.28, 0.40, body, 12, INK)
    callout(
        s, 0.42, 5.85, 12.50, 1.05,
        "合计约 70,000+ 是各渠道曝光或触达的计划目标，用户存在重叠，不能直接当成 7 万名独立消费者。抖音未计入该合计。统一话题：#尚9一滴水  #潮起潮落TIDELINE  #北外滩青年文化  #一滴水主题活动。",
        size=13,
    )


def slide_kol(prs):
    s = content(prs, "08  ·  传播、场地与执行", "资源核验与宣传节奏")
    add_rect(s, 0.42, 1.10, 6.15, 4.55, WHITE, LINE, 0.75)
    add_text(s, 0.58, 1.22, 5.85, 1.15, "原稿提到 UP主、Coser、B站百大，以及卡游、杭州银行等潜在合作方。粉丝规模写为 10 万至 1,000 万不等。正式合作前请提供：", 13, INK)
    bullets(s, 0.58, 2.45, 5.8, 3.0, D.KOL_CHECKS, size=13, spacing=3)
    rows = [("阶段", "时间", "主要内容")] + [row for row in D.RHYTHM]
    add_table(s, 6.75, 1.10, [1.55, 1.7, 3.0], rows, font=11, row_h=0.48, header_h=0.40)
    callout(
        s, 0.42, 5.82, 12.50, 1.12,
        "线下还包括高校动漫社和摄影社团、同人社团、合作商户、北外滩商圈合规展示、场地指引和品牌展位。原定筹备节点部分早于今天，执行时重排宣传排期。分享靠摄影、创作者合作和内容设计来促成。",
        size=13,
    )


def slide_floors(prs):
    s = content(prs, "08  ·  传播、场地与执行", "三个楼层分别承担什么")
    for i, (title, body) in enumerate(D.FLOORS):
        y = 1.15 + i * 1.45
        add_rect(s, 0.42, y, 12.50, 1.32, WHITE, LINE, 0.75)
        add_rect(s, 0.42, y, 0.08, 1.32, GOLD)
        add_text(s, 0.70, y + 0.12, 3.2, 1.05, title, 16, NAVY, True, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 4.00, y + 0.16, 8.7, 1.02, body, 14, INK, anchor=MSO_ANCHOR.MIDDLE)
    callout(
        s, 0.42, 5.60, 12.50, 1.20,
        "动线原稿：G层签到 → 安检 → 手环发放 → 摊位区 → 主题互动区 → 一楼宴会区域 → 三楼相应体验区。散场采取分批、单向和分流。具体动线须经场地方及专业安全管理人员审核。使用范围以最新平面图、使用权限、消防条件和核定容量为准。",
        size=13,
    )


def slide_areas(prs):
    s = content(prs, "08  ·  传播、场地与执行", "功能区面积需求")
    rows = [("区域", "拟设置位置", "计划面积")]
    for name, place, area in D.AREAS:
        rows.append((name, place, f"{area}㎡"))
    rows.append(("合计", "需求简单加总", "1,780㎡"))
    add_table(
        s, 0.42, 1.08, [4.3, 4.5, 3.7], rows, font=11, row_h=0.325, header_h=0.34,
        aligns=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT],
    )
    callout(
        s, 0.42, 6.20, 12.50, 0.78,
        "1,780㎡ 是方案需求的简单加总，不代表场地需要提供的独立净面积，也不能据此判断安全承载能力。各区域尚未获得场地方确认。",
        size=13,
    )


def slide_rundown(prs):
    s = content(prs, "08  ·  传播、场地与执行", "活动日流程")
    rows = [("时间", "活动内容", "场地方配合")] + [row for row in D.RUNDOWN]
    add_table(s, 0.42, 1.05, [2.15, 4.55, 5.80], rows, font=11, row_h=0.325, header_h=0.34)
    callout(
        s, 0.42, 6.15, 12.50, 0.82,
        "日程保留原稿时间，部分环节并行。VIP 约 200 人及分批人数，须按票务、空间和安全容量调整。布展写在 11月26日至28日，需等活动日确认后重排。占用周期按布展到恢复计算。",
        size=12,
    )


def slide_team(prs):
    s = content(prs, "08  ·  传播、场地与执行", "拟合作团队与自述规模")
    add_text(
        s, 0.42, 1.05, 12.5, 0.58,
        "北京惠迪文化传播有限公司，2016 年成立。定位为文化活动、会展、演出及 IP 运营服务机构。以下规模来自公司自述，正式合作前核验案例和履约情况。",
        14, INK,
    )
    boxes = place_grid(D.COMPANY_STATS, 6, 1.72, 1.15, gap_x=0.10)
    for x, y, w, h, (num, label) in boxes:
        add_rect(s, x, y, w, h, WHITE, LINE, 0.75)
        add_rect(s, x, y, w, 0.045, GOLD)
        add_text(s, x + 0.06, y + 0.16, w - 0.12, 0.42, num, 18, NAVY, True, align=PP_ALIGN.CENTER)
        add_text(s, x + 0.06, y + 0.64, w - 0.12, 0.40, label, 11, MUTED, align=PP_ALIGN.CENTER)
    add_text(s, 0.42, 3.05, 12.5, 0.30, "八个职能模块", 14, NAVY, True)
    tags = place_grid(D.FUNCTIONS, 8, 3.40, 0.48, gap_x=0.08)
    for x, y, w, h, name in tags:
        add_rect(s, x, y, w, h, ICE)
        add_text(s, x, y, w, h, name, 11, NAVY, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    boxes = place_grid(D.PHASE_MGMT, 4, 4.08, 1.55)
    phase_lines = {
        "孕育期": "立项、排期、预算框架\n场地落实、主题确定、招商\n嘉宾邀请及合作签约",
        "执行期": "现场方案、合规报批、采购物流\n人员到位、展前检查、搭建\n入场、运营及撤展",
    }
    for x, y, w, h, (title, body) in boxes:
        card(s, x, y, w, h, title, phase_lines.get(title, body), title_size=14, body_size=12)
    callout(
        s, 0.42, 5.80, 12.50, 1.15,
        "材料称覆盖全过程的 25 个管理节点，并有进度预警和分级管理。对一滴水有用的交付，是本次活动专属的执行计划、责任表和项目进度表。商业综合体合作经验涉及万达广场、大悦城、吾悦广场、爱琴海、永旺梦乐城、印象城、环宇城等，能否用于本项目需另行确认。",
        size=12,
    )


def slide_cases(prs):
    s = content(prs, "08  ·  传播、场地与执行", "材料所列代表项目")
    rows = [("项目名称", "材料所列规模", "材料所述传播")] + [row for row in D.CASES]
    add_table(s, 0.42, 1.08, [6.7, 2.1, 3.7], rows, font=11, row_h=0.40, header_h=0.36)
    callout(
        s, 0.42, 6.00, 12.50, 0.95,
        "传播数据以可核查资料为准。其他案例包括缀梦音游展、音律联觉舞美、主题展台、交通卡展台、旗忠劳力士功能展台、马文化周、舞马秀、马匹拍卖和沙地马球等。需确认团队在其中是主办、承办、联合执行还是单项供应商。",
        size=12,
    )


def slide_staff(prs):
    s = content(prs, "08  ·  传播、场地与执行", "现场人员配置")
    rows = [("工作组", "主要职责", "计划人数")] + [row for row in D.STAFF]
    add_table(s, 0.42, 1.15, [2.4, 7.3, 2.8], rows, font=14, row_h=0.52, header_h=0.44,
              aligns=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER])
    callout(
        s, 0.42, 5.45, 12.50, 1.35,
        "品牌赞助方案另设招商组，负责摊位招募、赞助对接和合同管理。上表人数属于策划阶段的初步安排。实际投入的人数、岗位资质和专业安保，依据场地条件、活动人数和相关管理要求确定。",
        size=14,
    )


def slide_safety(prs):
    s = content(prs, "08  ·  传播、场地与执行", "现场安全与环境保障")
    boxes = place_grid(D.SAFETY, 3, 1.12, 2.15, gap_y=0.12)
    for x, y, w, h, (title, body) in boxes:
        card(s, x, y, w, h, title, body, title_size=15, body_size=13)
    callout(
        s, 0.42, 5.70, 12.50, 1.15,
        "布展和撤展需提前确认：布展时间、施工要求、物料运输、货梯、用电、场地保护、撤展期限、垃圾清运和恢复标准。活动前完成实地踏勘，重点看 G层、一楼、三楼之间的电梯和应急通道。",
        size=14,
    )


def slide_tasks(prs):
    s = content(prs, "08  ·  传播、场地与执行", "需要一滴水配合确认的事项")
    for i, (num, title, body) in enumerate(D.VENUE_TASKS):
        col = i // 7
        row = i % 7
        x = 0.42 + col * 6.35
        y = 1.10 + row * 0.70
        add_rect(s, x, y, 6.20, 0.62, WHITE, LINE, 0.6)
        add_text(s, x + 0.1, y, 0.5, 0.62, num, 13, GOLD_DEEP, True, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 0.6, y, 1.7, 0.62, title, 13, NAVY, True, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + 2.35, y, 3.7, 0.62, body, 12, INK, anchor=MSO_ANCHOR.MIDDLE)
    callout(
        s, 0.42, 6.10, 12.50, 0.85,
        "另外确认：场地灯光、音响、餐饮及其他配套，哪些可以使用现有设备，哪些需要另行采购或租赁。十四项的完整表述见附表。",
        size=13,
    )


def slide_timeline(prs):
    s = content(prs, "08  ·  传播、场地与执行", "原定节点与今天的关系")
    add_text(s, 0.42, 1.02, 12.5, 0.28, "对照日期：" + D.AS_OF_CN + "。状态按原稿节点判断，供重排计划，不代表团队实际进度已经核验。", 12, MUTED)
    rows = [("阶段", "原计划时间", "主要工作", "相对今天")] + [row for row in D.TIMELINE]
    add_table(s, 0.42, 1.36, [2.3, 2.5, 5.5, 2.2], rows, font=11, row_h=0.38, header_h=0.34)
    callout(
        s, 0.42, 6.05, 12.50, 0.92,
        "明确合作意向后，请主办团队更新这些节点：场地签约、手续、核心嘉宾与品牌、正式开票、宣传物料。\n另附施工图、消防与现场安全方案，以及最终场地验收的时间。",
        size=13,
    )


def slide_raci(prs):
    s = content(prs, "09  ·  权责、评估与确认", "双方各自负责什么")
    cols = [
        ("主办及运营团队", D.HOST_DUTIES, NAVY),
        ("一滴水", D.VENUE_DUTIES, TEAL),
        ("双方共同", D.SHARED_DUTIES, GOLD_DEEP),
    ]
    for i, (title, items, color) in enumerate(cols):
        x = 0.42 + i * 4.22
        add_rect(s, x, 1.12, 4.06, 5.65, WHITE, LINE, 0.75)
        add_rect(s, x, 1.12, 4.06, 0.48, color)
        add_text(s, x, 1.12, 4.06, 0.48, title, 15, WHITE, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        bullets(s, x + 0.12, 1.72, 3.82, 4.85, items, size=12, spacing=2)


def slide_compliance(prs):
    s = content(prs, "09  ·  权责、评估与确认", "知识产权、保险与取消")
    add_rect(s, 0.42, 1.12, 6.15, 4.55, WHITE, LINE, 0.75)
    add_text(s, 0.58, 1.24, 5.85, 0.7, "活动为非官方同人文化活动。原稿提出不涉及《明日方舟》官方 IP 商业授权，并计划在宣传中注明非官方。正式活动仍需明确：", 13, INK)
    bullets(s, 0.58, 2.05, 5.8, 3.4, D.IP_POINTS, size=13, spacing=4)
    add_rect(s, 6.75, 1.12, 6.17, 4.55, WHITE, LINE, 0.75)
    add_text(s, 6.91, 1.26, 5.85, 0.34, "保险、损坏与取消", 16, NAVY, True)
    add_text(s, 6.91, 1.68, 5.85, 1.15, "合同中明确公众责任险及其他必要保险。因主办方原因造成场地设施损坏的，按合同承担修复或赔偿。布展前确认场地状态，撤场后完成验收。", 13, INK)
    add_text(s, 6.91, 2.90, 5.85, 0.55, "不可抗力导致延期或取消时协商处理，主办方提前通知。合同建议写明：", 13, INK)
    bullets(s, 6.91, 3.50, 5.8, 2.00, D.CANCEL_POINTS, size=12, spacing=1)
    callout(
        s, 0.42, 5.85, 12.50, 1.05,
        "注明“非官方活动”，并不能替代知识产权审核。正式主办方应提供合规说明和相应证明。右侧取消事项需要全部写入合同。",
        size=13,
    )


def slide_eval(prs):
    s = content(prs, "09  ·  权责、评估与确认", "用四类数据判断试点是否成立")
    blocks = [
        ("客流", "看活动是否真的发生", ["实际购票与入场", "VIP 与普通观众比例", "高峰人流与停留时间", "不同区域的参与情况"], False),
        ("经营", "看一滴水是否划算", ["场地、餐饮、配套收入", "实际投入成本", "净收益或净成本", "新增经营机会"], False),
        ("传播", "看内容是否可核查", ["B站、小红书、微博、短视频", "合作创作者发布", "场地话题与用户作品", "后台截图或平台统计"], False),
        ("转化", "建议新增的经营指标", ["场地与特色婚礼咨询", "生日、聚会、品牌活动咨询", "摄影及其他包场咨询", "实际形成的后续订单"], True),
    ]
    boxes = place_grid(blocks, 4, 1.12, 4.15)
    for x, y, w, h, (title, lead, items, focus) in boxes:
        add_rect(s, x, y, w, h, SAND if focus else WHITE, LINE, 0.75)
        add_rect(s, x, y, w, 0.045, GOLD)
        add_text(s, x + 0.14, y + 0.16, w - 0.28, 0.36, title, 18, NAVY, True)
        add_text(s, x + 0.14, y + 0.56, w - 0.28, 0.55, lead, 13, GOLD_DEEP, True)
        bullets(s, x + 0.10, y + 1.25, w - 0.22, 2.7, items, size=13, spacing=6)
    callout(
        s, 0.42, 5.45, 12.50, 1.35,
        "原稿提出活动结束后 3—7 天提交复盘报告。建议双方收到报告后开会，判断收益、资源投入、对原有经营的影响、年轻客群认知，以及哪些内容值得继续做。影像和内容的后续商业使用需取得授权。完整指标和交付清单见附表。",
        size=13,
    )


def slide_logic(prs):
    s = content(prs, "09  ·  权责、评估与确认", "从一次活动走到长期经营")
    for i, (title, body) in enumerate(D.LOGIC):
        hbar(s, 0.42, 1.20 + i * 1.20, 12.50, 1.08, f"0{i+1}", title, body)
    add_text(
        s, 0.42, 6.15, 12.5, 0.8,
        "成熟之后，还可以研究与北外滩其他商业资源、文化机构和品牌的联合活动，形成兼顾文化体验、商业消费和品牌传播的运营方式。是否进入这一步，由首期数据决定。",
        14, INK,
    )


def slide_steps(prs):
    s = content(prs, "09  ·  权责、评估与确认", "合作推进的五步")
    boxes = place_grid(D.STEPS, 5, 1.20, 4.55, gap_x=0.12)
    for x, y, w, h, (num, title, body) in boxes:
        add_rect(s, x, y, w, h, WHITE, LINE, 0.75)
        add_rect(s, x, y, w, 0.08, GOLD)
        add_text(s, x + 0.12, y + 0.25, w - 0.24, 0.5, num, 20, GOLD_DEEP, True)
        add_text(s, x + 0.12, y + 0.85, w - 0.24, 0.8, title, 15, NAVY, True)
        add_text(s, x + 0.12, y + 1.75, w - 0.24, 2.5, body, 13, INK)
    callout(
        s, 0.42, 5.95, 12.50, 0.95,
        "第三步的专项方案至少包括：商业预算、平面图、人流动线、安全及应急预案、嘉宾与赞助商清单、宣传计划、执行团队、场地使用需求和收益安排。",
        size=13,
    )


def slide_issues(prs):
    s = content(prs, "09  ·  权责、评估与确认", "签约前必须确认的八件事")
    short = [
        ("活动日期", "高", "28日与29日并存，且有首页与正文不一致；场地方案以29日为主", "对外写 11月下旬，日期待确认"),
        ("主办主体", "高", "已有北京惠迪介绍，法律角色仍待定", "未闭合前不承诺签约、收款和开票主体"),
        ("审批手续", "高", "一版写已办理，另一版待办理，未见完成证明", "按未完成管理，由主办方提交文件"),
        ("宣传资源", "高", "粉丝量、百大 UP主、卡游、杭州银行尚无确认清单", "只作为可推进的潜在资源"),
        ("场地范围", "高", "楼层、露台权限及对既有订单的影响未核实", "面积需求不等于可用面积或容量"),
        ("商业预算", "高", "尚无完整收入、成本、盈亏和风险准备金", "计划曝光和客流不折算成利润"),
        ("场地方收益", "高", "一滴水的收入、成本和权益尚未单列", "与主办方总收入分开列示"),
        ("长期合作", "中", "婚礼、虚拟庆典和特色餐饮的需求与授权未评估完", "用首期验证，不绑定长期免费场地"),
    ]
    rows = [("事项", "优先级", "当前缺口", "签约前的使用口径")] + [row for row in short]
    add_table(s, 0.28, 1.08, [1.7, 1.15, 5.15, 4.55], rows, font=11, row_h=0.58, header_h=0.40)
    note(s, "这页建议留在汇报最后。老板如果只记一件事，就是：方向可以谈，这八项没有闭合之前不签约、不对外宣传。")


def slide_close(prs):
    s = blank(prs)
    add_rect(s, 0, 0, 13.333, 7.5, NAVY)
    add_rect(s, 0, 0, 13.333, 0.08, GOLD)
    add_text(s, 0.62, 0.32, 12, 0.28, "合作愿景", 14, GOLD, True)
    add_text(s, 0.62, 0.68, 12, 0.55, "新客户、新收入、新品牌", 28, WHITE, True)
    items = [
        ("新客户", "用兴趣明确的青年文化活动，接触新的年轻消费者，为未来的婚礼、宴会、聚会和特色消费积累潜在客户。"),
        ("新收入", "用主题活动、特色餐饮、婚庆产品、摄影和包场，探索原有业务之外的增量经营机会。"),
        ("新品牌", "保持高品质宴会空间的定位，同时增加青年文化、特色活动和个性化消费场景的认知。"),
    ]
    boxes = place_grid(items, 3, 1.45, 2.15)
    for x, y, w, h, (title, body) in boxes:
        add_rect(s, x, y, w, h, NAVY_MID)
        add_text(s, x + 0.16, y + 0.14, w - 0.32, 0.36, title, 16, GOLD, True)
        add_text(s, x + 0.16, y + 0.58, w - 0.32, 1.4, body, 13, PALE)
    add_text(
        s, 0.62, 3.80, 12.1, 0.85,
        "希望把《潮起潮落 TIDELINE》做成一次可验证、可评估、可复制的试点。年轻人在体验一滴水之后，未来愿意再次选择这里，举办婚礼、生日、聚会、纪念庆典，以及其他值得纪念的重要活动。",
        15, WHITE,
    )
    pills = ["品牌年轻化", "客群多元化", "场景创新化", "经营高效化", "合作长期化"]
    for i, text in enumerate(pills):
        x = 0.62 + i * 2.48
        add_rect(s, x, 4.85, 2.32, 0.46, GOLD)
        add_text(s, x, 4.85, 2.32, 0.46, text, 13, NAVY, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(
        s, 0.62, 5.55, 12.1, 0.85,
        "以青年文化引入客户，以高品质空间创造体验，以特色婚庆和餐饮形成消费，以长期运营创造持续收益。",
        16, PALE_GOLD, True,
    )
    add_text(s, 0.62, 6.65, 12, 0.3, "北外滩尚9·一滴水    ·    青年文化活动、主题婚庆与精品社交", 13, PALE)
    add_page_slot(s, PALE_GOLD)
    note(s, "收束时重复三句话：新客户、新收入、新品牌。然后回到待确认事项，请对方指定下一步对接人。")


def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    prs.core_properties.title = "潮起潮落 TIDELINE × 北外滩尚9·一滴水 综合合作方案"
    prs.core_properties.subject = "青年文化消费场景创新暨长期战略合作"
    prs.core_properties.category = "内部研讨稿"
    prs.core_properties.keywords = "一滴水,TIDELINE,北外滩,婚庆,青年文化"
    builders = [
        slide_cover, slide_three, slide_toc,
        lambda p: section(p, "01", "合作契机", "年轻消费、宴会档期，以及一滴水已经具备的空间基础。"),
        slide_market, slide_traits, slide_pressure, slide_assets,
        lambda p: section(p, "02", "定位与路径", "一次活动是起点。目标是客群、产品和经营方式的长期变化。"),
        slide_goals, slide_phases,
        lambda p: section(p, "03", "首期活动与客群", "用《潮起潮落 TIDELINE》验证年轻文化活动能不能在一滴水成立。"),
        slide_event, slide_contents, slide_highlights, slide_audiences,
        lambda p: section(p, "04", "主题场景", "海洋主题与江景建筑的内容设计。使用官方素材前须完成授权审核。"),
        slide_zones, slide_interactions,
        lambda p: section(p, "05", "核心商业价值", "回答一滴水为什么要参与，以及参与之后能够经营什么。", True),
        slide_values, slide_wedding, slide_offpeak, slide_acquisition, slide_ops, slide_brand,
        lambda p: section(p, "06", "长期业务方向", "试点成立之后，才把这些方向做成可售卖的产品。", True),
        slide_longterm, slide_wedding_layers, slide_social, slide_more_products,
        lambda p: section(p, "07", "收入与合作模式", "赞助、场地方收入、成本和三种合作方式。金额均为建议或待填。", True),
        slide_sponsors, slide_barter, slide_income, slide_costs, slide_modes,
        lambda p: section(p, "08", "传播、场地与执行", "计划目标、空间需求、日程、团队和需要场地确认的清单。"),
        slide_channels, slide_kol, slide_floors, slide_areas, slide_rundown,
        slide_team, slide_cases, slide_staff, slide_safety, slide_tasks, slide_timeline,
        lambda p: section(p, "09", "权责、评估与待确认", "先分清责任和合规，再用数据决定要不要做长期合作。"),
        slide_raci, slide_compliance, slide_eval, slide_logic, slide_steps, slide_issues, slide_close,
    ]
    for build_slide in builders:
        build_slide(prs)
    total = stamp_pages(prs)
    text = extract_text(prs)
    required = [
        "5,000元", "15,000元", "30,000元", "1,780㎡", "北京惠迪", "非官方",
        "11月28", "卡游", "杭州银行", "优惠场租", "深海猎人", "伊比利亚", "海嗣",
        "痛车", "风岚", "城市之眼", "Lolita", "无限暖暖", "30 分钟", "两小时",
        "5%—10%", "#尚9一滴水", "个人信息", "机会成本", "140 万", "B站会员购",
        "杨浦复兴岛", "静安大悦城", "明日方舟",
    ]
    missing = [item for item in required if item not in text]
    if missing:
        raise RuntimeError("PPT 缺少关键内容：" + "、".join(missing))
    if total < 40:
        raise RuntimeError("页数异常：" + str(total))
    return prs


def extract_text(prs):
    parts = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                parts.append(shape.text_frame.text)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        parts.append(cell.text)
    return "\n".join(parts)


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "deliverables",
        "潮起潮落TIDELINE-北外滩尚9一滴水-综合合作方案.pptx",
    )
    os.makedirs(os.path.dirname(out), exist_ok=True)
    prs = build()
    prs.save(out)
    print(f"{out}\nslides={len(prs.slides)}")


if __name__ == "__main__":
    main()
