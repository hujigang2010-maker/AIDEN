"""赞助与展位的内部招商策略。不得外发。"""

from __future__ import annotations

import os

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

from theme import (
    CARD_GOLD,
    CARD_TEAL,
    CREAM,
    GOLD,
    INK,
    LINE,
    MIST,
    MUTED,
    PAPER,
    ROSE,
    ROSE_SOFT,
    SLATE,
    TEAL,
    WHITE,
    add_table,
    add_text,
    blank,
    content_header,
    notes,
    rect,
    rrect,
)

TOTAL = 9
OUT = os.path.normpath(
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "..",
        "deliverables",
        "2026康博会专题分论坛",
        "08-内部-赞助与展位招商策略.pptx",
    )
)


def new_prs():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = "2026康博会赞助与展位招商策略"
    prs.core_properties.subject = "内部"
    prs.core_properties.category = "内部"
    prs.core_properties.author = "专题分论坛筹备组"
    return prs


def cover(prs):
    s = blank(prs)
    rect(s, 0, 0, 13.333, 7.5, INK)
    rect(s, 0, 0, 0.14, 7.5, ROSE)
    add_text(s, "内部招商策略    ·    不得外发", 0.62, 0.42, 8, 0.32, size=16, bold=True, color=GOLD)
    add_text(s, "2026年11月17日", 9.3, 0.42, 3.4, 0.32, size=16, color=CREAM, align=PP_ALIGN.RIGHT)
    add_text(s, "赞助与展位", 0.6, 1.7, 12, 0.7, size=42, bold=True, color=WHITE)
    add_text(s, "对外发什么，对内先定什么", 0.62, 2.5, 12, 0.5, size=26, color=GOLD)
    rect(s, 0.62, 3.2, 1.35, 0.045, GOLD)
    add_text(
        s,
        "对外只发09号方案。\n本套说明授权、佣金、6,800元小型展台，以及给杨磊的说法。",
        0.62,
        3.5,
        12,
        0.8,
        size=18,
        color=CREAM,
        spacing=1.15,
    )
    add_text(
        s,
        "报价依据：大会公开页，标准展位16,800元/9㎡，双开口加10%，论坛演讲12,000元/人/20分钟。\n距11月28日还有11天。没有书面授权，不对外说20个展位已经可售。",
        0.62,
        5.15,
        12,
        0.8,
        size=16,
        color=WHITE,
        spacing=1.15,
    )
    notes(s, "公开价格见 https://sisae.com.cn/czfy 。核对用，不要把秘书处电话印到我们的招商页。")


def gates(prs):
    s = blank(prs)
    content_header(s, "发出之前", "五件事没有书面确认，09号先不要发出去", 2, TOTAL)
    rows = [
        ("01", "展位销售权", "组委会书面确认：是否有20个展位、在什么位置、可否由我们招、按16,800元销售。"),
        ("02", "折扣和佣金", "书面约定收款主体、结算方式、代理佣金、能否打折。未约定前，336,000元不是我们的收入。"),
        ("03", "论坛赞助权", "2万、5万、1万这三档是否允许由我们收取，且不与大会已售权益冲突。"),
        ("04", "署名授权", "杨浦区科技企业联合会、复旦大学住房政策研究中心，授权前不印成已经联合承办。"),
        ("05", "演讲不要叠售", "官方12,000元/20分钟，和我们审核制的10—15分钟分享，是两件东西，不能打包成一次收费。"),
    ]
    for i, (n, title, body) in enumerate(rows):
        y = 1.15 + i * 1.05
        rrect(s, 0.4, y, 12.52, 0.96, PAPER, line=LINE, radius=0.08)
        add_text(s, n, 0.58, y + 0.24, 0.7, 0.46, size=16, bold=True, color=TEAL, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, title, 1.4, y + 0.12, 2.5, 0.7, size=16, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, body, 4.0, y + 0.14, 8.6, 0.7, size=14, color=SLATE, anchor=MSO_ANCHOR.MIDDLE)


def who(prs):
    s = blank(prs)
    content_header(s, "对谁推什么", "同一种产品，不要对所有人用同一句开场", 3, TOTAL)
    cards = [
        (TEAL, "瑞健未来  ·  杨磊", "优先讲2万元产业创新合作伙伴。他若原本已被邀请做分享、且没谈过费用，赞助是可选择的权益，不是把演讲改成必须付费。展位作为第二种选择。"),
        (GOLD, "其他产品型企业", "优先讲16,800元的9㎡正式展位。企业如果还想上论坛，再单独谈赞助。没有授权时，不把演讲时间一并卖出。"),
        (ROSE, "20个展位整体", "先确认销售权、位置、价格、佣金和折扣，再给企业发方案。论坛场地免费，不代表展位可以免费使用或自行转售。"),
    ]
    for i, (color, title, body) in enumerate(cards):
        x = 0.4 + i * 4.28
        rrect(s, x, 1.25, 4.12, 4.15, PAPER, line=LINE, radius=0.08)
        rect(s, x, 1.25, 4.12, 0.08, color)
        add_text(s, title, x + 0.22, 1.5, 3.7, 0.7, size=18, bold=True, color=color)
        add_text(s, body, x + 0.22, 2.35, 3.7, 2.75, size=15, color=SLATE, spacing=1.12)
    add_text(s, "09号里的“瑞健未来”专页，只随发给该公司的版本保留。", 0.45, 5.65, 12.4, 0.45, size=16, bold=True, color=INK)


def prices(prs):
    s = blank(prs)
    content_header(s, "三套价格不要混", "对外主推两档，第三档没有授权就不能报", 4, TOTAL)
    add_table(
        s,
        0.4,
        1.22,
        12.52,
        3.35,
        [
            ["对外怎么说", "单价", "依据", "现在能不能报"],
            ["产业创新赞助", "20,000元/家", "本论坛拟议权益", "赞助权确认后可报"],
            ["大会标准展位", "16,800元/9㎡", "官网公开价，双开口18,480元", "销售权确认后可报"],
            ["官方演讲推介", "12,000元/人/20分钟", "官网公开价", "引导大会通道，不由我们另造"],
            ["论坛配套小展台", "6,800元/个（拟议）", "不是官方价格", "默认不报"],
        ],
        [3.1, 3.3, 3.5, 2.62],
        size=13,
    )
    rrect(s, 0.4, 4.8, 12.52, 1.65, ROSE_SOFT, radius=0.08)
    add_text(
        s,
        "04号、05号里产业创新是2家。09号按本次口径写成2—3家，生态支持写成3—5家。两套文件不要同时发给同一家企业。对外以09号为准。首席5万元、1家，没有新造权益；企业若问首席和生态的明细，用已经写过的权益清单，不要现场加码。",
        0.62,
        5.0,
        12.1,
        1.3,
        size=15,
        color=INK,
        anchor=MSO_ANCHOR.MIDDLE,
    )


def money(prs):
    s = blank(prs)
    content_header(s, "销售额不是承办收入", "20个标准展位满额是33.6万元营业额", 5, TOTAL)
    add_table(
        s,
        0.4,
        1.22,
        12.52,
        2.15,
        [
            ["销售情况", "展位数", "单价", "销售总额"],
            ["全部售出", "20", "16,800元", "336,000元"],
            ["售出15个", "15", "16,800元", "252,000元"],
            ["售出10个", "10", "16,800元", "168,000元"],
        ],
        [3.13, 3.13, 3.13, 3.13],
        size=15,
    )
    points = [
        ("光地不要混进来", "官网另有国内光地1,680元/㎡，36㎡起。那不是这20个标准展位，对外不报这个单价。"),
        ("佣金还没约定", "实际归属、代理佣金、结算和团购优惠，要和大会展位销售主体另签。签之前不承诺分成。"),
        ("个人不垫付", "展位款和赞助款进授权主体的对公账户。不要求个人先垫付展位费或赞助费。"),
    ]
    for i, (title, body) in enumerate(points):
        x = 0.4 + i * 4.28
        rrect(s, x, 3.6, 4.12, 2.45, MIST, radius=0.08)
        add_text(s, title, x + 0.2, 3.78, 3.72, 0.55, size=16, bold=True, color=TEAL)
        add_text(s, body, x + 0.2, 4.4, 3.72, 1.4, size=14, color=SLATE, spacing=1.08)


def mini(prs):
    s = blank(prs)
    content_header(s, "第三种只留在内部", "6,800元小展台，没有书面许可就不能卖", 6, TOTAL)
    add_text(
        s,
        "只有在已经掌握论坛附近的免费配套展示区、而且组委会书面允许我们单独销售时，才讨论这一档。它不是大会现有展位的官方价格。",
        0.45,
        1.18,
        12.4,
        0.7,
        size=16,
        color=SLATE,
    )
    add_table(
        s,
        0.4,
        2.0,
        12.52,
        1.7,
        [
            ["类型", "规格", "单价", "20个合计"],
            ["大会官方标准展位", "3m×3m，9㎡，三天", "16,800元", "336,000元"],
            ["论坛配套展示台（拟议）", "2m×2m，4㎡，配合论坛当天", "6,800元", "136,000元"],
        ],
        [3.6, 3.5, 2.5, 2.92],
        size=14,
    )
    blocks = [
        ("场地", "20个4㎡展台的净面积是80㎡。通道和疏散另计，实际区域通常明显大于80㎡。"),
        ("成本", "搭建、用电和现场人员另计。6,800元还没扣这些成本，不要当成利润。"),
        ("禁则", "组委会不允许论坛单独销售小型展台时，不得按6,800元收取展位费，也不写入09号。"),
    ]
    for i, (title, body) in enumerate(blocks):
        x = 0.4 + i * 4.28
        rrect(s, x, 3.95, 4.12, 2.15, PAPER, line=LINE, radius=0.08)
        add_text(s, title, x + 0.2, 4.1, 3.7, 0.36, size=16, bold=True, color=ROSE)
        add_text(s, body, x + 0.2, 4.55, 3.7, 1.35, size=14, color=SLATE)


def letter(prs):
    s = blank(prs)
    content_header(s, "可复制给杨磊", "赞助是一个选择，展位是另一个选择", 7, TOTAL)
    rrect(s, 0.4, 1.18, 12.52, 5.35, PAPER, line=LINE, radius=0.08)
    add_text(
        s,
        "杨总，您好。结合本次康博会，我们希望邀请瑞健未来更深入参与银发科技与具身智能产业创新论坛。目前有两种方式：",
        0.65,
        1.38,
        12.05,
        0.7,
        size=16,
        color=INK,
    )
    add_text(
        s,
        "方案一：论坛产业创新合作伙伴，2万元。包括品牌展示、专业分享申请、产业交流和相关宣传。若您方便，希望围绕“长寿科技走进社区”，分享社区门店和主动健康设备的实践。分享须经议程审核。",
        0.65,
        2.2,
        12.05,
        1.15,
        size=16,
        color=SLATE,
    )
    add_text(
        s,
        "方案二：9㎡标准展位，16,800元，展期11月27日至29日，用于设备和社区健康服务产品展示。我们也在协调一个约20家企业的智慧康养科技展示板块，位置和销售权以组委会确认为准。",
        0.65,
        3.45,
        12.05,
        1.15,
        size=16,
        color=SLATE,
    )
    add_text(
        s,
        "两种可以单独选。社区长寿科技和这次大会是同题的，想请您看哪种更符合公司的安排。",
        0.65,
        4.75,
        12.05,
        0.7,
        size=16,
        color=INK,
    )
    add_text(s, "不要写：演讲现已改为付费，不赞助就不能讲。", 0.65, 5.7, 12.05, 0.4, size=15, bold=True, color=ROSE)
    notes(s, "若此前邀请没有谈费用，这条红字是底线。企业主动要官方20分钟演讲时，按12,000元引导大会通道。")


def scrub(prs):
    s = blank(prs)
    content_header(s, "外发检查", "09号发出前，删掉这些内部说法", 8, TOTAL)
    left = [
        "6,800元小型展台",
        "336,000元是承办收入或利润",
        "佣金比例、分成、垫资",
        "20个展位已经批给我们",
        "杨磊已经确认出席或演讲",
    ]
    right = [
        "联合承办单位已经确定",
        "2万元自动包含20分钟官方演讲",
        "展位费包含主旨、圆桌或冠名",
        "大会秘书处电话写成我们的联系人",
        "04号、05号和09号同时发给同一家",
    ]
    for col, items in enumerate((left, right)):
        x = 0.4 + col * 6.45
        for i, item in enumerate(items):
            y = 1.25 + i * 0.85
            rrect(s, x, y, 6.25, 0.74, ROSE_SOFT, radius=0.08)
            add_text(s, item, x + 0.22, y + 0.16, 5.85, 0.44, size=16, color=INK, anchor=MSO_ANCHOR.MIDDLE)
    notes(s, "医疗器械、氧疗、机器人动态演示，参展前做资质、宣传合规和现场安全审核。这句已写在09号，对外保留。")


def decision(prs):
    s = blank(prs)
    content_header(s, "建议就这样定", "主推两万元赞助，展位按16,800元，授权先于发稿", 9, TOTAL)
    items = [
        ("对外报价", "赞助主推2万元。展位主推16,800元。满额33.6万元只在内部当作营业额。"),
        ("给杨磊", "两个选择都给。优先2万元。不把已发出的分享邀请改成强制付费。"),
        ("给其他企业", "先推展位。删掉瑞健未来专页。演讲需求单独谈，或引导官方12,000元通道。"),
        ("小型展台", "6,800元留在本页。组委会书面允许之前，口头不报、文件不发。"),
        ("收款", "合同、发票、收款走授权主体。个人不垫付。"),
    ]
    for i, (title, body) in enumerate(items):
        y = 1.18 + i * 1.02
        rrect(s, 0.4, y, 12.52, 0.94, CARD_TEAL if i == 0 else PAPER, line=LINE, radius=0.08)
        add_text(s, title, 0.62, y + 0.14, 2.3, 0.66, size=16, bold=True, color=TEAL, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, body, 3.05, y + 0.14, 9.6, 0.66, size=15, color=SLATE, anchor=MSO_ANCHOR.MIDDLE)
    notes(s, "09号联系人仍是待填写。填上授权联系人后再发。")


def build():
    prs = new_prs()
    cover(prs)
    gates(prs)
    who(prs)
    prices(prs)
    money(prs)
    mini(prs)
    letter(prs)
    scrub(prs)
    decision(prs)
    if len(prs.slides) != TOTAL:
        raise RuntimeError(f"页数{len(prs.slides)}与TOTAL {TOTAL}不一致")
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    prs.save(OUT)
    return OUT


if __name__ == "__main__":
    print(build())
