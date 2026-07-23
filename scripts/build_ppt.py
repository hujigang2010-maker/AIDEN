# -*- coding: utf-8 -*-
"""生成《复兴岛全球创客岛收官答卷大会》策划方案 PPT。"""
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

import content as C

OUT = Path(__file__).resolve().parent.parent / "deliverables"
OUT.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT / "复兴岛_全球创客岛_科创出海与具身智能国际大会_策划方案.pptx"

# 视觉：江岛青绿 + 深蓝（避免紫白/奶油陶土套路）
NAVY = RGBColor(0x0B, 0x1F, 0x3A)
NAVY2 = RGBColor(0x12, 0x2A, 0x4D)
TEAL = RGBColor(0x0A, 0x6E, 0x6A)
TEAL2 = RGBColor(0x1A, 0x9B, 0x8E)
GOLD = RGBColor(0xC4, 0x8A, 0x2A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xE6, 0xF0, 0xEE)
GREY = RGBColor(0x9A, 0xA7, 0xBD)
CARD = RGBColor(0x15, 0x2E, 0x45)
INK = RGBColor(0x1B, 0x2A, 0x44)

FONT = "Microsoft YaHei"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def set_cjk(run, name=FONT):
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", name)


def add_slide(bg=NAVY):
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    r.fill.solid()
    r.fill.fore_color.rgb = bg
    r.line.fill.background()
    r.shadow.inherit = False
    return s


def rect(s, x, y, w, h, color, line=None, line_w=None, shape=MSO_SHAPE.RECTANGLE):
    sp = s.shapes.add_shape(shape, x, y, w, h)
    if color is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = color
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = line_w or Pt(1)
    sp.shadow.inherit = False
    return sp


def text(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, space_after=4, line_spacing=1.05):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Pt(2)
    tf.margin_top = tf.margin_bottom = Pt(2)
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        p.line_spacing = line_spacing
        for t, size, color, bold in para:
            r = p.add_run()
            r.text = t
            r.font.size = Pt(size)
            r.font.color.rgb = color
            r.font.bold = bold
            set_cjk(r)
    return tb


def header(s, kicker, title, idx):
    rect(s, 0, 0, SW, Inches(1.2), NAVY2)
    rect(s, 0, Inches(1.2), SW, Pt(3), TEAL)
    rect(s, Inches(0.55), Inches(0.32), Pt(6), Inches(0.58), GOLD)
    text(
        s,
        Inches(0.75),
        Inches(0.2),
        Inches(10.5),
        Inches(0.9),
        [[(kicker, 11, TEAL2, True)], [(title, 24, WHITE, True)]],
        space_after=2,
    )
    text(
        s,
        Inches(11.7),
        Inches(0.4),
        Inches(1.2),
        Inches(0.5),
        [[("%02d" % idx, 22, GOLD, True)]],
        align=PP_ALIGN.RIGHT,
    )


def footer(s):
    text(
        s,
        Inches(0.7),
        Inches(7.05),
        Inches(9),
        Inches(0.35),
        [[("杨浦 · 复兴岛｜全球创客岛收官答卷大会", 9, GREY, False)]],
    )
    text(
        s,
        Inches(10.2),
        Inches(7.05),
        Inches(2.4),
        Inches(0.35),
        [[("建议日期 2026-09-12", 9, GREY, False)]],
        align=PP_ALIGN.RIGHT,
    )


def card(s, x, y, w, h, title, lines, accent=TEAL, tsize=14, bsize=11):
    rect(s, x, y, w, h, CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, y, w, Pt(4), accent, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    runs = [[(title, tsize, WHITE, True)]]
    for ln in lines:
        runs.append([("· " + ln, bsize, LIGHT, False)])
    text(s, x + Inches(0.2), y + Inches(0.18), w - Inches(0.4), h - Inches(0.35), runs, space_after=3)


# ===================== 01 封面 =====================
s = add_slide(NAVY)
rect(s, 0, Inches(5.7), SW, Pt(3), TEAL)
rect(s, 0, Inches(5.78), Inches(4.2), Pt(3), GOLD)
text(
    s,
    Inches(0.9),
    Inches(1.2),
    Inches(11.5),
    Inches(0.45),
    [[("杨浦区复兴岛｜全球创客岛阶段性成果发布暨国际合作揭牌", 14, TEAL2, True)]],
)
text(
    s,
    Inches(0.9),
    Inches(1.8),
    Inches(11.7),
    Inches(2.2),
    [
        [(C.PROJECT_NAME, 36, WHITE, True)],
        [(C.PROJECT_FULL, 22, GOLD, True)],
    ],
    space_after=10,
)
text(
    s,
    Inches(0.9),
    Inches(4.2),
    Inches(11.5),
    Inches(1.2),
    [
        [("建议举办：" + C.EVENT_DATE, 16, WHITE, True)],
        [("9 月 15 日前最靠近收官节点的开市 / 挂匾黄道吉日 ｜ 现场 200–300 人 ｜ 岛上主场", 13, LIGHT, False)],
        [("规格：区委区政府主要领导 +「一带一路」多国总领事 + 产业龙头", 13, GREY, False)],
    ],
    space_after=5,
)
text(
    s,
    Inches(0.9),
    Inches(5.95),
    Inches(11.5),
    Inches(0.6),
    [[("策划方案（建议稿）" + "  ·  " + C.VERSION, 14, WHITE, True)]],
)

# ===================== 02 目录 =====================
s = add_slide(NAVY)
header(s, "CONTENTS", "方案目录", 2)
footer(s)
items = [
    ("01  背景与战略意义", "创客岛收官 · 量子城市迎合"),
    ("02  择日结论", "首选 9/12，备选 9/9"),
    ("03  定位与四大板块", "出海 · AI · 具身 · 低空"),
    ("04  规模与场地", "200–300 人 · 岛上主场"),
    ("05  嘉宾与揭牌", "领导 + 领事 + 会客厅"),
    ("06  议程与组织", "全天流程 · 倒排期"),
    ("07  预算与 KPI", "成效可汇报"),
    ("08  风险与下一步", "立即决策清单"),
]
x0, y0 = Inches(0.7), Inches(1.6)
for i, (t, d) in enumerate(items):
    col, row = i % 2, i // 2
    x = x0 + col * Inches(6.2)
    y = y0 + row * Inches(1.2)
    card(s, x, y, Inches(5.9), Inches(1.0), t, [d], tsize=15, bsize=12)

# ===================== 03 背景 =====================
s = add_slide(NAVY)
header(s, "WHY NOW · 为何此刻", "全球创客岛收官 × 量子城市迎合", 3)
footer(s)
cards = [
    ("2025.12 启动", ["市委领导启动全球创客岛", "量子城市年度成果同步发布", "岛上实验基地叙事立住"]),
    ("2026.02 方案", ["国际创新创业集聚区实施方案", "数字智能 / 设计艺术 / 人民城市", "四大产业方向明确"]),
    ("2026 实践", ["具身智能实训实测平台启动", "空间智能体与场景持续积累", "需要高规格收官答卷"]),
    ("本场使命", ["政治规格拉满", "国际领事到场", "会客厅硬落位", "产业协同可感知"]),
]
for i, (t, lines) in enumerate(cards):
    card(s, Inches(0.55) + i * Inches(3.15), Inches(1.65), Inches(3.0), Inches(4.9), t, lines, tsize=16, bsize=12)

# ===================== 04 择日 =====================
s = add_slide(NAVY)
header(s, "AUSPICIOUS DATE · 择日", "首选：2026 年 9 月 12 日（周六）", 4)
footer(s)
rect(s, Inches(0.7), Inches(1.55), Inches(7.4), Inches(5.1), CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
rect(s, Inches(0.7), Inches(1.55), Inches(7.4), Pt(5), GOLD)
text(
    s,
    Inches(0.95),
    Inches(1.8),
    Inches(6.9),
    Inches(4.6),
    [
        [("黄历要点", 16, GOLD, True)],
        [("农历八月初二 · 定日 · 癸未日", 14, WHITE, True)],
        [("宜：开市 · 挂匾 · 立券 · 交易 · 出行 · 祈福", 13, LIGHT, False)],
        [("冲：冲羊煞东（属羊贵宾可调整位次）", 12, GREY, False)],
        [("", 8, GREY, False)],
        [("为何是它？", 16, TEAL2, True)],
        [("· 9/15 前最靠近收官节点的开市/挂匾吉日", 13, LIGHT, False)],
        [("· 同日宜出行与立券，契合领事到访与签约", 13, LIGHT, False)],
        [("· 避开 9/10 杨公忌日；9/16 已越过节点", 13, LIGHT, False)],
    ],
    space_after=5,
)
card(
    s,
    Inches(8.4),
    Inches(1.55),
    Inches(4.3),
    Inches(2.4),
    "备选一 · 9 月 9 日（周三）",
    ["最靠近 9/15 的工作日开市吉日", "宜开市 / 交易 / 立券 / 出行", "领导周六不便时启用"],
    accent=TEAL2,
)
card(
    s,
    Inches(8.4),
    Inches(4.2),
    Inches(4.3),
    Inches(2.45),
    "备选二 · 9 月 3 日（周四）",
    ["金匮黄道日", "宜开市 / 交易 / 立券", "可作预热或分论坛日"],
    accent=GOLD,
)

# ===================== 05 定位目标 =====================
s = add_slide(NAVY)
header(s, "POSITIONING · 定位目标", "一场可汇报、可传播、可复用的答卷", 5)
footer(s)
text(
    s,
    Inches(0.7),
    Inches(1.5),
    Inches(12),
    Inches(1.0),
    [[(C.ONE_LINER, 13, LIGHT, False)]],
    space_after=4,
)
for i, obj in enumerate(C.OBJECTIVES):
    y = Inches(2.6) + i * Inches(0.78)
    rect(s, Inches(0.7), y, Inches(12), Inches(0.68), CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, Inches(0.7), y, Pt(6), Inches(0.68), TEAL if i % 2 == 0 else GOLD)
    text(s, Inches(1.0), y + Inches(0.12), Inches(11.4), Inches(0.5), [[(obj, 13, WHITE, False)]])

# ===================== 06 四大板块 =====================
s = add_slide(NAVY)
header(s, "PILLARS · 主题板块", "出海与科创主轴，三大硬科技协同", 6)
footer(s)
accents = [TEAL, TEAL2, GOLD, RGBColor(0x3D, 0x7E, 0xA6)]
for i, p in enumerate(C.THEME_PILLARS):
    x = Inches(0.5) + i * Inches(3.2)
    card(s, x, Inches(1.6), Inches(3.05), Inches(5.0), f"{p['name']}\n{p['tag']}", p["points"], accent=accents[i], tsize=15, bsize=12)

# ===================== 07 规模场地 =====================
s = add_slide(NAVY)
header(s, "SCALE & VENUE · 规模场地", "200–300 人 · 必须在岛上办", 7)
footer(s)
card(
    s,
    Inches(0.55),
    Inches(1.55),
    Inches(6.0),
    Inches(5.1),
    "席位结构",
    [f"{a}｜{b}｜{c}" for a, b, c in C.SEAT_PLAN],
    tsize=15,
    bsize=11,
)
card(
    s,
    Inches(6.8),
    Inches(1.55),
    Inches(5.9),
    Inches(5.1),
    "场地原则与首选",
    [
        "政治象征：核心启动区 / 实验基地旁",
        "国际接待：贵宾室 · 同传 · 安保动线",
        "展示条件：成果展廊 + 具身演示区",
        "首选：创客岛核心区主会场（300–400）",
        "备选：滨江会客厅型 / 主场+分论坛分流",
        "底线：不离岛，保证叙事完整",
    ],
    accent=GOLD,
    tsize=15,
    bsize=12,
)

# ===================== 08 嘉宾 =====================
s = add_slide(NAVY)
header(s, "GUESTS · 嘉宾规格", "政治高位 + 国际高位同时在场", 8)
footer(s)
for i, g in enumerate(C.GUEST_TIERS):
    col, row = i % 2, i // 2
    x = Inches(0.55) + col * Inches(6.3)
    y = Inches(1.5) + row * Inches(2.55)
    card(s, x, y, Inches(6.05), Inches(2.35), g["tier"], g["targets"][:3] + [g["goal"]], tsize=15, bsize=11)

# ===================== 09 揭牌 =====================
s = add_slide(NAVY)
header(s, "UNVEILING · 硬成果", "国家厅 + 片区厅 + 产业平台落位", 9)
footer(s)
for i, u in enumerate(C.UNVEILING):
    x = Inches(0.55) + i * Inches(4.2)
    card(
        s,
        x,
        Inches(1.55),
        Inches(4.0),
        Inches(3.6),
        u["name"],
        [u["count"], u["form"], u["value"]],
        accent=[TEAL, GOLD, TEAL2][i],
        tsize=14,
        bsize=12,
    )
text(
    s,
    Inches(0.7),
    Inches(5.4),
    Inches(12),
    Inches(1.3),
    [
        [("落位原则", 14, GOLD, True)],
        [("先锁意向再上仪式 ｜ 一国一厅、一厅一责 ｜ 外事合规报批 ｜ 铭牌与媒体画面同步备妥", 13, LIGHT, False)],
    ],
)

# ===================== 10 议程 =====================
s = add_slide(NAVY)
header(s, "AGENDA · 议程节奏", "上午规格与揭牌，下午产业与对接", 10)
footer(s)
am = [a for a in C.AGENDA if a[0].startswith(("08", "09", "10", "11", "12"))]
pm = [a for a in C.AGENDA if a[0].startswith(("13", "15", "16"))]
card(
    s,
    Inches(0.5),
    Inches(1.5),
    Inches(6.1),
    Inches(5.2),
    "上午 · 政治答卷",
    [f"{t}  {n}" for t, n, _, __ in am[:8]],
    tsize=14,
    bsize=11,
)
card(
    s,
    Inches(6.85),
    Inches(1.5),
    Inches(5.9),
    Inches(5.2),
    "下午 · 产业与国际对接",
    [f"{t}  {n}" for t, n, _, __ in pm],
    accent=GOLD,
    tsize=14,
    bsize=11,
)

# ===================== 11 倒排期 =====================
s = add_slide(NAVY)
header(s, "TIMELINE · 倒排期", "从立项到 T+14 复盘", 11)
footer(s)
left = C.TIMELINE[:6]
right = C.TIMELINE[6:]
card(s, Inches(0.5), Inches(1.5), Inches(6.1), Inches(5.2), "前半程（锁日期 · 报批 · 揭牌）", [f"{a}：{b}" for a, b in left], tsize=13, bsize=11)
card(s, Inches(6.85), Inches(1.5), Inches(5.9), Inches(5.2), "后半程（确认 · 执行 · 转化）", [f"{a}：{b}" for a, b in right], accent=GOLD, tsize=13, bsize=11)

# ===================== 12 预算 KPI =====================
s = add_slide(NAVY)
header(s, "BUDGET & KPI · 预算成效", "中值约 55 万，成效必须可写进专报", 12)
footer(s)
card(
    s,
    Inches(0.5),
    Inches(1.5),
    Inches(6.1),
    Inches(5.2),
    "预算结构（万元，示意）",
    [f"{a}  {b}" for a, b, _ in C.BUDGET if a != "合计（示意）"]
    + [f"合计（示意）  {C.BUDGET[-1][1]}（场地置换后可压至约 40）"],
    tsize=14,
    bsize=11,
)
card(
    s,
    Inches(6.85),
    Inches(1.5),
    Inches(5.9),
    Inches(5.2),
    "核心 KPI",
    [f"{a}：{b}" for a, b in C.KPIS],
    accent=GOLD,
    tsize=14,
    bsize=11,
)

# ===================== 13 下一步 =====================
s = add_slide(NAVY)
header(s, "NEXT · 决策清单", "建议本周拍板的五件事", 13)
footer(s)
steps = [
    ("1", "确认主日期 9/12，备选 9/9"),
    ("2", "正式命名并成立六组专班"),
    ("3", "锁定主会场与 3 个揭牌对象"),
    ("4", "启动外事报批与领事邀约"),
    ("5", "提交一页纸请示供区领导决策"),
]
for i, (n, t) in enumerate(steps):
    y = Inches(1.6) + i * Inches(0.95)
    rect(s, Inches(1.5), y, Inches(10.3), Inches(0.8), CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, Inches(1.5), y, Inches(0.8), Inches(0.8), TEAL if i < 3 else GOLD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, Inches(1.5), y, Inches(0.8), Inches(0.8), [[(n, 22, WHITE, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(2.55), y + Inches(0.18), Inches(9), Inches(0.5), [[(t, 18, WHITE, True)]])

# ===================== 14 封底 =====================
s = add_slide(NAVY)
rect(s, 0, Inches(3.3), SW, Pt(3), TEAL)
text(
    s,
    Inches(0.9),
    Inches(2.0),
    Inches(11.5),
    Inches(1.2),
    [
        [("把 9 月 12 日，做成复兴岛的完美答卷。", 28, WHITE, True)],
        [("全球创客岛 · 量子城市 · 科创出海 · 具身智能", 16, GOLD, True)],
    ],
    align=PP_ALIGN.CENTER,
    space_after=10,
)
text(
    s,
    Inches(0.9),
    Inches(3.7),
    Inches(11.5),
    Inches(1.5),
    [
        [(C.EVENT_DATE, 18, TEAL2, True)],
        [("详细方案见 Word；执行台账见 Excel", 14, LIGHT, False)],
        [(C.VERSION, 12, GREY, False)],
    ],
    align=PP_ALIGN.CENTER,
    space_after=6,
)

prs.save(OUT_FILE)
print(f"已生成：{OUT_FILE}")
