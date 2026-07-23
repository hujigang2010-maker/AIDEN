# -*- coding: utf-8 -*-
"""生成《复兴岛全球创客岛收官答卷大会》策划方案 PPT（主呈现）。"""
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

NAVY = RGBColor(0x0B, 0x1F, 0x3A)
NAVY2 = RGBColor(0x12, 0x2A, 0x4D)
TEAL = RGBColor(0x0A, 0x6E, 0x6A)
TEAL2 = RGBColor(0x1A, 0x9B, 0x8E)
GOLD = RGBColor(0xC4, 0x8A, 0x2A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xE6, 0xF0, 0xEE)
GREY = RGBColor(0x9A, 0xA7, 0xBD)
CARD = RGBColor(0x15, 0x2E, 0x45)
ROW_ALT = RGBColor(0x1A, 0x36, 0x50)
INK = RGBColor(0x1B, 0x2A, 0x44)

FONT = "Microsoft YaHei"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]
TOTAL = 16


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
    rect(s, 0, 0, SW, Inches(1.15), NAVY2)
    rect(s, 0, Inches(1.15), SW, Pt(3), TEAL)
    rect(s, Inches(0.55), Inches(0.3), Pt(6), Inches(0.55), GOLD)
    text(
        s,
        Inches(0.75),
        Inches(0.18),
        Inches(10.2),
        Inches(0.85),
        [[(kicker, 11, TEAL2, True)], [(title, 22, WHITE, True)]],
        space_after=2,
    )
    text(
        s,
        Inches(11.3),
        Inches(0.35),
        Inches(1.6),
        Inches(0.5),
        [[(f"{idx:02d}/{TOTAL:02d}", 14, GOLD, True)]],
        align=PP_ALIGN.RIGHT,
    )


def footer(s):
    text(
        s,
        Inches(0.7),
        Inches(7.08),
        Inches(8.5),
        Inches(0.3),
        [[("杨浦·复兴岛｜全球创客岛收官答卷大会｜策划方案 PPT", 9, GREY, False)]],
    )
    text(
        s,
        Inches(9.5),
        Inches(7.08),
        Inches(3.2),
        Inches(0.3),
        [[("建议日期 2026-09-12", 9, GREY, False)]],
        align=PP_ALIGN.RIGHT,
    )


def card(s, x, y, w, h, title, lines, accent=TEAL, tsize=14, bsize=11):
    rect(s, x, y, w, h, CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, y, w, Pt(4), accent, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    runs = [[(title, tsize, WHITE, True)]]
    for ln in lines:
        runs.append([("· " + ln, bsize, LIGHT, False)])
    text(s, x + Inches(0.18), y + Inches(0.16), w - Inches(0.36), h - Inches(0.3), runs, space_after=3)


def add_table(s, left, top, width, height, headers, rows, col_widths=None, font_size=11):
    """原生表格，深色主题。"""
    ncol = len(headers)
    nrow = 1 + len(rows)
    table_shape = s.shapes.add_table(nrow, ncol, left, top, width, height)
    table = table_shape.table
    if col_widths:
        for i, w in enumerate(col_widths):
            table.columns[i].width = w
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = TEAL
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            for r in p.runs:
                r.font.size = Pt(font_size)
                r.font.bold = True
                r.font.color.rgb = WHITE
                set_cjk(r)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = CARD if i % 2 == 0 else ROW_ALT
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT if j > 0 else PP_ALIGN.CENTER
                for r in p.runs:
                    r.font.size = Pt(font_size - 1)
                    r.font.color.rgb = LIGHT
                    set_cjk(r)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    return table_shape


# ===================== 01 封面 =====================
s = add_slide(NAVY)
rect(s, 0, 0, Inches(0.18), SH, TEAL)
rect(s, 0, Inches(5.65), SW, Pt(3), TEAL)
rect(s, 0, Inches(5.73), Inches(4.5), Pt(3), GOLD)
text(s, Inches(0.9), Inches(1.15), Inches(11.5), Inches(0.4),
     [[("杨浦区复兴岛｜PPT 策划方案", 14, TEAL2, True)]])
text(s, Inches(0.9), Inches(1.7), Inches(11.7), Inches(2.0),
     [[(C.PROJECT_NAME, 34, WHITE, True)],
      [(C.PROJECT_FULL, 20, GOLD, True)]], space_after=10)
text(s, Inches(0.9), Inches(3.9), Inches(11.5), Inches(1.4),
     [[("建议举办：" + C.EVENT_DATE, 16, WHITE, True)],
      [("9/15 前最靠近收官节点的开市·挂匾黄道吉日 ｜ 现场 200–300 人 ｜ 岛上主场", 13, LIGHT, False)],
      [("规格：区委区政府主要领导  +  「一带一路」多国总领事  +  产业龙头", 13, GREY, False)]],
     space_after=5)
text(s, Inches(0.9), Inches(5.95), Inches(11.5), Inches(0.7),
     [[("主呈现：PPT 汇报稿  +  Excel 执行台账", 15, WHITE, True)],
      [(C.VERSION, 11, GREY, False)]], space_after=3)

# ===================== 02 一页总览 =====================
s = add_slide(NAVY)
header(s, "ONE-PAGER · 一页总览", "给领导决策的核心信息", 2)
footer(s)
ov_rows = [[k, v] for k, v in C.OVERVIEW]
add_table(
    s, Inches(0.55), Inches(1.4), Inches(12.2), Inches(5.4),
    ["项目", "内容"], ov_rows,
    col_widths=[Inches(2.4), Inches(9.8)], font_size=12,
)

# ===================== 03 目录 =====================
s = add_slide(NAVY)
header(s, "CONTENTS", "方案目录", 3)
footer(s)
items = [
    ("01", "一页总览", "核心信息速览"),
    ("02", "背景意义", "创客岛收官·量子城市"),
    ("03", "择日结论", "首选 9/12 · 备选 9/9"),
    ("04", "定位目标", "五大目标清单"),
    ("05", "主题板块", "出海·AI·具身·低空"),
    ("06", "规模场地", "200–300 人·岛上主场"),
    ("07", "嘉宾规格", "领导+领事分层"),
    ("08", "揭牌落位", "国家厅+片区厅"),
    ("09", "详细议程", "全天流程表"),
    ("10", "倒排期", "T-50 至 T+14"),
    ("11", "预算 KPI", "投入与成效"),
    ("12", "决策清单", "本周拍板事项"),
]
for i, (n, t, d) in enumerate(items):
    col, row = i % 3, i // 3
    x = Inches(0.55) + col * Inches(4.2)
    y = Inches(1.5) + row * Inches(1.25)
    rect(s, x, y, Inches(4.0), Inches(1.05), CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, y, Inches(0.7), Inches(1.05), TEAL if row < 2 else GOLD)
    text(s, x, y, Inches(0.7), Inches(1.05), [[(n, 16, WHITE, True)]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + Inches(0.85), y + Inches(0.2), Inches(3.0), Inches(0.7),
         [[(t, 15, WHITE, True)], [(d, 11, GREY, False)]], space_after=2)

# ===================== 04 背景 =====================
s = add_slide(NAVY)
header(s, "WHY NOW · 为何此刻", "全球创客岛收官 × 量子城市迎合", 4)
footer(s)
cards = [
    ("2025.12 启动", ["市委领导启动全球创客岛", "量子城市年度成果发布", "实验基地叙事立住"]),
    ("2026.02 方案", ["国际创新创业集聚区方案", "三岛定位落地", "四大产业方向明确"]),
    ("2026 实践", ["具身智能实训平台启动", "空间智能体持续积累", "亟需高规格收官"]),
    ("本场使命", ["政治规格拉满", "国际领事到场", "会客厅硬落位", "产业协同可感知"]),
]
for i, (t, lines) in enumerate(cards):
    card(s, Inches(0.5) + i * Inches(3.2), Inches(1.5), Inches(3.05), Inches(5.15),
         t, lines, accent=[TEAL, TEAL2, GOLD, RGBColor(0x3D, 0x7E, 0xA6)][i], tsize=16, bsize=12)

# ===================== 05 择日 =====================
s = add_slide(NAVY)
header(s, "AUSPICIOUS DATE · 择日", "首选：2026 年 9 月 12 日（周六）", 5)
footer(s)
hl_rows = [
    ["公历星期", C.HUANGLI["date"]],
    ["干支值日", f"{C.HUANGLI['ganzhi']} · {C.HUANGLI['zhiri']}"],
    ["宜", "开市 · 挂匾 · 立券 · 交易 · 出行 · 祈福"],
    ["冲煞", C.HUANGLI["chong"]],
    ["择日理由", "9/15 前最靠近收官节点的开市/挂匾吉日"],
]
add_table(s, Inches(0.5), Inches(1.4), Inches(7.5), Inches(3.6),
          ["项目", "内容"], hl_rows,
          col_widths=[Inches(1.8), Inches(5.7)], font_size=12)
card(s, Inches(8.3), Inches(1.4), Inches(4.5), Inches(2.4),
     "备选一 · 9 月 9 日（周三）",
     ["最靠近 9/15 的工作日开市吉日", "宜开市/交易/立券/出行", "领导周六不便时启用"],
     accent=TEAL2)
card(s, Inches(8.3), Inches(4.05), Inches(4.5), Inches(2.5),
     "备选二 · 9 月 3 日（周四）",
     ["金匮黄道日", "宜开市/交易/立券", "可作预热或分论坛日"],
     accent=GOLD)
text(s, Inches(0.55), Inches(5.3), Inches(7.4), Inches(1.3),
     [[("规避提示", 13, GOLD, True)],
      [("9 月 10 日为杨公忌日，大事勿用；9 月 16 日已越过 9/15 收官节点。", 12, LIGHT, False)],
      [("决策口径：主推 9/12；领导周六不便则改 9/9。", 12, TEAL2, True)]],
     space_after=4)

# ===================== 06 定位目标 =====================
s = add_slide(NAVY)
header(s, "POSITIONING · 定位目标", "一场可汇报、可传播、可复用的答卷", 6)
footer(s)
text(s, Inches(0.6), Inches(1.4), Inches(12.1), Inches(0.9),
     [[(C.ONE_LINER, 12, LIGHT, False)]], space_after=3)
for i, obj in enumerate(C.OBJECTIVES):
    y = Inches(2.45) + i * Inches(0.8)
    rect(s, Inches(0.6), y, Inches(12.1), Inches(0.7), CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, Inches(0.6), y, Inches(0.7), Inches(0.7), TEAL if i % 2 == 0 else GOLD)
    text(s, Inches(0.6), y, Inches(0.7), Inches(0.7),
         [[(str(i + 1), 18, WHITE, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(1.5), y + Inches(0.15), Inches(11), Inches(0.45),
         [[(obj, 13, WHITE, False)]])

# ===================== 07 四大板块 =====================
s = add_slide(NAVY)
header(s, "PILLARS · 主题板块", "出海与科创主轴，三大硬科技协同", 7)
footer(s)
accents = [TEAL, TEAL2, GOLD, RGBColor(0x3D, 0x7E, 0xA6)]
for i, p in enumerate(C.THEME_PILLARS):
    x = Inches(0.45) + i * Inches(3.2)
    card(s, x, Inches(1.45), Inches(3.05), Inches(5.2),
         f"{p['name']}\n{p['tag']}", p["points"], accent=accents[i], tsize=15, bsize=12)

# ===================== 08 规模场地 =====================
s = add_slide(NAVY)
header(s, "SCALE & VENUE · 规模场地", "200–300 人 · 必须在岛上办", 8)
footer(s)
add_table(
    s, Inches(0.45), Inches(1.4), Inches(6.3), Inches(4.0),
    ["席别", "人数", "组成"], list(C.SEAT_PLAN),
    col_widths=[Inches(2.0), Inches(1.2), Inches(3.1)], font_size=11,
)
card(s, Inches(7.0), Inches(1.4), Inches(5.8), Inches(5.2),
     "场地原则与首选",
     [
         "政治象征：核心启动区 / 实验基地旁",
         "国际接待：贵宾室·同传·安保动线",
         "展示条件：成果展廊 + 具身演示区",
         "首选：创客岛核心区主会场（300–400）",
         "备选：滨江会客厅型 / 主场+分论坛",
         "底线：不离岛，保证叙事完整",
     ],
     accent=GOLD, tsize=15, bsize=12)

# ===================== 09 嘉宾 =====================
s = add_slide(NAVY)
header(s, "GUESTS · 嘉宾规格", "政治高位 + 国际高位同时在场", 9)
footer(s)
guest_rows = [[g["tier"], "；".join(g["targets"][:2]), g["goal"]] for g in C.GUEST_TIERS]
add_table(
    s, Inches(0.45), Inches(1.4), Inches(12.4), Inches(3.6),
    ["层级", "邀约对象", "目标"], guest_rows,
    col_widths=[Inches(2.4), Inches(6.5), Inches(3.5)], font_size=11,
)
bri = "　".join(n for n, _ in C.BRI_COUNTRY_POOL[:8])
text(s, Inches(0.55), Inches(5.3), Inches(12.2), Inches(1.3),
     [[("「一带一路」拟邀国别参考池（首场 6–10 国）", 13, GOLD, True)],
      [(bri + " …", 12, LIGHT, False)],
      [("揭牌国总领事须本人出席；完整名单与对接重点见 Excel「一带一路国别池」表。", 11, GREY, False)]],
     space_after=4)

# ===================== 10 揭牌 =====================
s = add_slide(NAVY)
header(s, "UNVEILING · 硬成果", "国家厅 + 片区厅 + 产业平台落位", 10)
footer(s)
uv_rows = [[u["name"], u["count"], u["form"], u["value"]] for u in C.UNVEILING]
add_table(
    s, Inches(0.4), Inches(1.4), Inches(12.5), Inches(3.2),
    ["落位类型", "数量", "形式", "价值"], uv_rows,
    col_widths=[Inches(3.5), Inches(1.8), Inches(3.4), Inches(3.8)], font_size=11,
)
for i, p in enumerate(C.UNVEILING_PRINCIPLES):
    y = Inches(4.9) + i * Inches(0.45)
    text(s, Inches(0.55), y, Inches(12.2), Inches(0.4),
         [[(f"原则 {i+1}  ", 12, TEAL2, True), (p, 12, LIGHT, False)]])

# ===================== 11 议程上午 =====================
s = add_slide(NAVY)
header(s, "AGENDA · 上午议程", "规格致辞 · 成果发布 · 揭牌签约", 11)
footer(s)
am = [a for a in C.AGENDA if a[0][:2] in ("08", "09", "10", "11", "12")]
add_table(
    s, Inches(0.35), Inches(1.35), Inches(12.6), Inches(5.4),
    ["时间", "环节", "内容要点", "责任"],
    [[a[0], a[1], a[2], a[3]] for a in am],
    col_widths=[Inches(1.7), Inches(3.0), Inches(5.8), Inches(2.1)], font_size=11,
)

# ===================== 12 议程下午 =====================
s = add_slide(NAVY)
header(s, "AGENDA · 下午议程", "分论坛深耕 · 国际对接酒会", 12)
footer(s)
pm = [a for a in C.AGENDA if a[0][:2] in ("13", "15", "16")]
add_table(
    s, Inches(0.35), Inches(1.35), Inches(12.6), Inches(4.6),
    ["时间", "环节", "内容要点", "责任"],
    [[a[0], a[1], a[2], a[3]] for a in pm],
    col_widths=[Inches(1.7), Inches(3.2), Inches(5.6), Inches(2.1)], font_size=11,
)
text(s, Inches(0.55), Inches(6.2), Inches(12.2), Inches(0.5),
     [[("完整可编辑议程、责任人与状态跟踪见配套 Excel「3.详细议程」工作表。", 12, GREY, False)]])

# ===================== 13 倒排期 =====================
s = add_slide(NAVY)
header(s, "TIMELINE · 倒排期", "从立项到 T+14 复盘", 13)
footer(s)
add_table(
    s, Inches(0.4), Inches(1.35), Inches(12.5), Inches(5.4),
    ["节点", "关键任务"],
    [[a, b] for a, b in C.TIMELINE],
    col_widths=[Inches(3.2), Inches(9.3)], font_size=11,
)

# ===================== 14 预算 KPI =====================
s = add_slide(NAVY)
header(s, "BUDGET & KPI · 预算成效", "中值约 55 万，成效必须可写进专报", 14)
footer(s)
budget_rows = [[a, b, c] for a, b, c in C.BUDGET]
add_table(
    s, Inches(0.35), Inches(1.35), Inches(7.0), Inches(5.4),
    ["成本项", "万元", "说明"], budget_rows,
    col_widths=[Inches(2.6), Inches(1.0), Inches(3.4)], font_size=10,
)
add_table(
    s, Inches(7.55), Inches(1.35), Inches(5.35), Inches(5.4),
    ["维度", "量化目标"], [[a, b] for a, b in C.KPIS],
    col_widths=[Inches(1.5), Inches(3.85)], font_size=10,
)

# ===================== 15 决策清单 =====================
s = add_slide(NAVY)
header(s, "NEXT · 决策清单", "建议本周拍板的五件事", 15)
footer(s)
steps = [
    ("1", "确认主日期 9/12，备选 9/9"),
    ("2", "正式命名并成立六组专班"),
    ("3", "锁定主会场与 3 个揭牌对象"),
    ("4", "启动外事报批与领事邀约"),
    ("5", "用本 PPT + Excel 作一页请示附件"),
]
for i, (n, t) in enumerate(steps):
    y = Inches(1.55) + i * Inches(0.95)
    rect(s, Inches(1.4), y, Inches(10.5), Inches(0.8), CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, Inches(1.4), y, Inches(0.85), Inches(0.8), TEAL if i < 3 else GOLD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, Inches(1.4), y, Inches(0.85), Inches(0.8),
         [[(n, 22, WHITE, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(2.5), y + Inches(0.18), Inches(9.1), Inches(0.5),
         [[(t, 18, WHITE, True)]])

# ===================== 16 封底 =====================
s = add_slide(NAVY)
rect(s, 0, 0, Inches(0.18), SH, TEAL)
rect(s, 0, Inches(3.35), SW, Pt(3), TEAL)
text(s, Inches(0.9), Inches(2.0), Inches(11.5), Inches(1.2),
     [[("把 9 月 12 日，做成复兴岛的完美答卷。", 28, WHITE, True)],
      [("全球创客岛 · 量子城市 · 科创出海 · 具身智能", 16, GOLD, True)]],
     align=PP_ALIGN.CENTER, space_after=10)
text(s, Inches(0.9), Inches(3.75), Inches(11.5), Inches(1.8),
     [[(C.EVENT_DATE, 18, TEAL2, True)],
      [("主呈现文件", 13, GREY, False)],
      [("PPT 策划方案  +  Excel 执行计划表（13 张工作表）", 14, LIGHT, False)],
      [(C.VERSION, 12, GREY, False)]],
     align=PP_ALIGN.CENTER, space_after=5)

prs.save(OUT_FILE)
print(f"已生成：{OUT_FILE}（{TOTAL} 页）")
