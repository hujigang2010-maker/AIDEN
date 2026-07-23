# -*- coding: utf-8 -*-
"""生成《复兴岛全球创客岛收官答卷大会》策划方案 PPT（V3 · 敦煌+四层结构）。"""
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

FONT = "Microsoft YaHei"
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]
TOTAL = 26


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
    rect(s, 0, 0, SW, Inches(1.12), NAVY2)
    rect(s, 0, Inches(1.12), SW, Pt(3), TEAL)
    rect(s, Inches(0.55), Inches(0.28), Pt(6), Inches(0.55), GOLD)
    text(s, Inches(0.75), Inches(0.16), Inches(10.2), Inches(0.85),
         [[(kicker, 11, TEAL2, True)], [(title, 21, WHITE, True)]], space_after=2)
    text(s, Inches(11.3), Inches(0.35), Inches(1.6), Inches(0.5),
         [[(f"{idx:02d}/{TOTAL:02d}", 14, GOLD, True)]], align=PP_ALIGN.RIGHT)


def footer(s):
    text(s, Inches(0.7), Inches(7.08), Inches(8.8), Inches(0.3),
         [[("创客复兴·智汇杨浦｜全球创客岛成果发布暨国际峰会｜V3.2", 9, GREY, False)]])
    text(s, Inches(9.5), Inches(7.08), Inches(3.2), Inches(0.3),
         [[("建议日期 2026-09-12", 9, GREY, False)]], align=PP_ALIGN.RIGHT)


def card(s, x, y, w, h, title, lines, accent=TEAL, tsize=14, bsize=11):
    rect(s, x, y, w, h, CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, y, w, Pt(4), accent, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    runs = [[(title, tsize, WHITE, True)]]
    for ln in lines:
        runs.append([("· " + ln, bsize, LIGHT, False)])
    text(s, x + Inches(0.16), y + Inches(0.14), w - Inches(0.32), h - Inches(0.28),
         runs, space_after=2, line_spacing=1.02)


def add_table(s, left, top, width, height, headers, rows, col_widths=None, font_size=11):
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
                p.alignment = PP_ALIGN.LEFT if j else PP_ALIGN.CENTER
                for r in p.runs:
                    r.font.size = Pt(font_size - 1)
                    r.font.color.rgb = LIGHT
                    set_cjk(r)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    return table_shape


# ===== 01 封面 =====
s = add_slide(NAVY)
rect(s, 0, 0, Inches(0.18), SH, TEAL)
rect(s, 0, Inches(5.55), SW, Pt(3), TEAL)
rect(s, 0, Inches(5.63), Inches(4.8), Pt(3), GOLD)
text(s, Inches(0.9), Inches(0.95), Inches(11.5), Inches(0.35),
     [[("杨浦区 · 复兴岛｜政府口径主题已定稿", 13, TEAL2, True)]])
text(s, Inches(0.9), Inches(1.35), Inches(11.7), Inches(2.15),
     [[(C.PROJECT_NAME, 36, WHITE, True)],
      [(C.PROJECT_FULL, 15, GOLD, True)],
      [(C.PROJECT_SUBTITLE, 13, LIGHT, False)]],
     space_after=8)
text(s, Inches(0.9), Inches(3.75), Inches(11.5), Inches(1.5),
     [[("Slogan：锚定全球创客岛，答卷量子城市", 16, TEAL2, True)],
      [("主会日：" + C.EVENT_DATE + "　｜　服务 9/15 收官节点", 14, WHITE, True)],
      [("结构：主论坛 + 分论坛 + 体验层 + 黑客松　｜　杨浦AI企业同台互动", 13, LIGHT, False)],
      [("规格：区领导 + 东亚/东南亚总领事（按全部出席准备）+ 杨浦AI企业 + 创客", 12, GREY, False)]],
     space_after=4)
text(s, Inches(0.9), Inches(5.9), Inches(11.5), Inches(0.65),
     [[("主呈现：PPT + Excel　｜　" + C.VERSION, 14, WHITE, True)],
      [("高端大气主题 · 杨浦优质企业互动 · 国际揭牌 · 敦煌样本", 12, GREY, False)]],
     space_after=3)

# ===== 02 一页总览 =====
s = add_slide(NAVY)
header(s, "ONE-PAGER", "给领导决策的核心信息", 2)
footer(s)
add_table(s, Inches(0.5), Inches(1.35), Inches(12.3), Inches(5.45),
          ["项目", "内容"], [[k, v] for k, v in C.OVERVIEW],
          col_widths=[Inches(2.3), Inches(10.0)], font_size=11)

# ===== 03 目录 =====
s = add_slide(NAVY)
header(s, "CONTENTS", "方案目录", 3)
footer(s)
toc = [
    ("01", "一页总览"), ("02", "四层结构"), ("03", "历史文脉"),
    ("04", "领导寄托"), ("05", "规划课题"), ("06", "敦煌项目"),
    ("07", "冷启动角色"), ("08", "择日"), ("09", "定位目标"),
    ("10", "杨浦企业"), ("11", "主题板块"), ("12", "体验层"),
    ("13", "黑客松"), ("14", "规模场地"), ("15", "嘉宾领事"),
    ("16", "议程"), ("17", "倒排组织"), ("18", "预算KPI"),
]
for i, (n, t) in enumerate(toc):
    col, row = i % 6, i // 6
    x = Inches(0.4) + col * Inches(2.15)
    y = Inches(1.55) + row * Inches(1.65)
    rect(s, x, y, Inches(2.0), Inches(1.4), CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x, y + Inches(0.3), Inches(2.0), Inches(0.9),
         [[(n, 16, GOLD, True)], [(t, 12, WHITE, True)]],
         align=PP_ALIGN.CENTER, space_after=4)

# ===== 04 四层结构 =====
s = add_slide(NAVY)
header(s, "FORMAT · 活动结构", C.EVENT_FORMAT["title"], 4)
footer(s)
for i, (title, desc) in enumerate(C.EVENT_FORMAT["layers"]):
    y = Inches(1.35) + i * Inches(1.15)
    rect(s, Inches(0.45), y, Inches(12.4), Inches(1.05), CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, Inches(0.45), y, Inches(0.12), Inches(1.05), [GOLD, TEAL, TEAL2, RGBColor(0x3D, 0x7E, 0xA6)][i])
    text(s, Inches(0.75), y + Inches(0.12), Inches(11.8), Inches(0.85),
         [[(title, 14, GOLD if i == 0 else TEAL2, True)], [(desc, 12, LIGHT, False)]], space_after=2)
text(s, Inches(0.55), Inches(6.15), Inches(12.2), Inches(0.55),
     [[("组合逻辑：主论坛领导可见 · 分论坛专业可谈 · 体验层大众可感 · 黑客松青年可来", 12, GREY, False)]])

# ===== 05 历史文脉时间轴 =====
s = add_slide(NAVY)
header(s, "HERITAGE · 历史文脉", "从周家嘴浅滩到全球创客岛", 5)
footer(s)
rows = [[a, b, c] for a, b, c in C.HISTORY_TIMELINE]
add_table(s, Inches(0.4), Inches(1.3), Inches(12.5), Inches(5.5),
          ["时期", "节点", "要点"], rows,
          col_widths=[Inches(1.8), Inches(2.2), Inches(8.5)], font_size=10)

# ===== 06 文脉载体与衍生 =====
s = add_slide(NAVY)
header(s, "HERITAGE · 衍生议题", "工业遗存 + 敦煌文明如何转译为活动主题", 6)
footer(s)
card(s, Inches(0.4), Inches(1.35), Inches(5.9), Inches(5.3),
     "关键文脉载体",
     [f"{n}：{d[:42]}…" if len(d) > 42 else f"{n}：{d}" for n, d in C.HISTORY_HERITAGE],
     tsize=14, bsize=11)
# 右侧展示后 4 条衍生（含敦煌）
for i, d in enumerate(C.HISTORY_DERIVATIVES[-4:]):
    y = Inches(1.35) + i * Inches(1.3)
    card(s, Inches(6.55), y, Inches(6.3), Inches(1.2),
         d["name"].replace("议题衍生：", ""), d["points"],
         accent=GOLD if i % 2 else TEAL2, tsize=12, bsize=10)

# ===== 07 领导寄托 =====
s = add_slide(NAVY)
header(s, "MANDATE · 领导寄托", "市领导要求与区里如何接得住", 7)
footer(s)
# 陈吉宁要点精选
chen = C.LEADERSHIP[0]["points"][:5]
card(s, Inches(0.4), Inches(1.35), Inches(6.3), Inches(3.5),
     "市委书记陈吉宁（2025-12-18）", chen, accent=GOLD, tsize=14, bsize=11)
card(s, Inches(6.95), Inches(1.35), Inches(5.9), Inches(3.5),
     "杨浦区委区政府",
     C.LEADERSHIP[2]["points"][:5], accent=TEAL, tsize=14, bsize=11)
text(s, Inches(0.5), Inches(5.05), Inches(12.3), Inches(1.6),
     [[("本场活动如何回应寄托", 13, GOLD, True)]] +
     [[("· " + x, 12, LIGHT, False)] for x in C.LEADERSHIP_RESPONSE[:4]],
     space_after=2)

# ===== 08 规划对齐与课题 =====
s = add_slide(NAVY)
header(s, "PLANNING · 规划课题", "大上海与杨浦规划下的六大课题", 8)
footer(s)
add_table(s, Inches(0.35), Inches(1.3), Inches(6.3), Inches(5.4),
          ["维度", "对齐要点"], [[a, b] for a, b in C.PLANNING_ALIGN],
          col_widths=[Inches(1.8), Inches(4.5)], font_size=10)
add_table(s, Inches(6.85), Inches(1.3), Inches(6.0), Inches(5.4),
          ["课题", "题目", "聚焦"],
          [[a, b, c] for a, b, c in C.TOPIC_AGENDA],
          col_widths=[Inches(1.0), Inches(2.2), Inches(2.8)], font_size=10)

# ===== 09 敦煌项目 =====
s = add_slide(NAVY)
header(s, "DUNHUANG · 意向样本", C.DUNHUANG["name"], 9)
footer(s)
text(s, Inches(0.55), Inches(1.28), Inches(12.2), Inches(0.35),
     [[(C.DUNHUANG["status"], 12, GOLD, True)]])
card(s, Inches(0.4), Inches(1.7), Inches(6.2), Inches(3.0),
     "项目基本面（部分数据待核实）", C.DUNHUANG["basics"][:4], tsize=13, bsize=11)
card(s, Inches(6.85), Inches(1.7), Inches(5.95), Inches(3.0),
     "与复兴岛契合点", C.DUNHUANG["fit"], accent=TEAL2, tsize=13, bsize=11)
card(s, Inches(0.4), Inches(4.9), Inches(12.4), Inches(1.7),
     "活动用法：主论坛点题 · 分论坛 D 深讲 · 体验层打样 · 成熟则一事一议意向",
     C.DUNHUANG["event_use"], accent=GOLD, tsize=13, bsize=11)

# ===== 10 冷启动与角色 =====
s = add_slide(NAVY)
header(s, "COLD START · 角色原则", "智库顾问定位 · 一事一议 · 先调研后成交", 10)
footer(s)
card(s, Inches(0.4), Inches(1.35), Inches(6.2), Inches(5.3),
     "冷启动原则（会商纪要）", C.COLD_START, tsize=14, bsize=11)
card(s, Inches(6.85), Inches(1.35), Inches(5.95), Inches(5.3),
     "协同角色",
     [f"{a}：{b[:48]}…" if len(b) > 48 else f"{a}：{b}" for a, b in C.PARTNERS],
     accent=GOLD, tsize=13, bsize=11)

# ===== 11 择日 =====
s = add_slide(NAVY)
header(s, "DATE · 择日", "首选 2026-09-12（周六）", 11)
footer(s)
hl_rows = [
    ["公历星期", C.HUANGLI["date"]],
    ["宜", "开市 · 挂匾 · 立券 · 交易 · 出行"],
    ["节奏", "9/11 晚黑客松开营 → 9/12 主会+Demo Day"],
    ["理由", "9/15 前最近开市/挂匾吉日，契合揭牌与国际到访"],
]
add_table(s, Inches(0.45), Inches(1.4), Inches(7.6), Inches(3.2),
          ["项目", "内容"], hl_rows,
          col_widths=[Inches(1.8), Inches(5.8)], font_size=12)
card(s, Inches(8.3), Inches(1.4), Inches(4.5), Inches(2.5),
     "备选 · 9 月 9 日（周三）",
     ["工作日开市吉日", "领导周六不便时启用", "黑客松可前置周末"], accent=TEAL2)
card(s, Inches(8.3), Inches(4.15), Inches(4.5), Inches(2.4),
     "规避",
     ["9/10 杨公忌日，大事勿用", "9/16 已越过收官节点"], accent=GOLD)
text(s, Inches(0.55), Inches(4.9), Inches(7.5), Inches(1.6),
     [[("决策口径", 13, GOLD, True)],
      [("主推 9/12；领导周六不便则改 9/9。", 13, TEAL2, True)],
      [("黑客松与主会绑定，形成「创客夜→答卷日」完整叙事。", 12, LIGHT, False)]],
     space_after=4)

# ===== 12 定位目标 =====
s = add_slide(NAVY)
header(s, "GOALS · 定位目标", "可汇报、可传播、可复用的答卷", 12)
footer(s)
text(s, Inches(0.55), Inches(1.35), Inches(12.2), Inches(0.85),
     [[(C.ONE_LINER, 12, LIGHT, False)]])
for i, obj in enumerate(C.OBJECTIVES):
    y = Inches(2.35) + i * Inches(0.7)
    rect(s, Inches(0.55), y, Inches(12.2), Inches(0.62), CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, Inches(0.55), y, Inches(0.62), Inches(0.62), TEAL if i % 2 == 0 else GOLD)
    text(s, Inches(0.55), y, Inches(0.62), Inches(0.62),
         [[(str(i + 1), 16, WHITE, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(1.35), y + Inches(0.12), Inches(11.2), Inches(0.42),
         [[(obj, 12, WHITE, False)]])

# ===== 13 杨浦企业互动 =====
s = add_slide(NAVY)
header(s, "YANGPU AI · 企业互动", "优刻得 · 智谱 · 苏度 · 卓益得等上岛同台", 13)
footer(s)
rows = [[a, b, d] for a, b, _, d in C.YANGPU_ENTERPRISES[:8]]
add_table(s, Inches(0.3), Inches(1.28), Inches(12.7), Inches(4.3),
          ["企业", "赛道", "活动互动角色"], rows,
          col_widths=[Inches(2.4), Inches(2.0), Inches(8.3)], font_size=10)
text(s, Inches(0.5), Inches(5.75), Inches(12.3), Inches(0.9),
     [[("互动机制", 13, GOLD, True)],
      [("主论坛「杨浦AI力量」发言 ｜ 分论坛具身圆桌 ｜ 体验层真机 ｜ 黑客松导师/算力 ｜ 签约意向墙", 12, LIGHT, False)],
      [("说明：苏度科技为杨浦具身智能重点企业（会商口述「首度」对应此方向企业）。", 11, GREY, False)]],
     space_after=2)

# ===== 14 主题板块 =====
s = add_slide(NAVY)
header(s, "PILLARS · 主题板块", "文脉 · 出海 · 硬科技 · 敦煌沉浸", 14)
footer(s)
accents = [GOLD, TEAL, TEAL2, RGBColor(0x3D, 0x7E, 0xA6)]
for i, p in enumerate(C.THEME_PILLARS):
    x = Inches(0.4) + i * Inches(3.2)
    card(s, x, Inches(1.4), Inches(3.05), Inches(5.25),
         f"{p['name']}\n{p['tag']}", p["points"], accent=accents[i], tsize=14, bsize=12)

# ===== 14 体验/游戏层 =====
s = add_slide(NAVY)
header(s, "PLAY LAYER · 体验层", "论坛 × 游戏化体验：可玩可感的沉浸关卡", 15)
footer(s)
add_table(s, Inches(0.4), Inches(1.3), Inches(12.5), Inches(4.2),
          ["关卡", "内容"], [[a, b] for a, b in C.EXPERIENCE_LAYER],
          col_widths=[Inches(3.0), Inches(9.5)], font_size=12)
text(s, Inches(0.55), Inches(5.7), Inches(12.2), Inches(0.9),
     [[("说明：此处「游戏」指沉浸关卡与集章任务，非电竞赛事；与主论坛、分论坛 D（敦煌）、黑客松互相导流。", 12, LIGHT, False)],
      [("冷启动打样原则：空间未就绪可用多媒体样片；成熟后再上重资产永久落地。", 12, GREY, False)]],
     space_after=3)

# ===== 16 黑客松 =====
s = add_slide(NAVY)
header(s, "HACKATHON · 黑客松", C.HACKATHON["name"], 16)
footer(s)
text(s, Inches(0.55), Inches(1.3), Inches(12.2), Inches(0.45),
     [[(C.HACKATHON["slogan"] + "　｜　" + C.HACKATHON["window"], 14, GOLD, True)]])
card(s, Inches(0.4), Inches(1.85), Inches(6.2), Inches(2.4),
     "为何必须做黑客松", C.HACKATHON["why"], tsize=14, bsize=12)
card(s, Inches(6.85), Inches(1.85), Inches(5.95), Inches(2.4),
     "规模与场地",
     [C.HACKATHON["scale"], C.HACKATHON["venue"],
      "优胜项目登周六主舞台 Demo Day", "激励：奖金+券包+入孵绿色通道"],
     accent=GOLD, tsize=14, bsize=12)
tracks = [[a, b] for a, b in C.HACKATHON["tracks"]]
add_table(s, Inches(0.4), Inches(4.45), Inches(12.5), Inches(2.3),
          ["赛道", "课题方向"], tracks,
          col_widths=[Inches(3.5), Inches(9.0)], font_size=11)

# ===== 16 黑客松赛程 =====
s = add_slide(NAVY)
header(s, "HACKATHON · 赛程规则", "24 小时登岛出作品、出苗子", 17)
footer(s)
add_table(s, Inches(0.4), Inches(1.3), Inches(7.5), Inches(4.0),
          ["时间", "环节"], [[a, b] for a, b in C.HACKATHON["schedule"]],
          col_widths=[Inches(2.8), Inches(4.7)], font_size=11)
card(s, Inches(8.15), Inches(1.3), Inches(4.7), Inches(4.0),
     "规则与激励要点", C.HACKATHON["rules"][:5], accent=TEAL2, tsize=13, bsize=11)
text(s, Inches(0.5), Inches(5.55), Inches(12.3), Inches(1.1),
     [[("黑客松 KPI", 13, GOLD, True)],
      [("　｜　".join(C.HACKATHON["kpis"]), 13, LIGHT, False)]], space_after=4)

# ===== 17 规模场地 =====
s = add_slide(NAVY)
header(s, "SCALE & VENUE", "主会 200–300 人 · 体验/黑客松分流 · 必须在岛上", 18)
footer(s)
add_table(s, Inches(0.4), Inches(1.3), Inches(6.4), Inches(4.3),
          ["席别", "人数", "组成"], list(C.SEAT_PLAN),
          col_widths=[Inches(2.2), Inches(1.2), Inches(3.0)], font_size=10)
card(s, Inches(7.05), Inches(1.3), Inches(5.8), Inches(5.3),
     "场地首选：启动区 + 船台公园",
     [
         "主论坛/揭牌在核心启动区厂房会场",
         "领导领事短途步入船台公园文脉打卡",
         "黑客松置于相邻老厂房，动线最短",
         "可见船台、塔吊，保证历史—未来同框",
         "底线：不离岛，叙事完整",
     ],
     accent=GOLD, tsize=14, bsize=12)

# ===== 18 嘉宾揭牌 =====
s = add_slide(NAVY)
header(s, "GUESTS & UNVEILING", "政治高位 + 国际高位 + 杨浦企业 + 硬成果", 19)
footer(s)
guest_rows = [[g["tier"], "；".join(g["targets"][:2]), g["goal"]] for g in C.GUEST_TIERS]
add_table(s, Inches(0.35), Inches(1.3), Inches(12.6), Inches(2.9),
          ["层级", "邀约对象", "目标"], guest_rows,
          col_widths=[Inches(2.3), Inches(6.6), Inches(3.7)], font_size=10)
uv = [[u["name"], u["count"], u["form"]] for u in C.UNVEILING]
add_table(s, Inches(0.35), Inches(4.4), Inches(12.6), Inches(2.3),
          ["揭牌类型", "数量", "形式"], uv,
          col_widths=[Inches(4.5), Inches(2.2), Inches(5.9)], font_size=11)

# ===== 19 领事名单 =====
s = add_slide(NAVY)
header(s, "CONSULS · 东亚东南亚", "驻沪总领事名单 · 暂按全部出席准备", 20)
footer(s)
consul_rows = [
    [
        c["region"],
        c["country"],
        f"{c['name_zh']}（{c['name_en']}）",
        c["focus"],
        c["role"],
    ]
    for c in C.EAST_SE_ASIA_CONSULS
]
add_table(
    s,
    Inches(0.25),
    Inches(1.2),
    Inches(12.8),
    Inches(5.35),
    ["片区", "国家", "总领事", "合作侧重", "现场安排"],
    consul_rows,
    col_widths=[Inches(1.0), Inches(1.5), Inches(3.4), Inches(3.5), Inches(3.4)],
    font_size=9,
)
text(
    s,
    Inches(0.35),
    Inches(6.6),
    Inches(12.6),
    Inches(0.4),
    [[(
        "公开资料整理（2026-07）· 共 11 国驻沪总领事 · 未在沪设馆：缅甸/文莱/朝鲜/东帝汶（见 Excel）· 出席以外事确认为准",
        10, GREY, False,
    )]],
)

# ===== 20 议程上午 =====
s = add_slide(NAVY)
header(s, "AGENDA · 上午", "规格致辞 · 杨浦AI力量 · 揭牌签约", 21)
footer(s)
am = [a for a in C.AGENDA if a[0][:2] in ("08", "09", "10", "11", "12")]
add_table(s, Inches(0.3), Inches(1.28), Inches(12.7), Inches(5.5),
          ["时间", "环节", "内容要点", "责任"],
          [[a[0], a[1], a[2], a[3]] for a in am],
          col_widths=[Inches(1.6), Inches(2.9), Inches(6.1), Inches(2.1)], font_size=10)

# ===== 21 议程下午 =====
s = add_slide(NAVY)
header(s, "AGENDA · 下午", "四分论坛 · 体验层 · Demo Day · 酒会", 22)
footer(s)
pm = [a for a in C.AGENDA if a[0][:2] in ("13", "15", "16")]
add_table(s, Inches(0.3), Inches(1.28), Inches(12.7), Inches(5.0),
          ["时间", "环节", "内容要点", "责任"],
          [[a[0], a[1], a[2], a[3]] for a in pm],
          col_widths=[Inches(1.6), Inches(3.0), Inches(6.0), Inches(2.1)], font_size=9)
text(s, Inches(0.5), Inches(6.45), Inches(12.3), Inches(0.4),
     [[("完整议程、领事台账、敦煌对接与黑客松见 Excel。", 11, GREY, False)]])

# ===== 22 倒排组织 =====
s = add_slide(NAVY)
header(s, "TIMELINE & ORG", "倒排期与专班架构", 23)
footer(s)
add_table(s, Inches(0.35), Inches(1.28), Inches(7.7), Inches(5.5),
          ["节点", "关键任务"], [[a, b] for a, b in C.TIMELINE],
          col_widths=[Inches(2.5), Inches(5.2)], font_size=10)
card(s, Inches(8.25), Inches(1.28), Inches(4.65), Inches(5.5),
     "专班架构（新增文脉/黑客松组）",
     [f"{a}" for a, _ in C.ORG_STRUCTURE],
     accent=TEAL2, tsize=13, bsize=11)

# ===== 23 预算 KPI =====
s = add_slide(NAVY)
header(s, "BUDGET & KPI", "中值约 74 万 · 成效必须可写进专报", 24)
footer(s)
add_table(s, Inches(0.3), Inches(1.28), Inches(7.1), Inches(5.5),
          ["成本项", "万元", "说明"], [[a, b, c] for a, b, c in C.BUDGET],
          col_widths=[Inches(2.7), Inches(1.1), Inches(3.3)], font_size=9)
add_table(s, Inches(7.6), Inches(1.28), Inches(5.3), Inches(5.5),
          ["维度", "量化目标"], [[a, b] for a, b in C.KPIS],
          col_widths=[Inches(1.5), Inches(3.8)], font_size=9)

# ===== 24 决策清单 =====
s = add_slide(NAVY)
header(s, "NEXT · 决策清单", "建议立即拍板与跟进的关键事项", 25)
footer(s)
steps = [
    ("1", "确认主题「创客复兴·智汇杨浦」与主会日 9/12"),
    ("2", "锁定优刻得/智谱/苏度等企业互动与发言名单"),
    ("3", "48h 完成敦煌核查；锁定会场与揭牌对象"),
    ("4", "启动外事报批；按 11 国总领事全名单邀约"),
    ("5", "一事一议达线报区；用本 PPT+Excel 作请示附件"),
]
for i, (n, t) in enumerate(steps):
    y = Inches(1.5) + i * Inches(0.95)
    rect(s, Inches(1.35), y, Inches(10.6), Inches(0.8), CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, Inches(1.35), y, Inches(0.85), Inches(0.8), TEAL if i < 3 else GOLD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, Inches(1.35), y, Inches(0.85), Inches(0.8),
         [[(n, 22, WHITE, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(2.45), y + Inches(0.18), Inches(9.2), Inches(0.5),
         [[(t, 17, WHITE, True)]])

# ===== 24 封底 =====
s = add_slide(NAVY)
rect(s, 0, 0, Inches(0.18), SH, TEAL)
rect(s, 0, Inches(3.2), SW, Pt(3), TEAL)
text(s, Inches(0.9), Inches(1.7), Inches(11.5), Inches(1.4),
     [[(C.PROJECT_NAME, 32, WHITE, True)],
      [("锚定全球创客岛，答卷量子城市", 16, GOLD, True)]],
     align=PP_ALIGN.CENTER, space_after=10)
text(s, Inches(0.9), Inches(3.4), Inches(11.5), Inches(2.1),
     [[(C.PROJECT_FULL, 14, LIGHT, True)],
      [(C.EVENT_DATE + "　｜　服务 9/15 收官节点", 14, TEAL2, True)],
      [("优刻得 · 智谱 · 苏度 · 卓益得等杨浦企业同台互动", 13, LIGHT, False)],
      [("主呈现：PPT 策划方案  +  Excel 执行台账　｜　" + C.VERSION, 12, GREY, False)]],
     align=PP_ALIGN.CENTER, space_after=5)

prs.save(OUT_FILE)
print(f"已生成：{OUT_FILE}（{TOTAL} 页）")
