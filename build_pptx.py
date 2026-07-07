# -*- coding: utf-8 -*-
"""生成《下半年三方合作计划》汇报 PPT（16:9）。"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import plan_data as D

NAVY = RGBColor(0x1F, 0x38, 0x64)
BLUE = RGBColor(0x2E, 0x54, 0x96)
LTBLUE = RGBColor(0xD9, 0xE1, 0xF2)
STEEL = RGBColor(0x8F, 0xAA, 0xDC)
GOLD = RGBColor(0xBF, 0x90, 0x00)
GREY = RGBColor(0x59, 0x59, 0x59)
LGREY = RGBColor(0xF2, 0xF2, 0xF2)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x26, 0x26, 0x26)

FONT = "Microsoft YaHei"
EMU = 914400
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def slide():
    return prs.slides.add_slide(BLANK)


def _set_font(run, size, bold, color, font=FONT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    r = run._r
    rPr = r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', font)


def rect(s, x, y, w, h, fill=None, line=None, line_w=None, shadow=False, shape=MSO_SHAPE.RECTANGLE):
    sp = s.shapes.add_shape(shape, x, y, w, h)
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w or 1)
    sp.shadow.inherit = False
    if shadow:
        el = sp._element.spPr
        ef = el.makeelement(qn('a:effectLst'), {})
        sh = ef.makeelement(qn('a:outerShdw'),
                            {'blurRad': '40000', 'dist': '20000', 'dir': '5400000', 'rotWithShape': '0'})
        clr = sh.makeelement(qn('a:srgbClr'), {'val': 'A6A6A6'})
        alpha = clr.makeelement(qn('a:alpha'), {'val': '45000'})
        clr.append(alpha)
        sh.append(clr)
        ef.append(sh)
        el.append(ef)
    return sp


def text(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         space_after=4, line_spacing=1.0, wrap=True):
    """runs: list of paragraphs; each paragraph is list of (txt,size,bold,color)."""
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = Pt(2); tf.margin_right = Pt(2)
    tf.margin_top = Pt(1); tf.margin_bottom = Pt(1)
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        p.space_before = Pt(0)
        p.line_spacing = line_spacing
        for (txt, size, bold, color) in para:
            r = p.add_run()
            r.text = txt
            _set_font(r, size, bold, color)
    return tb


def bg(s, color=WHITE):
    rect(s, 0, 0, SW, SH, fill=color)


def header(s, kicker, title, idx=None):
    rect(s, 0, 0, SW, Inches(1.15), fill=NAVY)
    rect(s, 0, Inches(1.15), SW, Pt(3), fill=GOLD)
    text(s, Inches(0.55), Inches(0.16), Inches(11.5), Inches(0.32),
         [[(kicker, 11, True, STEEL)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(0.55), Inches(0.44), Inches(11.5), Inches(0.62),
         [[(title, 25, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)
    if idx:
        text(s, Inches(12.1), Inches(0.16), Inches(0.9), Inches(0.9),
             [[(idx, 30, True, RGBColor(0x3A,0x52,0x80))]], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def footer(s, n):
    text(s, Inches(0.55), Inches(7.08), Inches(9), Inches(0.3),
         [[("东方枢纽 × 复旦大学住房政策研究中心 × 杨浦区科技企业联合会  |  下半年合作计划", 8, False, GREY)]])
    text(s, Inches(12.3), Inches(7.08), Inches(0.8), Inches(0.3),
         [[(str(n), 8, False, GREY)]], align=PP_ALIGN.RIGHT)


PAGE = [0]
def bullet_card(s, x, y, w, h, title, lines, accent=BLUE, title_color=WHITE):
    rect(s, x, y, w, h, fill=WHITE, line=RGBColor(0xD0,0xD8,0xE8), line_w=1, shadow=True)
    rect(s, x, y, w, Inches(0.42), fill=accent)
    text(s, x+Inches(0.12), y, w-Inches(0.24), Inches(0.42),
         [[(title, 12, True, title_color)]], anchor=MSO_ANCHOR.MIDDLE)
    runs = [[("• ", 10, True, GOLD), (ln, 10, False, DARK)] for ln in lines]
    text(s, x+Inches(0.14), y+Inches(0.5), w-Inches(0.28), h-Inches(0.6),
         runs, space_after=4, line_spacing=1.02)


# ============================================================ 1. 封面
s = slide()
rect(s, 0, 0, SW, SH, fill=NAVY)
rect(s, 0, 0, SW, Inches(2.6), fill=BLUE)
rect(s, 0, Inches(2.6), SW, Pt(4), fill=GOLD)
# decorative blocks
for i, cx in enumerate([Inches(9.7), Inches(11.0), Inches(12.3)]):
    rect(s, cx, Inches(4.7), Inches(1.05), Inches(1.9), fill=RGBColor(0x2A,0x46,0x7A))
text(s, Inches(0.7), Inches(0.55), Inches(12), Inches(0.5),
     [[("三方战略合作 · 下半年活动计划", 15, True, STEEL)]])
text(s, Inches(0.7), Inches(1.15), Inches(12), Inches(1.5),
     [[("东方枢纽 133 万方招商", 40, True, WHITE)],
      [("12 场高质量闭门活动策划方案", 40, True, WHITE)]], space_after=2)
text(s, Inches(0.7), Inches(3.15), Inches(12), Inches(0.5),
     [[("复旦大学住房政策研究中心  ·  上海市杨浦区科技企业联合会  ·  东方枢纽", 15, True, WHITE)]])
text(s, Inches(0.7), Inches(3.75), Inches(11.5), Inches(0.5),
     [[("活动规划 · 时间建议 · 主题内容 · 详细策划 · 初步报价 · 招商引资价值分析", 13, False, STEEL)]])
text(s, Inches(0.7), Inches(6.6), Inches(12), Inches(0.5),
     [[("2026 年下半年（7–12 月）  |  内部审阅 & 客户汇报通用版", 11, False, STEEL)]])

# ============================================================ 2. 目录
s = slide(); bg(s); header(s, "AGENDA", "汇报目录", "01")
items = [
    ("01", "合作背景与三方定位", "谁在合作 · 各自角色"),
    ("02", "战略目标与合作原则", "小而美 · 重质量 · 闭环"),
    ("03", "六大核心产业与 12 场总览", "时间 · 主题 · 规模 · 报价"),
    ("04", "12 场活动详细策划案", "内容 · 嘉宾 · 招商衔接"),
    ("05", "报价体系", "三档模型 · 全年预算"),
    ("06", "招商引资价值分析（核心）", "133 万方去化转化"),
    ("07", "执行机制与下一步", "OA · 供应商 · 落地"),
]
cx = [Inches(0.6), Inches(6.9)]
for i, (num, t, sub) in enumerate(items):
    col = i % 2; row = i // 2
    x = cx[col]; y = Inches(1.55) + row * Inches(1.06)
    w = Inches(5.85); h = Inches(0.92)
    rect(s, x, y, w, h, fill=LGREY, line=RGBColor(0xD9,0xD9,0xD9), line_w=0.75)
    rect(s, x, y, Inches(0.95), h, fill=BLUE)
    text(s, x, y, Inches(0.95), h, [[(num, 22, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+Inches(1.1), y+Inches(0.12), w-Inches(1.2), Inches(0.4), [[(t, 15, True, NAVY)]])
    text(s, x+Inches(1.1), y+Inches(0.52), w-Inches(1.2), Inches(0.35), [[(sub, 10, False, GREY)]])
footer(s, 2)

# ============================================================ 3. 三方定位
s = slide(); bg(s); header(s, "背景 · BACKGROUND", "合作背景与三方定位", "01")
text(s, Inches(0.55), Inches(1.35), Inches(12.3), Inches(0.5),
     [[("依托东方枢纽资源，三方联合策划一系列聚焦六大产业、面向企业出海与政府飞地招商的高质量闭门活动。", 12, False, GREY)]])
colors = [BLUE, RGBColor(0x37,0x6A,0x4E), GOLD]
xs = [Inches(0.55), Inches(4.85), Inches(9.15)]
for i, (name, desc) in enumerate(D.PARTIES):
    x = xs[i]; y = Inches(2.05); w = Inches(3.9); h = Inches(4.4)
    rect(s, x, y, w, h, fill=WHITE, line=RGBColor(0xD0,0xD8,0xE8), line_w=1, shadow=True)
    rect(s, x, y, w, Inches(1.0), fill=colors[i])
    text(s, x+Inches(0.2), y+Inches(0.1), w-Inches(0.4), Inches(0.8),
         [[(name, 15, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+Inches(0.22), y+Inches(1.2), w-Inches(0.44), h-Inches(1.4),
         [[(desc, 12, False, DARK)]], line_spacing=1.25)
footer(s, 3)

# ============================================================ 4. 战略目标与原则
s = slide(); bg(s); header(s, "战略 · STRATEGY", "战略目标与六大合作原则", "02")
text(s, Inches(0.55), Inches(1.35), Inches(12.3), Inches(0.45),
     [[("总目标：以高质量活动为抓手，为东方枢纽 133 万方项目实现精准导流与高效去化转化。", 13, True, NAVY)]])
for i, (name, desc) in enumerate(D.PRINCIPLES):
    col = i % 2; row = i // 2
    x = Inches(0.55) + col * Inches(6.3); y = Inches(1.95) + row * Inches(1.55)
    w = Inches(6.0); h = Inches(1.4)
    rect(s, x, y, w, h, fill=LGREY, line=RGBColor(0xDD,0xE2,0xEC), line_w=0.75)
    rect(s, x, y, Pt(5), h, fill=GOLD)
    text(s, x+Inches(0.2), y+Inches(0.12), w-Inches(0.35), Inches(0.4), [[(name, 13, True, BLUE)]])
    text(s, x+Inches(0.2), y+Inches(0.52), w-Inches(0.35), Inches(0.85), [[(desc, 10.5, False, DARK)]], line_spacing=1.1)
footer(s, 4)

# ============================================================ 5. 六大产业
s = slide(); bg(s); header(s, "产业 · SECTORS", "东方枢纽六大核心产业板块", "03")
text(s, Inches(0.55), Inches(1.32), Inches(12), Inches(0.4),
     [[("12 场活动全部围绕以下六大板块定向邀约，而非泛化铺量。", 12, False, GREY)]])
icons = ["高服", "国贸", "民航", "医药", "工联", "绿能"]
for i, (name, desc) in enumerate(D.INDUSTRIES):
    col = i % 3; row = i // 3
    x = Inches(0.55) + col * Inches(4.2); y = Inches(1.9) + row * Inches(2.45)
    w = Inches(3.95); h = Inches(2.2)
    rect(s, x, y, w, h, fill=WHITE, line=RGBColor(0xD0,0xD8,0xE8), line_w=1, shadow=True)
    rect(s, x, y, w, Inches(0.12), fill=BLUE)
    circ = rect(s, x+Inches(0.25), y+Inches(0.35), Inches(0.9), Inches(0.9), fill=LTBLUE, shape=MSO_SHAPE.OVAL)
    text(s, x+Inches(0.25), y+Inches(0.35), Inches(0.9), Inches(0.9), [[(icons[i], 15, True, NAVY)]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+Inches(1.3), y+Inches(0.45), w-Inches(1.45), Inches(0.8),
         [[(name, 14, True, NAVY)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+Inches(0.28), y+Inches(1.42), w-Inches(0.5), Inches(0.7),
         [[(desc, 10.5, False, GREY)]], line_spacing=1.1)
footer(s, 5)

# ============================================================ 6. 12场总览表
s = slide(); bg(s); header(s, "总览 · OVERVIEW", "下半年 12 场活动总览", "03")
cols = [("#", 0.5), ("拟定时间", 1.85), ("活动主题", 4.35), ("产业板块", 2.05), ("规模", 1.35), ("档", 0.55), ("报价\n万元", 0.85)]
x0 = Inches(0.5); y0 = Inches(1.35); rowh = Inches(0.44)
cxs = []; acc = x0
for name, w in cols:
    cxs.append((acc, Inches(w))); acc += Inches(w)
# header
for (cx, cw), (name, _) in zip(cxs, cols):
    rect(s, cx, y0, cw, rowh, fill=NAVY)
    text(s, cx, y0, cw, rowh, [[(name, 9.5, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
for i, a in enumerate(D.ACTIVITIES):
    y = y0 + rowh * (i + 1)
    fill = LGREY if i % 2 else WHITE
    tierc = a["tier"].split(" ")[0]
    tcolor = {"A": RGBColor(0x37,0x6A,0x4E), "B": BLUE, "C": GOLD}[tierc]
    vals = [str(a["no"]), a["date"], a["title"], a["sector"], a["scale"].replace(" · ", "·"), tierc, str(a["price"])]
    for j, ((cx, cw), v) in enumerate(zip(cxs, vals)):
        rect(s, cx, y, cw, rowh, fill=fill, line=RGBColor(0xD9,0xD9,0xD9), line_w=0.5)
        al = PP_ALIGN.LEFT if j in (2, 3) else PP_ALIGN.CENTER
        col = tcolor if j == 5 else (GOLD if j == 6 else DARK)
        bold = j in (5, 6)
        pad = Inches(0.08) if j in (2,3) else 0
        text(s, cx+pad, y, cw-pad, rowh, [[(v, 8.6, bold, col)]], align=al, anchor=MSO_ANCHOR.MIDDLE)
# total
yt = y0 + rowh * (len(D.ACTIVITIES)+1)
rect(s, cxs[0][0], yt, sum(c[1] for c in cxs[:6]) if False else (cxs[5][0]+cxs[5][1]-cxs[0][0]), rowh, fill=LTBLUE)
text(s, cxs[0][0], yt, cxs[5][0]+cxs[5][1]-cxs[0][0], rowh, [[("12 场合计（初步报价）", 10, True, NAVY)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
rect(s, cxs[6][0], yt, cxs[6][1], rowh, fill=GOLD)
text(s, cxs[6][0], yt, cxs[6][1], rowh, [[(str(D.TOTAL_PRICE), 11, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, Inches(0.5), Inches(7.02), Inches(12), Inches(0.35),
     [[("档位：A 小而美闭门(≤30人·区内) / B 市区专场(≤80人) / C 旗舰论坛(≤200人)。节奏建议两周一场，资源成熟可提档、档期可调。", 8.5, False, GREY)]])

# ============================================================ 7-12. 逐场详细（2/页）
def act_card(s, x, y, w, h, a):
    tierc = a["tier"].split(" ")[0]
    accent = {"A": RGBColor(0x37,0x6A,0x4E), "B": BLUE, "C": GOLD}[tierc]
    rect(s, x, y, w, h, fill=WHITE, line=RGBColor(0xD0,0xD8,0xE8), line_w=1, shadow=True)
    rect(s, x, y, w, Inches(0.72), fill=accent)
    text(s, x+Inches(0.18), y+Inches(0.06), w-Inches(1.4), Inches(0.62),
         [[(f"NO.{a['no']}  {a['title']}", 12.5, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x+w-Inches(1.15), y+Inches(0.14), Inches(0.98), Inches(0.44), fill=WHITE)
    text(s, x+w-Inches(1.15), y+Inches(0.14), Inches(0.98), Inches(0.44),
         [[(f"{a['price']}万", 12, True, accent)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+Inches(0.18), y+Inches(0.8), w-Inches(0.3), Inches(0.3),
         [[("⏰ "+a["date"]+"    👥 "+a["scale"]+"    📍 "+a["venue"], 9, True, GGREY)]] if False else
         [[(a["date"]+"  |  "+a["scale"]+"  |  "+a["venue"], 9, True, GREY)]])
    yy = y+Inches(1.15)
    text(s, x+Inches(0.18), yy, w-Inches(0.3), Inches(0.25), [[("■ 内容建议", 9.5, True, accent)]])
    runs = [[("· ", 9, True, GOLD), (ln, 9, False, DARK)] for ln in a["content"]]
    text(s, x+Inches(0.2), yy+Inches(0.25), w-Inches(0.35), Inches(1.1), runs, space_after=2, line_spacing=1.0)
    yy2 = yy+Inches(1.5)
    text(s, x+Inches(0.18), yy2, w-Inches(0.3), Inches(0.25), [[("■ 拟邀嘉宾/资源", 9.5, True, accent)]])
    text(s, x+Inches(0.2), yy2+Inches(0.23), w-Inches(0.35), Inches(0.5),
         [[("、".join(a["guests"]), 8.7, False, DARK)]], line_spacing=1.0)
    yy3 = yy2+Inches(0.95)
    text(s, x+Inches(0.18), yy3, w-Inches(0.3), Inches(0.25), [[("■ 招商衔接价值", 9.5, True, accent)]])
    runs = [[("· ", 9, True, GOLD), (ln, 8.7, False, DARK)] for ln in a["invest"]]
    text(s, x+Inches(0.2), yy3+Inches(0.23), w-Inches(0.35), Inches(0.8), runs, space_after=1, line_spacing=1.0)

GGREY = GREY
pg = 7
for i in range(0, 12, 2):
    s = slide(); bg(s)
    header(s, "详细策划 · PLAYBOOK", f"12 场活动详细策划案（{i+1}–{i+2} / 12）", "04")
    act_card(s, Inches(0.5), Inches(1.4), Inches(6.05), Inches(5.35), D.ACTIVITIES[i])
    act_card(s, Inches(6.78), Inches(1.4), Inches(6.05), Inches(5.35), D.ACTIVITIES[i+1])
    footer(s, pg); pg += 1

# ============================================================ 13. 报价体系
s = slide(); bg(s); header(s, "报价 · QUOTE", "报价体系 · 三档模型与全年预算", "05")
tier_keys = list(D.TIERS.keys())
# 三档卡片
tcolors = [RGBColor(0x37,0x6A,0x4E), BLUE, GOLD]
for i, k in enumerate(tier_keys):
    x = Inches(0.5) + i * Inches(4.25); y = Inches(1.35); w = Inches(3.95); h = Inches(3.35)
    rect(s, x, y, w, h, fill=WHITE, line=RGBColor(0xD0,0xD8,0xE8), line_w=1, shadow=True)
    rect(s, x, y, w, Inches(0.95), fill=tcolors[i])
    parts = k.split(" · ")
    text(s, x+Inches(0.15), y+Inches(0.08), w-Inches(0.3), Inches(0.45), [[(parts[0], 14, True, WHITE)]])
    text(s, x+Inches(0.15), y+Inches(0.5), w-Inches(0.3), Inches(0.4), [[(parts[1] if len(parts)>1 else "", 10.5, False, WHITE)]])
    text(s, x+Inches(0.15), y+Inches(1.02), w-Inches(0.3), Inches(0.5),
         [[(f"{D.TIER_TOTAL[k]} ", 24, True, tcolors[i]), ("万元/场", 11, True, GREY)]])
    runs = []
    for item, v in zip(D.COST_ITEMS, D.TIERS[k]):
        runs.append([(f"{item}", 9, False, DARK), (f"  {v}", 9, True, GOLD)])
    text(s, x+Inches(0.18), y+Inches(1.6), w-Inches(0.32), Inches(1.7), runs, space_after=2, line_spacing=1.0)
# 汇总条
tier_counts = {}
for a in D.ACTIVITIES:
    tier_counts[a["tier"]] = tier_counts.get(a["tier"], 0) + 1
y = Inches(4.95)
rect(s, Inches(0.5), y, Inches(12.33), Inches(1.75), fill=LGREY, line=RGBColor(0xD9,0xD9,0xD9), line_w=0.75)
text(s, Inches(0.7), y+Inches(0.1), Inches(6), Inches(0.4), [[("全年预算汇总", 13, True, NAVY)]])
seg = [("A 档 × %d 场" % tier_counts[tier_keys[0]], round(tier_counts[tier_keys[0]]*D.TIER_TOTAL[tier_keys[0]],1)),
       ("B 档 × %d 场" % tier_counts[tier_keys[1]], round(tier_counts[tier_keys[1]]*D.TIER_TOTAL[tier_keys[1]],1)),
       ("C 档 × %d 场" % tier_counts[tier_keys[2]], round(tier_counts[tier_keys[2]]*D.TIER_TOTAL[tier_keys[2]],1))]
for i, (lab, val) in enumerate(seg):
    x = Inches(0.7) + i*Inches(3.15); yy = y+Inches(0.65)
    rect(s, x, yy, Inches(2.9), Inches(0.85), fill=WHITE, line=tcolors[i], line_w=1.2)
    text(s, x, yy+Inches(0.08), Inches(2.9), Inches(0.35), [[(lab, 11, True, tcolors[i])]], align=PP_ALIGN.CENTER)
    text(s, x, yy+Inches(0.42), Inches(2.9), Inches(0.4), [[(f"{val} 万元", 13, True, NAVY)]], align=PP_ALIGN.CENTER)
x = Inches(0.7) + 3*Inches(3.15); yy = y+Inches(0.65)
rect(s, x, yy, Inches(2.6), Inches(0.85), fill=NAVY)
text(s, x, yy+Inches(0.08), Inches(2.6), Inches(0.35), [[("12 场合计", 11, True, STEEL)]], align=PP_ALIGN.CENTER)
text(s, x, yy+Inches(0.42), Inches(2.6), Inches(0.4), [[(f"{D.TOTAL_PRICE} 万元", 14, True, WHITE)]], align=PP_ALIGN.CENTER)
footer(s, pg); pg += 1

# ============================================================ 14. 招商漏斗（核心）
s = slide(); bg(s); header(s, "价值 · 核心", "招商引资价值分析（一）转化漏斗", "06")
text(s, Inches(0.55), Inches(1.3), Inches(12.3), Inches(0.4),
     [[("核心利好：12 场活动为东方枢纽 133 万方项目构建“引流—深挖—对接—签约”的可量化招商漏斗。", 12, True, NAVY)]])
funnel_widths = [Inches(9.0), Inches(7.4), Inches(5.8), Inches(4.2), Inches(3.0)]
fcolors = [BLUE, RGBColor(0x35,0x5A,0x9A), RGBColor(0x3E,0x6A,0x9E), RGBColor(0x37,0x6A,0x4E), GOLD]
y = Inches(1.95)
for i, ((stage, basis, val), fw) in enumerate(zip(D.FUNNEL, funnel_widths)):
    x = Inches(0.7) + (Inches(9.0)-fw)/2
    rect(s, x, y, fw, Inches(0.72), fill=fcolors[i], shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x, y, fw, Inches(0.72), [[(f"{stage}", 11, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # right label
    text(s, Inches(10.0), y, Inches(2.9), Inches(0.72),
         [[(val, 15, True, GOLD)], [(basis, 8.5, False, GREY)]], anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.9)
text(s, Inches(0.7), Inches(6.75), Inches(12), Inches(0.35),
     [[("注：转化率为测算假设，用于展示价值逻辑；实际以邀约质量与后续跟进为准。", 8.5, False, GREY)]])
footer(s, pg); pg += 1

# ============================================================ 15. 去化测算 + 支柱
s = slide(); bg(s); header(s, "价值 · 核心", "招商引资价值分析（二）去化测算与价值支柱", "06")
# 去化情景
text(s, Inches(0.55), Inches(1.32), Inches(6), Inches(0.4), [[("133 万方 × 情景去化率测算", 13, True, NAVY)]])
for i, (name, rate, area) in enumerate(D.GMV_SCENARIOS):
    x = Inches(0.55) + i*Inches(2.05); y = Inches(1.8)
    col = [STEEL, BLUE, GOLD][i]
    rect(s, x, y, Inches(1.9), Inches(1.75), fill=WHITE, line=col, line_w=1.3, shadow=True)
    rect(s, x, y, Inches(1.9), Inches(0.42), fill=col)
    text(s, x, y, Inches(1.9), Inches(0.42), [[(name+"情景", 11, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x, y+Inches(0.55), Inches(1.9), Inches(0.4), [[(rate, 20, True, col)]], align=PP_ALIGN.CENTER)
    text(s, x, y+Inches(1.05), Inches(1.9), Inches(0.6), [[("去化", 9, False, GREY)],[(area, 11, True, NAVY)]], align=PP_ALIGN.CENTER)
# ROI 概览 右侧
text(s, Inches(6.95), Inches(1.32), Inches(6), Inches(0.4), [[("投入产出概览", 13, True, NAVY)]])
for i, (name, val) in enumerate(D.ROI_SUMMARY):
    y = Inches(1.8) + i*Inches(0.38)
    rect(s, Inches(6.95), y, Inches(5.85), Inches(0.32), fill=LGREY, line=RGBColor(0xDD,0xE2,0xEC), line_w=0.5)
    rect(s, Inches(6.95), y, Pt(4), Inches(0.32), fill=GOLD)
    text(s, Inches(7.15), y, Inches(3.3), Inches(0.32), [[(name, 9.5, False, DARK)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(10.3), y, Inches(2.4), Inches(0.32), [[(val, 10.5, True, BLUE)]], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
# 六大价值支柱
text(s, Inches(0.55), Inches(3.9), Inches(6), Inches(0.4), [[("六大价值支柱", 13, True, NAVY)]])
for i, (name, desc) in enumerate(D.VALUE_PILLARS):
    col = i % 3; row = i // 3
    x = Inches(0.55) + col*Inches(4.12); y = Inches(4.35) + row*Inches(1.35)
    w = Inches(3.9); h = Inches(1.2)
    rect(s, x, y, w, h, fill=WHITE, line=RGBColor(0xD0,0xD8,0xE8), line_w=1, shadow=True)
    text(s, x+Inches(0.15), y+Inches(0.08), w-Inches(0.3), Inches(0.32), [[(f"{i+1}. {name}", 11, True, BLUE)]])
    text(s, x+Inches(0.15), y+Inches(0.42), w-Inches(0.3), Inches(0.75), [[(desc, 8.6, False, DARK)]], line_spacing=1.0)
footer(s, pg); pg += 1

# ============================================================ 16. 执行机制
s = slide(); bg(s); header(s, "执行 · EXECUTION", "执行机制与合规保障", "07")
for i, (name, desc) in enumerate(D.EXECUTION):
    col = i % 2; row = i // 2
    x = Inches(0.55) + col*Inches(6.3); y = Inches(1.5) + row*Inches(1.75)
    w = Inches(6.0); h = Inches(1.55)
    rect(s, x, y, w, h, fill=LGREY, line=RGBColor(0xDD,0xE2,0xEC), line_w=0.75)
    rect(s, x, y, Inches(1.0), h, fill=BLUE)
    text(s, x, y, Inches(1.0), h, [[(f"{i+1:02d}", 22, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+Inches(1.15), y+Inches(0.15), w-Inches(1.3), Inches(0.4), [[(name, 13, True, NAVY)]])
    text(s, x+Inches(1.15), y+Inches(0.58), w-Inches(1.3), Inches(0.9), [[(desc, 10, False, DARK)]], line_spacing=1.1)
footer(s, pg); pg += 1

# ============================================================ 17. 下一步
s = slide()
rect(s, 0, 0, SW, SH, fill=NAVY)
rect(s, 0, Inches(2.3), SW, Pt(4), fill=GOLD)
text(s, Inches(0.7), Inches(0.9), Inches(12), Inches(0.6), [[("下一步 · NEXT STEPS", 15, True, STEEL)]])
text(s, Inches(0.7), Inches(1.4), Inches(12), Inches(0.8), [[("确认试点场次，即刻启动", 34, True, WHITE)]])
steps = [
    ("① 确认首场", "敲定 7 月中旬“企业出海·进境关外”试点主题、时间与 30 人以内精准邀约名单"),
    ("② 提交 OA 要件", "邀约名单 / 参会人数 / 人员背景 / 预算 / 时间，供东方枢纽走内部 OA 流程"),
    ("③ 供应商签约", "经指定策划供应商签约支付，落实国货区场地搭建通行证"),
    ("④ 滚动推进", "两周一场滚动执行，形成“活动—跟进—评估—再合作”长期闭环"),
]
for i, (t, d) in enumerate(steps):
    y = Inches(2.75) + i*Inches(1.02)
    rect(s, Inches(0.7), y, Inches(11.9), Inches(0.88), fill=RGBColor(0x2A,0x46,0x7A))
    text(s, Inches(0.9), y, Inches(3.0), Inches(0.88), [[(t, 15, True, GOLD)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(3.9), y, Inches(8.5), Inches(0.88), [[(d, 11.5, False, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)
text(s, Inches(0.7), Inches(6.95), Inches(12), Inches(0.4),
     [[("复旦大学住房政策研究中心  ·  上海市杨浦区科技企业联合会  ·  东方枢纽    |    合作愉快", 11, True, STEEL)]])

prs.save("东方枢纽三方合作计划_汇报PPT.pptx")
print("PPT 已生成：东方枢纽三方合作计划_汇报PPT.pptx  (页数: %d)" % len(prs.slides._sldIdLst))
