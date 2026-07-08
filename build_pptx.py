# -*- coding: utf-8 -*-
"""生成《下半年三方合作计划》汇报 PPT（16:9）· 商务紫主题。"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import plan_data as D

# ---- 商务紫配色 ----
PLUM   = RGBColor(0x2E, 0x1F, 0x47)   # 最深紫
PURPLE = RGBColor(0x5B, 0x3E, 0x8E)   # 主紫
PURPLE2= RGBColor(0x76, 0x58, 0xA8)   # 中紫
LILAC  = RGBColor(0xC0, 0xAE, 0xE0)   # 浅紫（kicker）
LAV    = RGBColor(0xEC, 0xE6, 0xF7)   # 浅紫填充
LAVED  = RGBColor(0xDD, 0xD2, 0xEF)   # 卡片描边
BGT    = RGBColor(0xFB, 0xFA, 0xFE)   # 页面底色
GOLD   = RGBColor(0xC1, 0x9A, 0x3A)   # 金
GREY   = RGBColor(0x60, 0x5A, 0x6B)
DARK   = RGBColor(0x2B, 0x25, 0x36)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
TIER_C = {"A": RGBColor(0x86, 0x6F, 0xB4), "B": PURPLE, "C": PLUM}

FONT = "Microsoft YaHei"
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def slide():
    return prs.slides.add_slide(BLANK)


def _set_font(run, size, bold, color, font=FONT):
    run.font.size = Pt(size); run.font.bold = bold
    run.font.color.rgb = color; run.font.name = font
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {}); rPr.append(ea)
    ea.set('typeface', font)


def rect(s, x, y, w, h, fill=None, line=None, line_w=None, shadow=False, shape=MSO_SHAPE.RECTANGLE):
    sp = s.shapes.add_shape(shape, x, y, w, h)
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(line_w or 1)
    sp.shadow.inherit = False
    if shadow:
        el = sp._element.spPr
        ef = el.makeelement(qn('a:effectLst'), {})
        sh = ef.makeelement(qn('a:outerShdw'),
                            {'blurRad': '55000', 'dist': '25000', 'dir': '5400000', 'rotWithShape': '0'})
        clr = sh.makeelement(qn('a:srgbClr'), {'val': '7A6BA0'})
        alpha = clr.makeelement(qn('a:alpha'), {'val': '38000'})
        clr.append(alpha); sh.append(clr); ef.append(sh); el.append(ef)
    return sp


def text(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         space_after=4, line_spacing=1.0, wrap=True):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap; tf.vertical_anchor = anchor
    tf.margin_left = Pt(2); tf.margin_right = Pt(2); tf.margin_top = Pt(1); tf.margin_bottom = Pt(1)
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.space_after = Pt(space_after); p.space_before = Pt(0); p.line_spacing = line_spacing
        for (txt, size, bold, color) in para:
            r = p.add_run(); r.text = txt; _set_font(r, size, bold, color)
    return tb


def bg(s):
    rect(s, 0, 0, SW, SH, fill=BGT)


def header(s, kicker, title, idx=None):
    rect(s, 0, 0, SW, Inches(1.18), fill=PLUM)
    rect(s, 0, Inches(1.18), SW, Pt(3), fill=GOLD)
    rect(s, Inches(0.55), Inches(0.5), Pt(4), Inches(0.5), fill=GOLD)   # 竖向金色小饰条
    text(s, Inches(0.72), Inches(0.18), Inches(11.4), Inches(0.32),
         [[(kicker, 11, True, LILAC)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(0.72), Inches(0.46), Inches(11.4), Inches(0.62),
         [[(title, 25, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)
    if idx:
        text(s, Inches(11.9), Inches(0.12), Inches(1.1), Inches(0.95),
             [[(idx, 34, True, PURPLE2)]], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def footer(s, n):
    rect(s, Inches(0.55), Inches(7.06), Inches(12.25), Pt(0.75), fill=LAVED)
    text(s, Inches(0.55), Inches(7.1), Inches(10), Inches(0.3),
         [[("东方枢纽 × 复旦大学 × 上海市科技企业联合会  |  下半年合作计划", 8, False, GREY)]])
    text(s, Inches(12.2), Inches(7.1), Inches(0.8), Inches(0.3),
         [[(str(n), 8, True, PURPLE)]], align=PP_ALIGN.RIGHT)


# ============================================================ 1. 封面
s = slide()
rect(s, 0, 0, SW, SH, fill=PLUM)
rect(s, 0, 0, SW, Inches(2.55), fill=PURPLE)
rect(s, 0, Inches(2.55), SW, Pt(4), fill=GOLD)
# 装饰细线方块
for cx in [Inches(9.55), Inches(10.9), Inches(12.25)]:
    rect(s, cx, Inches(4.55), Inches(1.05), Inches(2.0), fill=None, line=RGBColor(0x53,0x40,0x7C), line_w=1.2)
rect(s, Inches(12.25), Inches(4.55), Inches(1.05), Inches(2.0), fill=RGBColor(0x3A,0x27,0x5C))
text(s, Inches(0.72), Inches(0.55), Inches(12), Inches(0.5),
     [[("三方战略合作 · 下半年活动计划", 15, True, LILAC)]])
text(s, Inches(0.72), Inches(1.12), Inches(12), Inches(1.5),
     [[("东方枢纽 133 万方招商", 40, True, WHITE)],
      [("12 场高质量闭门活动策划方案", 40, True, WHITE)]], space_after=2)
rect(s, Inches(0.75), Inches(3.02), Inches(2.0), Pt(3), fill=GOLD)
text(s, Inches(0.72), Inches(3.2), Inches(12), Inches(0.5),
     [[("复旦大学  ·  上海市科技企业联合会  ·  东方枢纽", 15, True, WHITE)]])
text(s, Inches(0.72), Inches(3.78), Inches(11.8), Inches(0.5),
     [[("市级 + 浦东新区资源联动  |  活动规划 · 主题内容 · 详细策划 · 合理报价 · 招商价值", 12.5, False, LILAC)]])
text(s, Inches(0.72), Inches(6.62), Inches(12), Inches(0.5),
     [[("2026 年下半年（7–12 月）  |  内部审阅 & 客户汇报通用版", 11, False, LILAC)]])

# ============================================================ 2. 目录
s = slide(); bg(s); header(s, "AGENDA", "汇报目录")
items = [
    ("01", "合作背景与三方定位", "谁在合作 · 各自角色"),
    ("02", "政府资源背书矩阵", "市级 + 浦东新区联动"),
    ("03", "战略目标与合作原则", "小而美 · 重质量 · 科创赋能"),
    ("04", "六大核心产业与 12 场总览", "时间 · 主题 · 规模 · 报价"),
    ("05", "12 场活动详细策划案", "内容 · 嘉宾 · 招商衔接"),
    ("06", "报价体系（优化下调）", "三档模型 · 全年预算"),
    ("07", "招商引资价值分析（核心）", "133 万方去化转化"),
    ("08", "执行机制与下一步", "OA · 供应商 · 落地"),
]
cx = [Inches(0.6), Inches(6.9)]
for i, (num, t, sub) in enumerate(items):
    col = i % 2; row = i // 2
    x = cx[col]; y = Inches(1.5) + row * Inches(1.35)
    w = Inches(5.85); h = Inches(1.12)
    rect(s, x, y, w, h, fill=WHITE, line=LAVED, line_w=1, shadow=True)
    rect(s, x, y, Inches(1.0), h, fill=PURPLE)
    text(s, x, y, Inches(1.0), h, [[(num, 24, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+Inches(1.18), y+Inches(0.22), w-Inches(1.3), Inches(0.45), [[(t, 15, True, PLUM)]])
    text(s, x+Inches(1.18), y+Inches(0.66), w-Inches(1.3), Inches(0.35), [[(sub, 10, False, GREY)]])
footer(s, 2)

# ============================================================ 3. 三方定位
s = slide(); bg(s); header(s, "背景 · BACKGROUND", "合作背景与三方定位", "01")
text(s, Inches(0.72), Inches(1.38), Inches(12), Inches(0.5),
     [[("以“上海市级 + 浦东新区”政府资源联动为核心，复旦大学提供学术与科技智库背书，三方联合策划高质量招商活动。", 12, False, GREY)]])
colors = [PURPLE, PURPLE2, PLUM]
xs = [Inches(0.6), Inches(4.87), Inches(9.14)]
for i, (name, role, desc) in enumerate(D.PARTIES):
    x = xs[i]; y = Inches(2.0); w = Inches(3.87); h = Inches(3.85)
    rect(s, x, y, w, h, fill=WHITE, line=LAVED, line_w=1, shadow=True)
    rect(s, x, y, w, Inches(1.05), fill=colors[i])
    text(s, x+Inches(0.22), y+Inches(0.14), w-Inches(0.44), Inches(0.5), [[(name, 15, True, WHITE)]])
    text(s, x+Inches(0.22), y+Inches(0.62), w-Inches(0.44), Inches(0.35), [[(role, 10.5, True, LILAC)]])
    text(s, x+Inches(0.24), y+Inches(1.24), w-Inches(0.48), h-Inches(1.4),
         [[(desc, 11.5, False, DARK)]], line_spacing=1.28)
rect(s, Inches(0.6), Inches(6.05), Pt(4), Inches(0.55), fill=GOLD)
text(s, Inches(0.78), Inches(6.05), Inches(12), Inches(0.55),
     [[("备选/补充科技组织：", 10.5, True, PURPLE), ("、".join(D.ALT_TECH_ORGS), 10.5, False, GREY)]],
     anchor=MSO_ANCHOR.MIDDLE)
footer(s, 3)

# ============================================================ 4. 政府资源背书矩阵
s = slide(); bg(s); header(s, "资源 · GOVERNMENT", "政府资源背书矩阵（市级 + 浦东新区）", "02")
text(s, Inches(0.72), Inches(1.4), Inches(12), Inches(0.45),
     [[("本次活动的核心逻辑：市级资源与浦东新区资源相结合，商务委与投资促进部门共同参与背书。", 12, True, PLUM)]])
groups = list(D.GOV_RESOURCES.items())
gcolors = [PURPLE, PLUM]
for i, (group, gitems) in enumerate(groups):
    x = Inches(0.6) + i * Inches(6.35); y = Inches(2.05); w = Inches(6.0); h = Inches(3.35)
    rect(s, x, y, w, h, fill=WHITE, line=LAVED, line_w=1, shadow=True)
    rect(s, x, y, w, Inches(0.72), fill=gcolors[i])
    text(s, x+Inches(0.25), y, w-Inches(0.4), Inches(0.72), [[(group, 16, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)
    for j, it in enumerate(gitems):
        yy = y + Inches(0.95) + j * Inches(0.72)
        rect(s, x+Inches(0.25), yy, w-Inches(0.5), Inches(0.58), fill=LAV)
        rect(s, x+Inches(0.25), yy, Pt(4), Inches(0.58), fill=GOLD)
        text(s, x+Inches(0.45), yy, w-Inches(0.7), Inches(0.58), [[(it, 12, True, DARK)]], anchor=MSO_ANCHOR.MIDDLE)
# 中间连接符
text(s, Inches(6.35), Inches(3.3), Inches(0.65), Inches(0.7), [[("＋", 30, True, GOLD)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
# 底部联动逻辑条
y = Inches(5.7)
rect(s, Inches(0.6), y, Inches(12.35), Inches(0.95), fill=PLUM)
rect(s, Inches(0.6), y, Pt(5), Inches(0.95), fill=GOLD)
text(s, Inches(0.9), y, Inches(11.8), Inches(0.95), [[(D.GOV_TAGLINE, 13, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)
footer(s, 4)

# ============================================================ 5. 战略目标与原则
s = slide(); bg(s); header(s, "战略 · STRATEGY", "战略目标与六大合作原则", "03")
text(s, Inches(0.72), Inches(1.38), Inches(12.3), Inches(0.45),
     [[("总目标：以高质量活动为抓手，为东方枢纽 133 万方项目实现精准导流与高效去化转化。", 13, True, PLUM)]])
for i, (name, desc) in enumerate(D.PRINCIPLES):
    col = i % 2; row = i // 2
    x = Inches(0.6) + col * Inches(6.3); y = Inches(1.95) + row * Inches(1.55)
    w = Inches(6.0); h = Inches(1.4)
    rect(s, x, y, w, h, fill=WHITE, line=LAVED, line_w=0.75, shadow=True)
    rect(s, x, y, Pt(5), h, fill=GOLD)
    text(s, x+Inches(0.22), y+Inches(0.14), w-Inches(0.35), Inches(0.4), [[(name, 13, True, PURPLE)]])
    text(s, x+Inches(0.22), y+Inches(0.54), w-Inches(0.38), Inches(0.85), [[(desc, 10.5, False, DARK)]], line_spacing=1.12)
footer(s, 5)

# ============================================================ 6. 六大产业
s = slide(); bg(s); header(s, "产业 · SECTORS", "东方枢纽六大核心产业板块", "04")
text(s, Inches(0.72), Inches(1.35), Inches(12), Inches(0.4),
     [[("12 场活动全部围绕以下六大板块定向邀约，而非泛化铺量。", 12, False, GREY)]])
icons = ["高服", "国贸", "民航", "医药", "工联", "绿能"]
for i, (name, desc) in enumerate(D.INDUSTRIES):
    col = i % 3; row = i // 3
    x = Inches(0.6) + col * Inches(4.18); y = Inches(1.9) + row * Inches(2.45)
    w = Inches(3.95); h = Inches(2.2)
    rect(s, x, y, w, h, fill=WHITE, line=LAVED, line_w=1, shadow=True)
    rect(s, x, y, w, Inches(0.12), fill=PURPLE)
    rect(s, x+Inches(0.25), y+Inches(0.35), Inches(0.9), Inches(0.9), fill=LAV, shape=MSO_SHAPE.OVAL)
    text(s, x+Inches(0.25), y+Inches(0.35), Inches(0.9), Inches(0.9), [[(icons[i], 15, True, PURPLE)]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+Inches(1.3), y+Inches(0.45), w-Inches(1.45), Inches(0.8), [[(name, 14, True, PLUM)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+Inches(0.28), y+Inches(1.42), w-Inches(0.5), Inches(0.7), [[(desc, 10.5, False, GREY)]], line_spacing=1.1)
footer(s, 6)

# ============================================================ 7. 12场总览表
s = slide(); bg(s); header(s, "总览 · OVERVIEW", "下半年 12 场活动总览", "04")
cols = [("#", 0.5), ("拟定时间", 1.85), ("活动主题", 4.3), ("产业板块", 2.05), ("规模", 1.4), ("档", 0.55), ("报价\n万元", 0.85)]
x0 = Inches(0.5); y0 = Inches(1.35); rowh = Inches(0.44)
cxs = []; acc = x0
for name, w in cols:
    cxs.append((acc, Inches(w))); acc += Inches(w)
for (cx, cw), (name, _) in zip(cxs, cols):
    rect(s, cx, y0, cw, rowh, fill=PLUM)
    text(s, cx, y0, cw, rowh, [[(name, 9.5, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
for i, a in enumerate(D.ACTIVITIES):
    y = y0 + rowh * (i + 1)
    fill = LAV if i % 2 else WHITE
    tc = a["tier"].split(" ")[0]
    vals = [str(a["no"]), a["date"], a["title"], a["sector"], a["scale"].replace(" · ", "·"), tc, str(a["price"])]
    for j, ((cx, cw), v) in enumerate(zip(cxs, vals)):
        rect(s, cx, y, cw, rowh, fill=fill, line=LAVED, line_w=0.5)
        al = PP_ALIGN.LEFT if j in (2, 3) else PP_ALIGN.CENTER
        col = TIER_C[tc] if j == 5 else (GOLD if j == 6 else DARK)
        bold = j in (5, 6)
        pad = Inches(0.08) if j in (2, 3) else 0
        text(s, cx+pad, y, cw-pad, rowh, [[(v, 8.6, bold, col)]], align=al, anchor=MSO_ANCHOR.MIDDLE)
yt = y0 + rowh * (len(D.ACTIVITIES)+1)
rect(s, cxs[0][0], yt, cxs[5][0]+cxs[5][1]-cxs[0][0], rowh, fill=LILAC)
text(s, cxs[0][0], yt, cxs[5][0]+cxs[5][1]-cxs[0][0], rowh, [[("12 场合计（初步报价）", 10, True, PLUM)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
rect(s, cxs[6][0], yt, cxs[6][1], rowh, fill=GOLD)
text(s, cxs[6][0], yt, cxs[6][1], rowh, [[(str(D.TOTAL_PRICE), 11, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.35),
     [[("档位：A 小而美闭门(≤30人·区内) / B 市区专场(≤80人) / C 旗舰论坛(≤200人)。节奏两周一场，资源成熟可提档、档期可调。", 8.5, False, GREY)]])

# ============================================================ 8-13. 逐场详细（2/页）
def act_card(s, x, y, w, h, a):
    tc = a["tier"].split(" ")[0]; accent = TIER_C[tc]
    rect(s, x, y, w, h, fill=WHITE, line=LAVED, line_w=1, shadow=True)
    rect(s, x, y, w, Inches(0.72), fill=accent)
    text(s, x+Inches(0.18), y+Inches(0.06), w-Inches(1.4), Inches(0.62),
         [[(f"NO.{a['no']}  {a['title']}", 12.5, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x+w-Inches(1.15), y+Inches(0.14), Inches(0.98), Inches(0.44), fill=WHITE)
    text(s, x+w-Inches(1.15), y+Inches(0.14), Inches(0.98), Inches(0.44),
         [[(f"{a['price']}万", 12, True, accent)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+Inches(0.18), y+Inches(0.8), w-Inches(0.3), Inches(0.3),
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

pg = 8
for i in range(0, 12, 2):
    s = slide(); bg(s)
    header(s, "详细策划 · PLAYBOOK", f"12 场活动详细策划案（{i+1}–{i+2} / 12）", "05")
    act_card(s, Inches(0.5), Inches(1.4), Inches(6.05), Inches(5.35), D.ACTIVITIES[i])
    act_card(s, Inches(6.78), Inches(1.4), Inches(6.05), Inches(5.35), D.ACTIVITIES[i+1])
    footer(s, pg); pg += 1

# ============================================================ 14. 报价体系
s = slide(); bg(s); header(s, "报价 · QUOTE", "报价体系 · 按东方枢纽标准优化下调", "06")
tier_keys = list(D.TIERS.keys())
for i, k in enumerate(tier_keys):
    x = Inches(0.5) + i * Inches(4.25); y = Inches(1.35); w = Inches(3.95); h = Inches(3.35)
    rect(s, x, y, w, h, fill=WHITE, line=LAVED, line_w=1, shadow=True)
    rect(s, x, y, w, Inches(0.95), fill=TIER_C[k.split(" ")[0]])
    parts = k.split(" · ")
    text(s, x+Inches(0.15), y+Inches(0.1), w-Inches(0.3), Inches(0.45), [[(parts[0], 14, True, WHITE)]])
    text(s, x+Inches(0.15), y+Inches(0.52), w-Inches(0.3), Inches(0.4), [[(parts[1] if len(parts)>1 else "", 10.5, False, LILAC)]])
    text(s, x+Inches(0.15), y+Inches(1.02), w-Inches(0.3), Inches(0.5),
         [[(f"{D.TIER_TOTAL[k]} ", 24, True, PURPLE), ("万元/场", 11, True, GREY)]])
    runs = [[(item, 9, False, DARK), (f"  {v}", 9, True, GOLD)] for item, v in zip(D.COST_ITEMS, D.TIERS[k])]
    text(s, x+Inches(0.18), y+Inches(1.6), w-Inches(0.32), Inches(1.7), runs, space_after=2, line_spacing=1.0)
tier_counts = {}
for a in D.ACTIVITIES:
    tier_counts[a["tier"]] = tier_counts.get(a["tier"], 0) + 1
y = Inches(4.95)
rect(s, Inches(0.5), y, Inches(12.33), Inches(1.7), fill=LAV, line=LAVED, line_w=0.75)
drop = round((1 - D.TOTAL_PRICE / D.PREV_TOTAL) * 100)
text(s, Inches(0.7), y+Inches(0.12), Inches(11.8), Inches(0.4),
     [[("全年预算汇总", 13, True, PLUM), (f"    （较初版 {D.PREV_TOTAL} 万元下调约 {drop}%，上不封顶原则下的合理基准价）", 10, False, GREY)]])
tcolors = [TIER_C["A"], TIER_C["B"], TIER_C["C"]]
seg = [(f"A 档 × {tier_counts[tier_keys[0]]} 场", round(tier_counts[tier_keys[0]]*D.TIER_TOTAL[tier_keys[0]],1)),
       (f"B 档 × {tier_counts[tier_keys[1]]} 场", round(tier_counts[tier_keys[1]]*D.TIER_TOTAL[tier_keys[1]],1)),
       (f"C 档 × {tier_counts[tier_keys[2]]} 场", round(tier_counts[tier_keys[2]]*D.TIER_TOTAL[tier_keys[2]],1))]
for i, (lab, val) in enumerate(seg):
    x = Inches(0.7) + i*Inches(3.15); yy = y+Inches(0.62)
    rect(s, x, yy, Inches(2.9), Inches(0.85), fill=WHITE, line=tcolors[i], line_w=1.2)
    text(s, x, yy+Inches(0.08), Inches(2.9), Inches(0.35), [[(lab, 11, True, tcolors[i])]], align=PP_ALIGN.CENTER)
    text(s, x, yy+Inches(0.42), Inches(2.9), Inches(0.4), [[(f"{val} 万元", 13, True, PLUM)]], align=PP_ALIGN.CENTER)
x = Inches(0.7) + 3*Inches(3.15); yy = y+Inches(0.62)
rect(s, x, yy, Inches(2.6), Inches(0.85), fill=PLUM)
text(s, x, yy+Inches(0.08), Inches(2.6), Inches(0.35), [[("12 场合计", 11, True, LILAC)]], align=PP_ALIGN.CENTER)
text(s, x, yy+Inches(0.42), Inches(2.6), Inches(0.4), [[(f"{D.TOTAL_PRICE} 万元", 14, True, WHITE)]], align=PP_ALIGN.CENTER)
footer(s, pg); pg += 1

# ============================================================ 15. 招商漏斗
s = slide(); bg(s); header(s, "价值 · 核心", "招商引资价值分析（一）转化漏斗", "07")
text(s, Inches(0.72), Inches(1.3), Inches(12.3), Inches(0.4),
     [[("核心利好：12 场活动为东方枢纽 133 万方项目构建“引流—深挖—对接—签约”的可量化招商漏斗。", 12, True, PLUM)]])
fw_list = [Inches(9.0), Inches(7.4), Inches(5.8), Inches(4.2), Inches(3.0)]
fcolors = [PURPLE, RGBColor(0x53,0x39,0x82), RGBColor(0x47,0x30,0x72), RGBColor(0x3A,0x27,0x5C), GOLD]
y = Inches(1.95)
for i, ((stage, basis, val), fw) in enumerate(zip(D.FUNNEL, fw_list)):
    x = Inches(0.7) + (Inches(9.0)-fw)/2
    rect(s, x, y, fw, Inches(0.72), fill=fcolors[i], shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x, y, fw, Inches(0.72), [[(stage, 11, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(10.0), y, Inches(2.9), Inches(0.72),
         [[(val, 15, True, GOLD)], [(basis, 8.5, False, GREY)]], anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.9)
text(s, Inches(0.7), Inches(6.75), Inches(12), Inches(0.35),
     [[("注：转化率为测算假设，用于展示价值逻辑；实际以邀约质量与后续跟进为准。", 8.5, False, GREY)]])
footer(s, pg); pg += 1

# ============================================================ 16. 去化 + 支柱
s = slide(); bg(s); header(s, "价值 · 核心", "招商引资价值分析（二）去化测算与价值支柱", "07")
text(s, Inches(0.72), Inches(1.32), Inches(6), Inches(0.4), [[("133 万方 × 情景去化率测算", 13, True, PLUM)]])
scol = [PURPLE2, PURPLE, GOLD]
for i, (name, rate, area) in enumerate(D.GMV_SCENARIOS):
    x = Inches(0.6) + i*Inches(2.05); y = Inches(1.8)
    rect(s, x, y, Inches(1.9), Inches(1.75), fill=WHITE, line=scol[i], line_w=1.3, shadow=True)
    rect(s, x, y, Inches(1.9), Inches(0.42), fill=scol[i])
    text(s, x, y, Inches(1.9), Inches(0.42), [[(name+"情景", 11, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x, y+Inches(0.55), Inches(1.9), Inches(0.4), [[(rate, 20, True, scol[i])]], align=PP_ALIGN.CENTER)
    text(s, x, y+Inches(1.05), Inches(1.9), Inches(0.6), [[("去化", 9, False, GREY)],[(area, 11, True, PLUM)]], align=PP_ALIGN.CENTER)
text(s, Inches(6.95), Inches(1.32), Inches(6), Inches(0.4), [[("投入产出概览", 13, True, PLUM)]])
for i, (name, val) in enumerate(D.ROI_SUMMARY):
    y = Inches(1.8) + i*Inches(0.38)
    rect(s, Inches(6.95), y, Inches(5.85), Inches(0.32), fill=LAV, line=LAVED, line_w=0.5)
    rect(s, Inches(6.95), y, Pt(4), Inches(0.32), fill=GOLD)
    text(s, Inches(7.15), y, Inches(3.4), Inches(0.32), [[(name, 9.5, False, DARK)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(10.3), y, Inches(2.4), Inches(0.32), [[(val, 10.5, True, PURPLE)]], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
text(s, Inches(0.72), Inches(3.9), Inches(6), Inches(0.4), [[("六大价值支柱", 13, True, PLUM)]])
for i, (name, desc) in enumerate(D.VALUE_PILLARS):
    col = i % 3; row = i // 3
    x = Inches(0.6) + col*Inches(4.12); y = Inches(4.35) + row*Inches(1.35)
    w = Inches(3.9); h = Inches(1.2)
    rect(s, x, y, w, h, fill=WHITE, line=LAVED, line_w=1, shadow=True)
    text(s, x+Inches(0.15), y+Inches(0.08), w-Inches(0.3), Inches(0.32), [[(f"{i+1}. {name}", 11, True, PURPLE)]])
    text(s, x+Inches(0.15), y+Inches(0.42), w-Inches(0.3), Inches(0.75), [[(desc, 8.6, False, DARK)]], line_spacing=1.0)
footer(s, pg); pg += 1

# ============================================================ 17. 执行机制
s = slide(); bg(s); header(s, "执行 · EXECUTION", "执行机制与合规保障", "08")
for i, (name, desc) in enumerate(D.EXECUTION):
    col = i % 2; row = i // 2
    x = Inches(0.6) + col*Inches(6.3); y = Inches(1.5) + row*Inches(1.75)
    w = Inches(6.0); h = Inches(1.55)
    rect(s, x, y, w, h, fill=WHITE, line=LAVED, line_w=0.75, shadow=True)
    rect(s, x, y, Inches(1.0), h, fill=PURPLE)
    text(s, x, y, Inches(1.0), h, [[(f"{i+1:02d}", 22, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+Inches(1.15), y+Inches(0.15), w-Inches(1.3), Inches(0.4), [[(name, 13, True, PLUM)]])
    text(s, x+Inches(1.15), y+Inches(0.58), w-Inches(1.3), Inches(0.9), [[(desc, 10, False, DARK)]], line_spacing=1.1)
footer(s, pg); pg += 1

# ============================================================ 18. 下一步
s = slide()
rect(s, 0, 0, SW, SH, fill=PLUM)
rect(s, 0, Inches(2.3), SW, Pt(4), fill=GOLD)
text(s, Inches(0.72), Inches(0.9), Inches(12), Inches(0.6), [[("下一步 · NEXT STEPS", 15, True, LILAC)]])
text(s, Inches(0.72), Inches(1.4), Inches(12), Inches(0.8), [[("确认试点场次，即刻启动", 34, True, WHITE)]])
steps = [
    ("① 确认首场", "敲定 7 月中旬“企业出海·进境关外”试点主题、时间与 30 人以内精准邀约名单"),
    ("② 提交 OA 要件", "邀约名单 / 参会人数 / 人员背景 / 预算 / 时间，供东方枢纽走内部 OA 流程"),
    ("③ 供应商签约", "经指定策划供应商签约支付，落实国货区场地搭建通行证"),
    ("④ 滚动推进", "两周一场滚动执行，形成“活动—跟进—评估—再合作”长期闭环"),
]
for i, (t, d) in enumerate(steps):
    y = Inches(2.75) + i*Inches(1.02)
    rect(s, Inches(0.72), y, Inches(11.9), Inches(0.88), fill=PURPLE)
    rect(s, Inches(0.72), y, Pt(5), Inches(0.88), fill=GOLD)
    text(s, Inches(0.95), y, Inches(3.0), Inches(0.88), [[(t, 15, True, GOLD)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(3.9), y, Inches(8.5), Inches(0.88), [[(d, 11.5, False, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)
text(s, Inches(0.72), Inches(6.95), Inches(12), Inches(0.4),
     [[("复旦大学  ·  上海市科技企业联合会  ·  东方枢纽    |    合作愉快", 11, True, LILAC)]])

prs.save("东方枢纽三方合作计划_汇报PPT.pptx")
print("PPT 已生成：东方枢纽三方合作计划_汇报PPT.pptx  (页数: %d)" % len(prs.slides._sldIdLst))
