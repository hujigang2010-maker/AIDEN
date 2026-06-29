# -*- coding: utf-8 -*-
"""生成《源信网络 × 火山引擎/腾讯云 算力补贴三方合作方案》PPT。"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import copy

# ---------- 配色（科技蓝） ----------
NAVY      = RGBColor(0x0B, 0x2A, 0x5B)   # 深蓝
BLUE      = RGBColor(0x12, 0x5C, 0xC4)   # 主蓝
BLUE2     = RGBColor(0x2E, 0x8B, 0xE6)   # 亮蓝
CYAN      = RGBColor(0x00, 0xB4, 0xD8)   # 青
LIGHT     = RGBColor(0xEA, 0xF2, 0xFB)   # 浅蓝底
LIGHT2    = RGBColor(0xDB, 0xE9, 0xFA)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
INK       = RGBColor(0x1B, 0x24, 0x33)   # 正文深色
GRAY      = RGBColor(0x5A, 0x64, 0x72)
GOLD      = RGBColor(0xF2, 0xA0, 0x2D)   # 强调橙

FONT      = "Microsoft YaHei"   # 渲染时回退到文泉驿微米黑
FONT_FB   = "WenQuanYi Micro Hei"

EMU_W = Inches(13.333)
EMU_H = Inches(7.5)

prs = Presentation()
prs.slide_width = EMU_W
prs.slide_height = EMU_H
BLANK = prs.slide_layouts[6]


# ---------- 工具函数 ----------
def _set_font(run, size, color, bold=False, font=FONT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    # 东亚字体
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', font)


def add_slide():
    return prs.slides.add_slide(BLANK)


def rect(slide, x, y, w, h, fill, line=None, shadow=False, shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, x, y, w, h)
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(1)
    sp.shadow.inherit = False
    return sp


def grad(slide, x, y, w, h, c1, c2, angle=45):
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sp.line.fill.background()
    sp.shadow.inherit = False
    sp.fill.gradient()
    stops = sp.fill.gradient_stops
    stops[0].position = 0.0
    stops[0].color.rgb = c1
    stops[1].position = 1.0
    stops[1].color.rgb = c2
    try:
        sp.fill.gradient_angle = angle
    except Exception:
        pass
    return sp


def text(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         line_spacing=1.0, space_after=4, wrap=True):
    """runs: list of paragraphs; each paragraph is list of (txt,size,color,bold,font)."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = Pt(space_after)
        p.space_before = Pt(0)
        for seg in para:
            txt, size, color, bold = seg[0], seg[1], seg[2], seg[3]
            font = seg[4] if len(seg) > 4 else FONT
            r = p.add_run()
            r.text = txt
            _set_font(r, size, color, bold, font)
    return tb


def bg(slide, color=WHITE):
    rect(slide, 0, 0, EMU_W, EMU_H, color)


def header(slide, kicker, title, idx=None):
    """统一内容页页眉。"""
    bg(slide, WHITE)
    # 左侧竖色条
    rect(slide, 0, 0, Inches(0.22), EMU_H, BLUE)
    rect(slide, Inches(0.22), 0, Inches(0.08), EMU_H, CYAN)
    # 顶部标题块
    text(slide, Inches(0.7), Inches(0.45), Inches(11.5), Inches(0.4),
         [[(kicker, 13, BLUE2, True)]])
    text(slide, Inches(0.7), Inches(0.78), Inches(11.9), Inches(0.7),
         [[(title, 27, NAVY, True)]])
    rect(slide, Inches(0.72), Inches(1.5), Inches(1.5), Inches(0.06), GOLD)
    if idx is not None:
        text(slide, Inches(11.9), Inches(0.5), Inches(1.0), Inches(0.5),
             [[(idx, 12, GRAY, False)]], align=PP_ALIGN.RIGHT)
    # 页脚
    text(slide, Inches(0.7), Inches(7.02), Inches(9), Inches(0.3),
         [[("源信网络 × 复旦大学住房政策研究中心 × 杨浦区科技企业联合会  |  算力补贴合作方案", 9, GRAY, False)]])


def chip(slide, x, y, w, h, label, fill, tcolor=WHITE, size=12, bold=True, rounded=True):
    sp = rect(slide, x, y, w, h, fill,
              shape=MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE)
    tf = sp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = label
    _set_font(r, size, tcolor, bold)
    return sp


# ============================================================
# 1. 封面
# ============================================================
s = add_slide()
grad(s, 0, 0, EMU_W, EMU_H, NAVY, BLUE, angle=60)
# 右侧装饰圆
c = rect(s, Inches(9.3), Inches(-1.6), Inches(5.6), Inches(5.6), BLUE2, shape=MSO_SHAPE.OVAL)
c.fill.fore_color.rgb = RGBColor(0x1A, 0x4F, 0xA8)
c2 = rect(s, Inches(10.8), Inches(3.4), Inches(4.2), Inches(4.2), CYAN, shape=MSO_SHAPE.OVAL)
c2.fill.fore_color.rgb = RGBColor(0x12, 0x6A, 0xC0)
# 顶部标识
text(s, Inches(0.9), Inches(0.7), Inches(11), Inches(0.4),
     [[("AI 算力普惠 · 楼宇园区赋能计划", 14, CYAN, True)]])
text(s, Inches(0.9), Inches(2.35), Inches(11.5), Inches(2.0),
     [[("算力补贴三方合作方案", 46, WHITE, True)]])
text(s, Inches(0.9), Inches(3.5), Inches(11.5), Inches(0.9),
     [[("以大厂（火山引擎 / 腾讯云）token 算力补贴，", 20, LIGHT, False)],
      [("赋能办公楼、产业园区与科技载体内的 B 端企业", 20, LIGHT, False)]],
     line_spacing=1.2)
# 三方标签
chip(s, Inches(0.9), Inches(5.3), Inches(3.4), Inches(0.55), "复旦大学住房政策研究中心", WHITE, NAVY, 12)
chip(s, Inches(4.5), Inches(5.3), Inches(3.4), Inches(0.55), "杨浦区科技企业联合会", WHITE, NAVY, 12)
chip(s, Inches(8.1), Inches(5.3), Inches(2.0), Inches(0.55), "源信网络", GOLD, WHITE, 13)
text(s, Inches(0.9), Inches(6.5), Inches(8), Inches(0.4),
     [[("汇报交流稿  ·  2026 年 6 月", 13, LIGHT, False)]])

# ============================================================
# 2. 目录
# ============================================================
s = add_slide()
header(s, "AGENDA", "目录")
items = [
    ("01", "合作背景与机遇", "AI 算力时代与杨浦科创生态"),
    ("02", "三方角色与定位", "复旦研究中心 · 科企联合会 · 源信网络"),
    ("03", "合作总体框架", "大厂算力补贴的传导与分发链路"),
    ("04", "准入与申请条件", "高端办公门槛 + 5 万元无门槛消费券"),
    ("05", "补贴与优惠方式", "租金减免 · 算力补贴 · 运营产品补贴"),
    ("06", "补贴标准表", "按企业规模分档的补贴一览"),
    ("07", "三方共赢价值", "各方收益与协同效应"),
    ("08", "开放式可复制模式", "标准化模板 · 开放接入 · 可复制"),
    ("09", "落地路径与下一步", "试点—推广—复制的实施计划"),
]
y = 1.72
for i, (n, t, d) in enumerate(items):
    col = i % 2
    row = i // 2
    x = Inches(0.8 + col * 6.05)
    yy = Inches(y + row * 1.12)
    rect(s, x, yy, Inches(5.75), Inches(0.98), LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, yy, Inches(0.9), Inches(0.98), BLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x, yy + Inches(0.18), Inches(0.9), Inches(0.6),
         [[(n, 22, WHITE, True)]], align=PP_ALIGN.CENTER)
    text(s, x + Inches(1.1), yy + Inches(0.14), Inches(4.5), Inches(0.4),
         [[(t, 16, NAVY, True)]])
    text(s, x + Inches(1.1), yy + Inches(0.55), Inches(4.6), Inches(0.35),
         [[(d, 11, GRAY, False)]])

# ============================================================
# 3. 合作背景与机遇
# ============================================================
s = add_slide()
header(s, "01  合作背景", "AI 算力时代的普惠机遇", "01 / 09")
paras = [
    ("AI 算力成为新基建", "大模型与 AI Agent 应用爆发，token 算力成为企业研发与经营的刚需，但中小企业普遍面临用云成本高、门槛高的痛点。"),
    ("大厂算力红利可被引导", "火山引擎、腾讯云等持续投入算力补贴与生态扶持，源信网络作为合作方可对接补贴资源，向楼宇 / 园区的 B 端企业精准分发。"),
    ("杨浦科创载体丰富", "杨浦区甲级写字楼、产业园区、科技载体集聚，企业密度高，是算力普惠政策落地的优质试验田。"),
    ("产学研政协同", "复旦大学住房政策研究中心提供政策研究与评估，科技企业联合会链接企业资源，源信网络负责运营落地，形成闭环。"),
]
y = 1.9
for i, (t, d) in enumerate(paras):
    yy = Inches(y + i * 1.18)
    rect(s, Inches(0.8), yy, Inches(11.7), Inches(1.0), LIGHT if i % 2 == 0 else LIGHT2,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, Inches(0.8), yy, Inches(0.14), Inches(1.0), CYAN)
    text(s, Inches(1.15), yy + Inches(0.12), Inches(3.3), Inches(0.8),
         [[(t, 16, BLUE, True)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(4.6), yy + Inches(0.1), Inches(7.7), Inches(0.85),
         [[(d, 12.5, INK, False)]], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)

# ============================================================
# 4. 三方角色与定位
# ============================================================
s = add_slide()
header(s, "02  合作主体", "三方角色与定位", "02 / 09")
cards = [
    ("复旦大学住房政策研究中心", "政策智库 · 研究背书", BLUE,
     ["政策设计与合规研究", "补贴成效评估与课题", "产学研成果转化与背书"]),
    ("杨浦区科技企业联合会", "企业资源 · 组织协调", CYAN,
     ["对接区内企业与载体", "组织申报与政策宣贯", "汇集企业算力需求"]),
    ("源信网络", "运营落地 · 资源对接", GOLD,
     ["对接大厂算力补贴资源", "统筹楼宇 / 园区合作", "补贴发放与运营服务"]),
]
cw = Inches(3.85)
gap = Inches(0.18)
x0 = Inches(0.8)
for i, (name, role, color, pts) in enumerate(cards):
    x = x0 + i * (cw + gap)
    rect(s, x, Inches(1.95), cw, Inches(4.4), LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, Inches(1.95), cw, Inches(1.15), color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, Inches(2.7), cw, Inches(0.4), color)
    text(s, x + Inches(0.2), Inches(2.05), cw - Inches(0.4), Inches(0.7),
         [[(name, 16.5, WHITE, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + Inches(0.2), Inches(2.72), cw - Inches(0.4), Inches(0.36),
         [[(role, 12, WHITE, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    yy = 3.35
    for p in pts:
        text(s, x + Inches(0.35), Inches(yy), cw - Inches(0.6), Inches(0.7),
             [[("● ", 12, color, True), (p, 13, INK, False)]], line_spacing=1.05)
        yy += 0.92
# 底部链路
text(s, Inches(0.8), Inches(6.55), Inches(11.7), Inches(0.4),
     [[("协同闭环：研究背书 → 企业组织 → 运营落地，共同放大大厂算力补贴的政策价值", 13, NAVY, True)]],
     align=PP_ALIGN.CENTER)

# ============================================================
# 5. 合作总体框架（传导链路）
# ============================================================
s = add_slide()
header(s, "03  合作框架", "算力补贴的传导与分发链路", "03 / 09")
flow = [
    ("大厂", "火山引擎 / 腾讯云\n提供 token 算力补贴", BLUE),
    ("源信网络", "对接补贴 + 统筹运营\n方案设计与发放", GOLD),
    ("物业 / 楼宇 / 园区", "科技载体作为渠道\n承接并落地政策", CYAN),
    ("B 端企业", "园区楼宇内企业\n享受算力与租金福利", NAVY),
]
bw = Inches(2.75)
y = Inches(2.35)
x0 = 0.85
for i, (t, d, color) in enumerate(flow):
    x = Inches(x0 + i * 3.05)
    rect(s, x, y, bw, Inches(1.8), color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x, y + Inches(0.22), bw, Inches(0.5), [[(t, 18, WHITE, True)]], align=PP_ALIGN.CENTER)
    text(s, x + Inches(0.15), y + Inches(0.78), bw - Inches(0.3), Inches(0.95),
         [[(line, 11.5, WHITE, False)] for line in d.split("\n")],
         align=PP_ALIGN.CENTER, line_spacing=1.05)
    if i < 3:
        ar = s.shapes.add_shape(MSO_SHAPE.CHEVRON, x + bw + Inches(0.02), y + Inches(0.62),
                                Inches(0.26), Inches(0.55))
        ar.fill.solid(); ar.fill.fore_color.rgb = GOLD
        ar.line.fill.background(); ar.shadow.inherit = False
# 下方说明三块
notes = [
    ("补贴资源", "由源信网络对接大厂算力补贴与生态资源池，统一额度池管理。", BLUE),
    ("分发渠道", "以楼宇 / 园区为载体白名单，按准入条件向企业发放。", CYAN),
    ("政策福利", "企业端获得算力券、租金减免、云产品与运营补贴。", GOLD),
]
nw = Inches(3.8)
for i, (t, d, color) in enumerate(notes):
    x = Inches(0.85 + i * 3.95)
    rect(s, x, Inches(4.7), nw, Inches(1.55), LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, Inches(4.7), nw, Inches(0.5), color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x, Inches(4.76), nw, Inches(0.4), [[(t, 14, WHITE, True)]], align=PP_ALIGN.CENTER)
    text(s, x + Inches(0.25), Inches(5.3), nw - Inches(0.5), Inches(0.9),
         [[(d, 12, INK, False)]], line_spacing=1.1)

# ============================================================
# 6. 准入与申请条件
# ============================================================
s = add_slide()
header(s, "04  准入与申请", "准入门槛与申请条件", "04 / 09")
# 左：高端办公门槛
rect(s, Inches(0.8), Inches(1.9), Inches(7.0), Inches(4.55), LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
rect(s, Inches(0.8), Inches(1.9), Inches(7.0), Inches(0.7), BLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, Inches(0.8), Inches(1.97), Inches(7.0), Inches(0.55),
     [[("（一）高端办公场所 · 申请门槛（满足即可申请）", 15, WHITE, True)]], align=PP_ALIGN.CENTER)
conds = [
    ("入驻载体", "甲级写字楼 / 重点产业园 / 科技载体（园区·楼宇白名单）"),
    ("企业属性", "注册及纳税在杨浦区，科技型、AI / 数字经济相关企业优先"),
    ("租赁规模", "租赁面积 ≥ 200㎡ 或 ≥ 10 个工位"),
    ("合作期限", "签约期 ≥ 12 个月，并承诺使用大厂云算力服务"),
    ("配合事项", "配合补贴核销、案例展示与政策研究数据采集"),
]
yy = 2.78
for t, d in conds:
    text(s, Inches(1.1), Inches(yy), Inches(1.7), Inches(0.6),
         [[(t, 13, BLUE, True)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(2.85), Inches(yy), Inches(4.75), Inches(0.65),
         [[(d, 12, INK, False)]], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05)
    yy += 0.72
# 右：5万元消费券
rect(s, Inches(8.05), Inches(1.9), Inches(4.45), Inches(4.55), NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, Inches(8.25), Inches(2.2), Inches(4.05), Inches(0.5),
     [[("（二）无门槛福利", 15, CYAN, True)]], align=PP_ALIGN.CENTER)
text(s, Inches(8.25), Inches(2.85), Inches(4.05), Inches(1.1),
     [[("¥5万元", 44, WHITE, True)]], align=PP_ALIGN.CENTER)
text(s, Inches(8.25), Inches(3.95), Inches(4.05), Inches(0.5),
     [[("免费 · 无门槛消费券", 16, GOLD, True)]], align=PP_ALIGN.CENTER)
for i, d in enumerate(["园区 / 楼宇内企业均可领取", "用于抵扣大厂云产品与算力消费", "作为引流入口，先用起来再升级"]):
    text(s, Inches(8.45), Inches(4.55 + i * 0.5), Inches(3.9), Inches(0.45),
         [[("✓ ", 13, CYAN, True), (d, 12, LIGHT, False)]])

# ============================================================
# 7. 补贴与优惠方式（三支柱）
# ============================================================
s = add_slide()
header(s, "05  补贴方式", "补贴与优惠方式 · 三大支柱", "05 / 09")
pillars = [
    ("租金减免", "免几个月 · 补几个月", BLUE,
     ["按合作规模采取“免 X 月 + 补 X 月”折扣", "由载体让利 + 补贴池共担", "签约越长、规模越大，减免越多"]),
    ("算力补贴", "大厂算力直接补贴", CYAN,
     ["发放 token 算力券 / 代金额度", "对接火山引擎 / 腾讯云补贴池", "按企业规模分档授信额度"]),
    ("运营与产品消费补贴", "云服务 + 产品 + 运营", GOLD,
     ["云产品消费折扣与返券", "大厂产品采购专项补贴", "运营 / 培训 / 上云服务补贴"]),
]
cw = Inches(3.85)
for i, (t, sub, color, pts) in enumerate(pillars):
    x = Inches(0.8 + i * 4.03)
    rect(s, x, Inches(1.95), cw, Inches(4.5), LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, Inches(1.95), cw, Inches(1.2), color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, Inches(2.75), cw, Inches(0.4), color)
    text(s, x + Inches(0.15), Inches(2.05), cw - Inches(0.3), Inches(0.7),
         [[(t, 15.5, WHITE, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + Inches(0.15), Inches(2.78), cw - Inches(0.3), Inches(0.36),
         [[(sub, 11.5, WHITE, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    yy = 3.45
    for p in pts:
        rect(s, x + Inches(0.3), Inches(yy + 0.08), Inches(0.12), Inches(0.12), color, shape=MSO_SHAPE.OVAL)
        text(s, x + Inches(0.55), Inches(yy), cw - Inches(0.8), Inches(0.85),
             [[(p, 12.5, INK, False)]], line_spacing=1.05)
        yy += 0.95

# ============================================================
# 8. 补贴标准表（按企业规模分档）
# ============================================================
s = add_slide()
header(s, "06  补贴标准", "补贴标准一览（按企业规模分档）", "06 / 09")
rows = [
    ["企业规模（人）", "算力补贴额度\n(元/年)", "免费消费券", "租金减免\n方案", "云/运营\n消费补贴", "大额签约\n折扣"],
    ["0–50（含）", "5 万 token券", "5 万元", "免1补1", "8.5 折", "9.5 折"],
    ["50–100（含）", "15 万", "5 万元", "免2补1", "8.0 折", "9.0 折"],
    ["100–300（含）", "30 万", "5 万元", "免2补2", "7.5 折", "8.5 折"],
    ["300–500（含）", "60 万", "5 万元", "免3补2", "7.0 折", "8.5 折"],
    ["> 500", "100 万 + 定制", "5 万元", "免3补3", "6.5 折", "8.0 折"],
]
nrows, ncols = len(rows), len(rows[0])
tbl_x, tbl_y = Inches(0.8), Inches(2.0)
tbl_w, tbl_h = Inches(11.7), Inches(4.05)
gtbl = s.shapes.add_table(nrows, ncols, tbl_x, tbl_y, tbl_w, tbl_h).table
gtbl.first_row = False
gtbl.horz_banding = False
widths = [2.5, 2.0, 1.85, 1.6, 1.9, 1.85]
for c, w in enumerate(widths):
    gtbl.columns[c].width = Inches(w)
for r in range(nrows):
    gtbl.rows[r].height = Inches(0.58) if r == 0 else Inches(0.69)
for r in range(nrows):
    for c in range(ncols):
        cell = gtbl.cell(r, c)
        cell.margin_left = Inches(0.05); cell.margin_right = Inches(0.05)
        cell.margin_top = Inches(0.02); cell.margin_bottom = Inches(0.02)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        if r == 0:
            cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
        else:
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if r % 2 == 1 else LIGHT
        tf = cell.text_frame; tf.word_wrap = True
        lines = rows[r][c].split("\n")
        for li, ln in enumerate(lines):
            p = tf.paragraphs[0] if li == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            p.line_spacing = 1.0
            run = p.add_run(); run.text = ln
            if r == 0:
                _set_font(run, 12.5, WHITE, True)
            elif c == 0:
                _set_font(run, 12.5, NAVY, True)
            else:
                _set_font(run, 12.5, INK, False)
text(s, Inches(0.8), Inches(6.25), Inches(11.7), Inches(0.7),
     [[("说明：", 11, GOLD, True),
       ("以上为建议方案，最终额度以大厂补贴政策及三方协议为准；“免X补X”指首年租金由载体减免与补贴池共担。", 11, GRAY, False)]],
     line_spacing=1.1)

# ============================================================
# 9. 三方共赢价值
# ============================================================
s = add_slide()
header(s, "07  共赢价值", "三方共赢与协同价值", "07 / 09")
vals = [
    ("大厂（火山/腾讯）", BLUE, ["精准触达 B 端企业客户", "补贴转化为云消费与留存", "打造区域算力普惠样板"]),
    ("楼宇 / 园区物业", CYAN, ["差异化招商卖点", "提升入驻率与续约率", "升级为“算力友好型”载体"]),
    ("入驻 B 端企业", GOLD, ["降低用云与租金成本", "获得算力与上云支持", "加速 AI 应用落地"]),
    ("源信网络 + 研究中心 + 联合会", NAVY, ["运营服务与生态收益", "政策研究与成果背书", "区域科创影响力提升"]),
]
cw = Inches(5.75); ch = Inches(2.05)
for i, (t, color, pts) in enumerate(vals):
    col, row = i % 2, i // 2
    x = Inches(0.8 + col * 6.05)
    y = Inches(1.95 + row * 2.25)
    rect(s, x, y, cw, ch, LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, y, Inches(0.16), ch, color)
    text(s, x + Inches(0.4), y + Inches(0.18), cw - Inches(0.6), Inches(0.5),
         [[(t, 16, color, True)]])
    yy = 0.72
    for p in pts:
        text(s, x + Inches(0.45), y + Inches(yy), cw - Inches(0.7), Inches(0.42),
             [[("● ", 11, color, True), (p, 12.5, INK, False)]])
        yy += 0.43

# ============================================================
# 9b. 开放式可复制合作模式
# ============================================================
s = add_slide()
header(s, "08  可复制模式", "开放式 · 可复制合作模式", "08 / 09")
text(s, Inches(0.8), Inches(1.72), Inches(11.7), Inches(0.5),
     [[("把三方合作沉淀为“一次设计、处处可用”的开放式算力普惠样板——任何楼宇 / 园区 / 区域均可低成本接入与复制。",
        13.5, GRAY, False)]], line_spacing=1.1)
feats = [
    ("标准化模板", "合作协议、准入条件、补贴标准、核销流程全部模板化，开箱即用。", BLUE),
    ("模块化组合", "算力补贴 / 租金减免 / 消费补贴可按载体需求自由拼装组合。", CYAN),
    ("开放式接入", "对大厂、物业、园区、企业开放，白名单动态扩容、平台化运营。", GOLD),
    ("可复制推广", "一套打法复制到多楼宇、多园区乃至跨区域，边际成本递减。", NAVY),
]
cw = Inches(5.75); ch = Inches(1.95)
for i, (t, d, color) in enumerate(feats):
    col, row = i % 2, i // 2
    x = Inches(0.8 + col * 6.05)
    y = Inches(2.45 + row * 2.12)
    rect(s, x, y, cw, ch, LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, y, Inches(1.55), ch, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x + Inches(1.2), y, Inches(0.4), ch, color)
    text(s, x + Inches(0.1), y, Inches(1.35), ch,
         [[(str(i + 1).zfill(2), 30, WHITE, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + Inches(1.8), y + Inches(0.28), cw - Inches(2.0), Inches(0.5),
         [[(t, 17, color, True)]])
    text(s, x + Inches(1.8), y + Inches(0.85), cw - Inches(2.05), Inches(1.0),
         [[(d, 12.5, INK, False)]], line_spacing=1.15)
text(s, Inches(0.8), Inches(6.72), Inches(11.7), Inches(0.35),
     [[("开放接入 · 标准复制 · 平台运营 —— 形成可持续扩张的算力普惠生态", 13, NAVY, True)]],
     align=PP_ALIGN.CENTER)

# ============================================================
# 10. 落地路径与下一步
# ============================================================
s = add_slide()
header(s, "09  落地路径", "落地路径与下一步", "09 / 09")
phases = [
    ("第一阶段\n试点", "选取 1–2 个标杆楼宇 / 园区，确定补贴池与白名单，落地首批企业。", BLUE),
    ("第二阶段\n推广", "总结试点成效，由联合会组织区内载体规模化申报与宣贯。", CYAN),
    ("第三阶段\n复制", "形成标准化合作模板与政策研究报告，向全区乃至更大范围复制。", GOLD),
]
bw = Inches(3.8)
for i, (t, d, color) in enumerate(phases):
    x = Inches(0.85 + i * 3.95)
    rect(s, x, Inches(2.0), bw, Inches(2.5), LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, Inches(2.0), bw, Inches(0.95), color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x, Inches(2.1), bw, Inches(0.8),
         [[(line, 15, WHITE, True)] for line in t.split("\n")], align=PP_ALIGN.CENTER)
    text(s, x + Inches(0.3), Inches(3.15), bw - Inches(0.6), Inches(1.2),
         [[(d, 12.5, INK, False)]], line_spacing=1.15)
    if i < 2:
        ar = s.shapes.add_shape(MSO_SHAPE.CHEVRON, x + bw - Inches(0.02), Inches(2.95),
                                Inches(0.22), Inches(0.6))
        ar.fill.solid(); ar.fill.fore_color.rgb = GOLD
        ar.line.fill.background(); ar.shadow.inherit = False
# 下一步行动条
rect(s, Inches(0.85), Inches(4.95), Inches(11.65), Inches(1.55), NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, Inches(1.2), Inches(5.1), Inches(11), Inches(0.5),
     [[("下一步行动", 16, CYAN, True)]])
for i, d in enumerate(["① 三方明确分工与补贴池规模", "② 对接大厂确认算力补贴政策", "③ 确定首批试点载体与企业名单"]):
    text(s, Inches(1.2 + i * 3.8), Inches(5.7), Inches(3.7), Inches(0.7),
         [[(d, 12.5, WHITE, False)]], line_spacing=1.05)

# ============================================================
# 11. 封底
# ============================================================
s = add_slide()
grad(s, 0, 0, EMU_W, EMU_H, NAVY, BLUE, angle=60)
c = rect(s, Inches(-1.5), Inches(3.5), Inches(5), Inches(5), CYAN, shape=MSO_SHAPE.OVAL)
c.fill.fore_color.rgb = RGBColor(0x12, 0x6A, 0xC0)
text(s, Inches(0.9), Inches(2.7), Inches(11.5), Inches(1.2),
     [[("让算力普惠，赋能每一栋楼宇与园区", 34, WHITE, True)]])
text(s, Inches(0.9), Inches(4.0), Inches(11.5), Inches(0.6),
     [[("期待与各方深度合作，共建区域 AI 算力普惠样板", 18, LIGHT, False)]])
chip(s, Inches(0.9), Inches(5.3), Inches(3.4), Inches(0.55), "复旦大学住房政策研究中心", WHITE, NAVY, 12)
chip(s, Inches(4.5), Inches(5.3), Inches(3.4), Inches(0.55), "杨浦区科技企业联合会", WHITE, NAVY, 12)
chip(s, Inches(8.1), Inches(5.3), Inches(2.0), Inches(0.55), "源信网络", GOLD, WHITE, 13)

import os
out = "源信网络算力补贴三方合作方案.pptx"
prs.save(out)
print("saved", out, os.path.getsize(out))
