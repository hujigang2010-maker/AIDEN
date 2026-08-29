#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成《AI × 不动产资产管理高级研修班》对外展示方案 PPT。

定位：面向合作方、渠道与意向学员的 16:9 汇报稿。
主干：AI+不动产培训；澳洲考察仅作圈层延展，不作为主线。

运行：python3 scripts/build_ai_real_estate_deck.py
输出：output/AI+不动产资产管理高级研修班_合作方案.pptx
"""
from __future__ import annotations

import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# ---------------------------------------------------------------------------
# 设计规范：不动产稳健蓝 + AI 青绿 + 圈层金
# ---------------------------------------------------------------------------
NAVY = RGBColor(0x0B, 0x1F, 0x3A)
NAVY2 = RGBColor(0x14, 0x32, 0x56)
NAVY3 = RGBColor(0x1C, 0x42, 0x6E)
TEAL = RGBColor(0x0E, 0x8A, 0x7D)
TEAL_DEEP = RGBColor(0x0A, 0x6B, 0x61)
GOLD = RGBColor(0xC4, 0xA3, 0x5A)
GOLD_PALE = RGBColor(0xF4, 0xEB, 0xD3)
CREAM = RGBColor(0xF6, 0xF3, 0xEC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
INK = RGBColor(0x1E, 0x2A, 0x3A)
GRAY = RGBColor(0x5C, 0x6B, 0x7A)
MUTED = RGBColor(0x8A, 0x96, 0xA3)
CARD = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xE4, 0xDC, 0xCC)
SOFT = RGBColor(0xEE, 0xF6, 0xF4)

FONT = "WenQuanYi Micro Hei"
FOOTER = "AI × 不动产资产管理高级研修班 · 合作方案 · 2026 秋季首期"

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(ROOT, "output")
OUT = os.path.join(OUT_DIR, "AI+不动产资产管理高级研修班_合作方案.pptx")

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def slide():
    return prs.slides.add_slide(BLANK)


def _set_font(run, size, color, bold=False, italic=False, name=FONT):
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", name)


def rect(s, x, y, w, h, fill, line=None, line_w=None, round_=False, radius=0.06):
    shp = s.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if round_ else MSO_SHAPE.RECTANGLE, x, y, w, h
    )
    if round_:
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_w or 1)
    shp.shadow.inherit = False
    return shp


def txt(s, x, y, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, wrap=True, spacing=1.08):
    """lines: tuple 或 list of tuples (text, size, color, bold[, space_after])。"""
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    first = True
    for ln in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = spacing
        segs = ln if isinstance(ln, list) else [ln]
        sa = None
        for seg in segs:
            text, size, color = seg[0], seg[1], seg[2]
            bold = seg[3] if len(seg) > 3 else False
            if len(seg) > 4 and seg[4] is not None:
                sa = seg[4]
            r = p.add_run()
            r.text = text
            _set_font(r, size, color, bold)
        if sa is not None:
            p.space_after = Pt(sa)
        p.space_before = Pt(0)
    return tb


def header(s, idx, kicker, title, subtitle=None):
    rect(s, 0, 0, SW, SH, CREAM)
    rect(s, 0, 0, Inches(0.14), SH, TEAL)
    rect(s, 0, 0, SW, Inches(0.08), NAVY)
    txt(
        s,
        Inches(0.55),
        Inches(0.28),
        Inches(10.4),
        Inches(0.28),
        [(FOOTER, 10, MUTED, False)],
    )
    txt(
        s,
        Inches(0.55),
        Inches(0.56),
        Inches(11.6),
        Inches(0.62),
        [[(kicker + "  ", 13, TEAL, True), (title, 24, NAVY, True)]],
    )
    if subtitle:
        txt(
            s,
            Inches(0.55),
            Inches(1.18),
            Inches(11.8),
            Inches(0.34),
            [(subtitle, 12.5, GRAY, False)],
        )
        bar_y = Inches(1.56)
    else:
        bar_y = Inches(1.28)
    rect(s, Inches(0.55), bar_y, Inches(12.2), Pt(2.2), GOLD)
    txt(
        s,
        Inches(12.15),
        Inches(7.08),
        Inches(0.85),
        Inches(0.28),
        [(f"{idx:02d}", 12, MUTED, True)],
        align=PP_ALIGN.RIGHT,
    )


def card(s, x, y, w, h, accent=TEAL):
    rect(s, x, y, w, h, CARD, line=LINE, line_w=1, round_=True, radius=0.05)
    rect(s, x, y, Inches(0.10), h, accent)


# ===========================================================================
# 01 封面
# ===========================================================================
def cover():
    s = slide()
    rect(s, 0, 0, SW, SH, NAVY)
    rect(s, 0, 0, SW, Inches(0.12), TEAL)
    rect(s, Inches(9.85), 0, Inches(3.48), SH, NAVY2)
    rect(s, Inches(10.35), Inches(1.35), Inches(2.45), Inches(2.45), TEAL)
    rect(s, Inches(10.85), Inches(1.85), Inches(1.45), Inches(1.45), GOLD)

    txt(s, Inches(0.85), Inches(1.05), Inches(8.8), Inches(0.4),
        [("合 作 方 案  ·  2026 秋 季 首 期", 14, GOLD, True)])
    txt(s, Inches(0.82), Inches(1.7), Inches(9.2), Inches(2.15),
        [[("AI × 不动产", 44, WHITE, True, 6)],
         [("资产管理高级研修班", 36, WHITE, True)]])
    txt(s, Inches(0.85), Inches(4.05), Inches(8.8), Inches(0.7),
        [("不动产做主干，AI 做素养，圈层做长期价值。", 18, RGBColor(0xC5, 0xD4, 0xE4), False)])
    txt(s, Inches(0.85), Inches(4.7), Inches(8.8), Inches(0.45),
        [("面向上海及长三角不动产从业者的高性价比研修与资源连接方案", 14, MUTED, False)])

    chips = [
        ("周期", "6 个月 · 每月 1 天"),
        ("班型", "首期 25–60 人"),
        ("定价", "2.0–2.48 万元"),
        ("开班", "2026 年 11 月"),
    ]
    x = Inches(0.85)
    for k, v in chips:
        rect(s, x, Inches(5.55), Inches(2.05), Inches(1.12), NAVY2, round_=True, radius=0.08)
        txt(s, x + Inches(0.16), Inches(5.68), Inches(1.75), Inches(0.85),
            [[(k, 11, GOLD, True)], [(v, 13, WHITE, True)]])
        x += Inches(2.2)

    txt(s, Inches(10.2), Inches(4.55), Inches(2.8), Inches(2.3),
        [[("展示用途", 11, GOLD, True, 8)],
         [("合作方对齐", 14, WHITE, True, 4)],
         [("渠道说明", 14, WHITE, True, 4)],
         [("招生沟通", 14, WHITE, True)]],
        align=PP_ALIGN.CENTER)


# ===========================================================================
# 02 目录
# ===========================================================================
def toc():
    s = slide()
    header(s, 2, "目录", "今天对齐什么", "一条主线讲清楚：为什么做、做成什么样、怎么落地、何时启动。")
    items = [
        ("01", "机会判断", "地产进入资产管理时代，AI 正在改写岗位与工具。"),
        ("02", "项目定位", "不是纯 AI 课，而是不动产实操 + AI 素养 + 行业圈层。"),
        ("03", "产品设计", "六个月班型、课程地图、师资与社群机制。"),
        ("04", "落地模式", "两种主体、定价、渠道与收益对照。"),
        ("05", "推进节奏", "9 月启动招生，11 月开班；先把培训做成入口。"),
    ]
    y = Inches(1.85)
    for no, title, desc in items:
        rect(s, Inches(0.55), y, Inches(12.2), Inches(0.92), CARD, line=LINE, line_w=1, round_=True)
        rect(s, Inches(0.55), y, Inches(1.15), Inches(0.92), NAVY, round_=True)
        txt(s, Inches(0.55), y, Inches(1.15), Inches(0.92),
            [(no, 18, GOLD, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        txt(s, Inches(1.95), y + Inches(0.12), Inches(10.4), Inches(0.7),
            [[(title, 18, NAVY, True, 2)], [(desc, 13, GRAY, False)]])
        y += Inches(1.02)


# ===========================================================================
# 03 机会判断
# ===========================================================================
def opportunity():
    s = slide()
    header(s, 3, "01  机会判断", "为什么是现在，为什么是不动产",
           "行业在换赛道，人在换岗位。培训要接住这两股力，而不是再讲一遍通用 AI。")

    cols = [
        (NAVY, "行业侧", "从开发销售转向资产管理",
         ["存量时代，持有、运营、处置成为主业。",
          "不良资产、REITs、资本化退出成为刚需话题。",
          "纯开发逻辑的课已经很难卖，资产管理才有人买单。"]),
        (TEAL, "技术侧", "AI 改写工具，但不改行业本质",
         ["中年从业者有转型压力，但听不懂纯技术课。",
          "市场上 AI 课同质化严重，听半小时就触顶。",
          "真正能成交的，是「AI 用在不动产上能干什么」。"]),
        (RGBColor(0x8A, 0x6A, 0x2A), "需求侧", "买的是圈层，不是课件",
         ["EMBA 动辄几十万，周期长、门槛高。",
          "两万元左右换一个可走动的行业圈子，更贴近当下。",
          "上海及长三角商会会员、国企、银行是现成客源。"]),
    ]
    x = Inches(0.5)
    for accent, kicker, title, bullets in cols:
        rect(s, x, Inches(1.85), Inches(3.95), Inches(4.85), CARD, line=LINE, line_w=1, round_=True)
        rect(s, x, Inches(1.85), Inches(3.95), Inches(1.35), accent, round_=True)
        rect(s, x, Inches(2.75), Inches(3.95), Inches(0.45), accent)
        txt(s, x + Inches(0.22), Inches(1.98), Inches(3.5), Inches(1.1),
            [[(kicker, 12, GOLD_PALE, True, 2)], [(title, 16, WHITE, True)]])
        yy = Inches(3.4)
        for b in bullets:
            txt(s, x + Inches(0.28), yy, Inches(3.4), Inches(0.95),
                [[("●  ", 12, TEAL, True), (b, 13, INK, False)]])
            yy += Inches(1.0)
        x += Inches(4.15)


# ===========================================================================
# 04 项目定位
# ===========================================================================
def positioning():
    s = slide()
    header(s, 4, "02  项目定位", "一句话讲清楚这门课",
           "对外可以说 AI，对内必须守住不动产。否则既听不懂，也卖不掉。")

    rect(s, Inches(0.5), Inches(1.85), Inches(12.3), Inches(1.35), NAVY, round_=True)
    txt(s, Inches(0.75), Inches(2.0), Inches(11.8), Inches(1.05),
        [[("不动产资产管理是主干，AI 只做基础素养。", 20, WHITE, True, 4)],
         [("班集体、学员互访和长期圈层，才是复购与转介绍的真正产品。", 14, RGBColor(0xC5, 0xD4, 0xE4), False)]],
        anchor=MSO_ANCHOR.MIDDLE)

    left = [
        ("不是", "一场通用 AI 公开课", "不堆模型名词，不做成科技展讲解。"),
        ("不是", "几十万级商学院项目", "周期更短、价格更友好、行业更垂直。"),
        ("不是", "一次听完就散的沙龙", "六个月同班，才能把人变成可持续连接。"),
    ]
    right = [
        ("是", "不动产实操研修", "资产管理、不良资产、REITs，对接本行工作。"),
        ("是", "高性价比行业圈层", "用两万元左右进入可走动的不动产资源网。"),
        ("是", "后续业务的流量入口", "先把班做成品牌，再谈科创对接与 FA。"),
    ]
    y = Inches(3.4)
    for (a1, t1, d1), (a2, t2, d2) in zip(left, right):
        rect(s, Inches(0.5), y, Inches(6.0), Inches(1.12), CARD, line=LINE, line_w=1, round_=True)
        rect(s, Inches(0.5), y, Inches(0.9), Inches(1.12), RGBColor(0x8B, 0x3A, 0x3A), round_=True)
        txt(s, Inches(0.5), y, Inches(0.9), Inches(1.12),
            [(a1, 14, WHITE, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        txt(s, Inches(1.6), y + Inches(0.14), Inches(4.7), Inches(0.85),
            [[(t1, 15, NAVY, True, 2)], [(d1, 12, GRAY, False)]])

        rect(s, Inches(6.8), y, Inches(6.0), Inches(1.12), CARD, line=LINE, line_w=1, round_=True)
        rect(s, Inches(6.8), y, Inches(0.9), Inches(1.12), TEAL, round_=True)
        txt(s, Inches(6.8), y, Inches(0.9), Inches(1.12),
            [(a2, 14, WHITE, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        txt(s, Inches(7.9), y + Inches(0.14), Inches(4.7), Inches(0.85),
            [[(t2, 15, NAVY, True, 2)], [(d2, 12, GRAY, False)]])
        y += Inches(1.22)


# ===========================================================================
# 05 核心洞察
# ===========================================================================
def insight():
    s = slide()
    header(s, 5, "02  项目定位", "学员真正在买什么",
           "招生话术可以讲 AI，成交逻辑要讲圈层与转型。")

    rect(s, Inches(0.5), Inches(1.82), Inches(12.3), Inches(1.55), SOFT, round_=True)
    txt(s, Inches(0.8), Inches(1.98), Inches(11.7), Inches(1.25),
        [[("核心判断", 12, TEAL, True, 4)],
         [("多数人出来上课，不是来把 AI 学成专家，而是用可接受的成本进入一个高质量行业圈子，顺便补齐 AI 时代的常识。",
           16, NAVY, True)]],
        spacing=1.15)

    items = [
        ("低成本替代 EMBA", "长江商学院一类项目一上来就是几十万、周期长。本班用约 2 万元，换六个月可持续走动的同业关系。"),
        ("中年转型刚需", "付费主力是有支付能力、又面临岗位迭代的中年从业者，不是应届学生。"),
        ("行业垂直才成交", "客源来自房地产商会、国企、银行支行行长，以及广告、财税、律师等「百搭」上下游。"),
        ("AI 是入场券，不是主菜", "半小时讲清工具地图即可。再深，听不懂；只讲地产、不带 AI，又缺当下话题。"),
    ]
    positions = [
        (Inches(0.5), Inches(3.55)),
        (Inches(6.8), Inches(3.55)),
        (Inches(0.5), Inches(5.2)),
        (Inches(6.8), Inches(5.2)),
    ]
    for (title, body), (x, y) in zip(items, positions):
        rect(s, x, y, Inches(6.0), Inches(1.5), CARD, line=LINE, line_w=1, round_=True)
        rect(s, x, y, Inches(0.10), Inches(1.5), GOLD)
        txt(s, x + Inches(0.28), y + Inches(0.14), Inches(5.5), Inches(1.22),
            [[(title, 15, NAVY, True, 4)], [(body, 12.5, GRAY, False)]])


# ===========================================================================
# 06 目标客群
# ===========================================================================
def audience():
    s = slide()
    header(s, 6, "02  项目定位", "谁适合来，从哪里来",
           "先吃透上海及长三角不动产圈层，再谈扩科创、做出海。")

    personas = [
        ("核心盘", TEAL, "房地产商会会员", "上海本地会员 + 苏浙皖异地商会。每月来一次上海，可接受。"),
        ("支付盘", NAVY, "中年转型从业者", "有学费支付能力，对 AI 焦虑，需要新圈子和新话术。"),
        ("信用盘", NAVY3, "国企 / 银行支行行长", "不便入会，但需要资产、不良与客户资源对接。"),
        ("连接盘", RGBColor(0x8A, 0x6A, 0x2A), "地产上下游「百搭」", "广告、信息技术、财税、律师、二房东、社区商业。"),
    ]
    x = Inches(0.5)
    for tag, color, title, desc in personas:
        rect(s, x, Inches(1.85), Inches(3.0), Inches(3.15), CARD, line=LINE, line_w=1, round_=True)
        rect(s, x, Inches(1.85), Inches(3.0), Inches(0.7), color, round_=True)
        txt(s, x, Inches(1.85), Inches(3.0), Inches(0.7),
            [(tag, 13, WHITE, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        txt(s, x + Inches(0.18), Inches(2.7), Inches(2.64), Inches(2.1),
            [[(title, 16, NAVY, True, 8)], [(desc, 13, GRAY, False)]])
        x += Inches(3.2)

    rect(s, Inches(0.5), Inches(5.2), Inches(12.3), Inches(1.5), NAVY, round_=True)
    txt(s, Inches(0.8), Inches(5.38), Inches(11.7), Inches(1.15),
        [[("获客原则", 12, GOLD, True, 4)],
         [("自有会员给面子、异地商协会秘书长做渠道（返佣 20%–25%）、国企与银行点对点邀请。外地客源不必强拉入会，用班级把人留下即可。",
           14, WHITE, False)]],
        spacing=1.12)


# ===========================================================================
# 07 产品形态
# ===========================================================================
def product():
    s = slide()
    header(s, 7, "03  产品设计", "班型怎么开",
           "沿用已跑通的「六个月、每月一天、校内上课」标准配置，降低试错成本。")

    specs = [
        ("周期", "6 个月", "从 5 个月拉长，给社群和互访留足时间。"),
        ("课时", "每月 1 天", "周五或周末集中授课，便于外地学员往返。"),
        ("场地", "高校校内", "证书、发票、场地一次性对齐，观感更稳。"),
        ("规模", "25–60 人", "25 人起班，60 人更利于圈层生态；首期建议冲 40–60。"),
        ("滚动", "每两月一期", "不是一次性项目，做成可复制的班级产品。"),
        ("班费", "5–10 万元", "从学费中预留，用于互访、聚餐和班委活动。"),
    ]
    x0, y0 = Inches(0.5), Inches(1.85)
    cw, ch = Inches(4.0), Inches(2.35)
    for i, (k, v, d) in enumerate(specs):
        r, c = divmod(i, 3)
        x = x0 + c * Inches(4.15)
        y = y0 + r * Inches(2.5)
        rect(s, x, y, cw, ch, CARD, line=LINE, line_w=1, round_=True)
        txt(s, x + Inches(0.28), y + Inches(0.22), Inches(3.45), Inches(0.35),
            [(k, 12, TEAL, True)])
        txt(s, x + Inches(0.28), y + Inches(0.55), Inches(3.45), Inches(0.5),
            [(v, 22, NAVY, True)])
        rect(s, x + Inches(0.28), y + Inches(1.15), Inches(0.7), Pt(2.5), GOLD)
        txt(s, x + Inches(0.28), y + Inches(1.35), Inches(3.45), Inches(0.8),
            [(d, 13, GRAY, False)])


# ===========================================================================
# 08 课程地图
# ===========================================================================
def curriculum():
    s = slide()
    header(s, 8, "03  产品设计", "六个月课程地图",
           "每月一天主课 + 贯穿始终的学员单位互访。AI 只占必要篇幅，不动产实操占主干。")

    modules = [
        ("01", "开班与 AI 素养", "工具地图、场景入门、班委组建、破冰互访。"),
        ("02", "不动产资产管理", "资产盘点、持有运营、现金流与估值框架。"),
        ("03", "不良资产处置", "识别、尽调、重组、司法与市场化路径。"),
        ("04", "REITs 与资本化", "公募 REITs、证券化、退出与结构安排。"),
        ("05", "AI 赋能资产运营", "租赁、招商、客服、研报的轻量落地，不深挖算法。"),
        ("06", "结业路演对接", "小组课题、资源对接、结业仪式、同学会启动。"),
    ]
    x0, y0 = Inches(0.5), Inches(1.82)
    cw, ch = Inches(4.0), Inches(2.32)
    for i, (no, title, desc) in enumerate(modules):
        r, c = divmod(i, 3)
        x = x0 + c * Inches(4.15)
        y = y0 + r * Inches(2.48)
        rect(s, x, y, cw, ch, CARD, line=LINE, line_w=1, round_=True)
        rect(s, x, y, Inches(0.72), ch, NAVY if i < 4 else TEAL)
        txt(s, x, y, Inches(0.72), ch,
            [(no, 16, GOLD, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        txt(s, x + Inches(0.9), y + Inches(0.35), Inches(2.9), Inches(1.7),
            [[(title, 16, NAVY, True, 8)], [(desc, 13, GRAY, False)]])


# ===========================================================================
# 09 师资与资源
# ===========================================================================
def faculty():
    s = slide()
    header(s, 9, "03  产品设计", "谁来讲，凭什么相信",
           "师资成本可控：以头部企业高管分享 + 不动产实操老师为主，不堆高价商学院牌。")

    left_items = [
        ("AI 侧", "腾讯、阿里，以及宇树、智元机器人等参展企业高管。讲产业趋势与可用工具，不讲论文。"),
        ("不动产侧", "资产管理、不良处置、REITs 实操老师。这是学员本行，也是续班和转介绍的根基。"),
        ("高校侧", "校内场地、结业证与发票路径已可对齐。用不用高校抬头，按学员需求二选一。"),
    ]
    y = Inches(1.85)
    for title, body in left_items:
        rect(s, Inches(0.5), y, Inches(7.55), Inches(1.5), CARD, line=LINE, line_w=1, round_=True)
        rect(s, Inches(0.5), y, Inches(0.12), Inches(1.5), TEAL)
        txt(s, Inches(0.85), y + Inches(0.22), Inches(6.95), Inches(1.1),
            [[(title, 16, NAVY, True, 4)], [(body, 13, GRAY, False)]])
        y += Inches(1.62)

    rect(s, Inches(8.25), Inches(1.85), Inches(4.55), Inches(4.85), NAVY, round_=True)
    txt(s, Inches(8.5), Inches(2.1), Inches(4.05), Inches(0.4),
        [("课堂上要守住的边界", 14, GOLD, True)])
    notes = [
        "AI 讲皮毛即可，半小时能讲清工具地图。",
        "太深听不懂，全是 AI 又听腻。",
        "不动产课能不能卖，取决于客群是否垂直。",
        "本班优先服务房地产商会及上下游，垂直盘能撑住定价。",
        "讲师费用按「认圈子、滚动流量」来谈，不按商学院时薪来堆。",
    ]
    yy = Inches(2.65)
    for n in notes:
        txt(s, Inches(8.55), yy, Inches(4.0), Inches(0.7),
            [[("▸  ", 13, TEAL, True), (n, 13, WHITE, False)]])
        yy += Inches(0.72)


# ===========================================================================
# 10 社群运营
# ===========================================================================
def community():
    s = slide()
    header(s, 10, "03  产品设计", "班集体怎么运转",
           "课只是由头。真正把人留住的，是班委、互访和滚动开班形成的圈层密度。")

    steps = [
        ("1", "开班建制", "设班长、学习委员、组长。把 40–60 人拆成可走动的小组。"),
        ("2", "每月主课", "校内一天集中授课，统一节奏，便于外地学员安排。"),
        ("3", "学员互访", "课间穿插单位参访，把课堂关系变成业务关系。"),
        ("4", "班费活动", "预留 5–10 万元做聚餐、走访和小型沙龙，不另收费。"),
        ("5", "滚动续班", "每两个月新开一期，老学员可回流当链接人。"),
        ("6", "出口对接", "结业后筛种子选手，进入科创对接与考察名单。"),
    ]
    x0, y0 = Inches(0.5), Inches(1.85)
    for i, (no, title, desc) in enumerate(steps):
        r, c = divmod(i, 3)
        x = x0 + c * Inches(4.15)
        y = y0 + r * Inches(2.45)
        rect(s, x, y, Inches(4.0), Inches(2.25), CARD, line=LINE, line_w=1, round_=True)
        oval = s.shapes.add_shape(MSO_SHAPE.OVAL, x + Inches(0.25), y + Inches(0.28), Inches(0.48), Inches(0.48))
        oval.fill.solid()
        oval.fill.fore_color.rgb = TEAL if i < 4 else GOLD
        oval.line.fill.background()
        oval.shadow.inherit = False
        tf = oval.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r0 = p.add_run()
        r0.text = no
        _set_font(r0, 14, WHITE, True)
        txt(s, x + Inches(0.88), y + Inches(0.32), Inches(2.9), Inches(0.42),
            [(title, 16, NAVY, True)], anchor=MSO_ANCHOR.MIDDLE)
        txt(s, x + Inches(0.25), y + Inches(0.95), Inches(3.5), Inches(1.05),
            [(desc, 13, GRAY, False)])


# ===========================================================================
# 11 两种落地模式
# ===========================================================================
def models():
    s = slide()
    header(s, 11, "04  落地模式", "两种主体，按需求选用",
           "高校品牌能抬证书观感，但分成很重。第三方主体把收益留在项目方，灵活度更高。")

    cards = [
        (NAVY, "模式 A  ·  第三方主体运营（建议默认）",
         ["学费进入项目方指定账户，收益全部纳入项目方。",
          "结业证由中心 / 联合机构颁发，不占用高校抬头。",
          "发票路径清晰，不必向学校交高额份子钱。",
          "适合：渠道招生、商会会员班、快速滚动开班。",
          "代价：证书权威性弱于高校官方结业证。"]),
        (TEAL_DEEP, "模式 B  ·  高校抬头 + 官方结业证",
         ["用高校品牌开发票、发结业证，观感更好、更好对外讲。",
          "必须向学校分成。高校规则同类，交大口径可达约 80%。",
          "学校会过问安全、住宿、收费上限等合规事项。",
         "适合：学员明确要求「学校证书 / 学校发票」。",
          "代价：项目方利润空间显著收窄，决策前必须算清。"]),
    ]
    xs = [Inches(0.5), Inches(6.85)]
    for (clr, title, pts), x in zip(cards, xs):
        rect(s, x, Inches(1.82), Inches(5.95), Inches(4.9), CARD, line=LINE, line_w=1, round_=True)
        rect(s, x, Inches(1.82), Inches(5.95), Inches(0.85), clr, round_=True)
        rect(s, x, Inches(2.3), Inches(5.95), Inches(0.37), clr)
        txt(s, x + Inches(0.25), Inches(1.95), Inches(5.45), Inches(0.6),
            [(title, 15, WHITE, True)], anchor=MSO_ANCHOR.MIDDLE)
        yy = Inches(2.85)
        for p in pts:
            txt(s, x + Inches(0.3), yy, Inches(5.35), Inches(0.7),
                [[("●  ", 12, GOLD, True), (p, 13, INK, False)]])
            yy += Inches(0.72)


# ===========================================================================
# 12 收费与渠道
# ===========================================================================
def pricing():
    s = slide()
    header(s, 12, "04  落地模式", "怎么报价，怎么分账",
           "对外给一个干净的价格带；对内把渠道、班费、不可承诺事项一次说清。")

    # 价格带
    bands = [
        ("入门沟通价", "约 2.0 万元", "用于试探与关系户面子价，不作为对外主报价。"),
        ("标准对外价", "2.48 万元", "主推档。覆盖讲师、场地、班费与渠道空间。"),
        ("渠道结算", "返佣 20%–25%", "异地商协会秘书长、合作机构按此区间谈。"),
    ]
    x = Inches(0.5)
    accents = [NAVY, TEAL, RGBColor(0x8A, 0x6A, 0x2A)]
    for (k, v, d), acc in zip(bands, accents):
        rect(s, x, Inches(1.85), Inches(4.0), Inches(2.15), CARD, line=LINE, line_w=1, round_=True)
        txt(s, x + Inches(0.25), Inches(2.02), Inches(3.5), Inches(0.32),
            [(k, 12, acc, True)])
        txt(s, x + Inches(0.25), Inches(2.35), Inches(3.5), Inches(0.5),
            [(v, 22, NAVY, True)])
        txt(s, x + Inches(0.25), Inches(2.95), Inches(3.5), Inches(0.8),
            [(d, 13, GRAY, False)])
        x += Inches(4.15)

    notes = [
        ("班级活动经费", "学费中预留 5–10 万元，专款用于互访与班级活动，不另向学员加收。"),
        ("报价原则", "当前高端商学院招生难，本班主打高性价比圈层，不与几十万项目正面比课时。"),
        ("不对外承诺", "不承诺投资收益、移民结果或高校学历；证书性质在报名时书面说清。"),
        ("渠道纪律", "返佣只走约定渠道，不层层加价。关系户面子价需项目方书面确认。"),
    ]
    y = Inches(4.2)
    for i, (k, v) in enumerate(notes):
        c = i % 2
        r = i // 2
        x = Inches(0.5) + c * Inches(6.35)
        yy = y + r * Inches(1.2)
        rect(s, x, yy, Inches(6.15), Inches(1.08), CARD, line=LINE, line_w=1, round_=True)
        txt(s, x + Inches(0.22), yy + Inches(0.12), Inches(5.7), Inches(0.85),
            [[(k, 14, NAVY, True, 2)], [(v, 12.5, GRAY, False)]])


# ===========================================================================
# 13 收益测算
# ===========================================================================
def economics():
    s = slide()
    header(s, 13, "04  落地模式", "两种班型的收益对照（示意）",
           "数字用于合作对齐，不是对学员的承诺。高校分成按最严口径单列，避免事后被动。")

    rows = [
        ("项目", "轻量班 25 人", "生态班 60 人"),
        ("对外单价", "2.00 万元", "2.48 万元"),
        ("学费收入", "50.0 万元", "148.8 万元"),
        ("渠道返佣 20%–25%", "10.0–12.5 万元", "29.8–37.2 万元"),
        ("班级活动经费", "5 万元", "8–10 万元"),
        ("模式 A 项目方可留存（示意）", "约 32–35 万元", "约 102–111 万元"),
        ("模式 B 若按高校 80% 分成", "项目方空间显著收窄", "项目方空间显著收窄"),
    ]
    table_w = Inches(12.3)
    col_w = [Inches(4.5), Inches(3.9), Inches(3.9)]
    row_h = Inches(0.58)
    x0, y0 = Inches(0.5), Inches(1.82)
    for r, row in enumerate(rows):
        y = y0 + r * row_h
        bg = NAVY if r == 0 else (SOFT if r in (5, 6) else CARD)
        fg = WHITE if r == 0 else INK
        bold = r in (0, 5, 6)
        x = x0
        for c, cell in enumerate(row):
            w = col_w[c]
            fill = TEAL if (r == 0 and c > 0) else bg
            if r == 0 and c == 0:
                fill = NAVY
            rect(s, x, y, w, row_h, fill, line=LINE, line_w=1)
            color = WHITE if r == 0 or (r == 6 and c > 0) else (TEAL_DEEP if r == 5 and c > 0 else fg)
            if r == 6 and c > 0:
                color = RGBColor(0x8B, 0x3A, 0x3A)
            txt(s, x + Inches(0.12), y, w - Inches(0.2), row_h,
                [(cell, 13 if r else 13, color, bold)],
                anchor=MSO_ANCHOR.MIDDLE,
                align=PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER)
            x += w

    txt(s, Inches(0.55), Inches(6.05), Inches(12.2), Inches(0.7),
        [("建议：默认走模式 A，把班级做成可滚动产品；仅当学员明确要高校证书时，再单列模式 B 并重谈分成。首期按 40–60 人做圈层密度，比硬冲 25 人小班更划算。",
          13, GRAY, False)])


# ===========================================================================
# 14 三层生态
# ===========================================================================
def ecosystem():
    s = slide()
    header(s, 14, "05  推进节奏", "培训只是入口，后面还有两层",
           "战略可以画三层，执行必须先把第一层做成。至少跑满三期，再谈大赛与 FA。")

    layers = [
        ("第一层  ·  现在就做", TEAL, "AI+不动产培训",
         ["精准行业用户与班级品牌", "学费 + 渠道 + 班费可核算", "为后续两层提供名单与信任"]),
        ("第二层  ·  三期之后", NAVY, "长三角科创对接",
         ["从学员中筛 100–200 名种子", "特展约 4 万 / 场，系列约 50 万", "外地企业来沪参访 1–3 万 / 次"]),
        ("第三层  ·  偶得收益", GOLD, "FA 撮合佣金",
         ["只在有合作意向时抽取", "成功率低，不可当主营", "是补充，不是预算支柱"]),
    ]
    x = Inches(0.5)
    for tag, color, title, pts in layers:
        rect(s, x, Inches(1.85), Inches(4.0), Inches(3.95), CARD, line=LINE, line_w=1, round_=True)
        rect(s, x, Inches(1.85), Inches(4.0), Inches(1.2), color, round_=True)
        rect(s, x, Inches(2.6), Inches(4.0), Inches(0.45), color)
        txt(s, x + Inches(0.22), Inches(1.98), Inches(3.55), Inches(0.95),
            [[(tag, 12, GOLD_PALE if color != GOLD else NAVY, True, 2)],
             [(title, 18, WHITE if color != GOLD else NAVY, True)]])
        yy = Inches(3.25)
        for p in pts:
            txt(s, x + Inches(0.28), yy, Inches(3.45), Inches(0.7),
                [[("●  ", 12, TEAL, True), (p, 13.5, INK, False)]])
            yy += Inches(0.78)
        x += Inches(4.15)

    rect(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.7), SOFT, round_=True)
    txt(s, Inches(0.75), Inches(6.1), Inches(11.8), Inches(0.5),
        [("圈层延展（可选）：澳大利亚商务考察可作为高净值学员的深度连接产品，不并入本班主报价，单独成团、单独核算。",
          13, NAVY, False)],
        anchor=MSO_ANCHOR.MIDDLE)


# ===========================================================================
# 15 时间表
# ===========================================================================
def timeline():
    s = slide()
    header(s, 15, "05  推进节奏", "9–11 月黄金档，按周推进",
           "招生要两个月窗口。9 月中下旬启动，11 月开班，是当前最稳的节奏。")

    phases = [
        ("即刻", "方案对齐", ["确认主体：默认模式 A", "确认单价带与渠道口径", "确认课程六模块负责人"]),
        ("9 月中下旬", "启动招生", ["商会会员点对点邀请", "异地秘书长渠道启动", "高校场地与发票路径锁定"]),
        ("10 月", "满班冲刺", ["目标 40–60 人", "班委候选人预热", "首月课表与互访名单"]),
        ("11 月", "正式开班", ["开班仪式 + AI 素养", "班费账户与互访节奏", "同步筹备第二期名单"]),
    ]
    x = Inches(0.5)
    colors = [TEAL, NAVY, NAVY3, GOLD]
    for (when, title, pts), color in zip(phases, colors):
        rect(s, x, Inches(1.85), Inches(3.0), Inches(4.85), CARD, line=LINE, line_w=1, round_=True)
        rect(s, x, Inches(1.85), Inches(3.0), Inches(1.35), color, round_=True)
        rect(s, x, Inches(2.75), Inches(3.0), Inches(0.45), color)
        txt(s, x + Inches(0.15), Inches(2.0), Inches(2.7), Inches(1.05),
            [[(when, 12, GOLD_PALE if color != GOLD else NAVY, True, 2)],
             [(title, 18, WHITE if color != GOLD else NAVY, True)]],
            align=PP_ALIGN.CENTER)
        yy = Inches(3.4)
        for p in pts:
            txt(s, x + Inches(0.2), yy, Inches(2.6), Inches(0.85),
                [[("●  ", 12, TEAL, True), (p, 13, INK, False)]])
            yy += Inches(0.95)
        x += Inches(3.2)


# ===========================================================================
# 16 下一步
# ===========================================================================
def next_steps():
    s = slide()
    header(s, 16, "05  推进节奏", "今天需要拍板的四件事",
           "方案可以继续改细节，但下面四项不定，招生就开不了口。")

    decisions = [
        ("01", "主体", "默认模式 A（第三方账户），模式 B 仅作学员点单选项。"),
        ("02", "定价", "对外主报价 2.48 万元，关系户不低于 2.0 万元。"),
        ("03", "班型", "首期按 40–60 人招生，25 人为开班底线。"),
        ("04", "时间", "9 月中下旬启动招生，11 月开班，每两月滚动一期。"),
    ]
    y = Inches(1.82)
    for no, title, desc in decisions:
        rect(s, Inches(0.5), y, Inches(8.35), Inches(1.12), CARD, line=LINE, line_w=1, round_=True)
        rect(s, Inches(0.5), y, Inches(1.05), Inches(1.12), NAVY, round_=True)
        txt(s, Inches(0.5), y, Inches(1.05), Inches(1.12),
            [(no, 16, GOLD, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        txt(s, Inches(1.75), y + Inches(0.16), Inches(6.85), Inches(0.82),
            [[(title, 16, NAVY, True, 2)], [(desc, 13, GRAY, False)]])
        y += Inches(1.2)

    rect(s, Inches(9.05), Inches(1.82), Inches(3.75), Inches(4.85), NAVY, round_=True)
    txt(s, Inches(9.28), Inches(2.05), Inches(3.3), Inches(0.45),
        [("会后立即分头做", 14, GOLD, True)])
    todos = [
        "确认课程六模块与讲师档期",
        "测算模式 A / B 精确分成",
        "列出首批 80 人邀请名单",
        "锁定校内场地与开票主体",
        "准备一页招生海报与报名表",
    ]
    yy = Inches(2.6)
    for t in todos:
        txt(s, Inches(9.3), yy, Inches(3.25), Inches(0.7),
            [[("▸  ", 13, TEAL, True), (t, 13, WHITE, False)]])
        yy += Inches(0.72)


# ===========================================================================
# 17 结束页
# ===========================================================================
def closing():
    s = slide()
    rect(s, 0, 0, SW, SH, NAVY)
    rect(s, 0, 0, SW, Inches(0.12), TEAL)
    rect(s, 0, Inches(5.55), SW, Inches(0.06), GOLD)

    txt(s, Inches(0.9), Inches(1.7), Inches(11.5), Inches(0.4),
        [("先把一个班做成品牌，再谈后面的生态。", 16, GOLD, True)])
    txt(s, Inches(0.9), Inches(2.25), Inches(11.5), Inches(1.6),
        [[("不动产做主干", 32, WHITE, True, 8)],
         [("AI 做素养，圈层做长期价值", 28, WHITE, True)]])
    txt(s, Inches(0.9), Inches(4.15), Inches(11.5), Inches(0.7),
        [("AI × 不动产资产管理高级研修班  ·  2026 年 11 月首期", 16, RGBColor(0xC5, 0xD4, 0xE4), False)])

    chips = ["主体确认", "定价确认", "名单启动", "11 月开班"]
    x = Inches(0.9)
    for c in chips:
        rect(s, x, Inches(5.9), Inches(2.7), Inches(0.85), NAVY2, round_=True, radius=0.1)
        txt(s, x, Inches(5.9), Inches(2.7), Inches(0.85),
            [(c, 16, WHITE, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        x += Inches(3.0)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    cover()
    toc()
    opportunity()
    positioning()
    insight()
    audience()
    product()
    curriculum()
    faculty()
    community()
    models()
    pricing()
    economics()
    ecosystem()
    timeline()
    next_steps()
    closing()
    prs.save(OUT)
    print(f"已生成 {len(prs.slides)} 页：{OUT}")


if __name__ == "__main__":
    main()
