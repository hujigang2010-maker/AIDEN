# -*- coding: utf-8 -*-
"""港交所18C估值手册的投委会简报。数字来自 model_core，不另造一套。"""

from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
import model_core as mc

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "output" / "港交所18C分行业估值模板_投委会简报.pptx"

FONT = "微软雅黑"
SW, SH = 13.333333, 7.5
TOTAL = 30

PINE = RGBColor(0x0E, 0x2A, 0x24)
PINE_MID = RGBColor(0x1A, 0x43, 0x3A)
INK = RGBColor(0x1C, 0x24, 0x21)
CREAM = RGBColor(0xF7, 0xF4, 0xEE)
PAPER = RGBColor(0xFF, 0xFC, 0xF8)
GOLD = RGBColor(0xC6, 0xA1, 0x5B)
GOLD_DK = RGBColor(0x8A, 0x5A, 0x32)
TEAL = RGBColor(0x0F, 0x6E, 0x62)
TEAL_SOFT = RGBColor(0xE4, 0xF2, 0xEC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED = RGBColor(0xA3, 0x3B, 0x32)
RED_SOFT = RGBColor(0xF8, 0xEB, 0xE8)
LINE = RGBColor(0xE3, 0xDB, 0xCE)
MUTED = RGBColor(0x5C, 0x68, 0x62)
SLATE = RGBColor(0x3E, 0x4A, 0x45)


def rgb_hex(color):
    return f"{color[0]:02X}{color[1]:02X}{color[2]:02X}"


def set_run_font(run, size, color, bold=False, italic=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = FONT
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        node = rPr.find(qn(tag))
        if node is None:
            node = etree.SubElement(rPr, qn(tag))
        node.set("typeface", FONT)


def kill_style(shape):
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)
    try:
        shape.shadow.inherit = False
    except Exception:
        pass


def rect(slide, x, y, w, h, fill, line=None, radius=None):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    kill_style(shape)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(1.0)
    if radius:
        try:
            shape.adjustments[0] = radius
        except Exception:
            pass
    return shape


def bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def anchor_of(name):
    return {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}[name]


def align_of(name):
    return {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}[name]


def text(slide, x, y, w, h, content, size=14, color=INK, bold=False, align="left", anchor="top", italic=False, spacing=1.0):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor_of(anchor)
    lines = content if isinstance(content, (list, tuple)) else str(content).split("\n")
    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = align_of(align)
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        p.line_spacing = spacing
        run = p.add_run()
        run.text = line
        set_run_font(run, size, color, bold, italic)
    return shape


def rich(slide, x, y, w, h, paragraphs, anchor="top"):
    """paragraphs: list of list of (text, size, color, bold)."""
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor_of(anchor)
    for i, runs in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        p.line_spacing = 1.05
        if isinstance(runs, tuple):
            runs = [runs]
        for value, size, color, bold in runs:
            run = p.add_run()
            run.text = value
            set_run_font(run, size, color, bold)
    return shape


def chrome(slide, kicker, page):
    bg(slide, CREAM)
    rect(slide, 0, 0, 0.08, SH, PINE)
    text(slide, 0.38, 0.16, 7.2, 0.28, kicker, 11, GOLD_DK, True, anchor="middle")
    text(slide, 8.4, 0.16, 4.5, 0.28, "港交所 18C  ·  估值实务", 11, MUTED, False, "right", "middle")
    rect(slide, 0.38, 0.48, 12.55, 0.012, GOLD)
    text(slide, 0.38, 7.16, 8.5, 0.24, "虚构演示与方法模板  ·  非投资建议，非法律意见", 10, MUTED, anchor="middle")
    text(slide, 10.3, 7.16, 2.6, 0.24, f"{page:02d}   /   {TOTAL:02d}", 10, MUTED, False, "right", "middle")


def title(slide, content, size=26):
    text(slide, 0.38, 0.58, 12.55, 0.72, content, size, INK, True, anchor="middle")


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def footer_source(slide, source):
    text(slide, 0.38, 6.58, 12.5, 0.28, source, 11, MUTED, anchor="middle")


def n2(value):
    return f"{value:,.2f}"


def n3(value):
    return f"{value:,.3f}"


def s01(prs):
    slide = blank(prs)
    bg(slide, PINE)
    rect(slide, 0, 0, 0.12, SH, GOLD)
    text(slide, 0.72, 0.48, 11, 0.3, "HONG KONG EXCHANGES    ·    CHAPTER 18C", 13, GOLD, True)
    text(slide, 0.72, 2.05, 12, 0.85, "分行业估值模板", 48, WHITE, True)
    text(slide, 0.72, 2.95, 12, 0.6, "与模型回调实务", 32, RGBColor(0xF3, 0xEB, 0xD8), False)
    rect(slide, 0.72, 3.75, 2.1, 0.015, GOLD)
    metas = [("版本", "V1.0"), ("研究基准日", "2026-11-06"), ("测算币种", "人民币  ·  规则用港币")]
    for i, (k, v) in enumerate(metas):
        x = 0.72 + i * 3.5
        text(slide, x, 4.15, 3.2, 0.24, k, 12, GOLD)
        text(slide, x, 4.42, 3.2, 0.36, v, 18, WHITE, True)
    rect(slide, 0, 6.35, SW, 1.15, PINE_MID)
    text(slide, 0.72, 6.52, 11.8, 0.7, "股权融资、并购、投融资评审、Pre-IPO，以及港股18C申报可行性的初筛。\n表内比例、倍数和折现率，除标明为上市规则者外，都是虚构演示，不是行业均值。", 14, RGBColor(0xE7, 0xE1, 0xD4), spacing=1.15)


def s02(prs):
    slide = blank(prs)
    chrome(slide, "使用边界", 2)
    title(slide, "先把六种输出分开，再谈一个数")
    items = [
        ("VAL_EV", "经营企业价值", "经营性资产，不等于市值"),
        ("VAL_EQ_PRE", "融资前全部股权", "扣净债务，不加本轮新钱"),
        ("VAL_EQ_POST", "融资后全部股权", "新现金只能加一次"),
        ("VAL_SHARE", "完全摊薄每股", "优先股、期权、权证之后"),
        ("VAL_RANGE", "估值区间", "悲观、基准、乐观都留痕"),
        ("HK18C_GATE", "18C诊断", "不替代估值本身"),
    ]
    for i, (code, name, note) in enumerate(items):
        col, row = i % 3, i // 3
        x, y = 0.38 + col * 4.2, 1.55 + row * 1.85
        rect(slide, x, y, 4.0, 1.68, PAPER, LINE, 0.06)
        rect(slide, x, y, 0.08, 1.68, GOLD)
        text(slide, x + 0.28, y + 0.18, 3.5, 0.32, code, 13, GOLD_DK, True)
        text(slide, x + 0.28, y + 0.54, 3.5, 0.42, name, 20, INK, True)
        text(slide, x + 0.28, y + 1.08, 3.5, 0.4, note, 13, MUTED)
    text(slide, 0.42, 5.4, 12.4, 0.7, "灯号只有四种：PASS、CONDITIONAL、FAIL、NOT-ASSESSABLE。\n不能因为离 40 亿或 80 亿港元还差一截，就把模型调到刚好过线。", 16, SLATE, spacing=1.15)
    rect(slide, 0.38, 6.25, 12.55, 0.78, PINE, radius=0.06)
    text(slide, 0.62, 6.25, 12.1, 0.78, "融资前、融资后、企业价值和发行市值，是四套口径。", 18, WHITE, True, anchor="middle")


def s03(prs):
    slide = blank(prs)
    chrome(slide, "目录", 3)
    title(slide, "四幕。规则、方法、行业、回调。")
    chapters = [
        ("01", "规则", "收入、市值、研发、\n投资者、营运资金", "06–09"),
        ("02", "方法", "DCF、倍数、校准、\n一次只反解一个数", "11–13"),
        ("03", "行业", "AI做透，其余五条\n只改驱动和顺序", "15–22"),
        ("04", "回调", "四条路径分开走，\n桥必须轧平", "24–29"),
    ]
    for i, (num, name, body, pages) in enumerate(chapters):
        x = 0.38 + i * 3.22
        rect(slide, x, 1.7, 3.05, 4.55, PAPER, LINE, 0.05)
        rect(slide, x, 1.7, 3.05, 0.1, GOLD)
        text(slide, x + 0.22, 2.0, 2.6, 0.4, num, 14, GOLD_DK, True)
        text(slide, x + 0.22, 2.5, 2.6, 0.7, name, 32, INK, True)
        text(slide, x + 0.22, 3.5, 2.6, 1.3, body, 16, SLATE, spacing=1.2)
        text(slide, x + 0.22, 5.45, 2.6, 0.35, pages, 14, MUTED)


def s04(prs):
    slide = blank(prs)
    chrome(slide, "模型架构", 4)
    title(slide, "八步走完，才允许出现区间")
    steps = [
        ("01", "采集", "审计、订单、合同、股权条款"),
        ("02", "口径", "收入、研发、净债务、币种"),
        ("03", "驱动", "用户、Token、出货、临床"),
        ("04", "三表", "利润、资产负债、现金跑道"),
        ("05", "方法", "DCF、市场、交易、概率"),
        ("06", "校准", "情景、敏感性、股权桥"),
        ("07", "分配", "优先权、转换、期权池"),
        ("08", "分开", "18C门槛单独出，不倒推"),
    ]
    for i, (num, name, body) in enumerate(steps):
        col, row = i % 4, i // 4
        x, y = 0.38 + col * 3.22, 1.6 + row * 2.4
        rect(slide, x, y, 3.05, 2.15, PAPER, LINE, 0.06)
        text(slide, x + 0.2, y + 0.18, 2.6, 0.32, num, 13, GOLD_DK, True)
        text(slide, x + 0.2, y + 0.55, 2.6, 0.48, name, 24, INK, True)
        text(slide, x + 0.2, y + 1.2, 2.65, 0.7, body, 14, SLATE)


def divider(prs, page, numeral, name, sentence):
    slide = blank(prs)
    bg(slide, PINE)
    rect(slide, 0, 0, 0.12, SH, GOLD)
    text(slide, 0.72, 1.7, 4, 0.4, numeral, 18, GOLD, True)
    text(slide, 0.72, 2.3, 11, 1.1, name, 54, WHITE, True)
    rect(slide, 0.72, 3.65, 2.1, 0.015, GOLD)
    text(slide, 0.72, 4.0, 10.5, 1.1, sentence, 22, RGBColor(0xF3, 0xEB, 0xD8), spacing=1.2)
    text(slide, 10.4, 6.9, 2.4, 0.3, f"{page:02d}   /   {TOTAL:02d}", 12, GOLD, False, "right")


def s06(prs):
    slide = blank(prs)
    chrome(slide, "规则  ·  门槛", 6)
    title(slide, "先分已商业化，再读市值和研发")
    headers = ["核验项", "已商业化", "未商业化"]
    rows = [
        ("特专科技收入", "≥ 2.5 亿港元", "< 2.5 亿港元"),
        ("最低市值，临时", "40 亿港元", "80 亿港元"),
        ("规则正文原门槛", "60 亿港元", "100 亿港元"),
        ("研发 / 总经营支出", "≥ 15%", "≥ 30% 或 ≥ 50%"),
        ("记录", "经营与研发均至少 3 年", "同左"),
        ("营运资金", "适用一般要求", "12 个月成本的 125%"),
        ("配售", "至少 50% 由独立定价投资者认购", "同左"),
    ]
    x0, y0, widths, h = 0.38, 1.5, (3.3, 4.6, 4.6), 0.58
    for i, head in enumerate(headers):
        x = x0 + sum(widths[:i])
        rect(slide, x, y0, widths[i] - 0.06, 0.48, PINE)
        text(slide, x + 0.16, y0, widths[i] - 0.3, 0.48, head, 14, WHITE, True, anchor="middle")
    for r, row in enumerate(rows):
        y = y0 + 0.54 + r * h
        bg_color = PAPER if r % 2 == 0 else CREAM
        for i, value in enumerate(row):
            x = x0 + sum(widths[:i])
            rect(slide, x, y, widths[i] - 0.06, h - 0.06, bg_color)
            weight = i == 0
            color = INK if i == 0 else PINE
            text(slide, x + 0.16, y, widths[i] - 0.32, h - 0.06, value, 14, color, weight, anchor="middle")
    footer_source(slide, "临时门槛 2024-09-01 至 2027-08-31。研发至少两年各自达标，且三年合计达标。HKEX-1 至 HKEX-8。")


def s07(prs):
    slide = blank(prs)
    chrome(slide, "规则  ·  收入与领域", 7)
    title(slide, "18C看的是特专科技收入，不是总收入")
    rect(slide, 0.38, 1.55, 5.9, 5.15, PAPER, LINE, 0.05)
    text(slide, 0.62, 1.75, 5.4, 0.4, "这些都不要直接算进测试", 16, RED, True)
    bans = [("合同额", "还不是收入"), ("在手订单", "可能取消或延期"), ("ARR", "未按履约确认"), ("GMV", "不是公司收入"), ("未确认订阅", "仍在负债里")]
    for i, (name, note) in enumerate(bans):
        y = 2.35 + i * 0.78
        rect(slide, 0.62, y, 1.7, 0.6, RED_SOFT, radius=0.1)
        text(slide, 0.62, y, 1.7, 0.6, name, 14, RED, True, "center", "middle")
        text(slide, 2.5, y, 3.4, 0.6, note, 16, SLATE, anchor="middle")
    modules = [
        ("AI", "新一代信息技术", "转售算力、无核心研发的集成，不当然适用"),
        ("硬科技", "先进硬件、半导体、材料", "纯贸易或代工要审科技实质"),
        ("新能源", "生产、储存、输配", "资产运营和核心研发要分开"),
        ("生物医药", "多数先看18A", "平台可以讨论18C，药物本身不行"),
        ("高端装备", "机器人、自动化、航空", "传统机加工不会因为标签合格"),
        ("专精特新", "不是18C分类", "先回到上面某一个具体技术"),
    ]
    for i, (name, field, edge) in enumerate(modules):
        col, row = i % 2, i // 2
        x, y = 6.5 + col * 3.25, 1.55 + row * 1.72
        rect(slide, x, y, 3.1, 1.58, PAPER, LINE, 0.05)
        text(slide, x + 0.16, y + 0.12, 2.8, 0.32, name, 16, PINE, True)
        text(slide, x + 0.16, y + 0.48, 2.8, 0.3, field, 13, GOLD_DK, True)
        text(slide, x + 0.16, y + 0.86, 2.8, 0.55, edge, 12, SLATE)


def s08(prs):
    slide = blank(prs)
    chrome(slide, "规则  ·  研发与投资者", 8)
    title(slide, "比率要用指引口径，投资者不能只看一个比例")
    rect(slide, 0.38, 1.55, 6.35, 5.15, PAPER, LINE, 0.05)
    text(slide, 0.6, 1.75, 5.9, 0.35, "合资格研发 / 总经营支出", 16, PINE, True)
    lines = [
        "分子：费用化研发 + 符合条件的资本化开发 + 指引允许的其他项 − 已重复计入的摊销",
        "分母：财报总费用 − 销售成本 − 融资性质费用 + 未在费用中确认的合资格研发",
        "三年里至少两年各自达标，并且三年合计数也达标",
    ]
    for i, line in enumerate(lines):
        y = 2.3 + i * 1.05
        text(slide, 0.62, y, 0.4, 0.4, f"{i+1}", 16, GOLD_DK, True)
        text(slide, 1.1, y, 5.3, 0.9, line, 14, SLATE)
    rect(slide, 0.6, 5.55, 5.9, 0.85, RED_SOFT, radius=0.08)
    text(slide, 0.8, 5.55, 5.5, 0.85, "不要用“研发费用 / 全部成本费用”代替法定比率。", 15, RED, True, anchor="middle")

    rect(slide, 6.95, 1.55, 5.95, 5.15, PINE, radius=0.05)
    text(slide, 7.2, 1.78, 5.5, 0.35, "资深独立投资者，指示性基准", 16, GOLD, True)
    points = [
        ("2–5 名", "Pathfinder，申请前至少 12 个月"),
        ("10% 或 15 亿", "合计持股，或合计投资额。单一比例代替不了完整规则"),
        ("两名", "各自 3% 以上，或各自 4.5 亿港元"),
        ("还要看市值阶段", "全部资深投资者的比例，随预计市值和商业化阶段变化"),
    ]
    for i, (head, body) in enumerate(points):
        y = 2.4 + i * 0.95
        text(slide, 7.2, y, 5.4, 0.32, head, 18, WHITE, True)
        text(slide, 7.2, y + 0.34, 5.4, 0.45, body, 13, RGBColor(0xE7, 0xE1, 0xD4))


def s09(prs):
    slide = blank(prs)
    chrome(slide, "规则  ·  18A / 18C", 9)
    title(slide, "使用了AI，也不会把药物项目从18A搬到18C")
    rect(slide, 0.38, 1.6, 6.15, 4.35, PAPER, LINE, 0.05)
    rect(slide, 0.38, 1.6, 6.15, 0.62, PINE)
    text(slide, 0.6, 1.6, 5.7, 0.62, "18A   生物科技", 20, WHITE, True, anchor="middle")
    bits = ["初始市值至少 15 亿港元", "原则上至少两个财政年度", "至少一项核心产品超过概念阶段", "12 个月相关成本的 125% 营运资金", "核心产品通常要临床试验批准"]
    for i, bit in enumerate(bits):
        text(slide, 0.7, 2.5 + i * 0.6, 5.5, 0.5, bit, 16, INK, anchor="middle")
    rect(slide, 6.75, 1.6, 6.15, 4.35, PAPER, LINE, 0.05)
    rect(slide, 6.75, 1.6, 6.15, 0.62, GOLD_DK)
    text(slide, 6.95, 1.6, 5.7, 0.62, "18C   可能讨论的情形", 20, WHITE, True, anchor="middle")
    bits2 = ["价值在 AI 药物发现平台", "智能研发自动化、算法或软件", "收入来源不是自有受监管药物", "公开例子：晶泰科技 2024 招股书", "仍要按业务结构和交易所意见判断"]
    for i, bit in enumerate(bits2):
        text(slide, 7.05, 2.5 + i * 0.6, 5.5, 0.5, bit, 16, INK, anchor="middle")
    rect(slide, 0.38, 6.15, 12.52, 0.85, RED_SOFT, radius=0.06)
    text(slide, 0.6, 6.15, 12.1, 0.85, "15 亿港元不能套进 18C。40 亿或 80 亿也不能套回 18A。", 18, RED, True, anchor="middle")


def s11(prs):
    slide = blank(prs)
    chrome(slide, "方法", 11)
    title(slide, "先问信息在哪，再选方法")
    groups = [
        ("现金流", [("FCFF", "经营能建三表"), ("FCFE", "必须配股权成本"), ("rNPV", "概率和研发配套")]),
        ("市场", [("可比公司", "同期同口径"), ("可比交易", "控制权要剔除"), ("融资校准", "旧价格校准技术")]),
        ("路径与下限", [("PWERM", "路径乘回款"), ("OPM", "按断点分配"), ("NAV / VC", "下限或早期交叉")]),
    ]
    for i, (head, rows) in enumerate(groups):
        x = 0.38 + i * 4.25
        rect(slide, x, 1.55, 4.05, 4.55, PAPER, LINE, 0.05)
        rect(slide, x, 1.55, 4.05, 0.62, PINE)
        text(slide, x + 0.2, 1.55, 3.65, 0.62, head, 18, WHITE, True, anchor="middle")
        for j, (name, note) in enumerate(rows):
            y = 2.45 + j * 1.1
            text(slide, x + 0.24, y, 3.6, 0.4, name, 20, INK, True)
            text(slide, x + 0.24, y + 0.42, 3.6, 0.35, note, 14, MUTED)
    text(slide, 0.42, 6.3, 12.4, 0.55, "亏损期市盈率、终值异常的DCF，标成交叉检验。不要为了算出目标价而给它权重。", 16, SLATE)


def s12(prs):
    slide = blank(prs)
    chrome(slide, "方法  ·  DCF", 12)
    title(slide, "四段都要自洽，终值不能偷偷主导")
    blocks = [
        ("预测", "收入增长、毛利、费用化研发、销管。\nFCFF = NOPAT + 折旧摊销 − 资本开支 − 营运资金增加。"),
        ("折现", "股权成本用估值币种的无风险利率。\nWACC 用目标资本结构，税后债务成本。"),
        ("终值", "戈登增长要求 WACC > g，且利润、研发、重置资本开支已经稳定。\n还在高速扩张，就把预测期拉到 8–12 年。"),
        ("桥", "经营EV 加超额现金和非经营资产，减债务、租赁和少数股东。\n租赁只在口径一致时扣。"),
    ]
    for i, (head, body) in enumerate(blocks):
        y = 1.5 + i * 1.25
        rect(slide, 0.38, y, 8.55, 1.15, PAPER, LINE, 0.06)
        rect(slide, 0.38, y, 1.35, 1.15, PINE)
        text(slide, 0.38, y, 1.35, 1.15, head, 16, WHITE, True, "center", "middle")
        text(slide, 1.95, y + 0.12, 6.75, 0.92, body, 13, SLATE, spacing=1.08)
    rect(slide, 9.15, 1.5, 3.75, 5.0, PINE, radius=0.05)
    text(slide, 9.4, 1.75, 3.3, 0.4, "五件不要做的事", 16, GOLD, True)
    warns = ["币种和折现率各说各话", "WACC 小于或等于 g", "未稳态就永久外推", "资本化研发再扣一次", "把成熟公司WACC直接用于亏损AI"]
    for i, warn in enumerate(warns):
        text(slide, 9.4, 2.4 + i * 0.72, 3.3, 0.6, warn, 15, WHITE)


def s13(prs):
    slide = blank(prs)
    chrome(slide, "方法  ·  市场与反向", 13)
    title(slide, "倍数和分母必须是同一口径")
    pairs = [("EV / 收入", "看规模，低毛利会被放大"), ("EV / 毛利", "转售和集成收入多时更合适"), ("EV / EBITDA", "利润稳定之后再用"), ("市盈率", "亏损期不要用")]
    for i, (name, note) in enumerate(pairs):
        x = 0.38 + i * 3.22
        rect(slide, x, 1.55, 3.05, 1.85, PAPER, LINE, 0.06)
        text(slide, x + 0.18, 1.72, 2.7, 0.55, name, 18, PINE, True)
        text(slide, x + 0.18, 2.4, 2.7, 0.75, note, 14, SLATE)
    rect(slide, 0.38, 3.65, 8.3, 2.95, PAPER, LINE, 0.05)
    text(slide, 0.6, 3.85, 7.9, 0.4, "反向求解：其他参数固定，只解一个", 18, INK, True)
    for i, item in enumerate(("收入复合增长", "稳态利润率", "WACC", "永续增长 g")):
        x = 0.65 + (i % 2) * 3.9
        y = 4.5 + (i // 2) * 0.8
        rect(slide, x, y, 0.36, 0.36, GOLD)
        text(slide, x, y, 0.36, 0.36, str(i + 1), 14, WHITE, True, "center", "middle")
        text(slide, x + 0.5, y, 3.1, 0.36, item, 16, INK, anchor="middle")
    rect(slide, 8.9, 3.65, 4.0, 2.95, PINE, radius=0.05)
    text(slide, 9.15, 3.9, 3.55, 2.4, "调整倍数\n= 基准 ×（1 + 有证据的差异）\n\n不是一级市场一律七折。", 16, WHITE, spacing=1.15)


def s15(prs):
    slide = blank(prs)
    chrome(slide, "行业  ·  AI", 15)
    title(slide, "AI先分型。不同类型，第一方法不同。")
    types = [
        ("基础模型 / MaaS", "Token 与单价", "单位经济 + 多阶段 DCF"),
        ("企业级 SaaS", "ARR、留存、席位", "分群 DCF，EV / ARR"),
        ("芯片与算力", "出货、ASP、良率", "出货驱动 DCF"),
        ("项目制集成", "订单与确认率", "现金流 DCF，软件单独看"),
        ("尚未商业化", "里程碑与试点", "路径法 + 最近融资"),
    ]
    for i, (name, driver, method) in enumerate(types):
        x = 0.38 + i * 2.56
        rect(slide, x, 1.6, 2.42, 4.9, PAPER, LINE, 0.05)
        rect(slide, x, 1.6, 2.42, 0.1, GOLD)
        text(slide, x + 0.14, 1.9, 2.14, 0.4, f"0{i+1}", 14, GOLD_DK, True)
        text(slide, x + 0.14, 2.4, 2.14, 1.15, name, 20, INK, True)
        text(slide, x + 0.14, 3.7, 2.14, 0.3, "驱动", 12, MUTED, True)
        text(slide, x + 0.14, 4.05, 2.14, 0.7, driver, 14, SLATE)
        text(slide, x + 0.14, 4.85, 2.14, 0.3, "第一方法", 12, MUTED, True)
        text(slide, x + 0.14, 5.2, 2.14, 0.9, method, 14, PINE, True)


def s16(prs):
    slide = blank(prs)
    chrome(slide, "行业  ·  AI单位经济", 16)
    title(slide, "三处最容易把收入或成本算重")
    rect(slide, 0.38, 1.55, 7.9, 3.35, PAPER, LINE, 0.05)
    formulas = [
        "期末老客 ARR = 期初 ARR × NRR",
        "确认收入 ≈ 平均 ARR × 履约因子 + 已计费 Token",
        "Token 毛利率 =（实现单价 − 单位推理成本）/ 实现单价",
    ]
    for i, line in enumerate(formulas):
        text(slide, 0.65, 1.8 + i * 0.95, 7.4, 0.7, line, 18, INK, True, anchor="middle")
    bans = [
        ("NRR 之后再扣流失", "留存里已经包含增购和流失"),
        ("自有 GPU 进成本再进资本开支", "折旧在利润表，现金只在资本开支"),
        ("免费调用算成收入", "赠送、试用和非约束订单都要剔除"),
    ]
    for i, (head, body) in enumerate(bans):
        y = 5.1 + (i % 3) * 0
        x = 0.38 + i * 4.25
        rect(slide, x, 5.1, 4.05, 1.55, RED_SOFT, radius=0.06)
        text(slide, x + 0.18, 5.22, 3.7, 0.55, head, 15, RED, True)
        text(slide, x + 0.18, 5.8, 3.7, 0.65, body, 13, SLATE)


def s17(prs, rows):
    slide = blank(prs)
    chrome(slide, "行业  ·  虚构五年", 17)
    title(slide, "2026年收入 2.14 亿。现金流转正要到 2029 年。", 24)
    headers = ["年度", "增长", "收入", "毛利率", "研发", "销管", "EBITDA", "FCFF"]
    data = []
    for row in rows:
        data.append([
            str(row["year"]),
            f"{row['growth']:.0%}",
            n3(row["rev"]),
            f"{row['gm']:.0%}",
            f"{row['rnd']:.0%}",
            f"{row['sga']:.0%}",
            n3(row["ebitda"]),
            n3(row["fcff"]),
        ])
    widths = [1.25, 1.2, 1.7, 1.35, 1.25, 1.25, 2.15, 2.15]
    x0, y = 0.4, 1.55
    h = 0.52
    xx = x0
    for w, head in zip(widths, headers):
        rect(slide, xx, y, w - 0.05, h, PINE)
        text(slide, xx, y, w - 0.05, h, head, 13, WHITE, True, "center", "middle")
        xx += w
    for r, row in enumerate(data):
        yy = y + h + 0.04 + r * h
        xx = x0
        tone = PAPER if r % 2 == 0 else RGBColor(0xF3, 0xEF, 0xE6)
        for c, (w, value) in enumerate(zip(widths, row)):
            rect(slide, xx, yy, w - 0.05, h - 0.04, tone)
            color = RED if c >= 6 and value.startswith("-") else INK
            text(slide, xx, yy, w - 0.05, h - 0.04, value, 14, color, c == 0, "center", "middle")
            xx += w
    text(slide, 0.42, 4.85, 12.4, 0.7, "单位：亿元人民币。资本开支占收入 8%、8%、7%、6%、5%。折旧摊销为收入的 3%，营运资金为增量收入的 4%。税率 25%，亏损年税额为 0，没有做税损结转。", 13, MUTED)
    kpis = [("WACC", "16%"), ("永续增长", "3%"), ("折现", "2026年末起，年末")]
    for i, (k, v) in enumerate(kpis):
        x = 0.4 + i * 4.2
        rect(slide, x, 5.7, 4.0, 1.05, PINE, radius=0.08)
        text(slide, x + 0.2, 5.82, 3.6, 0.3, k, 12, GOLD)
        text(slide, x + 0.2, 6.12, 3.6, 0.45, v, 18, WHITE, True)


def s18(prs, result, grid):
    slide = blank(prs)
    chrome(slide, "行业  ·  敏感性", 18)
    title(slide, "5.52 亿元里，终值现值已经更高")
    tiles = [
        (n2(result["ev"]), "经营 EV", "亿元人民币"),
        (n2(result["tv_pv"]), "终值现值", "高于全部经营 EV"),
        (n2(result["pv_fcff"]), "五年 FCFF 现值", "前几年仍在消耗现金"),
        (n2(result["tv"]), "2031 终值", "尚未折现"),
    ]
    for i, (num, name, note) in enumerate(tiles):
        x = 0.38 + i * 3.22
        rect(slide, x, 1.5, 3.05, 1.7, PAPER if i else PINE, LINE if i else None, 0.06)
        text(slide, x + 0.16, 1.62, 2.75, 0.65, num, 28, WHITE if i == 0 else PINE, True)
        text(slide, x + 0.16, 2.3, 2.75, 0.32, name, 14, GOLD if i == 0 else INK, True)
        text(slide, x + 0.16, 2.65, 2.75, 0.35, note, 12, RGBColor(0xE7, 0xE1, 0xD4) if i == 0 else MUTED)
    waccs = (0.14, 0.16, 0.18)
    gs = (0.02, 0.03, 0.04)
    full_w = (0.12, 0.14, 0.16, 0.18, 0.20)
    full_g = (0.01, 0.02, 0.03, 0.04, 0.05)
    text(slide, 0.42, 3.4, 4, 0.35, "WACC  \\  终值 g", 14, MUTED, True)
    labels = ["2%", "3%", "4%"]
    for i, lab in enumerate(labels):
        text(slide, 2.3 + i * 1.7, 3.4, 1.5, 0.35, lab, 14, GOLD_DK, True, "center")
    for r, w in enumerate(waccs):
        text(slide, 0.4, 3.9 + r * 0.85, 1.6, 0.7, f"{w:.0%}", 16, INK, True, anchor="middle")
        for c, g in enumerate(gs):
            value = grid[full_w.index(w)][full_g.index(g)]
            x, y = 2.3 + c * 1.7, 3.9 + r * 0.85
            hot = abs(w - 0.16) < 1e-9 and abs(g - 0.03) < 1e-9
            rect(slide, x, y, 1.55, 0.72, PINE if hot else PAPER, GOLD if hot else LINE, 0.08)
            text(slide, x, y, 1.55, 0.72, n2(value), 18, WHITE if hot else INK, True, "center", "middle")
    rect(slide, 7.7, 3.45, 5.2, 3.15, RED_SOFT, radius=0.06)
    text(slide, 7.95, 3.65, 4.75, 2.75, "终值现值 5.83，经营 EV 5.52。\n\n短期现金消耗大，结论挂在远期盈利上。应拉长预测期、检查稳态资本开支，并与可比公司校准。\n\n不能把 5.52 亿元调高，去凑 40 亿或 80 亿港元。", 15, SLATE, spacing=1.05)


def s19(prs, mkt):
    slide = blank(prs)
    chrome(slide, "行业  ·  方法差额", 19)
    title(slide, "同一年收入，两种指示相差 2.70 亿元")
    bars = [
        ("原始可比  3.0x", mkt["raw"], False),
        ("质量调整后  −20%", mkt["adjusted"], False),
        ("DCF 经营 EV", mkt["dcf"], True),
    ]
    max_v = mkt["raw"]
    for i, (name, value, hot) in enumerate(bars):
        y = 1.65 + i * 1.15
        text(slide, 0.42, y, 3.3, 0.85, name, 16, INK, True, anchor="middle")
        width = 6.4 * value / max_v
        rect(slide, 3.8, y + 0.18, width, 0.5, PINE if hot else GOLD)
        text(slide, 3.9 + width, y, 2.2, 0.85, n2(value), 20, PINE, True, anchor="middle")
    rect(slide, 0.38, 5.15, 12.55, 1.7, PAPER, LINE, 0.05)
    text(slide, 0.6, 5.32, 12.1, 0.4, "这 2.70 亿元要解释，不能再贴一个“IPO 折价 20%”", 18, INK, True)
    text(slide, 0.6, 5.85, 12.1, 0.8, "解释来自增长质量、推理单位经济、正常化毛利、折现率，以及上市可比是否同一口径。未商业化企业还要单列：合资格收入的时间表、125% 营运资金、融资后股数、失败情景、客户与算力证据。", 14, SLATE)


def s20(prs):
    slide = blank(prs)
    chrome(slide, "行业  ·  硬科技与新能源", 20)
    title(slide, "先改物理量和时间，再谈倍数")
    rect(slide, 0.38, 1.55, 6.25, 5.2, PAPER, LINE, 0.05)
    text(slide, 0.6, 1.75, 5.8, 0.4, "硬科技", 22, PINE, True)
    text(slide, 0.6, 2.3, 5.8, 1.3, "可售量 = 投片 × 每片裸片 × 良率 × 交付率\n收入 = 可售量 × 订单加权 ASP", 16, INK, spacing=1.15)
    text(slide, 0.6, 3.85, 5.8, 0.35, "回调顺序", 14, GOLD_DK, True)
    for i, line in enumerate(("客户认证和流片是否延期", "单位成本、资本开支、现金跑道", "然后才重算倍数和退出时间")):
        text(slide, 0.6, 4.35 + i * 0.55, 5.8, 0.45, f"{i+1}    {line}", 16, SLATE, anchor="middle")
    rect(slide, 6.85, 1.55, 6.05, 5.2, PAPER, LINE, 0.05)
    text(slide, 7.05, 1.75, 5.6, 0.4, "新能源", 22, PINE, True)
    text(slide, 7.05, 2.3, 5.6, 1.3, "出货 = 产能 × 利用率 × 合格率 × 销售实现率\n长期电站按项目期限，不用永续高增长", 16, INK, spacing=1.15)
    text(slide, 7.05, 3.85, 5.6, 0.35, "价格和成本分开动", 14, GOLD_DK, True)
    for i, line in enumerate(("ASP 与原料各算各的", "再改交付、并网和回款", "利润和现金受到的影响不同")):
        text(slide, 7.05, 4.35 + i * 0.55, 5.6, 0.45, f"{i+1}    {line}", 16, SLATE, anchor="middle")


def s21(prs):
    slide = blank(prs)
    chrome(slide, "行业  ·  生物医药", 21)
    title(slide, "临床失败只惩罚一次")
    rect(slide, 0.38, 1.55, 7.5, 5.15, PAPER, LINE, 0.05)
    text(slide, 0.65, 1.8, 7.0, 0.4, "手册里的算术，不是成功率", 16, MUTED, True)
    text(slide, 0.65, 2.35, 7.0, 0.9, "12 亿  ×  15%  −  1.5 亿", 28, INK, True)
    text(slide, 0.65, 3.3, 7.0, 0.7, "=  0.30 亿元", 32, PINE, True)
    text(slide, 0.65, 4.3, 6.9, 2.0, "12 亿是“全部成功”时商业化现金流的现值。15% 是演示输入。1.5 亿是按阶段概率加权后的研发支出。\n正式模型要拆原理、毒理、各期临床、注册和上市后退出。", 15, SLATE, spacing=1.15)
    rect(slide, 8.1, 1.55, 4.85, 5.15, PINE, radius=0.05)
    text(slide, 8.4, 1.85, 4.3, 0.4, "分流", 16, GOLD, True)
    notes = ["有稳定销售：逐产品 DCF", "在研药物：rNPV", "AI 平台收入和自有管线分开", "18A 的 15 亿不用于 18C", "亏损药企不用市盈率"]
    for i, note in enumerate(notes):
        text(slide, 8.4, 2.5 + i * 0.7, 4.3, 0.55, note, 16, WHITE)


def s22(prs):
    slide = blank(prs)
    chrome(slide, "行业  ·  装备与专精特新", 22)
    title(slide, "签单不是收入。资质也不是溢价。")
    left = [
        ("设备", "出货台数 × 售价 × 验收系数"),
        ("订单", "期初 + 新签 − 确认 − 取消"),
        ("服务", "装机保有 × 渗透率 × 年度 ARPU"),
        ("延期", "同时改回款、质保和营运资金"),
    ]
    right = [
        ("它不是一个行业", "先回到具体技术，再决定是否 18C"),
        ("稳定零部件", "DCF 与 EV/EBITDA"),
        ("单一大客户", "情景，加上清算下限"),
        ("四个陷阱", "资质、专利件数、应收、重复加总"),
    ]
    text(slide, 0.42, 1.5, 6, 0.4, "高端装备", 16, GOLD_DK, True)
    text(slide, 6.9, 1.5, 6, 0.4, "专精特新", 16, GOLD_DK, True)
    for i, (head, body) in enumerate(left):
        y = 2.05 + i * 1.15
        rect(slide, 0.38, y, 6.15, 1.02, PAPER, LINE, 0.06)
        text(slide, 0.58, y + 0.1, 5.8, 0.35, head, 16, PINE, True)
        text(slide, 0.58, y + 0.48, 5.8, 0.4, body, 14, SLATE)
    for i, (head, body) in enumerate(right):
        y = 2.05 + i * 1.15
        rect(slide, 6.75, y, 6.15, 1.02, PAPER, LINE, 0.06)
        text(slide, 6.95, y + 0.1, 5.8, 0.35, head, 16, PINE, True)
        text(slide, 6.95, y + 0.48, 5.8, 0.4, body, 14, SLATE)


def s24(prs):
    slide = blank(prs)
    chrome(slide, "回调", 24)
    title(slide, "四种回调，不要合成一个折价")
    cards = [
        ("01", "经济估值再校准", "新订单、新现金流、新融资或新的可比。这是本手册的重点。"),
        ("02", "交易价格重谈", "融资价格和认购数量重谈。股权和摊薄要重算。"),
        ("03", "合同里的调整", "加权平均反稀释、对赌、清算优先。由合同和法律意见确定。"),
        ("04", "IPO 簿记", "发行价和发行数量来自需求和门槛。目标市值不是 DCF 的约束。"),
    ]
    for i, (num, head, body) in enumerate(cards):
        col, row = i % 2, i // 2
        x, y = 0.38 + col * 6.45, 1.55 + row * 2.5
        rect(slide, x, y, 6.25, 2.3, PAPER, LINE, 0.05)
        text(slide, x + 0.28, y + 0.25, 1.2, 0.45, num, 20, GOLD_DK, True)
        text(slide, x + 1.4, y + 0.28, 4.5, 0.45, head, 22, INK, True)
        text(slide, x + 0.28, y + 1.05, 5.7, 0.9, body, 16, SLATE)


def s25(prs):
    slide = blank(prs)
    chrome(slide, "回调  ·  顺序", 25)
    title(slide, "事实先进入模型，倍数最后动")
    steps = ["锁定上期", "核对实际", "归因驱动", "看资金跑道", "重抽可比", "改WACC和g", "重算检验", "拆变动桥", "签核"]
    for i, name in enumerate(steps):
        x = 0.38 + i * 1.42
        rect(slide, x, 1.55, 1.32, 1.55, PINE if i in (0, 8) else PAPER, LINE if i not in (0, 8) else None, 0.08)
        text(slide, x, 1.68, 1.32, 0.35, f"{i:02d}", 12, GOLD if i in (0, 8) else GOLD_DK, True, "center")
        text(slide, x + 0.06, 2.1, 1.2, 0.8, name, 13, WHITE if i in (0, 8) else INK, True, "center")
    triggers = [
        ("收费率下降", "先改单价和免费量", "不要直接提高 PS"),
        ("客户流失", "先改 NRR 和合同额", "不要只上调 WACC"),
        ("流片延期", "先改时间和概率", "不要保留原增长率"),
        ("临床失利", "先改该阶段路径", "不要再加一次溢价"),
        ("离门槛很远", "改结论和路径", "不要用门槛倒推终值"),
    ]
    for i, (fact, first, avoid) in enumerate(triggers):
        y = 3.4 + i * 0.68
        rect(slide, 0.38, y, 2.5, 0.58, CREAM, LINE, 0.08)
        text(slide, 0.5, y, 2.3, 0.58, fact, 14, INK, True, anchor="middle")
        text(slide, 3.1, y, 4.6, 0.58, first, 15, PINE, True, anchor="middle")
        text(slide, 7.8, y, 5.0, 0.58, avoid, 15, RED, anchor="middle")


def s26(prs, bridge):
    slide = blank(prs)
    chrome(slide, "回调  ·  变动桥", 26)
    title(slide, "8.00 到 5.90，每一档都要能单独重算")
    steps = [
        ("上期", bridge["old"], "abs"),
        ("经营下修", bridge["ops"], "neg"),
        ("算力毛利", bridge["margin"], "neg"),
        ("倍数下降", bridge["multiple"], "neg"),
        ("净现金", bridge["cash"], "pos"),
        ("本期", bridge["new"], "abs"),
    ]
    running = 0
    levels = []
    for name, delta, kind in steps:
        if kind == "abs":
            levels.append((name, 0, delta, kind))
            running = delta
        else:
            levels.append((name, running, running + delta, kind))
            running += delta
    max_v = 8.0
    top, bottom = 2.05, 5.55
    height = bottom - top
    left = 0.55
    slot = 12.2 / 6
    edges = []
    for i, (name, a, b, kind) in enumerate(levels):
        x = left + i * slot + 0.25
        y_hi = top + (1 - max(a, b) / max_v) * height
        bar_h = abs(b - a) / max_v * height
        color = PINE if kind == "abs" else (TEAL if kind == "pos" else RED)
        rect(slide, x, y_hi, 1.35, max(bar_h, 0.08), color, radius=0.04)
        label = n2(b if kind == "abs" else (b - a))
        text(slide, x - 0.1, max(y_hi - 0.36, 1.48), 1.55, 0.32, label, 14, color, True, "center")
        text(slide, x - 0.15, 5.7, 1.7, 0.4, name, 13, INK, True, "center")
        edges.append((x, x + 1.35, b))
    for i in range(len(edges) - 1):
        _, x_right, level = edges[i]
        x_next, _, _ = edges[i + 1]
        y = top + (1 - level / max_v) * height
        rect(slide, x_right, y, max(x_next - x_right, 0.02), 0.012, LINE)
    rect(slide, 0.4, 6.25, 12.5, 0.7, TEAL_SOFT, radius=0.06)
    text(slide, 0.6, 6.25, 12.1, 0.7, "残差 = 0。这是结构演示：冻结其他因素，一次只替换一项。它不是任何真实公司的减项。", 15, TEAL, True, anchor="middle")


def s27(prs):
    slide = blank(prs)
    chrome(slide, "股权  ·  稀释与 IPO", 27)
    title(slide, "价格从 8 亿谈到 6 亿，新股东多拿 5 个百分点")
    rect(slide, 0.38, 1.55, 6.15, 2.55, PAPER, LINE, 0.05)
    text(slide, 0.6, 1.7, 5.7, 0.35, "情景一", 14, GOLD_DK, True)
    text(slide, 0.6, 2.1, 5.7, 0.7, "8 亿  +  2 亿  =  10 亿", 26, INK, True)
    text(slide, 0.6, 2.9, 5.7, 0.8, "新投资者 20%　　原股东 80%", 18, PINE, True)
    rect(slide, 6.75, 1.55, 6.15, 2.55, PINE, radius=0.05)
    text(slide, 6.98, 1.7, 5.7, 0.35, "情景二  ·  同一笔 2 亿", 14, GOLD, True)
    text(slide, 6.98, 2.1, 5.7, 0.7, "6 亿  +  2 亿  =  8 亿", 26, WHITE, True)
    text(slide, 6.98, 2.9, 5.7, 0.8, "新投资者 25%　　原股东 75%", 18, RGBColor(0xF3, 0xEB, 0xD8), True)
    points = [
        ("发行市值", "发行价 × 发行后等价股数"),
        ("不要加两次", "已含现金，不要再加"),
        ("复杂条款", "先做瀑布，再谈普通股"),
        ("折让", "三项折让不一定叠加"),
    ]
    for i, (head, body) in enumerate(points):
        x = 0.38 + (i % 4) * 3.22
        y = 4.35
        rect(slide, x, y, 3.05, 2.3, PAPER, LINE, 0.06)
        text(slide, x + 0.18, y + 0.25, 2.7, 0.55, head, 18, PINE, True)
        text(slide, x + 0.18, y + 1.0, 2.7, 1.0, body, 14, SLATE)


def s28(prs):
    slide = blank(prs)
    chrome(slide, "模型", 28)
    title(slide, "工作簿用颜色管住谁可以改")
    colors = [
        (RGBColor(0xFF, 0xF6, 0xE4), BLUE := RGBColor(0x1F, 0x4E, 0x79), "蓝字", "输入。来源、日期、单位、责任人。"),
        (WHITE, INK, "黑字", "公式。要改就回到蓝字，不要覆盖。"),
        (TEAL_SOFT, TEAL, "绿色", "结果。EV、灯号、勾稽。"),
        (RGBColor(0xFB, 0xF0, 0xE4), GOLD_DK, "橙色", "提示。终值过高、演示汇率、税损未建。"),
        (RED_SOFT, RED, "红色", "异常。WACC≤g、桥不平、币种错。"),
    ]
    for i, (fill, color, name, body) in enumerate(colors):
        y = 1.5 + i * 0.95
        rect(slide, 0.38, y, 7.6, 0.85, fill, LINE, 0.08)
        text(slide, 0.6, y, 1.5, 0.85, name, 18, color, True, anchor="middle")
        text(slide, 2.2, y, 5.5, 0.85, body, 16, INK, anchor="middle")
    rect(slide, 8.2, 1.5, 4.7, 4.75, PINE, radius=0.05)
    text(slide, 8.45, 1.75, 4.25, 0.4, "打开 Excel 先看", 16, GOLD, True)
    lines = ["00  四条计算链不要加总", "02  本项目列默认不可评估", "03  改 AI 假设", "05  经营 EV 与终值占比", "06  敏感性", "17  异常数应为 0"]
    for i, line in enumerate(lines):
        text(slide, 8.45, 2.4 + i * 0.55, 4.2, 0.45, line, 15, WHITE)


def s29(prs):
    slide = blank(prs)
    chrome(slide, "落地", 29)
    title(slide, "资源有限时，按这个顺序做")
    order = [
        ("01", "18C 合规表", "收入、研发、投资者、现金"),
        ("02", "AI 单位毛利", "ARR、Token、推理成本"),
        ("03", "DCF 与跑道", "月度现金，而不只是终值"),
        ("04", "可比与融资", "用当时的价格做校准"),
        ("05", "股权与变动桥", "发行后股数，残差为零"),
        ("06", "其余五个行业", "只改驱动和回调顺序"),
    ]
    for i, (num, name, body) in enumerate(order):
        col, row = i % 3, i // 3
        x, y = 0.38 + col * 4.25, 1.55 + row * 2.15
        rect(slide, x, y, 4.05, 1.95, PAPER, LINE, 0.06)
        text(slide, x + 0.22, y + 0.2, 1.2, 0.4, num, 16, GOLD_DK, True)
        text(slide, x + 0.22, y + 0.7, 3.6, 0.45, name, 20, INK, True)
        text(slide, x + 0.22, y + 1.25, 3.6, 0.45, body, 14, SLATE)
    text(slide, 0.42, 6.05, 12.4, 0.85, "公开材料分两类：四大的方法指南，以及晶泰、越疆、MiniMax、壁仞和 Skadden 的发行或市场资料。前者不是客户底稿，后者不是估值报告。链接在 Word 和 Excel。", 14, MUTED)


def s30(prs):
    slide = blank(prs)
    bg(slide, PINE)
    rect(slide, 0, 0, 0.12, SH, GOLD)
    text(slide, 0.75, 1.45, 11.5, 0.4, "交给投委会的那一句", 16, GOLD, True)
    text(slide, 0.75, 2.05, 11.8, 2.3, "核心不是一个热门科技倍数。", 36, WHITE, True)
    text(slide, 0.75, 3.55, 11.5, 1.5, "把特专科技收入、研发持续性、推理单位经济、\n资金需求、股权优先权和可接受的定价，放进同一条可审阅的计算链。", 20, RGBColor(0xF3, 0xEB, 0xD8), spacing=1.25)
    files = [("PPT", "这本简报"), ("Word", "完整手册"), ("Excel", "可改的公式")]
    for i, (kind, name) in enumerate(files):
        x = 0.75 + i * 3.8
        rect(slide, x, 5.55, 3.45, 1.05, PINE_MID, radius=0.08)
        text(slide, x + 0.25, 5.68, 3.0, 0.35, kind, 12, GOLD, True)
        text(slide, x + 0.25, 6.02, 3.0, 0.4, name, 18, WHITE, True)


def build():
    mc.verify()
    rows = mc.forecast()
    result = mc.dcf(rows)
    grid_w, grid_g, grid = mc.sensitivity()
    mkt = mc.market_method(rows)
    bridge = mc.equity_bridge()
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    prs.core_properties.title = "港交所18C分行业估值模板"
    prs.core_properties.subject = "投委会简报。虚构演示，非投资建议。"
    prs.core_properties.category = "估值简报"
    s01(prs)
    s02(prs)
    s03(prs)
    s04(prs)
    divider(prs, 5, "01", "规则准入", "收入、市值、研发、投资者和现金，要同时说得通。")
    s06(prs)
    s07(prs)
    s08(prs)
    s09(prs)
    divider(prs, 10, "02", "方法", "口径一致、已经校准、确实有信息，才允许加权。")
    s11(prs)
    s12(prs)
    s13(prs)
    divider(prs, 14, "03", "六个行业", "AI做透。其余行业改的是驱动和回调顺序，不是另一套折扣。")
    s15(prs)
    s16(prs)
    s17(prs, rows)
    s18(prs, result, grid)
    s19(prs, mkt)
    s20(prs)
    s21(prs)
    s22(prs)
    divider(prs, 23, "04", "回调与交付", "经济再校准、价格重谈、合同条款和簿记，分开处理。")
    s24(prs)
    s25(prs)
    s26(prs, bridge)
    s27(prs)
    s28(prs)
    s29(prs)
    s30(prs)
    if len(prs.slides) != TOTAL:
        raise RuntimeError(f"页数 {len(prs.slides)} 与标注 {TOTAL} 不一致")
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUTPUT)
    return OUTPUT


if __name__ == "__main__":
    print(build())
