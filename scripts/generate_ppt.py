# -*- coding: utf-8 -*-
"""生成『群邦·领事会客厅』三方战略合作策划案 PPT。

用法：
    python3 scripts/generate_ppt.py [输出路径.pptx]
默认输出到 output/群邦-领事会客厅-三方战略合作策划案.pptx
"""
import sys
import os

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

import content as C

# ---- 主题 ----
NAVY = RGBColor(0x0B, 0x2A, 0x4A)      # 主色 深蓝
NAVY2 = RGBColor(0x12, 0x3A, 0x63)
GOLD = RGBColor(0xC9, 0xA2, 0x4B)      # 强调 金
LIGHT = RGBColor(0xF4, 0xF6, 0xF9)     # 浅底
GREY = RGBColor(0x5B, 0x6B, 0x7B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1E, 0x29, 0x33)

FONT = "Microsoft YaHei"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def set_font(run, size, color=DARK, bold=False, font=FONT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    # 中文字体
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', font)


def add_rect(slide, x, y, w, h, fill, line=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.75)
    shp.shadow.inherit = False
    return shp


def add_text(slide, x, y, w, h, text, size, color=DARK, bold=False,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font=FONT):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    set_font(r, size, color, bold, font)
    return tb


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def header(slide, title, idx=None, total=None):
    """标准内容页页眉。"""
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.1), NAVY)
    add_rect(slide, 0, Inches(1.1), SLIDE_W, Pt(3), GOLD)
    add_text(slide, Inches(0.6), Inches(0.12), Inches(11.0), Inches(0.85),
             title, 26, WHITE, True, anchor=MSO_ANCHOR.MIDDLE)
    if idx is not None:
        add_text(slide, Inches(11.6), Inches(0.12), Inches(1.3), Inches(0.85),
                 f"{idx:02d}/{total:02d}", 12, GOLD, True,
                 align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def footer(slide):
    add_text(slide, Inches(0.6), Inches(7.05), Inches(8), Inches(0.4),
             f"{C.PROJECT_NAME} · {C.PROJECT_TAG}", 9, GREY)


def bullets(slide, x, y, w, h, items, size=16, gap=10, color=DARK,
            marker="●", marker_color=GOLD, sub=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for it in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(gap)
        p.line_spacing = 1.12
        rm = p.add_run()
        rm.text = marker + "  "
        set_font(rm, size - 2, marker_color, True)
        r = p.add_run()
        r.text = it
        set_font(r, size, color)
    return tb


def add_table_slide(slide, tdata, idx, total):
    header(slide, tdata["title"], idx, total)
    headers = tdata["headers"]
    rows = tdata["rows"]
    ncol = len(headers)
    nrow = len(rows) + 1
    left = Inches(0.5)
    top = Inches(1.45)
    width = Inches(12.33)
    height = Inches(0.45) * nrow
    gtbl = slide.shapes.add_table(nrow, ncol, left, top, width, height).table
    # 列宽：首列略宽
    for ci in range(ncol):
        gtbl.columns[ci].width = Emu(int(width / ncol))
    # 表头
    for ci, htext in enumerate(headers):
        cell = gtbl.cell(0, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = Inches(0.06)
        cell.margin_right = Inches(0.06)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = htext
        set_font(r, 12, WHITE, True)
    # 数据
    for ri, row in enumerate(rows, start=1):
        is_total = (row[0] == "合计")
        for ci, val in enumerate(row):
            cell = gtbl.cell(ri, ci)
            cell.fill.solid()
            if is_total:
                cell.fill.fore_color.rgb = GOLD
            else:
                cell.fill.fore_color.rgb = WHITE if ri % 2 else LIGHT
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.06)
            cell.margin_right = Inches(0.06)
            cell.margin_top = Inches(0.02)
            cell.margin_bottom = Inches(0.02)
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if ci else PP_ALIGN.LEFT
            r = p.add_run(); r.text = str(val)
            set_font(r, 11, DARK, is_total or ci == 0)
    note = tdata.get("note")
    if note:
        ny = top + height + Inches(0.15)
        add_text(slide, Inches(0.5), ny, Inches(12.33), Inches(0.8),
                 "说明：" + note, 11, GREY)
    footer(slide)


# ===========================================================================
# 各页构建
# ===========================================================================

def slide_cover(prs):
    s = blank(prs)
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(s, 0, Inches(4.9), SLIDE_W, Pt(3), GOLD)
    # 顶部小标
    add_text(s, Inches(0.9), Inches(1.2), Inches(11.5), Inches(0.5),
             "复旦政策研究中心 × 杨浦科技企业联合会 × 郡邦·总领事俱乐部(CGC)",
             15, GOLD, True)
    add_text(s, Inches(0.9), Inches(2.4), Inches(11.5), Inches(1.4),
             C.PROJECT_NAME, 56, WHITE, True)
    add_text(s, Inches(0.92), Inches(3.7), Inches(11.5), Inches(0.8),
             C.PROJECT_SUBTITLE, 22, LIGHT)
    add_text(s, Inches(0.92), Inches(5.15), Inches(11.5), Inches(0.6),
             C.PROJECT_TAG, 18, GOLD, True)
    add_text(s, Inches(0.92), Inches(6.6), Inches(11.5), Inches(0.5),
             "地标空间：上海·金茂大厦86楼　|　" + C.VERSION, 12, GREY)
    return s


def slide_section(prs, no, title, subtitle=""):
    s = blank(prs)
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(s, Inches(0.9), Inches(2.6), Inches(1.6), Pt(4), GOLD)
    add_text(s, Inches(0.9), Inches(2.9), Inches(3), Inches(2),
             no, 96, GOLD, True)
    add_text(s, Inches(3.4), Inches(3.05), Inches(9), Inches(1.2),
             title, 40, WHITE, True, anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_text(s, Inches(3.45), Inches(4.25), Inches(9), Inches(1),
                 subtitle, 18, LIGHT)
    return s


def slide_background(prs, idx, total):
    s = blank(prs)
    header(s, "一、背景与机遇：稀缺资源的叠加", idx, total)
    add_text(s, Inches(0.6), Inches(1.4), Inches(6), Inches(0.5),
             "资源现状", 18, NAVY, True)
    bullets(s, Inches(0.6), Inches(1.95), Inches(6.0), Inches(4.5),
            C.BACKGROUND, size=15, gap=12)
    add_rect(s, Inches(6.95), Inches(1.45), Inches(5.8), Inches(5.0), LIGHT)
    add_text(s, Inches(7.2), Inches(1.6), Inches(5.4), Inches(0.5),
             "机遇与思路", 18, NAVY, True)
    bullets(s, Inches(7.2), Inches(2.2), Inches(5.4), Inches(4.0),
            C.OPPORTUNITY, size=15, gap=12)
    footer(s)
    return s


def slide_cgc_profile(prs, idx, total):
    s = blank(prs)
    header(s, "一、合作对象：总领事俱乐部（CGC）", idx, total)
    bullets(s, Inches(0.7), Inches(1.55), Inches(11.9), Inches(3.0),
            C.CGC_PROFILE, size=15.5, gap=12)
    add_rect(s, Inches(0.5), Inches(4.65), Inches(12.33), Inches(0.6), GOLD)
    add_text(s, Inches(0.7), Inches(4.65), Inches(12), Inches(0.6),
             "宗旨：促进成员交流合作 · 推动各国与领区双边关系、经贸与旅游 · 发展国际文化与友好关系",
             13, NAVY, True, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(0.7), Inches(5.45), Inches(12), Inches(1.0),
             "运营承办：郡邦（上海）文化交流发展有限公司　|　公众号：gh_79ebe44ac3d2（CGC）",
             12, GREY)
    footer(s)
    return s


def slide_assets(prs, idx, total):
    s = blank(prs)
    header(s, "二、CGC 既有品牌资产：在已有基础上升级变现", idx, total)
    add_text(s, Inches(0.6), Inches(1.35), Inches(12), Inches(0.5),
             "不必从零启动——CGC 已沉淀一批可直接升级为收入的品牌活动：",
             14, NAVY, True)
    bullets(s, Inches(0.7), Inches(2.0), Inches(11.9), Inches(4.2),
            C.EXISTING_ASSETS, size=15.5, gap=16)
    footer(s)
    return s


def slide_region(prs, idx, total):
    s = blank(prs)
    header(s, "三、长三角领区：把领事资源变成招商抓手", idx, total)
    add_rect(s, Inches(0.5), Inches(1.4), Inches(12.33), Inches(0.85), NAVY)
    add_text(s, Inches(0.7), Inches(1.4), Inches(12), Inches(0.85),
             C.REGION["intro"], 14, WHITE, True, anchor=MSO_ANCHOR.MIDDLE)
    # 四省市
    cities = ["上海", "江苏", "浙江", "安徽"]
    bw = Inches(2.85)
    gap = Inches(0.33)
    startx = Inches(0.6)
    top = Inches(2.55)
    for i, c in enumerate(cities):
        x = startx + i * (bw + gap)
        add_rect(s, x, top, bw, Inches(0.95), GOLD)
        add_text(s, x, top, bw, Inches(0.95), c, 26, NAVY, True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(0.6), Inches(3.75), Inches(12), Inches(0.5),
             "四大应用方向", 16, NAVY, True)
    bullets(s, Inches(0.7), Inches(4.25), Inches(11.9), Inches(2.4),
            C.REGION["apps"], size=15, gap=12)
    footer(s)
    return s


def slide_parties(prs, idx, total):
    s = blank(prs)
    header(s, "二、合作三方与定位", idx, total)
    cardw = Inches(4.0)
    gap = Inches(0.27)
    startx = Inches(0.5)
    top = Inches(1.5)
    h = Inches(5.0)
    for i, p in enumerate(C.PARTIES):
        x = startx + i * (cardw + gap)
        add_rect(s, x, top, cardw, h, WHITE, line=GOLD)
        add_rect(s, x, top, cardw, Inches(1.35), NAVY)
        add_text(s, x + Inches(0.15), top + Inches(0.12), cardw - Inches(0.3),
                 Inches(1.1), p["name"], 15, WHITE, True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, x + Inches(0.5), top + Inches(1.5), cardw - Inches(1.0),
                 Inches(0.55), GOLD)
        add_text(s, x + Inches(0.5), top + Inches(1.5), cardw - Inches(1.0),
                 Inches(0.55), p["role"], 14, NAVY, True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        bullets(s, x + Inches(0.25), top + Inches(2.3), cardw - Inches(0.5),
                Inches(2.5), p["duties"], size=12, gap=8, marker="·")
    footer(s)
    return s


def slide_loop(prs, idx, total):
    s = blank(prs)
    header(s, "三、价值闭环：活动—会籍—出海—招商", idx, total)
    stages = [
        ("活动获客", "小/中/大型领事活动\n建立深度链接", NAVY),
        ("会籍沉淀", "个人/企业/园区会籍\n形成稳定预收", NAVY2),
        ("出海·招商", "企业出海对接\n国家会客厅招商", GOLD),
        ("政企复购", "长期合作复购\n城市级平台", NAVY),
    ]
    bw = Inches(2.7)
    bh = Inches(2.3)
    top = Inches(2.6)
    startx = Inches(0.55)
    gap = Inches(0.5)
    for i, (t, d, col) in enumerate(stages):
        x = startx + i * (bw + gap)
        add_rect(s, x, top, bw, bh, col)
        add_text(s, x, top + Inches(0.25), bw, Inches(0.7), t, 22,
                 WHITE if col != GOLD else NAVY, True, align=PP_ALIGN.CENTER)
        add_text(s, x + Inches(0.2), top + Inches(1.05), bw - Inches(0.4),
                 Inches(1.1), d, 13,
                 LIGHT if col != GOLD else NAVY, align=PP_ALIGN.CENTER)
        if i < 3:
            add_text(s, x + bw - Inches(0.05), top + Inches(0.75),
                     gap + Inches(0.1), Inches(0.8), "➜", 28, GOLD, True,
                     align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.55), Inches(1.5), Inches(12), Inches(0.9),
             "以领事为纽带，三方联合，把稀缺连接沉淀为可复制、可收费的产品与现金流闭环。",
             17, NAVY, True)
    add_text(s, Inches(0.55), Inches(5.4), Inches(12), Inches(1.2),
             "‘小道大道’：先以小型高频活动建立链接（小道），再放大为城市级开放平台与变现体系（大道）。",
             14, GREY)
    footer(s)
    return s


def slide_strategy(prs, idx, total):
    s = blank(prs)
    header(s, "四、总体战略：以领事为纽带的引领模式", idx, total)
    top = Inches(1.6)
    rowh = Inches(0.96)
    for i, (k, v) in enumerate(C.STRATEGY_PILLARS):
        y = top + i * (rowh + Inches(0.12))
        add_rect(s, Inches(0.6), y, Inches(2.2), rowh, NAVY)
        add_text(s, Inches(0.6), y, Inches(2.2), rowh, k, 22, GOLD, True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, Inches(2.95), y, Inches(9.8), rowh, LIGHT)
        add_text(s, Inches(3.2), y, Inches(9.4), rowh, v, 15, DARK,
                 anchor=MSO_ANCHOR.MIDDLE)
    footer(s)
    return s


def slide_products(prs, idx, total):
    s = blank(prs)
    header(s, "五、产品矩阵：六大产品线", idx, total)
    cardw = Inches(4.05)
    cardh = Inches(2.55)
    gapx = Inches(0.13)
    gapy = Inches(0.2)
    startx = Inches(0.5)
    top = Inches(1.45)
    for i, pl in enumerate(C.PRODUCT_LINES):
        r, c = divmod(i, 3)
        x = startx + c * (cardw + gapx)
        y = top + r * (cardh + gapy)
        add_rect(s, x, y, cardw, cardh, WHITE, line=GOLD)
        add_rect(s, x, y, Inches(0.12), cardh, GOLD)
        add_text(s, x + Inches(0.25), y + Inches(0.1), Inches(1.1), Inches(0.5),
                 pl["no"], 20, GOLD, True)
        add_text(s, x + Inches(1.0), y + Inches(0.12), cardw - Inches(1.1),
                 Inches(0.5), pl["name"], 15, NAVY, True,
                 anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + Inches(0.25), y + Inches(0.62), cardw - Inches(0.4),
                 Inches(0.35), pl["tagline"], 11, GOLD, True)
        bullets(s, x + Inches(0.25), y + Inches(1.0), cardw - Inches(0.45),
                Inches(1.5), pl["points"], size=10.5, gap=4, marker="·")
    footer(s)
    return s


def slide_salon(prs, idx, total):
    s = blank(prs)
    header(s, "六、旗舰产品：国家会客厅", idx, total)
    add_rect(s, Inches(0.5), Inches(1.4), Inches(12.33), Inches(0.95), LIGHT)
    add_text(s, Inches(0.7), Inches(1.5), Inches(12), Inches(0.8),
             "定位：" + C.SALON_MODEL["concept"], 13, NAVY,
             anchor=MSO_ANCHOR.MIDDLE)
    cols = [
        ("四大功能", C.SALON_MODEL["functions"], NAVY),
        ("收入来源", C.SALON_MODEL["revenue"], GOLD),
        ("政府/园区价值", C.SALON_MODEL["gov_value"], NAVY2),
    ]
    cw = Inches(4.0)
    gap = Inches(0.16)
    startx = Inches(0.5)
    top = Inches(2.65)
    for i, (t, items, col) in enumerate(cols):
        x = startx + i * (cw + gap)
        add_rect(s, x, top, cw, Inches(0.6), col)
        add_text(s, x, top, cw, Inches(0.6), t, 16,
                 NAVY if col == GOLD else WHITE, True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        bullets(s, x + Inches(0.2), top + Inches(0.75), cw - Inches(0.35),
                Inches(3.4), items, size=12.5, gap=8, marker="·")
    footer(s)
    return s


def slide_outbound(prs, idx, total):
    s = blank(prs)
    header(s, "七、高端商务出海 + 人文特色", idx, total)
    add_rect(s, Inches(0.5), Inches(1.45), Inches(12.33), Inches(0.9), NAVY)
    add_text(s, Inches(0.7), Inches(1.45), Inches(12), Inches(0.9),
             "节奏：" + C.OUTBOUND["rhythm"], 16, WHITE, True,
             anchor=MSO_ANCHOR.MIDDLE)
    bullets(s, Inches(0.7), Inches(2.8), Inches(11.8), Inches(3.0),
            C.OUTBOUND["highlights"], size=17, gap=18)
    add_rect(s, Inches(0.5), Inches(5.9), Inches(12.33), Inches(0.85), LIGHT)
    add_text(s, Inches(0.7), Inches(5.9), Inches(12), Inches(0.85),
             "差异化亮点：与当地高僧大德面对面的人生智慧交流 —— 商务之外的精神价值，高复购、强口碑。",
             14, NAVY, True, anchor=MSO_ANCHOR.MIDDLE)
    footer(s)
    return s


def slide_compliance(prs, idx, total):
    s = blank(prs)
    header(s, "合规与风险管理", idx, total)
    bullets(s, Inches(0.7), Inches(1.7), Inches(12), Inches(4.8),
            C.COMPLIANCE, size=17, gap=18)
    footer(s)
    return s


def slide_next(prs, idx, total):
    s = blank(prs)
    header(s, "下一步行动建议", idx, total)
    top = Inches(1.6)
    rowh = Inches(0.92)
    for i, step in enumerate(C.NEXT_STEPS):
        y = top + i * (rowh + Inches(0.1))
        add_rect(s, Inches(0.6), y, Inches(0.85), rowh, GOLD)
        add_text(s, Inches(0.6), y, Inches(0.85), rowh, str(i + 1), 24, NAVY,
                 True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, Inches(1.6), y, Inches(11.1), rowh, LIGHT)
        add_text(s, Inches(1.85), y, Inches(10.7), rowh, step, 15, DARK,
                 anchor=MSO_ANCHOR.MIDDLE)
    footer(s)
    return s


def slide_end(prs):
    s = blank(prs)
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(s, Inches(0.9), Inches(3.5), Inches(2.0), Pt(4), GOLD)
    add_text(s, Inches(0.9), Inches(2.4), Inches(11.5), Inches(1.0),
             "以领事为桥，连接世界", 40, WHITE, True)
    add_text(s, Inches(0.92), Inches(3.8), Inches(11.5), Inches(0.8),
             "期待三方携手，把上海的开放资源转化为可持续的价值与现金流。", 18, LIGHT)
    add_text(s, Inches(0.92), Inches(6.4), Inches(11.5), Inches(0.6),
             C.PROJECT_NAME + " · " + C.PROJECT_TAG + " · " + C.VERSION,
             13, GOLD)
    return s


def build(path):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    # 内容页编号
    total = 20
    n = [0]

    def nxt():
        n[0] += 1
        return n[0]

    slide_cover(prs)
    slide_section(prs, "01", "背景与机遇",
                  "CGC 总领事俱乐部 · 长三角领区 · 稀缺资源叠加")
    slide_cgc_profile(prs, nxt(), total)
    slide_background(prs, nxt(), total)
    slide_assets(prs, nxt(), total)
    slide_parties(prs, nxt(), total)
    slide_region(prs, nxt(), total)
    slide_loop(prs, nxt(), total)
    slide_strategy(prs, nxt(), total)
    slide_section(prs, "02", "产品与变现",
                  "六大产品线 · 国家会客厅 · 出海与人文特色")
    slide_products(prs, nxt(), total)
    slide_salon(prs, nxt(), total)
    slide_outbound(prs, nxt(), total)
    # 表格页
    add_table_slide(blank(prs), C.TABLE_ASSETS, nxt(), total)
    add_table_slide(blank(prs), C.TABLE_ACTIVITIES, nxt(), total)
    add_table_slide(blank(prs), C.TABLE_SALON, nxt(), total)
    add_table_slide(blank(prs), C.TABLE_OUTBOUND, nxt(), total)
    add_table_slide(blank(prs), C.TABLE_MEMBERSHIP, nxt(), total)
    slide_section(prs, "03", "机制与落地", "分润机制 · 路线图 · 收入测算 · 合规")
    add_table_slide(blank(prs), C.TABLE_REVENUE_SHARE, nxt(), total)
    add_table_slide(blank(prs), C.TABLE_ROADMAP, nxt(), total)
    add_table_slide(blank(prs), C.TABLE_PROJECTION, nxt(), total)
    slide_compliance(prs, nxt(), total)
    slide_next(prs, nxt(), total)
    slide_end(prs)

    os.makedirs(os.path.dirname(path), exist_ok=True)
    prs.save(path)
    print(f"已生成 PPT：{path}（共 {len(prs.slides)} 页）")


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else \
        "output/群邦-领事会客厅-三方战略合作策划案.pptx"
    build(out)
