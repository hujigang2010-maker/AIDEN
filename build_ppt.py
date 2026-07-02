# -*- coding: utf-8 -*-
"""生成《上海 / 长三角产业研学考察计划与联合主办合作框架》PPT（对等 · 总会对总会）。"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

import content as C

# ---------------------------------------------------------------- 主题配色
NAVY = RGBColor(0x14, 0x2B, 0x45)
BLUE = RGBColor(0x1F, 0x4E, 0x79)
STEEL = RGBColor(0x2E, 0x6D, 0xA4)
AMBER = RGBColor(0xE8, 0x8A, 0x1A)
GREEN = RGBColor(0x2E, 0x7D, 0x53)
RED = RGBColor(0xB0, 0x3A, 0x2E)
LIGHT = RGBColor(0xEE, 0xF3, 0xF8)
GREY = RGBColor(0x5A, 0x63, 0x6E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x22, 0x2A, 0x33)

FONT = "微软雅黑"
SW = Inches(13.333)
SH = Inches(7.5)

prs = Presentation()
prs.slide_width = SW
prs.slide_height = SH
BLANK = prs.slide_layouts[6]


def set_cjk(run, font=FONT):
    run.font.name = font
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", font)


def add_slide():
    return prs.slides.add_slide(BLANK)


def rect(slide, x, y, w, h, fill=None, line=None):
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(1)
    return sp


def textbox(slide, x, y, w, h, lines, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Pt(4); tf.margin_right = Pt(4)
    tf.margin_top = Pt(2); tf.margin_bottom = Pt(2)
    for i, item in enumerate(lines):
        text, size, color, bold = item[0], item[1], item[2], item[3]
        sa = item[4] if len(item) > 4 else 4
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(sa)
        p.line_spacing = 1.08
        r = p.add_run(); r.text = text
        r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color
        set_cjk(r)
    return tb


def header(slide, tag, title, subtitle=None):
    rect(slide, 0, 0, SW, Inches(1.15), fill=NAVY)
    rect(slide, 0, Inches(1.15), SW, Pt(3), fill=AMBER)
    textbox(slide, Inches(0.5), Inches(0.12), Inches(6), Inches(0.4), [(tag, 13, AMBER, True)])
    ln = [(title, 25, WHITE, True)]
    if subtitle:
        ln.append((subtitle, 13, RGBColor(0xC9, 0xD6, 0xE3), False))
    textbox(slide, Inches(0.5), Inches(0.42), Inches(12.3), Inches(0.72), ln)


def footer(slide, page):
    textbox(slide, Inches(0.5), Inches(7.05), Inches(9), Inches(0.35),
            [("钢铁 × 不动产 × 产业园区 · 跨产业撮合与游学转化平台", 9, GREY, False)])
    textbox(slide, Inches(11.8), Inches(7.05), Inches(1.2), Inches(0.35),
            [(str(page), 9, GREY, False)], align=PP_ALIGN.RIGHT)


def simple_table(slide, x, y, w, headers, rows, col_w, hsize=11, bsize=10,
                 bold_cols=(), color_cols=None, rh=Inches(0.5)):
    nrow = len(rows) + 1
    ncol = len(headers)
    tbl = slide.shapes.add_table(nrow, ncol, x, y, w, rh * nrow).table
    for c, cw in enumerate(col_w):
        tbl.columns[c].width = cw
    for c, h in enumerate(headers):
        cell = tbl.cell(0, c)
        cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        r = cell.text_frame.paragraphs[0].add_run(); r.text = h
        r.font.size = Pt(hsize); r.font.bold = True; r.font.color.rgb = WHITE
        set_cjk(r)
    for ri, row in enumerate(rows, start=1):
        for ci, v in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.fill.solid(); cell.fill.fore_color.rgb = WHITE if ri % 2 else LIGHT
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            r = cell.text_frame.paragraphs[0].add_run(); r.text = v
            r.font.size = Pt(bsize)
            r.font.bold = (ci in bold_cols)
            r.font.color.rgb = (color_cols[ci] if color_cols else DARK)
            set_cjk(r)
    return tbl


# ================================================================ 1. 封面
s = add_slide()
rect(s, 0, 0, SW, SH, fill=NAVY)
rect(s, 0, Inches(3.55), SW, Pt(3), fill=AMBER)
textbox(s, Inches(1), Inches(1.75), Inches(11.3), Inches(0.7),
        [(C.PROJECT_TITLE, 40, WHITE, True)], align=PP_ALIGN.CENTER)
textbox(s, Inches(1), Inches(2.6), Inches(11.3), Inches(0.6),
        [(C.PROJECT_SUBTITLE, 21, AMBER, True)], align=PP_ALIGN.CENTER)
textbox(s, Inches(1), Inches(3.75), Inches(11.3), Inches(1.4),
        [("四期 · 每期两天半 · 上海（长三角延伸）", 16, RGBColor(0xC9, 0xD6, 0xE3), False, 8),
         (C.STRATEGIC_LEVEL, 15, WHITE, True, 8)], align=PP_ALIGN.CENTER)
textbox(s, Inches(1), Inches(6.0), Inches(11.3), Inches(1.1),
        [(C.LEAD_HOSTS, 13, WHITE, True, 6),
         (C.PARTNER, 13, RGBColor(0xC9, 0xD6, 0xE3), False)], align=PP_ALIGN.CENTER)

# ================================================================ 2. 总览 / 定位
s = add_slide()
header(s, "总览", "项目定位与四期安排", "以“产业需求驱动”为核心，形成“考察 + 对接 + 项目转化”闭环")
textbox(s, Inches(0.5), Inches(1.35), Inches(12.3), Inches(0.5),
        [(C.FRAMEWORK_INTRO, 12.5, DARK, False)])
cw = Inches(3.05); gap = Inches(0.18); x0 = Inches(0.5); y0 = Inches(2.15); ch = Inches(3.6)
for i, t in enumerate(C.TOURS):
    x = x0 + (cw + gap) * i
    rect(s, x, y0, cw, ch, fill=LIGHT, line=RGBColor(0xD5, 0xDE, 0xE8))
    rect(s, x, y0, cw, Inches(0.95), fill=BLUE)
    textbox(s, x + Inches(0.12), y0 + Inches(0.08), cw - Inches(0.24), Inches(0.85),
            [(t["code"], 15, AMBER, True, 2), (t["window"], 12, WHITE, False)])
    textbox(s, x + Inches(0.14), y0 + Inches(1.05), cw - Inches(0.28), ch - Inches(1.15),
            [(t["theme"], 12.5, NAVY, True, 6),
             ("聚焦：" + t["focus"], 9.5, GREY, False, 6),
             ("用钢场景：" + t["steel"], 9.5, STEEL, True, 0)])
textbox(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.85),
        [("规模：" + C.SCALE + "　|　成本口径：" + C.LOGISTICS, 11, DARK, False, 2),
         ("战略目标：" + C.STRATEGIC_GOAL, 12, AMBER, True)])
footer(s, 2)

# ================================================================ 3-6. 四期详情
def tour_slide(t, page):
    s = add_slide()
    header(s, t["code"] + " · " + t["window"], t["theme"])
    textbox(s, Inches(0.5), Inches(1.32), Inches(12.3), Inches(1.05),
            [("聚焦：" + t["focus"], 12.5, DARK, False, 4),
             ("主题课：" + t["lecture"], 12, BLUE, True, 4),
             ("用钢新场景：" + t["steel"], 12, STEEL, True, 0)])
    simple_table(
        s, Inches(0.5), Inches(2.55), Inches(12.3),
        ["类别", "考察点（项目 / 园区 / 工厂）", "看点 · 用钢关联"],
        [[c, n, note] for c, n, note in t["sites"]],
        [Inches(2.4), Inches(4.3), Inches(5.6)],
        hsize=12, bsize=11, bold_cols=(0, 1),
        color_cols=[STEEL, DARK, DARK])
    footer(s, page)

for i, t in enumerate(C.TOURS):
    tour_slide(t, 3 + i)

# ================================================================ 7. 2.5天标准行程
s = add_slide()
header(s, "行程模板", "每期两天半 · 标准时间节点排布", "四期共用，实际考察点按各期主题替换")
rows = len(C.DAY_TEMPLATE) + 1
tbl = s.shapes.add_table(rows, 4, Inches(0.5), Inches(1.35), Inches(12.3), Inches(0.34) * rows).table
for c, w in zip(range(4), [Inches(1.4), Inches(2.0), Inches(4.6), Inches(4.3)]):
    tbl.columns[c].width = w
for c, h in enumerate(["日期", "时间", "环节", "说明"]):
    cell = tbl.cell(0, c); cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    r = cell.text_frame.paragraphs[0].add_run(); r.text = h
    r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = WHITE; set_cjk(r)
day_color = {"第 1 天": RGBColor(0xE3, 0xEC, 0xF4), "第 2 天": RGBColor(0xED, 0xF3, 0xEA),
             "第 3 天": RGBColor(0xF6, 0xEE, 0xE1)}
for ri, (d, tm, act, note) in enumerate(C.DAY_TEMPLATE, start=1):
    for ci, v in enumerate([d, tm, act, note]):
        cell = tbl.cell(ri, ci); cell.fill.solid(); cell.fill.fore_color.rgb = day_color.get(d, WHITE)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        r = cell.text_frame.paragraphs[0].add_run(); r.text = v
        r.font.size = Pt(9.5); r.font.bold = (ci == 2)
        r.font.color.rgb = NAVY if ci == 0 else DARK; set_cjk(r)
footer(s, 7)

# ================================================================ 8. 合作定位：总会对总会 · 对等
s = add_slide()
header(s, "合作定位", "总会对总会 · 商务对商务 · 对等交流",
       "把学术公信力与买方网络，转化为可量化、可定价的商业筹码")
textbox(s, Inches(0.5), Inches(1.35), Inches(12.3), Inches(0.55),
        [(C.FRAMEWORK_INTRO, 12, DARK, False)])
# 双方主体（对等）
cw = Inches(6.0); gap = Inches(0.3); x0 = Inches(0.5); y0 = Inches(2.0); ch = Inches(1.75)
for i, (role, name, duty) in enumerate(C.PARTIES):
    x = x0 + (cw + gap) * i
    rect(s, x, y0, cw, ch, fill=LIGHT, line=RGBColor(0xD5, 0xDE, 0xE8))
    rect(s, x, y0, cw, Inches(0.55), fill=(STEEL if i == 0 else AMBER))
    textbox(s, x + Inches(0.16), y0 + Inches(0.1), cw - Inches(0.32), Inches(0.5),
            [(role, 13, WHITE, True)])
    textbox(s, x + Inches(0.18), y0 + Inches(0.65), cw - Inches(0.36), ch - Inches(0.75),
            [(name, 13, NAVY, True, 6), (duty, 11, GREY, False)])
# 对“各收各钱”的立场
textbox(s, Inches(0.5), Inches(3.95), Inches(12.3), Inches(0.4),
        [("针对对方“各收各钱、自负盈亏”方案的指导性立场（红线）：", 13, RED, True)])
textbox(s, Inches(0.6), Inches(4.4), Inches(12.2), Inches(2.3),
        [("• " + t, 12.5, DARK, False, 8) for t in C.STANCE])
footer(s, 8)

# ================================================================ 9. 我方核心商业资产
s = add_slide()
header(s, "商业资产", "我方输出的商业资产界定", "谈判开场先锚定：并非“帮忙”，而是高壁垒的商业资产")
cw = Inches(4.0); gap = Inches(0.15); x0 = Inches(0.5); y0 = Inches(1.5); ch = Inches(4.9)
for i, (a, b, c) in enumerate(C.ASSETS):
    x = x0 + (cw + gap) * i
    rect(s, x, y0, cw, ch, fill=LIGHT, line=RGBColor(0xD5, 0xDE, 0xE8))
    rect(s, x, y0, cw, Inches(1.0), fill=BLUE)
    textbox(s, x + Inches(0.14), y0 + Inches(0.12), cw - Inches(0.28), Inches(0.9),
            [(a, 13.5, WHITE, True)])
    textbox(s, x + Inches(0.16), y0 + Inches(1.15), cw - Inches(0.32), ch - Inches(1.3),
            [(b, 12, NAVY, True, 10), ("价值点：" + c, 11.5, DARK, False)])
footer(s, 9)

# ================================================================ 10. 三种对价方案
s = add_slide()
header(s, "对价方案", "合作模式与商业对价方案（三选一或组合）",
       "明确：仅直接运营成本独立核算，平台赋能层面必须有对价回流")
cw = Inches(4.0); gap = Inches(0.15); x0 = Inches(0.5); y0 = Inches(1.5); ch = Inches(5.0)
band = [STEEL, AMBER, GREEN]
for i, m in enumerate(C.MODELS):
    x = x0 + (cw + gap) * i
    rect(s, x, y0, cw, ch, fill=LIGHT, line=RGBColor(0xD5, 0xDE, 0xE8))
    rect(s, x, y0, cw, Inches(1.05), fill=band[i])
    textbox(s, x + Inches(0.14), y0 + Inches(0.1), cw - Inches(0.28), Inches(0.95),
            [(m["name"], 12.5, WHITE, True)])
    lines = [("适用：" + m["case"], 11, GREY, True, 8)]
    lines += [("• " + r, 11, DARK, False, 8) for r in m["rights"]]
    textbox(s, x + Inches(0.16), y0 + Inches(1.2), cw - Inches(0.32), ch - Inches(1.35), lines)
footer(s, 10)

# ================================================================ 11. 分成落位表
s = add_slide()
header(s, "分成落位", "分成落位（明确 · 可执行 · 可追溯）", "改前期模糊：每类收益都有计费方式、比例与结算凭证")
simple_table(
    s, Inches(0.5), Inches(1.5), Inches(12.3),
    ["收益类型", "计费方式", "我方对价 / 分成", "结算与凭证"],
    [list(r) for r in C.REVENUE_TABLE],
    [Inches(2.7), Inches(2.3), Inches(3.9), Inches(3.4)],
    hsize=12, bsize=11, bold_cols=(0, 2),
    color_cols=[NAVY, DARK, AMBER, GREY], rh=Inches(0.7))
textbox(s, Inches(0.5), Inches(5.9), Inches(12.3), Inches(0.9),
        [("说明：会务与接待等直接运营成本可各自独立核算；以上品牌 / 前端 / 后端 / 居间四类"
          "对价须约定于《联合主办合作框架与对价清单》，并以名单共管库作为分佣与追溯依据。",
          11.5, DARK, False)])
footer(s, 11)

# ================================================================ 12. 核心保护条款
s = add_slide()
header(s, "保护条款", "核心保护条款（数据与资源防火墙）", "防止过河拆桥、单方面洗走核心资源池")
y = Inches(1.5)
for i, (title, body) in enumerate(C.PROTECTION):
    rect(s, Inches(0.5), y, Inches(12.3), Inches(1.2), fill=LIGHT, line=RGBColor(0xD5, 0xDE, 0xE8))
    rect(s, Inches(0.5), y, Inches(0.14), Inches(1.2), fill=AMBER)
    textbox(s, Inches(0.8), y + Inches(0.12), Inches(11.8), Inches(1.0),
            [(str(i + 1) + "）" + title, 14, NAVY, True, 4), (body, 12, DARK, False)])
    y = y + Inches(1.35)
footer(s, 12)

# ================================================================ 13. 底线 + 反应分级
s = add_slide()
header(s, "底线与应对", "我方底线 与 对方反应分级应对")
textbox(s, Inches(0.5), Inches(1.4), Inches(6), Inches(0.4), [("我方要守住的底线", 15, RED, True)])
textbox(s, Inches(0.55), Inches(1.95), Inches(6.1), Inches(4.5),
        [("• " + t, 13, DARK, False, 14) for t in C.BOTTOM_LINES])
textbox(s, Inches(6.9), Inches(1.4), Inches(6), Inches(0.4), [("对方反应 · 分级应对", 15, NAVY, True)])
tier_color = {"接受": GREEN, "犹豫": AMBER, "拒绝": RED}
y = Inches(1.95)
for k, v in C.RESPONSE_TIERS:
    rect(s, Inches(6.9), y, Inches(5.9), Inches(1.35), fill=LIGHT, line=RGBColor(0xD5, 0xDE, 0xE8))
    rect(s, Inches(6.9), y, Inches(1.4), Inches(1.35), fill=tier_color[k])
    textbox(s, Inches(6.95), y + Inches(0.45), Inches(1.35), Inches(0.5),
            [(k, 14, WHITE, True)], align=PP_ALIGN.CENTER)
    textbox(s, Inches(8.45), y + Inches(0.14), Inches(4.25), Inches(1.05),
            [(v, 12, DARK, False)], anchor=MSO_ANCHOR.MIDDLE)
    y = y + Inches(1.5)
footer(s, 13)

# ================================================================ 14. 对外话术
s = add_slide()
header(s, "对外话术", "对外沟通话术（可直接用）", "先肯定、再升维、合理化收费、导向长期合作")
rect(s, Inches(0.5), Inches(1.45), Inches(12.3), Inches(2.55), fill=LIGHT, line=RGBColor(0xD5, 0xDE, 0xE8))
rect(s, Inches(0.5), Inches(1.45), Inches(0.14), Inches(2.55), fill=STEEL)
textbox(s, Inches(0.8), Inches(1.6), Inches(11.8), Inches(2.3),
        [("① 商业升维版（试水）", 13, STEEL, True, 8), (C.SCRIPT_UPGRADE, 12, DARK, False)])
rect(s, Inches(0.5), Inches(4.2), Inches(12.3), Inches(2.45), fill=LIGHT, line=RGBColor(0xD5, 0xDE, 0xE8))
rect(s, Inches(0.5), Inches(4.2), Inches(0.14), Inches(2.45), fill=AMBER)
textbox(s, Inches(0.8), Inches(4.35), Inches(11.8), Inches(2.2),
        [("② 战略高位切入版（正式磋商）", 13, AMBER, True, 8), (C.SCRIPT_HIGH, 12, DARK, False)])
footer(s, 14)

# ================================================================ 15. 目标里程碑
s = add_slide()
header(s, "目标", "目标里程碑 与 战略定位")
y = Inches(1.7)
for tm, goal in C.MILESTONES:
    rect(s, Inches(0.5), y, Inches(12.3), Inches(1.15), fill=LIGHT, line=RGBColor(0xD5, 0xDE, 0xE8))
    rect(s, Inches(0.5), y, Inches(2.0), Inches(1.15), fill=BLUE)
    textbox(s, Inches(0.55), y + Inches(0.35), Inches(1.9), Inches(0.5),
            [(tm, 15, WHITE, True)], align=PP_ALIGN.CENTER)
    textbox(s, Inches(2.7), y + Inches(0.14), Inches(9.9), Inches(0.9),
            [(goal, 13.5, DARK, False)], anchor=MSO_ANCHOR.MIDDLE)
    y = y + Inches(1.35)
rect(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.85), fill=NAVY)
textbox(s, Inches(0.7), Inches(6.12), Inches(11.9), Inches(0.65),
        [(C.STRATEGIC_GOAL, 15, AMBER, True)], anchor=MSO_ANCHOR.MIDDLE)
footer(s, 15)

# ================================================================ 16. 结束页
s = add_slide()
rect(s, 0, 0, SW, SH, fill=NAVY)
rect(s, 0, Inches(3.5), SW, Pt(3), fill=AMBER)
textbox(s, Inches(1), Inches(2.3), Inches(11.3), Inches(1.0),
        [("我们不是做游学活动，", 26, WHITE, True, 6),
         ("而是做钢铁产业与不动产 / 园区之间的需求撮合与项目转化平台。", 21, AMBER, True)],
        align=PP_ALIGN.CENTER)
textbox(s, Inches(1), Inches(4.0), Inches(11.3), Inches(1.4),
        [(C.STRATEGIC_LEVEL, 15, RGBColor(0xC9, 0xD6, 0xE3), False, 10),
         (C.LEAD_HOSTS, 13, WHITE, True, 4),
         (C.PARTNER, 12, RGBColor(0xC9, 0xD6, 0xE3), False)], align=PP_ALIGN.CENTER)

OUT = "output/中钢协_上海长三角产业研学考察计划与合作框架.pptx"
prs.save(OUT)
print("saved:", OUT, "slides:", len(prs.slides._sldIdLst))
