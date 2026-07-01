# -*- coding: utf-8 -*-
"""生成《上海 / 长三角产业研学考察计划与合作框架》PPT。"""
import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

import content as C

# ---------------------------------------------------------------- 主题配色
NAVY = RGBColor(0x14, 0x2B, 0x45)      # 深钢蓝（主色）
BLUE = RGBColor(0x1F, 0x4E, 0x79)      # 中蓝
STEEL = RGBColor(0x2E, 0x6D, 0xA4)     # 钢蓝亮
AMBER = RGBColor(0xE8, 0x8A, 0x1A)     # 琥珀（强调）
LIGHT = RGBColor(0xEE, 0xF3, 0xF8)     # 浅底
GREY = RGBColor(0x5A, 0x63, 0x6E)      # 灰字
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
    """同时设置拉丁与东亚字体，确保中文显示正常。"""
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
    from pptx.enum.shapes import MSO_SHAPE
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(1)
    return sp


def textbox(slide, x, y, w, h, lines, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT):
    """lines: list of (text, size, color, bold, [space_after_pt])"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Pt(4)
    tf.margin_right = Pt(4)
    tf.margin_top = Pt(2)
    tf.margin_bottom = Pt(2)
    for i, item in enumerate(lines):
        text, size, color, bold = item[0], item[1], item[2], item[3]
        sa = item[4] if len(item) > 4 else 4
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(sa)
        p.line_spacing = 1.08
        r = p.add_run()
        r.text = text
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        set_cjk(r)
    return tb


def header(slide, tag, title, subtitle=None):
    """标准内容页页眉：左侧色条 + 标签 + 标题。"""
    rect(slide, 0, 0, SW, Inches(1.15), fill=NAVY)
    rect(slide, 0, Inches(1.15), SW, Pt(3), fill=AMBER)
    textbox(slide, Inches(0.5), Inches(0.12), Inches(3), Inches(0.4),
            [(tag, 13, AMBER, True)])
    ln = [(title, 25, WHITE, True)]
    if subtitle:
        ln.append((subtitle, 13, RGBColor(0xC9, 0xD6, 0xE3), False))
    textbox(slide, Inches(0.5), Inches(0.42), Inches(12.3), Inches(0.72), ln)


def footer(slide, page):
    textbox(slide, Inches(0.5), Inches(7.05), Inches(9), Inches(0.35),
            [("钢铁 × 地产 × 产业园区 · 跨产业撮合与游学转化平台", 9, GREY, False)])
    textbox(slide, Inches(11.8), Inches(7.05), Inches(1.2), Inches(0.35),
            [(str(page), 9, GREY, False)], align=PP_ALIGN.RIGHT)


# ================================================================ 1. 封面
s = add_slide()
rect(s, 0, 0, SW, SH, fill=NAVY)
rect(s, 0, Inches(3.55), SW, Pt(3), fill=AMBER)
textbox(s, Inches(1), Inches(1.9), Inches(11.3), Inches(0.7),
        [(C.PROJECT_TITLE, 40, WHITE, True)], align=PP_ALIGN.CENTER)
textbox(s, Inches(1), Inches(2.75), Inches(11.3), Inches(0.6),
        [(C.PROJECT_SUBTITLE, 22, AMBER, True)], align=PP_ALIGN.CENTER)
textbox(s, Inches(1), Inches(3.9), Inches(11.3), Inches(1.4),
        [("四期 · 每期两天半 · 上海（长三角延伸）", 16, RGBColor(0xC9, 0xD6, 0xE3), False, 8),
         ("标杆项目 · 标杆园区 · 智能制造 · 科技企业 · 供需闭门撮合", 14, RGBColor(0xC9, 0xD6, 0xE3), False, 8)],
        align=PP_ALIGN.CENTER)
textbox(s, Inches(1), Inches(6.2), Inches(11.3), Inches(0.9),
        [(C.HOSTS, 13, WHITE, True, 4)], align=PP_ALIGN.CENTER)

# ================================================================ 2. 总览 / 定位
s = add_slide()
header(s, "总览", "项目定位与四期安排", "以“产业需求驱动”为核心，形成“考察 + 对接 + 项目转化”闭环")
textbox(s, Inches(0.5), Inches(1.35), Inches(12.3), Inches(0.5),
        [(C.FRAMEWORK_INTRO, 13, DARK, False)])
# 四张期次卡片
cw = Inches(3.05)
gap = Inches(0.18)
x0 = Inches(0.5)
y0 = Inches(2.05)
ch = Inches(3.7)
for i, t in enumerate(C.TOURS):
    x = x0 + (cw + gap) * i
    rect(s, x, y0, cw, ch, fill=LIGHT, line=RGBColor(0xD5, 0xDE, 0xE8))
    rect(s, x, y0, cw, Inches(0.95), fill=BLUE)
    textbox(s, x + Inches(0.12), y0 + Inches(0.08), cw - Inches(0.24), Inches(0.85),
            [(t["code"], 15, AMBER, True, 2),
             (t["window"], 12, WHITE, False)])
    body = [(t["theme"], 12.5, NAVY, True, 6)]
    body.append(("聚焦：" + t["focus"], 10, GREY, False, 6))
    body.append(("用钢场景：" + t["steel"], 10, STEEL, True, 0))
    textbox(s, x + Inches(0.14), y0 + Inches(1.05), cw - Inches(0.28), ch - Inches(1.15), body)
textbox(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.8),
        [("规模：" + C.SCALE + "　|　保障：" + C.LOGISTICS, 11, DARK, False, 2),
         ("战略目标：" + C.STRATEGIC_GOAL, 12, AMBER, True)])
footer(s, 2)

# ================================================================ 3-6. 四期详情
def tour_slide(t, page):
    s = add_slide()
    header(s, t["code"] + " · " + t["window"], t["theme"])
    # 聚焦 + 主题课 + 用钢
    textbox(s, Inches(0.5), Inches(1.32), Inches(12.3), Inches(1.05),
            [("聚焦：" + t["focus"], 12.5, DARK, False, 4),
             ("主题课：" + t["lecture"], 12, BLUE, True, 4),
             ("用钢新场景：" + t["steel"], 12, STEEL, True, 0)])
    # 参访点表
    rows = len(t["sites"]) + 1
    tbl_shape = s.shapes.add_table(rows, 3, Inches(0.5), Inches(2.55),
                                   Inches(12.3), Inches(0.5) * rows)
    table = tbl_shape.table
    table.columns[0].width = Inches(2.4)
    table.columns[1].width = Inches(4.3)
    table.columns[2].width = Inches(5.6)
    heads = ["类别", "考察点（项目 / 园区 / 工厂）", "看点 · 用钢关联"]
    for c, h in enumerate(heads):
        cell = table.cell(0, c)
        cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        r = p.add_run(); r.text = h; r.font.size = Pt(12); r.font.bold = True
        r.font.color.rgb = WHITE; set_cjk(r)
    for ri, (cat, name, note) in enumerate(t["sites"], start=1):
        vals = [cat, name, note]
        for ci, v in enumerate(vals):
            cell = table.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if ri % 2 else LIGHT
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
            r = p.add_run(); r.text = v
            r.font.size = Pt(11); r.font.bold = (ci == 1)
            r.font.color.rgb = STEEL if ci == 0 else DARK
            if ci == 0:
                r.font.bold = True
            set_cjk(r)
    footer(s, page)
    return s

for i, t in enumerate(C.TOURS):
    tour_slide(t, 3 + i)

# ================================================================ 7. 2.5天标准行程
s = add_slide()
header(s, "行程模板", "每期两天半 · 标准时间节点排布", "四期共用，实际考察点按各期主题替换")
rows = len(C.DAY_TEMPLATE) + 1
tbl_shape = s.shapes.add_table(rows, 4, Inches(0.5), Inches(1.35),
                               Inches(12.3), Inches(0.34) * rows)
table = tbl_shape.table
table.columns[0].width = Inches(1.4)
table.columns[1].width = Inches(2.0)
table.columns[2].width = Inches(4.6)
table.columns[3].width = Inches(4.3)
heads = ["日期", "时间", "环节", "说明"]
for c, h in enumerate(heads):
    cell = table.cell(0, c)
    cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = cell.text_frame.paragraphs[0]
    r = p.add_run(); r.text = h; r.font.size = Pt(11); r.font.bold = True
    r.font.color.rgb = WHITE; set_cjk(r)
day_color = {"第 1 天": RGBColor(0xE3, 0xEC, 0xF4),
             "第 2 天": RGBColor(0xED, 0xF3, 0xEA),
             "第 3 天": RGBColor(0xF6, 0xEE, 0xE1)}
for ri, (d, tm, act, note) in enumerate(C.DAY_TEMPLATE, start=1):
    vals = [d, tm, act, note]
    for ci, v in enumerate(vals):
        cell = table.cell(ri, ci)
        cell.fill.solid(); cell.fill.fore_color.rgb = day_color.get(d, WHITE)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        r = p.add_run(); r.text = v
        r.font.size = Pt(9.5); r.font.bold = (ci == 2)
        r.font.color.rgb = NAVY if ci == 0 else DARK
        set_cjk(r)
footer(s, 7)

# ================================================================ 8. 合作框架总览
s = add_slide()
header(s, "合作框架", "统筹合作框架 · 总览", "中钢协流通分会 × 复旦房地产研究中心 × 杨浦区科技企业联合会")
textbox(s, Inches(0.5), Inches(1.4), Inches(12.3), Inches(0.8),
        [(C.FRAMEWORK_INTRO, 13, DARK, False)])
# 三方角色
roles = [
    ("主导 / 控盘方", "复旦大学房地产研究中心 + 杨浦区科技企业联合会", "品牌 · 规则 · 内容 · 收费权 · 政府与园区资源"),
    ("战略协同 / 渠道方", "中钢协流通分会", "钢铁企业资源 · 组织参与 · 协助邀约"),
    ("我方目标定位", "上海及长三角总战略合作伙伴", "独家 / 优先承接中钢协在沪及长三角研学、撮合与项目转化"),
]
cw = Inches(4.0); gap = Inches(0.15); x0 = Inches(0.5); y0 = Inches(2.35); ch = Inches(2.4)
for i, (a, b, c) in enumerate(roles):
    x = x0 + (cw + gap) * i
    rect(s, x, y0, cw, ch, fill=LIGHT, line=RGBColor(0xD5, 0xDE, 0xE8))
    rect(s, x, y0, cw, Inches(0.6), fill=(AMBER if i == 2 else BLUE))
    textbox(s, x + Inches(0.14), y0 + Inches(0.1), cw - Inches(0.28), Inches(0.5),
            [(a, 13, WHITE, True)])
    textbox(s, x + Inches(0.16), y0 + Inches(0.75), cw - Inches(0.32), ch - Inches(0.9),
            [(b, 13.5, NAVY, True, 8), (c, 11, GREY, False)])
textbox(s, Inches(0.5), Inches(5.0), Inches(12.3), Inches(1.6),
        [("对方顾虑与我方应对：", 13, NAVY, True, 6)] +
        [("• " + k + " → " + v, 10.5, DARK, False, 3) for k, v in C.PAIN_RESPONSES])
footer(s, 8)

# ================================================================ 9. 两阶段
s = add_slide()
header(s, "合作节奏", "两阶段合作设计", "先共建、后盈利：初期不计营收，跑通后按比例提成")
cw = Inches(6.0); gap = Inches(0.3); x0 = Inches(0.5); y0 = Inches(1.6); ch = Inches(4.9)
for i, ph in enumerate(C.PHASES):
    x = x0 + (cw + gap) * i
    rect(s, x, y0, cw, ch, fill=LIGHT, line=RGBColor(0xD5, 0xDE, 0xE8))
    rect(s, x, y0, cw, Inches(0.75), fill=(BLUE if i == 0 else AMBER))
    textbox(s, x + Inches(0.18), y0 + Inches(0.14), cw - Inches(0.36), Inches(0.6),
            [(ph["name"], 15, WHITE, True)])
    lines = [("• " + p, 13, DARK, False, 10) for p in ph["points"]]
    textbox(s, x + Inches(0.22), y0 + Inches(0.95), cw - Inches(0.44), ch - Inches(1.1), lines)
footer(s, 9)

# ================================================================ 10. 盈利底线与分配
s = add_slide()
header(s, "红线", "盈利底线与分配原则", "我方无会员制收入，不以亏损为合作前提；分配向我方投入倾斜")
y = Inches(1.5)
for i, (title, body) in enumerate(C.PROFIT_PRINCIPLES):
    rect(s, Inches(0.5), y, Inches(12.3), Inches(1.35), fill=LIGHT, line=RGBColor(0xD5, 0xDE, 0xE8))
    rect(s, Inches(0.5), y, Inches(0.14), Inches(1.35), fill=AMBER)
    textbox(s, Inches(0.8), y + Inches(0.12), Inches(11.8), Inches(1.15),
            [(title, 15, NAVY, True, 5), (body, 12.5, DARK, False)])
    y = y + Inches(1.55)
footer(s, 10)

# ================================================================ 11. 收费结构 + 里程碑
s = add_slide()
header(s, "变现与目标", "收费结构（初期让利） 与 目标里程碑")
# 左：收费结构
textbox(s, Inches(0.5), Inches(1.4), Inches(6), Inches(0.4),
        [("收费结构（承接原报价逻辑）", 14, NAVY, True)])
rows = len(C.PRICING) + 1
tbl = s.shapes.add_table(rows, 2, Inches(0.5), Inches(1.9), Inches(6.1), Inches(0.5) * rows).table
tbl.columns[0].width = Inches(2.1); tbl.columns[1].width = Inches(4.0)
for c, h in enumerate(["项目", "说明"]):
    cell = tbl.cell(0, c); cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    r = cell.text_frame.paragraphs[0].add_run(); r.text = h
    r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = WHITE; set_cjk(r)
for ri, (a, b) in enumerate(C.PRICING, start=1):
    for ci, v in enumerate((a, b)):
        cell = tbl.cell(ri, ci); cell.fill.solid()
        cell.fill.fore_color.rgb = WHITE if ri % 2 else LIGHT
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        r = cell.text_frame.paragraphs[0].add_run(); r.text = v
        r.font.size = Pt(10); r.font.bold = (ci == 0)
        r.font.color.rgb = STEEL if ci == 0 else DARK; set_cjk(r)
# 右：里程碑
textbox(s, Inches(6.9), Inches(1.4), Inches(6), Inches(0.4),
        [("目标里程碑", 14, NAVY, True)])
y = Inches(1.95)
for tm, goal in C.MILESTONES:
    rect(s, Inches(6.9), y, Inches(5.9), Inches(1.0), fill=LIGHT, line=RGBColor(0xD5, 0xDE, 0xE8))
    rect(s, Inches(6.9), y, Inches(1.4), Inches(1.0), fill=BLUE)
    textbox(s, Inches(6.95), y + Inches(0.28), Inches(1.35), Inches(0.5),
            [(tm, 13, WHITE, True)], align=PP_ALIGN.CENTER)
    textbox(s, Inches(8.45), y + Inches(0.14), Inches(4.25), Inches(0.8),
            [(goal, 11.5, DARK, False)], anchor=MSO_ANCHOR.MIDDLE)
    y = y + Inches(1.2)
footer(s, 11)

# ================================================================ 12. 结束页
s = add_slide()
rect(s, 0, 0, SW, SH, fill=NAVY)
rect(s, 0, Inches(3.5), SW, Pt(3), fill=AMBER)
textbox(s, Inches(1), Inches(2.4), Inches(11.3), Inches(1.0),
        [("我们不是做游学活动，", 26, WHITE, True, 6),
         ("而是做钢铁产业与地产 / 园区之间的需求撮合与项目转化平台。", 22, AMBER, True)],
        align=PP_ALIGN.CENTER)
textbox(s, Inches(1), Inches(4.0), Inches(11.3), Inches(1.2),
        [(C.STRATEGIC_GOAL, 16, RGBColor(0xC9, 0xD6, 0xE3), False, 8),
         (C.HOSTS, 13, WHITE, True)], align=PP_ALIGN.CENTER)

OUT = "output/中钢协_上海长三角产业研学考察计划与合作框架.pptx"
prs.save(OUT)
print("saved:", OUT, "slides:", len(prs.slides._sldIdLst))
