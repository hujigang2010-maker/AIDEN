"""
生成《上海商办楼宇与产业园区市场深度报告》汇报版 PPT（精装版）
易居研究院 × 复旦大学住房政策研究中心
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree
import os

OUTPUT_DIR = "/workspace/上海商办楼宇与产业园区市场深度报告"
os.makedirs(OUTPUT_DIR, exist_ok=True)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SLIDE_W = prs.slide_width
SLIDE_H = prs.slide_height

# ====== 配色（杂志感 / 商业报告 / 暗蓝 + 金棕 + 留白）======
NAVY = RGBColor(0x10, 0x2A, 0x5B)        # 深海蓝（主色）
NAVY_DARK = RGBColor(0x07, 0x1A, 0x3F)
INK = RGBColor(0x1E, 0x2A, 0x4A)
BLUE = RGBColor(0x2A, 0x5B, 0xA8)        # 主蓝
BLUE_SOFT = RGBColor(0x4E, 0x7A, 0xC0)
LIGHT_BLUE = RGBColor(0xE9, 0xF0, 0xFB)
PAPER = RGBColor(0xF7, 0xF8, 0xFC)       # 留白底
GOLD = RGBColor(0xC8, 0x91, 0x3A)        # 金棕（点缀）
GOLD_SOFT = RGBColor(0xE6, 0xB6, 0x6E)
ORANGE = RGBColor(0xE0, 0x6A, 0x2C)
GREEN = RGBColor(0x2F, 0x7D, 0x5C)
TEAL = RGBColor(0x1E, 0x77, 0x8E)
PURPLE = RGBColor(0x6B, 0x4E, 0xA1)
RED = RGBColor(0xB8, 0x3A, 0x3F)
DARK = RGBColor(0x1A, 0x22, 0x2E)
GRAY = RGBColor(0x6E, 0x73, 0x7C)
GRAY_LIGHT = RGBColor(0xC9, 0xCE, 0xD8)
LINE = RGBColor(0xD8, 0xDE, 0xEA)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def set_font(run, font="微软雅黑", size=18, bold=False, color=DARK, italic=False):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    rpr = run._r.get_or_add_rPr()
    east = rpr.find(qn("a:ea"))
    if east is None:
        ea = etree.SubElement(rpr, qn("a:ea"))
        ea.set("typeface", font)
    else:
        east.set("typeface", font)
    latin = rpr.find(qn("a:latin"))
    if latin is None:
        l = etree.SubElement(rpr, qn("a:latin"))
        l.set("typeface", font)
    else:
        latin.set("typeface", font)


def add_rect(slide, x, y, w, h, fill, line=None, line_w=0.5):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_w)
    shp.shadow.inherit = False
    return shp


def add_round_rect(slide, x, y, w, h, fill, line=None, radius=0.04):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.5)
    shp.shadow.inherit = False
    try:
        shp.adjustments[0] = radius
    except Exception:
        pass
    return shp


def add_oval(slide, x, y, w, h, fill, line=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
    shp.shadow.inherit = False
    return shp


def add_text(slide, x, y, w, h, text, size=14, bold=False, color=DARK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font="微软雅黑",
             italic=False, line_spacing=1.15):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(2)
    tf.margin_right = Pt(2)
    tf.margin_top = Pt(1)
    tf.margin_bottom = Pt(1)
    tf.vertical_anchor = anchor
    lines = text if isinstance(text, list) else [text]
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = Pt(2)
        r = p.add_run()
        set_font(r, font=font, size=size, bold=bold, color=color, italic=italic)
        r.text = ln
    return tb


def add_bullets(slide, x, y, w, h, items, size=12, color=DARK,
                bullet_color=BLUE, bullet_char="●", line_spacing=1.25,
                space_after=4):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(2)
    tf.margin_right = Pt(2)
    tf.margin_top = Pt(1)
    tf.margin_bottom = Pt(1)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = line_spacing
        p.space_after = Pt(space_after)
        r1 = p.add_run()
        set_font(r1, size=size, color=bullet_color, bold=True)
        r1.text = bullet_char + "  "
        r2 = p.add_run()
        set_font(r2, size=size, color=color)
        r2.text = item


def add_line(slide, x1, y1, x2, y2, color=LINE, width=1.0):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    ln.line.color.rgb = color
    ln.line.width = Pt(width)
    ln.shadow.inherit = False
    return ln


def new_blank():
    return prs.slides.add_slide(prs.slide_layouts[6])


page_counter = [0]


def std_page_decor(slide, chapter, title, subtitle=None, section_color=BLUE):
    """ 标准内容页装饰：顶部留白、左侧色块、章节标识、标题、底部页脚 """
    # 顶部 - 极简色带 + 装饰金线
    add_rect(slide, 0, 0, Inches(0.32), SLIDE_H, NAVY)
    # 章节标签（旋转效果用文本框靠左竖排不太理想，仍水平）
    add_text(slide, Inches(0.0), Inches(0.0), Inches(0.32), SLIDE_H, "",
             size=10, color=WHITE)
    # 顶栏章节小标签
    add_text(slide, Inches(0.6), Inches(0.45), Inches(8.0), Inches(0.3),
             chapter, size=11, bold=True, color=section_color)
    # 主标题
    add_text(slide, Inches(0.6), Inches(0.75), Inches(12.0), Inches(0.55),
             title, size=24, bold=True, color=NAVY)
    # 装饰短线
    add_rect(slide, Inches(0.6), Inches(1.32), Inches(0.55), Inches(0.05), GOLD)
    if subtitle:
        add_text(slide, Inches(1.25), Inches(1.28), Inches(11.5), Inches(0.3),
                 subtitle, size=11, color=GRAY, italic=True)
    # 底部
    add_line(slide, Inches(0.6), SLIDE_H - Inches(0.45),
             SLIDE_W - Inches(0.6), SLIDE_H - Inches(0.45),
             color=LINE, width=0.75)
    add_text(slide, Inches(0.6), SLIDE_H - Inches(0.42), Inches(8.5), Inches(0.3),
             "上海商办楼宇与产业园区市场深度报告  ·  讨论稿",
             size=9, color=GRAY)
    add_text(slide, SLIDE_W - Inches(4.5), SLIDE_H - Inches(0.42),
             Inches(3.9), Inches(0.3),
             "易居研究院 × 复旦大学住房政策研究中心",
             size=9, color=GRAY, align=PP_ALIGN.RIGHT)
    page_counter[0] += 1
    add_text(slide, SLIDE_W - Inches(0.55), SLIDE_H - Inches(0.42),
             Inches(0.4), Inches(0.3),
             f"{page_counter[0]:02d}", size=10, bold=True, color=NAVY,
             align=PP_ALIGN.RIGHT)


def content_slide(chapter, title, subtitle=None, section_color=BLUE,
                  bg=WHITE):
    s = new_blank()
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, bg)
    std_page_decor(s, chapter, title, subtitle, section_color=section_color)
    return s


def section_divider(part_num, part_name, title, accent=BLUE,
                    bullets=None):
    """ 章节大隔页 """
    s = new_blank()
    # 深色底
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY_DARK)
    # 右侧装饰大块
    add_rect(s, SLIDE_W - Inches(4.5), 0, Inches(4.5), SLIDE_H, NAVY)
    # 金色装饰线
    add_rect(s, Inches(0.8), Inches(1.4), Inches(0.6), Inches(0.06), GOLD)
    # PART 编号
    add_text(s, Inches(0.8), Inches(1.6), Inches(6.0), Inches(0.8),
             f"PART  {part_num}", size=24, bold=True, color=GOLD_SOFT,
             font="微软雅黑")
    # 章节英文/小标识
    add_text(s, Inches(0.8), Inches(2.4), Inches(7.5), Inches(0.5),
             part_name, size=14, color=GRAY_LIGHT)
    # 章节标题
    add_text(s, Inches(0.8), Inches(3.2), Inches(10.5), Inches(1.6),
             title, size=42, bold=True, color=WHITE, line_spacing=1.15)
    # 副要点
    if bullets:
        for i, b in enumerate(bullets):
            y = Inches(5.4 + i * 0.4)
            add_oval(s, Inches(0.85), y + Inches(0.1), Inches(0.12),
                     Inches(0.12), GOLD)
            add_text(s, Inches(1.1), y, Inches(11.5), Inches(0.4),
                     b, size=14, color=WHITE)
    # 右下大装饰编号
    add_text(s, SLIDE_W - Inches(4.0), SLIDE_H - Inches(4.5),
             Inches(3.8), Inches(4.5),
             part_num, size=240, bold=True,
             color=RGBColor(0x1C, 0x36, 0x70), anchor=MSO_ANCHOR.BOTTOM,
             align=PP_ALIGN.CENTER)
    # 页脚
    add_text(s, Inches(0.8), SLIDE_H - Inches(0.55),
             Inches(8.0), Inches(0.3),
             "上海商办楼宇与产业园区市场深度报告  ·  讨论稿",
             size=9, color=GRAY_LIGHT)
    page_counter[0] += 1


def add_card(slide, x, y, w, h, title, items, accent=BLUE,
             title_size=13, body_size=11, icon_text="●"):
    add_rect(slide, x, y, w, h, WHITE, line=LINE)
    # 顶部色条
    add_rect(slide, x, y, w, Inches(0.36), accent)
    # icon
    add_text(slide, x + Inches(0.1), y, Inches(0.4), Inches(0.36),
             icon_text, size=14, bold=True, color=GOLD_SOFT,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_text(slide, x + Inches(0.5), y, w - Inches(0.55), Inches(0.36),
             title, size=title_size, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    add_bullets(slide, x + Inches(0.18), y + Inches(0.45),
                w - Inches(0.3), h - Inches(0.5),
                items, size=body_size, color=DARK, bullet_color=accent,
                bullet_char="•", line_spacing=1.22, space_after=2)


def stat_card(slide, x, y, w, h, value, label, color=BLUE, sub=None):
    add_rect(slide, x, y, w, h, WHITE, line=LINE)
    add_rect(slide, x, y, Inches(0.08), h, color)
    add_text(slide, x + Inches(0.2), y + Inches(0.08),
             w - Inches(0.3), Inches(0.55),
             value, size=24, bold=True, color=color)
    add_text(slide, x + Inches(0.2), y + Inches(0.65),
             w - Inches(0.3), Inches(0.3),
             label, size=11, color=DARK)
    if sub:
        add_text(slide, x + Inches(0.2), y + Inches(0.92),
                 w - Inches(0.3), Inches(0.28),
                 sub, size=9, color=GRAY, italic=True)


# ======================================================
# 1. 封面（杂志式）
# ======================================================
s = new_blank()
# 主底
add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY_DARK)
# 右侧大色块
add_rect(s, Inches(7.5), 0, Inches(5.83), SLIDE_H, NAVY)
# 顶部金线
add_rect(s, 0, 0, SLIDE_W, Inches(0.18), GOLD)
# 底部金线
add_rect(s, 0, SLIDE_H - Inches(0.18), SLIDE_W, Inches(0.18), GOLD)
# 左侧装饰竖条
add_rect(s, Inches(0.6), Inches(1.0), Inches(0.06), Inches(5.5), GOLD)
# 杂志式 issue 标识
add_text(s, Inches(0.9), Inches(0.7), Inches(10), Inches(0.4),
         "JOINT  REPORT  ·  2026", size=12, bold=True,
         color=GOLD_SOFT)
# 联合署名
add_text(s, Inches(0.9), Inches(1.1), Inches(10), Inches(0.4),
         "易居研究院  ×  复旦大学住房政策研究中心", size=15,
         color=GRAY_LIGHT)
# 主标题
add_text(s, Inches(0.9), Inches(2.2), Inches(11), Inches(1.1),
         "上海商办楼宇", size=56, bold=True, color=WHITE,
         line_spacing=1.0)
add_text(s, Inches(0.9), Inches(3.25), Inches(11), Inches(1.1),
         "与产业园区市场", size=56, bold=True, color=WHITE,
         line_spacing=1.0)
# 装饰
add_rect(s, Inches(0.95), Inches(4.45), Inches(0.5), Inches(0.06), GOLD)
add_text(s, Inches(0.9), Inches(4.55), Inches(11), Inches(0.6),
         "深度报告", size=40, bold=True, color=GOLD_SOFT)
# 副标题
add_text(s, Inches(0.9), Inches(5.3), Inches(11), Inches(0.4),
         "全域空间供给  ·  企业需求迁徙  ·  产业载体运营研究",
         size=17, color=GRAY_LIGHT)
add_text(s, Inches(0.9), Inches(5.75), Inches(11), Inches(0.35),
         "Shanghai Commercial Office & Industrial Park Market — In-depth Report",
         size=11, color=GRAY, italic=True)
# 右下角标签 - 研发方案
add_rect(s, SLIDE_W - Inches(5.5), SLIDE_H - Inches(2.0),
         Inches(5.0), Inches(0.7), GOLD)
add_text(s, SLIDE_W - Inches(5.5), SLIDE_H - Inches(2.0),
         Inches(5.0), Inches(0.7),
         "研发方案 · 大纲 · 数据调用 · 执行指引（讨论稿）",
         size=14, bold=True, color=NAVY_DARK,
         anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
# 底部联合课题组
add_text(s, Inches(0.9), SLIDE_H - Inches(1.1),
         Inches(11), Inches(0.4),
         "联合课题组：上海商办楼宇与产业园区市场研究课题组",
         size=12, color=GRAY_LIGHT)


# ======================================================
# 2. 执行摘要 Executive Summary
# ======================================================
s = content_slide("EXECUTIVE SUMMARY", "执行摘要",
                  "Five highlights — 一份报告·五个突破·四类对象·六大库·五大指数",
                  section_color=GOLD)
# 上方 5 个关键数字
nums = [
    ("16", "行政区", "全域覆盖"),
    ("20", "重点板块", "深度专项"),
    ("6", "数据库", "底层支撑"),
    ("5", "核心指数", "指标模型"),
    ("4", "采集阶段", "分期推进"),
]
for i, (v, l, sub) in enumerate(nums):
    x = Inches(0.6 + i * 2.5)
    stat_card(s, x, Inches(1.7), Inches(2.35), Inches(1.25), v, l,
              color=NAVY, sub=sub)

# 中段 - 报告四问
add_text(s, Inches(0.6), Inches(3.1), Inches(12.0), Inches(0.4),
         "本报告回答四个核心问题", size=14, bold=True, color=NAVY)
add_rect(s, Inches(0.6), Inches(3.45), Inches(0.4), Inches(0.04), GOLD)
qns = [
    ("Q1", "供给端在哪里？", "16区 → 街镇 → 楼宇/园区四级底数", BLUE),
    ("Q2", "需求端是谁？", "17 大行业 × 10 类规模 × 9 类清单", TEAL),
    ("Q3", "流动到哪里？", "8 类迁徙 × 多源信号 × 区域吸引力", ORANGE),
    ("Q4", "机会在哪里？", "区域机会 + 产业机会 + 企业线索", GOLD),
]
for i, (q, t, sub, c) in enumerate(qns):
    x = Inches(0.6 + i * 3.1)
    y = Inches(3.7)
    add_rect(s, x, y, Inches(3.0), Inches(1.5), WHITE, line=LINE)
    add_rect(s, x, y, Inches(0.7), Inches(1.5), c)
    add_text(s, x, y, Inches(0.7), Inches(1.5), q, size=22, bold=True,
             color=WHITE, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.85), y + Inches(0.12),
             Inches(2.1), Inches(0.4),
             t, size=13, bold=True, color=NAVY)
    add_text(s, x + Inches(0.85), y + Inches(0.55),
             Inches(2.1), Inches(0.85),
             sub, size=10, color=DARK)

# 下方 - 应用对象
add_text(s, Inches(0.6), Inches(5.4), Inches(12.0), Inches(0.4),
         "面向四类核心应用对象", size=14, bold=True, color=NAVY)
add_rect(s, Inches(0.6), Inches(5.75), Inches(0.4), Inches(0.04), GOLD)
audiences = [
    ("政府", "招商决策 / 空间治理", NAVY),
    ("园区/楼宇运营方", "招商策略 / 资管运营", BLUE),
    ("资产持有方", "资产定位 / 资本化", TEAL),
    ("企业选址方", "选址扩缩租决策", GOLD),
]
for i, (n, d, c) in enumerate(audiences):
    x = Inches(0.6 + i * 3.1)
    y = Inches(5.95)
    add_rect(s, x, y, Inches(3.0), Inches(0.5), c)
    add_text(s, x, y, Inches(3.0), Inches(0.5), n, size=12, bold=True,
             color=WHITE, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_rect(s, x, y + Inches(0.5), Inches(3.0), Inches(0.45),
             LIGHT_BLUE)
    add_text(s, x, y + Inches(0.5), Inches(3.0), Inches(0.45), d,
             size=11, color=DARK, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)


# ======================================================
# 3. 目录页
# ======================================================
s = content_slide("CONTENTS", "目录 · Outline",
                  "Five sections × Fourteen modules — 五个篇章·14个模块",
                  section_color=GOLD)
toc = [
    ("Part 01", "项目定位与背景", "联合发布·核心定位·五类读者·研究突破", BLUE),
    ("Part 02", "市场背景与逻辑重构", "城市更新·五个中心·企业选址逻辑变化", BLUE),
    ("Part 03", "全域供给格局", "16区·街镇·楼宇/园区·物业类型·重点板块", TEAL),
    ("Part 04", "租金、空置与价格体系", "分层租金·空置结构·综合承租成本", TEAL),
    ("Part 05", "入驻企业画像与产业需求", "行业结构·企业规模·重点清单", ORANGE),
    ("Part 06", "企业迁徙与产业流动", "8类迁徙·产业流向·区域吸引力", ORANGE),
    ("Part 07", "供需匹配与招商机会", "竞争力模型·三类招商机会", GOLD),
    ("Part 08", "市场趋势预测", "租金/空置/产业/空间策略", GOLD),
    ("Part 09", "政策建议与市场应用", "政府/园区/资产/企业 四类建议", PURPLE),
    ("Part 10", "成果体系与商业化", "月报/季报/白皮书/诊断/选址", PURPLE),
    ("Part 11", "数据体系（六大库）", "楼宇/租金/企业/迁徙/产业/政策", BLUE),
    ("Part 12", "核心指标（五大指数）", "景气/竞争力/需求/集聚/匹配", BLUE),
    ("Part 13", "数据调用与采集指引", "五项原则·四阶段·合规要点", TEAL),
    ("Part 14", "分工协同与首期落地", "三方分工·四类例会·试点先行", TEAL),
]
for i, (n, t, d, c) in enumerate(toc):
    col = i // 7
    row = i % 7
    x = Inches(0.6 + col * 6.4)
    y = Inches(1.85 + row * 0.72)
    add_text(s, x, y + Inches(0.04), Inches(0.75), Inches(0.35),
             n, size=10, bold=True, color=c)
    add_text(s, x + Inches(0.75), y, Inches(5.6), Inches(0.4),
             t, size=13, bold=True, color=NAVY)
    add_text(s, x + Inches(0.75), y + Inches(0.35),
             Inches(5.6), Inches(0.3),
             d, size=9, color=GRAY)
    # 装饰
    add_line(s, x, y + Inches(0.7), x + Inches(6.2), y + Inches(0.7),
             color=LINE, width=0.5)


# ======================================================
# 4. PART 1 隔页
# ======================================================
section_divider("01", "POSITIONING  &  BACKGROUND",
                "项目定位与背景",
                accent=BLUE,
                bullets=[
                    "联合发布主体：易居研究院 × 复旦大学住房政策研究中心",
                    "全域 / 全载体 / 双侧 / 动态 / 应用  五大研究突破",
                    "面向政府 · 园区 · 资产 · 企业 · 研究  五类读者",
                ])


# ======================================================
# 5. 联合发布与定位
# ======================================================
s = content_slide("PART 01 · 项目定位", "联合发布与核心定位",
                  "上海首份面向全域商办与产业园区的供需双侧·企业迁徙·产业流动决策报告")
# 左：联合发布
add_rect(s, Inches(0.6), Inches(1.85), Inches(5.6), Inches(2.6), LIGHT_BLUE)
add_rect(s, Inches(0.6), Inches(1.85), Inches(0.08), Inches(2.6), GOLD)
add_text(s, Inches(0.85), Inches(1.95), Inches(5.3), Inches(0.4),
         "JOINT  PUBLICATION", size=10, bold=True, color=NAVY)
add_text(s, Inches(0.85), Inches(2.3), Inches(5.3), Inches(0.4),
         "联合发布", size=14, bold=True, color=NAVY)
add_text(s, Inches(0.85), Inches(2.95), Inches(5.3), Inches(0.5),
         "易居房地产研究院", size=22, bold=True, color=NAVY,
         anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Inches(0.85), Inches(3.45), Inches(5.3), Inches(0.4),
         "×", size=20, bold=True, color=GOLD, anchor=MSO_ANCHOR.MIDDLE,
         align=PP_ALIGN.CENTER)
add_text(s, Inches(0.85), Inches(3.75), Inches(5.3), Inches(0.5),
         "复旦大学住房政策研究中心", size=22, bold=True, color=NAVY,
         anchor=MSO_ANCHOR.MIDDLE)
# 右：核心定位
add_text(s, Inches(6.5), Inches(1.95), Inches(6.3), Inches(0.4),
         "CORE  POSITIONING", size=10, bold=True, color=NAVY)
add_text(s, Inches(6.5), Inches(2.3), Inches(6.3), Inches(0.4),
         "核心定位", size=14, bold=True, color=NAVY)
add_bullets(s, Inches(6.5), Inches(2.8), Inches(6.3), Inches(1.7), [
    "全域覆盖：16区 → 街镇 → 楼宇/园区 → 企业 四级颗粒",
    "全类型载体：写字楼、产业园、科创园、孵化器、商务园",
    "供需双侧：供给底数 × 企业入驻 / 迁徙 / 扩缩租行为",
    "产业视角：集聚、流动、补链强链与新质生产力",
    "招商资管导向：政府 + 园区 + 资产 + 选址 一体应用",
], size=12, bullet_color=BLUE)

# 下：五大突破
add_text(s, Inches(0.6), Inches(4.7), Inches(12.0), Inches(0.4),
         "五大研究突破  ·  FIVE BREAKTHROUGHS", size=14, bold=True,
         color=NAVY)
add_rect(s, Inches(0.6), Inches(5.05), Inches(0.4), Inches(0.04), GOLD)
breaks = [
    ("全域", "核心区样本 → 上海全域", BLUE),
    ("全载体", "写字楼 → 全类型载体", TEAL),
    ("双侧", "租金空置 → 供需双侧", GREEN),
    ("动态", "静态描述 → 动态迁徙", ORANGE),
    ("应用", "宏观研究 → 招商资管选址", GOLD),
]
for i, (k, v, c) in enumerate(breaks):
    x = Inches(0.6 + i * 2.5)
    y = Inches(5.25)
    add_rect(s, x, y, Inches(2.35), Inches(1.6), WHITE, line=c)
    add_rect(s, x, y, Inches(2.35), Inches(0.45), c)
    add_text(s, x, y, Inches(2.35), Inches(0.45), k, size=14, bold=True,
             color=WHITE, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.15), y + Inches(0.55),
             Inches(2.05), Inches(0.95),
             v, size=10, color=DARK,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)


# ======================================================
# 6. 五类读者
# ======================================================
s = content_slide("PART 01 · 项目定位", "五类目标读者矩阵",
                  "Government · Operator · Capital · Enterprise · Research — 一份报告对应五个使用场景")
readers = [
    ("政府与产业招商", BLUE,
     ["市/区招商主管部门", "街镇招商中心", "开发区/功能区管委会"],
     "空间治理 · 精准招商 · 楼宇与园区经济政策"),
    ("楼宇/园区运营方", ORANGE,
     ["商办楼宇业主", "产业园区运营商", "城市更新操盘方", "国企平台公司"],
     "资产定位 · 租金策略 · 招商策略 · 运营服务"),
    ("资产持有与投资", GREEN,
     ["REITs / 类REITs资产方", "商办投资人", "资管公司", "地产基金"],
     "竞品对标 · 估值 · 改造 · 资本化路径"),
    ("企业选址决策方", GOLD,
     ["科创企业", "专精特新企业", "总部企业", "生产性服务业"],
     "区域比较 · 综合成本 · 政策匹配 · 扩缩租"),
    ("研究与咨询机构", PURPLE,
     ["房地产研究机构", "产业咨询机构", "金融研究部门", "城市更新服务商"],
     "标准化口径 · 行业研究 · 政策咨询"),
]
for i, (name, color, items, use) in enumerate(readers):
    col = i % 3
    row = i // 3
    x = Inches(0.6 + col * 4.21)
    y = Inches(1.95 + row * 2.55)
    add_rect(s, x, y, Inches(4.0), Inches(2.4), WHITE, line=LINE)
    add_rect(s, x, y, Inches(4.0), Inches(0.5), color)
    # 编号圆角
    add_oval(s, x + Inches(3.3), y + Inches(0.1), Inches(0.3),
             Inches(0.3), WHITE)
    add_text(s, x + Inches(3.3), y + Inches(0.1), Inches(0.3),
             Inches(0.3), f"0{i+1}", size=9, bold=True, color=color,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.18), y, Inches(3.5), Inches(0.5),
             name, size=14, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    add_bullets(s, x + Inches(0.2), y + Inches(0.65),
                Inches(3.7), Inches(1.25), items, size=11,
                bullet_color=color, bullet_char="·")
    add_rect(s, x, y + Inches(1.95), Inches(4.0), Inches(0.45), LIGHT_BLUE)
    add_text(s, x + Inches(0.15), y + Inches(1.95),
             Inches(3.7), Inches(0.45),
             "▎ " + use, size=10, color=NAVY,
             anchor=MSO_ANCHOR.MIDDLE)


# ======================================================
# 7. PART 02 隔页
# ======================================================
section_divider("02", "MARKET  CONTEXT",
                "市场背景与逻辑重构",
                bullets=[
                    "城市更新进入深水区  ·  存量盘活成为核心命题",
                    "新质生产力培育  ·  五个中心建设  ·  能级跃升",
                    "企业选址逻辑系统重构  ·  从地段优先转向生态优先",
                ])


# ======================================================
# 8. 市场背景
# ======================================================
s = content_slide("PART 02 · 市场背景", "上海商办与产业空间逻辑正在重构",
                  "From location-first to ecosystem-first — 七要素重构企业选址逻辑")
# 老逻辑 vs 新逻辑
add_rect(s, Inches(0.6), Inches(1.85), Inches(5.8), Inches(2.4), LIGHT_BLUE)
add_text(s, Inches(0.6), Inches(1.85), Inches(5.8), Inches(0.4),
         "  过去：地段 · 形象 · 总部标签", size=13, bold=True,
         color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
old = ["地段优先 / 形象优先", "核心区品牌办公", "高租金 = 高价值",
       "甲级写字楼为主流", "重资产 / 增量开发"]
add_bullets(s, Inches(0.85), Inches(2.3), Inches(5.4), Inches(1.95),
            old, size=12, bullet_color=GRAY, bullet_char="○",
            line_spacing=1.45)

# 箭头
arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(6.5), Inches(2.7),
                           Inches(0.6), Inches(0.7))
arrow.fill.solid(); arrow.fill.fore_color.rgb = GOLD
arrow.line.fill.background(); arrow.shadow.inherit = False

add_rect(s, Inches(7.2), Inches(1.85), Inches(5.6), Inches(2.4), NAVY)
add_text(s, Inches(7.2), Inches(1.85), Inches(5.6), Inches(0.4),
         "  现在：成本 · 产业 · 政策 · 人才", size=13, bold=True,
         color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
new = ["综合成本可控 · 政策可达", "产业协同 · 上下游生态",
       "空间灵活 · 服务完善", "交通便利 · 人才可达",
       "存量运营 · 产业服务 · 生态溢价"]
add_bullets(s, Inches(7.45), Inches(2.3), Inches(5.2), Inches(1.95),
            new, size=12, color=WHITE, bullet_color=GOLD, bullet_char="●",
            line_spacing=1.45)

# 四大不足
add_text(s, Inches(0.6), Inches(4.5), Inches(12.0), Inches(0.4),
         "传统研究体系的四大不足  ·  本报告的对应突破",
         size=14, bold=True, color=NAVY)
add_rect(s, Inches(0.6), Inches(4.85), Inches(0.4), Inches(0.04), GOLD)
issues = [
    ("重核心区，轻全域", "覆盖16区+街镇+楼宇/园区", BLUE),
    ("重资产指标，轻需求", "供需双侧+企业画像+行为研究", TEAL),
    ("重静态，轻动态", "月度/季度企业迁徙追踪", ORANGE),
    ("重描述，轻落地", "招商清单+资管+选址方案", GOLD),
]
for i, (a, b, c) in enumerate(issues):
    x = Inches(0.6 + i * 3.13)
    add_rect(s, x, Inches(5.1), Inches(3.0), Inches(0.5), WHITE, line=c)
    add_text(s, x, Inches(5.1), Inches(3.0), Inches(0.5), a,
             size=12, bold=True, color=c, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    arr = s.shapes.add_shape(MSO_SHAPE.DOWN_ARROW,
                             x + Inches(1.4), Inches(5.65),
                             Inches(0.22), Inches(0.25))
    arr.fill.solid(); arr.fill.fore_color.rgb = GOLD
    arr.line.fill.background(); arr.shadow.inherit = False
    add_rect(s, x, Inches(5.95), Inches(3.0), Inches(0.95), c)
    add_text(s, x + Inches(0.1), Inches(5.95), Inches(2.8),
             Inches(0.95), b, size=12, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)


# ======================================================
# 9. PART 03 隔页
# ======================================================
section_divider("03", "SUPPLY  LANDSCAPE",
                "全域供给格局",
                bullets=[
                    "上海 16 区  ·  20 大重点板块  ·  10 大物业类型",
                    "四级数据库  ·  从城市层级下沉到楼宇与企业",
                    "首次系统拆解全域商办与产业园区供给底数",
                ])


# ======================================================
# 10. 上海 16 区分组 + 重点指标
# ======================================================
s = content_slide("PART 03 · 全域供给", "上海 16 区供给格局总览",
                  "Four functional clusters — 核心商务 / 浦东主力 / 产业新兴 / 外溢承接")
# 6 大顶部指标
metrics = [
    ("商办楼宇", "全市数量盘点", BLUE),
    ("产业园区", "全市数量盘点", NAVY),
    ("总建筑面积", "存量底数", TEAL),
    ("可租赁面积", "市场流通", GOLD),
    ("可招商面积", "去化压力", ORANGE),
    ("新增供应", "供应节奏", PURPLE),
]
for i, (k, v, c) in enumerate(metrics):
    x = Inches(0.6 + i * 2.13)
    add_rect(s, x, Inches(1.85), Inches(2.0), Inches(1.05), c)
    add_text(s, x, Inches(1.85), Inches(2.0), Inches(0.45),
             k, size=12, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_text(s, x, Inches(2.3), Inches(2.0), Inches(0.55), v,
             size=10, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)

# 16 区组别 + 区列表
add_text(s, Inches(0.6), Inches(3.15), Inches(12), Inches(0.4),
         "16 区分组与代表板块", size=13, bold=True, color=NAVY)
add_rect(s, Inches(0.6), Inches(3.5), Inches(0.4), Inches(0.04), GOLD)

groups = [
    ("核心商务区", BLUE,
     [("黄浦", "外滩/南京东路"), ("静安", "南京西路/苏河湾"),
      ("徐汇", "徐家汇/滨江"), ("长宁", "虹桥/中山公园"),
      ("虹口", "北外滩")]),
    ("浦东主力区", NAVY,
     [("陆家嘴", "金融总部"), ("前滩", "总部展示"),
      ("张江", "硬科技/生医"), ("金桥", "智造/电子"),
      ("外高桥", "贸易"), ("临港", "硬科技/国际")]),
    ("产业新兴区", TEAL,
     [("杨浦", "在线新经济"), ("普陀", "数字经济"),
      ("闵行", "紫竹/虹桥"), ("嘉定", "智能汽车"),
      ("松江", "G60科创")]),
    ("外溢承接区", ORANGE,
     [("宝山", "南大智慧城"), ("青浦", "西虹桥"),
      ("奉贤", "东方美谷"), ("金山", "新材料"),
      ("崇明", "生态/海洋")]),
]
cy = Inches(3.7)
for i, (name, color, items) in enumerate(groups):
    x = Inches(0.6 + i * 3.13)
    add_rect(s, x, cy, Inches(3.0), Inches(0.4), color)
    add_text(s, x, cy, Inches(3.0), Inches(0.4), name, size=12,
             bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    box_h = Inches(0.5 * len(items) + 0.2)
    add_rect(s, x, cy + Inches(0.4), Inches(3.0), box_h, LIGHT_BLUE)
    for j, (d, b) in enumerate(items):
        y = cy + Inches(0.5 + j * 0.45)
        add_text(s, x + Inches(0.15), y, Inches(0.9), Inches(0.35),
                 d, size=11, bold=True, color=NAVY,
                 anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + Inches(1.0), y, Inches(1.95), Inches(0.35),
                 b, size=10, color=DARK,
                 anchor=MSO_ANCHOR.MIDDLE)


# ======================================================
# 11. 四级数据库结构
# ======================================================
s = content_slide("PART 03 · 全域供给", "四级精细化数据库结构",
                  "Four-level granularity — 行政区 / 街镇 / 楼宇园区 / 企业 ─ 报告核心差异化")
# 四级金字塔（横向流程）
levels = [
    ("L1", "行政区", "16 区", BLUE, "一级统计单元"),
    ("L2", "街道 / 镇", "200+ 街镇", TEAL, "二级精细化拆解"),
    ("L3", "楼宇 / 园区", "数千个载体", GOLD, "市场最小颗粒度"),
    ("L4", "入驻企业", "海量主体", ORANGE, "需求侧观察对象"),
]
y = Inches(2.0)
box_h = Inches(2.0)
for i, (code, name, scale, c, sub) in enumerate(levels):
    x = Inches(0.6 + i * 3.13)
    add_rect(s, x, y, Inches(2.85), box_h, WHITE, line=c)
    add_rect(s, x, y, Inches(2.85), Inches(0.4), c)
    add_text(s, x + Inches(0.12), y, Inches(0.8), Inches(0.4),
             code, size=12, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.85), y, Inches(2.0), Inches(0.4), name,
             size=13, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x, y + Inches(0.6), Inches(2.85), Inches(0.6), scale,
             size=20, bold=True, color=c, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    add_text(s, x, y + Inches(1.2), Inches(2.85), Inches(0.5), sub,
             size=10, color=GRAY, italic=True,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    if i < 3:
        arr = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                 x + Inches(2.88), y + Inches(0.85),
                                 Inches(0.22), Inches(0.3))
        arr.fill.solid(); arr.fill.fore_color.rgb = GOLD
        arr.line.fill.background(); arr.shadow.inherit = False

# 应用维度
add_text(s, Inches(0.6), Inches(4.3), Inches(12), Inches(0.4),
         "每一层级可输出的研究维度", size=13, bold=True, color=NAVY)
add_rect(s, Inches(0.6), Inches(4.65), Inches(0.4), Inches(0.04), GOLD)
dims = [
    ("行政区维度", "载体数量·总面积·租金水平·空置压力·主导产业·招商活跃度"),
    ("街镇维度", "载体清单·空置面积·平均租金·入驻企业数·产业集聚·供需匹配"),
    ("楼宇/园区维度", "竞争力评分·租金对标·客户画像·空置结构·招商建议"),
    ("企业维度", "行业·规模·阶段·迁徙·扩缩租·政策适配·选址需求"),
]
for i, (k, v) in enumerate(dims):
    y2 = Inches(4.9 + i * 0.48)
    add_rect(s, Inches(0.6), y2, Inches(2.5), Inches(0.42), NAVY)
    add_text(s, Inches(0.6), y2, Inches(2.5), Inches(0.42), k,
             size=12, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    add_rect(s, Inches(3.1), y2, Inches(9.7), Inches(0.42), LIGHT_BLUE)
    add_text(s, Inches(3.25), y2, Inches(9.4), Inches(0.42), v,
             size=11, color=DARK, anchor=MSO_ANCHOR.MIDDLE)


# ======================================================
# 12. 20 大重点板块网格
# ======================================================
s = content_slide("PART 03 · 全域供给", "20 大重点商务与产业板块",
                  "Twenty priority districts — 板块作为专项深度分析样本")
panels = [
    ("陆家嘴", "金融总部", NAVY), ("前滩", "总部展示", NAVY),
    ("张江科学城", "硬科技/生医", BLUE), ("临港新片区", "硬科技/国际", BLUE),
    ("徐汇滨江", "AI / 传媒", TEAL), ("北外滩", "金融 / 航运", TEAL),
    ("大虹桥", "国际商务", TEAL), ("杨浦滨江", "在线新经济", TEAL),
    ("五角场", "AI / 科创", GOLD), ("南京西路", "高端商务", GOLD),
    ("人民广场", "传统商务", GOLD), ("漕河泾", "信息技术", ORANGE),
    ("紫竹", "生命健康", ORANGE), ("金桥", "智造 / 通信", ORANGE),
    ("外高桥", "贸易 / 总部", ORANGE), ("嘉定汽车城", "智能汽车", PURPLE),
    ("G60科创走廊", "高端制造", PURPLE), ("青浦西虹桥", "数贸 / 进博", PURPLE),
    ("宝山南大智慧城", "新材料 / 数字", PURPLE), ("奉贤东方美谷", "美丽健康", PURPLE),
]
for i, (n, f, c) in enumerate(panels):
    col = i % 5
    row = i // 5
    x = Inches(0.6 + col * 2.55)
    y = Inches(1.95 + row * 1.05)
    add_rect(s, x, y, Inches(2.4), Inches(0.95), WHITE, line=c)
    add_rect(s, x, y, Inches(0.08), Inches(0.95), c)
    add_oval(s, x + Inches(2.05), y + Inches(0.12),
             Inches(0.22), Inches(0.22), c)
    add_text(s, x + Inches(2.05), y + Inches(0.12),
             Inches(0.22), Inches(0.22), str(i + 1).zfill(2),
             size=8, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.2), y + Inches(0.1),
             Inches(1.8), Inches(0.4),
             n, size=13, bold=True, color=NAVY)
    add_text(s, x + Inches(0.2), y + Inches(0.5),
             Inches(2.0), Inches(0.4),
             f, size=10, color=DARK)


# ======================================================
# 13. PART 04 隔页
# ======================================================
section_divider("04", "RENT  ·  VACANCY  ·  COST",
                "租金、空置与价格体系",
                bullets=[
                    "分层租金体系  ·  报价 / 中介 / 成交 三类标注",
                    "空置结构 8 维度  ·  招商可租面积为主口径",
                    "企业综合承租成本模型  ·  租金 + 五项 − 补贴",
                ])


# ======================================================
# 14. 租金/空置/成本
# ======================================================
s = content_slide("PART 04 · 租金空置", "租金 · 空置 · 综合承租成本",
                  "Three-layer rent quotes × Eight-dimension vacancy × Comprehensive cost model")
add_card(s, Inches(0.6), Inches(1.85), Inches(4.0), Inches(2.4),
         "分层租金体系", [
             "全市平均报价 / 成交租金",
             "核心 / 次核心 / 产业园 / 乙级以下",
             "16 区租金 + 物业等级租金",
             "报价 / 中介 / 真实成交（三类标注）",
         ], accent=BLUE, body_size=11)

add_card(s, Inches(4.75), Inches(1.85), Inches(4.0), Inches(2.4),
         "空置结构分析", [
             "总体空置率 + 各区/街镇 + 各物业",
             "小 / 中 / 大 · 整层 / 多层 / 整栋",
             "长期空置 vs 新增空置",
             "招商可租面积 = 报告主口径",
         ], accent=ORANGE, body_size=11)

add_card(s, Inches(8.9), Inches(1.85), Inches(3.9), Inches(2.4),
         "综合承租成本", [
             "= 租金 + 物业 + 停车",
             "  + 装修摊销 + 搬迁 + 通勤",
             "  − 政策补贴",
             "写字楼 vs 产业园对比 + 降本路径",
         ], accent=GREEN, body_size=11)

# 空置结构 8 维
add_text(s, Inches(0.6), Inches(4.4), Inches(12), Inches(0.4),
         "空置面积结构  ×  招商资管含义", size=13, bold=True, color=NAVY)
add_rect(s, Inches(0.6), Inches(4.75), Inches(0.4), Inches(0.04), GOLD)
structures = [
    ("小面积 ≤300㎡", "初创 / 小微团队", BLUE),
    ("中面积 300-1000㎡", "成长 / 中型", BLUE),
    ("大面积 1000-3000㎡", "总部 / 研发中心", TEAL),
    ("整层空置", "整层招商溢价", TEAL),
    ("多层连续 ≥2层", "总部 / 产业基地", ORANGE),
    ("整栋待招商", "产业主题改造", ORANGE),
    ("长期空置 ≥12月", "需重新定位 / 改造", RED),
    ("新增空置 3月", "短期波动信号", GOLD),
]
for i, (k, v, c) in enumerate(structures):
    col = i % 4
    row = i // 4
    x = Inches(0.6 + col * 3.13)
    y = Inches(5.0 + row * 0.92)
    add_rect(s, x, y, Inches(3.0), Inches(0.8), WHITE, line=LINE)
    add_rect(s, x, y, Inches(0.06), Inches(0.8), c)
    add_text(s, x + Inches(0.15), y + Inches(0.07),
             Inches(2.8), Inches(0.35),
             k, size=11, bold=True, color=NAVY)
    add_text(s, x + Inches(0.15), y + Inches(0.42),
             Inches(2.8), Inches(0.35),
             v, size=10, color=DARK)


# ======================================================
# 15. 综合成本公式可视化
# ======================================================
s = content_slide("PART 04 · 租金空置", "企业综合承租成本模型",
                  "Comprehensive Tenancy Cost Model — 不是租金最低，而是综合成本最优")
# 公式
formula_y = Inches(2.0)
formula_items = [
    ("租金", BLUE, "+"), ("物业费", BLUE, "+"), ("停车费", TEAL, "+"),
    ("装修摊销", TEAL, "+"), ("搬迁成本", ORANGE, "+"),
    ("通勤成本", ORANGE, "−"), ("政策补贴", GOLD, "="),
    ("综合承租成本", NAVY, "")
]
x_start = Inches(0.6)
w_item = Inches(1.45)
gap = Inches(0.05)
for i, (k, c, op) in enumerate(formula_items):
    x = x_start + (w_item + Inches(0.18)) * i
    add_rect(s, x, formula_y, w_item, Inches(0.8), WHITE, line=c)
    add_rect(s, x, formula_y, w_item, Inches(0.3), c)
    add_text(s, x, formula_y, w_item, Inches(0.3), k, size=11,
             bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    add_text(s, x, formula_y + Inches(0.3), w_item, Inches(0.5),
             "", size=11, color=DARK, anchor=MSO_ANCHOR.MIDDLE)
    if op:
        add_text(s, x + w_item, formula_y, Inches(0.18), Inches(0.8),
                 op, size=20, bold=True, color=GOLD,
                 anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)

# 应用分析
add_text(s, Inches(0.6), Inches(3.2), Inches(12), Inches(0.4),
         "可输出的对比与诊断", size=13, bold=True, color=NAVY)
add_rect(s, Inches(0.6), Inches(3.55), Inches(0.4), Inches(0.04), GOLD)
applies = [
    ("名义 vs 实际", "挂牌租金 vs 实际承租成本差异"),
    ("区域对比", "不同区域企业综合成本横向比较"),
    ("写字楼 vs 产业园", "两类载体的成本优势分析"),
    ("政策补贴影响", "补贴对选址决策的真实贡献"),
    ("降本迁移路径", "企业降本的最优迁移组合"),
]
for i, (k, v) in enumerate(applies):
    y2 = Inches(3.85 + i * 0.5)
    add_rect(s, Inches(0.6), y2, Inches(3.0), Inches(0.42), NAVY)
    add_text(s, Inches(0.6), y2, Inches(3.0), Inches(0.42), k,
             size=12, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    add_rect(s, Inches(3.6), y2, Inches(9.2), Inches(0.42), LIGHT_BLUE)
    add_text(s, Inches(3.75), y2, Inches(8.9), Inches(0.42), v,
             size=11, color=DARK, anchor=MSO_ANCHOR.MIDDLE)


# ======================================================
# 16. PART 05 隔页
# ======================================================
section_divider("05", "DEMAND  &  ENTERPRISE  PROFILES",
                "入驻企业画像与产业需求",
                bullets=[
                    "17 大重点行业  ·  10 类企业规模  ·  9 类重点清单",
                    "面积 / 租金 / 选址 / 政策 / 配套 五维需求矩阵",
                    "需求端不再是黑箱，可标签、可清单、可对接",
                ])


# ======================================================
# 17. 17 大重点行业
# ======================================================
s = content_slide("PART 05 · 企业画像", "17 大重点行业",
                  "Seventeen priority industries — 嵌入上海三大先导与新质生产力方向")
industries = [
    ("人工智能", BLUE), ("集成电路", BLUE), ("生物医药", BLUE),
    ("智能驾驶", TEAL), ("新能源汽车", TEAL), ("机器人", TEAL),
    ("低空经济", GREEN), ("软件信息", GREEN),
    ("金融服务", NAVY), ("专业服务", NAVY),
    ("文化传媒", ORANGE), ("数字贸易", ORANGE), ("跨境电商", ORANGE),
    ("高端制造", PURPLE), ("生产性服务业", PURPLE),
    ("总部经济", GOLD), ("专精特新", GOLD),
]
for i, (n, c) in enumerate(industries):
    col = i % 6
    row = i // 6
    x = Inches(0.6 + col * 2.13)
    y = Inches(1.95 + row * 0.85)
    add_rect(s, x, y, Inches(2.0), Inches(0.7), WHITE, line=c)
    add_rect(s, x, y, Inches(0.08), Inches(0.7), c)
    add_text(s, x + Inches(0.15), y, Inches(1.8), Inches(0.4),
             n, size=12, bold=True, color=NAVY,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.15), y + Inches(0.4),
             Inches(1.8), Inches(0.3),
             f"#{str(i+1).zfill(2)}", size=9, color=GRAY,
             italic=True)

# 行业需求矩阵
add_text(s, Inches(0.6), Inches(4.95), Inches(12), Inches(0.4),
         "典型行业空间需求矩阵（示例）", size=13, bold=True, color=NAVY)
add_rect(s, Inches(0.6), Inches(5.3), Inches(0.4), Inches(0.04), GOLD)
table_y = Inches(5.5)
headers = ["行业", "空间形态", "选址偏好", "关注因素"]
widths = [Inches(1.8), Inches(2.6), Inches(4.2), Inches(3.6)]
cx = Inches(0.6)
for i, h in enumerate(headers):
    add_rect(s, cx, table_y, widths[i], Inches(0.4), NAVY)
    add_text(s, cx, table_y, widths[i], Inches(0.4), h, size=11,
             bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    cx += widths[i]
rows = [
    ["人工智能", "中大型办公+研发展示", "徐汇滨江·杨浦·张江·临港", "人才/算力/资本/政策"],
    ["生物医药", "研发办公+实验空间", "张江·临港·闵行·奉贤", "实验/审批/产业链"],
    ["智能驾驶", "办公+研发+测试", "嘉定·浦东·杨浦·临港", "测试场景/整车/政策"],
]
for ri, row in enumerate(rows):
    cx = Inches(0.6)
    bg = LIGHT_BLUE if ri % 2 == 0 else WHITE
    for i, val in enumerate(row):
        add_rect(s, cx, table_y + Inches(0.4 + ri * 0.4),
                 widths[i], Inches(0.4), bg)
        add_text(s, cx + Inches(0.1), table_y + Inches(0.4 + ri * 0.4),
                 widths[i] - Inches(0.2), Inches(0.4),
                 val, size=10, color=DARK,
                 anchor=MSO_ANCHOR.MIDDLE)
        cx += widths[i]


# ======================================================
# 18. 9 类重点企业清单
# ======================================================
s = content_slide("PART 05 · 企业画像", "9 类重点企业清单体系",
                  "Nine priority lists — 报告商业化价值的核心产品形态")
lists = [
    ("高成长企业", "员工高速增长 + 融资 + 招聘暴增", BLUE),
    ("近期融资企业", "12 个月内完成新一轮融资", BLUE),
    ("专精特新企业", "经信委 / 工信部认定", TEAL),
    ("高新技术企业", "高新认定 + 三大研发条件", TEAL),
    ("上市 / 拟上市", "公告 + 招股书", GREEN),
    ("迁址可能企业", "多源信号 + 招聘异动", ORANGE),
    ("扩租潜力企业", "原址续约 + 业务公告", ORANGE),
    ("缩租风险企业", "营收下滑 / 招聘冻结", RED),
    ("跨区 / 补链企业", "外地企业上海布局 / 上下游", GOLD),
]
for i, (n, d, c) in enumerate(lists):
    col = i % 3
    row = i // 3
    x = Inches(0.6 + col * 4.21)
    y = Inches(1.95 + row * 1.55)
    add_rect(s, x, y, Inches(4.0), Inches(1.4), WHITE, line=LINE)
    add_rect(s, x, y, Inches(4.0), Inches(0.45), c)
    add_oval(s, x + Inches(3.45), y + Inches(0.08),
             Inches(0.3), Inches(0.3), WHITE)
    add_text(s, x + Inches(3.45), y + Inches(0.08),
             Inches(0.3), Inches(0.3), f"L{i+1}", size=9, bold=True,
             color=c, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.15), y, Inches(3.2), Inches(0.45),
             n, size=13, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.2), y + Inches(0.55),
             Inches(3.7), Inches(0.8),
             d, size=10, color=DARK, anchor=MSO_ANCHOR.MIDDLE)


# ======================================================
# 19. PART 06 隔页
# ======================================================
section_divider("06", "MIGRATION  &  INDUSTRY  FLOW",
                "企业迁徙与产业流动",
                bullets=[
                    "8 类迁徙行为  ·  多源信号交叉  ·  避免单一误判",
                    "重点产业流向地图  ·  AI / 智驾 / 生医 / 总部",
                    "9 项指标构建区域吸引力指数",
                ])


# ======================================================
# 20. 8 类迁徙类型
# ======================================================
s = content_slide("PART 06 · 企业迁徙", "8 类企业迁徙行为",
                  "Eight migration types — 每类迁徙背后的真实动因")
mtypes = [
    ("扩张型", "面积扩大 / 楼层升级 / 团队增加", GREEN),
    ("降本型", "迁出核心区 / 综合成本下降", ORANGE),
    ("总部升级型", "迁入高品质 / 滨江总部带", NAVY),
    ("产业集聚型", "向产业带集聚 / 上下游协同", BLUE),
    ("政策导向型", "政策窗口期内的迁移", GOLD),
    ("被动搬迁型", "楼宇改造 / 业主收回", GRAY),
    ("缩租型", "面积收缩 / 降本求生", RED),
    ("注册地变更", "注册地变 / 办公地不变", PURPLE),
]
for i, (n, d, c) in enumerate(mtypes):
    col = i % 4
    row = i // 4
    x = Inches(0.6 + col * 3.13)
    y = Inches(1.95 + row * 1.35)
    add_rect(s, x, y, Inches(3.0), Inches(1.2), WHITE, line=c)
    add_oval(s, x + Inches(0.15), y + Inches(0.15),
             Inches(0.55), Inches(0.55), c)
    add_text(s, x + Inches(0.15), y + Inches(0.15),
             Inches(0.55), Inches(0.55), f"M{i+1}", size=11, bold=True,
             color=WHITE, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.85), y + Inches(0.18),
             Inches(2.1), Inches(0.4),
             n, size=14, bold=True, color=NAVY)
    add_text(s, x + Inches(0.15), y + Inches(0.75),
             Inches(2.8), Inches(0.4),
             d, size=10, color=DARK)

# 校验提示
add_rect(s, Inches(0.6), Inches(4.85), Inches(12.2), Inches(1.0), NAVY)
add_text(s, Inches(0.8), Inches(4.95), Inches(12), Inches(0.4),
         "迁徙判断校验规则", size=14, bold=True, color=WHITE)
add_text(s, Inches(0.8), Inches(5.35), Inches(12), Inches(0.45),
         "工商地址变化  +  地图POI变化  +  招聘地址变化  +  官网/公众号公告  +  线下确认  —  至少两个以上信号交叉验证，避免误判",
         size=12, color=GOLD_SOFT, anchor=MSO_ANCHOR.MIDDLE)


# ======================================================
# 21. 产业迁徙流向 + 区域吸引力指数
# ======================================================
s = content_slide("PART 06 · 企业迁徙", "重点产业流向与区域吸引力",
                  "Industry flow map × Regional attractiveness index")
# 左侧：产业流向
add_text(s, Inches(0.6), Inches(1.85), Inches(6.0), Inches(0.4),
         "重点产业流向", size=13, bold=True, color=NAVY)
add_rect(s, Inches(0.6), Inches(2.2), Inches(0.4), Inches(0.04), GOLD)
flows = [
    ("人工智能", "→  徐汇滨江·杨浦·张江·临港", BLUE),
    ("智能驾驶", "→  嘉定·浦东·杨浦·临港", TEAL),
    ("生物医药", "→  张江·临港·闵行·奉贤", GREEN),
    ("集成电路", "→  张江·临港·嘉定", BLUE),
    ("金融/专业服务", "→  陆家嘴·北外滩·南京西路", NAVY),
    ("总部企业", "→  陆家嘴·前滩·徐汇滨江", GOLD),
    ("文化传媒", "→  徐汇滨江·杨浦滨江", ORANGE),
    ("硬科技/制造", "→  嘉定·松江·宝山·临港", PURPLE),
]
for i, (k, v, c) in enumerate(flows):
    y = Inches(2.4 + i * 0.5)
    add_rect(s, Inches(0.6), y, Inches(1.8), Inches(0.42), c)
    add_text(s, Inches(0.6), y, Inches(1.8), Inches(0.42),
             k, size=11, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_rect(s, Inches(2.4), y, Inches(4.5), Inches(0.42), LIGHT_BLUE)
    add_text(s, Inches(2.55), y, Inches(4.3), Inches(0.42),
             v, size=11, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

# 右侧：吸引力指数
add_text(s, Inches(7.2), Inches(1.85), Inches(6.0), Inches(0.4),
         "区域吸引力指数 · 9 项核心指标", size=13, bold=True, color=NAVY)
add_rect(s, Inches(7.2), Inches(2.2), Inches(0.4), Inches(0.04), GOLD)
inds = [
    ("企业净流入", BLUE), ("重点产业流入", BLUE),
    ("高成长企业流入", TEAL), ("租金性价比", TEAL),
    ("政策支持力度", ORANGE), ("交通可达性", ORANGE),
    ("产业配套成熟度", GOLD), ("空间供给适配", GOLD),
    ("资本服务集聚", PURPLE),
]
for i, (k, c) in enumerate(inds):
    col = i % 3
    row = i // 3
    x = Inches(7.2 + col * 2.0)
    y = Inches(2.4 + row * 1.0)
    add_rect(s, x, y, Inches(1.85), Inches(0.85), WHITE, line=c)
    add_oval(s, x + Inches(0.15), y + Inches(0.15),
             Inches(0.55), Inches(0.55), c)
    add_text(s, x + Inches(0.15), y + Inches(0.15),
             Inches(0.55), Inches(0.55), str(i + 1).zfill(2),
             size=11, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.75), y, Inches(1.05), Inches(0.85),
             k, size=10, bold=True, color=NAVY,
             anchor=MSO_ANCHOR.MIDDLE)


# ======================================================
# 22. PART 07 隔页
# ======================================================
section_divider("07", "MATCHING  &  OPPORTUNITIES",
                "供需匹配与招商机会",
                bullets=[
                    "楼宇竞争力 7 维模型  ·  产业园区竞争力扩展模型",
                    "三类招商机会：区域机会 · 产业机会 · 企业线索",
                    "形成可落地的招商清单与资产运营建议",
                ])


# ======================================================
# 23. 楼宇竞争力模型
# ======================================================
s = content_slide("PART 07 · 供需招商", "楼宇竞争力 7 维评分模型",
                  "Seven-dimension scoring — 可直接落地的资产诊断卡")
dims = [
    ("区位交通", 15, "地铁·主干路·高铁机场", BLUE),
    ("资产品质", 15, "建筑·层高·装修·智能化", BLUE),
    ("租金表现", 15, "报价·成交·弹性", TEAL),
    ("企业结构", 15, "质量·集中度·龙头", TEAL),
    ("运营服务", 15, "物业·企业服务·活动", ORANGE),
    ("政策资源", 10, "补贴·资质·招商", GOLD),
    ("去化能力", 15, "空置·周期·稳定性", PURPLE),
]
total_y = Inches(2.0)
# 评分条
for i, (k, w, d, c) in enumerate(dims):
    y = total_y + Inches(i * 0.55)
    add_rect(s, Inches(0.6), y, Inches(2.0), Inches(0.45), c)
    add_text(s, Inches(0.6), y, Inches(2.0), Inches(0.45),
             k, size=12, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_rect(s, Inches(2.6), y, Inches(1.0), Inches(0.45), LIGHT_BLUE)
    add_text(s, Inches(2.6), y, Inches(1.0), Inches(0.45),
             f"{w} 分", size=12, bold=True, color=NAVY,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    # 进度条
    add_rect(s, Inches(3.7), y + Inches(0.12), Inches(5.0),
             Inches(0.22), GRAY_LIGHT)
    bar_w = Inches(5.0 * w / 15)
    add_rect(s, Inches(3.7), y + Inches(0.12), bar_w, Inches(0.22), c)
    add_text(s, Inches(8.85), y, Inches(4.3), Inches(0.45),
             d, size=10, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

# 合计提示
add_rect(s, Inches(0.6), Inches(6.0), Inches(12.2), Inches(0.7), NAVY)
add_text(s, Inches(0.8), Inches(6.0), Inches(12), Inches(0.7),
         "合计 100 分  ·  支持横向竞品对比、纵向同区诊断、动态更新 ·  适配楼宇与园区双场景",
         size=13, bold=True, color=GOLD_SOFT,
         anchor=MSO_ANCHOR.MIDDLE)


# ======================================================
# 24. 三类招商机会
# ======================================================
s = content_slide("PART 07 · 供需招商", "三类招商机会",
                  "Three types of investment leads — 报告商业化价值的核心")
opps = [
    ("区域机会", BLUE, "Where",
     ["哪些区域适合承接核心区外溢",
      "哪些区域适合发展 AI / 智驾 / 生医",
      "哪些区域适合导入总部企业",
      "哪些区域适合做成本型办公承接"]),
    ("产业机会", ORANGE, "What",
     ["AI 产业链补链机会",
      "智能驾驶上下游企业导入",
      "生物医药研发服务企业导入",
      "专精特新企业集聚机会"]),
    ("企业线索", GOLD, "Who",
     ["近期融资企业",
      "新增招聘明显企业",
      "扩租 / 迁址可能企业",
      "产业链上下游企业"]),
]
for i, (n, c, en, items) in enumerate(opps):
    x = Inches(0.6 + i * 4.21)
    y = Inches(1.95)
    add_rect(s, x, y, Inches(4.0), Inches(4.5), WHITE, line=LINE)
    add_rect(s, x, y, Inches(4.0), Inches(1.0), c)
    add_text(s, x, y + Inches(0.1), Inches(4.0), Inches(0.4),
             en.upper(), size=10, color=GOLD_SOFT,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_text(s, x, y + Inches(0.45), Inches(4.0), Inches(0.5),
             n, size=20, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_bullets(s, x + Inches(0.3), y + Inches(1.2),
                Inches(3.6), Inches(3.2),
                items, size=12, bullet_color=c, bullet_char="●",
                line_spacing=1.4, space_after=6)


# ======================================================
# 25. PART 08 隔页
# ======================================================
section_divider("08", "TREND  FORECAST",
                "市场趋势预测",
                bullets=[
                    "租金 / 空置 / 产业 / 空间策略  四线预测",
                    "核心 / 滨江 / 张江临港 / 郊区 / 老旧更新  五大机会场域",
                    "面向未来 12-24 个月的市场判断",
                ])


# ======================================================
# 26. 趋势预测
# ======================================================
s = content_slide("PART 08 · 趋势预测", "市场趋势与未来机会",
                  "Four forecast lines × Five opportunity zones")
add_card(s, Inches(0.6), Inches(1.85), Inches(3.0), Inches(2.3),
         "租金趋势", [
             "核心商务区走势",
             "乙级写字楼承压",
             "产业园区韧性",
             "新兴板块成长",
             "老旧楼宇调整",
         ], accent=BLUE, body_size=10, title_size=12)
add_card(s, Inches(3.75), Inches(1.85), Inches(3.0), Inches(2.3),
         "空置趋势", [
             "高空置板块",
             "新增供应压力",
             "去化较快板块",
             "需求支撑较强",
             "存量改造压力",
         ], accent=ORANGE, body_size=10, title_size=12)
add_card(s, Inches(6.9), Inches(1.85), Inches(3.0), Inches(2.3),
         "产业需求趋势", [
             "AI · 智驾 · 机器人",
             "低空 · 生医 · 集电",
             "数贸 · 跨境电商",
             "绿色低碳 · 科技服务",
             "生产性服务业",
         ], accent=GREEN, body_size=10, title_size=12)
add_card(s, Inches(10.05), Inches(1.85), Inches(2.75), Inches(2.3),
         "空间策略", [
             "更灵活面积",
             "总部重品牌资源",
             "成长重政策",
             "硬科技重复合空间",
             "产业社区+服务",
         ], accent=GOLD, body_size=10, title_size=12)

# 五大机会
add_text(s, Inches(0.6), Inches(4.4), Inches(12), Inches(0.4),
         "未来五大机会场域", size=13, bold=True, color=NAVY)
add_rect(s, Inches(0.6), Inches(4.75), Inches(0.4), Inches(0.04), GOLD)
chances = [
    ("核心区", "总部经济·金融·专业服务·外资机构", NAVY),
    ("滨江区", "AI·数字经济·总部展示·科技服务", BLUE),
    ("张江/临港", "硬科技·生物医药·高端制造·智能驾驶", GREEN),
    ("郊区产业空间", "成本型外溢·研发制造结合·产业链集聚", ORANGE),
    ("老旧商办更新", "主题楼宇·垂直产业楼宇·科创服务空间", GOLD),
]
for i, (k, v, c) in enumerate(chances):
    y = Inches(5.0 + i * 0.42)
    add_rect(s, Inches(0.6), y, Inches(2.2), Inches(0.38), c)
    add_text(s, Inches(0.6), y, Inches(2.2), Inches(0.38),
             k, size=11, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_rect(s, Inches(2.8), y, Inches(10.0), Inches(0.38), LIGHT_BLUE)
    add_text(s, Inches(2.95), y, Inches(9.8), Inches(0.38),
             v, size=11, color=DARK, anchor=MSO_ANCHOR.MIDDLE)


# ======================================================
# 27. PART 09 隔页
# ======================================================
section_divider("09", "POLICY  &  APPLICATION",
                "政策建议与市场应用",
                bullets=[
                    "面向政府 / 园区 / 资产 / 企业  四类对象建议",
                    "可直接转化为政策语言、招商话术与资管动作",
                ])


# ======================================================
# 28. 四类建议
# ======================================================
s = content_slide("PART 09 · 政策应用", "面向四类对象的策略建议",
                  "Policy · Operator · Capital · Enterprise")
sugg = [
    ("对政府部门", NAVY, [
        "建立全域空间动态数据库",
        "以街镇为单位监测空置与迁徙",
        "对高空置区域实施精准招商",
        "推动楼宇 / 园区 / 城市更新联动",
        "用企业迁徙数据优化产业政策",
    ]),
    ("对园区/楼宇运营方", BLUE, [
        "从“租赁招商”转向“产业招商”",
        "建立企业画像与招商漏斗",
        "动态监测竞品租金和空置",
        "引入企业服务·政策·资本·产业链资源",
        "对老旧载体进行场景化改造",
    ]),
    ("对资产持有方", ORANGE, [
        "重新评估资产定位",
        "建立竞品对标体系",
        "关注真实成交而非挂牌租金",
        "通过产业主题提升资产溢价",
        "为 REITs / 资本化储备数据基础",
    ]),
    ("对企业选址方", GOLD, [
        "建立综合承租成本模型",
        "综合比较政策/产业/人才",
        "根据企业阶段选择空间类型",
        "成长企业关注扩租弹性",
        "总部企业关注品牌与资源",
    ]),
]
for i, (name, c, items) in enumerate(sugg):
    col = i % 2
    row = i // 2
    x = Inches(0.6 + col * 6.4)
    y = Inches(1.95 + row * 2.5)
    add_card(s, x, y, Inches(6.2), Inches(2.3), name, items,
             accent=c, body_size=11, title_size=14)


# ======================================================
# 29. PART 10 隔页
# ======================================================
section_divider("10", "DELIVERABLES  MATRIX",
                "成果体系与商业化产品",
                bullets=[
                    "月报 / 季报 / 白皮书 / 区域专项 / 楼宇诊断 / 选址服务",
                    "从一份报告升级为长期产品矩阵与研究品牌",
                ])


# ======================================================
# 30. 成果矩阵
# ======================================================
s = content_slide("PART 10 · 成果矩阵", "成果体系与商业化产品矩阵",
                  "Report + Database + Visualization + Productized services")
products = [
    ("月度市场监测简报", "月", "租金 / 空置 / 迁徙 / 招商热点 / 预警", BLUE),
    ("季度深度报告", "季", "区域 / 产业 / 迁徙 / 资管 / 趋势", NAVY),
    ("年度白皮书", "年", "全域格局 + 年度趋势 + 政策建议", GOLD),
    ("区域专项报告", "按需", "杨浦 / 浦东 / 徐汇滨江 / 嘉定 / 临港", ORANGE),
    ("楼宇 / 园区诊断", "按项目", "竞品 / 租金 / 空置 / 招商 / 改造", GREEN),
    ("企业选址服务", "按企业", "区域 + 楼宇 + 政策 + 成本", PURPLE),
]
for i, (n, f, c, color) in enumerate(products):
    col = i % 3
    row = i // 3
    x = Inches(0.6 + col * 4.21)
    y = Inches(1.95 + row * 2.5)
    add_rect(s, x, y, Inches(4.0), Inches(2.3), WHITE, line=LINE)
    add_rect(s, x, y, Inches(4.0), Inches(0.55), color)
    add_text(s, x + Inches(0.15), y, Inches(2.6), Inches(0.55),
             n, size=14, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x + Inches(2.85), y + Inches(0.1),
             Inches(1.0), Inches(0.35), GOLD)
    add_text(s, x + Inches(2.85), y + Inches(0.1),
             Inches(1.0), Inches(0.35), f, size=10, bold=True,
             color=NAVY_DARK, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.2), y + Inches(0.7),
             Inches(3.7), Inches(1.5),
             c, size=12, color=DARK, anchor=MSO_ANCHOR.MIDDLE)


# ======================================================
# 31. PART 11 隔页
# ======================================================
section_divider("11", "DATA  ARCHITECTURE",
                "数据体系（六大基础数据库）",
                bullets=[
                    "楼宇/园区  ·  租金/空置  ·  企业入驻",
                    "企业迁徙  ·  产业标签  ·  政策配套",
                ])


# ======================================================
# 32. 六大数据库
# ======================================================
s = content_slide("PART 11 · 数据体系", "六大基础数据库",
                  "Six core databases — 全域 / 动态 / 可校验 / 可落地的研究底座")
dbs = [
    ("楼宇与园区基础库", BLUE, "21 字段",
     "名称 / 类型 / 区街镇 / 地址 / 面积 / 层数 / 物业等级 / 产权 / 运营 / 招商"),
    ("租金与空置库", ORANGE, "15 字段",
     "报价 / 成交 / 物业费 / 可租 / 空置 / 楼层 / 装修 / 免租 / 付款"),
    ("企业入驻库", GREEN, "17 字段",
     "名称 / 信用代码 / 注册 / 办公 / 楼宇 / 行业 / 规模 / 融资 / 资质"),
    ("企业迁徙库", GOLD, "14 字段",
     "原/新 地址·楼宇·街镇 / 时间 / 类型 / 面积 / 租金 / 行业 / 规模"),
    ("产业标签库", NAVY, "15 字段",
     "三级行业 / 战新 / 重点 / 新质 / 产业链 / 产品 / 客户 / 专利 / 资质"),
    ("政策与配套库", PURPLE, "15 字段",
     "区/街镇政策 / 租金/装修/人才补贴 / 税收 / 基金 / 落户 / 交通商医"),
]
for i, (n, c, num, d) in enumerate(dbs):
    col = i % 2
    row = i // 2
    x = Inches(0.6 + col * 6.4)
    y = Inches(1.95 + row * 1.6)
    add_rect(s, x, y, Inches(6.2), Inches(1.45), WHITE, line=LINE)
    add_rect(s, x, y, Inches(6.2), Inches(0.45), c)
    add_oval(s, x + Inches(5.3), y + Inches(0.07),
             Inches(0.85), Inches(0.32), GOLD)
    add_text(s, x + Inches(5.3), y + Inches(0.07),
             Inches(0.85), Inches(0.32), num, size=10, bold=True,
             color=NAVY_DARK, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.2), y, Inches(5.0), Inches(0.45), n,
             size=14, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.25), y + Inches(0.55),
             Inches(5.9), Inches(0.85), d,
             size=11, color=DARK, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.3)


# ======================================================
# 33. 五大指数
# ======================================================
s = content_slide("PART 12 · 指标体系", "核心指标体系：五大指数",
                  "Five core indices — 让研究可量化、可对比、可发布")
indices = [
    ("区域市场景气指数", BLUE, "判断区域冷热·短期波动·预警",
     ["租金变化", "空置率变化", "新增成交面积",
      "新增企业数量", "企业净流入", "去化速度"]),
    ("楼宇/园区竞争力", NAVY, "横向对标·资产诊断",
     ["区位交通", "租金性价比", "空间品质", "产业集聚",
      "企业质量", "服务能力", "政策资源", "去化表现"]),
    ("企业需求热度", ORANGE, "需求强弱·增长方向·线索",
     ["新注册企业", "融资企业", "招聘岗位增长",
      "扩租企业", "新租企业", "重点产业活跃"]),
    ("产业集聚度", GREEN, "主导产业·集群成熟度",
     ["同行业数量", "龙头数量", "上下游完整度",
      "专精特新", "高新数量", "协同程度"]),
    ("供需匹配度", GOLD, "楼宇适配企业·园区适配产业",
     ["面积匹配", "租金匹配", "产业匹配",
      "政策匹配", "使用场景匹配"]),
]
for i, (n, c, u, items) in enumerate(indices):
    col = i % 3
    row = i // 3
    x = Inches(0.6 + col * 4.21)
    y = Inches(1.95 + row * 2.55)
    add_rect(s, x, y, Inches(4.0), Inches(2.4), WHITE, line=LINE)
    add_rect(s, x, y, Inches(4.0), Inches(0.5), c)
    add_text(s, x + Inches(0.2), y, Inches(3.6), Inches(0.5),
             n, size=13, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.2), y + Inches(0.55),
             Inches(3.7), Inches(0.3),
             "▎ " + u, size=10, color=GOLD, italic=True)
    add_bullets(s, x + Inches(0.25), y + Inches(0.9),
                Inches(3.6), Inches(1.45),
                items, size=10, bullet_color=c, bullet_char="·",
                line_spacing=1.25, space_after=2)


# ======================================================
# 34. 数据调用与采集（含路线图）
# ======================================================
s = content_slide("PART 13 · 数据调用", "数据调用原则与四阶段采集路线图",
                  "Five principles × Four-phase data collection roadmap")
# 五原则
add_text(s, Inches(0.6), Inches(1.85), Inches(6.0), Inches(0.4),
         "五项数据调用原则", size=13, bold=True, color=NAVY)
add_rect(s, Inches(0.6), Inches(2.2), Inches(0.4), Inches(0.04), GOLD)
princ = [
    ("① 公开合规", "公开/授权/合作/调研为主"),
    ("② 多源交叉", "关键字段 ≥ 2 来源验证"),
    ("③ 动态更新", "建立月度更新机制"),
    ("④ 分级可信", "A/B/C/D 四级可信度"),
    ("⑤ 人工校验", "AI 提效 + 关键样本人工"),
]
for i, (k, v) in enumerate(princ):
    y = Inches(2.4 + i * 0.55)
    add_rect(s, Inches(0.6), y, Inches(1.7), Inches(0.48), NAVY)
    add_text(s, Inches(0.6), y, Inches(1.7), Inches(0.48), k, size=11,
             bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    add_rect(s, Inches(2.3), y, Inches(4.4), Inches(0.48), LIGHT_BLUE)
    add_text(s, Inches(2.45), y, Inches(4.2), Inches(0.48), v,
             size=11, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

# 四阶段 - 时间线样式
add_text(s, Inches(6.9), Inches(1.85), Inches(6.0), Inches(0.4),
         "四阶段采集路线图", size=13, bold=True, color=NAVY)
add_rect(s, Inches(6.9), Inches(2.2), Inches(0.4), Inches(0.04), GOLD)

phases = [
    ("Phase 1", "中心城区试点", "黄静徐长普虹杨 7 区", BLUE),
    ("Phase 2", "浦东新区补全", "陆家嘴/张江/临港...", TEAL),
    ("Phase 3", "重点产业区", "闵嘉松青宝奉 + 重点板块", ORANGE),
    ("Phase 4", "16区全域", "上海全域动态数据库", GOLD),
]
ph_y = Inches(2.4)
# 主线
add_rect(s, Inches(7.0), ph_y + Inches(0.95), Inches(5.8),
         Inches(0.06), NAVY)
for i, (p, n, d, c) in enumerate(phases):
    x = Inches(6.95 + i * 1.45)
    # 圆点
    add_oval(s, x + Inches(0.55), ph_y + Inches(0.85),
             Inches(0.25), Inches(0.25), c)
    # 上方信息
    add_text(s, x, ph_y, Inches(1.4), Inches(0.3),
             p, size=10, bold=True, color=c,
             align=PP_ALIGN.CENTER)
    add_text(s, x, ph_y + Inches(0.3), Inches(1.4), Inches(0.35),
             n, size=11, bold=True, color=NAVY,
             align=PP_ALIGN.CENTER)
    add_text(s, x, ph_y + Inches(0.6), Inches(1.4), Inches(0.3),
             "▽", size=10, color=c, align=PP_ALIGN.CENTER)
    # 下方
    add_rect(s, x + Inches(0.05), ph_y + Inches(1.3), Inches(1.3),
             Inches(0.7), LIGHT_BLUE)
    add_text(s, x + Inches(0.1), ph_y + Inches(1.3),
             Inches(1.2), Inches(0.7),
             d, size=9, color=DARK,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)

# 难点提示条
add_rect(s, Inches(0.6), Inches(5.65), Inches(12.2), Inches(1.25), NAVY)
add_text(s, Inches(0.8), Inches(5.7), Inches(12), Inches(0.4),
         "五大难点 → 解决方案", size=13, bold=True, color=GOLD_SOFT)
diffs = [
    "实际办公地址 → 工商+POI+招聘+官网+水牌+物业  多源交叉",
    "成交租金难获取 → 挂牌/中介/成交 三类标注，访谈积累",
    "空置口径不统一 → 招商可租口径为主，物理/隐性辅助",
    "迁徙易误判 → 至少两个信号交叉，线下确认",
    "数据合规 → 公开优先，统计汇总，发布前合规审查",
]
for i, d in enumerate(diffs):
    add_text(s, Inches(0.85), Inches(6.1 + i * 0.16), Inches(12), Inches(0.16),
             f"●  {d}", size=10, color=WHITE)


# ======================================================
# 35. 三方分工与首期落地
# ======================================================
s = content_slide("PART 14 · 分工落地", "三方协同 × 四类例会 × 首期试点",
                  "Joint workflow & pilot rollout")
roles = [
    ("易居研究院", BLUE, [
        "研究框架 / 统计口径",
        "市场指标 / 区域板块研究",
        "资管 / 招商策略",
        "线下访谈调研",
        "报告撰写 / 可视化 / 发布",
    ]),
    ("复旦大学住房政策研究中心", NAVY, [
        "城市更新与空间治理理论",
        "政策评价体系",
        "公共政策视角与合规把关",
        "学术背书 / 专家研讨",
        "白皮书及政策建议联合撰写",
    ]),
    ("数据技术与 AI 采集方", GOLD, [
        "POI/AOI 采集",
        "工商 / 招聘 / 公告抓取",
        "企业地址识别匹配",
        "迁徙线索识别 / 标准化",
        "数据库搭建 / 接口",
    ]),
]
for i, (n, c, items) in enumerate(roles):
    x = Inches(0.6 + i * 4.21)
    add_card(s, x, Inches(1.95), Inches(4.0), Inches(2.4), n, items,
             accent=c, body_size=11, title_size=13)

# 例会
add_text(s, Inches(0.6), Inches(4.55), Inches(12), Inches(0.4),
         "四类联合工作机制", size=13, bold=True, color=NAVY)
add_rect(s, Inches(0.6), Inches(4.9), Inches(0.4), Inches(0.04), GOLD)
meets = [
    ("周度", "数据进度会", "字段/样本/异常"),
    ("双周", "研究模型评审", "口径/指标/算法"),
    ("月度", "报告选题会", "区域/产业/案例"),
    ("季度", "成果复盘会", "准确性/反馈/规划"),
]
for i, (f, n, c) in enumerate(meets):
    x = Inches(0.6 + i * 3.13)
    add_rect(s, x, Inches(5.1), Inches(3.0), Inches(0.48), BLUE)
    add_text(s, x, Inches(5.1), Inches(3.0), Inches(0.48),
             f"{f} · {n}", size=12, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_rect(s, x, Inches(5.58), Inches(3.0), Inches(0.48), LIGHT_BLUE)
    add_text(s, x, Inches(5.58), Inches(3.0), Inches(0.48), c,
             size=11, color=DARK, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)

# 首期试点
add_rect(s, Inches(0.6), Inches(6.2), Inches(12.2), Inches(0.85), NAVY)
add_oval(s, Inches(0.8), Inches(6.35), Inches(0.55), Inches(0.55), GOLD)
add_text(s, Inches(0.8), Inches(6.35), Inches(0.55), Inches(0.55), "首",
         size=14, bold=True, color=NAVY_DARK,
         anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
add_text(s, Inches(1.55), Inches(6.25), Inches(11.2), Inches(0.4),
         "首期落地 · 《上海中心城区商办楼宇与产业园区市场试点报告》",
         size=14, bold=True, color=WHITE)
add_text(s, Inches(1.55), Inches(6.6), Inches(11.2), Inches(0.4),
         "覆盖：黄浦 · 静安 · 徐汇 · 长宁 · 普陀 · 虹口 · 杨浦  —  完成字段、口径、迁徙模型试点",
         size=11, color=GOLD_SOFT)


# ======================================================
# 36. 关键时间节点（路线图）
# ======================================================
s = content_slide("ROADMAP", "项目研发路线图",
                  "Stage-gate development roadmap")
stages = [
    ("S1 · 启动期", "课题立项 · 字段冻结 · 工具链准备 · 试点区选定",
     BLUE),
    ("S2 · 试点期", "中心城区试点 · 数据采集 · 模型测试 · 首份试点报告",
     TEAL),
    ("S3 · 扩展期", "浦东补全 + 重点产业区补全 · 月报上线 · 季报体系",
     ORANGE),
    ("S4 · 全域期", "16 区全域数据库 · 年度白皮书 · 区域专项 · 商业化产品矩阵",
     GOLD),
    ("S5 · 长期", "动态监测 · 看板 · 政策咨询 · 资管诊断 · 选址服务",
     PURPLE),
]
add_rect(s, Inches(1.0), Inches(3.7), Inches(11.4), Inches(0.08), NAVY)
for i, (name, desc, c) in enumerate(stages):
    x = Inches(0.6 + i * 2.55)
    box_w = Inches(2.35)
    cy = Inches(3.65)
    add_oval(s, x + Inches(1.1), cy, Inches(0.3), Inches(0.3), c)
    add_text(s, x + Inches(1.1), cy, Inches(0.3), Inches(0.3),
             "", anchor=MSO_ANCHOR.MIDDLE)
    if i % 2 == 0:
        # 上方
        add_rect(s, x, Inches(2.0), box_w, Inches(1.5), WHITE, line=c)
        add_rect(s, x, Inches(2.0), box_w, Inches(0.4), c)
        add_text(s, x, Inches(2.0), box_w, Inches(0.4), name,
                 size=11, bold=True, color=WHITE,
                 anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
        add_text(s, x + Inches(0.15), Inches(2.45),
                 box_w - Inches(0.25), Inches(1.0),
                 desc, size=10, color=DARK,
                 anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
        # 连接线
        add_line(s, x + Inches(1.25), Inches(3.5),
                 x + Inches(1.25), Inches(3.65),
                 color=c, width=1.5)
    else:
        add_rect(s, x, Inches(4.0), box_w, Inches(1.5), WHITE, line=c)
        add_rect(s, x, Inches(4.0), box_w, Inches(0.4), c)
        add_text(s, x, Inches(4.0), box_w, Inches(0.4), name,
                 size=11, bold=True, color=WHITE,
                 anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
        add_text(s, x + Inches(0.15), Inches(4.45),
                 box_w - Inches(0.25), Inches(1.0),
                 desc, size=10, color=DARK,
                 anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
        add_line(s, x + Inches(1.25), Inches(3.95),
                 x + Inches(1.25), Inches(4.0),
                 color=c, width=1.5)

# 底部价值定位
add_rect(s, Inches(0.6), Inches(5.9), Inches(12.2), Inches(1.0), NAVY)
add_text(s, Inches(0.8), Inches(5.95), Inches(12), Inches(0.4),
         "长期愿景", size=12, bold=True, color=GOLD_SOFT)
add_text(s, Inches(0.8), Inches(6.3), Inches(12), Inches(0.55),
         "将该项目打造为 — 上海商办楼宇与产业园区市场年度旗舰研究产品 + 行业标准化研究样本 + 城市空间治理决策参考",
         size=12, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)


# ======================================================
# 37. 封底
# ======================================================
s = new_blank()
add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY_DARK)
add_rect(s, 0, 0, SLIDE_W, Inches(0.18), GOLD)
add_rect(s, 0, SLIDE_H - Inches(0.18), SLIDE_W, Inches(0.18), GOLD)
# 大装饰
add_rect(s, Inches(0.6), Inches(1.2), Inches(0.06), Inches(5.2), GOLD)

add_text(s, Inches(0.9), Inches(1.4), Inches(12), Inches(0.5),
         "THANKS  FOR  YOUR  ATTENTION", size=14, bold=True,
         color=GOLD_SOFT)
add_text(s, Inches(0.9), Inches(2.0), Inches(12), Inches(1.4),
         "感谢聆听 · 欢迎共建", size=64, bold=True, color=WHITE,
         line_spacing=1.0)
# 装饰线
add_rect(s, Inches(0.95), Inches(3.7), Inches(0.6), Inches(0.06), GOLD)

add_text(s, Inches(0.9), Inches(4.0), Inches(12), Inches(0.5),
         "上海商办楼宇与产业园区市场研究课题组", size=22, bold=True,
         color=WHITE)
add_text(s, Inches(0.9), Inches(4.55), Inches(12), Inches(0.45),
         "易居房地产研究院  ×  复旦大学住房政策研究中心", size=18,
         color=GOLD_SOFT)

# 联系信息样
add_rect(s, Inches(0.9), Inches(5.6), Inches(11.5), Inches(0.06), GOLD)
contacts = [
    ("项目主体", "易居研究院 × 复旦大学住房政策研究中心"),
    ("发布形式", "月度简报 + 季度报告 + 年度白皮书"),
    ("适用对象", "政府 · 园区 · 资产 · 企业 · 研究机构"),
]
for i, (k, v) in enumerate(contacts):
    x = Inches(0.9 + i * 3.85)
    add_text(s, x, Inches(5.85), Inches(3.6), Inches(0.3),
             k, size=10, color=GOLD_SOFT)
    add_text(s, x, Inches(6.2), Inches(3.6), Inches(0.4),
             v, size=11, bold=True, color=WHITE)

add_text(s, Inches(0.9), SLIDE_H - Inches(0.6), Inches(12), Inches(0.3),
         "— 讨论稿 · 仅供联合课题组内部研讨使用 —", size=11,
         color=GRAY_LIGHT, italic=True)


out = os.path.join(OUTPUT_DIR, "上海商办楼宇与产业园区市场深度报告_汇报版.pptx")
prs.save(out)
print(f"PPT 已生成：{out}  ·  共 {len(prs.slides)} 页")
