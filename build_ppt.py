"""
Generate the "十五五" 房地产高质量发展 presentation deck.

Theme: modern, restrained, policy / academic-friendly
Palette: deep navy + gold accent + soft neutrals
Aspect: 16:9 widescreen
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# -------------------- Theme --------------------
NAVY        = RGBColor(0x0B, 0x24, 0x47)   # primary deep navy
NAVY_DARK   = RGBColor(0x07, 0x1A, 0x33)
NAVY_SOFT   = RGBColor(0x19, 0x3A, 0x6B)
GOLD        = RGBColor(0xC9, 0xA2, 0x27)   # accent gold
GOLD_LIGHT  = RGBColor(0xE6, 0xC2, 0x4A)
CRIMSON     = RGBColor(0xB3, 0x1B, 0x1B)   # subtle Chinese-policy red accent
INK         = RGBColor(0x1B, 0x1F, 0x27)
TEXT        = RGBColor(0x2C, 0x33, 0x40)
SUBTEXT     = RGBColor(0x55, 0x5E, 0x6C)
LINE        = RGBColor(0xCF, 0xD4, 0xDC)
PAPER       = RGBColor(0xF7, 0xF5, 0xEF)   # warm off-white
PAPER_DEEP  = RGBColor(0xEC, 0xE6, 0xD8)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
TEAL        = RGBColor(0x1F, 0x6F, 0x6B)

CN_FONT      = "Microsoft YaHei"
CN_FONT_BOLD = "Microsoft YaHei"
EN_FONT      = "Calibri"

# -------------------- Presentation setup --------------------
prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

# -------------------- Helpers --------------------
def set_run_font(run, size=18, bold=False, color=TEXT, name=CN_FONT):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    # Ensure East Asian font is set on the rPr
    rPr = run._r.get_or_add_rPr()
    for tag in ("ea", "cs"):
        el = rPr.find(qn(f"a:{tag}"))
        if el is None:
            from lxml import etree
            el = etree.SubElement(rPr, qn(f"a:{tag}"))
        el.set("typeface", name)


def add_text(slide, left, top, width, height, text,
             size=18, bold=False, color=TEXT, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, font=CN_FONT, line_spacing=1.25):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    lines = text if isinstance(text, list) else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = line
        set_run_font(r, size=size, bold=bold, color=color, name=font)
    return tb


def add_rect(slide, left, top, width, height, fill=NAVY, line=None, shape=MSO_SHAPE.RECTANGLE):
    shp = slide.shapes.add_shape(shape, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.5)
    shp.shadow.inherit = False
    return shp


def add_line(slide, x1, y1, x2, y2, color=GOLD, weight=1.5):
    from pptx.util import Emu
    ln = slide.shapes.add_connector(1, x1, y1, x2, y2)
    ln.line.color.rgb = color
    ln.line.width = Pt(weight)
    return ln


def page_background(slide, paper=True):
    add_rect(slide, 0, 0, SW, SH, fill=PAPER if paper else WHITE)


def header_bar(slide, kicker, title, page_num, total=18):
    # Top thin gold bar
    add_rect(slide, 0, 0, SW, Inches(0.18), fill=GOLD)
    # Navy band
    add_rect(slide, 0, Inches(0.18), SW, Inches(0.92), fill=NAVY)
    # Kicker (small label)
    add_text(slide, Inches(0.6), Inches(0.28), Inches(6), Inches(0.32),
             kicker, size=11, color=GOLD_LIGHT, bold=True)
    # Main title
    add_text(slide, Inches(0.6), Inches(0.55), Inches(11), Inches(0.55),
             title, size=22, color=WHITE, bold=True)
    # Page indicator (right)
    add_text(slide, Inches(11.6), Inches(0.4), Inches(1.2), Inches(0.5),
             f"{page_num:02d} / {total:02d}", size=11, color=GOLD_LIGHT,
             bold=True, align=PP_ALIGN.RIGHT)


def footer(slide, text="上海市房产经济学会 · 中青年学术研讨会  |  2026.05.27"):
    add_text(slide, Inches(0.6), Inches(7.05), Inches(12), Inches(0.35),
             text, size=10, color=SUBTEXT)
    add_line(slide, Inches(0.6), Inches(7.02), Inches(12.73), Inches(7.02),
             color=LINE, weight=0.75)


def bullet_block(slide, left, top, width, items, size=15, color=TEXT,
                 bullet_color=GOLD, line_spacing=1.4, gap=Inches(0.05),
                 heading_h=Inches(0.40), body_h=Inches(0.50)):
    """Render bullets: items is a list of (heading, body) or strings."""
    cur_top = top
    for item in items:
        if isinstance(item, tuple):
            head, body = item
        else:
            head, body = item, None
        add_rect(slide, left, cur_top + Inches(0.10),
                 Inches(0.12), Inches(0.12),
                 fill=bullet_color, shape=MSO_SHAPE.OVAL)
        add_text(slide, left + Inches(0.28), cur_top,
                 width - Inches(0.3), heading_h,
                 head, size=size+1, bold=True, color=NAVY)
        next_top = cur_top + heading_h + Inches(0.02)
        if body:
            add_text(slide, left + Inches(0.28), next_top,
                     width - Inches(0.3), body_h,
                     body, size=size-1, color=TEXT, line_spacing=line_spacing)
            next_top += body_h
        cur_top = next_top + gap
    return cur_top


# ====================================================================
# Slide 1 — Cover
# ====================================================================
def slide_cover():
    s = prs.slides.add_slide(BLANK)
    # Solid navy background
    add_rect(s, 0, 0, SW, SH, fill=NAVY)
    # Decorative geometric shapes
    add_rect(s, Inches(9.5), 0, Inches(3.83), SH, fill=NAVY_DARK)
    # Gold vertical accent
    add_rect(s, Inches(9.3), 0, Inches(0.08), SH, fill=GOLD)
    # Big faded number/character feel — use abstract block
    add_rect(s, Inches(10.0), Inches(0.6), Inches(2.8), Inches(2.8),
             fill=NAVY_SOFT)
    add_text(s, Inches(10.0), Inches(0.9), Inches(2.8), Inches(2.4),
             "十五五", size=72, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(10.0), Inches(2.6), Inches(2.8), Inches(0.6),
             "2026 — 2030", size=14, color=GOLD_LIGHT,
             align=PP_ALIGN.CENTER, font=EN_FONT)

    # Subtle 印章-style red square
    add_rect(s, Inches(10.5), Inches(5.8), Inches(0.9), Inches(0.9),
             fill=CRIMSON)
    add_text(s, Inches(10.5), Inches(5.85), Inches(0.9), Inches(0.8),
             "高\n质\n量", size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.0)

    # Top gold bar
    add_rect(s, 0, 0, Inches(9.3), Inches(0.18), fill=GOLD)

    # Kicker
    add_text(s, Inches(0.8), Inches(0.6), Inches(8), Inches(0.4),
             "上海市房产经济学会  ·  中青年学术研讨会",
             size=13, color=GOLD_LIGHT, bold=True)
    add_line(s, Inches(0.8), Inches(1.05), Inches(3.6), Inches(1.05),
             color=GOLD, weight=2)

    # Main title (Chinese)
    add_text(s, Inches(0.8), Inches(1.4), Inches(8.4), Inches(1.4),
             "“十五五”时期", size=36, bold=True, color=WHITE)
    add_text(s, Inches(0.8), Inches(2.1), Inches(8.4), Inches(1.4),
             "房地产高质量发展的新模式", size=40, bold=True, color=WHITE)

    # Subtitle
    add_text(s, Inches(0.8), Inches(3.3), Inches(8.4), Inches(0.6),
             "——以科技创新与产城融合为驱动",
             size=20, color=GOLD_LIGHT, bold=True)

    # English line
    add_text(s, Inches(0.8), Inches(4.1), Inches(8.4), Inches(0.4),
             "A New Model of High-Quality Real-Estate Development for the 15th Five-Year Plan",
             size=12, color=RGBColor(0xB8, 0xC4, 0xD6), font=EN_FONT)

    # Divider
    add_line(s, Inches(0.8), Inches(4.7), Inches(5.5), Inches(4.7),
             color=GOLD, weight=1)

    # Footer block — speaker / venue
    add_text(s, Inches(0.8), Inches(5.0), Inches(8), Inches(0.4),
             "汇报场合：中青年学术研讨会  ·  规模 50–60 人",
             size=13, color=WHITE)
    add_text(s, Inches(0.8), Inches(5.45), Inches(8), Inches(0.4),
             "汇报时间：2026 年 5 月 27 日  下午  ·  时长 30 分钟",
             size=13, color=WHITE)
    add_text(s, Inches(0.8), Inches(5.9), Inches(8), Inches(0.4),
             "议题方向：十五五规划 · 房地产新模式 · 高质量发展",
             size=13, color=GOLD_LIGHT, bold=True)

    add_text(s, Inches(0.8), Inches(6.8), Inches(8), Inches(0.4),
             "汇报人：（拟定稿  ·  供老同志审阅）",
             size=11, color=RGBColor(0x9A, 0xA9, 0xC2))


# ====================================================================
# Slide 2 — Abstract (摘要)
# ====================================================================
def slide_abstract():
    s = prs.slides.add_slide(BLANK)
    page_background(s)
    header_bar(s, "ABSTRACT  ·  内容摘要", "150 字摘要  ·  供主办方先行审定", 2)

    # Left vertical accent
    add_rect(s, Inches(0.6), Inches(1.5), Inches(0.12), Inches(5.0),
             fill=GOLD)

    # Quote-style abstract
    abstract = (
        "立足“十五五”规划开局，房地产行业正从“高杠杆、高周转”的旧模式，"
        "全面转向“稳健经营、提质增效”的新阶段。本汇报围绕“房地产新模式”"
        "与“高质量发展”两大主题，提出：未来五年，行业竞争的核心已由土储与"
        "去化，转向存量资产精细化运营与科技产业化融合。报告将聚焦人工智能"
        "与大语言模型在不动产管理中的落地、科创生态对空间价值的二次赋能，"
        "以及“学术—产业—投资”跨界协同的实践路径，探讨服务实体经济、孕育"
        "新质生产力的不动产新范式。"
    )

    add_text(s, Inches(1.0), Inches(1.55), Inches(7.6), Inches(0.5),
             "内容摘要  ·  Executive Abstract",
             size=20, bold=True, color=NAVY)
    add_text(s, Inches(1.0), Inches(2.05), Inches(7.6), Inches(0.35),
             "Approx. 150 字，紧扣“十五五 / 高质量 / 新模式 / 转型 / 实体经济”",
             size=11, color=SUBTEXT)

    add_text(s, Inches(1.0), Inches(2.6), Inches(7.6), Inches(3.8),
             abstract, size=15, color=TEXT, line_spacing=1.8)

    # Keyword cloud (right side)
    kx, ky = Inches(9.4), Inches(1.55)
    add_rect(s, kx, ky, Inches(3.4), Inches(5.0),
             fill=PAPER_DEEP)
    add_rect(s, kx, ky, Inches(3.4), Inches(0.5), fill=NAVY)
    add_text(s, kx, ky, Inches(3.4), Inches(0.5),
             "关 键 词 / KEYWORDS", size=13, bold=True, color=GOLD,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    keywords = [
        ("十五五规划",  20, NAVY,    True),
        ("高质量发展",  22, CRIMSON, True),
        ("房地产新模式", 18, NAVY,   True),
        ("产城融合",    16, TEAL,    False),
        ("科技赋能",    16, NAVY_SOFT, False),
        ("存量运营",    14, SUBTEXT, False),
        ("新质生产力",  18, GOLD,    True),
        ("实体经济",    15, NAVY,    False),
        ("AI · 大模型", 14, TEAL,    False),
        ("产业转型",    14, SUBTEXT, False),
    ]
    cy = ky + Inches(0.7)
    for kw, sz, col, bd in keywords:
        add_text(s, kx + Inches(0.2), cy, Inches(3.0), Inches(0.45),
                 kw, size=sz, bold=bd, color=col, align=PP_ALIGN.CENTER)
        cy += Inches(0.42)

    footer(s)


# ====================================================================
# Slide 3 — Agenda (汇报大纲)
# ====================================================================
def slide_agenda():
    s = prs.slides.add_slide(BLANK)
    page_background(s)
    header_bar(s, "AGENDA  ·  汇报大纲", "四个板块  ·  共 30 分钟", 3)

    add_text(s, Inches(0.8), Inches(1.4), Inches(11), Inches(0.5),
             "汇报大纲   Outline",
             size=24, bold=True, color=NAVY)
    add_text(s, Inches(0.8), Inches(1.9), Inches(11), Inches(0.4),
             "以宏观政策切入，落地到实战领域，兼顾学术严谨与产业落地",
             size=12, color=SUBTEXT)

    sections = [
        ("01", "时代定调",   "“十五五”开局与房地产周期的新旧转换",   "约 5 分钟",  NAVY),
        ("02", "破局寻路",   "房地产“新模式”的核心内涵与高质量标准", "约 8 分钟",  TEAL),
        ("03", "核心动能",   "科技创新如何驱动不动产高质量发展",     "约 12 分钟", CRIMSON),
        ("04", "总结展望",   "“十五五”时期从业者的知与行",             "约 5 分钟",  GOLD),
    ]

    top = Inches(2.65)
    card_w = Inches(2.9)
    card_h = Inches(4.0)
    gap = Inches(0.18)
    start_x = Inches(0.8)
    for i, (num, title, sub, dur, col) in enumerate(sections):
        x = start_x + (card_w + gap) * i
        # Card
        add_rect(s, x, top, card_w, card_h, fill=WHITE)
        # Color bar on top
        add_rect(s, x, top, card_w, Inches(0.18), fill=col)
        # Big number
        add_text(s, x, top + Inches(0.4), card_w, Inches(1.0),
                 num, size=54, bold=True, color=col,
                 align=PP_ALIGN.CENTER, font=EN_FONT)
        # underline
        add_line(s, x + Inches(0.8), top + Inches(1.55),
                 x + card_w - Inches(0.8), top + Inches(1.55),
                 color=GOLD, weight=1.2)
        # Title
        add_text(s, x + Inches(0.2), top + Inches(1.7), card_w - Inches(0.4),
                 Inches(0.6), title, size=22, bold=True, color=NAVY,
                 align=PP_ALIGN.CENTER)
        # Subtitle
        add_text(s, x + Inches(0.25), top + Inches(2.4),
                 card_w - Inches(0.5), Inches(1.0),
                 sub, size=12, color=TEXT,
                 align=PP_ALIGN.CENTER, line_spacing=1.5)
        # Duration pill
        add_rect(s, x + Inches(0.7), top + Inches(3.5),
                 card_w - Inches(1.4), Inches(0.35),
                 fill=PAPER_DEEP, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(s, x + Inches(0.7), top + Inches(3.5),
                 card_w - Inches(1.4), Inches(0.35),
                 dur, size=11, bold=True, color=NAVY,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    footer(s)


# ====================================================================
# Section divider helper
# ====================================================================
def section_divider(num, kicker, title, subtitle, page_num):
    s = prs.slides.add_slide(BLANK)
    add_rect(s, 0, 0, SW, SH, fill=NAVY)
    # Right diagonal block
    add_rect(s, Inches(8.5), 0, Inches(4.83), SH, fill=NAVY_DARK)
    # gold accent line
    add_rect(s, Inches(8.42), 0, Inches(0.08), SH, fill=GOLD)
    # Top gold bar
    add_rect(s, 0, 0, SW, Inches(0.15), fill=GOLD)

    # Section number, large
    add_text(s, Inches(0.8), Inches(1.4), Inches(4), Inches(2.5),
             num, size=180, bold=True, color=NAVY_SOFT,
             font=EN_FONT, anchor=MSO_ANCHOR.TOP)

    # PART label
    add_text(s, Inches(0.85), Inches(3.6), Inches(6), Inches(0.4),
             f"PART {num}  ·  {kicker}",
             size=14, color=GOLD, bold=True, font=EN_FONT)
    add_line(s, Inches(0.85), Inches(4.05), Inches(2.5), Inches(4.05),
             color=GOLD, weight=2)

    # Title
    add_text(s, Inches(0.85), Inches(4.25), Inches(11), Inches(1.0),
             title, size=40, bold=True, color=WHITE)

    # Subtitle
    add_text(s, Inches(0.85), Inches(5.2), Inches(11), Inches(0.6),
             subtitle, size=16, color=RGBColor(0xC8, 0xD3, 0xE4))

    # Page number
    add_text(s, Inches(11.6), Inches(0.4), Inches(1.2), Inches(0.5),
             f"{page_num:02d} / 18", size=11, color=GOLD_LIGHT,
             bold=True, align=PP_ALIGN.RIGHT)

    # Bottom thin line
    add_line(s, Inches(0.8), Inches(6.8), Inches(12.5), Inches(6.8),
             color=NAVY_SOFT, weight=0.75)
    add_text(s, Inches(0.8), Inches(6.85), Inches(12), Inches(0.4),
             "上海市房产经济学会  ·  中青年学术研讨会  ·  2026.05.27",
             size=10, color=RGBColor(0x9A, 0xA9, 0xC2))


# ====================================================================
# Section 1 — 时代定调
# ====================================================================
def slide_s1_macro():
    s = prs.slides.add_slide(BLANK)
    page_background(s)
    header_bar(s, "PART 01  ·  时代定调", "宏观背景解读：“十五五”的战略地位", 5)

    add_text(s, Inches(0.8), Inches(1.4), Inches(11), Inches(0.5),
             "“十五五”：国家经济结构转型的关键五年",
             size=24, bold=True, color=NAVY)
    add_text(s, Inches(0.8), Inches(1.95), Inches(11), Inches(0.4),
             "从“数量扩张” 走向 “质量跃迁”，从“要素驱动” 走向 “创新驱动”",
             size=13, color=SUBTEXT)

    # Three pillar cards
    pillars = [
        ("战略定位", "国家由 中高速增长 转向 高质量发展 的关键期；\n"
                     "也是构建 新发展格局、培育 新质生产力 的攻坚期。"),
        ("政策主线", "扩大内需 + 科技自立自强 + 现代化产业体系；\n"
                     "房地产作为支柱产业，承担稳投资、稳预期的功能。"),
        ("行业坐标", "“房住不炒”长期不变，行业重心从 增量开发 \n"
                     "转向 存量经营 与 服务实体经济。"),
    ]
    top = Inches(2.7)
    cw = Inches(3.95); ch = Inches(3.5)
    for i, (h, b) in enumerate(pillars):
        x = Inches(0.8) + (cw + Inches(0.2)) * i
        add_rect(s, x, top, cw, ch, fill=WHITE)
        add_rect(s, x, top, Inches(0.12), ch, fill=GOLD)
        add_text(s, x + Inches(0.4), top + Inches(0.35), cw - Inches(0.6),
                 Inches(0.5), h, size=18, bold=True, color=NAVY)
        add_line(s, x + Inches(0.4), top + Inches(0.95),
                 x + Inches(1.4), top + Inches(0.95), color=GOLD, weight=1.5)
        add_text(s, x + Inches(0.4), top + Inches(1.15), cw - Inches(0.6),
                 ch - Inches(1.3), b, size=13, color=TEXT, line_spacing=1.7)
    # Footnote
    add_text(s, Inches(0.8), Inches(6.35), Inches(12), Inches(0.4),
             "▌ 论点：宏观叙事下，房地产的角色从“增长引擎”转向“质量底盘”。",
             size=12, bold=True, color=CRIMSON)

    footer(s)


def slide_s1_farewell():
    s = prs.slides.add_slide(BLANK)
    page_background(s)
    header_bar(s, "PART 01  ·  时代定调", "告别旧模式：从“三高”到“稳健经营、提质增效”", 6)

    add_text(s, Inches(0.8), Inches(1.4), Inches(11.5), Inches(0.5),
             "旧模式 vs 新模式 · 对照表",
             size=24, bold=True, color=NAVY)

    # Two-column compare table
    headers = ["维度", "旧模式（增量时代）", "新模式（“十五五”）"]
    rows = [
        ("发展逻辑", "高杠杆 · 高周转 · 高增长",       "稳健经营 · 提质增效 · 长周期"),
        ("竞争核心", "土地储备 · 销售去化",            "存量运营 · 服务能力"),
        ("收益来源", "开发利润 · 资产升值",            "运营收益 · 资管费 · 增值服务"),
        ("产品形态", "标准化住宅 · 商办堆量",          "产品+服务+生态 · 场景化复合空间"),
        ("评价指标", "规模 · 排名 · 市占率",           "ROE · ROA · NOI · 客户复购"),
        ("行业关系", "土地财政 · 金融加杠杆",          "服务实体 · 赋能产业 · 协同创新"),
    ]
    top = Inches(2.05)
    col_x = [Inches(0.8), Inches(2.8), Inches(7.7)]
    col_w = [Inches(2.0), Inches(4.9), Inches(4.85)]
    rh   = Inches(0.62)

    # Header row
    add_rect(s, col_x[0], top, col_w[0]+col_w[1]+col_w[2], rh, fill=NAVY)
    for i, h in enumerate(headers):
        add_text(s, col_x[i] + Inches(0.2), top, col_w[i] - Inches(0.2), rh,
                 h, size=14, bold=True, color=WHITE,
                 anchor=MSO_ANCHOR.MIDDLE)

    for ri, row in enumerate(rows):
        y = top + rh + rh * ri
        bg = WHITE if ri % 2 == 0 else PAPER_DEEP
        add_rect(s, col_x[0], y, col_w[0]+col_w[1]+col_w[2], rh, fill=bg)
        # cell texts
        add_text(s, col_x[0] + Inches(0.2), y, col_w[0]-Inches(0.2), rh,
                 row[0], size=13, bold=True, color=NAVY,
                 anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, col_x[1] + Inches(0.2), y, col_w[1]-Inches(0.2), rh,
                 row[1], size=13, color=SUBTEXT,
                 anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, col_x[2] + Inches(0.2), y, col_w[2]-Inches(0.2), rh,
                 row[2], size=13, bold=True, color=CRIMSON,
                 anchor=MSO_ANCHOR.MIDDLE)

    # Bottom line
    add_text(s, Inches(0.8), Inches(6.7), Inches(12), Inches(0.4),
             "▌ 转变不是周期性的回调，而是结构性、范式性的“换轨”。",
             size=12, bold=True, color=NAVY)
    footer(s)


def slide_s1_thesis():
    s = prs.slides.add_slide(BLANK)
    page_background(s)
    header_bar(s, "PART 01  ·  时代定调", "核心论点：竞争维度的根本转移", 7)

    # Big quote block left
    add_rect(s, Inches(0.8), Inches(1.6), Inches(7.2), Inches(5.0),
             fill=NAVY)
    add_rect(s, Inches(0.8), Inches(1.6), Inches(0.18), Inches(5.0),
             fill=GOLD)
    add_text(s, Inches(1.1), Inches(1.85), Inches(0.8), Inches(0.6),
             "“", size=72, bold=True, color=GOLD, font=EN_FONT)
    add_text(s, Inches(1.1), Inches(2.55), Inches(6.8), Inches(3.2),
             "“十五五”期间的房地产高质量发展，\n"
             "不再是单纯的去化和土储竞争，\n"
             "而是 存量资产精细化运营 \n"
             "与 科技产业化融合 的竞争。",
             size=22, bold=True, color=WHITE, line_spacing=1.6)
    add_line(s, Inches(1.1), Inches(5.75), Inches(3.0), Inches(5.75),
             color=GOLD, weight=2)
    add_text(s, Inches(1.1), Inches(5.85), Inches(6.5), Inches(0.4),
             "— 本汇报的核心论点",
             size=13, color=GOLD_LIGHT, bold=True)

    # Right hand: 3 key shifts
    add_text(s, Inches(8.4), Inches(1.6), Inches(4.5), Inches(0.5),
             "三个根本转移", size=20, bold=True, color=NAVY)
    add_line(s, Inches(8.4), Inches(2.1), Inches(9.4), Inches(2.1),
             color=GOLD, weight=2)
    shifts = [
        ("从 开发 → 运营", "拼建造、拼速度，转为拼服务、拼体验、拼长期收益。"),
        ("从 资产 → 平台", "重资产持有，转为轻资产管理输出与品牌赋能。"),
        ("从 空间 → 生态", "提供物理空间，转为构建产业 + 科创 + 资本生态。"),
    ]
    bullet_block(s, Inches(8.4), Inches(2.3), Inches(4.5), shifts,
                 size=14, gap=Inches(0.15))
    footer(s)


# ====================================================================
# Section 2 — 破局寻路
# ====================================================================
def slide_s2_space():
    s = prs.slides.add_slide(BLANK)
    page_background(s)
    header_bar(s, "PART 02  ·  破局寻路",
               "高质量发展的空间载体：从“居住与办公”到“科创与产业”", 9)

    add_text(s, Inches(0.8), Inches(1.4), Inches(12), Inches(0.5),
             "空间载体的代际跃迁",
             size=24, bold=True, color=NAVY)
    add_text(s, Inches(0.8), Inches(1.95), Inches(12), Inches(0.4),
             "不动产不再只是“盛放生活与办公的容器”，而是“激发创新与产业的舞台”。",
             size=13, color=SUBTEXT)

    # Evolution arrow
    stages = [
        ("1.0", "居住空间",  "解决“有没有”\n刚需 · 改善",            NAVY_SOFT),
        ("2.0", "办公空间",  "解决“好不好”\n商办 · 商业 · 物流",     TEAL),
        ("3.0", "产城融合",  "解决“活不活”\n科创园 · 总部基地",     CRIMSON),
        ("4.0", "科创生态",  "解决“能不能孕育新质生产力”\n超级会客厅", GOLD),
    ]
    top = Inches(2.7)
    cw = Inches(2.95); ch = Inches(3.3)
    for i, (n, h, b, col) in enumerate(stages):
        x = Inches(0.8) + (cw + Inches(0.15)) * i
        add_rect(s, x, top, cw, ch, fill=WHITE)
        add_rect(s, x, top, cw, Inches(0.6), fill=col)
        add_text(s, x, top, cw, Inches(0.6),
                 f"阶段 {n}", size=14, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font=EN_FONT)
        add_text(s, x + Inches(0.2), top + Inches(0.85),
                 cw - Inches(0.4), Inches(0.6),
                 h, size=20, bold=True, color=NAVY,
                 align=PP_ALIGN.CENTER)
        add_line(s, x + Inches(0.8), top + Inches(1.55),
                 x + cw - Inches(0.8), top + Inches(1.55),
                 color=GOLD, weight=1.2)
        add_text(s, x + Inches(0.25), top + Inches(1.75),
                 cw - Inches(0.5), ch - Inches(1.9),
                 b, size=12, color=TEXT,
                 align=PP_ALIGN.CENTER, line_spacing=1.6)

    # Arrow callout
    add_text(s, Inches(0.8), Inches(6.3), Inches(12), Inches(0.4),
             "▌ 高质量发展的空间载体，是 “能让科技 · 资本 · 人才共振” 的产业舞台。",
             size=13, bold=True, color=NAVY)
    footer(s)


def slide_s2_two_features():
    s = prs.slides.add_slide(BLANK)
    page_background(s)
    header_bar(s, "PART 02  ·  破局寻路",
               "新模式的两大特征：轻重分离 + 产城融合", 10)

    add_text(s, Inches(0.8), Inches(1.4), Inches(12), Inches(0.5),
             "“新模式”的两大底层特征",
             size=24, bold=True, color=NAVY)

    # Two halves
    left_x  = Inches(0.8);  right_x = Inches(7.0)
    bw      = Inches(5.95); bh      = Inches(4.6)
    top     = Inches(2.05)

    # Left card — 轻重分离
    add_rect(s, left_x, top, bw, bh, fill=WHITE)
    add_rect(s, left_x, top, bw, Inches(0.6), fill=NAVY)
    add_text(s, left_x, top, bw, Inches(0.6),
             "特征一  ·  轻重分离",
             size=18, bold=True, color=GOLD,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, left_x + Inches(0.5), top + Inches(0.9),
             bw - Inches(1), Inches(0.5),
             "重资产持有 → 轻资产管理输出 · 品牌输出",
             size=14, bold=True, color=NAVY)
    add_line(s, left_x + Inches(0.5), top + Inches(1.45),
             left_x + Inches(2), top + Inches(1.45),
             color=GOLD, weight=1.5)
    items_l = [
        ("资本结构",  "降低对单一开发利润和高杠杆融资的依赖。"),
        ("能力输出",  "把管理、运营、招商、科技四大能力对外赋能。"),
        ("收益模式",  "管理费 + 业绩分成 + 长期租金 + 资管增值。"),
        ("退出路径",  "对接 公募 REITs · 类 REITs · 不动产私募基金。"),
    ]
    bullet_block(s, left_x + Inches(0.4), top + Inches(1.6),
                 bw - Inches(0.8), items_l, size=12, gap=Inches(0.04),
                 heading_h=Inches(0.34), body_h=Inches(0.36),
                 line_spacing=1.3)

    # Right card — 产城融合
    add_rect(s, right_x, top, bw, bh, fill=WHITE)
    add_rect(s, right_x, top, bw, Inches(0.6), fill=CRIMSON)
    add_text(s, right_x, top, bw, Inches(0.6),
             "特征二  ·  产城融合",
             size=18, bold=True, color=GOLD,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, right_x + Inches(0.5), top + Inches(0.9),
             bw - Inches(1), Inches(0.5),
             "聚焦实体经济  ·  打造科创“超级会客厅”",
             size=14, bold=True, color=NAVY)
    add_line(s, right_x + Inches(0.5), top + Inches(1.45),
             right_x + Inches(2), top + Inches(1.45),
             color=GOLD, weight=1.5)
    items_r = [
        ("产业链接",  "导入高校智库 · 科技企业 · 国际创新资源。"),
        ("场景运营",  "围绕“专精特新”企业、OPC（单一人公司）等新主体打造场景。"),
        ("城市价值",  "让不动产成为产业升级与城市更新的“连接器”。"),
        ("社会贡献",  "服务实体经济，孕育新质生产力的物理底座。"),
    ]
    bullet_block(s, right_x + Inches(0.4), top + Inches(1.6),
                 bw - Inches(0.8), items_r, size=12, gap=Inches(0.04),
                 heading_h=Inches(0.34), body_h=Inches(0.36),
                 line_spacing=1.3)

    add_text(s, Inches(0.8), Inches(6.75), Inches(12), Inches(0.4),
             "▌ 一手做减法（轻），一手做加法（产）——这是新模式的两条腿。",
             size=12, bold=True, color=NAVY)
    footer(s)


# ====================================================================
# Section 3 — 核心动能 (the speaker's main territory)
# ====================================================================
def slide_s3_overview():
    s = prs.slides.add_slide(BLANK)
    page_background(s)
    header_bar(s, "PART 03  ·  核心动能",
               "科技创新驱动不动产高质量发展 —— 三条主线", 12)

    add_text(s, Inches(0.8), Inches(1.4), Inches(12), Inches(0.5),
             "三条主线：管理提效 · 空间增值 · 跨界协同",
             size=24, bold=True, color=NAVY)
    add_text(s, Inches(0.8), Inches(1.95), Inches(12), Inches(0.4),
             "从“砖头”到“算法”，科技把不动产从“资产”升级为“平台”。",
             size=13, color=SUBTEXT)

    lanes = [
        ("管理提效", "AI · 大模型 · 数字化",
         "投资测算 · 资产管理 · 招商引资全链路自动化",  NAVY),
        ("空间增值", "科创生态  ·  场景运营",
         "前沿沙龙 · 国际路演 · 专精特新 · OPC 服务",   CRIMSON),
        ("跨界协同", "学术 · 产业 · 投资",
         "打通三圈壁垒  ·  不动产价值二次跃升",         GOLD),
    ]
    top = Inches(2.75)
    cw = Inches(3.95); ch = Inches(3.6)
    for i, (h, sub, b, col) in enumerate(lanes):
        x = Inches(0.8) + (cw + Inches(0.2)) * i
        add_rect(s, x, top, cw, ch, fill=WHITE)
        # number badge
        add_rect(s, x + Inches(0.3), top + Inches(0.3),
                 Inches(0.7), Inches(0.7), fill=col, shape=MSO_SHAPE.OVAL)
        add_text(s, x + Inches(0.3), top + Inches(0.3),
                 Inches(0.7), Inches(0.7),
                 f"0{i+1}", size=18, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font=EN_FONT)
        add_text(s, x + Inches(1.15), top + Inches(0.35),
                 cw - Inches(1.4), Inches(0.45),
                 h, size=20, bold=True, color=NAVY)
        add_text(s, x + Inches(1.15), top + Inches(0.85),
                 cw - Inches(1.4), Inches(0.45),
                 sub, size=12, color=col, bold=True)
        add_line(s, x + Inches(0.3), top + Inches(1.55),
                 x + cw - Inches(0.3), top + Inches(1.55),
                 color=LINE, weight=1)
        add_text(s, x + Inches(0.3), top + Inches(1.7),
                 cw - Inches(0.6), ch - Inches(1.9),
                 b, size=14, color=TEXT, line_spacing=1.7)

    add_text(s, Inches(0.8), Inches(6.55), Inches(12), Inches(0.4),
             "▌ 接下来逐条展开 —— 这三条主线，决定行业未来五年的“代差”。",
             size=12, bold=True, color=CRIMSON)
    footer(s)


def slide_s3_efficiency():
    s = prs.slides.add_slide(BLANK)
    page_background(s)
    header_bar(s, "PART 03 · 核心动能 ①",
               "管理提效：AI 与数字化重塑不动产经营", 13)

    add_text(s, Inches(0.8), Inches(1.4), Inches(12), Inches(0.5),
             "大语言模型 + AI，让不动产管理“看得清、算得快、跑得稳”",
             size=22, bold=True, color=NAVY)

    # Four pillars (2x2)
    cells = [
        ("投资测算",  "LLM 辅助尽调",
         "项目可研、租赁条款抽取、敏感性分析自动化；\n投资委员会的决策周期由“周”缩到“天”。", NAVY),
        ("资产管理",  "AI 驱动 NOI 优化",
         "租户画像 · 续约预测 · 能耗优化 · 维保排程；\n实现资产 ROE / NOI 的可视化追踪。", TEAL),
        ("招商引资",  "自动化获客与匹配",
         "依托大模型构建“产业图谱 + 招商雷达”，\n精准触达专精特新、总部企业与外资客户。", CRIMSON),
        ("风险合规",  "智能审单与监控",
         "合同审查、合规巡检、舆情监测全流程 AI 化；\n降低人工成本，提高风控韧性。", GOLD),
    ]
    grid_top = Inches(2.05)
    cw, ch = Inches(5.95), Inches(2.15)
    for i, (h, k, b, col) in enumerate(cells):
        r, c = divmod(i, 2)
        x = Inches(0.8) + c * (cw + Inches(0.2))
        y = grid_top + r * (ch + Inches(0.2))
        add_rect(s, x, y, cw, ch, fill=WHITE)
        add_rect(s, x, y, Inches(0.15), ch, fill=col)
        add_text(s, x + Inches(0.4), y + Inches(0.25),
                 cw - Inches(0.6), Inches(0.45),
                 h, size=17, bold=True, color=NAVY)
        add_text(s, x + Inches(0.4), y + Inches(0.7),
                 cw - Inches(0.6), Inches(0.35),
                 k, size=12, color=col, bold=True)
        add_text(s, x + Inches(0.4), y + Inches(1.1),
                 cw - Inches(0.6), ch - Inches(1.2),
                 b, size=13, color=TEXT, line_spacing=1.6)

    add_text(s, Inches(0.8), Inches(6.85), Inches(12), Inches(0.4),
             "▌ 关键词：自动化  ·  数据资产化  ·  人机协同  ·  端到端流程再造。",
             size=12, bold=True, color=NAVY)
    footer(s)


def slide_s3_space_value():
    s = prs.slides.add_slide(BLANK)
    page_background(s)
    header_bar(s, "PART 03 · 核心动能 ②",
               "空间增值：科创生态对物理空间的二次赋能", 14)

    add_text(s, Inches(0.8), Inches(1.4), Inches(12), Inches(0.5),
             "把物业打造成 “产业、人才、资本” 共振的活力场",
             size=22, bold=True, color=NAVY)

    # Left: ecosystem layers
    lx, ly = Inches(0.8), Inches(2.1)
    lw, lh = Inches(6.0), Inches(4.5)
    add_rect(s, lx, ly, lw, lh, fill=WHITE)
    add_rect(s, lx, ly, lw, Inches(0.55), fill=NAVY)
    add_text(s, lx, ly, lw, Inches(0.55),
             "生态四层结构  ·  Ecosystem Stack",
             size=15, bold=True, color=GOLD,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    layers = [
        ("L4 · 资本层", "国资基金 · 产业基金 · 不动产 REITs",     CRIMSON),
        ("L3 · 产业层", "专精特新企业 · 总部基地 · OPC 单一人公司", GOLD),
        ("L2 · 创新层", "前沿科技沙龙 · 国际创新路演 · 高校智库", TEAL),
        ("L1 · 空间层", "高品质物业 · 共享场景 · 数字基础设施",   NAVY_SOFT),
    ]
    layer_top = ly + Inches(0.8)
    layer_h = Inches(0.78)
    for i, (h, b, col) in enumerate(layers):
        y = layer_top + i * (layer_h + Inches(0.08))
        add_rect(s, lx + Inches(0.3), y, lw - Inches(0.6), layer_h,
                 fill=PAPER_DEEP)
        add_rect(s, lx + Inches(0.3), y, Inches(0.18), layer_h, fill=col)
        add_text(s, lx + Inches(0.6), y, Inches(1.7), layer_h,
                 h, size=13, bold=True, color=NAVY,
                 anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, lx + Inches(2.3), y, lw - Inches(2.6), layer_h,
                 b, size=12, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)

    # Right: typical scenarios
    rx, ry = Inches(7.0), Inches(2.1)
    rw, rh = Inches(5.8), Inches(4.5)
    add_rect(s, rx, ry, rw, rh, fill=WHITE)
    add_rect(s, rx, ry, rw, Inches(0.55), fill=CRIMSON)
    add_text(s, rx, ry, rw, Inches(0.55),
             "典型场景  ·  Signature Scenarios",
             size=15, bold=True, color=GOLD,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    scenes = [
        ("前沿科技沙龙",  "AI · 生物医药 · 新能源等议题主题夜，\n面向园区企业与高校研究者。"),
        ("国际创新路演",  "对接海外创投与孵化器，\n承接技术引进与跨境合作。"),
        ("专精特新加速器", "围绕“小巨人”企业提供选址、补贴、融资、招才一站式服务。"),
        ("OPC 共享总部",  "为“单一人公司 OPC”定制\n弹性办公 · 行政托管 · 合规服务。"),
    ]
    bullet_block(s, rx + Inches(0.4), ry + Inches(0.75),
                 rw - Inches(0.8), scenes, size=12, gap=Inches(0.04),
                 heading_h=Inches(0.34), body_h=Inches(0.5),
                 line_spacing=1.3)

    add_text(s, Inches(0.8), Inches(6.85), Inches(12), Inches(0.4),
             "▌ 空间增值的本质：从 “出租平方米” 升级为 “出租生态位”。",
             size=12, bold=True, color=CRIMSON)
    footer(s)


def slide_s3_collab():
    s = prs.slides.add_slide(BLANK)
    page_background(s)
    header_bar(s, "PART 03 · 核心动能 ③",
               "跨界协同：打通“学术—产业—投资”三圈壁垒", 15)

    add_text(s, Inches(0.8), Inches(1.4), Inches(12), Inches(0.5),
             "三圈协同：让不动产成为创新要素的“路由器”",
             size=22, bold=True, color=NAVY)

    # Three big circles (conceptual)
    cy = Inches(4.4)
    rds = Inches(2.1)
    centers = [
        (Inches(3.5),  cy, "学术圈", "高校 · 智库 · 研究院",   NAVY),
        (Inches(6.7),  cy, "产业圈", "实体企业 · 园区 · 协会", CRIMSON),
        (Inches(9.9),  cy, "投资圈", "基金 · 银行 · REITs",    GOLD),
    ]
    for cx, c_y, h, sub, col in centers:
        add_rect(s, cx - rds/2, c_y - rds/2, rds, rds,
                 fill=col, shape=MSO_SHAPE.OVAL)
        add_text(s, cx - rds/2, c_y - rds/2 + Inches(0.45),
                 rds, Inches(0.5), h, size=20, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER)
        add_text(s, cx - rds/2, c_y - rds/2 + Inches(1.05),
                 rds, Inches(0.4), sub, size=12, color=WHITE,
                 align=PP_ALIGN.CENTER)

    # Center label "不动产平台"
    add_rect(s, Inches(5.95), Inches(5.5), Inches(1.45), Inches(0.55),
             fill=WHITE, line=GOLD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, Inches(5.95), Inches(5.5), Inches(1.45), Inches(0.55),
             "不动产平台",
             size=13, bold=True, color=NAVY,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Value statements
    vals = [
        ("理论 → 落地", "高校研究成果在园区中试与产业化"),
        ("产业 → 资本", "项目通过 REITs / 私募实现价值变现"),
        ("资本 → 学术", "投资反哺研究与人才培养，形成闭环"),
    ]
    vy = Inches(2.05)
    for i, (h, b) in enumerate(vals):
        x = Inches(0.8) + i * Inches(4.15)
        add_rect(s, x, vy, Inches(3.9), Inches(0.85), fill=WHITE)
        add_rect(s, x, vy, Inches(0.12), Inches(0.85), fill=GOLD)
        add_text(s, x + Inches(0.3), vy + Inches(0.1),
                 Inches(3.5), Inches(0.4),
                 h, size=14, bold=True, color=NAVY)
        add_text(s, x + Inches(0.3), vy + Inches(0.45),
                 Inches(3.5), Inches(0.4),
                 b, size=12, color=TEXT)

    add_text(s, Inches(0.8), Inches(6.85), Inches(12), Inches(0.4),
             "▌ 协同的红利：不动产价值的“二次跃升”——从空间溢价到生态溢价。",
             size=12, bold=True, color=CRIMSON)
    footer(s)


# ====================================================================
# Section 4 — 总结展望
# ====================================================================
def slide_s4_consensus():
    s = prs.slides.add_slide(BLANK)
    page_background(s)
    header_bar(s, "PART 04  ·  总结展望",
               "发展共识：学术与实业的“双轮驱动”", 17)

    add_text(s, Inches(0.8), Inches(1.4), Inches(12), Inches(0.5),
             "高质量发展，是“理论 × 实战”的合奏",
             size=24, bold=True, color=NAVY)

    # Two columns: 学术 / 实业
    cols = [
        ("学术界  ·  提供方向",
         CRIMSON,
         [("理论指引",  "为新模式提供分析框架、政策解读与价值判断。"),
          ("方法供给",  "把产业经济、城市经济与金融工程的方法落到不动产。"),
          ("人才培育",  "为行业培养兼具学术素养与产业视角的中青年骨干。")]),
        ("实业界  ·  蹚出新路",
         NAVY,
         [("工具创新", "用 AI、数字化等科技工具，重塑投资 · 运营 · 招商流程。"),
          ("场景试验", "把科创沙龙、国际路演、OPC 服务在园区中规模化。"),
          ("生态构建", "把单体物业升级为连接学术、产业、资本的综合平台。")]),
    ]
    top = Inches(2.05)
    cw, ch = Inches(5.95), Inches(4.6)
    for i, (title, col, items) in enumerate(cols):
        x = Inches(0.8) + i * (cw + Inches(0.2))
        add_rect(s, x, top, cw, ch, fill=WHITE)
        add_rect(s, x, top, cw, Inches(0.65), fill=col)
        add_text(s, x, top, cw, Inches(0.65),
                 title, size=18, bold=True, color=GOLD,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        bullet_block(s, x + Inches(0.4), top + Inches(0.95),
                     cw - Inches(0.8), items, size=14, gap=Inches(0.1))

    add_text(s, Inches(0.8), Inches(6.85), Inches(12), Inches(0.4),
             "▌ 共识：没有学术的“望远镜”和实业的“显微镜”，就没有真正的高质量发展。",
             size=12, bold=True, color=NAVY)
    footer(s)


def slide_s4_closing():
    s = prs.slides.add_slide(BLANK)
    add_rect(s, 0, 0, SW, SH, fill=NAVY)
    add_rect(s, 0, 0, SW, Inches(0.15), fill=GOLD)

    # Left big quote
    add_text(s, Inches(0.8), Inches(1.0), Inches(12), Inches(0.5),
             "结语  ·  Closing",
             size=14, color=GOLD, bold=True, font=EN_FONT)
    add_line(s, Inches(0.8), Inches(1.5), Inches(2.2), Inches(1.5),
             color=GOLD, weight=2)

    add_text(s, Inches(0.8), Inches(1.8), Inches(12), Inches(1.2),
             "行业虽在阵痛期，",
             size=36, bold=True, color=WHITE)
    add_text(s, Inches(0.8), Inches(2.7), Inches(12), Inches(1.2),
             "但 “和光同尘” —— 顺应科技大势与国家宏观规划。",
             size=28, bold=True, color=GOLD_LIGHT)

    add_text(s, Inches(0.8), Inches(4.0), Inches(12), Inches(1.2),
             "不动产，依然是孕育 ",
             size=26, color=WHITE)
    add_text(s, Inches(4.85), Inches(4.0), Inches(8), Inches(1.2),
             "新质生产力",
             size=32, bold=True, color=GOLD)
    add_text(s, Inches(7.5), Inches(4.0), Inches(8), Inches(1.2),
             " 的最佳土壤。",
             size=26, color=WHITE)

    # Big seal
    add_rect(s, Inches(11.2), Inches(5.6), Inches(1.4), Inches(1.4),
             fill=CRIMSON)
    add_text(s, Inches(11.2), Inches(5.6), Inches(1.4), Inches(1.4),
             "高质量\n发展",
             size=18, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.2)

    # Thank you
    add_text(s, Inches(0.8), Inches(5.8), Inches(8), Inches(0.6),
             "敬请批评指正  ·  THANK YOU",
             size=22, bold=True, color=WHITE)
    add_line(s, Inches(0.8), Inches(6.35), Inches(4.5), Inches(6.35),
             color=GOLD, weight=1.5)
    add_text(s, Inches(0.8), Inches(6.5), Inches(10), Inches(0.4),
             "上海市房产经济学会  ·  中青年学术研讨会  ·  2026.05.27",
             size=12, color=GOLD_LIGHT)


# ====================================================================
# Build deck
# ====================================================================
def build():
    slide_cover()             # 1
    slide_abstract()          # 2
    slide_agenda()            # 3
    section_divider("01", "时代定调", "“十五五”开局与房地产周期的新旧转换",
                    "宏观背景 · 旧模式告别 · 核心论点", 4)  # 4
    slide_s1_macro()          # 5
    slide_s1_farewell()       # 6
    slide_s1_thesis()         # 7
    section_divider("02", "破局寻路", "房地产“新模式”的核心内涵",
                    "空间载体跃迁 · 轻重分离 · 产城融合", 8)  # 8
    slide_s2_space()          # 9
    slide_s2_two_features()   # 10
    section_divider("03", "核心动能", "科技创新如何驱动不动产高质量发展",
                    "管理提效 · 空间增值 · 跨界协同", 11)  # 11
    slide_s3_overview()       # 12
    slide_s3_efficiency()     # 13
    slide_s3_space_value()    # 14
    slide_s3_collab()          # 15
    section_divider("04", "总结展望", "“十五五”时期从业者的知与行",
                    "学术与实业的双轮驱动 · 结语", 16)  # 16
    slide_s4_consensus()      # 17
    slide_s4_closing()        # 18

    out = "/workspace/十五五-房地产高质量发展-汇报大纲.pptx"
    prs.save(out)
    print(f"Saved: {out}")
    return out


if __name__ == "__main__":
    build()
