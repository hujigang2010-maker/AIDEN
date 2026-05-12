"""
森马（上海）国际运营中心 · 创意展示 PPT 生成器
Semir Global Headquarter — Creative Presentation Builder

Design system:
- Dark "潮玩元宇宙" theme (deep navy + magenta / violet / cyan accents)
- Widescreen 16:9 (13.33in × 7.5in)
- Cover · chapter dividers · content slides · finale
"""

from __future__ import annotations
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ---------- Theme ----------
BG       = RGBColor(0x07, 0x07, 0x10)   # deep ink
BG2      = RGBColor(0x10, 0x10, 0x1E)
SURFACE  = RGBColor(0x18, 0x18, 0x2A)
SURFACE2 = RGBColor(0x22, 0x22, 0x38)
TEXT     = RGBColor(0xEC, 0xEC, 0xF4)
DIM      = RGBColor(0x9A, 0xA0, 0xB8)
MUTE     = RGBColor(0x6B, 0x70, 0x90)
ACCENT   = RGBColor(0xFF, 0x4F, 0x8B)   # magenta
ACCENT2  = RGBColor(0x8B, 0x5C, 0xF6)   # violet
ACCENT3  = RGBColor(0x06, 0xB6, 0xD4)   # cyan
ACCENT4  = RGBColor(0xF5, 0x9E, 0x0B)   # amber
ACCENT5  = RGBColor(0x10, 0xB9, 0x81)   # green

CN_FONT  = '思源黑体 CN'  # falls back gracefully
EN_FONT  = 'Inter'

# ---------- Helpers ----------

def set_slide_bg(slide, color: RGBColor) -> None:
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_rect(slide, x, y, w, h, fill=None, line=None, shadow=False):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.shadow.inherit = False
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(0.75)
    if not shadow:
        # turn off default shadow via XML
        sp = shape.shadow._element
    return shape


def add_round_rect(slide, x, y, w, h, fill=None, line=None, radius=0.08):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shape.shadow.inherit = False
    shape.adjustments[0] = radius
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(0.75)
    return shape


def add_oval(slide, x, y, w, h, fill=None, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)
    shape.shadow.inherit = False
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(0.75)
    return shape


def add_text(
    slide, x, y, w, h, text,
    size=14, color=TEXT, bold=False, italic=False,
    font=CN_FONT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
    line_spacing=1.2, letter_spacing=0,
):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    tf.word_wrap = True
    tf.vertical_anchor = anchor

    if isinstance(text, str):
        lines = text.split('\n')
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = line_spacing
            r = p.add_run()
            r.text = line
            r.font.name = font
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.italic = italic
            r.font.color.rgb = color
            if letter_spacing:
                # spc attribute (hundredths of a point); use 100 for ~1pt
                rPr = r._r.get_or_add_rPr()
                rPr.set('spc', str(letter_spacing))
    else:
        # text is list of (segment, opts)
        for i, item in enumerate(text):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = line_spacing
            segs = item if isinstance(item, list) else [item]
            for seg in segs:
                txt = seg.get('text', '')
                r = p.add_run()
                r.text = txt
                r.font.name = seg.get('font', font)
                r.font.size = Pt(seg.get('size', size))
                r.font.bold = seg.get('bold', bold)
                r.font.italic = seg.get('italic', italic)
                r.font.color.rgb = seg.get('color', color)
                ls = seg.get('letter_spacing', letter_spacing)
                if ls:
                    rPr = r._r.get_or_add_rPr()
                    rPr.set('spc', str(ls))
    return tb


def add_line(slide, x1, y1, x2, y2, color=ACCENT, width=1.0):
    ln = slide.shapes.add_connector(1, x1, y1, x2, y2)
    ln.line.color.rgb = color
    ln.line.width = Pt(width)
    return ln


# Decorative background patterns

def hero_bg(slide):
    """Cover-style background: layered glow circles + corner accents."""
    set_slide_bg(slide, BG)
    # giant soft glow circles
    add_oval(slide, Inches(-3), Inches(-2), Inches(8), Inches(8), fill=ACCENT)
    last = slide.shapes[-1]
    last.fill.transparency = 0  # python-pptx can't easily set transparency; mimic via lighter color
    # repaint to lighter mixed color
    last.fill.fore_color.rgb = RGBColor(0x40, 0x14, 0x28)

    add_oval(slide, Inches(8), Inches(-3), Inches(9), Inches(9), fill=RGBColor(0x2A, 0x16, 0x44))
    add_oval(slide, Inches(5), Inches(4), Inches(7), Inches(7), fill=RGBColor(0x06, 0x2A, 0x3A))

    # subtle grid (diagonal lines)
    for i in range(-4, 18):
        add_line(
            slide,
            Inches(i * 0.9), Inches(0),
            Inches(i * 0.9 + 4), Inches(7.5),
            color=RGBColor(0x18, 0x18, 0x2A), width=0.5,
        )

    # overlay a dark scrim
    add_rect(slide, 0, 0, Inches(13.333), Inches(7.5), fill=BG)
    slide.shapes[-1].fill.fore_color.rgb = RGBColor(0x08, 0x08, 0x14)


def section_bg(slide, tint=ACCENT):
    """Chapter-divider background."""
    set_slide_bg(slide, BG)
    # large tinted glow blocks
    add_oval(slide, Inches(-3), Inches(-3), Inches(9), Inches(9),
             fill=RGBColor(min(tint[0] // 4 + 16, 255), min(tint[1] // 4 + 8, 255), min(tint[2] // 4 + 24, 255)))
    add_oval(slide, Inches(7), Inches(3), Inches(9), Inches(9),
             fill=RGBColor(0x2A, 0x16, 0x44))


def content_bg(slide):
    set_slide_bg(slide, BG)
    # subtle top accent bar
    add_rect(slide, 0, 0, Inches(13.333), Inches(0.06), fill=ACCENT)
    add_rect(slide, Inches(8), 0, Inches(5.333), Inches(0.06), fill=ACCENT2)
    add_rect(slide, Inches(11), 0, Inches(2.333), Inches(0.06), fill=ACCENT3)


def page_footer(slide, page_no: int, total: int, chapter_label: str = ''):
    # left: brand
    add_text(
        slide, Inches(0.5), Inches(7.05), Inches(8), Inches(0.4),
        '森马（上海）国际运营中心 · SEMIR GROUP\'S GLOBAL HEADQUARTER',
        size=8, color=MUTE, letter_spacing=200,
    )
    # right: page no
    add_text(
        slide, Inches(11.5), Inches(7.05), Inches(1.3), Inches(0.4),
        f'{page_no:02d} / {total:02d}',
        size=9, color=DIM, align=PP_ALIGN.RIGHT, font=EN_FONT, letter_spacing=200,
    )
    if chapter_label:
        add_text(
            slide, Inches(0.5), Inches(0.18), Inches(6), Inches(0.3),
            chapter_label,
            size=9, color=ACCENT, letter_spacing=300, font=EN_FONT, bold=True,
        )


def kicker(slide, text, x=Inches(0.7), y=Inches(0.55), color=ACCENT):
    add_text(slide, x, y, Inches(6), Inches(0.4), text,
             size=11, color=color, letter_spacing=300, bold=True)


def title(slide, text, x=Inches(0.7), y=Inches(0.95), w=Inches(12), h=Inches(1.2), size=36, color=TEXT):
    add_text(slide, x, y, w, h, text, size=size, color=color, bold=True, line_spacing=1.1)


def subtitle(slide, text, x=Inches(0.7), y=Inches(1.85), w=Inches(12), h=Inches(0.6), size=13, color=DIM):
    add_text(slide, x, y, w, h, text, size=size, color=color, line_spacing=1.6)


# ============================================================
# Build the deck
# ============================================================
def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    TOTAL = 24  # rough count for footer

    # -------- 1. Cover --------
    s = prs.slides.add_slide(blank)
    hero_bg(s)

    # Eyebrow pill
    pill = add_round_rect(s, Inches(5.0), Inches(1.55), Inches(3.3), Inches(0.36),
                          fill=None, line=RGBColor(0x33, 0x33, 0x4A), radius=0.5)
    add_text(s, Inches(5.0), Inches(1.55), Inches(3.3), Inches(0.36),
             'SEMIR · SHANGHAI · MINHANG · WUJING',
             size=9, color=DIM, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             letter_spacing=400, font=EN_FONT)

    # Title
    add_text(s, Inches(0.5), Inches(2.2), Inches(12.333), Inches(1.4),
             '森马（上海）', size=58, color=TEXT, bold=True, align=PP_ALIGN.CENTER, line_spacing=1.0)
    # Gradient-feel via 3 colored words
    add_text(s, Inches(0.5), Inches(3.3), Inches(12.333), Inches(1.4),
             [[
                 {'text': '国际', 'color': ACCENT, 'size': 72, 'bold': True},
                 {'text': '运营', 'color': ACCENT2, 'size': 72, 'bold': True},
                 {'text': '中心', 'color': ACCENT3, 'size': 72, 'bold': True},
             ]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP, line_spacing=1.0)

    add_text(s, Inches(0.5), Inches(4.7), Inches(12.333), Inches(0.5),
             "SEMIR GROUP'S GLOBAL HEADQUARTER",
             size=14, color=DIM, align=PP_ALIGN.CENTER, letter_spacing=400, font=EN_FONT)

    add_text(s, Inches(0.5), Inches(5.25), Inches(12.333), Inches(0.5),
             '科技潮玩产业策源高地  ·  Z·世代潮玩社交主场',
             size=16, color=TEXT, align=PP_ALIGN.CENTER, letter_spacing=300)

    # Three pillar tags
    tags = [('IMPORTANT CARRIER', '南部科创走廊\n新兴载体', ACCENT),
            ('INNOVATION NODE',  '大零号湾科创区\n示范枢纽', ACCENT2),
            ('SUPER PLATFORM',   '长三角一核三带\n融合标杆', ACCENT3)]
    x0 = 1.2
    for i, (en, cn, c) in enumerate(tags):
        x = Inches(x0 + i * 3.7)
        box = add_round_rect(s, x, Inches(6.1), Inches(3.4), Inches(0.85),
                             fill=None, line=RGBColor(0x33, 0x33, 0x4A), radius=0.2)
        add_text(s, x + Inches(0.15), Inches(6.18), Inches(3.1), Inches(0.28),
                 en, size=8, color=c, letter_spacing=300, font=EN_FONT, bold=True)
        add_text(s, x + Inches(0.15), Inches(6.42), Inches(3.1), Inches(0.5),
                 cn, size=11, color=TEXT, bold=True, line_spacing=1.2)

    # -------- 2. Table of Contents --------
    s = prs.slides.add_slide(blank)
    content_bg(s)
    kicker(s, 'CONTENTS / 目  录')
    title(s, '一座潮玩之城的四章叙事',  size=40)
    subtitle(s, '从战略卡位到产业生态、从商业落位到运营协同。')

    items = [
        ('01', '项目概况篇', 'PROJECT OVERVIEW',
         '战略定位 · 宏观区位 · 交通辐射 · 周边产业 · 六栋楼宇程序', ACCENT),
        ('02', '产业规划篇', 'INDUSTRY PLANNING',
         '产业定位 · 5 层金字塔 · 双牌照 · 6 大配套 · 三轨运营',  ACCENT2),
        ('03', '商业规划篇', 'COMMERCIAL PLANNING',
         'Z·世代主场 · 业态垂直落位 · 5 大主力业态',                ACCENT3),
        ('04', '运营合作篇', 'OPERATIONS & COOPERATION',
         '特色集会 · 产品首发 · 主题大赛 · 拟合作品牌矩阵',            ACCENT4),
    ]
    y = 2.7
    for num, cn, en, desc, c in items:
        add_text(s, Inches(0.7), Inches(y), Inches(1.6), Inches(0.9),
                 num, size=56, color=c, bold=True, font=EN_FONT, line_spacing=1.0)
        add_text(s, Inches(2.3), Inches(y + 0.08), Inches(3.5), Inches(0.5),
                 cn, size=22, color=TEXT, bold=True)
        add_text(s, Inches(2.3), Inches(y + 0.6), Inches(4), Inches(0.3),
                 en, size=9, color=DIM, letter_spacing=400, font=EN_FONT)
        add_text(s, Inches(6.5), Inches(y + 0.25), Inches(6.5), Inches(0.6),
                 desc, size=13, color=DIM)
        add_line(s, Inches(0.7), Inches(y + 1.0), Inches(12.6), Inches(y + 1.0),
                 color=RGBColor(0x22, 0x22, 0x38), width=0.6)
        y += 1.05
    page_footer(s, 2, TOTAL)

    # -------- 3. Chapter 01 divider --------
    s = prs.slides.add_slide(blank)
    section_bg(s, ACCENT)
    add_text(s, Inches(0.5), Inches(0.6), Inches(12), Inches(0.4),
             'CHAPTER 01', size=11, color=ACCENT,
             align=PP_ALIGN.CENTER, letter_spacing=600, bold=True, font=EN_FONT)
    add_text(s, Inches(0.5), Inches(1.3), Inches(12.333), Inches(2.6),
             '01', size=240, color=TEXT, bold=True,
             align=PP_ALIGN.CENTER, font=EN_FONT, line_spacing=0.95)
    add_text(s, Inches(0.5), Inches(4.7), Inches(12.333), Inches(0.8),
             '项目概况篇', size=44, color=TEXT, bold=True,
             align=PP_ALIGN.CENTER, line_spacing=1.1)
    add_text(s, Inches(0.5), Inches(5.6), Inches(12.333), Inches(0.5),
             'PROJECT OVERVIEW', size=13, color=DIM,
             align=PP_ALIGN.CENTER, letter_spacing=600, font=EN_FONT)
    # accent underline
    add_rect(s, Inches(6.16), Inches(6.3), Inches(1), Inches(0.06), fill=ACCENT)

    # -------- 4. 战略定位 --------
    s = prs.slides.add_slide(blank)
    content_bg(s)
    page_footer(s, 4, TOTAL, '01 · 项目概况篇')
    kicker(s, '（一）战略定位 · STRATEGIC POSITIONING')
    title(s, '锚定双重身份  ·  构建示范性融合标杆')
    subtitle(s,
             '锚定「大零号湾科技创新策源功能区文创融合核心区」与「上海市唯一科技时尚特色小镇」双重身份，\n构建时尚研发与文创转化双向赋能的示范性融合标杆。')

    cards = [
        ('01', '南部科创走廊', '的新兴载体', 'IMPORTANT CARRIER', ACCENT),
        ('02', '大零号湾科创区', '示范枢纽',   'INNOVATION NODE',  ACCENT2),
        ('03', '长三角一核三带', '融合标杆',   'SUPER PLATFORM',   ACCENT3),
    ]
    for i, (n, t1, t2, en, c) in enumerate(cards):
        x = Inches(0.7 + i * 4.1)
        card = add_round_rect(s, x, Inches(3.4), Inches(3.9), Inches(2.9),
                              fill=SURFACE, line=RGBColor(0x22, 0x22, 0x38), radius=0.08)
        # top accent bar
        add_rect(s, x, Inches(3.4), Inches(3.9), Inches(0.06), fill=c)
        add_text(s, x + Inches(0.3), Inches(3.65), Inches(2), Inches(0.5),
                 n, size=26, color=c, bold=True, font=EN_FONT)
        add_text(s, x + Inches(0.3), Inches(4.4), Inches(3.5), Inches(0.6),
                 t1, size=22, color=TEXT, bold=True)
        add_text(s, x + Inches(0.3), Inches(4.95), Inches(3.5), Inches(0.6),
                 t2, size=18, color=DIM)
        add_text(s, x + Inches(0.3), Inches(5.85), Inches(3.5), Inches(0.4),
                 en, size=9, color=MUTE, letter_spacing=400, font=EN_FONT, bold=True)

    # -------- 5. 宏观区位 --------
    s = prs.slides.add_slide(blank)
    content_bg(s)
    page_footer(s, 5, TOTAL, '01 · 项目概况篇')
    kicker(s, '（二）宏观区位 · MACRO LOCATION')
    title(s, '闵行千亿发展极  ·  超大产业集群集聚', size=32)
    subtitle(s, '元江路-剑川路地区中心为闵行五大中心之一，能级比肩漕河泾、张江，与梅陇、七宝共筑闵行 TOD 金三角。')

    # Left: five industries
    add_round_rect(s, Inches(0.7), Inches(2.7), Inches(5.6), Inches(3.7),
                   fill=SURFACE, line=RGBColor(0x22, 0x22, 0x38), radius=0.04)
    add_text(s, Inches(1.0), Inches(2.9), Inches(5), Inches(0.4),
             '五大产业鼎立', size=14, color=ACCENT, letter_spacing=300, bold=True)
    five = ['元宇宙', '未来能源', '智慧医疗', '人工智能', '低空经济']
    for i, w in enumerate(five):
        col = i % 3
        row = i // 3
        x = Inches(1.0 + col * 1.65)
        y = Inches(3.55 + row * 0.95)
        add_round_rect(s, x, y, Inches(1.5), Inches(0.75),
                       fill=SURFACE2, line=RGBColor(0x33, 0x33, 0x4A), radius=0.2)
        add_text(s, x, y, Inches(1.5), Inches(0.75), w,
                 size=14, color=TEXT, bold=True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Right: 8 functional zones
    add_round_rect(s, Inches(6.6), Inches(2.7), Inches(6.1), Inches(3.7),
                   fill=SURFACE, line=RGBColor(0x22, 0x22, 0x38), radius=0.04)
    add_text(s, Inches(6.9), Inches(2.9), Inches(5.5), Inches(0.4),
             '上海南部科创中心 · 主要产业功能区',
             size=14, color=ACCENT2, letter_spacing=200, bold=True)
    zones = [
        ('"零号湾"创新创业集聚区', '创新策源区',     ACCENT),
        ('紫竹国家高新技术产业开发区', '高新产业承载区', ACCENT3),
        ('临港浦江国际科技城',     '高新产业承载区', ACCENT3),
        ('向阳工业互联网基地',     '高新产业承载区', ACCENT3),
        ('闵行经济技术开发区',     '先进制造业',  ACCENT4),
        ('莘庄工业区',          '先进制造业',  ACCENT4),
        ('上海航天产业基地',     '战略产业',    ACCENT5),
        ('马桥人工智能创新试验区', '战略产业',    ACCENT5),
    ]
    for i, (name, tag, c) in enumerate(zones):
        col = i % 2
        row = i // 2
        x = Inches(6.85 + col * 2.95)
        y = Inches(3.5 + row * 0.7)
        add_text(s, x, y, Inches(2.85), Inches(0.32),
                 tag, size=8, color=c, letter_spacing=200, bold=True)
        add_text(s, x, y + Inches(0.25), Inches(2.85), Inches(0.4),
                 name, size=10, color=TEXT)

    # -------- 6. 交通辐射 --------
    s = prs.slides.add_slide(blank)
    content_bg(s)
    page_footer(s, 6, TOTAL, '01 · 项目概况篇')
    kicker(s, '（四）交通辐射 · TRAFFIC RADIATION')
    title(s, '科创新城南拓客厅  ·  主城南轴，通达全城', size=30)
    subtitle(s, '元江路定位为闵行南部科创功能轴、东西交通主廊、产城融合纽带。')

    traffic = [
        ('🚄', '紫竹高新区',    '5 km',   '车程 10 min', ACCENT),
        ('🛣️', '申嘉湖高速入口', '4 km',  '车程 15 min', ACCENT2),
        ('✈️', '浦东国际机场',  '36 km',  '车程 1.5 h · 快线 1 h', ACCENT3),
        ('🛫', '虹桥国际机场',  '17 km',  '车程 1 h · 快线 45 min', ACCENT4),
        ('🎓', '大学城',        '4 km',   '车程 10 min · 地铁直达',  ACCENT5),
        ('Ⓜ️', '15 号线元江路 TOD', '日均 5–7 万人次', '主干轨交节点',  ACCENT),
    ]
    for i, (icon, name, dist, time, c) in enumerate(traffic):
        col = i % 3
        row = i // 3
        x = Inches(0.7 + col * 4.1)
        y = Inches(2.85 + row * 1.95)
        add_round_rect(s, x, y, Inches(3.9), Inches(1.7),
                       fill=SURFACE, line=RGBColor(0x22, 0x22, 0x38), radius=0.08)
        # icon block
        add_round_rect(s, x + Inches(0.25), y + Inches(0.35), Inches(1.0), Inches(1.0),
                       fill=SURFACE2, line=None, radius=0.18)
        add_text(s, x + Inches(0.25), y + Inches(0.35), Inches(1.0), Inches(1.0),
                 icon, size=30, color=TEXT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + Inches(1.4), y + Inches(0.3), Inches(2.4), Inches(0.4),
                 name, size=14, color=TEXT, bold=True)
        add_text(s, x + Inches(1.4), y + Inches(0.75), Inches(2.4), Inches(0.4),
                 dist, size=16, color=c, bold=True, font=EN_FONT)
        add_text(s, x + Inches(1.4), y + Inches(1.15), Inches(2.4), Inches(0.4),
                 time, size=10, color=DIM)

    # -------- 7. 立足周边 / 数据 --------
    s = prs.slides.add_slide(blank)
    content_bg(s)
    page_footer(s, 7, TOTAL, '01 · 项目概况篇')
    kicker(s, '（五）立足周边 · SURROUNDINGS')
    title(s, '产业氛围浓厚成熟  ·  人居环境持续优化', size=30)

    stats = [
        ('24', '万', '15 分钟车行覆盖\n居住人口', ACCENT),
        ('12', '万', '15 分钟车行覆盖\n产业办公人口', ACCENT2),
        ('3',  'km', '范围内聚集大量\n产业园 + 企业办公',  ACCENT3),
    ]
    for i, (n, u, lab, c) in enumerate(stats):
        x = Inches(0.7 + i * 4.1)
        add_round_rect(s, x, Inches(2.7), Inches(3.9), Inches(2.2),
                       fill=SURFACE, line=RGBColor(0x22, 0x22, 0x38), radius=0.06)
        add_text(s, x, Inches(2.85), Inches(3.9), Inches(1.2),
                 n, size=72, color=c, bold=True, align=PP_ALIGN.CENTER, font=EN_FONT, line_spacing=1.0)
        add_text(s, x + Inches(2.2), Inches(3.05), Inches(1.5), Inches(0.6),
                 u, size=18, color=TEXT, bold=True)
        add_text(s, x, Inches(4.05), Inches(3.9), Inches(0.8),
                 lab, size=12, color=DIM, align=PP_ALIGN.CENTER, line_spacing=1.4)

    # brand cloud
    add_round_rect(s, Inches(0.7), Inches(5.15), Inches(12), Inches(1.6),
                   fill=SURFACE, line=RGBColor(0x22, 0x22, 0x38), radius=0.04)
    add_text(s, Inches(1.0), Inches(5.25), Inches(11), Inches(0.4),
             '周边产业坐标 · ECOSYSTEM AROUND',
             size=10, color=ACCENT, letter_spacing=300, bold=True)
    brands = ['森马产业', '衣恋', '普洛斯', '雅诗兰黛', '莲谷科技园',
              '保利光合跃城', '颛桥科技绿洲', '拉夏贝尔',
              '舜江集团总部 1 号', '闵行物流园', '中建产研', '首开塘湾基地']
    px, py = 1.0, 5.7
    for i, b in enumerate(brands):
        w = max(0.85, len(b) * 0.18)
        if px + w > 12.6:
            px, py = 1.0, py + 0.45
        chip = add_round_rect(s, Inches(px), Inches(py), Inches(w), Inches(0.35),
                              fill=SURFACE2, line=RGBColor(0x33, 0x33, 0x4A), radius=0.4)
        add_text(s, Inches(px), Inches(py), Inches(w), Inches(0.35),
                 b, size=9, color=TEXT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        px += w + 0.1

    # -------- 8. 项目概况 / 六栋楼 --------
    s = prs.slides.add_slide(blank)
    content_bg(s)
    page_footer(s, 8, TOTAL, '01 · 项目概况篇')
    kicker(s, '（六）项目概况 · SIX BUILDINGS')
    title(s, '六栋楼宇 · 一座潮玩之城', size=30)

    # Totals strip
    add_round_rect(s, Inches(0.7), Inches(2.0), Inches(12), Inches(1.0),
                   fill=SURFACE, line=RGBColor(0x22, 0x22, 0x38), radius=0.1)
    totals = [('22 万㎡', '总体建筑面积', ACCENT),
              ('5.2 万㎡', '商业建筑面积', ACCENT2),
              ('1500+', '商业停车位', ACCENT3)]
    for i, (n, lab, c) in enumerate(totals):
        x = Inches(0.7 + i * 4.0)
        add_text(s, x, Inches(2.1), Inches(4.0), Inches(0.55),
                 n, size=24, color=c, bold=True, align=PP_ALIGN.CENTER, font=EN_FONT)
        add_text(s, x, Inches(2.65), Inches(4.0), Inches(0.35),
                 lab, size=10, color=DIM, align=PP_ALIGN.CENTER, letter_spacing=300)
    add_line(s, Inches(4.7), Inches(2.2), Inches(4.7), Inches(2.85), color=RGBColor(0x33, 0x33, 0x4A))
    add_line(s, Inches(8.7), Inches(2.2), Inches(8.7), Inches(2.85), color=RGBColor(0x33, 0x33, 0x4A))

    buildings = [
        ('1', '零售 · 森马总部',     '1–4F 零售\n5F+ 森马总部办公',           ACCENT),
        ('2', '二次元主题秀场',      '整栋 · Livehouse / 秀场',              ACCENT2),
        ('3', '休闲 · 萌宠 · 酒店',  '1–4F 休闲运动、萌宠\n5F+ 酒店',          ACCENT3),
        ('4', '潮玩艺术 · 直播',     '1–3F 潮玩艺术中心\n4F 直播中心 / 5F+ 集群', ACCENT4),
        ('5', '动漫书店 · 潮玩产业',  '1–4F 动漫书店、娱乐\n5F+ 潮玩产业集群',   ACCENT5),
        ('6', '品质生活 · 商务宴请',  '1–5F 品质生活、餐饮、配套、宴请',         RGBColor(0xF4, 0x3F, 0x5E)),
    ]
    for i, (n, t, prog, c) in enumerate(buildings):
        col = i % 3
        row = i // 3
        x = Inches(0.7 + col * 4.1)
        y = Inches(3.3 + row * 1.85)
        add_round_rect(s, x, y, Inches(3.9), Inches(1.65),
                       fill=SURFACE, line=RGBColor(0x22, 0x22, 0x38), radius=0.08)
        # left vertical accent
        add_rect(s, x, y, Inches(0.08), Inches(1.65), fill=c)
        add_text(s, x + Inches(0.3), y + Inches(0.15), Inches(1.5), Inches(0.5),
                 f'#{n}', size=24, color=c, bold=True, font=EN_FONT)
        add_text(s, x + Inches(1.3), y + Inches(0.22), Inches(2.6), Inches(0.4),
                 t, size=13, color=TEXT, bold=True)
        add_text(s, x + Inches(0.3), y + Inches(0.8), Inches(3.5), Inches(0.8),
                 prog, size=10, color=DIM, line_spacing=1.5)

    # ============================================================
    # CHAPTER 02
    # ============================================================
    s = prs.slides.add_slide(blank)
    section_bg(s, ACCENT2)
    add_text(s, Inches(0.5), Inches(0.6), Inches(12), Inches(0.4),
             'CHAPTER 02', size=11, color=ACCENT2,
             align=PP_ALIGN.CENTER, letter_spacing=600, bold=True, font=EN_FONT)
    add_text(s, Inches(0.5), Inches(1.3), Inches(12.333), Inches(2.6),
             '02', size=240, color=TEXT, bold=True,
             align=PP_ALIGN.CENTER, font=EN_FONT, line_spacing=0.95)
    add_text(s, Inches(0.5), Inches(4.7), Inches(12.333), Inches(0.8),
             '产业规划篇', size=44, color=TEXT, bold=True,
             align=PP_ALIGN.CENTER, line_spacing=1.1)
    add_text(s, Inches(0.5), Inches(5.6), Inches(12.333), Inches(0.5),
             'INDUSTRY PLANNING', size=13, color=DIM,
             align=PP_ALIGN.CENTER, letter_spacing=600, font=EN_FONT)
    add_rect(s, Inches(6.16), Inches(6.3), Inches(1), Inches(0.06), fill=ACCENT2)

    # -------- 10. 产业定位 --------
    s = prs.slides.add_slide(blank)
    content_bg(s)
    page_footer(s, 10, TOTAL, '02 · 产业规划篇')
    kicker(s, '（一）产业定位 · INDUSTRY POSITIONING', color=ACCENT2)
    title(s, '科技潮玩产业策源高地', size=34)

    # 3 goals
    add_text(s, Inches(0.7), Inches(2.4), Inches(6), Inches(0.4),
             '三大战略目标 · 3 STRATEGIC GOALS',
             size=11, color=ACCENT4, letter_spacing=300, bold=True)
    goals = ['国际化产业枢纽', '品牌运营高地', 'IP 创制中心']
    for i, g in enumerate(goals):
        y = Inches(2.9 + i * 0.85)
        add_round_rect(s, Inches(0.7), y, Inches(5.9), Inches(0.7),
                       fill=SURFACE, line=RGBColor(0x22, 0x22, 0x38), radius=0.15)
        add_rect(s, Inches(0.7), y, Inches(0.08), Inches(0.7), fill=ACCENT4)
        add_text(s, Inches(0.95), y, Inches(5.5), Inches(0.7),
                 g, size=18, color=TEXT, bold=True, anchor=MSO_ANCHOR.MIDDLE)

    # 5 advantages
    add_text(s, Inches(7.0), Inches(2.4), Inches(6), Inches(0.4),
             '五大优势聚焦 · 5 STRENGTHS',
             size=11, color=ACCENT3, letter_spacing=300, bold=True)
    advs = ['创意集聚', '源头孵化', '场景体验', '生态复合', '集约选品']
    for i, a in enumerate(advs):
        col = i % 2
        row = i // 2
        x = Inches(7.0 + col * 3.0)
        y = Inches(2.9 + row * 0.85)
        w = Inches(2.85 if i < 4 else 5.85)  # last one wide
        if i == 4:
            x = Inches(7.0); w = Inches(5.85)
        add_round_rect(s, x, y, w, Inches(0.7),
                       fill=SURFACE2, line=RGBColor(0x06, 0xB6, 0xD4), radius=0.15)
        add_text(s, x, y, w, Inches(0.7), a,
                 size=16, color=TEXT, bold=True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Twin drivers
    drivers = [
        ('"产业转化" + "转化产业"', '双轮驱动 · 为吴泾贡献特色地标产业', ACCENT),
        ('"投资驱动" + "市场驱动"', '双管齐下 · 为闵行构建潮玩产业生态', ACCENT2),
    ]
    for i, (h, d, c) in enumerate(drivers):
        x = Inches(0.7 + i * 6.3)
        y = Inches(6.0)
        add_round_rect(s, x, y, Inches(6.0), Inches(0.9),
                       fill=SURFACE, line=c, radius=0.1)
        add_text(s, x + Inches(0.3), y + Inches(0.1), Inches(5.6), Inches(0.4),
                 h, size=13, color=c, bold=True)
        add_text(s, x + Inches(0.3), y + Inches(0.48), Inches(5.6), Inches(0.4),
                 d, size=10, color=DIM)

    # -------- 11. 产业金字塔 --------
    s = prs.slides.add_slide(blank)
    content_bg(s)
    page_footer(s, 11, TOTAL, '02 · 产业规划篇')
    kicker(s, '（二）产业规划 · INDUSTRY ECOSYSTEM', color=ACCENT2)
    title(s, '一体化潮玩产业生态体系', size=30)
    subtitle(s, '依托共享配套形成强大吸附效应，夯实中型企业产业基础，筑牢小型企业发展骨架，机构服务血肉补充。')

    rows = [
        ('10%', '头部央企 · 行业协会 · 导向',    '3 个 · 2000㎡',      0.32, ACCENT),
        ('10%', '共享配套服务体系 · 吸附点',     '3 个 · 2000㎡',      0.45, ACCENT4),
        ('20%', '中型潮玩运营企业 · 基础',       '4–6 个 · 5000㎡',    0.60, ACCENT3),
        ('40%', '小型潮玩运营企业 · 骨架',       '30 个 · 200–500㎡',  0.80, ACCENT2),
        ('20%', '中小型潮玩服务机构 · 血肉',     '15 个 · 200–500㎡',  1.00, ACCENT5),
    ]
    y0 = 2.7
    bar_max = 9.0  # inches
    bar_left = 1.7
    for i, (pct, label, meta, ratio, c) in enumerate(rows):
        y = Inches(y0 + i * 0.72)
        add_text(s, Inches(0.7), y, Inches(0.9), Inches(0.55),
                 pct, size=20, color=c, bold=True,
                 align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE, font=EN_FONT)
        w = bar_max * ratio
        bar = add_round_rect(s, Inches(bar_left), y, Inches(w), Inches(0.55),
                             fill=c, line=None, radius=0.18)
        add_text(s, Inches(bar_left + 0.2), y, Inches(w - 0.4), Inches(0.55),
                 label, size=12, color=BG, bold=True, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(bar_left + w + 0.15), y, Inches(3), Inches(0.55),
                 meta, size=10, color=DIM, anchor=MSO_ANCHOR.MIDDLE)

    # -------- 12. 产业牌照 --------
    s = prs.slides.add_slide(blank)
    content_bg(s)
    page_footer(s, 12, TOTAL, '02 · 产业规划篇')
    kicker(s, '（三）产业牌照 · INDUSTRY LICENSES', color=ACCENT2)
    title(s, '双牌照赋能  ·  共建两大产业平台', size=30)
    subtitle(s, '确认与中国百货协会和中国动漫集团在本项目共同设立两大产业牌照。')

    # license 1
    add_round_rect(s, Inches(0.7), Inches(2.7), Inches(5.9), Inches(4.0),
                   fill=SURFACE, line=RGBColor(0x22, 0x22, 0x38), radius=0.05)
    add_text(s, Inches(1.0), Inches(2.95), Inches(2), Inches(0.7),
             '🏛️', size=40, color=TEXT)
    add_text(s, Inches(1.0), Inches(3.85), Inches(5.5), Inches(0.6),
             '潮玩次元商业专委会', size=22, color=TEXT, bold=True)
    add_text(s, Inches(1.0), Inches(4.55), Inches(5.5), Inches(0.4),
             'CHINA RETAIL · TRENDY-TOY COMMITTEE',
             size=9, color=ACCENT, letter_spacing=300, font=EN_FONT, bold=True)
    add_round_rect(s, Inches(1.0), Inches(5.1), Inches(0.55), Inches(0.32),
                   fill=RGBColor(0x44, 0x18, 0x28), line=None, radius=0.4)
    add_text(s, Inches(1.0), Inches(5.1), Inches(0.55), Inches(0.32),
             '4月', size=9, color=ACCENT, bold=True,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(1.65), Inches(5.1), Inches(4.8), Inches(0.32),
             '中百协潮玩次元商业专委会（筹）启动筹备',
             size=11, color=DIM, anchor=MSO_ANCHOR.MIDDLE)
    add_round_rect(s, Inches(1.0), Inches(5.6), Inches(0.55), Inches(0.32),
                   fill=RGBColor(0x44, 0x18, 0x28), line=None, radius=0.4)
    add_text(s, Inches(1.0), Inches(5.6), Inches(0.55), Inches(0.32),
             '6月', size=9, color=ACCENT, bold=True,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(1.65), Inches(5.6), Inches(4.8), Inches(0.32),
             '济南理事会会议上正式挂牌成立',
             size=11, color=DIM, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(1.0), Inches(6.1), Inches(5.5), Inches(0.4),
             '已开始正常业务往来展开 →',
             size=10, color=ACCENT, bold=True)

    # license 2
    add_round_rect(s, Inches(6.8), Inches(2.7), Inches(5.9), Inches(4.0),
                   fill=RGBColor(0x12, 0x18, 0x28), line=ACCENT3, radius=0.05)
    add_text(s, Inches(7.1), Inches(2.95), Inches(2), Inches(0.7),
             '🤖', size=40, color=TEXT)
    add_text(s, Inches(7.1), Inches(3.85), Inches(5.5), Inches(0.6),
             'AI 潮玩产业基地', size=22, color=TEXT, bold=True)
    add_text(s, Inches(7.1), Inches(4.55), Inches(5.5), Inches(0.4),
             'AI · TRENDY-TOY INDUSTRIAL BASE',
             size=9, color=ACCENT3, letter_spacing=300, font=EN_FONT, bold=True)
    add_text(s, Inches(7.1), Inches(5.05), Inches(5.5), Inches(0.3),
             '合作资源', size=10, color=ACCENT3, letter_spacing=200, bold=True)
    add_text(s, Inches(7.1), Inches(5.32), Inches(5.5), Inches(0.5),
             '中国动漫集团 · 上海交大设计学院 · 闵行科协 · 森马集团',
             size=11, color=DIM, line_spacing=1.4)
    add_text(s, Inches(7.1), Inches(5.85), Inches(5.5), Inches(0.3),
             '未来活动', size=10, color=ACCENT3, letter_spacing=200, bold=True)
    add_text(s, Inches(7.1), Inches(6.12), Inches(5.5), Inches(0.6),
             '全国潮玩设计技能大赛 · 国漫·潮游集 · 动漫新品及跨界产品首发',
             size=11, color=DIM, line_spacing=1.4)

    # -------- 13. 产业配套 --------
    s = prs.slides.add_slide(blank)
    content_bg(s)
    page_footer(s, 13, TOTAL, '02 · 产业规划篇')
    kicker(s, '（四）产业配套 · INDUSTRY SUPPORTING PLATFORMS', color=ACCENT2)
    title(s, '六大产业配套平台', size=30)

    services = [
        ('选品中心',           '1# 4F',
         '华东首个 IP 潮玩选品 + 仓储式销售空间',                          ACCENT),
        ('代运营物流中心',     '森马二期仓库',
         '联动骏耀科技 · 智能仓储 + 物流代运营',                          ACCENT2),
        ('共享直播中心',       '4# 4F',
         '潮玩电商直播间 · 联动绮丽少女女团',                              ACCENT3),
        ('AI 共享设计中心',   '4# 5F',
         '联动高校人才资源 · 引入 AI 科技 · 共建 AI 潮玩设计',           ACCENT4),
        ('AI 共享打样 / DIY', '4# 5F',
         '助力客户快速打样 · 提高面世效率',                                 ACCENT5),
        ('潮玩产业展厅',       '5# 5F',
         '聚焦品牌 IP 叙事 · 强化行业交流与渠道拓展',                       ACCENT),
    ]
    for i, (name, loc, desc, c) in enumerate(services):
        col = i % 3
        row = i // 3
        x = Inches(0.7 + col * 4.1)
        y = Inches(2.85 + row * 2.0)
        add_round_rect(s, x, y, Inches(3.9), Inches(1.8),
                       fill=SURFACE, line=RGBColor(0x22, 0x22, 0x38), radius=0.08)
        # top accent stripe
        add_rect(s, x, y, Inches(3.9), Inches(0.06), fill=c)
        # location pill
        add_round_rect(s, x + Inches(0.3), y + Inches(0.25), Inches(1.6), Inches(0.32),
                       fill=SURFACE2, line=None, radius=0.4)
        add_text(s, x + Inches(0.3), y + Inches(0.25), Inches(1.6), Inches(0.32),
                 loc, size=9, color=c, bold=True, font=EN_FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, letter_spacing=200)
        add_text(s, x + Inches(0.3), y + Inches(0.7), Inches(3.5), Inches(0.45),
                 name, size=16, color=TEXT, bold=True)
        add_text(s, x + Inches(0.3), y + Inches(1.18), Inches(3.5), Inches(0.55),
                 desc, size=10, color=DIM, line_spacing=1.45)

    # -------- 14. 产业服务三轨 --------
    s = prs.slides.add_slide(blank)
    content_bg(s)
    page_footer(s, 14, TOTAL, '02 · 产业规划篇')
    kicker(s, '（五）产业服务 · OPERATION STRATEGIES', color=ACCENT2)
    title(s, '三轨并行  ·  创新园区服务综合解决方案', size=30)
    subtitle(s, '吸引企业集聚 · 培育产业发展。')

    strategies = [
        ('01', '企业主导自营',
         '保障配套服务提供能力 · 自持综合服务中心、园区餐饮、商务办公、设计工作室、人才/财务服务中心。',
         ACCENT),
        ('02', '绿色通道协助',
         '发挥行政力量能动性 · 申请街道政府开启绿色通道 · 行政审批 / 企业代办 / 申报协调。',
         ACCENT2),
        ('03', '机构专业运营',
         '引入专业潮玩产业服务机构 · 客户招聘、人力资源、人才培训、财务服务、科技金融等专业服务。',
         ACCENT3),
    ]
    for i, (n, t, d, c) in enumerate(strategies):
        x = Inches(0.7 + i * 4.1)
        y = Inches(2.7)
        h = Inches(4.0 if i == 1 else 3.6)
        if i == 1:
            y = Inches(2.5)
        add_round_rect(s, x, y, Inches(3.9), h,
                       fill=SURFACE if i != 1 else RGBColor(0x1E, 0x10, 0x24),
                       line=c if i == 1 else RGBColor(0x22, 0x22, 0x38), radius=0.06)
        add_text(s, x + Inches(0.4), y + Inches(0.4), Inches(2), Inches(0.6),
                 n, size=32, color=c, bold=True, font=EN_FONT)
        add_text(s, x + Inches(0.4), y + Inches(1.2), Inches(3.2), Inches(0.6),
                 t, size=20, color=TEXT, bold=True)
        # divider
        add_rect(s, x + Inches(0.4), y + Inches(1.85), Inches(0.6), Inches(0.04), fill=c)
        add_text(s, x + Inches(0.4), y + Inches(2.05), Inches(3.2), Inches(2),
                 d, size=11, color=DIM, line_spacing=1.7)

    # ============================================================
    # CHAPTER 03
    # ============================================================
    s = prs.slides.add_slide(blank)
    section_bg(s, ACCENT3)
    add_text(s, Inches(0.5), Inches(0.6), Inches(12), Inches(0.4),
             'CHAPTER 03', size=11, color=ACCENT3,
             align=PP_ALIGN.CENTER, letter_spacing=600, bold=True, font=EN_FONT)
    add_text(s, Inches(0.5), Inches(1.3), Inches(12.333), Inches(2.6),
             '03', size=240, color=TEXT, bold=True,
             align=PP_ALIGN.CENTER, font=EN_FONT, line_spacing=0.95)
    add_text(s, Inches(0.5), Inches(4.7), Inches(12.333), Inches(0.8),
             '商业规划篇', size=44, color=TEXT, bold=True,
             align=PP_ALIGN.CENTER, line_spacing=1.1)
    add_text(s, Inches(0.5), Inches(5.6), Inches(12.333), Inches(0.5),
             'COMMERCIAL PLANNING', size=13, color=DIM,
             align=PP_ALIGN.CENTER, letter_spacing=600, font=EN_FONT)
    add_rect(s, Inches(6.16), Inches(6.3), Inches(1), Inches(0.06), fill=ACCENT3)

    # -------- 16. 商业定位 大字 --------
    s = prs.slides.add_slide(blank)
    content_bg(s)
    page_footer(s, 16, TOTAL, '03 · 商业规划篇')
    kicker(s, '（一）商业定位 · COMMERCIAL POSITIONING', color=ACCENT3)
    add_text(s, Inches(0.5), Inches(2.0), Inches(12.333), Inches(1.4),
             [[
                 {'text': 'Z·世代', 'color': ACCENT, 'size': 76, 'bold': True},
                 {'text': '潮玩社交主场', 'color': TEXT, 'size': 76, 'bold': True},
             ]], align=PP_ALIGN.CENTER, line_spacing=1.0)
    add_text(s, Inches(0.5), Inches(3.6), Inches(12.333), Inches(0.8),
             '潮玩元宇宙  ·  青年引力场',
             size=34, color=TEXT, align=PP_ALIGN.CENTER, bold=True, line_spacing=1.1)
    add_text(s, Inches(0.5), Inches(4.5), Inches(12.333), Inches(0.6),
             '潮玩艺术爱好者聚集地',
             size=24, color=DIM, align=PP_ALIGN.CENTER, letter_spacing=200)

    # decorative bottom bar
    add_rect(s, Inches(5.66), Inches(5.7), Inches(0.66), Inches(0.06), fill=ACCENT)
    add_rect(s, Inches(6.32), Inches(5.7), Inches(0.66), Inches(0.06), fill=ACCENT2)
    add_rect(s, Inches(6.98), Inches(5.7), Inches(0.66), Inches(0.06), fill=ACCENT3)

    # -------- 17. 业态落位 表 --------
    s = prs.slides.add_slide(blank)
    content_bg(s)
    page_footer(s, 17, TOTAL, '03 · 商业规划篇')
    kicker(s, '（二）业态落位 · VERTICAL PROGRAM', color=ACCENT3)
    title(s, '楼层业态垂直分布', size=28)

    floors = ['6F+', '5F', '4F', '3F', '2F', '1F', 'B1/B2']
    program = [
        # 1#, 2#, 3#, 4#, 5#, 6#
        ['森马总部办公', '—',           '酒店',         '潮玩产业办公',  '潮玩产业办公', '—'],
        ['森马总部办公', '—',           '酒店',         '潮玩产业办公',  '潮玩产业办公', '商务宴请'],
        ['★ 旗舰零售',  '二次元 Live', '休闲运动',     '★ 直播基地',  '动漫书店',   '特色餐饮'],
        ['★ 旗舰零售',  '二次元 Live', '萌宠空间',     '★ 潮玩艺术',  '动漫书店',   '品质生活'],
        ['★ 旗舰零售',  '二次元 Live', '休闲运动',     '★ 潮玩艺术',  'IP 潮玩零售', '服务配套'],
        ['★ IP 潮玩街区', '', '', '', '', ''],   # span
        ['停车场', '', '', '', '', ''],            # span
    ]

    # Header
    col_x = [0.7, 1.7, 3.5, 5.3, 7.1, 8.9, 10.7]  # 7 cols boundaries
    col_w = [1.0, 1.8, 1.8, 1.8, 1.8, 1.8, 1.8]
    header_y = Inches(2.4)
    headers = ['楼层', '1#', '2#', '3#', '4#', '5#', '6#']
    for i, h in enumerate(headers):
        x = Inches(col_x[i])
        add_rect(s, x, header_y, Inches(col_w[i]), Inches(0.45),
                 fill=RGBColor(0x22, 0x18, 0x30), line=None)
        add_text(s, x, header_y, Inches(col_w[i]), Inches(0.45), h,
                 size=11, color=TEXT, bold=True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    row_h = 0.52
    for ri, floor in enumerate(floors):
        y = Inches(2.85 + ri * row_h)
        # floor cell
        add_rect(s, Inches(col_x[0]), y, Inches(col_w[0]), Inches(row_h),
                 fill=RGBColor(0x12, 0x12, 0x20), line=RGBColor(0x22, 0x22, 0x38))
        add_text(s, Inches(col_x[0]), y, Inches(col_w[0]), Inches(row_h),
                 floor, size=12, color=ACCENT, bold=True, font=EN_FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        if ri == 5:  # 1F  full span IP 潮玩街区
            x = Inches(col_x[1])
            tot_w = sum(col_w[1:])
            add_rect(s, x, y, Inches(tot_w), Inches(row_h),
                     fill=RGBColor(0x44, 0x18, 0x44), line=RGBColor(0x33, 0x33, 0x4A))
            add_text(s, x, y, Inches(tot_w), Inches(row_h),
                     '★ IP 潮玩街区', size=14, color=TEXT, bold=True,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        elif ri == 6:  # B1/B2 parking span
            x = Inches(col_x[1])
            tot_w = sum(col_w[1:])
            add_rect(s, x, y, Inches(tot_w), Inches(row_h),
                     fill=RGBColor(0x18, 0x18, 0x28), line=RGBColor(0x22, 0x22, 0x38))
            add_text(s, x, y, Inches(tot_w), Inches(row_h),
                     '停车场 · PARKING', size=11, color=DIM,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, letter_spacing=300)
        else:
            for ci, val in enumerate(program[ri]):
                x = Inches(col_x[ci + 1])
                w = Inches(col_w[ci + 1])
                fill = SURFACE if ri % 2 == 0 else RGBColor(0x10, 0x10, 0x1C)
                if val.startswith('★'):
                    fill = RGBColor(0x2A, 0x14, 0x30)
                add_rect(s, x, y, w, Inches(row_h),
                         fill=fill, line=RGBColor(0x22, 0x22, 0x38))
                col = ACCENT if val.startswith('★') else (MUTE if val == '—' else TEXT)
                add_text(s, x + Inches(0.05), y, w - Inches(0.1), Inches(row_h),
                         val, size=10, color=col,
                         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                         bold=val.startswith('★'))

    add_text(s, Inches(0.7), Inches(6.8), Inches(12), Inches(0.3),
             '★ 主力业态 · ANCHOR FORMAT',
             size=9, color=ACCENT, letter_spacing=200, bold=True)

    # -------- 18. 主力业态 总览 --------
    s = prs.slides.add_slide(blank)
    content_bg(s)
    page_footer(s, 18, TOTAL, '03 · 商业规划篇')
    kicker(s, '（三）主力业态 · ANCHOR FORMATS', color=ACCENT3)
    title(s, '五大主力业态 · 锚定潮玩之城', size=30)

    anchors = [
        ('01', '动漫潮玩谷主题街区',  '~3,000 ㎡', '南上海首个动漫潮玩主题商业街区', ACCENT),
        ('02', 'IP 潮玩选品 & 仓储零售', '~5,000 ㎡', '华东首个 IP 潮玩选品 + 仓储零售中心', ACCENT2),
        ('03', '潮玩艺术中心',         '2,000 ㎡ · 4# 1–3F', '南上海首个特色艺术文化体验空间', ACCENT3),
        ('04', '森马展厅 & 二次元 Livehouse', '700 ㎡', '国内首个二次元主题 Livehouse', ACCENT4),
        ('05', '动漫主题书店',         '1,500 ㎡', '南上海首个动漫主题书店 · 展售 + 签售 + 咖啡', ACCENT5),
    ]
    for i, (n, name, area, hl, c) in enumerate(anchors):
        if i < 3:
            x = Inches(0.7 + i * 4.1)
            y = Inches(2.7)
            w = Inches(3.9); h = Inches(2.3)
        else:
            x = Inches(0.7 + (i - 3) * 6.15 + (0 if i < 4 else 0))
            x = Inches(0.7 + (i - 3) * 6.15)
            y = Inches(5.15)
            w = Inches(5.95); h = Inches(1.6)
        add_round_rect(s, x, y, w, h,
                       fill=SURFACE, line=RGBColor(0x22, 0x22, 0x38), radius=0.06)
        add_rect(s, x, y, w, Inches(0.06), fill=c)
        add_text(s, x + Inches(0.3), y + Inches(0.25), Inches(1.5), Inches(0.4),
                 n, size=14, color=c, bold=True, font=EN_FONT, letter_spacing=300)
        add_text(s, x + Inches(0.3), y + Inches(0.65), w - Inches(0.6), Inches(0.5),
                 name, size=15, color=TEXT, bold=True, line_spacing=1.25)
        add_round_rect(s, x + Inches(0.3), y + Inches(1.2), Inches(2.2), Inches(0.32),
                       fill=SURFACE2, line=None, radius=0.5)
        add_text(s, x + Inches(0.3), y + Inches(1.2), Inches(2.2), Inches(0.32),
                 area, size=9, color=c, bold=True, font=EN_FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, letter_spacing=200)
        add_text(s, x + Inches(0.3), y + Inches(1.62), w - Inches(0.6), Inches(0.55),
                 hl, size=10, color=DIM, line_spacing=1.45)

    # ============================================================
    # CHAPTER 04
    # ============================================================
    s = prs.slides.add_slide(blank)
    section_bg(s, ACCENT4)
    add_text(s, Inches(0.5), Inches(0.6), Inches(12), Inches(0.4),
             'CHAPTER 04', size=11, color=ACCENT4,
             align=PP_ALIGN.CENTER, letter_spacing=600, bold=True, font=EN_FONT)
    add_text(s, Inches(0.5), Inches(1.3), Inches(12.333), Inches(2.6),
             '04', size=240, color=TEXT, bold=True,
             align=PP_ALIGN.CENTER, font=EN_FONT, line_spacing=0.95)
    add_text(s, Inches(0.5), Inches(4.7), Inches(12.333), Inches(0.8),
             '运营合作篇', size=44, color=TEXT, bold=True,
             align=PP_ALIGN.CENTER, line_spacing=1.1)
    add_text(s, Inches(0.5), Inches(5.6), Inches(12.333), Inches(0.5),
             'OPERATIONS & COOPERATION', size=13, color=DIM,
             align=PP_ALIGN.CENTER, letter_spacing=600, font=EN_FONT)
    add_rect(s, Inches(6.16), Inches(6.3), Inches(1), Inches(0.06), fill=ACCENT4)

    # -------- 20. 三大活动 --------
    s = prs.slides.add_slide(blank)
    content_bg(s)
    page_footer(s, 20, TOTAL, '04 · 运营合作篇')
    kicker(s, '三大旗舰活动 · FLAGSHIP EVENTS', color=ACCENT4)
    title(s, '从垂类热度  延展广域客群', size=30)

    events = [
        ('（一）特色集会', '潮玩集市',
         '📍 中心广场',
         'IP 艺术展 · 次元市集 · 嘉年华 · 动漫展会 · 二次元路演 · 品牌出海交流会',
         '', ACCENT),
        ('（二）产品首发', '动漫新品 & 跨界首发会',
         '📍 中心广场 · 4#',
         '年度全新动漫内容 · 原创 IP 形象 · 潮玩手办 · 周边衍生 · 多品类跨界联名',
         '现场体验 + 互动打卡 + 合作签约', ACCENT2),
        ('（三）主题大赛', '全国潮玩设计大赛',
         '📍 中心广场',
         '原创 IP 形象 · 手办盲盒 · 国潮文创 · 数字潮玩',
         '合作单位：中国动漫集团 · 上海交大设计学院', ACCENT3),
    ]
    for i, (k, name, loc, desc, foot, c) in enumerate(events):
        x = Inches(0.7 + i * 4.1)
        y = Inches(2.7)
        add_round_rect(s, x, y, Inches(3.9), Inches(4.1),
                       fill=SURFACE if i != 2 else RGBColor(0x1E, 0x14, 0x10),
                       line=c if i == 2 else RGBColor(0x22, 0x22, 0x38), radius=0.06)
        add_round_rect(s, x + Inches(0.3), y + Inches(0.3), Inches(1.7), Inches(0.3),
                       fill=RGBColor(0x44, 0x2A, 0x10), line=None, radius=0.5)
        add_text(s, x + Inches(0.3), y + Inches(0.3), Inches(1.7), Inches(0.3),
                 k, size=8, color=ACCENT4, bold=True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, letter_spacing=200)
        add_text(s, x + Inches(0.3), y + Inches(0.75), Inches(3.4), Inches(0.7),
                 name, size=20, color=TEXT, bold=True, line_spacing=1.2)
        add_text(s, x + Inches(0.3), y + Inches(1.6), Inches(3.4), Inches(0.35),
                 loc, size=10, color=c, bold=True)
        add_rect(s, x + Inches(0.3), y + Inches(2.0), Inches(0.5), Inches(0.04), fill=c)
        add_text(s, x + Inches(0.3), y + Inches(2.15), Inches(3.4), Inches(1.5),
                 desc, size=11, color=DIM, line_spacing=1.7)
        if foot:
            add_text(s, x + Inches(0.3), y + Inches(3.55), Inches(3.4), Inches(0.5),
                     foot, size=10, color=TEXT, bold=True, line_spacing=1.4)

    # -------- 21. 拟合作品牌 --------
    s = prs.slides.add_slide(blank)
    content_bg(s)
    page_footer(s, 21, TOTAL, '04 · 运营合作篇')
    kicker(s, '（四）拟合作品牌 · PARTNER MATRIX', color=ACCENT4)
    title(s, '生态合作伙伴矩阵', size=30)

    cols = [
        ('IP 潮玩 / 二次元',
         ['绮丽少女', '秋子 ACG 超级贩卖仓', '轻语有品', '宏腾玩具', '超级悦沢', '樱漫书店'],
         ACCENT),
        ('产业合作机构',
         ['中国百货协会', '中国动漫集团', '上海交大设计学院', '闵行科协', '森马集团', '骏耀科技'],
         ACCENT2),
        ('周边产业生态',
         ['森马产业', '衣恋', '普洛斯', '雅诗兰黛', '保利', '中建产研'],
         ACCENT3),
    ]
    for i, (head, items, c) in enumerate(cols):
        x = Inches(0.7 + i * 4.1)
        y = Inches(2.7)
        add_round_rect(s, x, y, Inches(3.9), Inches(4.1),
                       fill=SURFACE, line=RGBColor(0x22, 0x22, 0x38), radius=0.06)
        add_text(s, x + Inches(0.3), y + Inches(0.25), Inches(3.4), Inches(0.4),
                 head, size=11, color=c, letter_spacing=300, bold=True)
        add_rect(s, x + Inches(0.3), y + Inches(0.7), Inches(0.6), Inches(0.04), fill=c)
        for j, it in enumerate(items):
            ty = y + Inches(0.95 + j * 0.5)
            add_oval(s, x + Inches(0.35), ty + Inches(0.13), Inches(0.1), Inches(0.1),
                     fill=c, line=None)
            add_text(s, x + Inches(0.6), ty, Inches(3.0), Inches(0.45),
                     it, size=13, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)

    # -------- 22. 项目核心价值 总结 --------
    s = prs.slides.add_slide(blank)
    content_bg(s)
    page_footer(s, 22, TOTAL)
    kicker(s, '核心价值 · CORE VALUE')
    title(s, '一站式潮玩产业  ·  全要素城市级解决方案', size=28)
    subtitle(s, '从战略卡位到生态协同，从空间叙事到客群运营，森马国际运营中心提供完整答案。')

    cells = [
        ('🏛️', '战略级卡位',       '南部科创走廊 + 大零号湾 + 长三角', ACCENT),
        ('🚄', 'TOD 级流量',       '15 号线元江路日均 5–7 万人次',     ACCENT2),
        ('🏢', '22 万㎡ 综合体',    '6 栋楼宇 · 商业 5.2 万㎡',         ACCENT3),
        ('🎯', '产业级牌照',       '中百协 + 中国动漫集团 双牌照',     ACCENT4),
        ('🤖', 'AI · IP 全链路',  '设计 / 打样 / 选品 / 直播 / 仓储',  ACCENT5),
        ('🎉', '潮玩级活动',       '集市 / 首发 / 大赛 持续引流',       ACCENT),
    ]
    for i, (icon, name, desc, c) in enumerate(cells):
        col = i % 3
        row = i // 3
        x = Inches(0.7 + col * 4.1)
        y = Inches(2.7 + row * 2.0)
        add_round_rect(s, x, y, Inches(3.9), Inches(1.85),
                       fill=SURFACE, line=RGBColor(0x22, 0x22, 0x38), radius=0.06)
        add_text(s, x + Inches(0.3), y + Inches(0.3), Inches(0.8), Inches(0.8),
                 icon, size=28, color=TEXT)
        add_text(s, x + Inches(1.2), y + Inches(0.3), Inches(2.6), Inches(0.4),
                 name, size=15, color=c, bold=True)
        add_text(s, x + Inches(1.2), y + Inches(0.75), Inches(2.6), Inches(1.0),
                 desc, size=10, color=DIM, line_spacing=1.5)

    # -------- 23. Quote / 战略愿景 --------
    s = prs.slides.add_slide(blank)
    section_bg(s, ACCENT2)
    add_text(s, Inches(1.0), Inches(2.2), Inches(11.333), Inches(1.0),
             '"', size=120, color=ACCENT, bold=True,
             align=PP_ALIGN.CENTER, line_spacing=0.8, font=EN_FONT)
    add_text(s, Inches(1.0), Inches(3.3), Inches(11.333), Inches(2.0),
             '不只是一座综合体，\n而是一片 Z·世代心目中的潮玩元宇宙。',
             size=32, color=TEXT, bold=True,
             align=PP_ALIGN.CENTER, line_spacing=1.4)
    add_text(s, Inches(1.0), Inches(5.4), Inches(11.333), Inches(0.5),
             '— SEMIR GROUP\'S GLOBAL HEADQUARTER',
             size=12, color=DIM, align=PP_ALIGN.CENTER, letter_spacing=400, font=EN_FONT)

    # -------- 24. THANKS --------
    s = prs.slides.add_slide(blank)
    hero_bg(s)
    add_text(s, Inches(0.5), Inches(2.4), Inches(12.333), Inches(2.5),
             [[
                 {'text': 'THA', 'color': ACCENT, 'size': 200, 'bold': True, 'font': EN_FONT},
                 {'text': 'NK',  'color': ACCENT2, 'size': 200, 'bold': True, 'font': EN_FONT},
                 {'text': 'S',   'color': ACCENT3, 'size': 200, 'bold': True, 'font': EN_FONT},
             ]], align=PP_ALIGN.CENTER, line_spacing=1.0)
    add_text(s, Inches(0.5), Inches(5.0), Inches(12.333), Inches(0.6),
             '森马（上海）国际运营中心  ·  SEMIR GROUP\'S GLOBAL HEADQUARTER',
             size=14, color=TEXT, align=PP_ALIGN.CENTER, letter_spacing=300)
    add_text(s, Inches(0.5), Inches(5.7), Inches(12.333), Inches(0.5),
             'SHANGHAI  ·  MINHANG  ·  WUJING  ·  15 号线元江路 TOD',
             size=10, color=DIM, align=PP_ALIGN.CENTER, letter_spacing=500, font=EN_FONT)
    # divider
    add_rect(s, Inches(6.16), Inches(6.5), Inches(1), Inches(0.04), fill=ACCENT)

    # -------- save --------
    out_path = os.path.join(os.path.dirname(__file__), '..', 'docs',
                            '森马国际运营中心-项目介绍.pptx')
    out_path = os.path.normpath(out_path)
    prs.save(out_path)
    print(f'PPT saved: {out_path}  ·  slides: {len(prs.slides)}')


if __name__ == '__main__':
    build()
