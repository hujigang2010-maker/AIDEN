# -*- coding: utf-8 -*-
"""
上海市云计算创新基地 · 招商合作版宣传册（PPT）生成脚本

特点：
- 16:9 商务大气版式，深海军蓝 + 科技青 + 鎏金 高端配色
- 图文结合，多图：封面/价值/三位一体/产业实力/明星企业/产业方向/服务/合作/历程/封底
- 结构按"招商转化型"重排：价值 → 证明 → 资源 → 入口
- 所有中文文字由 PPT 文本框渲染（清晰可编辑），图片仅作背景与配图

运行: python3 build_deck.py
输出: dist/上海云基地-招商合作版宣传册.pptx
"""

import os
from PIL import Image, ImageEnhance, ImageFilter

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ----------------------------------------------------------------------------
# 配色与字体
# ----------------------------------------------------------------------------
NAVY      = "0A1A33"   # 主背景深蓝
NAVY_DK   = "061123"   # 更深
NAVY_CARD = "112B4A"   # 卡片底
CYAN      = "33C9FF"   # 科技青（主强调）
CYAN_DK   = "1E88C7"
GOLD      = "E2B663"   # 鎏金（高端点缀）
WHITE     = "FFFFFF"
MUTED     = "A9C0D6"   # 次要文字

FONT      = "Microsoft YaHei"   # 微软雅黑：通用中文商务字体

ASSETS = os.path.join(os.path.dirname(__file__), "assets")
BUILD  = os.path.join(os.path.dirname(__file__), "build")
DIST   = os.path.join(os.path.dirname(__file__), "dist")
os.makedirs(BUILD, exist_ok=True)
os.makedirs(DIST, exist_ok=True)

EMU_PER_IN = 914400
SLIDE_W_IN = 13.333
SLIDE_H_IN = 7.5

# ----------------------------------------------------------------------------
# 图片预处理：中心裁切为 16:9，并可选压暗，保证文字可读
# ----------------------------------------------------------------------------
def prep_bg(name, darken=1.0, blur=0):
    """裁切为 16:9 输出到 build/ 并返回路径。darken<1 压暗，blur 模糊半径。"""
    src = os.path.join(ASSETS, name)
    im = Image.open(src).convert("RGB")
    target = 16 / 9
    w, h = im.size
    if w / h > target:          # 太宽 -> 裁左右
        nw = int(h * target)
        left = (w - nw) // 2
        im = im.crop((left, 0, left + nw, h))
    else:                        # 太高 -> 裁上下
        nh = int(w / target)
        top = (h - nh) // 2
        im = im.crop((0, top, w, top + nh))
    im = im.resize((1920, 1080), Image.LANCZOS)
    if blur:
        im = im.filter(ImageFilter.GaussianBlur(blur))
    if darken != 1.0:
        im = ImageEnhance.Brightness(im).enhance(darken)
    out = os.path.join(BUILD, name)
    im.save(out, quality=92)
    return out

# ----------------------------------------------------------------------------
# 低层工具
# ----------------------------------------------------------------------------
prs = Presentation()
prs.slide_width = Inches(SLIDE_W_IN)
prs.slide_height = Inches(SLIDE_H_IN)
BLANK = prs.slide_layouts[6]


def add_slide():
    return prs.slides.add_slide(BLANK)


def _set_run_font(run, name, size, color, bold, italic=False, spacing=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)
    rPr = run._r.get_or_add_rPr()
    # 东亚字体
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', name)
    cs = rPr.find(qn('a:cs'))
    if cs is None:
        cs = rPr.makeelement(qn('a:cs'), {})
        rPr.append(cs)
    cs.set('typeface', name)
    if spacing is not None:
        rPr.set('spc', str(int(spacing * 100)))


def add_text(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, line_spacing=1.0, space_after=0,
             wrap=True, para_gap=None):
    """runs: list of paragraphs; each paragraph is list of (text,font,size,color,bold[,italic])."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
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
        if line_spacing:
            p.line_spacing = line_spacing
        if space_after:
            p.space_after = Pt(space_after)
        if para_gap is not None:
            p.space_before = Pt(para_gap if i > 0 else 0)
        for seg in para:
            text, fnt, size, color, bold = seg[0], seg[1], seg[2], seg[3], seg[4]
            italic = seg[5] if len(seg) > 5 else False
            spacing = seg[6] if len(seg) > 6 else None
            r = p.add_run()
            r.text = text
            _set_run_font(r, fnt, size, color, bold, italic, spacing)
    return tb


def add_rect(slide, x, y, w, h, fill=None, alpha=None, line_color=None,
             line_w=None, shape=MSO_SHAPE.RECTANGLE, shadow=False):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = RGBColor.from_string(fill)
        if alpha is not None:
            _apply_alpha(sp.fill.fore_color, alpha)
    if line_color is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = RGBColor.from_string(line_color)
        sp.line.width = Pt(line_w or 1)
    sp.shadow.inherit = False
    if shadow:
        _soft_shadow(sp)
    return sp


def _apply_alpha(fore_color, alpha_pct):
    """alpha_pct: 0-100 不透明度。"""
    srgb = fore_color._xFill.find(qn('a:srgbClr'))
    a = srgb.makeelement(qn('a:alpha'), {'val': str(int(alpha_pct * 1000))})
    srgb.append(a)


def _soft_shadow(sp):
    spPr = sp._element.spPr
    effLst = spPr.makeelement(qn('a:effectLst'), {})
    shdw = effLst.makeelement(qn('a:outerShdw'),
                              {'blurRad': '90000', 'dist': '40000',
                               'dir': '5400000', 'rotWithShape': '0'})
    clr = shdw.makeelement(qn('a:srgbClr'), {'val': '000000'})
    alpha = clr.makeelement(qn('a:alpha'), {'val': '42000'})
    clr.append(alpha)
    shdw.append(clr)
    effLst.append(shdw)
    spPr.append(effLst)


def full_bg(slide, img):
    slide.shapes.add_picture(img, 0, 0, Inches(SLIDE_W_IN), Inches(SLIDE_H_IN))


def overlay(slide, x, y, w, h, color, alpha):
    add_rect(slide, x, y, w, h, fill=color, alpha=alpha)


def gradient_left(slide, color=NAVY_DK, width=8.2, alpha=82):
    """左侧渐隐遮罩（用一组阶梯矩形模拟），保证左侧文字清晰。"""
    steps = 14
    for i in range(steps):
        a = alpha * (1 - i / steps)
        add_rect(slide, width * i / steps, 0, width / steps + 0.02,
                 SLIDE_H_IN, fill=color, alpha=max(a, 0))


def kicker(slide, x, y, text, color=CYAN):
    """小标题前的装饰条 + 英文/分类标签。"""
    add_rect(slide, x, y + 0.05, 0.42, 0.07, fill=color)
    add_text(slide, x + 0.55, y - 0.16, 6, 0.4,
             [[(text, FONT, 13, color, True, False, 2)]])


def section_title(slide, x, y, zh, color=WHITE, size=30):
    add_text(slide, x, y, 11, 0.9, [[(zh, FONT, size, color, True)]])


def page_no(slide, n, total, light=False):
    c = MUTED if not light else "5B7287"
    add_text(slide, SLIDE_W_IN - 1.4, SLIDE_H_IN - 0.55, 1.1, 0.35,
             [[(f"{n:02d} / {total:02d}", FONT, 10, c, False)]],
             align=PP_ALIGN.RIGHT)
    add_text(slide, 0.6, SLIDE_H_IN - 0.55, 5, 0.35,
             [[("上海云基地  SHANGHAI CLOUD BASE", FONT, 10, c, False, False, 1)]])


TOTAL = 11

# ============================================================================
# Slide 1 · 封面
# ============================================================================
def slide_cover():
    s = add_slide()
    full_bg(s, prep_bg("cover.png"))
    gradient_left(s, NAVY_DK, width=8.4, alpha=80)
    # 顶部资质条
    add_rect(s, 0.85, 0.7, 0.12, 0.95, fill=GOLD)
    add_text(s, 1.1, 0.62, 8, 0.5,
             [[("国家级科技企业孵化器（专业类）", FONT, 14, GOLD, True, False, 1)]])
    add_text(s, 1.1, 1.06, 9, 0.5,
             [[("National Professional Technology Business Incubator", FONT, 11, MUTED, False, False, 1)]])
    # 主标题
    add_text(s, 0.82, 2.55, 9.2, 2.2, [
        [("上海市云计算创新基地", FONT, 52, WHITE, True)],
        [("上 海 云 基 地", FONT, 26, CYAN, True, False, 6)],
    ], line_spacing=1.05)
    # 一句话钩子
    add_rect(s, 0.9, 4.55, 0.42, 0.07, fill=CYAN)
    add_text(s, 0.86, 4.78, 10.5, 1.5, [
        [("中国云计算与 AI 产业创新策源地", FONT, 22, WHITE, True)],
        [("让技术企业从 ", FONT, 16, MUTED, False),
         ("0 → 1 → 上市", FONT, 16, GOLD, True),
         (" 的产业加速器", FONT, 16, MUTED, False)],
    ], line_spacing=1.2, para_gap=10)
    # 底部主体
    add_text(s, 0.86, SLIDE_H_IN - 0.7, 9, 0.4,
             [[("上海市杨浦云计算创新基地发展有限公司  ·  招商合作版", FONT, 11, MUTED, False, False, 1)]])


# ============================================================================
# Slide 2 · 一句话价值
# ============================================================================
def slide_value():
    s = add_slide()
    full_bg(s, prep_bg("value.png", darken=0.92))
    overlay(s, 0, 0, SLIDE_W_IN, SLIDE_H_IN, NAVY_DK, 30)
    gradient_left(s, NAVY_DK, width=8.6, alpha=86)
    kicker(s, 0.85, 0.95, "01  我们是谁 · WHO WE ARE")
    add_text(s, 0.82, 1.7, 9.2, 2.4, [
        [("把「技术公司」", FONT, 40, WHITE, True)],
        [("变成「产业公司」的地方", FONT, 40, CYAN, True)],
    ], line_spacing=1.12)
    add_text(s, 0.86, 3.95, 8.6, 2.2, [
        [("我们解决什么问题 —— ", FONT, 15, GOLD, True),
         ("帮助硬科技企业跨越从技术到产业的鸿沟。", FONT, 15, WHITE, False)],
        [("依托杨浦高校科研资源，首创 ", FONT, 14, MUTED, False),
         ("「把上下楼变上下游」", FONT, 14, CYAN, True),
         (" 理念，", FONT, 14, MUTED, False)],
        [("以「基金 + 基地 + 服务平台」连接 ", FONT, 14, MUTED, False),
         ("技术、用户、资本、服务", FONT, 14, WHITE, True),
         (" 四大生态，", FONT, 14, MUTED, False)],
        [("构建融合「政产学研投用」的产业生态系统。", FONT, 14, MUTED, False)],
    ], line_spacing=1.35, para_gap=8)
    page_no(s, 2, TOTAL)


# ============================================================================
# Slide 3 · 为什么选我们（核心卖点）
# ============================================================================
def slide_why():
    s = add_slide()
    full_bg(s, prep_bg("texture.png", darken=0.78))
    kicker(s, 0.85, 0.7, "02  为什么选我们 · WHY US")
    section_title(s, 0.82, 1.15, "四大核心优势，一站赋能产业成长")
    cards = [
        ("政　策", "国家级孵化器", "国家级科技企业孵化器（专业类），2025 拟推荐工信部卓越级，政策与资质背书。"),
        ("资　本", "云海 + 云天使", "自有「云海创投」1.5 亿元，联合发起 3 亿元「云天使基金」，投资 + 孵化双驱动。"),
        ("产　业", "近 200 家企业生态", "15 年沉淀，从初创孵化到上市扩张的完整链条，2025 年企业总产值超 60 亿元。"),
        ("场　景", "方案孵化器", "为创新方案提供验证基础、应用资源对接与实验测试，让技术快速跑通真实场景。"),
    ]
    cw, gap = 2.86, 0.28
    x0 = 0.85
    y0 = 2.15
    ch = 4.35
    for i, (tag, title, desc) in enumerate(cards):
        x = x0 + i * (cw + gap)
        add_rect(s, x, y0, cw, ch, fill=NAVY_CARD, alpha=86, shadow=True)
        add_rect(s, x, y0, cw, 0.1, fill=CYAN if i % 2 == 0 else GOLD)
        add_text(s, x + 0.32, y0 + 0.45, cw - 0.6, 0.6,
                 [[(f"0{i+1}", FONT, 30, GOLD if i % 2 else CYAN, True)]])
        add_text(s, x + 0.32, y0 + 1.25, cw - 0.6, 0.5,
                 [[(tag, FONT, 13, CYAN if i % 2 == 0 else GOLD, True, False, 4)]])
        add_text(s, x + 0.32, y0 + 1.62, cw - 0.6, 0.7,
                 [[(title, FONT, 19, WHITE, True)]])
        add_rect(s, x + 0.32, y0 + 2.28, 0.7, 0.035, fill=MUTED, alpha=60)
        add_text(s, x + 0.32, y0 + 2.5, cw - 0.62, 1.7,
                 [[(desc, FONT, 12.5, MUTED, False)]], line_spacing=1.3)
    page_no(s, 3, TOTAL)


# ============================================================================
# Slide 4 · 三位一体模式
# ============================================================================
def slide_trinity():
    s = add_slide()
    full_bg(s, prep_bg("trinity.png", darken=0.62))
    overlay(s, 0, 0, SLIDE_W_IN, SLIDE_H_IN, NAVY_DK, 28)
    kicker(s, 0.85, 0.7, "03  运营模式 · OPERATING MODEL")
    section_title(s, 0.82, 1.15, "「三位一体」全生命周期赋能")
    # 流程：基金 → 基地 → 服务
    chain = [("基金", "钱", CYAN), ("基地", "空间", GOLD), ("服务", "能力", CYAN)]
    bx, by, bw, bh = 0.85, 2.2, 2.4, 0.95
    gap = 0.95
    for i, (a, b, c) in enumerate(chain):
        x = bx + i * (bw + gap)
        add_rect(s, x, by, bw, bh, fill=NAVY_CARD, alpha=92, line_color=c, line_w=1.5)
        add_text(s, x, by + 0.13, bw, bh,
                 [[(a, FONT, 22, WHITE, True), ("（", FONT, 14, MUTED, False),
                   (b, FONT, 14, c, True), ("）", FONT, 14, MUTED, False)]],
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        if i < 2:
            add_text(s, x + bw, by, gap, bh,
                     [[("→", FONT, 26, GOLD, True)]],
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 三个平台详述
    cols = [
        ("产业孵化平台", CYAN, [
            "累计孵化、引进 100 余家科技企业",
            "2025 年企业总产值超 60 亿元",
            "从初创到上市的完整培育链条",
        ]),
        ("产业投资平台", GOLD, [
            "「云海创投」基金 1.5 亿元",
            "作为 LP 参与 3 亿元「云天使基金」",
            "聚焦云计算、网络安全早期项目",
        ]),
        ("产业服务平台", CYAN, [
            "云计算创新展示中心",
            "五角场创新创业学院",
            "云基地方案孵化器（市场·技术·场景·人才）",
        ]),
    ]
    cw, gap2 = 3.85, 0.32
    x0, y0, ch = 0.85, 3.65, 3.15
    for i, (title, c, items) in enumerate(cols):
        x = x0 + i * (cw + gap2)
        add_rect(s, x, y0, cw, ch, fill=NAVY_CARD, alpha=82, shadow=True)
        add_rect(s, x, y0, 0.1, ch, fill=c)
        add_text(s, x + 0.35, y0 + 0.3, cw - 0.6, 0.6,
                 [[(title, FONT, 18, WHITE, True)]])
        runs = [[("●  ", FONT, 11, c, True), (it, FONT, 12.5, MUTED, False)] for it in items]
        add_text(s, x + 0.35, y0 + 1.15, cw - 0.62, ch - 1.3, runs,
                 line_spacing=1.25, para_gap=10)
    page_no(s, 4, TOTAL)


# ============================================================================
# Slide 5 · 产业实力（数字）
# ============================================================================
def slide_strength():
    s = add_slide()
    full_bg(s, prep_bg("achievement.png", darken=0.8))
    overlay(s, 0, 0, SLIDE_W_IN, SLIDE_H_IN, NAVY_DK, 32)
    kicker(s, 0.85, 0.7, "04  产业实力 · BY THE NUMBERS")
    section_title(s, 0.82, 1.15, "我们已经做成了什么")
    stats = [
        ("近 200", "家", "累计孵化创新企业"),
        ("60 亿+", "元", "2025 企业总产值"),
        ("67", "家", "累计获投融资企业"),
        ("72", "家", "高新技术企业"),
        ("1", "家", "科创板上市企业"),
        ("15", "年", "深耕云计算产业"),
    ]
    cw, ch, gx, gy = 3.85, 1.95, 0.32, 0.3
    x0, y0 = 0.85, 2.15
    for idx, (num, unit, label) in enumerate(stats):
        r, col = divmod(idx, 3)
        x = x0 + col * (cw + gx)
        y = y0 + r * (ch + gy)
        add_rect(s, x, y, cw, ch, fill=NAVY_CARD, alpha=80, shadow=True)
        add_rect(s, x + 0.32, y + 0.4, 0.5, 0.06, fill=CYAN if idx % 2 == 0 else GOLD)
        add_text(s, x + 0.3, y + 0.55, cw - 0.5, 0.9,
                 [[(num, FONT, 40, WHITE, True), (" " + unit, FONT, 16, CYAN if idx % 2 == 0 else GOLD, True)]])
        add_text(s, x + 0.32, y + 1.42, cw - 0.5, 0.4,
                 [[(label, FONT, 13, MUTED, False)]])
    add_text(s, 0.85, SLIDE_H_IN - 0.95, 11.6, 0.5,
             [[("另含 国家级专精特新「小巨人」1 家 · 上海市专精特新 26 家 · 科技小巨人 3 家 · 创新型中小企业 42 家", FONT, 11.5, MUTED, False)]])
    page_no(s, 5, TOTAL)


# ============================================================================
# Slide 6 · 明星企业（分层）
# ============================================================================
def slide_stars():
    s = add_slide()
    full_bg(s, prep_bg("ecosystem.png", darken=0.66))
    overlay(s, 0, 0, SLIDE_W_IN, SLIDE_H_IN, NAVY_DK, 30)
    kicker(s, 0.85, 0.7, "05  明星企业 · PORTFOLIO")
    section_title(s, 0.82, 1.15, "层次分明的产业创新集群")
    tiers = [
        ("已上市", GOLD, "优刻得（科创板·中国云计算首批上市企业之一）"),
        ("Pre-IPO", CYAN, "上海道客（云原生操作系统·国家级专精特新「小巨人」）"),
        ("高成长", CYAN, "骞云科技 B+ · 贝塔数据 B+ · 速石科技 A · 炎凰数据 A+ · 骥步科技"),
        ("细分龙头", GOLD, "即科智能（营收 5 亿+）· 秒针系统（拟港股上市）· 伟仕佳杰（全球第八大 IT 分销）"),
        ("早期高潜", CYAN, "时加跳动 · 剀蕊得 · 康进讯达 · 云鹊医疗"),
    ]
    y = 2.05
    rh = 0.92
    for i, (tier, c, names) in enumerate(tiers):
        add_rect(s, 0.85, y, 11.63, rh - 0.16, fill=NAVY_CARD, alpha=78)
        add_rect(s, 0.85, y, 0.09, rh - 0.16, fill=c)
        add_rect(s, 1.1, y + 0.16, 2.0, rh - 0.48, fill=c, alpha=16)
        add_text(s, 1.1, y, 2.0, rh - 0.16,
                 [[(tier, FONT, 17, c, True)]],
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, 3.35, y, 9.0, rh - 0.16,
                 [[(names, FONT, 13.5, WHITE, False)]],
                 anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)
        y += rh
    page_no(s, 6, TOTAL)


# ============================================================================
# Slide 7 · 产业方向（四大赛道）
# ============================================================================
def slide_tracks():
    s = add_slide()
    full_bg(s, prep_bg("tracks.png", darken=0.72))
    overlay(s, 0, 0, SLIDE_W_IN, SLIDE_H_IN, NAVY_DK, 30)
    kicker(s, 0.85, 0.7, "06  产业方向 · FOCUS TRACKS")
    section_title(s, 0.82, 1.15, "聚焦四大前沿赛道")
    tracks = [
        ("云计算 / 云原生", "IaaS · PaaS · 多云管理 · 云原生中间件与存储", CYAN),
        ("大数据 / 人工智能", "AI 数据治理 · 智能算法 · 行业大模型应用", GOLD),
        ("边缘计算", "云边协同 · 实时数据 · 智能终端融合", CYAN),
        ("SaaS / 智能终端", "垂直行业 SaaS · FPGA 可编程芯片 · 平台型创新", GOLD),
    ]
    cw, gap = 2.86, 0.28
    x0, y0, ch = 0.85, 2.3, 3.9
    for i, (title, desc, c) in enumerate(tracks):
        x = x0 + i * (cw + gap)
        add_rect(s, x, y0, cw, ch, fill=NAVY_CARD, alpha=84, shadow=True)
        add_rect(s, x, y0, cw, 0.1, fill=c)
        add_text(s, x + 0.32, y0 + 0.5, cw - 0.6, 0.7,
                 [[(f"0{i+1}", FONT, 28, c, True)]])
        add_text(s, x + 0.32, y0 + 1.35, cw - 0.6, 1.0,
                 [[(title, FONT, 18, WHITE, True)]], line_spacing=1.1)
        add_rect(s, x + 0.32, y0 + 2.35, 0.7, 0.035, fill=MUTED, alpha=60)
        add_text(s, x + 0.32, y0 + 2.6, cw - 0.62, 1.2,
                 [[(desc, FONT, 12.5, MUTED, False)]], line_spacing=1.3)
    page_no(s, 7, TOTAL)


# ============================================================================
# Slide 8 · 你将获得（服务，客户视角）
# ============================================================================
def slide_services():
    s = add_slide()
    full_bg(s, prep_bg("services.png", darken=0.7))
    overlay(s, 0, 0, SLIDE_W_IN, SLIDE_H_IN, NAVY_DK, 40)
    gradient_left(s, NAVY_DK, width=7.6, alpha=70)
    kicker(s, 0.85, 0.7, "07  你将获得 · WHAT YOU GAIN")
    section_title(s, 0.82, 1.15, "入驻云基地，你将获得")
    items = [
        ("投融资对接", "云海创投 + 云天使基金，链接红杉、高瓴、腾讯、百度等一线机构"),
        ("场景验证", "方案孵化器提供需求调研、资源整合与实验测试，加速产品落地"),
        ("客户与市场资源", "端到端产业生态，打通金融、营销、医疗、轨交等行业客户"),
        ("政策与战略支持", "国家级孵化器资质背书，提供顶层设计与战略（投资）咨询"),
        ("国际化拓展", "依托五角场创院、港中大上海中心，助力「走出去·引进来」"),
        ("人才与培训", "五角场创新创业学院，体系化人才培养与产学研合作"),
    ]
    cw, ch, gx, gy = 5.7, 1.35, 0.3, 0.25
    x0, y0 = 0.85, 2.2
    for idx, (title, desc) in enumerate(items):
        r, col = divmod(idx, 2)
        x = x0 + col * (cw + gx)
        y = y0 + r * (ch + gy)
        add_rect(s, x, y, cw, ch, fill=NAVY_CARD, alpha=76)
        add_rect(s, x, y, 0.09, ch, fill=CYAN if idx % 2 == 0 else GOLD)
        add_text(s, x + 0.35, y + 0.2, cw - 0.6, 0.5,
                 [[("✓ ", FONT, 15, GOLD, True), (title, FONT, 16, WHITE, True)]])
        add_text(s, x + 0.62, y + 0.68, cw - 0.9, 0.6,
                 [[(desc, FONT, 12, MUTED, False)]], line_spacing=1.2)
    page_no(s, 8, TOTAL)


# ============================================================================
# Slide 9 · 合作方式（转化页）
# ============================================================================
def slide_cooperation():
    s = add_slide()
    full_bg(s, prep_bg("cooperation.png", darken=0.78))
    overlay(s, 0, 0, SLIDE_W_IN, SLIDE_H_IN, NAVY_DK, 34)
    gradient_left(s, NAVY_DK, width=8.0, alpha=78)
    kicker(s, 0.85, 0.7, "08  合作方式 · LET'S PARTNER")
    section_title(s, 0.82, 1.15, "我们诚挚欢迎")
    cats = [
        ("AI / 硬科技企业", "入驻孵化"),
        ("SaaS 企业", "加速成长"),
        ("产业链合作伙伴", "生态共建"),
        ("投资机构 / 资本", "项目共投"),
        ("政府 / 载体", "平台共建"),
    ]
    cw, gap = 2.22, 0.18
    x0, y0, ch = 0.85, 2.4, 1.95
    for i, (a, b) in enumerate(cats):
        x = x0 + i * (cw + gap)
        c = CYAN if i % 2 == 0 else GOLD
        add_rect(s, x, y0, cw, ch, fill=NAVY_CARD, alpha=88, line_color=c, line_w=1.2, shadow=True)
        add_text(s, x + 0.15, y0 + 0.45, cw - 0.3, 0.9,
                 [[(a, FONT, 15.5, WHITE, True)]],
                 align=PP_ALIGN.CENTER, line_spacing=1.1)
        add_rect(s, x + cw/2 - 0.35, y0 + 1.28, 0.7, 0.03, fill=c, alpha=70)
        add_text(s, x + 0.1, y0 + 1.4, cw - 0.2, 0.45,
                 [[(b, FONT, 12.5, c, True)]], align=PP_ALIGN.CENTER)
    # CTA
    add_rect(s, 0.85, 5.15, 11.63, 1.25, fill=CYAN_DK, alpha=22, line_color=CYAN, line_w=1.2)
    add_text(s, 1.25, 5.4, 8.5, 0.8, [
        [("把上下楼变上下游 —— 让我们成为你产业化的第一站", FONT, 19, WHITE, True)],
        [("欢迎洽谈入驻、投资与生态合作", FONT, 13, MUTED, False)],
    ], line_spacing=1.2, para_gap=6)
    add_text(s, 9.6, 5.4, 2.7, 0.8,
             [[("立即联系", FONT, 18, GOLD, True)],
              [("大学路 322 号 205 室", FONT, 12, MUTED, False)]],
             align=PP_ALIGN.RIGHT, line_spacing=1.2, para_gap=6)
    page_no(s, 9, TOTAL)


# ============================================================================
# Slide 10 · 发展历程（时间轴）
# ============================================================================
def slide_history():
    s = add_slide()
    full_bg(s, prep_bg("texture.png", darken=0.74))
    kicker(s, 0.85, 0.7, "09  发展历程 · MILESTONES")
    section_title(s, 0.82, 1.15, "十五年，从无到有构建产业生态")
    milestones = [
        ("2010.10", "正式成立", "上海「云海计划」杨浦核心执行主体"),
        ("2014", "市级孵化器", "认定为上海市科技企业孵化器"),
        ("2016", "大数据基地", "获授牌「上海市大数据创新基地」"),
        ("2019", "国家级跃升", "晋升国家级科技企业孵化器（专业类）"),
        ("2025", "迈向卓越级", "拟推荐工信部卓越级·总产值超 60 亿"),
    ]
    # 横向时间轴
    axis_y = 3.4
    add_rect(s, 1.0, axis_y, 11.3, 0.04, fill=CYAN, alpha=70)
    n = len(milestones)
    seg = 11.3 / n
    for i, (yr, title, desc) in enumerate(milestones):
        cx = 1.0 + seg * i + seg / 2
        c = GOLD if i % 2 == 0 else CYAN
        # 节点
        add_rect(s, cx - 0.1, axis_y - 0.07, 0.2, 0.2, fill=c, shape=MSO_SHAPE.OVAL)
        above = i % 2 == 0
        if above:
            add_text(s, cx - seg/2, axis_y - 1.55, seg, 1.3, [
                [(yr, FONT, 20, c, True)],
                [(title, FONT, 15, WHITE, True)],
                [(desc, FONT, 11, MUTED, False)],
            ], align=PP_ALIGN.CENTER, line_spacing=1.1, para_gap=4)
            add_rect(s, cx - 0.01, axis_y - 0.25, 0.02, 0.25, fill=c, alpha=60)
        else:
            add_text(s, cx - seg/2, axis_y + 0.45, seg, 1.3, [
                [(yr, FONT, 20, c, True)],
                [(title, FONT, 15, WHITE, True)],
                [(desc, FONT, 11, MUTED, False)],
            ], align=PP_ALIGN.CENTER, line_spacing=1.1, para_gap=4)
            add_rect(s, cx - 0.01, axis_y + 0.12, 0.02, 0.25, fill=c, alpha=60)
    page_no(s, 10, TOTAL)


# ============================================================================
# Slide 11 · 封底
# ============================================================================
def slide_back():
    s = add_slide()
    full_bg(s, prep_bg("backcover.png", darken=0.6))
    overlay(s, 0, 0, SLIDE_W_IN, SLIDE_H_IN, NAVY_DK, 46)
    add_text(s, 0, 2.0, SLIDE_W_IN, 1.6, [
        [("上海市云计算创新基地", FONT, 40, WHITE, True)],
        [("让技术企业从 0 → 1 → 上市", FONT, 18, CYAN, True)],
    ], align=PP_ALIGN.CENTER, line_spacing=1.2, para_gap=10)
    add_rect(s, SLIDE_W_IN/2 - 0.6, 3.95, 1.2, 0.05, fill=GOLD)
    add_text(s, 0, 4.4, SLIDE_W_IN, 1.6, [
        [("运营主体  上海市杨浦云计算创新基地发展有限公司", FONT, 14, WHITE, False)],
        [("地　　址  上海市杨浦区大学路 322 号 205 室", FONT, 14, MUTED, False)],
    ], align=PP_ALIGN.CENTER, line_spacing=1.4, para_gap=6)
    add_text(s, 0, SLIDE_H_IN - 0.75, SLIDE_W_IN, 0.4,
             [[("SHANGHAI CLOUD BASE  ·  把上下楼变上下游", FONT, 11, MUTED, False, False, 2)]],
             align=PP_ALIGN.CENTER)


# ----------------------------------------------------------------------------
def main():
    slide_cover()
    slide_value()
    slide_why()
    slide_trinity()
    slide_strength()
    slide_stars()
    slide_tracks()
    slide_services()
    slide_cooperation()
    slide_history()
    slide_back()
    out = os.path.join(DIST, "上海云基地-招商合作版宣传册.pptx")
    prs.save(out)
    print("Saved:", out, "| slides:", len(prs.slides._sldIdLst))


if __name__ == "__main__":
    main()
