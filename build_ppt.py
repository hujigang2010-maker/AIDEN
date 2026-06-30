# -*- coding: utf-8 -*-
"""
易居研究院 × 复旦大学住房政策研究中心 · 数据合作方案
生成 16:9 PPT。
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------- 主题 ----------
NAVY   = RGBColor(0x13, 0x29, 0x4B)   # 主色 深蓝
NAVY2  = RGBColor(0x1F, 0x3C, 0x66)
CRIMSON= RGBColor(0xB2, 0x22, 0x34)   # 复旦红 强调
GOLD   = RGBColor(0xC9, 0xA2, 0x27)   # 金 点缀
TEAL   = RGBColor(0x2E, 0x8B, 0x8B)
LIGHT  = RGBColor(0xF4, 0xF6, 0xF9)
CARD   = RGBColor(0xFF, 0xFF, 0xFF)
GRAYBG = RGBColor(0xEC, 0xEF, 0xF3)
INK    = RGBColor(0x22, 0x2A, 0x35)
GRAY   = RGBColor(0x5B, 0x66, 0x73)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
GREEN  = RGBColor(0x2E, 0x7D, 0x46)

FONT = "WenQuanYi Micro Hei"

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


def rect(s, x, y, w, h, fill, line=None, line_w=None, shadow=False, round_=False):
    shp = s.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if round_ else MSO_SHAPE.RECTANGLE,
        x, y, w, h)
    if round_:
        try:
            shp.adjustments[0] = 0.06
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
        shp.line.width = line_w or Pt(1)
    shp.shadow.inherit = False
    if shadow:
        el = shp._element.spPr
        ef = el.makeelement(qn('a:effectLst'), {})
        sh = el.makeelement(qn('a:outerShdw'),
                            {'blurRad': '90000', 'dist': '38100',
                             'dir': '5400000', 'rotWithShape': '0'})
        clr = el.makeelement(qn('a:srgbClr'), {'val': '6B7785'})
        alpha = el.makeelement(qn('a:alpha'), {'val': '38000'})
        clr.append(alpha); sh.append(clr); ef.append(sh); el.append(ef)
    return shp


def txt(s, x, y, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        wrap=True):
    """lines: list of (text, size, color, bold, [space_after]) ; or list of segments per para"""
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0; tf.margin_right = 0
    tf.margin_top = 0; tf.margin_bottom = 0
    first = True
    for ln in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        # ln can be a tuple (single run) or list of runs (segments)
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
        p.line_spacing = 1.0
    return tb


def header(s, idx, kicker, title, subtitle=None):
    rect(s, 0, 0, SW, SH, LIGHT)
    rect(s, 0, 0, Inches(0.16), SH, CRIMSON)
    # top kicker bar
    txt(s, Inches(0.6), Inches(0.42), Inches(10), Inches(0.35),
        [( "易居研究院  ×  复旦大学住房政策研究中心", 11, GRAY, False)])
    txt(s, Inches(0.6), Inches(0.72), Inches(11.6), Inches(0.7),
        [[(kicker + "  ", 14, CRIMSON, True), (title, 26, NAVY, True)]])
    if subtitle:
        txt(s, Inches(0.62), Inches(1.42), Inches(11.6), Inches(0.4),
            [(subtitle, 12.5, GRAY, False)])
    rect(s, Inches(0.62), Inches(1.86), Inches(12.1), Pt(2), GOLD)
    # page num
    txt(s, Inches(12.4), Inches(7.02), Inches(0.7), Inches(0.3),
        [(f"{idx:02d}", 11, GRAY, True)], align=PP_ALIGN.RIGHT)


# ============================================================
# 1. 封面
# ============================================================
def cover():
    s = slide()
    rect(s, 0, 0, SW, SH, NAVY)
    # decorative bands
    rect(s, 0, 0, SW, Inches(0.18), CRIMSON)
    rect(s, 0, Inches(5.55), SW, Inches(0.06), GOLD)
    # subtle blocks
    rect(s, Inches(9.4), Inches(1.0), Inches(3.4), Inches(3.4), NAVY2)
    rect(s, Inches(10.0), Inches(1.6), Inches(2.2), Inches(2.2), CRIMSON)

    txt(s, Inches(0.9), Inches(1.15), Inches(9), Inches(0.5),
        [("数 据 合 作 方 案", 16, GOLD, True)])
    txt(s, Inches(0.9), Inches(1.75), Inches(11), Inches(1.8),
        [[("易居研究院", 40, WHITE, True)],
         [("  ×  ", 30, GOLD, True),
          ("复旦大学住房政策研究中心", 34, WHITE, True)]])
    txt(s, Inches(0.92), Inches(3.5), Inches(11.5), Inches(0.6),
        [("《上海商办楼宇和产业园区市场报告》· 试点合作（杨浦 PoC）", 18, RGBColor(0xCF,0xD8,0xE6), False)])
    txt(s, Inches(0.92), Inches(4.15), Inches(11.5), Inches(0.5),
        [("全域全量数据 + AI 热力活跃度指数（替代出租率）· 双方联合发布、共同署名", 13,
          RGBColor(0xAE,0xBA,0xCC), False)])

    # info chips
    info = [("项目名称", "杨浦区办公楼宇数据采集试点"),
            ("报价日期", "2026 年 6 月 24 日"),
            ("有效期", "30 个自然日")]
    cx = Inches(0.9)
    for k, v in info:
        rect(s, cx, Inches(5.9), Inches(3.7), Inches(0.95), NAVY2, round_=True)
        txt(s, cx + Inches(0.25), Inches(6.05), Inches(3.3), Inches(0.7),
            [[(k, 11, GOLD, True)], [(v, 13.5, WHITE, True, 0)]])
        cx += Inches(3.95)


# ============================================================
# 2. 项目概况
# ============================================================
def overview():
    s = slide()
    header(s, 2, "一、", "项目概况", "Project Overview · 双方合作的范围、维度与产出")
    items = [
        ("试点范围", "杨浦区全域（五角场、杨浦滨江为重点）"),
        ("数据维度", "贵方 16 字段需求 + 空置率 + 衍生分析指标"),
        ("交付周期", "首期 2-3 周 · 季度更新机制"),
        ("数据规模", "约 600 栋办公楼宇 + 5,000+ 入驻企业"),
        ("合作模式", "数据 + 分析包，可按范围弹性升级"),
        ("联合发布", "研究院与我方共同署名、共同对外发布"),
    ]
    x0, y0 = Inches(0.62), Inches(2.25)
    cw, ch = Inches(3.86), Inches(2.15)
    gx, gy = Inches(0.26), Inches(0.32)
    for i, (k, v) in enumerate(items):
        r, c = divmod(i, 3)
        x = x0 + c * (cw + gx)
        y = y0 + r * (ch + gy)
        rect(s, x, y, cw, ch, CARD, shadow=True, round_=True)
        rect(s, x, y, Inches(0.12), ch, NAVY, round_=False)
        txt(s, x + Inches(0.38), y + Inches(0.32), cw - Inches(0.6), Inches(0.5),
            [(k, 16, CRIMSON, True)])
        rect(s, x + Inches(0.4), y + Inches(0.86), Inches(0.7), Pt(2), GOLD)
        txt(s, x + Inches(0.38), y + Inches(1.05), cw - Inches(0.66), Inches(1.0),
            [(v, 13.5, INK, False)], anchor=MSO_ANCHOR.TOP)


# ============================================================
# 3. 合作模式与分工
# ============================================================
def partnership():
    s = slide()
    header(s, 4, "三、", "合作双方与分工",
           "强强联合：AI 数据能力 + 学术与政策影响力，共同打造行业标杆")
    # two partner cards + middle plus
    cards = [
        (NAVY, "易居研究院", "AI 数据采集与分析方",
         ["AI 自动化多源数据采集（高德 / 企查查 / 公开网络）",
          "空置率自有模型测算、行业分布与板块迁徙分析",
          "可视化报告、数据系统与招商线索产出",
          "16 区一键复制的方法论与采集脚本资产"]),
        (CRIMSON, "复旦大学住房政策研究中心", "学术背书与联合发布方",
         ["权威学术背书与住房 / 产业政策研究框架",
          "数据口径与研究方法学审核把关",
          "联合署名、共同对外发布行业报告",
          "对接决策层、提升成果政策影响力"]),
    ]
    cw, ch = Inches(5.55), Inches(4.0)
    xs = [Inches(0.62), Inches(7.16)]
    for (clr, name, role, pts), x in zip(cards, xs):
        rect(s, x, Inches(2.3), cw, ch, CARD, shadow=True, round_=True)
        rect(s, x, Inches(2.3), cw, Inches(0.95), clr, round_=True)
        rect(s, x, Inches(2.72), cw, Inches(0.53), clr)
        txt(s, x + Inches(0.35), Inches(2.42), cw - Inches(0.6), Inches(0.8),
            [[(name, 18, WHITE, True)], [(role, 12, RGBColor(0xE6,0xEC,0xF4), True, 0)]])
        ty = Inches(3.5)
        for p in pts:
            txt(s, x + Inches(0.45), ty, cw - Inches(0.8), Inches(0.7),
                [[("●  ", 11, GOLD, True), (p, 13, INK, False)]])
            ty += Inches(0.74)
    # center plus badge
    rect(s, Inches(6.27), Inches(4.0), Inches(0.78), Inches(0.78), GOLD, round_=True)
    txt(s, Inches(6.27), Inches(4.05), Inches(0.78), Inches(0.68),
        [("+", 30, WHITE, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ============================================================
# 4. 数据方案内容
# ============================================================
def data_package():
    s = slide()
    header(s, 6, "五、", "数据 + 分析合作内容",
           "一体化的数据采集、分析与联合发布产出（空置率由我方 AI 自有模型测算）")
    cols = [
        ("基础数据层", NAVY,
         ["楼宇名称、所属板块、物业类型", "物业等级、总建筑面积、竣工时间",
          "入驻企业、所属行业、经营状态", "结构化 Excel 数据表"]),
        ("分析模型层", TEAL,
         ["报价 / 成交租金、物业费", "租赁面积、企业画像",
          "楼宇活跃度（热力）指数", "行业分布图谱 + 板块迁徙分析"]),
        ("成果交付层", CRIMSON,
         ["Excel 数据 + 40-60 页分析报告", "可视化图表与专题图",
          "联合署名行业专报", "季度更新机制"]),
    ]
    cw = Inches(3.92)
    xs = [Inches(0.62), Inches(4.7), Inches(8.78)]
    for (title, clr, pts), x in zip(cols, xs):
        rect(s, x, Inches(2.35), cw, Inches(4.0), CARD, shadow=True, round_=True)
        rect(s, x, Inches(2.35), cw, Inches(0.78), clr, round_=True)
        rect(s, x, Inches(2.72), cw, Inches(0.41), clr)
        txt(s, x, Inches(2.46), cw, Inches(0.6),
            [(title, 16, WHITE, True)], align=PP_ALIGN.CENTER)
        ty = Inches(3.4)
        for p in pts:
            txt(s, x + Inches(0.34), ty, cw - Inches(0.6), Inches(0.8),
                [[("✓ ", 13, GREEN, True), (p, 13, INK, False)]])
            ty += Inches(0.66)
    txt(s, Inches(0.62), Inches(6.55), Inches(12.2), Inches(0.5),
        [[("说明：", 12, CRIMSON, True),
          ("出租率属闭源数据、不作核心指标，改以「楼宇活跃度（热力）指数」替代；空置率以企业入驻密度近似，均标注口径，不含第三方采购付费。", 12, GRAY, False)]])


# ============================================================
# 5. 16字段覆盖与AI可达性
# ============================================================
def fields_table():
    s = slide()
    header(s, 7, "六、", "16 字段覆盖 · AI 可达性对照",
           "逐项对照贵方原始需求；标注需实地踏勘（“爬楼”）的字段")
    rows = [
        ("楼宇名称", "易", "100%", "AI 自动"),
        ("所属板块", "易", "100%", "AI 自动"),
        ("物业类型", "易", "95%", "AI 自动"),
        ("物业等级", "中", "80%", "AI + 推断"),
        ("总建筑面积", "中", "85%", "AI 自动"),
        ("竣工时间", "易", "95%", "AI 自动"),
        ("入驻企业名称", "易", "100%", "AI 自动"),
        ("入驻企业行业", "易", "100%", "AI 自动"),
        ("企业经营状态", "易", "100%", "AI 自动"),
        ("报价租金", "中", "85%", "AI + 抽样"),
        ("物业费", "中", "75%", "AI + 抽样"),
        ("标准层层高", "难", "60%", "需实地踏勘"),
        ("成交租金", "难", "60%", "需实地踏勘"),
        ("免租期", "难", "50%", "需实地踏勘"),
        ("租赁面积", "难", "50%", "需实地踏勘"),
        ("入驻楼层", "难", "40%", "需实地踏勘"),
    ]
    # two columns of 8 rows
    colspec = [Inches(2.2), Inches(0.9), Inches(1.05), Inches(1.65)]
    headers = ["字段", "难度", "AI可达", "采集方式"]
    def draw_block(x0, subset):
        y = Inches(2.3)
        # header row
        cx = x0
        rect(s, x0, y, sum(colspec, Emu(0)), Inches(0.46), NAVY)
        for w, htext in zip(colspec, headers):
            txt(s, cx, y, w, Inches(0.46), [(htext, 12.5, WHITE, True)],
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            cx += w
        y += Inches(0.46)
        for i, (f, d, a, m) in enumerate(subset):
            bg = CARD if i % 2 == 0 else GRAYBG
            rowh = Inches(0.455)
            rect(s, x0, y, sum(colspec, Emu(0)), rowh, bg)
            hard = m == "需实地踏勘"
            cx = x0
            vals = [f, d, a, m]
            for j, (w, val) in enumerate(zip(colspec, vals)):
                color = CRIMSON if (hard and j >= 1) else INK
                bold = hard and j == 3
                al = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
                pad = Inches(0.18) if j == 0 else 0
                txt(s, cx + pad, y, w - pad, rowh,
                    [(val, 11.5, color, bold)], align=al, anchor=MSO_ANCHOR.MIDDLE)
                cx += w
            y += rowh
        return y
    draw_block(Inches(0.62), rows[:8])
    draw_block(Inches(6.75), rows[8:])
    txt(s, Inches(0.62), Inches(6.82), Inches(12), Inches(0.5),
        [[("● 字段整体覆盖率约 100%；", 12, NAVY, True),
          ("标红 4 项为需“爬楼”实地踏勘字段，AI 仅能抽样估算 —— 详见第 08 页免责说明。", 12, GRAY, False)]])


# ============================================================
# 6. 免责声明与责任共担
# ============================================================
def disclaimer():
    s = slide()
    header(s, 8, "七、", "免责声明与责任共担",
           "明确数据边界与双方权责，确保合作长期、稳健、可持续")
    # left: disclaimer
    x = Inches(0.62); w = Inches(6.0)
    rect(s, x, Inches(2.3), w, Inches(4.55), CARD, shadow=True, round_=True)
    rect(s, x, Inches(2.3), w, Inches(0.85), CRIMSON, round_=True)
    rect(s, x, Inches(2.72), w, Inches(0.43), CRIMSON)
    txt(s, x + Inches(0.35), Inches(2.45), w - Inches(0.6), Inches(0.6),
        [("⚠  数据免责声明", 17, WHITE, True)])
    pts = [
        "对于 AI 无法通过自动化手段、需“爬楼”实地踏勘方可获取的字段"
        "（如入驻楼层、租赁面积、成交租金、免租期、标准层层高等），",
        "我方提供尽力采集与多源抽样估算，但不对该类数据的完整性与"
        "绝对准确性承担全责保障。",
        "上述字段以“抽样 + 估算 + 标注口径”方式交付，并明确标识置信区间，"
        "供研究与决策参考。",
        "可 AI 自动采集字段，我方按约定精度（误差 ±5%）负责。",
    ]
    ty = Inches(3.4)
    for p in pts:
        txt(s, x + Inches(0.4), ty, w - Inches(0.75), Inches(1.0),
            [[("•  ", 13, CRIMSON, True), (p, 12.5, INK, False)]])
        ty += Inches(0.86)
    # right: shared responsibility
    x2 = Inches(6.92); w2 = Inches(5.8)
    rect(s, x2, Inches(2.3), w2, Inches(4.55), CARD, shadow=True, round_=True)
    rect(s, x2, Inches(2.3), w2, Inches(0.85), NAVY, round_=True)
    rect(s, x2, Inches(2.72), w2, Inches(0.43), NAVY)
    txt(s, x2 + Inches(0.35), Inches(2.45), w2 - Inches(0.6), Inches(0.6),
        [("⚖  责任共担条款", 17, WHITE, True)])
    pts2 = [
        "双方对共同对外发布的数据承担共同责任与履约义务，"
        "而非仅由复旦大学住房政策研究中心单方承担。",
        "联合署名报告中的数据口径、结论与对外解释，由双方共同确认后发布。",
        "如因数据引发对外争议，双方依约定分工共同应对、共同处置。",
        "我方对采集与建模流程负责，研究院对研究框架与发布合规共同负责。",
    ]
    ty = Inches(3.4)
    for p in pts2:
        txt(s, x2 + Inches(0.4), ty, w2 - Inches(0.75), Inches(1.0),
            [[("•  ", 13, NAVY, True), (p, 12.5, INK, False)]])
        ty += Inches(0.86)


# ============================================================
# 7. 报价方案 · 一页总览（四卡）
# ============================================================
def pricing_page():
    s = slide()
    header(s, 11, "十、", "报价方案 · 一页总览",
           "已删除原 ¥28,000 基础包；新增 Token 预付费与变动报价 —— 全部报价一页可见")
    ORANGE = RGBColor(0xDE, 0x82, 0x0A)
    VIOLET = RGBColor(0x5E, 0x4B, 0x8C)
    # (badge, name, sub, header_color, bg, price, note, feats, button, recommended)
    cards = [
        ("试点 PoC", "Token 预付费", "试点验证 · 预付费", NAVY, RGBColor(0xE9, 0xF1, 0xFB),
         "≥ ¥80,000", "先付后开票 · 可抵扣",
         ["最低不少于 ¥80,000", "杨浦≈1,000 栋样本试跑", "含热力活跃度指标测试",
          "Token 按量 · 先付后开票", "可抵扣正式合作费用"],
         "试点启动 · 预付费", False),
        ("方案 B", "数据+分析包", "数据 + 分析模型", ORANGE, RGBColor(0xFF, 0xF4, 0xD6),
         "¥78,000", "维护：¥18,000/季",
         ["报价/成交租金、物业费", "租赁面积、企业画像", "楼宇活跃度（热力）指数",
          "行业分布 + 板块迁徙", "Excel + 40-60 页报告"],
         "★ 研究院首选 · 标杆", True),
        ("方案 C", "长期合作包", "系统 + 服务分成", GREEN, RGBColor(0xE7, 0xF4, 0xEA),
         "¥168,000", "维护：¥28K/季 + 分成 10-15%",
         ["方案 B 全部内容", "AI 企业匹配推荐", "招商策略生成",
          "项目诊断 + 自动化季报", "数据系统 + 招商线索库"],
         "长期战略合作伙伴", False),
        ("第四方案", "变动报价", "规模/频次/成效浮动", VIOLET, RGBColor(0xED, 0xEA, 0xF6),
         "浮动计价", "随规模与成效调整",
         ["全市 16 区复制单价递减", "扩大范围按项加收", "热力更新 / 高颗粒度加收",
          "季度/月度更新订阅", "招商成效分成 10-15%"],
         "全市 · 长期弹性补充", False),
    ]
    x = Inches(0.55); cw = Inches(2.9); gap = Inches(0.2)
    top = Inches(2.12); ch = Inches(4.32)
    for (badge, name, sub, hclr, bg, price, note, feats, btn, rec) in cards:
        y = top
        rect(s, x, y, cw, ch, bg, shadow=True, round_=True)
        rect(s, x, y, cw, Inches(1.0), hclr, round_=True)
        rect(s, x, y + Inches(0.5), cw, Inches(0.5), hclr)
        if rec:
            rect(s, x + cw - Inches(1.15), y - Inches(0.28), Inches(1.15), Inches(0.42),
                 ORANGE, round_=True)
            txt(s, x + cw - Inches(1.15), y - Inches(0.27), Inches(1.15), Inches(0.4),
                [("★ 推荐", 11, WHITE, True)], align=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.MIDDLE)
        txt(s, x + Inches(0.25), y + Inches(0.13), cw - Inches(0.5), Inches(0.8),
            [[(badge, 11, RGBColor(0xFF, 0xFF, 0xFF), True)],
             [(name, 16, WHITE, True, 0)]])
        txt(s, x + Inches(0.25), y + Inches(1.06), cw - Inches(0.5), Inches(0.32),
            [(sub, 10.5, GRAY, False)])
        txt(s, x + Inches(0.25), y + Inches(1.42), cw - Inches(0.5), Inches(0.6),
            [(price, 22, hclr, True)])
        txt(s, x + Inches(0.27), y + Inches(2.02), cw - Inches(0.5), Inches(0.3),
            [(note, 10, GRAY, False)])
        ty = y + Inches(2.34)
        for f in feats:
            txt(s, x + Inches(0.27), ty, cw - Inches(0.5), Inches(0.32),
                [[("✓ ", 10.5, hclr, True), (f, 10.5, INK, False)]])
            ty += Inches(0.29)
        rect(s, x + Inches(0.22), y + ch - Inches(0.52), cw - Inches(0.44), Inches(0.42),
             hclr, round_=True)
        txt(s, x + Inches(0.22), y + ch - Inches(0.52), cw - Inches(0.44), Inches(0.42),
            [(btn, 10.5, WHITE, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        x += cw + gap
    rect(s, Inches(0.55), Inches(6.6), Inches(12.23), Inches(0.6), LIGHT,
         line=NAVY, line_w=Pt(1), round_=True)
    txt(s, Inches(0.8), Inches(6.6), Inches(11.8), Inches(0.6),
        [[("附加条款：", 11.5, CRIMSON, True),
          ("出租率闭源、不作核心指标，改以热力活跃度替代 · Token 据实开票 · 热力指数收费点见第 10 页 · "
           "数据支持另计/共担、发布审核前置。", 11.5, GRAY, False)]], anchor=MSO_ANCHOR.MIDDLE)


# ============================================================
# 10. 全市复制路径
# ============================================================
def rollout():
    s = slide()
    header(s, 12, "十一、", "试点 → 全市 16 区复制路径",
           "杨浦做透 → 中心城区 → 浦东 → 外围 · 规模效应让单区报价递减")
    stages = [
        ("阶段一", "杨浦区（试点）", "Month 1-2", "¥70K", "单区保底", NAVY),
        ("阶段二", "静安+徐汇+黄浦", "Month 3-4", "¥70K/区", "中心 3 区", NAVY2),
        ("阶段三", "长宁+虹口+普陀+闸北", "Month 5-6", "¥60K/区", "中心 4 区", TEAL),
        ("阶段四", "浦东（陆家嘴/张江/临港）", "Month 7-9", "¥80K/区", "单区最大", CRIMSON),
        ("阶段五", "外围 7 区", "Month 10-12", "¥40K/区", "量大价低", GOLD),
    ]
    cw = Inches(2.32); gap = Inches(0.13)
    x = Inches(0.62); y = Inches(2.35)
    for (st, area, mo, price, tag, clr) in stages:
        rect(s, x, y, cw, Inches(3.0), CARD, shadow=True, round_=True)
        rect(s, x, y, cw, Inches(0.62), clr, round_=True)
        rect(s, x, y + Inches(0.3), cw, Inches(0.32), clr)
        txt(s, x, y + Inches(0.1), cw, Inches(0.5),
            [(st, 15, WHITE, True)], align=PP_ALIGN.CENTER)
        txt(s, x + Inches(0.2), y + Inches(0.8), cw - Inches(0.4), Inches(0.95),
            [(area, 12.5, INK, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        txt(s, x, y + Inches(1.75), cw, Inches(0.35),
            [(mo, 11, GRAY, False)], align=PP_ALIGN.CENTER)
        txt(s, x, y + Inches(2.12), cw, Inches(0.5),
            [(price, 20, clr, True)], align=PP_ALIGN.CENTER)
        txt(s, x, y + Inches(2.62), cw, Inches(0.3),
            [(tag, 11, GRAY, False)], align=PP_ALIGN.CENTER)
        x += cw + gap
    # total
    rect(s, Inches(0.62), Inches(5.7), Inches(12.1), Inches(1.2), NAVY, round_=True)
    txt(s, Inches(1.0), Inches(5.82), Inches(7.5), Inches(1.0),
        [[("上海全 16 区 · 全年覆盖（累计）", 15, GOLD, True)],
         [("行业首份全域数据 · 双方联合发布权", 12.5, RGBColor(0xCF,0xD8,0xE6), False, 0)]],
        anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(8.3), Inches(5.82), Inches(4.2), Inches(1.0),
        [[("约 ¥1,048,000", 26, WHITE, True)], [("（含一年维护，按变动报价递减）", 11, RGBColor(0xCF,0xD8,0xE6), False, 0)]],
        align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


# ============================================================
# 11. 差异化价值
# ============================================================
def value():
    s = slide()
    header(s, 13, "十二、", "我方差异化价值",
           "8 维度对比传统数据服务，AI 驱动、可复制、可联合发布")
    items = [
        ("数据颗粒度", "楼宇级 / 企业级 / 楼层级，细 10-100 倍"),
        ("数据动态性", "月度更新可选 / 实时企业变更追踪"),
        ("租金口径", "楼宇均值 + 报价/成交双口径，可下钻"),
        ("定制能力", "完全可定制 / 1-2 周快速交付"),
        ("招商可用性", "自带企业匹配 + 招商策略，数据转线索"),
        ("AI 能力", "AI 自动化 + 多源交叉，16 区可快速复制"),
        ("合作模式", "数据 + 分析 + 长期合作 + 联合发布"),
        ("联合署名", "研究院主名 · 我方副署 · 行业首份"),
    ]
    x0, y0 = Inches(0.62), Inches(2.3)
    cw, ch = Inches(5.95), Inches(1.0)
    gx, gy = Inches(0.2), Inches(0.18)
    for i, (k, v) in enumerate(items):
        r, c = divmod(i, 2)
        x = x0 + c * (cw + gx); y = y0 + r * (ch + gy)
        rect(s, x, y, cw, ch, CARD, shadow=True, round_=True)
        rect(s, x, y, Inches(1.9), ch, NAVY, round_=True)
        rect(s, x + Inches(1.78), y, Inches(0.12), ch, NAVY)
        txt(s, x + Inches(0.1), y, Inches(1.7), ch,
            [(k, 14, WHITE, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        txt(s, x + Inches(2.1), y, cw - Inches(2.35), ch,
            [(v, 12.5, INK, False)], anchor=MSO_ANCHOR.MIDDLE)


# ============================================================
# 12. 落地时间表
# ============================================================
def timeline():
    s = slide()
    header(s, 14, "十三、", "落地时间表",
           "8 周内完成杨浦试点项目，节点清晰、交付可控")
    steps = [
        ("W0", "合同确认", "签订协议、首付款到账"),
        ("W1", "试点启动", "字段定稿、采集脚本部署"),
        ("W2-3", "首轮采集", "楼宇 + 企业基础数据"),
        ("W4", "高价值字段", "租金 / 物业费 / 免租期补全"),
        ("W5", "分析建模", "空置率模型 + 行业分布"),
        ("W6", "报告编制", "40-60 页分析报告"),
        ("W7", "验收复核", "数据交叉验证 ±5%"),
        ("W8", "正式交付", "成果 + 培训 + 收尾款"),
    ]
    # horizontal timeline line
    rect(s, Inches(0.9), Inches(3.95), Inches(11.5), Pt(3), GOLD)
    cw = Inches(1.46); x = Inches(0.62)
    for i, (w, t, d) in enumerate(steps):
        cx = x + i * Inches(1.52)
        up = i % 2 == 0
        cy = Inches(2.45) if up else Inches(4.45)
        clr = NAVY if up else CRIMSON
        rect(s, cx, cy, cw, Inches(1.35), CARD, shadow=True, round_=True)
        rect(s, cx, cy, cw, Inches(0.42), clr, round_=True)
        rect(s, cx, cy + Inches(0.2), cw, Inches(0.22), clr)
        txt(s, cx, cy + Inches(0.04), cw, Inches(0.36),
            [(w, 13, WHITE, True)], align=PP_ALIGN.CENTER)
        txt(s, cx + Inches(0.08), cy + Inches(0.5), cw - Inches(0.16), Inches(0.4),
            [(t, 12.5, INK, True)], align=PP_ALIGN.CENTER)
        txt(s, cx + Inches(0.06), cy + Inches(0.88), cw - Inches(0.12), Inches(0.45),
            [(d, 9.5, GRAY, False)], align=PP_ALIGN.CENTER)
        # dot on line
        dot = rect(s, cx + cw/2 - Inches(0.09), Inches(3.86), Inches(0.18), Inches(0.18),
                   GOLD)
        dot.adjustments  # noop


# ============================================================
# 13. 合作条款
# ============================================================
def terms():
    s = slide()
    header(s, 15, "十四、", "合作条款",
           "明确双方权责 · 确保合作长期稳定")
    items = [
        ("数据所有权", "原始数据双方共有；研究院享独家发布权"),
        ("付款方式", "50% 启动金 + 30% 中期款 + 20% 验收款"),
        ("数据合规", "仅采集公开数据，涉及 PII 均脱敏处理"),
        ("验收标准", "可 AI 采集字段完整度 ≥ 95%、准确度 ±5%"),
        ("知识产权", "AI 模型与采集方法论归我方，合作期内授权使用"),
        ("发布署名", "联合署名：研究院主名 · 我方副署"),
        ("责任共担", "共同发布数据由双方共同担责、共同履约"),
        ("免责边界", "需“爬楼”实地字段不承担全责，按抽样估算交付"),
    ]
    x0, y0 = Inches(0.62), Inches(2.3)
    cw, ch = Inches(5.95), Inches(1.0)
    gx, gy = Inches(0.2), Inches(0.18)
    for i, (k, v) in enumerate(items):
        r, c = divmod(i, 2)
        x = x0 + c * (cw + gx); y = y0 + r * (ch + gy)
        rect(s, x, y, cw, ch, CARD, shadow=True, round_=True)
        rect(s, x, y, Inches(0.12), ch, CRIMSON)
        txt(s, x + Inches(0.35), y + Inches(0.14), cw - Inches(0.6), Inches(0.4),
            [[("▎ ", 13, CRIMSON, True), (k, 14.5, NAVY, True)]])
        txt(s, x + Inches(0.4), y + Inches(0.55), cw - Inches(0.65), Inches(0.4),
            [(v, 12, INK, False)])


# ============================================================
# 14. 推荐报价
# ============================================================
def recommendation():
    s = slide()
    rect(s, 0, 0, SW, SH, NAVY)
    rect(s, 0, 0, SW, Inches(0.16), CRIMSON)
    txt(s, Inches(0.9), Inches(0.6), Inches(11), Inches(0.5),
        [[("十五、", 14, GOLD, True), ("推荐报价", 26, WHITE, True)]])
    txt(s, Inches(0.92), Inches(1.18), Inches(11.5), Inches(0.4),
        [("综合贵方需求、数据边界与长期合作潜力，我方郑重推荐以下报价", 13,
          RGBColor(0xCF,0xD8,0xE6), False)])
    rect(s, Inches(0.92), Inches(1.7), Inches(11.5), Pt(2), GOLD)

    # main recommended card
    rect(s, Inches(0.92), Inches(2.0), Inches(7.0), Inches(4.5), CARD, shadow=True, round_=True)
    rect(s, Inches(0.92), Inches(2.0), Inches(7.0), Inches(1.1), CRIMSON, round_=True)
    rect(s, Inches(0.92), Inches(2.55), Inches(7.0), Inches(0.55), CRIMSON)
    rect(s, Inches(6.7), Inches(2.18), Inches(1.05), Inches(0.36), GOLD, round_=True)
    txt(s, Inches(6.7), Inches(2.19), Inches(1.05), Inches(0.34),
        [("★ 首选", 11, WHITE, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(1.3), Inches(2.16), Inches(6), Inches(0.9),
        [[("方案 B · 数据 + 分析包", 20, WHITE, True)],
         [("¥78,000 · 数据+分析模型 · 研究院标杆首选", 13, RGBColor(0xF0,0xD9,0xDC), True, 0)]])
    rpts = [
        "先以 Token 预付费 ≥¥80,000 启动 PoC（先付后开票 · 可抵扣）",
        "正式合作以【方案 B 数据+分析包】¥78,000 为标杆首选",
        "热力活跃度指数替代闭源出租率，不含第三方采购付费",
        "需“爬楼”字段以抽样估算交付，权责清晰、风险共担",
        "数据支持另计 / 共担，免责与发布审核前置确认",
    ]
    ty = Inches(3.4)
    for p in rpts:
        txt(s, Inches(1.35), ty, Inches(6.3), Inches(0.55),
            [[("✓  ", 14, GREEN, True), (p, 13, INK, False)]])
        ty += Inches(0.6)

    # right: ladder
    x2 = Inches(8.2)
    rect(s, x2, Inches(2.0), Inches(4.22), Inches(4.5), NAVY2, shadow=True, round_=True)
    txt(s, x2 + Inches(0.35), Inches(2.2), Inches(3.6), Inches(0.5),
        [("报价阶梯一览", 15, GOLD, True)])
    rect(s, x2 + Inches(0.35), Inches(2.7), Inches(1.0), Pt(2), GOLD)
    ladder = [
        ("试点 PoC", "≥¥80,000", "先付后开票"),
        ("方案 B 数据+分析", "¥78,000", "★ 推荐"),
        ("方案 C 长期合作", "¥168,000", "系统+分成"),
        ("变动报价", "浮动", "全市/长期"),
    ]
    ty = Inches(2.9)
    for name, price, tag in ladder:
        hl = "推荐" in tag
        rect(s, x2 + Inches(0.3), ty, Inches(3.62), Inches(0.6),
             CRIMSON if hl else NAVY, round_=True)
        txt(s, x2 + Inches(0.5), ty, Inches(2.0), Inches(0.6),
            [[(name, 12, WHITE, True)], [(tag, 10, GOLD if hl else RGBColor(0xAE,0xBA,0xCC), True, 0)]],
            anchor=MSO_ANCHOR.MIDDLE)
        txt(s, x2 + Inches(2.3), ty, Inches(1.5), Inches(0.6),
            [(price, 14, WHITE, True)], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        ty += Inches(0.7)
    txt(s, x2 + Inches(0.32), Inches(6.42), Inches(3.7), Inches(0.4),
        [("付款：50% 启动 + 30% 中期 + 20% 验收", 9.5, RGBColor(0xCF,0xD8,0xE6), False)])

    txt(s, Inches(0.92), Inches(6.7), Inches(11.5), Inches(0.5),
        [[("结论：", 13, GOLD, True),
          ("先以【Token 预付费 ≥¥80,000】启动 PoC 验证，正式合作以【方案 B 数据+分析包】为标杆，"
           "规模化升级【方案 C】或转【变动报价】。", 12.5, WHITE, False)]])


# ============================================================
# 15. 封底
# ============================================================
def closing():
    s = slide()
    rect(s, 0, 0, SW, SH, NAVY)
    rect(s, 0, Inches(2.6), SW, Inches(0.06), GOLD)
    rect(s, 0, Inches(4.7), SW, Inches(0.06), CRIMSON)
    txt(s, Inches(0), Inches(2.9), SW, Inches(0.8),
        [("期待与易居研究院的长期合作", 30, WHITE, True)], align=PP_ALIGN.CENTER)
    txt(s, Inches(0), Inches(3.8), SW, Inches(0.6),
        [("—— 双方联合发布行业首份全域报告，共建数据标杆 ——", 16, GOLD, True)],
        align=PP_ALIGN.CENTER)
    txt(s, Inches(0), Inches(5.0), SW, Inches(0.5),
        [("本方案为讨论稿，可根据双方实际需求灵活调整", 13, RGBColor(0xCF,0xD8,0xE6), False)],
        align=PP_ALIGN.CENTER)
    txt(s, Inches(0), Inches(5.55), SW, Inches(0.5),
        [("易居研究院  ×  复旦大学住房政策研究中心  ·  2026 年 6 月 24 日  ·  报价有效期 30 天", 12,
          RGBColor(0xAE,0xBA,0xCC), False)], align=PP_ALIGN.CENTER)


# ============================================================
# 项目边界 · 为什么需要收费
# ============================================================
def boundaries():
    s = slide()
    header(s, 3, "二、", "项目边界 · 为什么需要收费",
           "数据采集、建模与发布存在客观成本与风险，合理收费是项目可持续的前提")
    cards = [
        ("AI 识别空置率存在不确定性", CRIMSON,
         ["空置率依赖 AI 多源推断（POI / 招聘 / 入驻面积反推）",
          "客观存在 ±5% 误差，需持续建模、校准与抽样验证",
          "模型研发与维护构成真实、持续的投入"]),
        ("数据源不统一", NAVY,
         ["高德 / 企查查 / 中介 / 公开网络口径各不相同",
          "需大量清洗、对齐、交叉校验与去重",
          "“爬楼”字段更需人工实地补采与整合"]),
        ("涉及对外发布合规风险", TEAL,
         ["联合署名对外发布，需共同承担数据责任",
          "PII 脱敏、口径合规、对外解释均需把关",
          "合规审核与风险兜底产生实际成本"]),
    ]
    cw = Inches(3.92)
    xs = [Inches(0.62), Inches(4.7), Inches(8.78)]
    for (t, clr, pts), x in zip(cards, xs):
        rect(s, x, Inches(2.3), cw, Inches(3.55), CARD, shadow=True, round_=True)
        rect(s, x, Inches(2.3), cw, Inches(1.0), clr, round_=True)
        rect(s, x, Inches(2.75), cw, Inches(0.55), clr)
        txt(s, x + Inches(0.3), Inches(2.4), cw - Inches(0.55), Inches(0.8),
            [(t, 15, WHITE, True)], anchor=MSO_ANCHOR.MIDDLE)
        ty = Inches(3.55)
        for p in pts:
            txt(s, x + Inches(0.34), ty, cw - Inches(0.6), Inches(0.7),
                [[("• ", 12, clr, True), (p, 12.5, INK, False)]])
            ty += Inches(0.72)
    rect(s, Inches(0.62), Inches(6.15), Inches(12.1), Inches(0.95), NAVY, round_=True)
    rect(s, Inches(0.62), Inches(6.15), Inches(2.4), Inches(0.95), GOLD, round_=True)
    txt(s, Inches(0.62), Inches(6.15), Inches(2.4), Inches(0.95),
        [("收费合理化", 16, WHITE, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(3.25), Inches(6.15), Inches(9.2), Inches(0.95),
        [("正因建模不确定性、数据源整合与发布合规均需真实投入，项目以专业服务计费 —— "
          "收费对应的是“可信、可用、可对外发布”的数据成果。", 13, WHITE, False)],
        anchor=MSO_ANCHOR.MIDDLE)


# ============================================================
# 合作模式 · 服务三段结构
# ============================================================
def service_structure():
    s = slide()
    header(s, 5, "四、", "合作模式 · 服务三段结构",
           "把“服务”拆清楚：数据研究建模 → AI 分析输出 → 报告与对外发布")
    segs = [
        ("①", "数据研究与模型部分", NAVY,
         ["字段体系与口径定义", "空置率模型与方法学设计",
          "数据标准、误差边界与验证规则", "学术框架 + 研究方法共建"]),
        ("②", "AI 分析与输出部分", TEAL,
         ["多源数据自动化采集", "AI 交叉验证、清洗与对齐",
          "行业分布 / 板块迁徙分析", "可视化图表与数据系统"]),
        ("③", "报告与对外发布部分", CRIMSON,
         ["分析报告编制与排版", "联合署名、共同对外发布",
          "发布前合规审核流程", "对外口径与解释统一"]),
    ]
    cw = Inches(3.78)
    xs = [Inches(0.62), Inches(4.78), Inches(8.94)]
    for (no, t, clr, pts), x in zip(segs, xs):
        rect(s, x, Inches(2.45), cw, Inches(3.95), CARD, shadow=True, round_=True)
        rect(s, x, Inches(2.45), cw, Inches(1.05), clr, round_=True)
        rect(s, x, Inches(3.0), cw, Inches(0.5), clr)
        txt(s, x + Inches(0.28), Inches(2.5), Inches(0.85), Inches(0.95),
            [(no, 30, WHITE, True)], anchor=MSO_ANCHOR.MIDDLE)
        txt(s, x + Inches(1.1), Inches(2.5), cw - Inches(1.35), Inches(0.95),
            [(t, 15, WHITE, True)], anchor=MSO_ANCHOR.MIDDLE)
        ty = Inches(3.78)
        for p in pts:
            txt(s, x + Inches(0.35), ty, cw - Inches(0.6), Inches(0.6),
                [[("· ", 12, clr, True), (p, 12.5, INK, False)]])
            ty += Inches(0.62)
    for ax in [Inches(4.43), Inches(8.59)]:
        a = s.shapes.add_shape(MSO_SHAPE.CHEVRON, ax, Inches(4.05), Inches(0.34), Inches(0.62))
        a.fill.solid(); a.fill.fore_color.rgb = GOLD; a.line.fill.background()
        a.shadow.inherit = False
    txt(s, Inches(0.62), Inches(6.62), Inches(12), Inches(0.4),
        [[("说明：", 12, CRIMSON, True),
          ("三段服务既相互独立又彼此衔接，分别对应研究、技术与发布三类投入，便于分工计价与责任界定。",
           12, GRAY, False)]])


# ============================================================
# 报价定义 · 任务标准与范围
# ============================================================
def quote_definition():
    s = slide()
    header(s, 9, "八、", "报价定义 · 任务标准与范围",
           "借用《上海商办楼宇和产业园区市场报告》题目界定范围；融合最近一次 AI 评审会需求")
    # left: 任务标准（借用对方题目）
    x = Inches(0.62); w = Inches(6.0)
    rect(s, x, Inches(2.3), w, Inches(4.6), CARD, shadow=True, round_=True)
    rect(s, x, Inches(2.3), w, Inches(0.85), NAVY, round_=True)
    rect(s, x, Inches(2.72), w, Inches(0.43), NAVY)
    txt(s, x + Inches(0.32), Inches(2.43), w - Inches(0.6), Inches(0.6),
        [("▎任务标准（借用对方题目）", 15, WHITE, True)])
    left = [
        ("题目", "《上海商办楼宇和产业园区市场报告》"),
        ("覆盖", "超甲/甲/乙级及以下写字楼 + 各类产业园区"),
        ("颗粒度", "16 区 → 街道 / 镇二级（适配高德 POI/AOI）"),
        ("字段五类", "资产基础 / 企业画像 / 租赁交易 / 迁徙 / 衍生"),
        ("分期", "中心 7 区 → 浦东 → 全市全域"),
        ("报告框架", "概览 / 分区 / 画像 / 迁徙 / 供需匹配 / 趋势"),
    ]
    ty = Inches(3.35)
    for k, v in left:
        txt(s, x + Inches(0.34), ty, w - Inches(0.65), Inches(0.6),
            [[("▸ " + k + "：", 12, NAVY, True), (v, 11.5, INK, False)]])
        ty += Inches(0.57)
    # right: AI 评审会边界（写入报价定义）
    x2 = Inches(6.92); w2 = Inches(5.8)
    rect(s, x2, Inches(2.3), w2, Inches(4.6), CARD, shadow=True, round_=True)
    rect(s, x2, Inches(2.3), w2, Inches(0.85), CRIMSON, round_=True)
    rect(s, x2, Inches(2.72), w2, Inches(0.43), CRIMSON)
    txt(s, x2 + Inches(0.32), Inches(2.43), w2 - Inches(0.6), Inches(0.6),
        [("▎AI 评审会边界（写入报价定义）", 15, WHITE, True)])
    right = [
        "数据依赖网络“留痕”，静态 / 未公开信息无法纳入",
        "出租率属闭源数据 → 不作为核心指标",
        "改以“楼宇活跃度（热力）”替代出租率",
        "层高等基于行业常识估算，准确率约 80%",
        "成交租金缺时间戳、真实性存疑 → 标注口径",
        "Token 消耗按量计费、据实开具发票",
        "杨浦约 1,000 栋作为 PoC 验证样本",
    ]
    ty = Inches(3.35)
    for p in right:
        txt(s, x2 + Inches(0.34), ty, w2 - Inches(0.65), Inches(0.6),
            [[("•  ", 12, CRIMSON, True), (p, 11.5, INK, False)]])
        ty += Inches(0.49)


# ============================================================
# 热力活跃度指数 · 替代出租率与收费点
# ============================================================
def heat_index():
    s = slide()
    header(s, 10, "九、", "热力活跃度指数 · 替代出租率与收费点",
           "以百度/高德热力图构建“楼宇活跃度”，绕过闭源出租率；明确各项收费点")
    VIOLET = RGBColor(0x5E, 0x4B, 0x8C)
    # left definition
    x = Inches(0.62); w = Inches(4.5)
    rect(s, x, Inches(2.3), w, Inches(4.58), CARD, shadow=True, round_=True)
    rect(s, x, Inches(2.3), w, Inches(0.85), NAVY, round_=True)
    rect(s, x, Inches(2.72), w, Inches(0.43), NAVY)
    txt(s, x + Inches(0.3), Inches(2.43), w - Inches(0.55), Inches(0.6),
        [("▎指标定义", 15, WHITE, True)])
    defs = [
        ("数据源", "百度 / 高德热力图（定位请求密度）"),
        ("时段", "早 8:30-9:30 · 晚 18:00-19:00"),
        ("颗粒度", "10 米级，可区分单栋出入口"),
        ("计算", "活跃度 ÷ 总建筑面积 = 单位面积活跃度"),
        ("作用", "替代闭源出租率，辨办公 / 居住"),
    ]
    ty = Inches(3.4)
    for k, v in defs:
        txt(s, x + Inches(0.32), ty, w - Inches(0.6), Inches(0.7),
            [[("▎" + k, 12.5, CRIMSON, True)], [(v, 11, INK, False, 0)]])
        ty += Inches(0.68)
    # right charging points
    x2 = Inches(5.35); w2 = Inches(7.38)
    txt(s, x2, Inches(2.28), w2, Inches(0.5),
        [("★ 热力指数收费点", 16, CRIMSON, True)])
    items = [
        ("①", "热力数据采集费", "百度/高德热力 API 调用", "按 楼宇 × 时段 × 天数", NAVY),
        ("②", "Token / 算力费", "AI 运行按量计费、据实开票", "≈1,000栋约 ¥1,500/次", TEAL),
        ("③", "热力指标建模费", "时间窗口 + 提示词 + 面积归一化", "一次性建模费", CRIMSON),
        ("④", "热力更新订阅费", "月度 / 季度刷新与活跃度追踪", "按周期订阅", VIOLET),
        ("⑤", "高颗粒度加收", "单栋出入口级（10 米）热力", "按需加收", GOLD),
    ]
    y = Inches(2.82)
    rh = Inches(0.74); gap = Inches(0.12)
    for (no, t, d, price, c) in items:
        rect(s, x2, y, w2, rh, CARD, shadow=True, round_=True)
        rect(s, x2, y, Inches(0.68), rh, c, round_=True)
        txt(s, x2, y, Inches(0.68), rh, [(no, 22, WHITE, True)],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        txt(s, x2 + Inches(0.88), y + Inches(0.09), Inches(3.4), Inches(0.34),
            [(t, 13.5, NAVY, True)])
        txt(s, x2 + Inches(0.9), y + Inches(0.43), Inches(3.5), Inches(0.28),
            [(d, 10, GRAY, False)])
        rect(s, x2 + w2 - Inches(2.5), y + Inches(0.17), Inches(2.35), Inches(0.4),
             LIGHT, line=c, line_w=Pt(1.2), round_=True)
        txt(s, x2 + w2 - Inches(2.5), y + Inches(0.17), Inches(2.35), Inches(0.4),
            [(price, 10.5, c, True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        y += rh + gap


cover()
overview()
boundaries()
partnership()
service_structure()
data_package()
fields_table()
disclaimer()
quote_definition()
heat_index()
pricing_page()
rollout()
value()
timeline()
terms()
recommendation()
closing()

out = "易居研究院_复旦大学住房政策研究中心_合作方案.pptx"
prs.save(out)
print("saved", out, "slides=", len(prs.slides._sldIdLst))
