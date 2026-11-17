# -*- coding: utf-8 -*-
"""复旦大学住房政策研究中心 × 东方枢纽集团 战略合作计划（建议稿）PPT。"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

PLUM = RGBColor(0x2E, 0x1F, 0x47)
PURPLE = RGBColor(0x5B, 0x3E, 0x8E)
PURPLE2 = RGBColor(0x76, 0x58, 0xA8)
LILAC = RGBColor(0xC0, 0xAE, 0xE0)
LAV = RGBColor(0xF4, 0xEF, 0xFA)
LAVED = RGBColor(0xDD, 0xD2, 0xEF)
BGT = RGBColor(0xFB, 0xFA, 0xFE)
GOLD = RGBColor(0xB0, 0x86, 0x2C)
GREY = RGBColor(0x5E, 0x58, 0x68)
DARK = RGBColor(0x2B, 0x25, 0x36)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xE6, 0xDF, 0xF0)

FONT = "Microsoft YaHei"
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]
PAGE = [0]


def _font(run, size, bold, color):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        rPr.append(ea)
    ea.set("typeface", FONT)


def rect(s, x, y, w, h, fill=None, line=None, lw=1):
    sp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
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
        sp.line.width = Pt(lw)
    return sp


def rrect(s, x, y, w, h, fill=None, line=None, lw=1):
    sp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    sp.shadow.inherit = False
    # tighter corners
    try:
        sp.adjustments[0] = 0.08
    except Exception:
        pass
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(lw)
    return sp


def txt(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, after=3, ls=1.0):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Pt(2)
    tf.margin_right = Pt(2)
    tf.margin_top = Pt(1)
    tf.margin_bottom = Pt(1)
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(after)
        p.space_before = Pt(0)
        p.line_spacing = ls
        for item in para:
            r = p.add_run()
            r.text = item[0]
            _font(r, item[1], item[2], item[3])
    return tb


def new():
    PAGE[0] += 1
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, SW, SH, fill=BGT)
    return s


def header(s, kicker, title):
    rect(s, 0, 0, SW, Inches(1.05), fill=PLUM)
    rect(s, 0, Inches(1.05), SW, Pt(3), fill=GOLD)
    txt(s, Inches(0.55), Inches(0.12), Inches(11), Inches(0.28),
        [[(kicker, 11, True, LILAC)]], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(0.55), Inches(0.40), Inches(11.2), Inches(0.52),
        [[(title, 22, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(11.6), Inches(0.18), Inches(1.3), Inches(0.7),
        [[(f"{PAGE[0]:02d}", 18, True, PURPLE2)]], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def footer(s):
    txt(s, Inches(0.5), Inches(7.15), Inches(10.2), Inches(0.28),
        [[("复旦大学住房政策研究中心  ×  上海东方枢纽投资建设发展集团有限公司  |  战略合作计划（建议稿）", 8, False, GREY)]])
    txt(s, Inches(11.4), Inches(7.15), Inches(1.5), Inches(0.28),
        [[("内部呈送", 8, False, GREY)]], align=PP_ALIGN.RIGHT)


def card(s, x, y, w, h, fill=WHITE):
    return rrect(s, x, y, w, h, fill=fill, line=LINE, lw=1)


# ============================================================ 1 封面
s = prs.slides.add_slide(BLANK)
PAGE[0] = 1
rect(s, 0, 0, SW, SH, fill=PLUM)
rect(s, 0, 0, Inches(0.18), SH, fill=GOLD)
txt(s, Inches(0.7), Inches(0.55), Inches(11.5), Inches(0.35),
    [[("建议稿  ·  呈送东方枢纽集团  ·  2026年11月", 13, True, LILAC)]])
txt(s, Inches(0.7), Inches(1.35), Inches(12), Inches(1.7),
    [[("复旦大学住房政策研究中心", 28, True, WHITE)],
     [("与东方枢纽集团战略合作计划", 32, True, WHITE)]], after=6)
rect(s, Inches(0.72), Inches(3.25), Inches(2.2), Pt(3), fill=GOLD)
txt(s, Inches(0.7), Inches(3.5), Inches(11.8), Inches(1.1),
    [[("共同推进国际商务交流、产业招商、城市更新与科技创新合作", 16, False, LILAC)],
     [("首期：2026年第四季度 — 2027年12月", 15, True, WHITE)]], after=8)
txt(s, Inches(0.7), Inches(5.15), Inches(11.8), Inches(1.3),
    [[("学术及研究牵头：复旦大学住房政策研究中心", 13, False, WHITE)],
     [("产业及空间合作：上海东方枢纽投资建设发展集团有限公司", 13, False, WHITE)],
     [("拟参与实施：上海汇在东方建设发展有限公司、复旦大学科技园等相关单位", 13, False, WHITE)]], after=4)
txt(s, Inches(0.7), Inches(6.85), Inches(11.5), Inches(0.35),
    [[("本方案为合作建议稿，不代表各单位已完成内部审批或形成正式合作关系。", 10, False, LILAC)]])

# ============================================================ 2 升级定位
s = new(); header(s, "定位  ·  WHY NOW", "从联合办活动，升级为长期产业服务伙伴"); footer(s)
txt(s, Inches(0.55), Inches(1.25), Inches(12.2), Inches(0.55),
    [[("合作核心不再只是邀请嘉宾、举办论坛，而是形成一套可签约、可交付、可延续的商业与产业服务体系。", 14, True, PLUM)]])
chain = ["学术研究\n与智库咨询", "企业资源\n筛选", "精准招商\n对接", "国际商务\n活动", "项目落地与\n持续服务"]
for i, name in enumerate(chain):
    x = Inches(0.5) + i * Inches(2.55)
    rrect(s, x, Inches(2.05), Inches(2.25), Inches(1.35), fill=PURPLE if i % 2 == 0 else PLUM)
    txt(s, x, Inches(2.05), Inches(2.25), Inches(1.35),
        [[(f"0{i+1}", 11, True, GOLD)], [(name, 14, True, WHITE)]],
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, after=2)
    if i < 4:
        txt(s, x + Inches(2.15), Inches(2.4), Inches(0.45), Inches(0.6),
            [[("→", 18, True, GOLD)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
# three points
pts = [
    ("不是活动执行商", "单独两场活动的执行收入，不是这次合作最值得争取的结果。"),
    ("是长期服务伙伴", "成为东方枢纽在产业研究、企业导入和商务合作上的持续伙伴。"),
    ("先试点、再年度化", "先用两场活动打开合作，再用年度研究与精准招商建立持续关系。"),
]
for i, (t, d) in enumerate(pts):
    x = Inches(0.5) + i * Inches(4.2)
    card(s, x, Inches(3.75), Inches(4.0), Inches(2.0))
    rect(s, x, Inches(3.75), Inches(0.08), Inches(2.0), fill=GOLD)
    txt(s, x + Inches(0.28), Inches(3.95), Inches(3.55), Inches(0.45), [[(t, 15, True, PURPLE)]])
    txt(s, x + Inches(0.28), Inches(4.5), Inches(3.55), Inches(1.05), [[(d, 13, False, DARK)]], ls=1.15)
txt(s, Inches(0.55), Inches(6.0), Inches(12.2), Inches(0.9),
    [[("当前窗口：首期自 2026 年第四季度启动。11—12 月先完成合作机制与需求梳理，原 10—12 月活动排期按实际推进重排，优先保证首批项目质量。", 12, False, GREY)]], ls=1.15)

# ============================================================ 3 主体关系
s = new(); header(s, "主体  ·  WHO", "合作主体关系"); footer(s)
txt(s, Inches(0.55), Inches(1.22), Inches(12.2), Inches(0.4),
    [[("先把签约与实施主体说清楚，避免把平台、场景和实施单位混为一家。", 13, False, GREY)]])
blocks = [
    ("学术及研究牵头", "复旦大学住房政策研究中心", "研究统筹、方案策划、专家与学术资源协调、白皮书与课题组织、企业对接协调。"),
    ("拟协调高校资源", "复旦相关学院、研究机构\n及科技创新平台", "按项目和审批情况参与专题研究、专家交流、企业培训、技术对接与成果转化。"),
    ("产业及空间合作", "上海东方枢纽投资建设\n发展集团有限公司", "上海市政府批准组建的市级国资体系相关平台。提供产业场景、招商需求与国际商务平台。"),
    ("场景与实施", "国合区  ·  汇在东方\n复旦大学科技园等", "上海东方枢纽国际商务合作区是具体合作场景；汇在东方等为开发建设与项目实施主体。"),
]
for i, (k, name, d) in enumerate(blocks):
    x = Inches(0.4) + (i % 4) * Inches(3.22)
    card(s, x, Inches(1.8), Inches(3.05), Inches(3.55))
    rect(s, x, Inches(1.8), Inches(3.05), Inches(0.12), fill=GOLD if i == 2 else PURPLE)
    txt(s, x + Inches(0.18), Inches(2.05), Inches(2.7), Inches(0.35), [[(k, 11, True, GOLD)]])
    txt(s, x + Inches(0.18), Inches(2.4), Inches(2.7), Inches(1.15), [[(name, 14, True, PLUM)]], ls=1.05)
    txt(s, x + Inches(0.18), Inches(3.65), Inches(2.7), Inches(1.5), [[(d, 11, False, DARK)]], ls=1.12)
txt(s, Inches(0.5), Inches(5.55), Inches(12.3), Inches(1.35),
    [[("合同安排建议", 13, True, PURPLE)],
     [("不建议以复旦大学名义签署全部商业服务合同。住房政策研究中心承担学术研究与项目协调；相关学院、科技园及其他实施单位按授权参与；东方枢纽集团或其指定子公司承担对应采购与项目合作。学校名称使用、经费与责任，以最终授权和书面协议为准。", 12, False, DARK)]],
    ls=1.15)

# ============================================================ 4 目标
s = new(); header(s, "目标  ·  GOAL", "总体目标与五位一体合作体系"); footer(s)
txt(s, Inches(0.55), Inches(1.22), Inches(12.2), Inches(0.7),
    [[("服务上海国际经济、金融、贸易、航运和科技创新中心建设，发挥国合区在国际商务交流、科技合作、国际会展、国际培训上的制度与区位优势。", 13, False, DARK)]], ls=1.12)
funcs = ["国际商务交流", "科技合作", "国际会展", "国际培训"]
for i, f in enumerate(funcs):
    x = Inches(0.5) + i * Inches(3.2)
    rrect(s, x, Inches(2.1), Inches(3.0), Inches(0.7), fill=LAV, line=LAVED)
    txt(s, x, Inches(2.1), Inches(3.0), Inches(0.7),
        [[(f, 14, True, PURPLE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
five = [
    ("高校智库研究", "区域定位、站城融合、空间运营与总部经济"),
    ("产业创新资源", "科技企业、校友网络与成果转化"),
    ("国际商务交流", "商协会、领事机构与出海服务"),
    ("精准招商服务", "筛企业、做对接、跟踪落地"),
    ("空间载体运营", "办公、商务与文商旅场景匹配"),
]
for i, (t, d) in enumerate(five):
    x = Inches(0.4) + i * Inches(2.56)
    card(s, x, Inches(3.1), Inches(2.42), Inches(2.15))
    txt(s, x + Inches(0.15), Inches(3.25), Inches(2.12), Inches(0.7), [[(f"0{i+1}", 12, True, GOLD)], [(t, 14, True, PLUM)]], after=1)
    txt(s, x + Inches(0.15), Inches(4.15), Inches(2.12), Inches(0.9), [[(d, 12, False, DARK)]], ls=1.1)
txt(s, Inches(0.55), Inches(5.5), Inches(12.2), Inches(1.3),
    [[("共同推动东方枢纽从大型交通与商务基础设施，进一步发展为具有产业集聚、国际交流和创新服务能力的综合性产业平台。", 14, True, PLUM)]], ls=1.15)

# ============================================================ 5 契合点
s = new(); header(s, "契合  ·  FIT", "双方资源与合作价值"); footer(s)
rows = [
    ("城市发展、住房政策、城市更新研究", "区域定位、站城融合、空间运营、总部经济", "区域发展策略与运营咨询"),
    ("科技园、科创项目与校友企业网络", "航空、跨境贸易、科技服务精准招商", "可持续的招商项目来源"),
    ("AI、集成电路、生物医药等学术产业资源", "高附加值科技服务与创新项目导入", "科技成果对接与产业集聚"),
    ("白皮书、产业研究与企业数据库经验", "重点产业识别、企业画像、招商决策", "提升招商精准度"),
    ("专家、企业家、投资机构与国际网络", "国际商务交流、论坛与项目路演", "提升影响力与转化效率"),
    ("继续教育与国际培训资源", "国际培训、产业人才与企业管理培训", "可持续的人才服务产品"),
]
# header row
ys = Inches(1.28)
rect(s, Inches(0.4), ys, Inches(4.15), Inches(0.42), fill=PLUM)
rect(s, Inches(4.55), ys, Inches(4.35), Inches(0.42), fill=PURPLE)
rect(s, Inches(8.9), ys, Inches(4.0), Inches(0.42), fill=GOLD)
for x, w, lab in [
    (Inches(0.4), Inches(4.15), "复旦方面的合作资源"),
    (Inches(4.55), Inches(4.35), "东方枢纽的实际需求"),
    (Inches(8.9), Inches(4.0), "预期合作价值"),
]:
    txt(s, x, ys, w, Inches(0.42), [[(lab, 12, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
for i, (a, b, c) in enumerate(rows):
    y = Inches(1.7) + i * Inches(0.85)
    bgc = WHITE if i % 2 == 0 else LAV
    rect(s, Inches(0.4), y, Inches(4.15), Inches(0.85), fill=bgc, line=LINE, lw=0.5)
    rect(s, Inches(4.55), y, Inches(4.35), Inches(0.85), fill=bgc, line=LINE, lw=0.5)
    rect(s, Inches(8.9), y, Inches(4.0), Inches(0.85), fill=bgc, line=LINE, lw=0.5)
    txt(s, Inches(0.52), y, Inches(3.95), Inches(0.85), [[(a, 12, False, DARK)]], anchor=MSO_ANCHOR.MIDDLE, ls=1.05)
    txt(s, Inches(4.67), y, Inches(4.15), Inches(0.85), [[(b, 12, False, DARK)]], anchor=MSO_ANCHOR.MIDDLE, ls=1.05)
    txt(s, Inches(9.02), y, Inches(3.75), Inches(0.85), [[(c, 12, True, PURPLE)]], anchor=MSO_ANCHOR.MIDDLE, ls=1.05)

# ============================================================ 6 原则
s = new(); header(s, "原则  ·  PRINCIPLES", "三条合作原则"); footer(s)
principles = [
    ("01", "以产业落地为导向", "不以活动数量作为唯一成果。活动、研究和招商都要回答：能否带来真实企业、真实需求和可跟踪项目。"),
    ("02", "以研究与真实需求为基础", "形成可执行、可评价的服务项目。每项合作有范围、有交付、有书面确认。"),
    ("03", "以具体项目带动长期平台", "先把能签约的项目做实，再逐步形成稳定合作机制，不先搭一个空的战略框架。"),
]
for i, (n, t, d) in enumerate(principles):
    y = Inches(1.45) + i * Inches(1.75)
    card(s, Inches(0.7), y, Inches(11.9), Inches(1.55))
    rrect(s, Inches(0.95), y + Inches(0.35), Inches(1.15), Inches(0.85), fill=PLUM)
    txt(s, Inches(0.95), y + Inches(0.35), Inches(1.15), Inches(0.85),
        [[(n, 20, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(2.4), y + Inches(0.28), Inches(9.8), Inches(0.45), [[(t, 18, True, PLUM)]])
    txt(s, Inches(2.4), y + Inches(0.8), Inches(9.8), Inches(0.55), [[(d, 14, False, DARK)]])

# ============================================================ 7 七大内容总览
s = new(); header(s, "内容  ·  SCOPE", "七项重点合作内容"); footer(s)
items = [
    ("01", "城市发展与国际商务研究", "专题研究与决策咨询，形成区域策略和运营建议。"),
    ("02", "科技企业精准招商", "四类产业资源导入，建立“两库一清单”。"),
    ("03", "国际商务与科创系列活动", "两个活动品牌：成果论坛 + 闭门对接。"),
    ("04", "国际商务与企业出海", "商协会、领事资源、出海研究与外资需求调查。"),
    ("05", "科技成果转化与产学研", "需求导向的技术对接、路演与应用试点。"),
    ("06", "国际商务人才研修", "协调终身教育等授权单位，开发专题课程。"),
    ("07", "文商旅与特色运营", "站前商业文化空间的场景匹配与可行性研究。"),
]
for i, (n, t, d) in enumerate(items):
    col, row = i % 2, i // 2
    if i == 6:
        x, y = Inches(3.45), Inches(1.35) + 3 * Inches(1.4)
        # last one centered? better keep grid and let 7 sit left of row 4... 4 rows with last full? 
    x = Inches(0.45) + (i % 2) * Inches(6.45)
    y = Inches(1.3) + (i // 2) * Inches(1.42)
    if i == 6:
        x = Inches(0.45)
    card(s, x, y, Inches(6.25) if i < 6 else Inches(12.4), Inches(1.28))
    rrect(s, x + Inches(0.15), y + Inches(0.28), Inches(0.7), Inches(0.7), fill=PURPLE)
    txt(s, x + Inches(0.15), y + Inches(0.28), Inches(0.7), Inches(0.7),
        [[(n, 12, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + Inches(1.05), y + Inches(0.18), Inches(5.0 if i < 6 else 10.8), Inches(0.4), [[(t, 15, True, PLUM)]])
    txt(s, x + Inches(1.05), y + Inches(0.62), Inches(5.0 if i < 6 else 10.8), Inches(0.5), [[(d, 12, False, DARK)]])

# ============================================================ 8 研究
s = new(); header(s, "合作一  ·  RESEARCH", "共建城市发展与国际商务产业研究机制"); footer(s)
txt(s, Inches(0.55), Inches(1.22), Inches(12.2), Inches(0.45),
    [[("住房政策研究中心发挥城市发展与空间政策研究优势，对接集团的区域开发、产业招商和载体运营需求。", 13, False, GREY)]])
topics = [
    "国合区产业定位与发展策略",
    "上海东站与浦东机场联动的站城融合、商务空间与运营模式",
    "总部经济与现代服务业集聚路径",
    "国际商务载体、产业办公及商业文化设施的运营效率",
    "AI 与数字化在园区运营、智慧楼宇、招商管理和空间治理中的应用",
    "长三角产业协同与企业资源导入机制",
]
for i, t in enumerate(topics):
    col, row = i % 2, i // 2
    x = Inches(0.45) + col * Inches(6.45)
    y = Inches(1.8) + row * Inches(1.05)
    card(s, x, y, Inches(6.25), Inches(0.92))
    rect(s, x, y, Inches(0.08), Inches(0.92), fill=GOLD)
    txt(s, x + Inches(0.28), y, Inches(5.8), Inches(0.92),
        [[(f"{i+1}   {t}", 13, False, DARK)]], anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(0.55), Inches(5.05), Inches(12.2), Inches(0.35), [[("首批建议研究成果", 14, True, PURPLE)]])
papers = [
    "《东方枢纽国际商务合作区产业发展与空间运营策略研究》",
    "《重点产业招商机会及目标企业研究报告》",
    "《站城融合与总部经济发展研究》",
]
for i, p in enumerate(papers):
    x = Inches(0.45) + i * Inches(4.25)
    rrect(s, x, Inches(5.5), Inches(4.05), Inches(1.25), fill=PLUM)
    txt(s, x + Inches(0.2), Inches(5.5), Inches(3.65), Inches(1.25),
        [[(p, 13, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, ls=1.1)

# ============================================================ 9 招商
s = new(); header(s, "合作二  ·  INVESTMENT", "科技企业精准招商与产业资源导入"); footer(s)
inds = [
    ("民用航空及配套", "航空装备研发、供应链、航材、航空贸易、航空金融与专业服务。吸引有国际商务、贸易结算、销售服务和总部办公需求的企业。"),
    ("跨境贸易与供应链", "跨境贸易、国际货代、供应链管理、跨境电商、数字贸易和专业服务企业。"),
    ("人工智能与科技服务", "人工智能、集成电路、智能装备、数字服务和高端科技服务。优先选择有国际业务或与航空、贸易链协同的企业。"),
    ("总部经济与专业服务", "科技企业区域总部、国际业务中心、跨境投资与贸易服务机构、专业咨询机构。"),
]
for i, (t, d) in enumerate(inds):
    y = Inches(1.25) + i * Inches(1.05)
    card(s, Inches(0.4), y, Inches(8.15), Inches(0.95))
    rrect(s, Inches(0.55), y + Inches(0.22), Inches(0.5), Inches(0.5), fill=PURPLE)
    txt(s, Inches(0.55), y + Inches(0.22), Inches(0.5), Inches(0.5),
        [[(str(i + 1), 14, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(1.2), y + Inches(0.08), Inches(7.15), Inches(0.32), [[(t, 13, True, PLUM)]])
    txt(s, Inches(1.2), y + Inches(0.4), Inches(7.15), Inches(0.48), [[(d, 11, False, DARK)]], ls=1.05)
card(s, Inches(8.75), Inches(1.25), Inches(4.15), Inches(5.15), fill=PLUM)
txt(s, Inches(8.95), Inches(1.45), Inches(3.8), Inches(0.4), [[("两库一清单", 16, True, GOLD)]])
libs = [
    ("目标企业库", "行业、阶段、融资、国际业务与扩张需求"),
    ("合作项目库", "对接、考察、谈判与推进状态"),
    ("载体匹配清单", "空间、功能、政策、商务与人才需求"),
]
for i, (t, d) in enumerate(libs):
    txt(s, Inches(8.95), Inches(2.05) + i * Inches(1.05), Inches(3.75), Inches(0.9),
        [[(t, 14, True, WHITE)], [(d, 12, False, LILAC)]], after=2)
txt(s, Inches(8.95), Inches(5.3), Inches(3.75), Inches(0.85),
    [[("每季度一次重点项目筛选会，形成明确的企业对接任务。", 12, False, WHITE)]], ls=1.1)

# ============================================================ 10 活动品牌
s = new(); header(s, "合作三  ·  PROGRAMS", "两个活动品牌，而不是普通论坛"); footer(s)
card(s, Inches(0.4), Inches(1.3), Inches(6.2), Inches(4.35))
rect(s, Inches(0.4), Inches(1.3), Inches(6.2), Inches(0.7), fill=PURPLE)
txt(s, Inches(0.6), Inches(1.3), Inches(5.85), Inches(0.7),
    [[("品牌一", 11, True, GOLD)], [("  东方枢纽 · 复旦国际商务与产业创新论坛", 13, True, WHITE)]],
    anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(0.6), Inches(2.15), Inches(5.8), Inches(0.7),
    [[("中高端交流与成果发布。政策研究、产业趋势、国际商务合作。", 13, False, DARK)]], ls=1.1)
for i, t in enumerate(["科技企业国际化与出海新机遇", "国际商务枢纽与长三角总部经济", "人工智能与数字贸易", "城市更新、站城融合与商务空间"]):
    txt(s, Inches(0.6), Inches(2.9) + i * Inches(0.48), Inches(5.8), Inches(0.45),
        [[(f"·  {t}", 13, False, PLUM)]], anchor=MSO_ANCHOR.MIDDLE)
card(s, Inches(6.8), Inches(1.3), Inches(6.1), Inches(4.35))
rect(s, Inches(6.8), Inches(1.3), Inches(6.1), Inches(0.7), fill=PLUM)
txt(s, Inches(7.0), Inches(1.3), Inches(5.75), Inches(0.7),
    [[("品牌二", 11, True, GOLD)], [("  东方枢纽产业招商闭门对接会", 14, True, WHITE)]],
    anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(7.0), Inches(2.15), Inches(5.7), Inches(0.85),
    [[("20—30 人精准招商。政策介绍、项目推介、需求交流、实地考察、后续跟踪。", 13, False, DARK)]], ls=1.1)
txt(s, Inches(7.0), Inches(3.15), Inches(5.7), Inches(2.2),
    [[("市区一场：", 13, True, PURPLE), ("经批准使用的复旦或科技园场地，扩大企业线索。", 13, False, DARK)],
     [("国合区一场：", 13, True, PURPLE), ("有明确意向的企业现场考察功能、环境与载体。", 13, False, DARK)],
     [("每场必交：", 13, True, GOLD), ("企业需求清单、招商意向清单、跟踪反馈报告。", 13, False, DARK)]],
    after=8, ls=1.05)
txt(s, Inches(0.5), Inches(5.85), Inches(12.3), Inches(1.0),
    [[("双场景联动的目的，是把“听论坛”变成“看空间、对需求、能跟踪”。活动结束后必须有清单，避免只办一场、没有转化。", 13, True, PLUM)]], ls=1.1)

# ============================================================ 11 出海 + 产学研
s = new(); header(s, "合作四、五  ·  GLOBAL & TECH", "国际商务出海服务，与科技成果转化"); footer(s)
card(s, Inches(0.4), Inches(1.28), Inches(6.25), Inches(5.45))
txt(s, Inches(0.6), Inches(1.42), Inches(5.9), Inches(0.4), [[("国际商务与科技企业出海", 16, True, PLUM)]])
for i, t in enumerate([
    "国际商协会、驻沪领事机构及外资企业对接",
    "科技企业出海政策与海外市场研究",
    "国际产业链合作及跨境投资项目交流",
    "国际企业在沪投资发展需求调查",
    "海外科技项目与上海产业资源对接",
    "国际专业服务机构合作交流",
]):
    txt(s, Inches(0.65), Inches(2.0) + i * Inches(0.48), Inches(5.8), Inches(0.45),
        [[(f"{i+1}   {t}", 13, False, DARK)]], anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(0.65), Inches(5.05), Inches(5.8), Inches(1.4),
    [[("可探索“国际科技企业进入中国市场交流会”“长三角科技企业全球化发展闭门会”。境外机构与专家参与，以实际邀约及相关规定为准。", 12, False, GREY)]], ls=1.12)
card(s, Inches(6.85), Inches(1.28), Inches(6.05), Inches(5.45))
txt(s, Inches(7.05), Inches(1.42), Inches(5.7), Inches(0.4), [[("科技成果转化与产学研", 16, True, PLUM)]])
for i, t in enumerate([
    "AI 与智慧商务空间",
    "智能机器人与智慧园区服务",
    "民用航空数字化、供应链与智能装备",
    "跨境贸易数字化服务",
    "高端专业服务与创新企业孵化",
    "未来健康及国际科技服务",
]):
    txt(s, Inches(7.1), Inches(2.0) + i * Inches(0.42), Inches(5.6), Inches(0.4),
        [[(f"·  {t}", 13, False, DARK)]], anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(7.1), Inches(4.65), Inches(5.6), Inches(1.85),
    [[("形式：技术需求发布、团队对接、成果展示、路演、应用试点、产业化咨询。", 12, False, DARK)],
     [("具备条件的项目，探索“高校科研团队 — 企业 — 东方枢纽产业空间”联合对接。知识产权与技术转移，由权利主体按学校制度另行协商。", 12, False, GREY)]],
    after=6, ls=1.1)

# ============================================================ 12 培训 + 文商旅
s = new(); header(s, "合作六、七  ·  TALENT & PLACE", "人才研修与文商旅场景运营"); footer(s)
card(s, Inches(0.4), Inches(1.28), Inches(6.25), Inches(5.45))
txt(s, Inches(0.6), Inches(1.45), Inches(5.85), Inches(0.7),
    [[("国际商务人才培养与高端研修", 16, True, PLUM)],
     [("拟协调终身教育学院等经授权的培训单位，不由研究中心自行承诺办学资质。", 12, False, GREY)]], after=2)
courses = [
    "国际商务与全球产业链管理高级研修班",
    "中国科技企业出海与跨境投资实务",
    "人工智能与产业数字化高级管理者研修",
    "国际商务合作区运营与产业招商专题",
    "城市更新、产业空间与智慧运营专题",
]
for i, t in enumerate(courses):
    y = Inches(2.45) + i * Inches(0.58)
    rrect(s, Inches(0.65), y, Inches(5.75), Inches(0.5), fill=LAV)
    txt(s, Inches(0.8), y, Inches(5.5), Inches(0.5), [[(t, 13, False, DARK)]], anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(0.65), Inches(5.5), Inches(5.75), Inches(1.0),
    [[("课程名称、师资、收费、结业证明和招生，由实际承办的授权单位确定。", 12, False, GREY)]], ls=1.1)
card(s, Inches(6.85), Inches(1.28), Inches(6.05), Inches(5.45))
txt(s, Inches(7.05), Inches(1.45), Inches(5.65), Inches(0.7),
    [[("文商旅融合与特色运营", 16, True, PLUM)],
     [("结合站前办公、商业及文化设施，做场景匹配和商业可行性研究。", 12, False, GREY)]], after=2)
places = [
    "国际文化交流与科技艺术展演",
    "沉浸式文化体验及数字文旅",
    "国际商务活动与文化商业空间融合",
    "科技展览展示及产业体验场景",
    "枢纽配套商业与文化空间运营",
]
for i, t in enumerate(places):
    y = Inches(2.45) + i * Inches(0.58)
    rrect(s, Inches(7.1), y, Inches(5.55), Inches(0.5), fill=LAV)
    txt(s, Inches(7.25), y, Inches(5.3), Inches(0.5), [[(t, 13, False, DARK)]], anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(7.1), Inches(5.5), Inches(5.55), Inches(1.0),
    [[("优先选择有成熟内容、明确经营主体和真实投资需求的项目。", 12, False, GREY)]], ls=1.1)

# ============================================================ 13 首批项目
s = new(); header(s, "首期  ·  FIRST BATCH", "首批合作项目计划"); footer(s)
head = ["项目", "主要工作", "拟实施时间", "预期交付"]
widths = [2.7, 4.3, 2.5, 2.7]
xs = [Inches(0.4)]
for w in widths[:-1]:
    xs.append(xs[-1] + Inches(w))
y0 = Inches(1.3)
for x, w, hname in zip(xs, widths, head):
    rect(s, x, y0, Inches(w), Inches(0.48), fill=PLUM)
    txt(s, x, y0, Inches(w), Inches(0.48), [[(hname, 12, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
data = [
    ("合作机制与需求梳理", "建立联合工作组，确认产业和研究需求", "2026年11—12月", "年度合作项目清单"),
    ("出海与国际商务活动", "白皮书交流、政策解读、国际商务对接", "2026年12月—2027年一季度", "1 场主题活动、企业线索清单"),
    ("航空与国际供应链闭门会", "组织航空产业链及贸易企业赴国合区考察", "2027年一季度", "1 场精准对接会、项目跟踪表"),
    ("重点产业研究", "分析产业布局、重点企业与招商路径", "2027年上半年", "1 份专题研究报告"),
    ("国际商务人才培训", "设计课程、师资及培训实施模式", "2027年上半年", "1 套培训方案及试点计划"),
    ("文商旅及科技应用对接", "筛选场景运营项目和技术应用项目", "2027年二至四季度", "可行性与合作建议"),
]
for i, row in enumerate(data):
    y = Inches(1.78) + i * Inches(0.78)
    bgc = WHITE if i % 2 == 0 else LAV
    for x, w, val in zip(xs, widths, row):
        rect(s, x, y, Inches(w), Inches(0.78), fill=bgc, line=LINE, lw=0.5)
        txt(s, x + Inches(0.08), y, Inches(w) - Inches(0.12), Inches(0.78),
            [[(val, 11, True if x == xs[0] else False, PLUM if x == xs[0] else DARK)]],
            anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, ls=1.05)
txt(s, Inches(0.45), Inches(6.55), Inches(12.4), Inches(0.5),
    [[("原计划 2026 年 10—12 月的活动，按实际推进重新排期。当前优先把 11—12 月的机制梳理和首场活动做扎实。", 12, False, GREY)]])

# ============================================================ 14 职责
s = new(); header(s, "组织  ·  GOVERNANCE", "职责分工与联合工作机制"); footer(s)
roles = [
    ("住房政策研究中心", [
        "合作议题研究与总体方案策划",
        "城市发展、空间运营及产业政策研究",
        "专家与学术资源协调",
        "白皮书、研究报告与课题组织",
        "企业对接、主题活动及项目推进协调",
    ]),
    ("东方枢纽集团及实施主体", [
        "提供产业定位、招商方向及载体需求",
        "组织空间、商务设施和项目场景考察",
        "协调园区运营、招商和商务活动资源",
        "企业落地的政策咨询与服务对接",
        "按协议推动采购、实施与验收",
    ]),
    ("复旦相关学院与科创平台", [
        "按项目及审批参与专题研究",
        "专家交流与课程资源支持",
        "技术对接与成果展示",
        "科技成果转化的合规路径",
        "不替代授权单位对外签约办学或技术转移",
    ]),
]
for i, (title, lines) in enumerate(roles):
    x = Inches(0.4) + i * Inches(4.28)
    card(s, x, Inches(1.28), Inches(4.1), Inches(4.15))
    rect(s, x, Inches(1.28), Inches(4.1), Inches(0.6), fill=PURPLE if i != 1 else PLUM)
    txt(s, x + Inches(0.15), Inches(1.28), Inches(3.8), Inches(0.6),
        [[(title, 14, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)
    for j, line in enumerate(lines):
        txt(s, x + Inches(0.2), Inches(2.05) + j * Inches(0.58), Inches(3.7), Inches(0.55),
            [[(f"·  {line}", 12, False, DARK)]], anchor=MSO_ANCHOR.MIDDLE)
rrect(s, Inches(0.4), Inches(5.6), Inches(12.5), Inches(1.25), fill=LAV)
txt(s, Inches(0.6), Inches(5.72), Inches(12.1), Inches(1.0),
    [[("联合项目工作组：月度沟通、季度重点项目推进、年度成果评价。重大活动、课题和招商项目分别签专项协议，写清实施主体、工作内容、交付标准、费用与权责。", 13, False, PLUM)]],
    anchor=MSO_ANCHOR.MIDDLE, ls=1.12)

# ============================================================ 15 经费
s = new(); header(s, "经费  ·  COMMERCIAL", "框架合作 + 专项项目合同"); footer(s)
txt(s, Inches(0.55), Inches(1.22), Inches(12.2), Inches(0.4),
    [[("分别核定内容与预算。财政专项资金可以研究申报，但不作为合作经费的必然来源。", 13, False, GREY)]])
fees = [
    ("第一类", "专题研究与咨询", "产业发展、招商策略、城市空间运营。按研究范围、调研量、专家投入和交付成果协商费用。"),
    ("第二类", "产业活动与商务对接", "按规模、场地、嘉宾、企业邀约、会务执行和成果跟踪分项核定。"),
    ("第三类", "产业招商与项目服务", "企业筛选、需求调查、商务对接、落地跟踪。涉及费用的，按国资采购、招商管理和财务制度执行。"),
    ("第四类", "教育培训与科技合作", "由具备资质和授权的实施单位，按具体培训或技术合作项目另行制定方案。"),
]
for i, (k, t, d) in enumerate(fees):
    y = Inches(1.75) + i * Inches(1.2)
    card(s, Inches(0.5), y, Inches(12.3), Inches(1.08))
    rrect(s, Inches(0.68), y + Inches(0.24), Inches(1.35), Inches(0.6), fill=PLUM)
    txt(s, Inches(0.68), y + Inches(0.24), Inches(1.35), Inches(0.6),
        [[(k, 13, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(2.2), y + Inches(0.12), Inches(10.3), Inches(0.35), [[(t, 14, True, PURPLE)]])
    txt(s, Inches(2.2), y + Inches(0.5), Inches(10.3), Inches(0.48), [[(d, 12, False, DARK)]])

# ============================================================ 16 三个提醒
s = new(); header(s, "边界  ·  CONDITIONS", "费用与合作条件上的三个提醒"); footer(s)
warns = [
    ("纪要报价不是合同",
     "此前沟通中出现过 30 人小型活动 3.6 万元、市区专场 8.8 万元，以及单场 3 万—8 万元的预算设想。这些是洽谈记录，可作为谈判参考，不能视为已落实的活动预算或正式合同。"),
    ("专项资金有门槛",
     "国合区会议、培训等专项资金可以依法依规研究申报。现行细则中，30 人级会议活动的支持条件还包括举办时间超过两天、境外参会人数超过 20 人等。普通 20—30 人招商沙龙，不能直接认定为符合补贴条件。"),
    ("签约主体要分开",
     "不建议以复旦大学名义签署所有商业服务合同。研究中心做学术研究与项目协调；学院、科技园等按授权参与；集团或其指定子公司承担采购与项目合作。这样后续收费、交付和执行才站得住。"),
]
for i, (t, d) in enumerate(warns):
    y = Inches(1.35) + i * Inches(1.8)
    card(s, Inches(0.5), y, Inches(12.3), Inches(1.65))
    rect(s, Inches(0.5), y, Inches(0.1), Inches(1.65), fill=GOLD)
    txt(s, Inches(0.85), y + Inches(0.15), Inches(11.7), Inches(0.4), [[(f"0{i+1}   {t}", 16, True, PLUM)]])
    txt(s, Inches(0.85), y + Inches(0.6), Inches(11.7), Inches(0.9), [[(d, 13, False, DARK)]], ls=1.12)

# ============================================================ 17 2027目标
s = new(); header(s, "目标  ·  2027", "2027 年度建议合作目标"); footer(s)
txt(s, Inches(0.55), Inches(1.22), Inches(12.2), Inches(0.4),
    [[("以下是协商用的工作目标，不构成招商落地数量或合作结果的保证。评价以实际完成、书面确认和可核查成果为准。", 12, False, GREY)]])
goals = [
    ("2 份", "专题研究与产业报告"),
    ("6—8 场", "产业与国际商务活动"),
    ("80—100 家", "初步筛选重点目标企业"),
    ("20—30 家", "有效招商对接线索"),
    ("8—12 个", "实地考察或深度洽谈"),
    ("2—4 个", "落地或实质性合作成果"),
    ("1—2 项", "国际商务与产业管理培训"),
]
for i, (n, lab) in enumerate(goals):
    if i < 4:
        x = Inches(0.45) + i * Inches(3.2)
        y = Inches(1.85)
    else:
        x = Inches(2.05) + (i - 4) * Inches(3.2)
        y = Inches(4.15)
    card(s, x, y, Inches(3.02), Inches(1.95))
    txt(s, x, y + Inches(0.3), Inches(3.02), Inches(0.7),
        [[(n, 26, True, PURPLE)]], align=PP_ALIGN.CENTER)
    txt(s, x + Inches(0.15), y + Inches(1.1), Inches(2.72), Inches(0.65),
        [[(lab, 13, False, DARK)]], align=PP_ALIGN.CENTER)

# ============================================================ 18 四阶段
s = new(); header(s, "路径  ·  ROADMAP", "四阶段推进"); footer(s)
phases = [
    ("第一阶段", "现在", "明确关系与项目边界", "双方负责人确认牵头单位、业务需求、首批项目和实施主体，形成可执行任务清单。"),
    ("第二阶段", "试点", "两场活动 + 一项研究", "优先做科技企业出海主题活动、民用航空产业链闭门会、产业招商研究。用较小投入验证资源、执行和转化。"),
    ("第三阶段", "年度化", "协议与项目化管理", "在试点基础上拓展研究、精准招商、国际培训和成果转化，形成稳定的年度项目安排。"),
    ("第四阶段", "平台化", "长期合作载体", "条件成熟后，探索常态化联合研究机制或国际商务产业创新合作平台。名称、架构和对外发布按程序审批。"),
]
for i, (ph, tag, t, d) in enumerate(phases):
    y = Inches(1.3) + i * Inches(1.4)
    rrect(s, Inches(0.45), y, Inches(1.7), Inches(1.22), fill=PLUM if i < 2 else PURPLE)
    txt(s, Inches(0.45), y, Inches(1.7), Inches(1.22),
        [[(ph, 13, True, WHITE)], [(tag, 12, True, GOLD)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, after=2)
    card(s, Inches(2.3), y, Inches(10.55), Inches(1.22))
    txt(s, Inches(2.5), y + Inches(0.12), Inches(10.15), Inches(0.4), [[(t, 15, True, PLUM)]])
    txt(s, Inches(2.5), y + Inches(0.55), Inches(10.15), Inches(0.55), [[(d, 13, False, DARK)]])

# ============================================================ 19 三个优先项目
s = new(); header(s, "决策  ·  PRIORITY", "建议优先谈成的三个项目"); footer(s)
txt(s, Inches(0.55), Inches(1.2), Inches(12.2), Inches(0.4),
    [[("完整方案覆盖面较广。第一次推进，先抓最容易形成项目合同的三件事。", 14, True, PLUM)]])
pri = [
    ("第一优先", "产业招商与城市空间\n运营专题研究",
     "最契合住房政策研究中心的专业定位，也直接回应产业导入、办公空间和总部经济需求。",
     "先争取一项有明确范围、成果和预算的委托研究。"),
    ("第二优先", "国际商务及航空产业\n精准招商",
     "纪要里已有沟通基础。按“市区一场成果发布 + 国合区一场精准对接”执行。",
     "分别形成企业需求清单和招商成果清单，不做普通论坛。"),
    ("第三优先", "国际商务与\n企业家研修",
     "协调终身教育学院等资源，面向跨境贸易企业、高管和园区运营人员。",
     "可形成长期业务，但须先落实学校授权和实际承办单位。"),
]
for i, (tag, t, d, act) in enumerate(pri):
    x = Inches(0.4) + i * Inches(4.28)
    card(s, x, Inches(1.75), Inches(4.1), Inches(4.95))
    rect(s, x, Inches(1.75), Inches(4.1), Inches(1.35), fill=GOLD if i == 0 else (PURPLE if i == 1 else PLUM))
    title_color = PLUM if i == 0 else WHITE
    txt(s, x + Inches(0.2), Inches(1.82), Inches(3.7), Inches(0.28), [[(tag, 12, True, title_color)]])
    txt(s, x + Inches(0.2), Inches(2.1), Inches(3.7), Inches(0.9), [[(t, 16, True, title_color)]], ls=1.05)
    txt(s, x + Inches(0.22), Inches(3.25), Inches(3.65), Inches(1.55), [[(d, 13, False, DARK)]], ls=1.15)
    txt(s, x + Inches(0.22), Inches(4.9), Inches(3.65), Inches(1.55),
        [[("建议动作", 12, True, PURPLE)], [(act, 13, False, DARK)]], after=3, ls=1.12)

# The first card title got messy. Let me fix by rewriting that slide more cleanly after I see render.
# I'll adjust the loop - I already wrote a confusing title for i==0. Let me fix the source before running.
# I'll do a search replace after writing... actually I'll fix now before the file is used. The write already happened in this call.
# I'll patch after.

# ============================================================ 20 成效
s = new(); header(s, "成效  ·  VALUE", "预期合作成效"); footer(s)
vals = [
    ("对于东方枢纽集团",
     "提升招商能力与项目储备质量，扩大高校科技创新和企业导入渠道，丰富国际商务和高端培训内容，为大型产业载体的中长期运营提供研究与产业支持。"),
    ("对于住房政策研究中心",
     "把城市发展、国际商务空间、城市更新和产业运营研究落到真实场景，形成可示范的政策研究与产学研案例。"),
    ("对于相关院系和企业",
     "打通成果转化、企业国际化、人才培养和场景应用，增强高校科研成果与上海重大产业平台的协同。"),
]
for i, (t, d) in enumerate(vals):
    y = Inches(1.4) + i * Inches(1.55)
    card(s, Inches(0.5), y, Inches(12.3), Inches(1.4))
    rect(s, Inches(0.5), y, Inches(0.1), Inches(1.4), fill=PURPLE)
    txt(s, Inches(0.85), y + Inches(0.15), Inches(11.7), Inches(0.4), [[(t, 16, True, PLUM)]])
    txt(s, Inches(0.85), y + Inches(0.6), Inches(11.7), Inches(0.65), [[(d, 13, False, DARK)]], ls=1.1)
txt(s, Inches(0.55), Inches(6.15), Inches(12.2), Inches(0.7),
    [[("长期模式：研究引领、产业导入、国际链接、空间承载、项目落地。", 15, True, PURPLE)]],
    anchor=MSO_ANCHOR.MIDDLE)

# ============================================================ 21 收束
s = prs.slides.add_slide(BLANK)
PAGE[0] += 1
rect(s, 0, 0, SW, SH, fill=PLUM)
rect(s, 0, 0, Inches(0.18), SH, fill=GOLD)
txt(s, Inches(0.75), Inches(0.7), Inches(11.5), Inches(0.4),
    [[("建议的下一步", 14, True, GOLD)]])
txt(s, Inches(0.75), Inches(1.2), Inches(11.8), Inches(1.1),
    [[("先谈成三个能签约的项目，", 28, True, WHITE)],
     [("再谈年度合作。", 28, True, WHITE)]], after=2)
nexts = [
    ("1", "确认一项专题研究的范围、成果和预算"),
    ("2", "确认“市区一场 + 国合区一场”的航空/国际商务精准招商"),
    ("3", "明确研修项目的授权承办单位，再谈课程方案"),
    ("4", "11—12 月成立联合工作组，形成 2027 年项目清单"),
]
for i, (n, t) in enumerate(nexts):
    y = Inches(2.7) + i * Inches(0.85)
    txt(s, Inches(0.75), y, Inches(0.6), Inches(0.6), [[(n, 20, True, GOLD)]], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(1.5), y, Inches(10.5), Inches(0.6), [[(t, 18, False, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(0.75), Inches(6.5), Inches(11.5), Inches(0.6),
    [[("本方案为合作建议稿。合作主体、学校名称使用、项目经费与实施责任，以最终授权和书面协议为准。", 12, False, LILAC)]], ls=1.1)

out = "复旦住房政策研究中心×东方枢纽集团战略合作计划.pptx"
prs.save(out)
print("saved", out, "slides", len(prs.slides))
