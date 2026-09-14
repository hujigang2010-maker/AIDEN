# -*- coding: utf-8 -*-
"""生成 PPT：东昇聚变政府事务与上海产业落地 90 天工作设想。"""
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

import content as C

OUT = Path(__file__).resolve().parent.parent / "deliverables"
OUT.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT / "东昇聚变_政府事务与上海产业落地_90天工作设想.pptx"

NAVY = RGBColor(0x07, 0x1A, 0x2B)
NAVY2 = RGBColor(0x0C, 0x2C, 0x48)
CYAN = RGBColor(0x1F, 0xA8, 0xB4)
EMBER = RGBColor(0xE3, 0x6B, 0x2C)
SAND = RGBColor(0xF4, 0xF1, 0xEA)
MIST = RGBColor(0xEE, 0xF3, 0xF6)
INK = RGBColor(0x1A, 0x24, 0x33)
GREY = RGBColor(0x5B, 0x67, 0x75)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xD5, 0xDD, 0xE5)


def set_font(run, name="微软雅黑", size=18, bold=False, color=INK):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        rPr.append(rPr.makeelement(qn("a:ea"), {"typeface": name}))
    else:
        ea.set("typeface", name)


def add_rect(slide, left, top, width, height, fill=NAVY):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def add_round(slide, left, top, width, height, fill=CYAN):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    shape.shadow.inherit = False
    # 略收圆角
    try:
        shape.adjustments[0] = 0.08
    except Exception:
        pass
    return shape


def add_text(
    slide,
    left,
    top,
    width,
    height,
    text,
    size=18,
    bold=False,
    color=INK,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, Emu(0))
    if isinstance(text, str):
        text = [text]
    for i, t in enumerate(text):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = t
        set_font(run, size=size, bold=bold, color=color)
    return tb


def add_runs(slide, left, top, width, height, paragraphs, anchor=MSO_ANCHOR.TOP):
    """paragraphs: list of list of (text, size, bold, color) or str."""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, Emu(0))
    for i, para in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        if isinstance(para, str):
            para = [(para, 12, False, INK)]
        for text, size, bold, color in para:
            run = p.add_run()
            run.text = text
            set_font(run, size=size, bold=bold, color=color)
    return tb


def add_bullets(slide, left, top, width, height, items, size=13, color=INK, spacing=1.12):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = spacing
        run = p.add_run()
        run.text = f"·  {it}"
        set_font(run, size=size, color=color)
    return tb


def header(slide, title, subtitle=None, page=None, total=None):
    add_rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.08), fill=CYAN)
    add_rect(slide, Inches(0), Inches(0.08), Inches(13.333), Inches(0.78), fill=NAVY)
    add_text(
        slide,
        Inches(0.45),
        Inches(0.18),
        Inches(10.6),
        Inches(0.32),
        title,
        size=18,
        bold=True,
        color=WHITE,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    if subtitle:
        add_text(
            slide,
            Inches(0.45),
            Inches(0.48),
            Inches(10.6),
            Inches(0.28),
            subtitle,
            size=11,
            color=CYAN,
            anchor=MSO_ANCHOR.TOP,
        )
    add_text(
        slide,
        Inches(10.9),
        Inches(7.18),
        Inches(2.05),
        Inches(0.22),
        f"{page} / {total}" if page and total else "",
        size=10,
        color=GREY,
        align=PP_ALIGN.RIGHT,
    )
    add_text(
        slide,
        Inches(0.45),
        Inches(7.18),
        Inches(8),
        Inches(0.22),
        f"{C.COMPANY}  ·  {C.CONFIDENTIAL}",
        size=9,
        color=GREY,
    )


def make_table(slide, left, top, width, height, headers, rows, font_size=10, col_widths=None):
    shape = slide.shapes.add_table(len(rows) + 1, len(headers), left, top, width, height)
    table = shape.table
    if col_widths:
        for i, w in enumerate(col_widths):
            table.columns[i].width = w
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        tf = cell.text_frame
        tf.clear()
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h
        set_font(r, size=font_size, bold=True, color=WHITE)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = table.cell(ri + 1, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if ri % 2 == 0 else MIST
            tf = cell.text_frame
            tf.clear()
            tf.word_wrap = True
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = str(val)
            set_font(r, size=max(font_size - 1, 9), color=INK)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    n = len(rows) + 1
    row_h = int(height / n)
    for row in table.rows:
        row.height = row_h
    return table


def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    TOTAL = 16
    page = [0]

    def new():
        page[0] += 1
        return prs.slides.add_slide(blank), page[0]

    # 1 封面
    s, _ = new()
    add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), fill=NAVY)
    add_rect(s, Inches(0), Inches(0), Inches(0.18), Inches(7.5), fill=CYAN)
    add_rect(s, Inches(0), Inches(6.85), Inches(13.333), Inches(0.65), fill=NAVY2)
    add_text(
        s,
        Inches(0.7),
        Inches(1.15),
        Inches(11),
        Inches(0.32),
        C.CONFIDENTIAL,
        size=12,
        color=CYAN,
    )
    add_text(s, Inches(0.7), Inches(1.65), Inches(12), Inches(0.55), C.COMPANY, size=20, bold=True, color=WHITE)
    add_text(s, Inches(0.7), Inches(2.25), Inches(12), Inches(0.9), C.DOC_TITLE, size=32, bold=True, color=WHITE)
    add_rect(s, Inches(0.7), Inches(3.28), Inches(2.2), Inches(0.06), fill=EMBER)
    add_text(s, Inches(0.7), Inches(3.55), Inches(11.5), Inches(0.4), C.TAGLINE, size=16, color=CYAN)
    add_text(
        s,
        Inches(0.7),
        Inches(4.2),
        Inches(11.5),
        Inches(1.1),
        [
            "给公司高管的一份入职工作设想：把政府事务做成可调用的地图，",
            "把上海落地做成可决策的比选，而不是再设一套拿地施工班子。",
        ],
        size=15,
        color=WHITE,
    )
    add_text(
        s,
        Inches(0.7),
        Inches(6.98),
        Inches(12),
        Inches(0.38),
        f"提交人  {C.AUTHOR}      {C.DATE_STR}      {C.VERSION}      {C.HORIZON}",
        size=13,
        color=WHITE,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # 2 目录
    s, n = new()
    header(s, "目录", "先对齐岗位边界，再谈 90 天能交出什么", n, TOTAL)
    toc = [
        ("01", "岗位怎么理解", "三张地图，一条边界"),
        ("02", "上海落地判断", "重资产看浦东，杨浦做窗口"),
        ("03", "90 天总目标", "作战图 · 比选包 · 闭环清单"),
        ("04", "三阶段安排", "入局对齐 → 开渠比选 → 交卷闭环"),
        ("05", "部门与风控", "补位拜访，指标翻译，不承包施工"),
        ("06", "节奏与支持", "周报、风险、需要公司打开的门"),
    ]
    for i, (num, title, sub) in enumerate(toc):
        col, row = i % 3, i // 3
        left = Inches(0.45 + col * 4.2)
        top = Inches(1.2 + row * 2.7)
        add_round(s, left, top, Inches(3.95), Inches(2.4), fill=SAND)
        add_rect(s, left, top, Inches(0.1), Inches(2.4), fill=CYAN if row == 0 else EMBER)
        add_text(s, left + Inches(0.3), top + Inches(0.28), Inches(3.4), Inches(0.4), num, size=20, bold=True, color=CYAN)
        add_text(s, left + Inches(0.3), top + Inches(0.85), Inches(3.4), Inches(0.45), title, size=18, bold=True, color=NAVY)
        add_text(s, left + Inches(0.3), top + Inches(1.4), Inches(3.4), Inches(0.7), sub, size=13, color=GREY)

    # 3 岗位理解
    s, n = new()
    header(s, "岗位怎么理解", "三张地图，一条边界——辅助高管，而不是替代合作方", n, TOTAL)
    add_round(s, Inches(0.45), Inches(1.08), Inches(12.45), Inches(0.85), fill=MIST)
    add_text(
        s,
        Inches(0.65),
        Inches(1.18),
        Inches(12.05),
        Inches(0.65),
        C.ROLE_ONE_LINER,
        size=13,
        color=INK,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    accents = [CYAN, EMBER, NAVY2]
    for i, m in enumerate(C.MAPS):
        left = Inches(0.45 + i * 4.2)
        add_round(s, left, Inches(2.1), Inches(4.0), Inches(3.55), fill=SAND)
        add_rect(s, left, Inches(2.1), Inches(4.0), Inches(0.1), fill=accents[i])
        add_text(s, left + Inches(0.22), Inches(2.3), Inches(3.55), Inches(0.35), m["name"], size=16, bold=True, color=NAVY)
        add_text(s, left + Inches(0.22), Inches(2.68), Inches(3.55), Inches(0.4), m["goal"], size=11, color=GREY)
        add_bullets(s, left + Inches(0.22), Inches(3.15), Inches(3.55), Inches(2.3), m["points"], size=12)

    # 4 边界 + 落地判断
    s, n = new()
    header(s, "一条边界，一个判断", "先写清不做什么，再写上海该往哪落", n, TOTAL)
    add_text(s, Inches(0.45), Inches(1.05), Inches(6.1), Inches(0.35), "本岗明确不做什么", size=16, bold=True, color=NAVY)
    add_round(s, Inches(0.45), Inches(1.45), Inches(6.2), Inches(5.4), fill=SAND)
    add_bullets(s, Inches(0.7), Inches(1.65), Inches(5.75), Inches(5.05), C.BOUNDARY, size=13)
    add_text(s, Inches(6.9), Inches(1.05), Inches(6.0), Inches(0.35), "上海落地的基本判断", size=16, bold=True, color=NAVY)
    add_round(s, Inches(6.9), Inches(1.45), Inches(6.0), Inches(2.55), fill=NAVY)
    add_text(s, Inches(7.1), Inches(1.6), Inches(5.6), Inches(2.25), C.LANDING_JUDGEMENT, size=12, color=WHITE)
    add_round(s, Inches(6.9), Inches(4.15), Inches(6.0), Inches(2.7), fill=MIST)
    add_bullets(s, Inches(7.1), Inches(4.3), Inches(5.6), Inches(2.4), C.SITE_PRINCIPLES, size=12)

    # 5 选址角色
    s, n = new()
    header(s, "选址角色分工", "主方案、协同方案、排除与对照一次写死，避免情感选点", n, TOTAL)
    cols = [
        (CYAN, "主方案 · 浦东", "金桥看装置与制造，张江看研发与人才", [
            "新区财政能够托住重资产对话",
            "金桥工业场域与承重更对口大装置",
            "张江科研配套完整，须逐栋核荷载",
            "服务公司已在推进的拿地沟通，做条款翻译",
        ]),
        (EMBER, "协同 · 杨浦", "滨江与复兴岛做窗口，不扛重资产", [
            "十五五将聚变列入区级重点，高校红利在此",
            "滨江平台相对有资金，其他平台负债偏高",
            "复兴岛窗口短、剩余空间有限，轻资产才考虑",
            "适合总部、研发、人才与学术背书",
        ]),
        (NAVY2, "排除 / 对照", "马桥排除，徐汇对照，枢纽观察", [
            "马桥政策浓度高，但同类企业同场互挤",
            "徐汇财政强、楼宇承重弱，排除大装置",
            "东方枢纽作外向协同观察，不纳入主装置",
            "比选表进高管会，不靠口头偏好",
        ]),
    ]
    for i, (color, title, sub, items) in enumerate(cols):
        left = Inches(0.4 + i * 4.3)
        add_round(s, left, Inches(1.12), Inches(4.1), Inches(5.7), fill=SAND)
        add_rect(s, left, Inches(1.12), Inches(4.1), Inches(1.15), fill=color)
        add_text(s, left + Inches(0.22), Inches(1.22), Inches(3.7), Inches(0.4), title, size=18, bold=True, color=WHITE)
        add_text(s, left + Inches(0.22), Inches(1.62), Inches(3.7), Inches(0.5), sub, size=12, color=WHITE)
        add_bullets(s, left + Inches(0.22), Inches(2.45), Inches(3.7), Inches(4.1), items, size=13)

    # 6 90天目标
    s, n = new()
    header(s, "90 天总目标", "可汇报、可决策、可继续往下干的三件套", n, TOTAL)
    add_round(s, Inches(0.45), Inches(1.1), Inches(12.45), Inches(0.85), fill=NAVY)
    add_text(
        s,
        Inches(0.7),
        Inches(1.22),
        Inches(12.0),
        Inches(0.6),
        C.NINETY_GOAL,
        size=15,
        color=WHITE,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    for i, d in enumerate(C.DELIVERABLES):
        left = Inches(0.45 + i * 4.2)
        add_round(s, left, Inches(2.15), Inches(4.0), Inches(4.55), fill=SAND)
        add_text(s, left + Inches(0.25), Inches(2.35), Inches(3.5), Inches(0.35), f"交付 {i + 1}", size=12, bold=True, color=CYAN)
        add_text(s, left + Inches(0.25), Inches(2.75), Inches(3.5), Inches(1.1), d["name"], size=16, bold=True, color=NAVY)
        add_text(s, left + Inches(0.25), Inches(3.95), Inches(3.5), Inches(1.4), d["what"], size=13, color=INK)
        add_text(s, left + Inches(0.25), Inches(5.5), Inches(3.5), Inches(0.8), f"审定：{d['owner']}", size=12, color=GREY)

    # 7 三阶段
    s, n = new()
    header(s, "三阶段总览", f"{C.HORIZON}", n, TOTAL)
    for i, ph in enumerate(C.PHASES):
        top = Inches(1.15 + i * 1.9)
        add_round(s, Inches(0.45), top, Inches(12.45), Inches(1.75), fill=SAND)
        add_rect(s, Inches(0.45), top, Inches(1.7), Inches(1.75), fill=NAVY if i != 1 else CYAN)
        add_text(
            s,
            Inches(0.55),
            top + Inches(0.4),
            Inches(1.5),
            Inches(0.4),
            ph["key"],
            size=22,
            bold=True,
            color=WHITE,
            align=PP_ALIGN.CENTER,
        )
        add_text(
            s,
            Inches(0.55),
            top + Inches(0.9),
            Inches(1.5),
            Inches(0.4),
            ph["span"],
            size=10,
            color=WHITE,
            align=PP_ALIGN.CENTER,
        )
        add_text(s, Inches(2.35), top + Inches(0.15), Inches(4.2), Inches(0.4), ph["name"], size=18, bold=True, color=NAVY)
        add_text(s, Inches(2.35), top + Inches(0.55), Inches(4.2), Inches(1.0), ph["one"], size=13, color=GREY)
        add_bullets(s, Inches(6.7), top + Inches(0.18), Inches(5.95), Inches(1.5), ph["outcomes"][:3], size=12)

    # 8 第一阶段
    s, n = new()
    header(s, "第一阶段  T0–T+30  入局 · 对齐", "先把装置、公司诉求和已有关系搞清楚，再出门", n, TOTAL)
    for i, tr in enumerate(C.PHASE1_TRACKS):
        left = Inches(0.4 + i * 4.3)
        add_round(s, left, Inches(1.12), Inches(4.1), Inches(5.7), fill=SAND)
        add_rect(s, left, Inches(1.12), Inches(0.1), Inches(5.7), fill=CYAN)
        add_text(s, left + Inches(0.28), Inches(1.3), Inches(3.6), Inches(0.45), tr["title"], size=16, bold=True, color=NAVY)
        add_bullets(s, left + Inches(0.28), Inches(1.85), Inches(3.6), Inches(4.7), tr["items"], size=13)

    # 9 第二阶段 政府
    s, n = new()
    header(s, "第二阶段  T+31–T+60  开渠 · 部门", "按地图拜访：补位不抢位，服务已有谈判", n, TOTAL)
    headers = C.PHASE2_GOV[0]
    rows = [r for r in C.PHASE2_GOV[1:]]
    make_table(
        s,
        Inches(0.4),
        Inches(1.15),
        Inches(12.55),
        Inches(5.7),
        headers,
        rows,
        font_size=12,
        col_widths=[Inches(1.2), Inches(2.8), Inches(5.0), Inches(3.55)],
    )

    # 10 第二阶段 现场 + 第三阶段
    s, n = new()
    header(s, "现场六指标与第三阶段交卷", "选址只核事实；90 天结束只留可执行事项", n, TOTAL)
    add_text(s, Inches(0.45), Inches(1.08), Inches(6.2), Inches(0.35), "现场只核六件事", size=16, bold=True, color=NAVY)
    add_round(s, Inches(0.45), Inches(1.48), Inches(6.25), Inches(5.35), fill=SAND)
    add_bullets(s, Inches(0.7), Inches(1.65), Inches(5.8), Inches(5.0), C.PHASE2_SITE, size=13)
    add_text(s, Inches(6.95), Inches(1.08), Inches(6.0), Inches(0.35), "T+61–T+90 交卷", size=16, bold=True, color=NAVY)
    add_round(s, Inches(6.95), Inches(1.48), Inches(5.95), Inches(5.35), fill=MIST)
    add_bullets(s, Inches(7.2), Inches(1.65), Inches(5.5), Inches(5.0), C.PHASE3_ITEMS, size=13)

    # 11 用地
    s, n = new()
    header(s, "用地与对赌：给决策用的风控，不是去谈地", "低地价换长期税收。曲线讲不清，就不要急着拿地", n, TOTAL)
    headers = C.LAND_TYPES[0]
    rows = C.LAND_TYPES[1:]
    make_table(
        s,
        Inches(0.4),
        Inches(1.12),
        Inches(12.55),
        Inches(2.7),
        headers,
        rows,
        font_size=12,
        col_widths=[Inches(2.6), Inches(3.3), Inches(3.4), Inches(3.25)],
    )
    add_round(s, Inches(0.4), Inches(4.05), Inches(12.55), Inches(2.75), fill=SAND)
    add_text(s, Inches(0.65), Inches(4.2), Inches(12), Inches(0.35), "谈判时必须提前说清的三句话", size=14, bold=True, color=EMBER)
    add_bullets(s, Inches(0.65), Inches(4.6), Inches(12.05), Inches(2.0), C.TAX_NOTES, size=14)

    # 12 政策包
    s, n = new()
    header(s, "政策包：写成路径，不写成收入", "90 天能认路、备材料；不能承诺名单和金额", n, TOTAL)
    headers = C.POLICY_PACK[0]
    rows = C.POLICY_PACK[1:]
    make_table(
        s,
        Inches(0.4),
        Inches(1.15),
        Inches(12.55),
        Inches(5.7),
        headers,
        rows,
        font_size=12,
        col_widths=[Inches(1.8), Inches(3.3), Inches(3.8), Inches(3.65)],
    )

    # 13 融资 + 节奏
    s, n = new()
    header(s, "融资协同与周节奏", "融资是公司级事项；本岗只做背书材料与机构过滤", n, TOTAL)
    add_round(s, Inches(0.4), Inches(1.1), Inches(12.55), Inches(1.35), fill=NAVY)
    add_text(s, Inches(0.65), Inches(1.22), Inches(12.1), Inches(1.1), C.FINANCE_NOTE, size=13, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_round(s, Inches(0.4), Inches(2.6), Inches(6.2), Inches(4.15), fill=SAND)
    add_text(s, Inches(0.6), Inches(2.75), Inches(5.8), Inches(0.35), "本岗只做三件事", size=15, bold=True, color=NAVY)
    add_bullets(s, Inches(0.6), Inches(3.2), Inches(5.8), Inches(3.3), C.FINANCE_ACTIONS, size=13)
    add_text(s, Inches(6.85), Inches(2.6), Inches(6.0), Inches(0.35), "固定节奏", size=15, bold=True, color=NAVY)
    make_table(
        s,
        Inches(6.85),
        Inches(3.0),
        Inches(6.05),
        Inches(3.75),
        C.CADENCE[0],
        C.CADENCE[1:],
        font_size=11,
        col_widths=[Inches(1.5), Inches(3.0), Inches(1.55)],
    )

    # 14 产出物
    s, n = new()
    header(s, "90 天产出物时间表", "每个阶段都有纸面结果，避免“一直在跑、没有交卷”", n, TOTAL)
    make_table(
        s,
        Inches(0.4),
        Inches(1.15),
        Inches(12.55),
        Inches(5.7),
        C.OUTPUT_LIST[0],
        C.OUTPUT_LIST[1:],
        font_size=13,
        col_widths=[Inches(1.8), Inches(6.3), Inches(4.45)],
    )

    # 15 风险与支持
    s, n = new()
    header(s, "风险、边界、需要公司打开的门", "岗位价值取决于补位是否干净、口径是否经得起问", n, TOTAL)
    make_table(
        s,
        Inches(0.35),
        Inches(1.08),
        Inches(12.65),
        Inches(3.35),
        C.RISKS[0],
        C.RISKS[1:],
        font_size=10,
        col_widths=[Inches(2.5), Inches(4.3), Inches(5.85)],
    )
    add_text(s, Inches(0.45), Inches(4.52), Inches(12), Inches(0.3), "需要公司支持的五件事", size=14, bold=True, color=NAVY)
    add_round(s, Inches(0.35), Inches(4.85), Inches(12.65), Inches(1.95), fill=MIST)
    add_bullets(s, Inches(0.55), Inches(4.95), Inches(12.25), Inches(1.75), C.SUPPORT_NEEDED, size=12)

    # 16 结束
    s, _ = new()
    add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), fill=NAVY)
    add_rect(s, Inches(0), Inches(0), Inches(0.18), Inches(7.5), fill=EMBER)
    add_text(s, Inches(0.7), Inches(1.5), Inches(12), Inches(0.4), "90 天结束时，高管手里应有三样东西", size=16, color=CYAN)
    add_text(s, Inches(0.7), Inches(2.05), Inches(12), Inches(0.7), "地图、比选、下一季度只留可执行事项", size=26, bold=True, color=WHITE)
    items = [
        "谁已对接、谁该补、谁不要再去——写在作战图上",
        "重资产主方案在浦东，杨浦做窗口——写在比选表上",
        "1–2 条可启动的政策路径，15–20 家机构短名单——写在闭环清单上",
    ]
    add_bullets(s, Inches(0.7), Inches(3.05), Inches(11.5), Inches(2.0), items, size=16, color=WHITE)
    add_text(
        s,
        Inches(0.7),
        Inches(5.4),
        Inches(11.5),
        Inches(0.8),
        "先懂装置，再出门；补位不抢位；政策不口头化。\n这是本岗对东昇聚变最有用的 90 天。",
        size=16,
        color=CYAN,
    )
    add_text(
        s,
        Inches(0.7),
        Inches(6.7),
        Inches(12),
        Inches(0.35),
        f"{C.AUTHOR}  ·  {C.DATE_STR}  ·  {C.VERSION}  ·  {C.CONFIDENTIAL}",
        size=12,
        color=WHITE,
    )

    prs.save(OUT_FILE)
    print(f"已生成 {OUT_FILE}  共 {len(prs.slides)} 页")
    return OUT_FILE


if __name__ == "__main__":
    build()
