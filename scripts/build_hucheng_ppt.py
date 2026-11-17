#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成《沪蓉科技产业协同发展合作建议书》讨论稿 PPT。

口径可直接发给成都民营产业园集团管理层。
载体面积沿用 2026 年 9 月 24 日交流口径，正式签约前需再核对。
"""

from __future__ import annotations

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

FONT = "Microsoft YaHei"
W, H = 13.333, 7.5
TOTAL = 16

NAVY = "10243A"
NAVY_2 = "17344C"
INK = "1C2833"
MUTED = "5C6B76"
TEAL = "1B6B7A"
TEAL_SOFT = "E6F3F1"
GOLD = "8C6A2F"
GOLD_SOFT = "F8F1DE"
GOLD_ON_NAVY = "E4C98A"
CREAM = "F6F3EC"
WHITE = "FFFFFF"
CARD = "FFFCF8"
LINE = "E4DACC"
CINNABAR = "8E3B34"
CINNABAR_SOFT = "F8EEEE"
NAVY_TEXT = "E7EEF2"
NAVY_MUTED = "C5D0D8"


def rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def font_run(run, size, color, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    run.font.name = FONT
    rpr = run._r.get_or_add_rPr()
    rpr.set("lang", "zh-CN")
    rpr.set("altLang", "en-US")
    for tag in ("a:latin", "a:ea", "a:cs"):
        node = rpr.find(qn(tag))
        if node is None:
            node = etree.SubElement(rpr, qn(tag))
        node.set("typeface", FONT)


def rect(slide, x, y, w, h, fill, line=None, rounded=False, adj=0.08):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = rgb(line)
        shape.line.width = Pt(1)
    shape.shadow.inherit = False
    if rounded:
        try:
            shape.adjustments[0] = adj
        except Exception:
            pass
    return shape


def paint(slide, color):
    shape = rect(slide, 0, 0, W, H, color)
    tree = slide.shapes._spTree
    element = shape._element
    tree.remove(element)
    tree.insert(2, element)


def write(slide, paragraphs, x, y, w, h, anchor="t"):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.word_wrap = True
    frame.auto_size = None
    frame.margin_left = Emu(0)
    frame.margin_right = Emu(0)
    frame.margin_top = Emu(0)
    frame.margin_bottom = Emu(0)
    frame._txBody.bodyPr.set("anchor", {"t": "t", "m": "ctr", "b": "b"}[anchor])
    align_map = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
    for index, para in enumerate(paragraphs):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.alignment = align_map[para.get("align", "left")]
        paragraph.space_before = Pt(para.get("before", 0))
        paragraph.space_after = Pt(para.get("after", 0))
        paragraph.line_spacing = para.get("spacing", 1.05)
        run = paragraph.add_run()
        run.text = para["text"]
        font_run(run, para.get("size", 14), para.get("color", INK), para.get("bold", False))
    return box


def t(text, size=14, color=INK, bold=False, align="left", before=0, after=0, spacing=1.05):
    return {
        "text": text,
        "size": size,
        "color": color,
        "bold": bold,
        "align": align,
        "before": before,
        "after": after,
        "spacing": spacing,
    }


def blank(prs, color=CREAM):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    paint(slide, color)
    return slide


def footer(slide, page):
    rect(slide, 0.52, 7.02, 12.3, 0.01, LINE)
    write(
        slide,
        [t("上海市杨浦区科技企业联合会  ·  战略合作建议（讨论稿）", 11, MUTED)],
        0.52,
        7.1,
        9.4,
        0.28,
    )
    write(
        slide,
        [t(f"{page:02d}  /  {TOTAL:02d}", 11, MUTED, align="right")],
        10.2,
        7.1,
        2.6,
        0.28,
    )


def header(slide, kicker, title):
    rect(slide, 0, 0, W, 0.06, TEAL)
    write(slide, [t(kicker, 12, GOLD, True)], 0.52, 0.2, 12.2, 0.26)
    write(slide, [t(title, 26, NAVY, True, spacing=1.0)], 0.5, 0.46, 12.3, 0.46)
    rect(slide, 0.52, 1.02, 0.78, 0.035, GOLD)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def card(slide, x, y, w, h, fill=CARD, line=LINE):
    return rect(slide, x, y, w, h, fill, line, rounded=True, adj=0.06)


def slide_cover(prs):
    slide = blank(prs, NAVY)
    rect(slide, 0, 0, 0.12, H, GOLD)
    rect(slide, 8.55, 0, 4.783, H, NAVY_2)
    rect(slide, 8.55, 0, 0.015, H, GOLD)

    write(slide, [t("讨论稿", 14, GOLD_ON_NAVY, True)], 0.62, 0.38, 3, 0.3)
    write(slide, [t("2026年11月", 14, NAVY_MUTED)], 6.3, 0.38, 1.9, 0.3, anchor="m")
    write(
        slide,
        [t("战略合作建议书", 36, WHITE, True, spacing=1.0)],
        0.6,
        1.15,
        7.6,
        0.7,
    )
    write(
        slide,
        [
            t("共建沪蓉科技产业协同发展", 20, NAVY_TEXT, spacing=1.05),
            t("与创新企业服务平台", 20, NAVY_TEXT, before=2, spacing=1.05),
        ],
        0.62,
        1.95,
        7.5,
        0.85,
    )
    rect(slide, 0.64, 3.0, 1.35, 0.035, GOLD)
    write(slide, [t("建议发起单位", 12, GOLD_ON_NAVY, True)], 0.62, 3.25, 6.5, 0.26)
    write(
        slide,
        [t("上海市杨浦区科技企业联合会", 18, WHITE, True)],
        0.62,
        3.52,
        7.5,
        0.38,
    )
    write(slide, [t("拟合作单位", 12, GOLD_ON_NAVY, True)], 0.62, 4.08, 6.5, 0.26)
    write(
        slide,
        [t("成都民营产业园集团", 18, WHITE, True)],
        0.62,
        4.35,
        7.5,
        0.38,
    )
    write(
        slide,
        [t("具体签约主体以双方确认为准", 13, NAVY_MUTED)],
        0.62,
        4.78,
        7.4,
        0.28,
    )
    write(
        slide,
        [t("拟合作周期    2027 — 2029年", 15, NAVY_TEXT, True)],
        0.62,
        5.4,
        7.2,
        0.34,
    )
    write(
        slide,
        [t("依据 2026年9月24日双方交流整理", 13, NAVY_MUTED)],
        0.62,
        5.82,
        7.2,
        0.28,
    )
    write(
        slide,
        [t("本文件供战略合作交流，不构成法律约束力协议。", 12, NAVY_MUTED)],
        0.62,
        6.85,
        7.5,
        0.3,
    )

    write(slide, [t("五个合作方向", 14, GOLD_ON_NAVY, True)], 8.9, 0.42, 4.0, 0.32)
    rect(slide, 8.92, 0.84, 0.7, 0.03, GOLD)
    items = [
        ("01", "上海产业招商", "以上海园区连接杨浦企业与成都载体"),
        ("02", "人工智能产业集聚", "人工智能、机器人与硬科技企业服务"),
        ("03", "科技金融赋能", "路演、资本对接与企业成长辅导"),
        ("04", "沪蓉双城联动", "上海研发创新，成都转化与拓展"),
        ("05", "国际化出海", "新加坡、东南亚与企业出海服务"),
    ]
    for index, (num, title, body) in enumerate(items):
        y = 1.15 + index * 1.05
        write(slide, [t(num, 14, GOLD_ON_NAVY, True)], 8.9, y, 0.7, 0.28)
        write(slide, [t(title, 16, WHITE, True)], 9.65, y - 0.02, 2.9, 0.32)
        write(slide, [t(body, 12, NAVY_MUTED, spacing=1.05)], 9.65, y + 0.32, 3.2, 0.5)
    notes(
        slide,
        "开场说明这是讨论稿，签约主体尚未最终确认。五个方向是全文主线，"
        "上海园区是试点，成都载体是合作放大的方向。",
    )


def slide_background(prs):
    slide = blank(prs)
    header(slide, "01    合作背景", "从一次交流，走向长期合作机制")
    footer(slide, 2)
    stats = [
        ("约80万㎡", "成都产业载体", "以成都高新区为核心，覆盖温江、新津，并延伸至上海与海外孵化平台。"),
        ("约6000㎡", "上海合作窗口", "上海项目拥有自有载体，并已形成一定规模的科技企业集群，适合作为首批试点。"),
        ("2026.09.24", "已有交流基础", "双方已就招商、企业服务、川渝资源、国际化与园区联合运营交换意见。"),
    ]
    for index, (num, title, body) in enumerate(stats):
        x = 0.5 + index * 4.22
        card(slide, x, 1.28, 4.05, 2.28)
        rect(slide, x, 1.28, 4.05, 0.06, TEAL if index < 2 else GOLD)
        write(slide, [t(num, 22, NAVY, True)], x + 0.22, 1.46, 3.6, 0.42)
        write(slide, [t(title, 15, TEAL, True)], x + 0.22, 1.96, 3.6, 0.3)
        write(slide, [t(body, 13, INK, spacing=1.12)], x + 0.22, 2.36, 3.6, 1.0)
    card(slide, 0.5, 3.72, 12.33, 1.62)
    write(slide, [t("联合会可以接上的资源", 15, TEAL, True)], 0.74, 3.88, 6, 0.3)
    write(
        slide,
        [
            t(
                "上海市杨浦区科技企业联合会长期服务区域科技企业，重点围绕人工智能、科技金融、企业国际化、产业招商及创新创业开展工作。依托《杨浦一百》研究成果、产业活动和创新企业资源库，持续对接资本、园区、高校及产业链。",
                14,
                INK,
                spacing=1.15,
            )
        ],
        0.74,
        4.26,
        11.85,
        0.9,
    )
    card(slide, 0.5, 5.5, 12.33, 1.35, TEAL_SOFT, "C9DDD8")
    write(
        slide,
        [t("合作原则    资源共享  ·  产业协同  ·  市场化运作  ·  优势互补  ·  长期共赢", 15, TEAL, True)],
        0.74,
        5.66,
        11.8,
        0.34,
    )
    write(
        slide,
        [
            t(
                "建议这次合作围绕上海产业招商、人工智能产业集聚、科技金融赋能、沪蓉双城联动、国际化出海五个方向，建立可持续的合作机制。",
                14,
                INK,
                spacing=1.1,
            )
        ],
        0.74,
        6.08,
        11.85,
        0.58,
    )
    notes(slide, "面积为交流口径，正式对外承诺前请集团确认范围。原则页不要展开成招商话术。")


def slide_strengths(prs):
    slide = blank(prs)
    header(slide, "02    资源互补", "双方各自拿出最能互补的资源")
    footer(slide, 3)
    columns = [
        (
            TEAL,
            "成都民营产业园集团",
            [
                ("产业载体", "成都约80万平方米，上海约6000平方米自有空间。"),
                ("孵化运营", "长期从事科技企业孵化与产业园区运营。"),
                ("网络布局", "成都高新区、温江、新津、上海及海外孵化平台。"),
                ("企业基础", "上海项目已集聚一批科技企业，可作为服务样本。"),
                ("国际化", "具备海外孵化经验，以及新加坡等业务布局基础。"),
            ],
        ),
        (
            GOLD,
            "杨浦区科技企业联合会",
            [
                ("科技企业", "区域企业服务积累，以及《杨浦一百》研究成果。"),
                ("产业方向", "人工智能、科技金融、国际化、招商与创新创业。"),
                ("资本对接", "投资机构、专业服务机构与融资路演组织能力。"),
                ("活动品牌", "产业活动的内容策划、嘉宾邀请和企业组织。"),
                ("创新资源", "高校、科技成果转化与产业链资源的对接通道。"),
            ],
        ),
    ]
    for index, (accent, title, items) in enumerate(columns):
        x = 0.5 + index * 6.4
        card(slide, x, 1.26, 6.2, 5.55)
        rect(slide, x, 1.26, 6.2, 0.62, NAVY)
        write(slide, [t(title, 16, WHITE, True)], x + 0.24, 1.4, 5.7, 0.36)
        for row, (name, body) in enumerate(items):
            y = 2.08 + row * 0.9
            rect(slide, x + 0.24, y, 0.08, 0.62, accent)
            write(slide, [t(name, 15, NAVY, True)], x + 0.48, y - 0.02, 5.4, 0.3)
            write(slide, [t(body, 13, INK, spacing=1.05)], x + 0.48, y + 0.3, 5.45, 0.48)
    notes(slide, "这一页只讲资源，不讲费用。让对方看到联合会带来的是企业、资本和活动，载体仍由集团掌握。")


def slide_directions(prs):
    slide = blank(prs)
    header(slide, "03    合作方向", "五个方向，构成长期合作机制")
    footer(slide, 4)
    items = [
        ("01", "上海产业招商", "把集团上海园区建成连接复旦周边创新资源、杨浦科技企业和成都载体的服务节点。"),
        ("02", "人工智能产业集聚", "重点引入人工智能、机器人、先进制造、集成电路和科技服务企业，提升出租率与入驻质量。"),
        ("03", "科技金融赋能", "为入驻和储备企业提供融资路演、投资机构对接、投后服务与成长辅导。"),
        ("04", "沪蓉双城联动", "促进上海科技成果、创新企业和产业资本，与成都产业空间、产业链需求精准匹配。"),
        ("05", "国际化出海", "面向新加坡、东南亚及其他重点市场，探索企业国际化服务和跨境创新合作。"),
    ]
    positions = [(0, 0), (1, 0), (2, 0), (0, 1), (1, 1)]
    widths = [4.02, 4.02, 4.02, 6.12, 6.12]
    xs = [0.5, 4.66, 8.82, 0.5, 6.76]
    ys = [1.26, 1.26, 1.26, 3.72, 3.72]
    for (num, title, body), w, x, y in zip(items, widths, xs, ys):
        card(slide, x, y, w, 2.3)
        rect(slide, x, y, 0.08, 2.3, TEAL)
        write(slide, [t(num, 14, GOLD, True)], x + 0.26, y + 0.16, 1.2, 0.28)
        write(slide, [t(title, 18, NAVY, True)], x + 0.26, y + 0.5, w - 0.46, 0.4)
        write(slide, [t(body, 13, INK, spacing=1.12)], x + 0.26, y + 1.05, w - 0.48, 1.05)
    notes(slide, "五个方向是建议书的结构，不是五个同时铺开的项目。后文按这个顺序展开。")


def slide_leasing(prs):
    slide = blank(prs)
    header(slide, "04    产业招商", "以上海园区为试点，联合开展产业招商")
    footer(slide, 5)
    columns = [
        (
            "01",
            "共享招商资源",
            [
                "集团提供可招商面积、租赁条件、空间产品和适配行业。",
                "联合会结合《杨浦一百》，筛选扩租、搬迁、研发办公和区域布局需求。",
                "定期组织企业考察、项目路演、专题推介和现场带看。",
                "重点跟踪被投企业、成果转化企业和成长型科技企业。",
            ],
        ),
        (
            "02",
            "聚焦三个方向",
            [
                "人工智能与具身智能产业。",
                "硬科技、集成电路与先进制造。",
                "科技服务与产业国际化业务。",
                "按企业阶段探索灵活租赁、成长型优惠和产业服务组合。",
            ],
        ),
        (
            "03",
            "市场化合作",
            [
                "可按项目采用资源推荐、招商代理或专题活动招商。",
                "专项服务另行约定内容、考核标准和费用结算。",
                "招商服务费与会员会费分开约定。",
                "成交认定和结算方式，以专项协议为准。",
            ],
        ),
    ]
    for index, (num, title, lines) in enumerate(columns):
        x = 0.5 + index * 4.22
        card(slide, x, 1.26, 4.05, 5.52)
        write(slide, [t(num, 14, GOLD, True)], x + 0.22, 1.44, 3.5, 0.28)
        write(slide, [t(title, 18, NAVY, True)], x + 0.22, 1.78, 3.6, 0.4)
        rect(slide, x + 0.24, 2.3, 0.7, 0.03, GOLD)
        body = [t("·  " + line, 13, INK, before=0 if i == 0 else 8, spacing=1.12) for i, line in enumerate(lines)]
        write(slide, body, x + 0.22, 2.52, 3.62, 3.9)
    notes(slide, "强调上海是试点，专项招商另行计费。不要在这一页承诺出租率数字。")


def slide_platform(prs):
    slide = blank(prs)
    header(slide, "05    创新平台", "共建人工智能与硬科技创新服务平台")
    footer(slide, 6)
    card(slide, 0.5, 1.24, 12.33, 1.18, TEAL_SOFT, "C9DDD8")
    write(slide, [t("建议名称", 12, TEAL, True)], 0.74, 1.38, 2, 0.24)
    write(
        slide,
        [t("沪蓉人工智能与硬科技产业协同创新中心", 20, NAVY, True)],
        0.74,
        1.64,
        11.8,
        0.4,
    )
    write(
        slide,
        [t("平台名称、设立形式及对外挂牌，由双方另行协商确认。", 13, MUTED)],
        0.74,
        2.08,
        11.5,
        0.24,
    )
    items = [
        ("01", "联合筛选人工智能、机器人、集成电路、可控核聚变等领域的创新项目。"),
        ("02", "围绕研发合作、产业应用、科技成果转化和企业融资开展专题交流。"),
        ("03", "邀请高校科研团队、科技企业及投资机构参与产业对接。"),
        ("04", "推动园区现有企业与联合会服务企业开展产业链协作。"),
        ("05", "探索开发者交流、技术展示、产业沙龙和创新项目路演。"),
    ]
    for index, (num, text) in enumerate(items):
        y = 2.58 + index * 0.72
        card(slide, 0.5, y, 12.33, 0.64)
        write(slide, [t(num, 14, GOLD, True, align="center")], 0.66, y + 0.14, 0.6, 0.34)
        write(slide, [t(text, 15, INK)], 1.4, y + 0.14, 11.1, 0.36)
    card(slide, 0.5, 6.22, 12.33, 0.68, TEAL_SOFT, "C9DDD8")
    write(
        slide,
        [t("具备条件的项目，可协调集团成都总部孵化平台，提供进一步的产业空间及配套服务。", 14, TEAL, True)],
        0.74,
        6.38,
        11.85,
        0.36,
    )
    notes(slide, "名称是建议，不要说已经同意挂牌。可控核聚变按对方交流中的产业方向保留，作为筛选范围。")


def slide_finance(prs):
    slide = blank(prs)
    header(slide, "06    科技金融", "为科技企业建立常态化融资服务")
    footer(slide, 7)
    columns = [
        (
            "01",
            "融资对接机制",
            "联合会协助引入创业投资基金、产业投资基金、银行、证券机构和专业服务机构。",
            ["融资咨询", "路演展示", "投资机构对接"],
        ),
        (
            "02",
            "季度主题路演",
            "优先组织有明确融资需求、较好技术基础或商业化进展的项目。",
            ["人工智能与具身智能", "硬科技与集成电路", "新能源及未来产业", "专精特新与先进制造"],
        ),
        (
            "03",
            "孵化与资本协同",
            "对优质企业，探索由空间、辅导、资本对接和产业资源组成的综合培育。",
            ["园区空间支持", "创业成长辅导", "产业资源服务"],
        ),
    ]
    for index, (num, title, lead, lines) in enumerate(columns):
        x = 0.5 + index * 4.22
        card(slide, x, 1.26, 4.05, 4.28)
        write(slide, [t(num, 14, GOLD, True)], x + 0.22, 1.42, 3.5, 0.26)
        write(slide, [t(title, 18, NAVY, True)], x + 0.22, 1.74, 3.6, 0.4)
        write(slide, [t(lead, 13, INK, spacing=1.12)], x + 0.22, 2.26, 3.62, 1.15)
        body = [t("·  " + line, 14, TEAL, True, before=0 if i == 0 else 6) for i, line in enumerate(lines)]
        write(slide, body, x + 0.22, 3.5, 3.62, 1.8)
    card(slide, 0.5, 5.7, 12.33, 1.12, GOLD_SOFT, "E6D7B4")
    write(slide, [t("合规边界", 14, GOLD, True)], 0.74, 5.86, 2, 0.28)
    write(
        slide,
        [
            t(
                "具体基金投资、投融资收益和投资管理，由具备相应资质的市场主体依法独立开展。联合会与集团负责组织对接，不替代持牌机构作投资决策。",
                14,
                INK,
                spacing=1.1,
            )
        ],
        0.74,
        6.2,
        11.85,
        0.5,
    )
    notes(slide, "科技金融页必须讲清合规边界，避免被理解成联合会或集团直接募资、代销或承诺收益。")


def slide_dual_city(prs):
    slide = blank(prs)
    header(slide, "07    沪蓉双城", "建立上海创新与成都转化的双向通道")
    footer(slide, 8)
    roles = [
        ("上海方面", "组织对接", "科技创新资源、科研成果、科技企业及投资机构。"),
        ("成都方面", "在地支撑", "产业空间、当地产业链、园区运营，以及企业投资考察。"),
    ]
    for index, (place, tag, body) in enumerate(roles):
        x = 0.5 + index * 6.4
        card(slide, x, 1.24, 6.2, 1.42)
        write(slide, [t(place, 16, NAVY, True)], x + 0.24, 1.38, 3.2, 0.32)
        write(slide, [t(tag, 13, TEAL, True, align="right")], x + 3.3, 1.4, 2.65, 0.28)
        write(slide, [t(body, 14, INK)], x + 0.24, 1.84, 5.72, 0.58)
    flows = [
        ("01", "上海企业西进", "支持成长型科技企业赴成都设立研发、生产或区域业务机构。"),
        ("02", "成都企业东进", "支持优质科技企业来上海开展研发合作、市场拓展和资本对接。"),
        ("03", "企业之间协作", "推动两地企业开展产品应用、供应链合作和产业项目落地。"),
    ]
    for index, (num, title, body) in enumerate(flows):
        x = 0.5 + index * 4.22
        card(slide, x, 2.82, 4.05, 1.95)
        write(slide, [t(num, 13, GOLD, True)], x + 0.22, 2.98, 3.5, 0.24)
        write(slide, [t(title, 16, NAVY, True)], x + 0.22, 3.28, 3.6, 0.34)
        write(slide, [t(body, 13, INK, spacing=1.12)], x + 0.22, 3.72, 3.62, 0.85)
    card(slide, 0.5, 4.94, 12.33, 1.88, NAVY)
    write(slide, [t("每年一次", 13, GOLD_ON_NAVY, True)], 0.76, 5.12, 3, 0.26)
    write(slide, [t("上海 — 成都科技企业产业考察", 18, WHITE, True)], 0.76, 5.42, 8, 0.38)
    write(
        slide,
        [t("围绕具体产业需求，安排园区参访、企业交流和项目洽谈。", 14, NAVY_TEXT)],
        0.76,
        5.88,
        11.5,
        0.3,
    )
    write(
        slide,
        [t("协同模式：上海研发创新，成都产业转化与市场拓展。", 15, GOLD_ON_NAVY, True)],
        0.76,
        6.26,
        11.5,
        0.34,
    )
    notes(slide, "双城是长期价值所在。首年先完成需求摸排和至少一次考察，不承诺当年落地面积。")


def slide_global(prs):
    slide = blank(prs)
    header(slide, "08    国际化出海", "结合海外孵化经验，服务企业走向海外")
    footer(slide, 9)
    columns = [
        (
            "01",
            "新加坡市场合作",
            "围绕国际市场拓展、海外商务对接、创新项目孵化和跨境商业合作，建立专项业务交流机制。",
        ),
        (
            "02",
            "企业出海资源库",
            "联合专业机构，对接海外市场信息、商务渠道、法律税务咨询、知识产权保护及跨境业务服务。",
        ),
        (
            "03",
            "国际化专题活动",
            "联合举办出海沙龙、东南亚市场交流和国际合作项目对接。条件成熟后，再探索其他海外创新平台。",
        ),
    ]
    for index, (num, title, body) in enumerate(columns):
        x = 0.5 + index * 4.22
        card(slide, x, 1.28, 4.05, 4.15)
        rect(slide, x, 1.28, 4.05, 0.06, TEAL)
        write(slide, [t(num, 14, GOLD, True)], x + 0.24, 1.52, 3.5, 0.28)
        write(slide, [t(title, 18, NAVY, True)], x + 0.24, 1.95, 3.58, 0.7)
        write(slide, [t(body, 14, INK, spacing=1.18)], x + 0.24, 2.8, 3.58, 2.2)
    card(slide, 0.5, 5.6, 12.33, 1.22, TEAL_SOFT, "C9DDD8")
    write(
        slide,
        [
            t("合作起点", 14, TEAL, True),
            t(
                "集团提供现有海外孵化经验和新加坡业务布局，联合会提供国际化业务渠道和专业机构对接。先做交流机制和资源库，再决定是否延伸到新的海外平台。",
                14,
                INK,
                before=4,
                spacing=1.1,
            ),
        ],
        0.74,
        5.76,
        11.85,
        0.95,
    )
    notes(slide, "海外部分保持克制。只写新加坡、东南亚和现有海外孵化经验，不新增未经交流确认的城市。")


def slide_events(prs):
    slide = blank(prs)
    header(slide, "09    产业活动", "用联合活动把企业、资本和场景组织到一起")
    footer(slide, 10)
    card(slide, 0.5, 1.24, 4.15, 4.22)
    write(slide, [t("场地支持", 18, NAVY, True)], 0.74, 1.44, 3.7, 0.4)
    rect(slide, 0.76, 1.96, 0.7, 0.03, GOLD)
    write(
        slide,
        [
            t(
                "按双方前期交流意向，集团可为联合会举办适宜规模的科技企业活动提供场地支持。",
                14,
                INK,
                spacing=1.15,
            )
        ],
        0.74,
        2.2,
        3.68,
        1.35,
    )
    write(
        slide,
        [
            t("另行协商", 13, TEAL, True),
            t("场次、容量、设施及使用安排，在首月对接时确认。", 13, INK, before=4, spacing=1.1),
        ],
        0.74,
        3.7,
        3.68,
        1.3,
    )
    events = [
        "杨浦人工智能与硬科技创新沙龙",
        "沪蓉科技企业产业协同交流会",
        "科技企业投融资与产业资本对接会",
        "科技成果转化与企业成长专题沙龙",
        "科技企业出海与国际市场拓展交流会",
    ]
    card(slide, 4.82, 1.24, 8.01, 4.22)
    write(slide, [t("建议共同策划的活动品牌", 18, NAVY, True)], 5.06, 1.44, 7.5, 0.38)
    for index, name in enumerate(events):
        y = 2.02 + index * 0.64
        rect(slide, 5.06, y, 0.46, 0.46, TEAL_SOFT)
        write(slide, [t(f"{index + 1:02d}", 12, TEAL, True, align="center")], 5.06, y + 0.08, 0.46, 0.3)
        write(slide, [t(name, 15, INK)], 5.68, y + 0.06, 6.8, 0.36)
    card(slide, 0.5, 5.62, 12.33, 1.2, GOLD_SOFT, "E6D7B4")
    write(slide, [t("《杨浦一百》企业服务", 15, GOLD, True)], 0.74, 5.76, 6, 0.28)
    write(
        slide,
        [
            t(
                "结合研究成果，开展园区优秀企业走访、创新案例交流和产业研究。企业宣传与资源对接，在双方授权且企业同意后进行。",
                14,
                INK,
                spacing=1.1,
            )
        ],
        0.74,
        6.12,
        11.85,
        0.55,
    )
    notes(slide, "活动品牌是建议菜单。首年目标是 4 到 6 场，不必五类各做一遍。场地条件仍待确认。")


def slide_roles(prs):
    slide = blank(prs)
    header(slide, "10    职责分工", "每件事都写明谁组织、谁承接")
    footer(slide, 11)
    headers = ["合作事项", "杨浦区科技企业联合会", "成都民营产业园集团"]
    rows = [
        ("园区招商", "企业筛选、项目推荐、招商活动", "空间提供、租赁商务、落地服务"),
        ("人工智能合作", "产业资源、创新项目、活动组织", "企业孵化、产业场景、载体支持"),
        ("科技金融", "投资机构、融资路演、专业服务对接", "企业需求组织、项目推荐"),
        ("沪蓉产业协同", "上海企业和科研创新资源", "成都园区、产业链及在地服务"),
        ("企业国际化", "国际合作渠道及专业机构对接", "海外平台及相关业务资源对接"),
        ("品牌活动", "内容策划、嘉宾邀请、企业组织", "场地及必要配套支持"),
    ]
    widths = [2.2, 4.95, 4.95]
    xs = [0.52]
    for width in widths[:-1]:
        xs.append(xs[-1] + width + 0.08)
    header_h = 0.48
    row_h = 0.74
    y = 1.28
    for x, width, text in zip(xs, widths, headers):
        rect(slide, x, y, width, header_h, NAVY, rounded=True, adj=0.08)
        write(slide, [t(text, 14, WHITE, True, align="center")], x, y + 0.08, width, 0.32)
    for index, row in enumerate(rows):
        yy = y + header_h + 0.08 + index * (row_h + 0.06)
        fill = CARD if index % 2 == 0 else "F7F4EC"
        for col, (x, width, text) in enumerate(zip(xs, widths, row)):
            rect(slide, x, yy, width, row_h, fill, LINE, rounded=True, adj=0.08)
            color = NAVY if col == 0 else INK
            write(
                slide,
                [t(text, 14, color, col == 0, align="center" if col == 0 else "left")],
                x + (0.08 if col else 0.1),
                yy + 0.18,
                width - 0.2,
                0.4,
            )
    notes(slide, "分工表是框架。具体项目仍以双方业务负责人和专项协议为准。")


def slide_phases(prs):
    slide = blank(prs)
    header(slide, "11    推进节奏", "十二个月内，从试点走到可复制机制")
    footer(slide, 12)
    phases = [
        (
            "第一阶段",
            "资源梳理与启动",
            "首月",
            [
                "确定业务负责人，建立常态对接。",
                "梳理上海可招商空间与租赁政策。",
                "明确首批重点服务企业名单。",
                "选定首场产业活动。",
                "对接成都总部产业合作负责人。",
                "确认重点项目与合作优先级。",
            ],
        ),
        (
            "第二阶段",
            "试点与活动落地",
            "第2—3个月",
            [
                "以上海园区作为首批试点载体。",
                "举办人工智能或科技金融主题活动。",
                "建立招商线索与项目进展台账。",
                "组织实地看场与商务洽谈。",
                "完成首批融资企业与机构匹配。",
                "启动沪蓉产业合作需求摸排。",
            ],
        ),
        (
            "第三阶段",
            "深化与机制复制",
            "第4—12个月",
            [
                "招商、孵化、金融转入常态推进。",
                "完成至少一次沪蓉考察或交流。",
                "形成一项可持续的协同或出海项目。",
                "条件成熟后延伸到成都总部园区。",
                "复盘首年成果，再确定量化指标。",
                "重大项目另行论证，不在试点中捆绑。",
            ],
        ),
    ]
    for index, (phase, title, when, lines) in enumerate(phases):
        x = 0.5 + index * 4.22
        card(slide, x, 1.24, 4.05, 5.55)
        rect(slide, x, 1.24, 4.05, 1.18, NAVY)
        write(slide, [t(phase, 13, GOLD_ON_NAVY, True)], x + 0.2, 1.34, 3.65, 0.24)
        write(slide, [t(title, 16, WHITE, True)], x + 0.2, 1.6, 3.65, 0.34)
        write(slide, [t(when, 13, NAVY_TEXT)], x + 0.2, 1.98, 3.65, 0.26)
        body = [t(line, 13, INK, before=0 if i == 0 else 7, spacing=1.05) for i, line in enumerate(lines)]
        write(slide, body, x + 0.22, 2.6, 3.62, 3.9)
    notes(slide, "节奏按首月、2到3个月、4到12个月。成都总部在第三阶段扩展，不作为首月交付。")


def slide_commercial(prs):
    slide = blank(prs)
    header(slide, "12    商业机制", "五种合作方式并行，三类费用分开")
    footer(slide, 13)
    card(slide, 0.5, 1.22, 12.33, 1.08, GOLD_SOFT, "E6D7B4")
    write(slide, [t("费用边界", 14, GOLD, True)], 0.72, 1.34, 2.2, 0.26)
    write(
        slide,
        [
            t(
                "会费、招商服务费、活动合作费分别约定。副会长单位年费用于建立正式合作关系，不包含后续招商及专业服务。",
                14,
                INK,
            )
        ],
        0.72,
        1.68,
        11.9,
        0.46,
    )
    rows = [
        ("01", "战略合作", "可先签署框架协议，明确资源共享、产业协作和联合活动的总体方向。"),
        ("02", "会员合作", "建议集团上海项目申请副会长单位。参考会费每年人民币1万元，以正式会员管理规定和审批程序为准。"),
        ("03", "专项招商", "企业租赁、项目引进及其他商业合作，另签协议，约定工作范围、成果认定、收费标准和结算方式。"),
        ("04", "产业活动", "优先使用双方现有资源。第三方赞助、项目服务费及其他商业安排，按活动方案和独立协议执行。"),
        ("05", "重大项目", "园区改造、载体运营、跨区域投资及基金合作，另行组织论证和商务洽谈。"),
    ]
    for index, (num, title, body) in enumerate(rows):
        y = 2.46 + index * 0.86
        card(slide, 0.5, y, 12.33, 0.78)
        rect(slide, 0.5, y, 0.08, 0.78, TEAL if index != 1 else GOLD)
        write(slide, [t(num, 14, GOLD, True)], 0.76, y + 0.2, 0.55, 0.36)
        write(slide, [t(title, 15, NAVY, True)], 1.4, y + 0.18, 1.7, 0.4)
        write(slide, [t(body, 13, INK, spacing=1.05)], 3.2, y + 0.12, 9.35, 0.54)
    notes(slide, "1万元是参考会费和关系入口。招商代理、活动赞助和基金合作都要另约，不能写进会员权益。")


def slide_outcomes(prs):
    slide = blank(prs)
    header(slide, "13    首年成果", "首年重点争取六项可见成果")
    footer(slide, 14)
    items = [
        ("01", "跑通联合工作机制", "联合招商和企业服务有固定负责人、对接节奏和项目台账。"),
        ("02", "完成4—6场产业活动", "围绕人工智能、硬科技和科技金融，形成持续的企业与机构流量。"),
        ("03", "建立一批项目储备", "储备有明确空间、融资或跨区域布局需求的科技企业项目。"),
        ("04", "形成上海招商成果", "在特色方向上推动上海园区获得可核对的企业入驻或租赁进展。"),
        ("05", "完成一次沪蓉交流", "至少组织一次上海与成都的科技企业产业交流或实地考察。"),
        ("06", "形成一个持续项目", "探索一项可以继续运营的双城产业协同或国际化服务项目。"),
    ]
    for index, (num, title, body) in enumerate(items):
        col = index % 2
        row = index // 2
        x = 0.5 + col * 6.4
        y = 1.24 + row * 1.72
        card(slide, x, y, 6.2, 1.58)
        write(slide, [t(num, 16, GOLD, True)], x + 0.22, y + 0.18, 0.7, 0.32)
        write(slide, [t(title, 16, NAVY, True)], x + 1.0, y + 0.18, 4.9, 0.34)
        write(slide, [t(body, 13, INK, spacing=1.08)], x + 0.22, y + 0.66, 5.76, 0.72)
    write(
        slide,
        [
            t(
                "以上六项为建议性工作目标。具体任务和量化指标，在合作启动后结合双方资源确定；本稿不预设成交面积或融资金额。",
                13,
                MUTED,
            )
        ],
        0.55,
        6.48,
        12.2,
        0.42,
    )
    notes(slide, "六项都是建议性目标。出租面积、融资金额等量化指标，启动后再共同确定，本稿不预设数字。")


def slide_priorities(prs):
    slide = blank(prs)
    header(slide, "14    优先抓手", "建议先做三件事，并看向成都载体")
    footer(slide, 15)
    rows = [
        (TEAL, "第一优先", "副会长单位入会 + 园区联合招商", "尽快建立正式合作关系，并形成可核对的业务成果。"),
        (GOLD, "第二优先", "人工智能产业活动 + 科技金融路演", "把企业和投资机构组织到园区，形成持续的项目来源。"),
        (CINNABAR, "第三优先", "成都约80万平方米载体的沪蓉招商", "由上海试点延伸到集团层面，服务长三角科技企业与成都产业空间的长期协作。"),
    ]
    for index, (color, rank, title, body) in enumerate(rows):
        y = 1.24 + index * 1.42
        card(slide, 0.5, y, 12.33, 1.3)
        rect(slide, 0.5, y, 1.85, 1.3, color, rounded=True, adj=0.06)
        write(slide, [t(rank, 15, WHITE, True, align="center")], 0.58, y + 0.44, 1.7, 0.4)
        write(slide, [t(title, 18, NAVY, True)], 2.58, y + 0.22, 9.9, 0.4)
        write(slide, [t(body, 14, INK)], 2.58, y + 0.7, 9.9, 0.4)
    card(slide, 0.5, 5.56, 12.33, 1.28, NAVY)
    write(slide, [t("建议的合作尺度", 13, GOLD_ON_NAVY, True)], 0.74, 5.7, 4, 0.26)
    write(
        slide,
        [
            t(
                "上海约6000平方米园区适合作为首批试点。合作价值的放大，在于同步对接成都约80万平方米产业载体，使双方形成面向长三角科技企业的沪蓉招商协作。",
                14,
                WHITE,
                spacing=1.1,
            )
        ],
        0.74,
        6.04,
        11.85,
        0.64,
    )
    notes(
        slide,
        "三个抓手按易落地程度排序。讲解时说明：第一、第二优先负责把关系做成，第三优先决定合作是否具备集团层面的商业价值。",
    )


def slide_close(prs):
    slide = blank(prs)
    header(slide, "15    下一步", "先确认主体和权限，再启动首月事项")
    card(slide, 0.5, 1.22, 6.2, 3.72)
    write(slide, [t("首月建议完成", 16, NAVY, True)], 0.74, 1.38, 5.6, 0.34)
    start_items = [
        "双方确定业务负责人，建立常态对接。",
        "完成上海园区可招商空间及租赁政策梳理。",
        "明确首批重点服务的科技企业名单。",
        "筛选适宜举办的首场产业活动。",
        "对接集团成都总部产业合作负责人。",
        "确认重点项目及合作优先级。",
    ]
    write(
        slide,
        [t("·  " + item, 13, INK, before=0 if i == 0 else 5, spacing=1.02) for i, item in enumerate(start_items)],
        0.74,
        1.86,
        5.7,
        2.85,
    )
    card(slide, 6.86, 1.22, 5.97, 3.72, GOLD_SOFT, "E6D7B4")
    write(slide, [t("正式文本前请确认", 16, NAVY, True)], 7.1, 1.38, 5.5, 0.34)
    checks = [
        "集团的准确工商主体名称。",
        "上海项目的签约主体与权限。",
        "成都总部的签约或授权权限。",
        "约80万㎡、约6000㎡的统计范围。",
        "副会长申请主体及当年会费标准。",
    ]
    write(
        slide,
        [t("·  " + item, 13, INK, before=0 if i == 0 else 6, spacing=1.02) for i, item in enumerate(checks)],
        7.1,
        1.86,
        5.5,
        2.8,
    )
    card(slide, 0.5, 5.08, 12.33, 0.72, TEAL_SOFT, "C9DDD8")
    write(
        slide,
        [
            t(
                "建议以集团上海园区为首批试点，以具体项目形成可落地成果，再逐步深化集团层面的战略协作。",
                14,
                TEAL,
                True,
            )
        ],
        0.72,
        5.26,
        11.9,
        0.4,
    )
    rect(slide, 0, 6.0, W, 1.5, NAVY)
    write(slide, [t("上海市杨浦区科技企业联合会", 16, WHITE, True)], 0.55, 6.16, 8, 0.32)
    write(slide, [t("2026年11月", 14, GOLD_ON_NAVY)], 10.2, 6.18, 2.6, 0.3, anchor="m")
    write(
        slide,
        [
            t(
                "本建议书为双方战略合作交流讨论文件，不构成具有法律约束力的合作协议。具体合作主体、项目内容及商业条款，以双方后续正式确认文件为准。",
                12,
                NAVY_TEXT,
                spacing=1.08,
            )
        ],
        0.55,
        6.58,
        11.3,
        0.72,
    )
    write(slide, [t("16  /  16", 12, NAVY_MUTED, align="right")], 11.5, 6.95, 1.4, 0.28)
    notes(slide, "结束时把讨论稿性质说清楚，并请对方指定上海项目和成都总部的对接人。")


def audit(prs):
    issues = []
    for index, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if shape.left is None or shape.top is None:
                continue
            right = shape.left + shape.width
            bottom = shape.top + shape.height
            if shape.left < -Emu(Inches(0.02)) or right > prs.slide_width + Emu(Inches(0.02)):
                issues.append(f"第{index}页横向越界")
            if shape.top < -Emu(Inches(0.02)) or bottom > prs.slide_height + Emu(Inches(0.02)):
                issues.append(f"第{index}页纵向越界")
        text = "\n".join(shape.text_frame.text for shape in slide.shapes if shape.has_text_frame)
        if len(text.strip()) < 8:
            issues.append(f"第{index}页文字过少")
    if len(prs.slides) != TOTAL:
        issues.append(f"页数{len(prs.slides)}与标注{TOTAL}不一致")
    return issues


def build(path):
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    prs.core_properties.title = "上海市杨浦区科技企业联合会与成都民营产业园集团战略合作建议书"
    prs.core_properties.subject = "共建沪蓉科技产业协同发展与创新企业服务平台（讨论稿）"
    prs.core_properties.author = "上海市杨浦区科技企业联合会"
    prs.core_properties.category = "讨论稿"
    slide_cover(prs)
    slide_background(prs)
    slide_strengths(prs)
    slide_directions(prs)
    slide_leasing(prs)
    slide_platform(prs)
    slide_finance(prs)
    slide_dual_city(prs)
    slide_global(prs)
    slide_events(prs)
    slide_roles(prs)
    slide_phases(prs)
    slide_commercial(prs)
    slide_outcomes(prs)
    slide_priorities(prs)
    slide_close(prs)
    issues = audit(prs)
    if issues:
        raise SystemExit("版式检查未通过：\n" + "\n".join(issues))
    prs.save(path)
    print(f"已生成 {path} ，共 {len(prs.slides)} 页")


if __name__ == "__main__":
    build("deliverables/杨浦科技企业联合会_成都民营产业园集团_战略合作建议书_讨论稿.pptx")
