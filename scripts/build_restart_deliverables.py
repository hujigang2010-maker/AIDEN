#!/usr/bin/env python3
"""生成《银发青年·重启第二人生》合作落地 PPT 与 Word。

综合三份材料：10月13日洽谈、内部落地指引，以及《银发青年_重启第二人生_项目落地策划方案_2026-10》。
对方为陈红苗，运营主理为胡总。价格与预算均为测算假设。
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from pptx import Presentation
from pptx.dml.color import RGBColor as PColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn as pqn
from pptx.util import Emu, Inches, Pt as PPt

FONT = "WenQuanYi Micro Hei"
OUT_DIR = Path("/workspace/docs")
PPT_NAME = "银发青年_重启第二人生_项目落地策划方案_2026-10.pptx"
DOC_NAME = "银发青年_重启第二人生_项目落地策划方案_2026-10.docx"

INK = PColor(0x1C, 0x24, 0x2B)
PAPER = PColor(0xF3, 0xEE, 0xE4)
WHITE = PColor(0xFF, 0xFC, 0xF8)
CINA = PColor(0x8C, 0x34, 0x2C)
BRONZE = PColor(0x8A, 0x6A, 0x3E)
SAGE = PColor(0x2F, 0x55, 0x4C)
MUTED = PColor(0x6A, 0x61, 0x56)
LINE = PColor(0xE3, 0xD8, 0xC6)
SOFT = PColor(0xF7, 0xF1, 0xE6)
CREAM = PColor(0xFB, 0xF7, 0xF1)

SW = Inches(13.333)
SH = Inches(7.5)


def style_run(run, size, color, bold=False):
    run.font.size = PPt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    rpr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        node = rpr.find(pqn(f"a:{tag}"))
        if node is None:
            node = rpr.makeelement(pqn(f"a:{tag}"), {})
            rpr.append(node)
        node.set("typeface", FONT)


def blank(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = PAPER
    return slide


def rect(slide, l, t, w, h, fill, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = PPt(1)
    # 去掉圆角默认阴影带来的突兀
    sppr = shape._element.spPr
    effect = sppr.find(pqn("a:effectLst"))
    if effect is not None:
        sppr.remove(effect)
    return shape


def box(slide, l, t, w, h, lines, size=15, color=INK, bold=False, align="left", anchor="top", spacing=6):
    shape = slide.shapes.add_textbox(l, t, w, h)
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}[anchor]
    align_map = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
    if isinstance(lines, str):
        lines = [lines]
    normalized = []
    for line in lines:
        normalized.extend(str(line).split("\n"))
    lines = normalized
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align_map[align]
        p.space_before = PPt(0)
        p.space_after = PPt(spacing if i < len(lines) - 1 else 0)
        run = p.add_run()
        run.text = "" if line is None else str(line)
        style_run(run, size, color, bold)
    return shape


def rich(slide, l, t, w, h, paragraphs, anchor="top"):
    """paragraphs: list of list of (text, size, color, bold)."""
    shape = slide.shapes.add_textbox(l, t, w, h)
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE}[anchor]
    for i, runs in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = PPt(0)
        p.space_after = PPt(4)
        for text, size, color, bold in runs:
            run = p.add_run()
            run.text = text
            style_run(run, size, color, bold)
    return shape


def footer(slide, page, total=17):
    rect(slide, Inches(0.55), Inches(7.18), Inches(12.25), Inches(0.01), LINE)
    box(slide, Inches(0.55), Inches(7.22), Inches(8.5), Inches(0.24),
        "重启第二人生  ·  呈陈红苗  ·  测算假设，待双方确认", 10, MUTED)
    box(slide, Inches(11.4), Inches(7.22), Inches(1.4), Inches(0.24),
        f"{page}  /  {total}", 10, MUTED, align="right")


def kicker(slide, text, l=0.55, t=0.28):
    box(slide, Inches(l), Inches(t), Inches(8), Inches(0.28), text, 12, CINA, True)


def title(slide, text, t=0.52, size=30):
    box(slide, Inches(0.55), Inches(t), Inches(12.2), Inches(0.58), text, size, INK, True)


def card(slide, l, t, w, h, accent=CINA):
    rect(slide, l, t, w, h, WHITE, LINE)
    rect(slide, l, t, Inches(0.07), h, accent)


def build_ppt(path: Path):
    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH
    prs.core_properties.title = "银发青年·重启第二人生项目落地策划方案"
    prs.core_properties.author = "胡总团队"
    prs.core_properties.subject = "呈陈红苗  2026年10月"
    prs.core_properties.category = "合作落地稿"

    cover(prs)
    principle(prs)
    two_lines(prs)
    money(prs)
    first_steps(prs)
    tea(prs)
    care(prs)
    trip(prs)
    roles(prs)
    thirty(prs)
    exchange(prs)
    community(prs)
    numbers(prs)
    ninety(prs)
    guard(prs)
    confirm(prs)
    this_week(prs)

    prs.save(path)


def cover(prs):
    s = blank(prs)
    rect(s, 0, 0, SW, SH, INK)
    rect(s, 0, 0, Inches(0.16), SH, CINA)
    box(s, Inches(0.75), Inches(0.48), Inches(10), Inches(0.3),
        "合作落地稿   ·   2026年10月19日   ·   依据10月13日洽谈", 13, BRONZE)
    box(s, Inches(0.75), Inches(1.7), Inches(11), Inches(0.4), "银发青年", 20, BRONZE, True)
    box(s, Inches(0.72), Inches(2.15), Inches(12), Inches(1.15), "重启第二人生", 54, WHITE, True)
    box(s, Inches(0.75), Inches(3.45), Inches(10), Inches(0.45), "项目落地策划方案", 24, PAPER)
    rect(s, Inches(0.75), Inches(4.15), Inches(2.1), Inches(0.045), CINA)
    box(s, Inches(0.75), Inches(4.4), Inches(11.2), Inches(0.85),
        "先做出两场付费会客厅和第一批陪诊、陪聊订单。\n公司、股权和小程序，等真实付款出现后再谈。",
        18, PColor(0xD9, 0xD0, 0xC4))
    box(s, Inches(0.75), Inches(6.45), Inches(6), Inches(0.35), "呈    陈红苗", 16, WHITE, True)
    box(s, Inches(8.2), Inches(6.45), Inches(4.4), Inches(0.35), "胡总团队整理", 16, PAPER, align="right")


def principle(prs):
    s = blank(prs)
    kicker(s, "01    这次一起做成的事")
    title(s, "钱要落到产品上，陪伴要让人愿意再约")
    # 三张引用
    quotes = [
        ("01", "钱从哪里来", "每一笔都要写清谁付钱、买到什么、什么时候到账。"),
        ("02", "真陪伴", "客人能指出这次独有的一件事。说不出，这单不算做成。"),
        ("03", "25岁与65岁", "青年按单获得收入，长者获得在场的人，并可以留下一段经验。"),
    ]
    for i, (n, h, b) in enumerate(quotes):
        x = Inches(0.55 + i * 4.2)
        card(s, x, Inches(1.55), Inches(3.95), Inches(2.55), CINA if i == 0 else (SAGE if i == 1 else BRONZE))
        box(s, x + Inches(0.28), Inches(1.75), Inches(3.4), Inches(0.35), n, 14, CINA if i == 0 else (SAGE if i == 1 else BRONZE), True)
        box(s, x + Inches(0.28), Inches(2.15), Inches(3.4), Inches(0.45), h, 22, INK, True)
        box(s, x + Inches(0.28), Inches(2.75), Inches(3.4), Inches(1.05), b, 15, MUTED, spacing=2)
    rect(s, Inches(0.55), Inches(4.35), Inches(12.25), Inches(2.55), WHITE, LINE)
    box(s, Inches(0.85), Inches(4.55), Inches(11.6), Inches(0.35), "总原则", 14, CINA, True)
    box(s, Inches(0.85), Inches(5.0), Inches(11.6), Inches(1.6),
        [
            "一个品牌，两条分开算账的业务。前台是文化会客厅，另一条是范围写死的陪诊与陪聊。",
            "1.8万私域用来组织年轻人供给和活动执行，不写成现成的银发客户，也不自动并入新公司。",
            "陈红苗主责客人、品位和场地。胡总主责供给、机构、交付和台账。每周对一次账。",
        ],
        16, INK, spacing=8)
    footer(s, 2)


def two_lines(prs):
    s = blank(prs)
    kicker(s, "02    品牌结构")
    title(s, "一个品牌，两条业务，分开算账")
    cols = [
        (CINA, "A线", "第二人生会客厅", [
            "高端文化生活与同龄社交",
            "客人：50–75岁，有支付能力，看重话题和人",
            "场景：茶室微缩、思想会客厅、主理人轮值",
            "收入：门票，之后才是系列与会员",
            "第一优先：陈红苗的圈子是否愿意付钱",
        ]),
        (SAGE, "B线", "品质陪伴", [
            "非医疗陪诊、公共场所陪聊、数字技能辅导",
            "付款人：长者本人，或成年子女",
            "陪诊由审核过的服务商承接",
            "陪聊由培训过的年轻人承接",
            "并行试单：先有安全的订单，再谈规模",
        ]),
    ]
    for i, (accent, tag, name, items) in enumerate(cols):
        x = Inches(0.55 + i * 6.4)
        card(s, x, Inches(1.5), Inches(6.1), Inches(4.35), accent)
        box(s, x + Inches(0.32), Inches(1.7), Inches(5.5), Inches(0.3), tag, 13, accent, True)
        box(s, x + Inches(0.32), Inches(2.05), Inches(5.5), Inches(0.45), name, 24, INK, True)
        box(s, x + Inches(0.32), Inches(2.7), Inches(5.5), Inches(2.8),
            [f"·  {t}" for t in items], 15, INK, spacing=10)
    box(s, Inches(0.55), Inches(6.05), Inches(12.2), Inches(0.85),
        "两条线共用品牌和回访，不共用一张海报。会客厅不宣传陪诊钟点价；陪诊单上不借用高校、联合会或海上讲堂的名义。",
        14, MUTED)
    footer(s, 3)


def money(prs):
    s = blank(prs)
    kicker(s, "03    钱从哪里来")
    title(s, "三条变现路径，按这个顺序收第一笔")
    paths = [
        ("01", "会客厅票款", "现在就卖", "长者本人", "茶室 8–12 人 × 299 元", "付款满 8 人就办"),
        ("02", "陪诊与陪聊", "审核通过就卖", "本人或子女", "陪诊 399 元 / 4 小时；谈话陪同测 800 元", "11 月开始试单"),
        ("03", "冬季微度假", "有名单再卖", "参团客人", "意向金 1,000 元，可退", "12 月满 8 人再交给旅行社"),
    ]
    for i, (n, name, when, who, price, go) in enumerate(paths):
        x = Inches(0.55 + i * 4.2)
        card(s, x, Inches(1.5), Inches(3.95), Inches(4.55), CINA if i < 2 else BRONZE)
        box(s, x + Inches(0.25), Inches(1.68), Inches(3.45), Inches(0.3), n, 13, CINA if i < 2 else BRONZE, True)
        box(s, x + Inches(0.25), Inches(2.05), Inches(3.45), Inches(0.7), name, 22, INK, True)
        box(s, x + Inches(0.25), Inches(2.8), Inches(3.45), Inches(0.35), when, 13, SAGE, True)
        rich(s, x + Inches(0.25), Inches(3.3), Inches(3.45), Inches(2.4), [
            [("谁付钱  ", 13, MUTED, False), (who, 14, INK, True)],
            [("第一笔", 13, MUTED, False)],
            [(price, 15, INK, True)],
            [("启动", 13, MUTED, False)],
            [(go, 15, INK, False)],
        ])
    footer(s, 4)


def first_steps(prs):
    s = blank(prs)
    kicker(s, "04    第一步")
    title(s, "四步准备，按闸门依次打开")
    steps = [
        ("10月20日–26日", "把口头约定写成一页备忘录", "陈红苗给出茶室两个日期和 10 位愿意见面的人。胡总指定调度、发出供给报名、搭好台账。收款主体没定，谁也不收款。"),
        ("11月第一周", "茶室微缩测试", "陈红苗当一次客人，完成 40 分钟模拟陪聊。通过后，8–12 人微缩场按 299 元预售。满 8 人付款才办。"),
        ("11月同步", "陪诊、陪聊试单", "陪诊只派给审核过的服务商。陪聊只派给培训过的 8 人名单。每周付费单不超过 4 单。"),
        ("12月", "第二场售票，短途只登记", "约 20 人的会客厅按 399 元测算。短途满 8 位可退意向金，才进入旅行社合同。不满就停在方案。"),
    ]
    for i, (when, head, body) in enumerate(steps):
        y = Inches(1.48 + i * 1.35)
        rect(s, Inches(0.55), y, Inches(0.08), Inches(1.15), CINA if i < 3 else BRONZE)
        box(s, Inches(0.85), y, Inches(3.1), Inches(1.15), when, 15, CINA, True, anchor="middle")
        box(s, Inches(4.0), y, Inches(8.6), Inches(0.38), head, 18, INK, True)
        box(s, Inches(4.0), y + Inches(0.4), Inches(8.6), Inches(0.75), body, 13, MUTED)
    footer(s, 5)


def tea(prs):
    s = blank(prs)
    kicker(s, "05    茶室微缩测试")
    title(s, "第一场就在茶室，用付款检验圈子")
    left = [
        ("规模", "8–12 人。陈红苗亲自邀请，胡总不把名单拿去群发。"),
        ("价格", "体验价 299 元/人。这是测算价，访谈后可以改，已收款不补差。"),
        ("时长", "120 分钟：茶叙、一个问题、两人深谈、收集下一笔预付。"),
        ("模拟", "开工前 40 分钟。陈红苗当客人。若觉得被对付，就换人重来一次。"),
    ]
    for i, (h, b) in enumerate(left):
        y = Inches(1.5 + i * 1.25)
        card(s, Inches(0.55), y, Inches(7.55), Inches(1.12), SAGE)
        box(s, Inches(0.85), y + Inches(0.14), Inches(6.9), Inches(0.32), h, 14, SAGE, True)
        box(s, Inches(0.85), y + Inches(0.48), Inches(6.9), Inches(0.5), b, 14, INK)
    card(s, Inches(8.3), Inches(1.5), Inches(4.5), Inches(5.0), CINA)
    box(s, Inches(8.6), Inches(1.72), Inches(4.0), Inches(0.3), "通过标准", 14, CINA, True)
    box(s, Inches(8.6), Inches(2.2), Inches(4.0), Inches(3.9),
        [
            "7 天内有 4 位客人预付下一场，或预付一单陪同。",
            "现场气氛好、没有预付，这场记为未通过。",
            "两次模拟都未通过，微缩场延期，不改为免费场。",
            "名牌只写名字。没有赞助，不介绍保险和课程。",
        ],
        15, INK, spacing=12)
    footer(s, 6)


def care(prs):
    s = blank(prs)
    kicker(s, "06    第一批订单")
    title(s, "陪诊走服务商，陪聊走培训名单")
    # two columns
    blocks = [
        (SAGE, "办事陪诊", [
            "4 小时流程陪同，测算价 399 元",
            "挂号费、车费、药费客人自付",
            "只做排队、路线、医嘱原文、取药",
            "合格服务商签约后才能派单",
            "胡总主责尽调、派单和回访",
        ]),
        (CINA, "谈话与学习陪同", [
            "公共场所，一次只做一件事",
            "高端档先按 2 小时 800 元访谈",
            "同时询问 480 元和 1,200 元",
            "陪伴者来自录取的 8 人，周结约 70%",
            "陈红苗可介绍客人，不负责排班",
        ]),
    ]
    for i, (accent, name, items) in enumerate(blocks):
        x = Inches(0.55 + i * 6.4)
        card(s, x, Inches(1.48), Inches(6.15), Inches(3.55), accent)
        box(s, x + Inches(0.3), Inches(1.65), Inches(5.6), Inches(0.4), name, 20, INK, True)
        box(s, x + Inches(0.3), Inches(2.2), Inches(5.6), Inches(2.55),
            [f"·  {t}" for t in items], 15, INK, spacing=8)
    rect(s, Inches(0.55), Inches(5.2), Inches(12.25), Inches(1.7), INK)
    box(s, Inches(0.85), Inches(5.38), Inches(11.7), Inches(0.3), "每一单出门前都要有的四样东西", 14, BRONZE, True)
    box(s, Inches(0.85), Inches(5.8), Inches(11.7), Inches(0.85),
        "书面四问   ·   预付款到共同可见账户   ·   当日意外险保单号   ·   两小时内一页纪要\n给药、助浴、翻身、陪夜、解释病情、推销保险，出现任何一项即停单。",
        15, WHITE, spacing=6)
    footer(s, 7)


def trip(prs):
    s = blank(prs)
    kicker(s, "07    冬季过渡")
    title(s, "短途微度假等名单，不抢 11 月的第一笔")
    rules = [
        ("目的地", "高铁大约两小时，优先苏州、嘉兴、杭州。两天内返回。"),
        ("成团", "8 位客人各自预付可退意向金 1,000 元。出发前七日可退。"),
        ("承办", "交通与住宿由持证旅行社签约收款。我们负责内容和陪伴。"),
        ("行程", "每天留白。不进购物店，不办保险说明会，不办速配。"),
    ]
    for i, (h, b) in enumerate(rules):
        y = Inches(1.5 + (i % 4) * 1.15)
        card(s, Inches(0.55), y, Inches(8.35), Inches(1.02), BRONZE)
        box(s, Inches(0.85), y + Inches(0.12), Inches(1.5), Inches(0.75), h, 16, BRONZE, True, anchor="middle")
        box(s, Inches(2.4), y + Inches(0.18), Inches(6.2), Inches(0.7), b, 14, INK, anchor="middle")
    card(s, Inches(9.1), Inches(1.5), Inches(3.7), Inches(4.6), CINA)
    box(s, Inches(9.35), Inches(1.72), Inches(3.25), Inches(0.7), "今冬先不走", 22, INK, True)
    box(s, Inches(9.35), Inches(2.7), Inches(3.25), Inches(3.0),
        ["敦煌、云南、厦门、青岛\n留到 2027 年春季单独立项。", "", "西藏和高海拔不排。", "", "领队没有书面档期之前，\n对外不宣布人名。"],
        14, INK, spacing=2)
    footer(s, 8)


def roles(prs):
    s = blank(prs)
    kicker(s, "08    分工")
    title(s, "陈红苗守客人和品位，胡总守供给和交付")
    headers = ["事项", "陈红苗", "胡总"]
    rows = [
        ["客人与访谈", "主责。20 人候选，自己联络", "提供问卷和记录表"],
        ["茶室与场地", "主责。档期、费用、品牌授权", "准备流程和备选场地报价"],
        ["内容品位", "否决虚的场、推销和凑数", "按确认后的流程执行"],
        ["年轻供给", "参与标准，可试一次模拟", "主责招募、培训、8 人名单"],
        ["机构与陪诊商", "参与高端机构那一次", "主责 8 家拜访和服务商尽调"],
        ["活动执行回访", "出席微缩场和首场复盘", "主责现场、纪要、24 小时回访"],
        ["账与合同", "确认主体、审批支出", "起草备忘录，每周出台账"],
        ["公司与股权", "验证后共同谈", "验证后共同谈"],
    ]
    table_shape = s.shapes.add_table(1 + len(rows), 3, Inches(0.55), Inches(1.45), Inches(12.25), Inches(5.45))
    table = table_shape.table
    table.columns[0].width = Inches(2.5)
    table.columns[1].width = Inches(4.9)
    table.columns[2].width = Inches(4.85)
    data = [headers] + rows
    for r, row in enumerate(data):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = val
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.1)
            cell.margin_right = Inches(0.08)
            fill = INK if r == 0 else (SOFT if r % 2 == 0 else WHITE)
            cell.fill.solid()
            cell.fill.fore_color.rgb = fill
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT
                for run in p.runs:
                    style_run(run, 13, WHITE if r == 0 else INK, r == 0 or c == 0)
    footer(s, 9)


def thirty(prs):
    s = blank(prs)
    kicker(s, "09    三十天交付")
    title(s, "10月26日以前，各自先交得出来的东西")
    card(s, Inches(0.55), Inches(1.5), Inches(6.15), Inches(5.35), CINA)
    box(s, Inches(0.85), Inches(1.7), Inches(5.6), Inches(0.4), "陈红苗", 22, INK, True)
    box(s, Inches(0.85), Inches(2.25), Inches(5.55), Inches(4.3),
        [
            "·  茶室两个可办日期，写清可坐人数和费用",
            "·  20 位候选人的联络计划，电话留在自己这里",
            "·  其中 10 位答应在 11 月 2 日前深访",
            "·  海上讲堂商用条件：收费、档期、名称",
            "·  回复文末确认事项，尤其是体验价和收款主体",
        ],
        15, INK, spacing=12)
    card(s, Inches(6.9), Inches(1.5), Inches(5.9), Inches(5.35), SAGE)
    box(s, Inches(7.2), Inches(1.7), Inches(5.35), Inches(0.4), "胡总", 22, INK, True)
    box(s, Inches(7.2), Inches(2.25), Inches(5.3), Inches(4.3),
        [
            "·  指定 1 名调度，并定下自己每周那个半天",
            "·  私域只发一次供给报名，限额 20，录取 8",
            "·  8 家机构与服务商的拜访排期",
            "·  六张表：访谈、机构、服务商、订单、收支、复购",
            "·  两页备忘录草稿、保险样张购买路径",
        ],
        15, INK, spacing=12)
    footer(s, 10)


def exchange(prs):
    s = blank(prs)
    kicker(s, "10    25岁 × 65岁")
    title(s, "双向发生在每一单的结尾，不发生在口号里")
    layers = [
        ("会上十分钟", "叫做「交换一件小事」。青年请教一段经历，长者请教一个操作。可以拒绝，工钱照结。"),
        ("八对共创", "公共场所，90 分钟。交付故事卡，以及一条已解决的数字问题。由会客厅客人付费。"),
        ("上门另立产品", "一对一上门只走陪聊或陪诊规则：预付、保险、纪要、停单。不把共创活动直接变成进门服务。"),
    ]
    for i, (h, b) in enumerate(layers):
        y = Inches(1.5 + i * 1.55)
        rect(s, Inches(0.55), y, Inches(12.25), Inches(1.4), WHITE, LINE)
        rect(s, Inches(0.55), y, Inches(0.08), Inches(1.4), SAGE)
        box(s, Inches(0.9), y + Inches(0.18), Inches(2.6), Inches(1.0), h, 18, INK, True, anchor="middle")
        box(s, Inches(3.6), y + Inches(0.22), Inches(8.8), Inches(1.0), b, 15, MUTED, anchor="middle")
    footer(s, 11)


def community(prs):
    s = blank(prs)
    kicker(s, "11    私域信任")
    title(s, "1.8万人是供给壁垒，不是银发客户池")
    points = [
        ("画像已经对过", "群内以三十岁上下为主。前十个群更活跃，也已满员。陈红苗所在的群里，同龄人是少数。"),
        ("信任怎么用", "持续运营多年，所以招得到愿意按规则出门的年轻人，也比闲鱼更容易让熟人放心。"),
        ("一次怎么发", "只在胡总亲自维护的高活跃群发一次。文案只写：上海、按单、先面试、做陪同和记录，不做护理。"),
        ("产权怎么留", "社群、白皮书、联合会和高校关联保持独立授权。未书面同意，不批量营销，不折成股权。"),
    ]
    for i, (h, b) in enumerate(points):
        x = Inches(0.55 + (i % 2) * 6.4)
        y = Inches(1.5 + (i // 2) * 2.55)
        card(s, x, y, Inches(6.15), Inches(2.35), BRONZE if i % 2 else CINA)
        box(s, x + Inches(0.3), y + Inches(0.25), Inches(5.6), Inches(0.45), h, 18, INK, True)
        box(s, x + Inches(0.3), y + Inches(0.85), Inches(5.6), Inches(1.2), b, 14, MUTED)
    footer(s, 12)


def numbers(prs):
    s = blank(prs)
    kicker(s, "12    单位经济  ·  全部是测算假设")
    title(s, "先看清一笔账，再决定能不能加场")
    headers = ["测算", "收入", "直接成本", "贡献结余"]
    rows = [
        ["会客厅 20人 × 399元", "7,980", "5,900", "2,080（约26%）"],
        ["陪诊 20单 × 399元", "7,980", "6,700", "1,280（约16%）"],
        ["60天：2场 + 20单", "23,940", "18,500", "5,440"],
    ]
    table_shape = s.shapes.add_table(4, 4, Inches(0.55), Inches(1.5), Inches(12.25), Inches(2.35))
    table = table_shape.table
    widths = [Inches(4.3), Inches(2.4), Inches(2.5), Inches(3.05)]
    for i, w in enumerate(widths):
        table.columns[i].width = w
    data = [headers] + rows
    for r, row in enumerate(data):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = val
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.1)
            cell.fill.solid()
            cell.fill.fore_color.rgb = INK if r == 0 else (SOFT if r == 3 else WHITE)
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    style_run(run, 14, WHITE if r == 0 else INK, r in (0, 3))
    notes = [
        ("会客厅成本里有什么", ["场地 1,200，主持 1,800。", "茶歇 800，执行 1,000。", "物料 500，沟通 600。", "茶室零租金单独记账。"]),
        ("陪诊成本里有什么", ["服务商 280，保险准备金 15。", "排单回访 30，杂费 10。", "16% 是贡献率。", "尚未扣除固定费用和税费。"]),
        ("这张表证明什么", ["两个月成交约 2.4 万元。", "贡献结余约 5,440 元。", "固定成本和税费还没扣。", "用来验证，不用来宣布盈利。"]),
    ]
    for i, (h, b) in enumerate(notes):
        x = Inches(0.55 + i * 4.2)
        box(s, x, Inches(4.15), Inches(3.95), Inches(0.4), h, 14, CINA, True)
        box(s, x, Inches(4.55), Inches(3.95), Inches(2.2), b, 13, MUTED, spacing=2)
    footer(s, 13)


def ninety(prs):
    s = blank(prs)
    kicker(s, "13    九十天")
    title(s, "做到第 90 天，用这些数复盘")
    phases = [
        ("D1–D7", "10/20–10/26", "备忘录、调度、名单、账户"),
        ("D8–D30", "10/27–11/18", "访谈、微缩场、服务商、首批单"),
        ("D31–D60", "11/19–12/18", "第二场售票、累计 10–20 单陪诊"),
        ("D61–D90", "12/19–1/17", "复盘后决定公司、会员、短途"),
    ]
    for i, (d, date, what) in enumerate(phases):
        x = Inches(0.55 + i * 3.2)
        card(s, x, Inches(1.48), Inches(3.02), Inches(2.15), CINA)
        box(s, x + Inches(0.18), Inches(1.62), Inches(2.7), Inches(0.3), d, 12, CINA, True)
        box(s, x + Inches(0.18), Inches(1.95), Inches(2.7), Inches(0.55), date, 16, INK, True)
        box(s, x + Inches(0.18), Inches(2.55), Inches(2.7), Inches(0.85), what, 13, MUTED)
    # KPI
    kpis = [
        ("15", "高端深访"),
        ("2", "收费会客厅"),
        ("10", "完成陪诊单"),
        ("0", "重大安全事件"),
    ]
    for i, (n, label) in enumerate(kpis):
        x = Inches(0.55 + i * 3.2)
        box(s, x, Inches(3.9), Inches(3.0), Inches(0.55), n, 28, CINA, True)
        box(s, x, Inches(4.5), Inches(3.0), Inches(0.35), "底线  ·  " + label, 13, MUTED)
    box(s, Inches(0.55), Inches(5.15), Inches(12.2), Inches(1.7),
        [
            "停止：连续两场没有真实支付；服务商过不了审核；优化后单笔仍是负贡献；出现处理不了的安全或隐私事件；双方写不下资金和客户归属。",
            "10月31日的讲座、11月29日的漫展保持原计划，不改成这个项目的场，也不把报名数据转过来。",
        ],
        14, INK, spacing=8)
    footer(s, 14)


def guard(prs):
    s = blank(prs)
    kicker(s, "14    底线")
    title(s, "这六条不让步，项目才走得下去")
    items = [
        ("身体", "前 90 天不做给药、助浴、翻身、陪夜。专业护理交给有资质的机构，不自营。"),
        ("医疗", "陪诊停在流程和原话记录。不解释诊断，不代替家属决定。"),
        ("保险", "意外险只保这一次行程的意外。销售保单不进业务。样张看过再放人出门。"),
        ("旅行", "面向客人的交通住宿由持证旅行社承办。我们不靠社群名义自行组团收款。"),
        ("婚恋", "认识可以发生。公司不收费配对，不承诺固定伴侣，不公开财产和住址。"),
        ("数据", "诊疗、行程、家庭关系单独同意。纪要不进群，不进朋友圈。"),
    ]
    for i, (h, b) in enumerate(items):
        x = Inches(0.55 + (i % 3) * 4.2)
        y = Inches(1.5 + (i // 3) * 2.6)
        card(s, x, y, Inches(4.0), Inches(2.4), CINA if i % 2 == 0 else SAGE)
        box(s, x + Inches(0.25), y + Inches(0.25), Inches(3.5), Inches(0.4), h, 18, INK, True)
        box(s, x + Inches(0.25), y + Inches(0.8), Inches(3.5), Inches(1.35), b, 14, MUTED)
    footer(s, 15)


def confirm(prs):
    s = blank(prs)
    kicker(s, "15    请陈红苗确认")
    title(s, "九件事，回复同意或改一句")
    items = [
        "1   前 90 天做会客厅、陪诊、陪聊。护工、保险销售、小程序放到复盘后。",
        "2   茶室微缩按 299 元/人预售，满 8 人付款才办。您当一次模拟客人。",
        "3   谈话陪同先拿 800 元档去做访谈，陪诊试单按 399 元测算，两条不写在同一张海报。",
        "4   陪同类订单陪伴者到手约 70%。带来客人的一方，在平台留存里单独记账。",
        "5   10月26日前指定收款主体。在此之前不收客人的钱。",
        "6   您的客人名单和我的社群都保持各自所有，项目只共用订单和纪要。",
        "7   短途满 8 人再交给旅行社。敦煌不在今冬产品里。",
        "8   每周一次对账。我每周只投入一个固定半天，另加调度的执行时间。",
        "9   公司等到连续四周有预付款，且前 6 单里至少 4 单愿意再约。",
    ]
    box(s, Inches(0.7), Inches(1.45), Inches(12), Inches(5.4), items, 16, INK, spacing=8)
    footer(s, 16)


def this_week(prs):
    s = blank(prs)
    kicker(s, "16    本周")
    title(s, "10月20日就可以开始的六件事")
    actions = [
        ("陈红苗", "今天", "看完分工页，把不能接受的句子改掉发回。"),
        ("陈红苗", "10月22日前", "给出茶室两个日期，并开始约前 5 位深访。"),
        ("陈红苗", "10月26日前", "确认场地能否商用，以及体验价 299 元是否接受。"),
        ("胡总", "10月21日前", "指定调度，定下每周那个半天。"),
        ("胡总", "10月23日前", "在指定社群发出一次供给报名。"),
        ("双方", "10月26日", "通话 40 分钟，定备忘录、账户和第一场能不能预售。"),
    ]
    for i, (who, when, what) in enumerate(actions):
        y = Inches(1.45 + i * 0.88)
        accent = CINA if who == "陈红苗" else (SAGE if who == "胡总" else BRONZE)
        rect(s, Inches(0.55), y, Inches(12.25), Inches(0.78), WHITE, LINE)
        rect(s, Inches(0.55), y, Inches(0.08), Inches(0.78), accent)
        box(s, Inches(0.85), y, Inches(1.5), Inches(0.78), who, 14, accent, True, anchor="middle")
        box(s, Inches(2.4), y, Inches(2.3), Inches(0.78), when, 14, INK, True, anchor="middle")
        box(s, Inches(4.7), y, Inches(7.8), Inches(0.78), what, 15, INK, anchor="middle")
    footer(s, 17)


# ----- Word -----

D_INK = RGBColor(0x1C, 0x24, 0x2B)
D_CINA = RGBColor(0x8C, 0x34, 0x2C)
D_MUTED = RGBColor(0x5C, 0x54, 0x4A)
D_WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def set_run_font(run, size, color, bold=False, east="文泉驿微米黑"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), FONT)
    rfonts.set(qn("w:hAnsi"), FONT)
    rfonts.set(qn("w:eastAsia"), east)


def shade_cell(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")


def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge in ("top", "left", "bottom", "right"):
        el = tcBorders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            tcBorders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), "E3D8C6")


def add_page_field(paragraph):
    run = paragraph.add_run()
    set_run_font(run, 9, D_MUTED)
    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "end")
    run._r.append(fld1)
    run2 = paragraph.add_run()
    set_run_font(run2, 9, D_MUTED)
    run2._r.append(instr)
    run3 = paragraph.add_run()
    set_run_font(run3, 9, D_MUTED)
    run3._r.append(fld2)


def build_doc(path: Path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.2)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.2)
    sec.right_margin = Cm(2.2)
    sec.header_distance = Cm(0.8)
    sec.footer_distance = Cm(0.6)

    header = sec.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = hp.add_run("银发青年 · 重启第二人生")
    set_run_font(r, 9, D_CINA, True)
    r2 = hp.add_run("        呈陈红苗        合作落地稿")
    set_run_font(r2, 9, D_MUTED)

    footer = sec.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    fr = fp.add_run("内部讨论稿  ·  价格与预算均为测算假设  ·  ")
    set_run_font(fr, 9, D_MUTED)
    add_page_field(fp)

    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(11)
    normal.font.color.rgb = D_INK
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "文泉驿微米黑")
    pf = normal.paragraph_format
    pf.line_spacing = 1.15
    pf.space_after = Pt(8)

    def h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(8)
        r = p.add_run(text)
        set_run_font(r, 16, D_CINA, True)
        return p

    def h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(text)
        set_run_font(r, 13, D_INK, True)
        return p

    def para(text, size=11, color=D_INK, bold=False, after=8, before=0):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(text)
        set_run_font(r, size, color, bold)
        return p

    def bullets(items):
        for item in items:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.75)
            p.paragraph_format.first_line_indent = Cm(-0.4)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.12
            r = p.add_run("·  " + item)
            set_run_font(r, 11, D_INK)

    def make_table(headers, rows, widths):
        table = doc.add_table(rows=1 + len(rows), cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        for i, w in enumerate(widths):
            for cell in table.columns[i].cells:
                cell.width = Cm(w)
        for c, text in enumerate(headers):
            cell = table.rows[0].cells[c]
            cell.text = ""
            p = cell.paragraphs[0]
            r = p.add_run(text)
            set_run_font(r, 10, D_WHITE, True)
            shade_cell(cell, "1C242B")
            set_cell_border(cell)
        for ri, row in enumerate(rows):
            for c, text in enumerate(row):
                cell = table.rows[ri + 1].cells[c]
                cell.text = ""
                p = cell.paragraphs[0]
                r = p.add_run(text)
                set_run_font(r, 10, D_INK, c == 0)
                if ri % 2 == 1:
                    shade_cell(cell, "F7F1E6")
                set_cell_border(cell)
        doc.add_paragraph().paragraph_format.space_after = Pt(6)
        return table

    # 封面
    para("银发青年", 12, D_CINA, True, after=2, before=12)
    para("重启第二人生", 26, D_INK, True, after=2)
    para("项目落地策划方案", 16, D_CINA, True, after=10)
    para("呈  陈红苗", 14, D_INK, True, after=2)
    para("胡总团队整理  ·  2026年10月19日", 11, D_MUTED, after=8)
    para("依据 2026年10月13日《银发经济情绪价值项目合作洽谈》综合整理。会上事实、策划建议和测算假设在正文中分开写。尚未签约、尚未付款、尚未授权的资源，不写成已经落实。", 11, D_MUTED, after=12)

    h1("致陈红苗")
    para("陈红苗：")
    para("10月13日您反复问的是两件事：钱从哪里来，以及第一步做什么。这份方案按这个顺序写，方便您直接改。")
    para("我们用一个品牌做两条分开算账的业务。前台叫「第二人生会客厅」，卖小规模的文化社交。另一条叫「品质陪伴」，卖范围写死的陪诊和陪聊。冬季的江浙沪短途是第三条收入，用来过渡到明年春天的主题研学；它等会客厅里有了 8 位愿意预付意向金的客人，再交给有资质的旅行社。")
    para("您主责高端客人、内容品位和场地。我主责年轻人供给、杨浦与虹口的机构走访、交付和台账。1.8 万人的社群是我这边持续运营多年的信任基础，用来组织供给和执行，不当成现成的银发客户，也不自动变成尚未成立的公司的财产。")
    para("请您先看「双方分工」和文末九条确认事项。10月26日我们通话 40 分钟，把备忘录、收款主体和茶室第一场能不能预售定下来。")
    para("胡总", bold=True, after=2)
    para("2026年10月19日", color=D_MUTED)

    h1("一、先看结论")
    para("这件事做成上海的一项可复购服务：50–75 岁、有支付能力、看重话题和人的客人，为一次被认真对待的时间付钱。青年按单获得收入，长者获得在场的人，也可以把自己的经验留在场合里。")
    para("前 90 天只验证三笔钱。")
    bullets([
        "会客厅票款。茶室微缩场 8–12 人，体验价 299 元/人，满 8 人付款才办。第二场约 20 人，测算价 399 元/人。",
        "陪诊与陪聊。陪诊 4 小时测算 399 元，由审核过的服务商交付。谈话陪同先按 800 元 / 2 小时做访谈测档。",
        "冬季微度假。意向金 1,000 元、可退，满 8 人后由持证旅行社承办。敦煌放到 2027 年春季。",
    ])
    para("公司登记放在现金出现之后。连续四个星期有预付款，并且前 6 单里至少 4 单回访说愿意再约，再谈注册和股权。股权不在这份方案里写死。")

    h1("二、会上已经对齐、还没落死的地方")
    h2("2.1 可以当作会议事实的判断")
    bullets([
        "您希望做的是「银发青年重启第二人生」：老年仍然有精神，成员以后可以成为主理人，而不是只被服务。",
        "客人要的是有钱有闲的一档，要真陪伴。2024 年社区养老院的公益互动证明年轻人进得去，没有证明这一档客人会付钱。",
        "您把护工看成偏重、有风险的业务，要求先讲清赚钱的节点，并要有市场对标。",
        "我这边的核心执行大约 7 人，社群约 1.8 万人、40 多个群，画像以三十岁上下为主。白皮书之后约 18 个组织来谈，其中大量仍是意向。",
        "分工的方向当时已经说了：我对接机构需求和年轻供给，您对接高端银发客人。",
    ])
    h2("2.2 四个矛盾在本方案里的收法")
    make_table(
        ["矛盾", "本方案的收法"],
        [
            ["高端生活，和刚需陪诊", "品牌前台是会客厅。陪诊、陪聊作为独立服务出售，不写进同一张海报。"],
            ["您的同龄人，和我的年轻社群", "您负责付费客人。我负责供给、执行和机构拜访。1.8 万人不乘任何转化率去预测老年客户。"],
            ["内容和生意", "第一场就预售。7 天内没有下一笔预付，这场记为未通过。"],
            ["轻资产，和老年服务的责任", "轻，指不先租办公室、不先做软件。每一单仍有合同、保险样张、停单和回访。"],
        ],
        [4.2, 12.2],
    )
    para("会上还没有确定的，包括公司名称、股权、出资金额、场地商业价格、真实可邀请的客户数量、收款主体，以及海上讲堂名称能否用于商业合作。这些在确认前保持开放。", color=D_MUTED)

    h1("三、双方分工")
    para("试点期不新设部门。陈红苗是高端客群与内容主理人。胡总是运营、供给与机构对接主理人。7 人团队里只指定 1 名调度，其余人保持原来的工作。")
    make_table(
        ["事项", "陈红苗", "胡总", "时间"],
        [
            ["试点备忘录", "确认并签署", "起草两页以内", "10月26日"],
            ["收款主体", "共同确认", "准备可追溯账户方案", "未定之前不收款"],
            ["茶室档期", "主责，两个日期", "配合流程", "10月26日"],
            ["其他场地", "确认海上讲堂商用条件", "报价备选场地", "使用前"],
            ["高端访谈", "主责 15 人，名单不外传", "问卷与记录", "11月2日前至少 6 人"],
            ["年轻供给", "参与标准，完成一次模拟", "报名、面试、录取 8 人", "11月2日"],
            ["机构与服务商", "参加高端机构那一次", "主责 8 家", "11月18日前"],
            ["微缩会客厅", "邀请，守品位", "执行、记录、回访", "满 8 人付款后"],
            ["陪诊试单", "可介绍家庭客户", "尽调、派单、回访", "11月"],
            ["短途", "定目的地与品味", "对接旅行社文本", "12月下旬"],
            ["周对账", "参加", "出订单与收支", "每周"],
            ["公司股权", "验证后共同谈", "验证后共同谈", "第90天"],
        ],
        [3.2, 4.6, 4.8, 3.8],
    )
    h2("3.1 权限")
    bullets([
        "陈红苗可以否决一场她认为虚的局、一个不合适的场合，以及任何推销。",
        "胡总可以否决超出每周一个半天的安排、没有保险的上门，以及范围外的新产品。",
        "调度维护预约表，不改价格，不答应范围外的服务。",
    ])
    h2("3.2 资源边界")
    bullets([
        "陈红苗的客人联系方式留在她自己的联络计划里。项目组只接收客人同意之后的订单信息。",
        "胡总的社群、白皮书、杨浦科技企业联合会与复旦相关合作，保持独立。商业使用要单独授权，不能当成新公司的官方背书。",
        "海上讲堂的名称、档期和摄影，以权利人书面条件为准。自用和转给共同项目不是同一件事。",
        "订单纪要归试点项目共有。任何一方退出后，不把纪要拿去销售保险或其他产品。",
    ])
    h2("3.3 钱怎么分")
    para("陪诊、陪聊类订单：陪伴者或服务商结算约 70%，平台留存不超过 30%。平台留存里，交付与调度占 15 个点，带来这位付款人的一方占 10 个点，5 个点留作保险差额和退款。双方共同介绍的，10 个点对半。")
    para("会客厅：票款先扣除茶歇、场地、人员补贴和保障。剩余的 50% 归带来客人的一方，40% 归交付侧，10% 进入备用金。客人由双方共同邀请时，前两项合并后对半。茶室前两场若使用陈红苗的私人空间，场地按零记，同时在台账里记为她的资源投入；从第三场起再商量茶歇成本。")
    para("客人没有预付，这一单不出门。陪伴者按周结算，不从平台留存里临时借支。")

    h1("四、三条变现路径")
    h2("4.1 路径一：会客厅票款")
    para("这是品牌的第一笔钱，也是检验陈红苗的圈子是否成立的那一笔。付款人是长者本人。产品是一个下午的小规模深谈，客人离开时认识同频的人，并被邀请预付下一次。")
    para("茶室微缩场用来做第一次检验，第二场才把人数放到 20 人左右、把价格放到 399 元测算档。两场都要单独复盘。嘉宾费用或场地租金一旦升高，单场可能变成亏损，不能用「票卖出去了」代替复盘。")
    h2("4.2 路径二：陪诊与陪聊")
    para("这是家庭和机构更愿意立刻付钱的刚需。两条 SKU 分开。")
    bullets([
        "办事陪诊：4 小时，测算价 399 元。只做预约提醒、院内路线、排队、医嘱原文和取药。由审核过的服务商派单。胡总主责。",
        "谈话或学习陪同：公共场所，一次一件事。高端测档默认 800 元 / 2 小时，同时问 480 元和 1,200 元。至少 5 位目标客人点头，且陪伴者到手不低于每小时 140 元，才采用该档。陪伴者来自录取的 8 人。",
    ])
    para("2026 年上海普惠陪诊大约每小时 50–80 元，杨浦公共平台 4 小时大约 298 元。399 元买的是指定服务商、过程记录和回访，不跟公益低价竞争。800 元档只对陈红苗介绍的、在意「对方是谁」的客人开放，不写成全市统一价。")
    h2("4.3 路径三：冬季短途微度假")
    para("旅游在会上被看成刚需，同时已经入冬，长途不适合作为第一笔。今冬只做高铁大约两小时内的一日或两日，优先苏州、嘉兴、杭州。它是从会客厅走向明年春季研学的过渡，不替代 11 月的票款和陪诊。")
    bullets([
        "8 位客人各自预付可退意向金 1,000 元。出发前七日退出，或未成团，原路退回。",
        "交通、住宿和面向客人的招徕，由持证旅行社签合同、收取团费。我们负责主题、客人组织和现场陪伴设计。",
        "每天留出不被景点填满的时间。不进购物店，不安排保险说明会，不办速配。",
        "敦煌、云南、厦门、青岛在短途走通之后单独立项。领队档期没有书面确认，对外不宣布人名。西藏和高海拔不排。",
    ])

    h1("五、茶室微缩测试，按这个流程做")
    para("名称：第二人生 · 茶室会客厅（第 1 期）。题目：一个人的晚上，怎样安排成自己还想过的样子。")
    make_table(
        ["环节", "时间", "做什么", "谁"],
        [
            ["模拟陪聊", "开工前 40 分钟", "陈红苗当客人，1 名已录取青年当陪伴者", "胡总在场计时，不插话"],
            ["茶叙", "20 分钟", "只写名字的名牌，互相认识", "调度"],
            ["开场", "10 分钟", "讲清这一场要谈的问题，不做销售", "陈红苗"],
            ["两人深谈", "40 分钟", "退休后最想完成的一件事", "主持人"],
            ["交换一件小事", "15 分钟", "青年与长者各请教一件，可拒绝", "现场"],
            ["下一笔", "15 分钟", "只收三种预付意向：下一场、陪同、短途名单", "胡总"],
        ],
        [3.2, 3.4, 6.6, 3.2],
    )
    para("模拟结束，陈红苗只回答两句：下次是否愿意付钱给这个人；有没有感觉被对付。第二句若是「有」，换一个人重做一次。两次都不通过，微缩场延期，不改成免费场。")
    para("通过标准：7 天内 4 位客人预付下一场或一单陪同。到场人数和气氛不单独作为通过。")
    para("筛选看交流意愿。不要求客人说明存款或保单。不承诺结识对象、名人辅导或高校背书。")

    h1("六、陪诊与陪聊，按这个流程做")
    h2("6.1 出门前的四问")
    bullets([
        "这次要完成的一件事是什么？",
        "哪些话题不要碰？",
        "步行、座位和节奏有什么限制？",
        "希望当次顺手完成的一件小事是什么？",
    ])
    para("同时写明付款人、纪要发给谁、长者是否同意把谈话内容发给子女。长者不同意，当天只保留流程信息。")
    h2("6.2 当天")
    bullets([
        "调度在共享表写下人员、地点、开始与结束时间。",
        "出门前购买当日意外险，保单号入表。保险只覆盖这一次行程中的意外，不覆盖医疗结果和护理责任。",
        "结束后两小时内交出一页纪要：完成的事、听到的原话、下次要带的材料、只有这一次才会出现的事实。",
        "二十四小时内回访付款人，只问还约不约这个人。前 6 单由胡总亲自回访。",
    ])
    h2("6.3 停单")
    para("推销保险、保健品、课程或理财；解释病情或代替家属决定；收受红包、密码或贵重礼物；身体护理；未经同意传播纪要；绕开预约表私下收款。出现任何一项，当天停止，该人员退出试点。")
    h2("6.4 服务商准入")
    para("陪诊不从微信群里临时问谁有空。服务商要有营业信息、人员身份、培训记录、保单及除外责任、应急流程和投诉处理记录。普通兼职青年可以做公共场所的谈话和数字辅导，不承担注射、给药、助浴、翻身和陪夜。")
    para("机构拜访首批 8 家：3 家社区机构，用来看代际活动需求；3 家中高端机构，用来看是否真有采购预算；2 家陪诊或家政供应商，用来看报价和责任。证据分四级：已经付款、有预算和决策人、只是感兴趣、普通介绍。60 天内只推进前两级。约 18 个来谈的组织，不自动算成订单。")

    h1("七、25 岁与 65 岁，以及 1.8 万私域")
    para("会上说的双向，落在服务结尾，不落在招聘口号里。25 岁上下的人有时间、需要一份体面的按单收入。65 岁上下的人有阅历、需要一个真的在场的人。青年可以请教一段经历，长者可以请教一个手机或软件操作。拒绝不影响结算。")
    para("第一期共创是 8 对、90 分钟、公共场所。交付两样东西：一份本人同意留下的人生故事卡，一条已经解决的数字问题。上门服务不从这场活动里直接延伸，要另下陪同订单。")
    para("1.8 万人、40 多个群，是 2018 年以来持续运营的结果。前段的群更活跃。这个规模说明供给招募和活动执行有信任基础。它不说明群里有足够多的 65 岁付费客人。供给报名只发一次，限额 20 人报名、录取 8 人。复购仍走同一张预约表，私下成交则退出试点。")
    para("青年按周结算。培训和模拟的交通实报，不把第一场做成无偿劳动。胡总每周投入上限是一个固定半天，用来面试、回访和拍板；10 月 31 日的讲座和 11 月 29 日的漫展不插入这个项目。")

    h1("八、测算、预算和财务纪律")
    para("下面的数字用来讨论，不是报价，也不是已经同意的结算价。真实场地、嘉宾和服务商报价回来之后，要重算。")
    h2("8.1 会客厅：20 人 × 399 元")
    make_table(
        ["项目", "金额（元）", "说明"],
        [
            ["收入", "7,980", "测试票价"],
            ["场地", "1,200", "须实询；茶室零租金要另记资源投入"],
            ["嘉宾与主持", "1,800", "按合适主讲者估算，不是知名教授报价"],
            ["茶歇", "800", "约 40 元/人"],
            ["执行", "1,000", "人工不按零计算"],
            ["物料与保障", "500", "含必要保险预算，以保单为准"],
            ["沟通与售后", "600", "邀请、回访"],
            ["贡献结余", "2,080", "约 26%，未扣固定费用和税费"],
        ],
        [4.0, 3.5, 9.0],
    )
    h2("8.2 陪诊：20 单 × 399 元")
    make_table(
        ["项目", "单笔（元）", "20 单（元）"],
        [
            ["客户支付", "399", "7,980"],
            ["服务方结算", "280", "5,600"],
            ["保险与准备金", "15", "300"],
            ["排单与回访", "30", "600"],
            ["支付杂费", "10", "200"],
            ["贡献结余", "64", "1,280"],
        ],
        [6.0, 5.0, 5.5],
    )
    para("若 60 天做成 2 场满额会客厅和 20 单陪诊，成交约 23,940 元，贡献结余约 5,440 元。这还没有覆盖固定人员、交通、法律审核和税费。试点的价值是看单位经济是否为正，不是宣布模式已经成熟。")
    para("60 天现金支出设置 20,000–30,000 元上限，作为控制线，不是必须花完的预算。没有合同和客户预算，不预付大额场地和嘉宾费用。每周共享订单、流水和异常，不由个人长期代收。")

    h1("九、90 天日程、底线和停止条件")
    make_table(
        ["阶段", "日期", "动作", "主责"],
        [
            ["D1–D7", "10/20–10/26", "备忘录、调度、报名、茶室日期、账户", "双方"],
            ["D8–D14", "10/27–11/02", "深访、8 家排期、录取 8 人、定主题", "陈红苗邀约，胡总供给"],
            ["D15–D30", "11/03–11/18", "模拟、微缩场、服务商审核、首批试单", "胡总执行，陈红苗出席"],
            ["D31–D60", "11/19–12/18", "第二场售票，累计 10–20 单陪诊", "双方"],
            ["D61–D90", "12/19–1/17", "短途是否成团，是否注册公司", "双方复盘"],
        ],
        [2.6, 3.4, 7.2, 3.2],
    )
    h2("9.1 底线指标")
    make_table(
        ["指标", "底线", "较好"],
        [
            ["高端深访", "15 人", "30 人"],
            ["机构与服务商访谈", "8 家", "15 家"],
            ["收费会客厅", "2 场", "3–4 场"],
            ["会客厅付费人次", "25", "50"],
            ["完成的陪诊单", "10", "20"],
            ["陪诊按时完成", "95%", "98%"],
            ["重大安全事件", "0", "0"],
            ["付费机构", "有可验证的采购意向", "1 家签约付款"],
        ],
        [6.0, 5.2, 5.2],
    )
    h2("9.2 出现这些情况就停")
    bullets([
        "模拟两次未通过：不办付费微缩场。",
        "11月22日前预付款不足 3 笔：12 月不办售票场，先补访谈。",
        "前 6 单愿意再约的不足 4 单：更换名单，不扩大名额。",
        "微缩场 7 天内预付不足 4 位：第二场不售票。",
        "短途意向不足 8 人：不订车、不订房。",
        "优化后单笔仍是负贡献，或出现处理不了的安全、隐私纠纷：暂停该产品。",
        "双方写不下资金审批和客户归属：停止扩张。",
    ])

    h1("十、风险与合规提示")
    para("本节是经营底线，不代替律师、保险经纪、医疗机构或旅行社的意见。正式收款前，按选定的主体再做一次专业核对。")
    bullets([
        "长者是否参加、是否透露家庭信息、是否认识新朋友，由本人决定。紧急联系人用于应急，不替代本人的选择。",
        "前 90 天不做给药、助浴、翻身、陪夜和失能转移。专业护理只转介有资质的机构。",
        "单次意外险不能自动覆盖护理事故、医疗损害和信息泄露。看清承保主体、保额和除外责任再出门。",
        "含交通、住宿、行程并面向客人招徕的团队旅游，由取得许可的旅行社组织。内容策划不能用来替代旅行社许可。",
        "交友活动不承诺配对成功。工作人员不索取礼物、借款和投资。财产、住址、病史不公开。",
        "诊疗信息、行踪和家庭关系按敏感个人信息处理，单独同意，限制访问。",
        "复旦相关机构、杨浦科技企业联合会的名称和成果，未经授权不用于商业背书，也不用于保险或养老投资推广。",
        "按单合作不等于自动排除用工责任。单量上升前，先请律师看合作关系，再扩大名单。",
    ])

    h1("十一、请陈红苗确认")
    para("回复时按编号写「同意」，或直接改一句。第 1、2、4、5 条没有回复之前，报名可以先发，收款和上门不开始。")
    bullets([
        "前 90 天的收入产品是会客厅、陪诊和陪聊。护工、保险销售和小程序放在复盘之后。",
        "茶室微缩按 299 元/人预售，满 8 人付款才办。您亲自当一次模拟客人。",
        "谈话陪同先用 800 元档访谈；陪诊试单按 399 元测算。两条不出现在同一张海报上。",
        "陪同类订单按约 70% 结算给在场的一方，平台留存不超过 30%，其中带来客户的 10 个点单独记账。",
        "10月26日前指定收款主体。此前不向客人收款。",
        "您的名单和我的社群各自所有。项目共有的是订单和纪要。",
        "短途满 8 人再交给旅行社。敦煌不进入今冬产品。",
        "每周对一次账。我每周投入一个固定半天，另加调度的执行时间。",
        "公司等到连续四周有预付款，且前 6 单里至少 4 单愿意再约，再讨论注册和股权。",
    ])

    h1("十二、本周动作")
    make_table(
        ["谁", "何时", "做什么"],
        [
            ["陈红苗", "10月20日", "看完分工，把不能接受的句子改回"],
            ["陈红苗", "10月22日前", "给出茶室两个日期，约前 5 位深访"],
            ["陈红苗", "10月26日前", "确认场地商用条件和 299 元体验价"],
            ["胡总", "10月21日前", "指定调度，定下每周那个半天"],
            ["胡总", "10月23日前", "在指定社群发出一次供给报名"],
            ["双方", "10月26日", "通话 40 分钟，定备忘录、账户和第一场预售"],
        ],
        [3.2, 3.6, 9.6],
    )

    h1("附录一  可以发给客人的邀请")
    para("人生下半场，依然是自己的主场。")
    para("我们在筹备一场小规模的「第二人生 · 茶室会客厅」，邀请对文化、友谊和一段有质量的谈话有兴趣的朋友。这一场不设赞助，名牌只写名字。时间、人数和费用以正式邀请为准。首期体验名额 8–12 人。")

    h1("附录二  供给报名的四句话")
    bullets([
        "上海市内的按单陪同，先面试再上岗，名额 8 人。",
        "做就医流程记录、谈话或教会一个操作，不做护理。",
        "每一单有预付、有意外险、有当天纪要。",
        "报名后单独联系，不在群里接龙。",
    ])

    h1("附录三  备忘录至少写上的句子")
    bullets([
        "试点 60 天，非排他。项目名称暂定。",
        "各自原有资源不因本备忘录转移所有权。",
        "前 90 天的产品、分成、停单，以双方确认后的本方案为准。",
        "预付款进入指定账户后，在场一方按周结算。",
        "名单和纪要的使用范围写死。退出后继续履行已收款订单，或按客人选择退款。",
        "注册公司、股权和薪酬另纸约定。未另纸约定之前，双方都没有股权承诺。",
    ])

    h1("附录四  材料与测算说明")
    para("内部材料为 2026年10月13日洽谈记录，录音约 14:03–15:38。重点包括高端养老体验与项目缘起、情绪价值与活动清单、陪诊护工与抽成、社群画像和收尾分工。")
    para("399 元票价、399 元陪诊、800 元谈话档、1,200 元场地、20,000–30,000 元预算上限和各阶段人数，都是方案测算。上海普惠陪诊价格带参考 2026 年公开报道，大约每小时 50–80 元，杨浦分级服务 4 小时大约 298 元。静安八号只作为品质参照，不写成合作方。")
    para("本方案用于推进陈红苗与胡总的试点合作，不构成法律、财税、保险、医疗或旅游业务意见。", color=D_MUTED)

    core = doc.core_properties
    core.title = "银发青年·重启第二人生项目落地策划方案"
    core.author = "胡总团队"
    core.subject = "呈陈红苗"
    core.category = "合作落地稿"
    core.comments = "2026年10月19日综合整理。价格为测算假设。"
    doc.save(path)


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    ppt = OUT_DIR / PPT_NAME
    doc = OUT_DIR / DOC_NAME
    build_ppt(ppt)
    build_doc(doc)
    print(ppt)
    print(doc)


if __name__ == "__main__":
    main()
