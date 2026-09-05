# -*- coding: utf-8 -*-
"""给港大经管上海中心看的汇报 PPT。口吻按当面汇报来写。"""

from __future__ import annotations

import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

OUT = Path(__file__).resolve().parent.parent / "output" / "港大经管上海中心_联合合作汇报.pptx"

GREEN = RGBColor(0x00, 0x3D, 0x2E)
GREEN2 = RGBColor(0x0A, 0x5C, 0x46)
GOLD = RGBColor(0xC4, 0xA3, 0x5A)
CREAM = RGBColor(0xF7, 0xF4, 0xEC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1A, 0x24, 0x20)
GREY = RGBColor(0x5B, 0x6B, 0x64)
LIGHT = RGBColor(0xE8, 0xEF, 0xEA)
FONT = "微软雅黑"
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
TOTAL = 10


def set_font(run, size, color=DARK, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    rpr = run._r.get_or_add_rPr()
    ea = rpr.find(qn("a:ea"))
    if ea is None:
        ea = rpr.makeelement(qn("a:ea"), {})
        rpr.append(ea)
    ea.set("typeface", FONT)


def add_rect(slide, x, y, w, h, fill):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_round(slide, x, y, w, h, fill):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    try:
        shp.adjustments[0] = 0.08
    except Exception:
        pass
    return shp


def add_text(slide, x, y, w, h, text, size, color=DARK, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    set_font(r, size, color, bold)
    return tb


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def header(slide, title, idx):
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.02), GREEN)
    add_rect(slide, 0, Inches(1.02), SLIDE_W, Pt(3), GOLD)
    add_text(slide, Inches(0.55), Inches(0.16), Inches(11.2), Inches(0.7), title, 24, WHITE, True, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(11.5), Inches(0.16), Inches(1.4), Inches(0.7), f"{idx}/{TOTAL}", 12, GOLD, True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def footer(slide):
    add_text(
        slide,
        Inches(0.55),
        Inches(7.12),
        Inches(12.2),
        Inches(0.28),
        "复旦大学住房政策研究中心  ·  上海市杨浦区科技企业联合会  ·  致香港大学经管学院上海中心",
        10,
        GREY,
    )


def bullets(slide, x, y, w, h, items, size=16, gap=10):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for it in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(gap)
        p.line_spacing = 1.2
        rm = p.add_run()
        rm.text = "·  "
        set_font(rm, size, GOLD, True)
        r = p.add_run()
        r.text = it
        set_font(r, size, DARK)
    return tb


def slide_cover(prs):
    s = blank(prs)
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, GREEN)
    add_rect(s, 0, 0, Inches(0.16), SLIDE_H, GOLD)
    add_text(s, Inches(0.85), Inches(1.35), Inches(11.5), Inches(0.4), "致　香港大学经管学院上海中心", 16, GOLD, False)
    add_text(s, Inches(0.85), Inches(2.0), Inches(11.5), Inches(0.9), "联合合作汇报", 44, WHITE, True)
    add_text(s, Inches(0.85), Inches(3.0), Inches(11.5), Inches(0.5), "出海专题课，以及请中心主任到场一事", 20, GOLD, False)
    add_rect(s, Inches(0.85), Inches(3.7), Inches(2.0), Pt(3), GOLD)
    add_text(
        s,
        Inches(0.85),
        Inches(4.05),
        Inches(11.2),
        Inches(1.1),
        "上次在中心见面之后，把可以马上做的事写成这份汇报，并附上协议，请中心审阅。",
        18,
        WHITE,
        False,
    )
    add_text(
        s,
        Inches(0.85),
        Inches(5.55),
        Inches(11.2),
        Inches(0.9),
        "复旦大学住房政策研究中心\n上海市杨浦区科技企业联合会",
        16,
        RGBColor(0xD5, 0xE4, 0xDC),
        False,
    )
    add_text(s, Inches(0.85), Inches(6.7), Inches(11.2), Inches(0.35), "2026年9月　　联系人：潘嘉琰老师", 14, GOLD, False)


def slide_why(prs):
    s = blank(prs)
    header(s, "这份材料是什么", 2)
    add_round(s, Inches(0.5), Inches(1.35), Inches(12.3), Inches(1.7), CREAM)
    add_text(
        s,
        Inches(0.75),
        Inches(1.5),
        Inches(11.8),
        Inches(1.4),
        "潘老师：上次在外滩中心，谈到华东、华中的招生和市场，也谈到出海，以及请中心主任到现场看一看。回来后我们只写了两件事，把费用和交付写清楚，方便中心内部过目。",
        18,
        DARK,
        False,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    cards = [
        ("第一件", "在外滩中心办一场出海专题课。三十到四十人。我们把人请来，中心出场地和老师，课后由招生老师单独沟通。"),
        ("第二件", "请中心主任参加我们正在办的一场出海活动。主任方不方便，完全按中心的时间。"),
        ("先不做的", "港大本部的出海课程怎么推广、原创谷怎么合作、学生在上海实习参访，本期不写进交付，以后另说。"),
    ]
    for i, (t, d) in enumerate(cards):
        y = Inches(3.25 + i * 1.2)
        add_round(s, Inches(0.5), y, Inches(12.3), Inches(1.08), WHITE)
        add_rect(s, Inches(0.5), y, Inches(1.6), Inches(1.08), GREEN if i < 2 else GREEN2)
        add_text(s, Inches(0.5), y, Inches(1.6), Inches(1.08), t, 16, WHITE, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(2.3), y + Inches(0.15), Inches(10.2), Inches(0.78), d, 15, DARK, False, anchor=MSO_ANCHOR.MIDDLE)
    footer(s)


def slide_fit(prs):
    s = blank(prs)
    header(s, "我们为什么来谈这件事", 3)
    bullets(
        s,
        Inches(0.6),
        Inches(1.4),
        Inches(12.1),
        Inches(5.3),
        [
            "中心在上海已经开始运转。招生宣传、学生来沪学习实习、复旦与港大合办课程迁回中心上课，这些都需要把合适的人请到外滩来。",
            "潘老师负责华东、华中高端教育的招生和市场。我们长期在上海办企业和校友活动，手里有一批企业负责人。",
            "港大在企业出海、赴港上市这件事上，内地学校很难替代。我们能做的，是把真正在做出海打算的人请到中心，听老师讲、再单独谈。",
            "所以这场合作就放在出海上，不办空的招生说明会。",
        ],
        size=17,
        gap=14,
    )
    footer(s)


def slide_class(prs):
    s = blank(prs)
    header(s, "外滩出海专题课", 4)
    rows = [
        ("时间", "协议签订后一个月左右，选一个工作日下午，四点左右结束。"),
        ("地点", "香港大学经管学院上海中心，外滩 SOHO F 栋。"),
        ("人数", "三十到四十人。"),
        ("来宾", "正在考虑出海或赴港上市的企业负责人。"),
        ("内容", "港大老师或校友讲出海和港股路径；企业之间交换看法；课后由招生老师单独沟通。"),
        ("分工", "中心提供场地、老师、招生安排。我们负责请人、现场和会后把名单整理好交给中心。"),
    ]
    for i, (k, v) in enumerate(rows):
        y = Inches(1.28 + i * 0.88)
        add_round(s, Inches(0.5), y, Inches(12.3), Inches(0.78), WHITE if i % 2 == 0 else CREAM)
        add_text(s, Inches(0.7), y, Inches(1.5), Inches(0.78), k, 16, GREEN, True, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(2.4), y, Inches(10.1), Inches(0.78), v, 15, DARK, False, anchor=MSO_ANCHOR.MIDDLE)
    footer(s)


def slide_director(prs):
    s = blank(prs)
    header(s, "请中心主任到场", 5)
    add_round(s, Inches(0.5), Inches(1.4), Inches(12.3), Inches(2.2), CREAM)
    add_text(
        s,
        Inches(0.8),
        Inches(1.6),
        Inches(11.7),
        Inches(1.85),
        "招生和市场是潘老师的工作。若以后还要在学院层面合作，需要中心主任先看过现场。\n\n我们近期有出海主题的活动，会请总领事或企业负责人到场。希望请潘老师代为邀请主任出席一次。来不来，以中心的安排为准。",
        18,
        DARK,
        False,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    bullets(
        s,
        Inches(0.7),
        Inches(3.9),
        Inches(12.0),
        Inches(2.8),
        [
            "我们发出书面邀请，写明时间、地点和到场嘉宾。",
            "主任因公务不能来，不影响这一项算已完成。",
            "这一次不另收费，算在下面写的八万八千元里面。",
        ],
        size=17,
        gap=12,
    )
    footer(s)


def slide_90(prs):
    s = blank(prs)
    header(s, "九十天怎么排", 6)
    steps = [
        ("第 1 周", "签订协议，支付费用，把外滩那场的日期定下来。"),
        ("第 2–3 周", "拿出拟邀请名单，不少于八十人，双方看过再定。"),
        ("第 4–7 周", "办外滩专题课；同时发出请主任出席的邀请。"),
        ("第 8–12 周", "把课上的情况和名单交给中心。后面若要续办或谈别的事，再另写。"),
    ]
    for i, (t, d) in enumerate(steps):
        y = Inches(1.35 + i * 1.3)
        add_round(s, Inches(0.5), y, Inches(12.3), Inches(1.15), WHITE)
        add_rect(s, Inches(0.5), y, Inches(2.4), Inches(1.15), GREEN)
        add_text(s, Inches(0.5), y, Inches(2.4), Inches(1.15), t, 18, WHITE, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(3.15), y, Inches(9.4), Inches(1.15), d, 17, DARK, False, anchor=MSO_ANCHOR.MIDDLE)
    footer(s)


def slide_roles(prs):
    s = blank(prs)
    header(s, "各自做什么", 7)
    add_round(s, Inches(0.45), Inches(1.35), Inches(6.1), Inches(5.4), WHITE)
    add_rect(s, Inches(0.45), Inches(1.35), Inches(6.1), Inches(0.7), GREEN)
    add_text(s, Inches(0.45), Inches(1.35), Inches(6.1), Inches(0.7), "香港大学经管学院上海中心", 18, WHITE, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    bullets(
        s,
        Inches(0.7),
        Inches(2.25),
        Inches(5.6),
        Inches(4.2),
        [
            "提供外滩中心场地和名义",
            "安排老师或校友",
            "招生跟进、面试和录取仍由学院决定",
            "请潘老师协助邀请中心主任",
        ],
        size=16,
        gap=12,
    )
    add_round(s, Inches(6.8), Inches(1.35), Inches(6.1), Inches(5.4), WHITE)
    add_rect(s, Inches(6.8), Inches(1.35), Inches(6.1), Inches(0.7), GREEN2)
    add_text(s, Inches(6.8), Inches(1.35), Inches(6.1), Inches(0.7), "住房政策研究中心、杨浦科企联", 18, WHITE, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    bullets(
        s,
        Inches(7.05),
        Inches(2.25),
        Inches(5.6),
        Inches(4.2),
        [
            "策划、请人、现场",
            "会后把纪要和名单交给中心",
            "安排请主任出席的那一场活动",
            "开票、收款",
        ],
        size=16,
        gap=12,
    )
    footer(s)


def slide_fee(prs):
    s = blank(prs)
    header(s, "费用", 8)
    add_round(s, Inches(0.5), Inches(1.3), Inches(12.3), Inches(1.55), GREEN)
    add_text(s, Inches(0.8), Inches(1.42), Inches(11.7), Inches(0.4), "前期工作费用", 14, GOLD, True)
    add_text(
        s,
        Inches(0.8),
        Inches(1.82),
        Inches(11.7),
        Inches(0.8),
        "人民币捌万捌仟元整（88,000 元）　　协议生效后十个工作日内一次付清，开增值税发票",
        20,
        WHITE,
        True,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    add_text(s, Inches(0.6), Inches(3.1), Inches(6.0), Inches(0.4), "这八万八千元包括", 16, GREEN, True)
    bullets(
        s,
        Inches(0.55),
        Inches(3.55),
        Inches(6.0),
        Inches(3.2),
        [
            "九十天的筹备",
            "拟邀请名单（不少于八十人）",
            "外滩出海专题课一场",
            "请主任出席的一次活动安排",
            "课后纪要和名单",
        ],
        size=16,
        gap=8,
    )
    add_text(s, Inches(6.9), Inches(3.1), Inches(6.0), Inches(0.4), "不包括", 16, GREEN, True)
    bullets(
        s,
        Inches(6.85),
        Inches(3.55),
        Inches(6.0),
        Inches(3.2),
        [
            "学费、报名费，也不按人头分成",
            "第二场及以后（如需续办，每场六万八千元，另签）",
            "港大本部出海课程的招生",
            "原创谷的日常事务",
        ],
        size=16,
        gap=8,
    )
    footer(s)


def slide_ask(prs):
    s = blank(prs)
    header(s, "请中心定几件事", 9)
    items = [
        "外滩专题课放在哪一天。",
        "主任方便出席的大致时间。来不了也没有关系，我们按书面邀请办理。",
        "协议由哪一家主体签，发票抬头是谁，款什么时候付。",
    ]
    for i, t in enumerate(items):
        y = Inches(1.4 + i * 1.35)
        add_round(s, Inches(0.55), y, Inches(12.2), Inches(1.2), WHITE)
        add_round(s, Inches(0.8), y + Inches(0.28), Inches(0.65), Inches(0.65), GOLD)
        add_text(s, Inches(0.8), y + Inches(0.28), Inches(0.65), Inches(0.65), str(i + 1), 20, GREEN, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(1.7), y + Inches(0.2), Inches(10.7), Inches(0.8), t, 18, DARK, False, anchor=MSO_ANCHOR.MIDDLE)
    add_text(
        s,
        Inches(0.6),
        Inches(5.6),
        Inches(12.1),
        Inches(1.2),
        "协议文本随这份汇报一并送上。签好、费用到账后七日内，我们提交拟邀请名单初稿。\n联系人：潘嘉琰　　(86) 180 1860 6086　　jyanpan@hku.hk",
        15,
        GREY,
        False,
        align=PP_ALIGN.CENTER,
    )
    footer(s)


def slide_close(prs):
    s = blank(prs)
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, GREEN)
    add_rect(s, 0, 0, Inches(0.16), SLIDE_H, GOLD)
    add_text(s, Inches(0.9), Inches(2.2), Inches(11.5), Inches(0.7), "请中心审阅。有要改的地方，我们按中心的意见改。", 24, WHITE, True)
    add_text(
        s,
        Inches(0.9),
        Inches(3.2),
        Inches(11.5),
        Inches(1.2),
        "复旦大学住房政策研究中心\n上海市杨浦区科技企业联合会\n2026年9月",
        18,
        GOLD,
        False,
    )
    add_text(s, Inches(0.9), Inches(5.5), Inches(11.5), Inches(0.8), "随附：联合策划服务合作协议", 16, WHITE, False)


def build(path: Path | None = None) -> Path:
    path = path or OUT
    path.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    slide_cover(prs)
    slide_why(prs)
    slide_fit(prs)
    slide_class(prs)
    slide_director(prs)
    slide_90(prs)
    slide_roles(prs)
    slide_fee(prs)
    slide_ask(prs)
    slide_close(prs)
    prs.save(path)
    return path


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else OUT
    print(f"已生成 {build(out)}")
