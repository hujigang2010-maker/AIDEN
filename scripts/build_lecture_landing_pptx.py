#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把《10月31日讲座_落地方案_价值实施_赞助落地》做成内部执行演示文稿。

画面为 16:9。颜色沿用同一场讲座赞助方案的宣纸、朱砂和墨色，
便于两套材料放在一起；本文件含收支、决策和取消规则，整册不对外发送。
"""

from __future__ import annotations

import math
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output"
PPTX_PATH = OUT / "10月31日讲座_落地方案_价值实施_赞助落地.pptx"

W = 13.333333
H = 7.5

INK = RGBColor(0x14, 0x11, 0x0E)
INK_SOFT = RGBColor(0x23, 0x1C, 0x18)
PAPER = RGBColor(0xF6, 0xF1, 0xE8)
PAPER_2 = RGBColor(0xEF, 0xE8, 0xDC)
CARD = RGBColor(0xFF, 0xFC, 0xF7)
CINNABAR = RGBColor(0x8E, 0x2A, 0x28)
CINNABAR_DEEP = RGBColor(0x6E, 0x21, 0x1F)
GOLD = RGBColor(0xA6, 0x84, 0x45)
GOLD_LINE = RGBColor(0xC4, 0xA2, 0x65)
GOLD_TEXT = RGBColor(0x7A, 0x5A, 0x22)
CREAM = RGBColor(0xF7, 0xF1, 0xE6)
CREAM_DIM = RGBColor(0xD9, 0xCB, 0xB6)
TEXT = RGBColor(0x1C, 0x19, 0x16)
MUTED = RGBColor(0x5C, 0x56, 0x4E)
LINE = RGBColor(0xE3, 0xD9, 0xC8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PINE = RGBColor(0x1E, 0x3A, 0x32)
PINE_SOFT = RGBColor(0xE7, 0xF0, 0xEC)
WASH = RGBColor(0xF8, 0xF0, 0xEA)
GOLD_WASH = RGBColor(0xF8, 0xF3, 0xE6)

FONT = "微软雅黑"
BOXES: list[dict] = []


def set_run_font(run, size: float, color: RGBColor, bold: bool = False, font: str = FONT) -> None:
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    run.font.italic = False
    r_pr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs", "a:sym"):
        node = r_pr.find(qn(tag))
        if node is not None:
            r_pr.remove(node)
    for tag in ("a:latin", "a:ea", "a:cs"):
        r_pr.append(r_pr.makeelement(qn(tag), {"typeface": font}))


def _solid(shape, color: RGBColor) -> None:
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.shadow.inherit = False


def add_rect(slide, x, y, w, h, color: RGBColor):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    _solid(shape, color)
    return shape


def add_round(slide, x, y, w, h, fill: RGBColor, line: RGBColor | None = None, radius: float = 0.08):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)
    )
    _solid(shape, fill)
    if line is not None:
        shape.line.color.rgb = line
        shape.line.width = Pt(0.75)
    try:
        shape.adjustments[0] = radius
    except Exception:
        pass
    return shape


def add_text(slide, x, y, w, h, paragraphs: list[dict], *, anchor=MSO_ANCHOR.TOP, page: int = 0, pad: float = 0.06):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(pad)
    tf.margin_right = Inches(pad)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    expanded: list[dict] = []
    for item in paragraphs:
        parts = str(item.get("text", "")).split("\n")
        for j, part in enumerate(parts):
            piece = dict(item)
            piece["text"] = part
            if j > 0:
                piece["before"] = 0
            expanded.append(piece)
    for i, item in enumerate(expanded):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = item.get("align", PP_ALIGN.LEFT)
        p.level = 0
        p.space_before = Pt(item.get("before", 0))
        p.space_after = Pt(item.get("after", 0))
        p.line_spacing = item.get("spc", 1.0)
        run = p.add_run()
        run.text = item.get("text", "")
        set_run_font(run, item.get("size", 14), item.get("color", TEXT), item.get("bold", False), item.get("font", FONT))
    BOXES.append(
        {
            "page": page,
            "x": x,
            "y": y,
            "w": w,
            "h": h,
            "pad": pad,
            "paras": [
                {
                    "text": item.get("text", ""),
                    "size": item.get("size", 14),
                    "before": item.get("before", 0),
                    "after": item.get("after", 0),
                    "spc": item.get("spc", 1.0),
                }
                for item in expanded
            ],
        }
    )
    return box


def t(text, size=14, color=TEXT, bold=False, align=PP_ALIGN.LEFT, before=0, after=0, spc=1.0):
    return {
        "text": text,
        "size": size,
        "color": color,
        "bold": bold,
        "align": align,
        "before": before,
        "after": after,
        "spc": spc,
    }


def new_slide(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def paint(slide, color=PAPER):
    add_rect(slide, 0, 0, W, H, color)


def footer(slide, page: int, total: int):
    add_rect(slide, 0.52, 7.08, W - 1.04, 0.01, LINE)
    add_text(
        slide,
        0.52,
        7.14,
        9.4,
        0.28,
        [t("内部执行稿  ·  请勿整册外发  ·  10 月 31 日讲座落地方案", 11, MUTED)],
        anchor=MSO_ANCHOR.MIDDLE,
        page=page,
    )
    add_text(
        slide,
        10.3,
        7.14,
        2.5,
        0.28,
        [t(f"{page:02d}  /  {total:02d}", 11, MUTED, align=PP_ALIGN.RIGHT)],
        anchor=MSO_ANCHOR.MIDDLE,
        page=page,
    )


def header(slide, kicker: str, title: str, page: int, total: int):
    paint(slide)
    add_rect(slide, 0, 0, W, 0.08, CINNABAR)
    add_text(slide, 0.52, 0.18, 12.2, 0.28, [t(kicker, 13, CINNABAR, True)], page=page)
    add_text(slide, 0.52, 0.46, 12.3, 0.50, [t(title, 26, TEXT, True)], page=page)
    add_rect(slide, 0.52, 1.04, 0.72, 0.035, GOLD_LINE)
    footer(slide, page, total)


def cell(text, color=TEXT, bold=False, align=PP_ALIGN.LEFT, bg=None):
    return {"text": text, "color": color, "bold": bold, "align": align, "bg": bg}


def paint_table(slide, x, y, widths, headers, rows, page, row_h=0.48, header_h=0.42, size=13, header_colors=None):
    cursor = x
    for i, label in enumerate(headers):
        bg = INK if not header_colors else header_colors[i]
        add_rect(slide, cursor, y, widths[i], header_h, bg)
        add_text(
            slide,
            cursor,
            y,
            widths[i],
            header_h,
            [t(label, 13, CREAM, True, PP_ALIGN.CENTER)],
            anchor=MSO_ANCHOR.MIDDLE,
            page=page,
            pad=0.04,
        )
        cursor += widths[i]
    for r, row in enumerate(rows):
        yy = y + header_h + r * row_h
        cursor = x
        for c, item in enumerate(row):
            if not isinstance(item, dict):
                item = cell(item, align=PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER)
            bg = item.get("bg")
            if bg is None:
                bg = CARD if r % 2 == 0 else PAPER_2
            add_rect(slide, cursor, yy, widths[c], row_h, bg)
            add_text(
                slide,
                cursor,
                yy,
                widths[c],
                row_h,
                [t(item["text"], size, item.get("color", TEXT), item.get("bold", False), item.get("align", PP_ALIGN.LEFT), spc=1.0)],
                anchor=MSO_ANCHOR.MIDDLE,
                page=page,
                pad=0.08,
            )
            cursor += widths[c]
    return y + header_h + len(rows) * row_h


# ---------------------------------------------------------------------------
# 各页
# ---------------------------------------------------------------------------

def slide_cover(prs, page, total):
    s = new_slide(prs)
    paint(s, INK)
    add_rect(s, 0, 0, 0.14, H, CINNABAR)
    add_round(s, 0.58, 0.42, 0.62, 0.62, CINNABAR, radius=0.12)
    add_text(s, 0.58, 0.50, 0.62, 0.48, [t("心", 22, CREAM, True, PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.MIDDLE, page=page)
    add_text(s, 1.40, 0.48, 8, 0.28, [t("内部执行稿  ·  请勿整册外发", 14, GOLD_LINE, True)], anchor=MSO_ANCHOR.MIDDLE, page=page)
    add_text(s, 10.2, 0.48, 2.6, 0.28, [t("2026.10.10", 14, CREAM_DIM, align=PP_ALIGN.RIGHT)], anchor=MSO_ANCHOR.MIDDLE, page=page)

    add_text(s, 0.58, 1.55, 12.1, 0.36, [t("2026 年 10 月 31 日（星期六）09:00—12:00  ·  上海（待公布）", 16, GOLD_LINE, True)], page=page)
    add_text(s, 0.58, 2.05, 12.1, 0.70, [t("王德峰教授《传习录》与阳明心学", 32, CREAM, True)], page=page)
    add_text(s, 0.58, 2.82, 12.1, 0.50, [t("落地方案  ·  价值实施  ·  赞助落地", 22, CREAM, True)], page=page)
    add_rect(s, 0.58, 3.48, 1.15, 0.035, GOLD_LINE)
    add_text(
        s,
        0.58,
        3.68,
        12.1,
        0.70,
        [t("联合主办\n复旦大学住房政策研究中心  ·  上海市杨浦区科技企业联合会  ·  首乾书院", 15, CREAM_DIM, spc=1.15)],
        page=page,
    )

    chips = [
        ("10 万元", "讲师费测算，目前唯一现金成本"),
        ("约 200 座", "场地容量，以书面确认为准"),
        ("4 人", "总负责、内容、票务、现场"),
    ]
    for i, (head, sub) in enumerate(chips):
        x = 0.58 + i * 4.1
        add_round(s, x, 4.70, 3.92, 1.28, INK_SOFT, radius=0.1)
        add_text(s, x + 0.22, 4.84, 3.5, 0.48, [t(head, 22, GOLD_LINE, True)], page=page)
        add_text(s, x + 0.22, 5.36, 3.5, 0.42, [t(sub, 13, CREAM)], page=page)

    add_text(
        s,
        0.58,
        6.25,
        12.1,
        0.70,
        [t("成稿距活动 21 天。讲师费、场地容量、上座率都是待确认假设，不是已核实数据。\n拿到讲师费书面数字和场地座位数后，按收支页、讲师费页重算。", 13, CREAM_DIM, spc=1.2)],
        page=page,
    )
    add_text(s, 11.3, 6.95, 1.5, 0.28, [t(f"{page:02d}  /  {total:02d}", 12, CREAM_DIM, align=PP_ALIGN.RIGHT)], page=page)


def slide_decision(prs, page, total):
    s = new_slide(prs)
    header(s, "一页结论", "六项决定，两条线同时推进", page, total)
    cards = [
        ("活动性质", "一个上午", "约 3 小时单场公开讲座\n不拆“半天票”"),
        ("票价", "699 元", "早鸟 499，10/20 前限 60 张\n团体 599，同一单位 5 张起\n前排 999，限 20 张"),
        ("赞助", "6 万保底", "争取 8 万。战略 3 万，限 1 家\n特别 1 万，2—3 家\n品牌 5,000 元，3—5 家"),
        ("保本门票", "162 / 71 / 41", "无赞助须卖 162 张\n赞助 6 万到账后 71 张\n赞助 8 万到账后 41 张"),
        ("关键日期", "10 月 22 日", "10/12 书面确认，10/13 开售\n10/20 赞助到账，早鸟截止\n10/22 决定是否继续"),
        ("团队与底线", "4 个角色", "不先付清讲师费\n赞助未到账不计入收入\n未确认的权益不写进宣传"),
    ]
    for i, (kicker, value, body) in enumerate(cards):
        col, row = i % 3, i // 3
        x = 0.52 + col * 4.16
        y = 1.26 + row * 2.28
        add_round(s, x, y, 4.00, 2.14, CARD, LINE, 0.08)
        add_text(s, x + 0.16, y + 0.10, 3.68, 0.26, [t(kicker, 12, CINNABAR, True)], page=page)
        add_text(s, x + 0.16, y + 0.36, 3.68, 0.42, [t(value, 20, TEXT, True)], page=page)
        add_text(s, x + 0.16, y + 0.84, 3.68, 1.16, [t(body, 13, MUTED, spc=1.08)], page=page)
    add_round(s, 0.52, 5.90, 12.30, 1.02, INK, radius=0.08)
    add_text(
        s,
        0.72,
        6.02,
        11.9,
        0.78,
        [t("门票负责“场子满”，赞助负责“讲师费落袋”。\n两条线同时推进，10 月 22 日看已到账数据，决定是否继续。", 16, CREAM, True, spc=1.12)],
        anchor=MSO_ANCHOR.MIDDLE,
        page=page,
    )


def slide_gates(prs, page, total):
    s = new_slide(prs)
    header(s, "落地方案  ·  开售门槛", "四份书面确认齐全，才允许开售", page, total)
    items = [
        ("01", "讲师费与付款", "找首乾书院", ["写明含税或不含税", "签约付 30% 定金", "活动后 3 个工作日付清", "老师取消，已付全额退还"]),
        ("02", "场地与容量", "找场地方", ["免费使用的书面确认", "可用座位到底有多少", "进场时间与彩排时间", "音响、投影、话筒"]),
        ("03", "名义与肖像", "三家主办＋讲师方", ["同意联合主办并收费", "同意使用姓名和肖像", "同意使用课程题目", "校内未批则改学术支持"]),
        ("04", "收款与开票", "书院或联合会", ["门票由谁收款", "赞助由谁收款", "由谁开具发票", "退款由谁执行"]),
    ]
    for i, (num, title, who, lines) in enumerate(items):
        x = 0.52 + i * 3.20
        add_round(s, x, 1.26, 3.06, 4.52, CARD, LINE, 0.08)
        add_text(s, x + 0.16, 1.40, 2.74, 0.34, [t(num, 18, CINNABAR, True)], page=page)
        add_text(s, x + 0.16, 1.80, 2.74, 0.64, [t(title, 18, TEXT, True)], page=page)
        add_text(s, x + 0.16, 2.48, 2.74, 0.36, [t(who, 13, GOLD_TEXT, True)], page=page)
        add_text(s, x + 0.16, 2.96, 2.74, 2.50, [t("\n".join(lines), 14, TEXT, spc=1.35)], page=page)
    add_round(s, 0.52, 5.94, 12.30, 0.98, WASH, radius=0.08)
    add_text(
        s,
        0.74,
        6.06,
        11.9,
        0.74,
        [t("缺任何一份，10 月 13 日不开售。\n复旦中心若未获校内同意，改列“学术支持单位”，不作为收费方。", 15, TEXT, spc=1.15)],
        anchor=MSO_ANCHOR.MIDDLE,
        page=page,
    )


def slide_tickets(prs, page, total):
    s = new_slide(prs)
    header(s, "落地方案  ·  票务", "四档票价，早鸟只卖到 10 月 20 日", page, total)
    tickets = [
        ("早鸟席", "499", "限 60 张", "10/20 24:00 截止", "入场听讲\n纸质讲义\n茶歇", False),
        ("标准席", "699", "主售票档", "数量以容量为准", "入场听讲\n纸质讲义\n茶歇", True),
        ("企业团体", "599", "5 张起", "同一单位，可统一开票", "权益与标准席相同\n按人计价\n便于企业报名", False),
        ("前排席", "999", "限 20 张", "前三排座位", "含上述全部权益\n签名书须确认后再加", False),
    ]
    for i, (name, price, quota, note, rights, main) in enumerate(tickets):
        x = 0.52 + i * 3.20
        add_round(s, x, 1.28, 3.06, 4.55, CARD, LINE, 0.08)
        add_rect(s, x, 1.28, 3.06, 0.12, CINNABAR if main else INK)
        add_text(s, x + 0.16, 1.52, 2.74, 0.32, [t(name, 14, CINNABAR if main else MUTED, True)], page=page)
        add_text(s, x + 0.14, 1.88, 2.76, 0.78, [t(price, 40, TEXT, True)], page=page)
        add_text(s, x + 0.16, 2.70, 2.74, 0.30, [t("元 / 人", 13, MUTED)], page=page)
        add_text(s, x + 0.16, 3.12, 2.74, 0.34, [t(quota, 16, TEXT, True)], page=page)
        add_text(s, x + 0.16, 3.48, 2.74, 0.40, [t(note, 13, GOLD_TEXT)], page=page)
        add_rect(s, x + 0.16, 4.02, 2.70, 0.012, LINE)
        add_text(s, x + 0.16, 4.16, 2.74, 1.45, [t(rights, 14, TEXT, spc=1.2)], page=page)
    add_text(
        s,
        0.52,
        6.00,
        12.3,
        0.90,
        [t("单场约 3 小时，不另设半天票。按推荐结构，平均票面约 659 元，到手约 619 元/张。\n平台手续费、税费和少量退款按 94% 计入到手；赞助到手也按 94% 计。", 14, MUTED, spc=1.15)],
        page=page,
    )


def slide_seats(prs, page, total):
    s = new_slide(prs)
    header(s, "落地方案  ·  座位", "200 个座位，先扣赠票再卖票", page, total)
    add_text(s, 0.52, 1.28, 12.3, 0.36, [t("下面按赞助 12 席画示意。实际赞助席 10—14，可售门票在 168—182 之间调整。", 14, MUTED)], page=page)

    segments = [
        (12, CINNABAR, "赞助"),
        (12, GOLD, "学员"),
        (6, PINE, "嘉宾"),
        (170, INK, "可售"),
    ]
    bar_x, bar_y, bar_w, bar_h = 0.52, 1.82, 12.30, 0.62
    cursor = bar_x
    for seats, color, _label in segments:
        seg_w = bar_w * seats / 200
        add_rect(s, cursor, bar_y, seg_w, bar_h, color)
        if seg_w > 0.7:
            add_text(
                s,
                cursor,
                bar_y,
                seg_w,
                bar_h,
                [t(str(seats), 16, CREAM, True, PP_ALIGN.CENTER)],
                anchor=MSO_ANCHOR.MIDDLE,
                page=page,
            )
        cursor += seg_w

    blocks = [
        (CINNABAR, "10—14 席", "赞助企业", "随档位赠送，占用容量。\n不与门票收入重复计算。"),
        (GOLD, "12 席", "此前 2,399 元学员", "若已到账，安排第一排。\n并保留当时承诺的权益。"),
        (PINE, "约 6 席", "嘉宾、主办、媒体", "控制在 6 席以内。\n不挤占可售门票。"),
        (INK, "168—182 席", "可售门票", "按赞助签约动态调整。\n不卖站票。"),
    ]
    for i, (color, num, title, desc) in enumerate(blocks):
        x = 0.52 + (i % 4) * 3.20
        y = 2.72
        add_round(s, x, y, 3.06, 2.55, CARD, LINE, 0.08)
        add_rect(s, x, y, 3.06, 0.08, color)
        add_text(s, x + 0.16, y + 0.22, 2.74, 0.46, [t(num, 22, TEXT, True)], page=page)
        add_text(s, x + 0.16, y + 0.72, 2.74, 0.36, [t(title, 14, CINNABAR, True)], page=page)
        add_text(s, x + 0.16, y + 1.16, 2.74, 1.18, [t(desc, 13, TEXT, spc=1.12)], page=page)
    add_round(s, 0.52, 5.48, 12.30, 1.42, GOLD_WASH, radius=0.08)
    add_text(
        s,
        0.74,
        5.64,
        11.9,
        1.12,
        [t("2,399 元学员是否占这 12 席，只看收款记录。未到账不预留，也不计入结余。\n到账的话，在收支表之外再多约 2.7 万元。嘉宾席含主办方和媒体，不要临时加座。", 15, TEXT, spc=1.2)],
        anchor=MSO_ANCHOR.MIDDLE,
        page=page,
    )


def slide_math(prs, page, total):
    s = new_slide(prs)
    header(s, "落地方案  ·  收支", "赞助一到账，保本线立刻下降", page, total)
    heroes = [
        ("162", "张", "无赞助时保本", CINNABAR),
        ("71", "张", "赞助 6 万到账后", GOLD_TEXT),
        ("41", "张", "赞助 8 万到账后", PINE),
    ]
    for i, (num, unit, label, color) in enumerate(heroes):
        x = 0.52 + i * 4.16
        add_round(s, x, 1.26, 4.00, 1.72, CARD, LINE, 0.08)
        add_text(s, x + 0.2, 1.38, 2.5, 0.78, [t(num, 40, color, True)], page=page)
        add_text(s, x + 2.35, 1.62, 1.4, 0.42, [t(unit, 16, MUTED)], anchor=MSO_ANCHOR.MIDDLE, page=page)
        add_text(s, x + 0.2, 2.22, 3.6, 0.48, [t(label, 15, TEXT, True)], page=page)

    headers = ["情景", "保本需卖", "卖 120 张结余", "卖 150 张结余"]
    rows = [
        [
            cell("无赞助", bold=True),
            cell("162 张", bold=True, align=PP_ALIGN.CENTER),
            cell("-2.57 万", CINNABAR, True, PP_ALIGN.CENTER),
            cell("-0.71 万", CINNABAR, True, PP_ALIGN.CENTER),
        ],
        [
            cell("赞助 6 万到账", bold=True, bg=PINE_SOFT),
            cell("71 张", bold=True, align=PP_ALIGN.CENTER, bg=PINE_SOFT),
            cell("+3.07 万", PINE, True, PP_ALIGN.CENTER, PINE_SOFT),
            cell("+4.93 万", PINE, True, PP_ALIGN.CENTER, PINE_SOFT),
        ],
        [
            cell("赞助 8 万到账", bold=True, bg=GOLD_WASH),
            cell("41 张", bold=True, align=PP_ALIGN.CENTER, bg=GOLD_WASH),
            cell("+4.95 万", PINE, True, PP_ALIGN.CENTER, GOLD_WASH),
            cell("+6.81 万", PINE, True, PP_ALIGN.CENTER, GOLD_WASH),
        ],
    ]
    paint_table(s, 0.52, 3.18, [3.15, 3.05, 3.05, 3.05], headers, rows, page, row_h=0.58, header_h=0.46, size=15)
    add_round(s, 0.52, 5.50, 12.30, 1.40, WASH, radius=0.08)
    add_text(
        s,
        0.74,
        5.64,
        11.9,
        1.14,
        [t("讲师费按 10 万元测算。平均票面约 659 元，到手约 619 元/张；赞助到手同样按 94%。\n此前 12 位学员若已按 2,399 元到账，结余再多约 2.7 万元。未到账不计入。", 15, TEXT, spc=1.2)],
        anchor=MSO_ANCHOR.MIDDLE,
        page=page,
    )


def slide_fee(prs, page, total):
    s = new_slide(prs)
    header(s, "落地方案  ·  讲师费", "讲师费变了，就换一种打法", page, total)
    rows = [
        ("分成，或 ≤ 6 万", "票价可降为统一 499 元，早鸟 399 元。赞助记为利润。", False),
        ("6—10 万", "按本方案执行。这是当前测算所在的区间。", True),
        ("10—12 万", "方案不变，但赞助至少 6 万到账后，再全面推广。", False),
        ("12—18 万", "赞助至少 8 万签约并且到账之后，才开售。", False),
        ("> 18 万，或单场 2,500 元", "不做公开售票。改为分成，或我方只做宣传支持。", False),
    ]
    for i, (band, action, current) in enumerate(rows):
        y = 1.26 + i * 1.08
        bg = WASH if current else CARD
        add_round(s, 0.52, y, 12.30, 0.98, bg, LINE, 0.08)
        add_text(s, 0.74, y + 0.14, 4.3, 0.70, [t(band, 16, CINNABAR if current else TEXT, True)], anchor=MSO_ANCHOR.MIDDLE, page=page)
        add_text(s, 5.15, y + 0.14, 7.4, 0.70, [t(action, 15, TEXT)], anchor=MSO_ANCHOR.MIDDLE, page=page)
    add_text(s, 0.52, 6.68, 12.3, 0.30, [t("对照的是书面讲师费，不是口头报价。还没拿到数字之前，对外仍按本方案筹备，但不付清全款。", 13, MUTED)], page=page)


def slide_roles(prs, page, total):
    s = new_slide(prs)
    header(s, "落地方案  ·  分工", "对外签字只留在 A", page, total)
    people = [
        ("A  总负责", "签约与拍板", ["讲师合同与付款", "对接首乾书院", "三家名义授权", "赞助谈判与签约"], "赞助进展：意向、签约、到账"),
        ("B  内容传播", "卖点与露出", ["海报定稿", "公众号与朋友圈", "倒计时和社群", "赞助方露出物料"], "当天发布内容与阅读数据"),
        ("C  票务客服", "报名与名单", ["小鹅通或活动行", "收款、开票、退款", "群内答疑", "名单与座位表"], "每天 20:00 销售日报"),
        ("D  渠道现场", "企业与现场", ["会员企业逐家邀约", "团体票和赞助线索", "场地、设备、茶歇", "签到与动线"], "当天拜访企业数与线索"),
    ]
    for i, (name, role, lines, daily) in enumerate(people):
        x = 0.52 + i * 3.20
        add_round(s, x, 1.26, 3.06, 4.48, CARD, LINE, 0.08)
        add_text(s, x + 0.16, 1.40, 2.74, 0.40, [t(name, 16, CINNABAR, True)], page=page)
        add_text(s, x + 0.16, 1.82, 2.74, 0.34, [t(role, 13, GOLD_TEXT, True)], page=page)
        add_text(s, x + 0.16, 2.30, 2.74, 2.00, [t("\n".join(lines), 14, TEXT, spc=1.28)], page=page)
        add_rect(s, x + 0.16, 4.40, 2.74, 0.012, LINE)
        add_text(s, x + 0.16, 4.50, 2.74, 0.26, [t("每天交出", 12, MUTED, True)], page=page)
        add_text(s, x + 0.16, 4.78, 2.74, 0.74, [t(daily, 13, TEXT, True, spc=1.05)], page=page)
    add_round(s, 0.52, 5.90, 12.30, 1.02, INK, radius=0.08)
    add_text(
        s,
        0.74,
        6.02,
        11.9,
        0.78,
        [t("讲师、授权、收款、赞助签约，全部由 A 对外签字。\nB 管传播，C 管票务，D 管渠道和现场。", 16, CREAM, True, spc=1.12)],
        anchor=MSO_ANCHOR.MIDDLE,
        page=page,
    )


def _schedule(prs, page, total, kicker, title, rows, note):
    s = new_slide(prs)
    header(s, kicker, title, page, total)
    widths = [2.05, 3.45, 3.45, 3.35]
    headers = ["日期", "门票线", "赞助线", "检查点"]
    parsed = []
    for date, ticket, sponsor, check, kind in rows:
        bg = WASH if kind == "alert" else (GOLD_WASH if kind == "go" else None)
        weight = kind in ("alert", "go")
        parsed.append(
            [
                cell(date, CINNABAR if weight else TEXT, True, PP_ALIGN.LEFT, bg),
                cell(ticket, TEXT, weight, PP_ALIGN.LEFT, bg),
                cell(sponsor, TEXT, weight, PP_ALIGN.LEFT, bg),
                cell(check, TEXT, weight, PP_ALIGN.LEFT, bg),
            ]
        )
    row_h = 0.70 if len(rows) >= 7 else 0.78
    end = paint_table(s, 0.52, 1.26, widths, headers, parsed, page, row_h=row_h, header_h=0.42, size=13)
    add_text(s, 0.52, max(end + 0.12, 6.55), 12.3, 0.42, [t(note, 13, MUTED)], page=page)


def slide_schedule_a(prs, page, total):
    rows = [
        ("10/10—12", "改海报，搭报名页\n删半天票，改成第一讲/第二讲", "列出 20 家目标企业\n发出赞助邀请", "四份书面确认齐全", ""),
        ("10/13", "开售：公众号、朋友圈\n三方社群同步", "首轮电话或面谈", "首日销量 ≥ 30 张", "go"),
        ("10/14—15", "推进早鸟", "确定主要意向", "至少 1 家战略，或 3 家特别", ""),
        ("10/16—17", "第二波推文，讲讲座看点", "发出合同\n核对开票信息", "门票 ≥ 49 张", ""),
        ("10/18—20", "早鸟倒计时", "签约并催款到账", "门票 ≥ 98；到账 ≥ 3 万", ""),
        ("10/21", "切换为 699 元标准价", "收集 LOGO、简介\n和参会名单", "物料素材齐", ""),
        ("10/22", "是否继续", "按到账金额重算保本", "看决策页", "alert"),
    ]
    _schedule(
        prs,
        page,
        total,
        "落地方案  ·  时间表",
        "10 月 22 日之前，门票和赞助并行走",
        rows,
        "意向书或微信书面确认算意向；没有到账，就不算赞助收入。",
    )


def slide_schedule_b(prs, page, total):
    rows = [
        ("10/23—25", "企业团体票冲刺", "赞助物料定稿", "背景板、易拉宝、鸣谢页", ""),
        ("10/26—28", "最后一波推广\n发参会须知", "物料下单制作", "达到重算后的保本线", ""),
        ("10/29", "停售，锁定名单", "锁定赞助席位名单", "座位表定稿", ""),
        ("10/30", "场地彩排", "物料进场布置", "设备全部试一遍", ""),
        ("10/31", "活动当天", "现场兑现权益并拍照", "按当天流程", "go"),
        ("11/1—7", "讲师尾款、开票、复盘", "发出赞助结案报告", "11 月 7 日前完成", "alert"),
    ]
    _schedule(
        prs,
        page,
        total,
        "落地方案  ·  时间表",
        "决策日之后，一直做到结案",
        rows,
        "10 月 20 日之后才签约的赞助，背景板能否加 LOGO，看制作进度。",
    )


def slide_go(prs, page, total):
    s = new_slide(prs)
    header(s, "落地方案  ·  决策", "只拿已到账的赞助重算保本", page, total)
    add_round(s, 0.52, 1.26, 12.30, 1.20, INK, radius=0.08)
    add_text(
        s,
        0.74,
        1.40,
        11.9,
        0.92,
        [t("保本张数  =  （讲师费 − 已到账赞助 × 94%）÷ 619\n口头意向不进入这个公式。", 20, CREAM, True, spc=1.15)],
        anchor=MSO_ANCHOR.MIDDLE,
        page=page,
    )
    bands = [
        (PINE_SOFT, PINE, "≥ 保本张数的 60%", "正常推进", "维持现有票价和推广节奏。"),
        (GOLD_WASH, GOLD_TEXT, "保本张数的 50%—60%", "加推，并继续追赞助", "C、D 加推企业团体票，A 追还没到账的赞助。"),
        (WASH, CINNABAR, "低于保本张数的 50%", "改条件，否则取消", "与讲师方改成分成或降低讲师费。谈不成就取消，按规则全额退款。"),
    ]
    for i, (bg, color, title, action, detail) in enumerate(bands):
        y = 2.66 + i * 1.38
        add_round(s, 0.52, y, 12.30, 1.26, bg, radius=0.08)
        add_text(s, 0.74, y + 0.16, 5.3, 0.94, [t(title, 18, color, True)], anchor=MSO_ANCHOR.MIDDLE, page=page)
        add_text(s, 6.15, y + 0.14, 6.4, 0.42, [t(action, 16, TEXT, True)], page=page)
        add_text(s, 6.15, y + 0.58, 6.4, 0.52, [t(detail, 14, TEXT)], page=page)


def slide_day(prs, page, total):
    s = new_slide(prs)
    header(s, "落地方案  ·  当天", "三小时怎么过", page, total)
    rows = [
        ("07:30", "全员到场，检查音响、投影、话筒、签到台、背景板", "D"),
        ("08:00", "赞助企业接待区就位", "A"),
        ("08:15", "签到核销，发讲义，引导入座", "C、D"),
        ("09:00", "主持开场，介绍主办单位，鸣谢赞助单位", "A"),
        ("09:05", "第一讲：走进王阳明——心即理、致良知", "讲座"),
        ("10:20", "茶歇 20 分钟，赞助企业展示与交流", "D"),
        ("10:40", "第二讲：知行合一——在不确定的时代践行心学", "讲座"),
        ("11:40", "现场问答。问题提前在群里征集，主持人筛选", "A"),
        ("12:00", "合影以讲师同意为准，引导离场，扫码入群", "B、C"),
    ]
    for i, (when, what, who) in enumerate(rows):
        y = 1.24 + i * 0.62
        bg = GOLD_WASH if what.startswith("第一讲") or what.startswith("第二讲") else (CARD if i % 2 == 0 else PAPER_2)
        add_rect(s, 0.52, y, 12.30, 0.58, bg)
        add_text(s, 0.62, y, 1.35, 0.58, [t(when, 14, CINNABAR, True)], anchor=MSO_ANCHOR.MIDDLE, page=page)
        add_text(s, 2.05, y, 8.9, 0.58, [t(what, 14, TEXT, True if what.startswith("第") else False)], anchor=MSO_ANCHOR.MIDDLE, page=page)
        add_text(s, 11.05, y, 1.6, 0.58, [t(who, 14, GOLD_TEXT, True, PP_ALIGN.RIGHT)], anchor=MSO_ANCHOR.MIDDLE, page=page, pad=0.04)


def slide_risk(prs, page, total):
    s = new_slide(prs)
    header(s, "落地方案  ·  规则", "退费写清楚，风险有人接", page, total)
    add_text(s, 0.52, 1.24, 6.0, 0.32, [t("写在报名页上的退费", 14, CINNABAR, True)], page=page)
    refunds = [
        ("10/24  24:00 前", "全额退款"),
        ("10/25—10/29", "退 80%，或免费转让"),
        ("10/30 起", "不退款，仍可转让"),
        ("讲师或主办方取消、改期", "全额退款"),
    ]
    for i, (when, rule) in enumerate(refunds):
        y = 1.64 + i * 1.22
        add_round(s, 0.52, y, 6.05, 1.10, CARD, LINE, 0.08)
        add_text(s, 0.70, y + 0.12, 5.7, 0.36, [t(when, 13, GOLD_TEXT, True)], page=page)
        add_text(s, 0.70, y + 0.48, 5.7, 0.46, [t(rule, 18, TEXT, True)], page=page)
    add_text(s, 6.80, 1.24, 6.0, 0.32, [t("事先写明的风险", 14, CINNABAR, True)], page=page)
    risks = [
        ("讲师临时取消", "合同写明已付款项全额退还"),
        ("复旦名义未获校内同意", "改列学术支持单位，不收费"),
        ("销量不足", "10/22 处理，不硬撑"),
        ("赞助谈好但未到账", "只按到账金额做支出"),
        ("未确认权益被写进宣传", "签名书、合影、问答确认前不写"),
        ("现场超员", "按核定人数售票，不卖站票"),
        ("与书院主课价格冲突", "称公开讲座；群内不报主课价"),
    ]
    for i, (name, action) in enumerate(risks):
        y = 1.64 + i * 0.70
        add_round(s, 6.80, y, 6.02, 0.64, CARD, LINE, 0.06)
        add_text(s, 6.94, y, 5.74, 0.30, [t(name, 12, TEXT, True)], anchor=MSO_ANCHOR.MIDDLE, page=page)
        add_text(s, 6.94, y + 0.28, 5.74, 0.32, [t(action, 12, MUTED)], anchor=MSO_ANCHOR.MIDDLE, page=page)


def slide_parties(prs, page, total):
    s = new_slide(prs)
    header(s, "价值实施  ·  四方", "每一方要什么、交什么、如何证明", page, total)
    headers = ["对象", "他们要什么", "我们交付什么", "怎么证明"]
    rows = [
        [
            cell("听众", bold=True),
            cell("值回票价，现场听到网上没有的内容"),
            cell("3 小时讲座、茶歇、讲义、问答、交流群"),
            cell("满意度、到场率、入群率"),
        ],
        [
            cell("赞助企业", bold=True),
            cell("品牌曝光、接待客户、高管学习"),
            cell("展示、席位、鸣谢、推文、结案报告"),
            cell("照片、阅读量、到场人数"),
        ],
        [
            cell("主办机构", bold=True),
            cell("影响力、会员服务、企业资源"),
            cell("联合露出、会员团体价、名单沉淀"),
            cell("参会企业名单与合作线索"),
        ],
        [
            cell("讲师与书院", bold=True),
            cell("讲师费、影响力、口碑"),
            cell("按时付款、规范组织、专业现场"),
            cell("按期履约、现场秩序与反馈"),
        ],
    ]
    paint_table(s, 0.52, 1.30, [2.15, 3.45, 3.85, 2.85], headers, rows, page, row_h=1.05, header_h=0.46, size=13)
    add_text(
        s,
        0.52,
        6.15,
        12.3,
        0.75,
        [t("价值要能被指出来。没有照片、名单或问卷的权益，结案时不算兑现。\n录音录像不作为交付，除非讲师书面同意。", 14, MUTED, spc=1.15)],
        page=page,
    )


def slide_audience(prs, page, total):
    s = new_slide(prs)
    header(s, "价值实施  ·  听众", "699 元要让人带走五样东西", page, total)
    rows = [
        ("听得懂", "会前 3 天发 1 页导读：心即理、致良知、知行合一", "B", "10/28"),
        ("带得走", "纸质讲义：提纲、名句摘录、笔记页。内容须讲师确认", "B、D", "10/28 定稿，10/29 印"),
        ("问得到", "群里征集问题，主持人筛选 5—8 个，现场问答 20 分钟", "A、C", "10/26 起"),
        ("有交流", "茶歇 20 分钟，分科技、金融、制造、教育四张交流桌", "D", "10/30 布置"),
        ("有后续", "会后交流群。3 天内发要点；无书面同意不发录音录像", "B、C", "11/3 前"),
    ]
    widths = [1.70, 6.15, 1.35, 3.10]
    headers = ["价值", "具体做法", "负责", "时间"]
    parsed = []
    for name, how, who, when in rows:
        parsed.append(
            [
                cell(name, CINNABAR, True),
                cell(how),
                cell(who, align=PP_ALIGN.CENTER, bold=True),
                cell(when),
            ]
        )
    paint_table(s, 0.52, 1.28, widths, headers, parsed, page, row_h=0.92, header_h=0.44, size=14)
    add_text(s, 0.52, 6.50, 12.3, 0.42, [t("一场听完就走的讲座撑不起 699 元。讲义内容未确认前，海报上不写具体篇目。", 14, MUTED)], page=page)


def slide_hosts(prs, page, total):
    s = new_slide(prs)
    header(s, "价值实施  ·  主办", "三家主办，各做各的事", page, total)
    hosts = [
        ("复旦大学住房政策研究中心", "学术主办，或学术支持", ["以学术主办或学术支持出现", "体现城市与人文的交叉研究", "不承担收费责任", "校内未同意则改列支持单位"]),
        ("上海市杨浦区科技企业联合会", "会员入口", ["会员享团体价 599 元/人", "赞助优先从会员企业开发", "作为年度会员服务活动", "沉淀参会企业与合作线索"]),
        ("首乾书院", "讲师资源方", ["提供讲师，获得讲师费", "同时获得品牌曝光", "不在交流群里报主课价格", "主课与本场公开讲座分开说"]),
    ]
    for i, (name, role, lines) in enumerate(hosts):
        x = 0.52 + i * 4.16
        add_round(s, x, 1.26, 4.00, 4.35, CARD, LINE, 0.08)
        add_text(s, x + 0.18, 1.42, 3.64, 0.90, [t(name, 16, TEXT, True, spc=1.05)], page=page)
        add_text(s, x + 0.18, 2.36, 3.64, 0.36, [t(role, 14, CINNABAR, True)], page=page)
        add_text(s, x + 0.18, 2.90, 3.64, 2.35, [t("\n".join(lines), 14, TEXT, spc=1.35)], page=page)
    add_round(s, 0.52, 5.78, 12.30, 1.12, INK, radius=0.08)
    add_text(
        s,
        0.74,
        5.90,
        11.9,
        0.88,
        [t("交流群由我方建群并担任群主，首乾书院担任管理员。\n群内不报主课价格，对外只称公开讲座，不称体验课。", 16, CREAM, True, spc=1.12)],
        anchor=MSO_ANCHOR.MIDDLE,
        page=page,
    )


def slide_kpi(prs, page, total):
    s = new_slide(prs)
    header(s, "价值实施  ·  复盘", "结束后一周，用八个数复盘", page, total)
    kpis = [
        ("到场率", "≥ 90%", "签到核销"),
        ("满意度", "≥ 4.5 / 5", "会后扫码，3 道题"),
        ("愿意再参加", "≥ 70%", "同一份问卷"),
        ("入群率", "≥ 80%", "入群人数 ÷ 到场人数"),
        ("开售推文", "≥ 5,000", "公众号阅读"),
        ("企业团体票", "≥ 20 张", "票务系统"),
        ("赞助续约意向", "≥ 50%", "结案回访"),
        ("新增合作线索", "≥ 10 家", "拜访记录与现场交流"),
    ]
    for i, (name, value, source) in enumerate(kpis):
        col, row = i % 4, i // 4
        x = 0.52 + col * 3.20
        y = 1.30 + row * 2.55
        add_round(s, x, y, 3.06, 2.38, CARD, LINE, 0.08)
        add_text(s, x + 0.16, y + 0.16, 2.74, 0.36, [t(name, 14, MUTED, True)], page=page)
        add_text(s, x + 0.16, y + 0.58, 2.74, 0.78, [t(value, 26, CINNABAR, True)], page=page)
        add_text(s, x + 0.16, y + 1.55, 2.74, 0.58, [t(source, 14, TEXT)], page=page)


def slide_position(prs, page, total):
    s = new_slide(prs)
    header(s, "赞助落地  ·  定位", "卖的是一次完整合作", page, total)
    add_round(s, 0.52, 1.26, 12.30, 1.45, INK, radius=0.08)
    add_text(
        s,
        0.74,
        1.40,
        11.9,
        1.18,
        [t("不卖 LOGO 位置。卖正式席位、现场展示和活动传播。\n企业可以派管理层来学习，也可以邀请重要客户。不设 5—10 万元总冠名。", 16, CREAM, spc=1.2)],
        anchor=MSO_ANCHOR.MIDDLE,
        page=page,
    )
    tiers = [
        ("战略支持", "3 万", "限 1 家", "4 个前排席位\n背景板单独大标\n开场单独鸣谢", CINNABAR),
        ("特别支持", "1 万", "2—3 家", "2 个讲座席位\n背景板中等标识\n合并鸣谢", GOLD_TEXT),
        ("品牌支持", "5,000", "3—5 家", "1 个讲座席位\n背景板小标并列\n合并鸣谢", PINE),
    ]
    for i, (name, price, quota, body, color) in enumerate(tiers):
        x = 0.52 + i * 4.16
        add_round(s, x, 2.92, 4.00, 3.00, CARD, LINE, 0.08)
        add_text(s, x + 0.2, 3.06, 3.6, 0.32, [t(name, 14, color, True)], page=page)
        add_text(s, x + 0.2, 3.40, 3.6, 0.70, [t(price, 36, TEXT, True)], page=page)
        add_text(s, x + 0.2, 4.14, 3.6, 0.32, [t("元    ·    " + quota, 14, MUTED)], page=page)
        add_text(s, x + 0.2, 4.56, 3.6, 1.15, [t(body, 14, TEXT, spc=1.15)], page=page)
    add_text(s, 0.52, 6.08, 12.3, 0.82, [t("门槛定低、家数做多，是为了在 10 天内签下来。现金目标：6 万保底，8 万争取。\n实物支持另计，不进现金目标。", 14, MUTED, spc=1.15)], page=page)


def slide_benefits(prs, page, total):
    s = new_slide(prs)
    header(s, "赞助落地  ·  权益", "权益逐项写死，不靠口头", page, total)
    headers = ["权益", "战略支持 · 3 万", "特别支持 · 1 万", "品牌支持 · 5,000"]
    wash = WASH
    rows_src = [
        ("数量", "限 1 家", "2—3 家", "3—5 家"),
        ("讲座席位", "4 席，安排前排", "2 席", "1 席"),
        ("主背景板", "大尺寸，单独位置", "中尺寸", "小尺寸，并列"),
        ("开场鸣谢", "单独鸣谢", "合并鸣谢", "合并鸣谢"),
        ("PPT 鸣谢页", "单独一页", "合并一页", "合并一页"),
        ("公众号", "开售＋会后，100 字简介", "会后 LOGO", "会后 LOGO"),
        ("现场展示", "易拉宝 2 个＋茶歇展台", "易拉宝 1 个", "资料放入签到袋"),
        ("茶歇交流", "负责人 3 分钟介绍", "—", "—"),
        ("结案报告", "有", "有", "有"),
    ]
    rows = []
    for name, a, b, c in rows_src:
        rows.append(
            [
                cell(name, bold=True),
                cell(a, bold=True, align=PP_ALIGN.CENTER, bg=wash),
                cell(b, align=PP_ALIGN.CENTER),
                cell(c, align=PP_ALIGN.CENTER),
            ]
        )
    paint_table(
        s,
        0.52,
        1.24,
        [2.20, 3.50, 3.30, 3.30],
        headers,
        rows,
        page,
        row_h=0.50,
        header_h=0.44,
        size=13,
        header_colors=[INK, CINNABAR, INK, INK],
    )
    add_text(
        s,
        0.52,
        6.28,
        12.3,
        0.70,
        [t("席位占用容量，不重复计票房。合影、签名书、闭门交流不是赞助权益，除非讲师书面同意。\n茶歇 3 分钟介绍也要讲师和主办方同意。位置、尺寸、数量写入协议。", 13, MUTED, spc=1.12)],
        page=page,
    )


def slide_targets(prs, page, total):
    s = new_slide(prs)
    header(s, "赞助落地  ·  目标", "保底 6 万，争取 8 万", page, total)
    panels = [
        ("保底", "6 万元", "10 个赠送席位", [("战略支持", "1 家", "3 万"), ("特别支持", "2 家", "2 万"), ("品牌支持", "2 家", "1 万")], CINNABAR),
        ("争取", "8 万元", "14 个赠送席位", [("战略支持", "1 家", "3 万"), ("特别支持", "3 家", "3 万"), ("品牌支持", "4 家", "2 万")], PINE),
    ]
    for i, (name, total_cash, seats, lines, color) in enumerate(panels):
        x = 0.52 + i * 6.40
        add_round(s, x, 1.26, 6.20, 3.85, CARD, LINE, 0.08)
        add_text(s, x + 0.24, 1.40, 5.7, 0.30, [t(name, 14, color, True)], page=page)
        add_text(s, x + 0.24, 1.72, 5.7, 0.62, [t(total_cash, 32, TEXT, True)], page=page)
        add_text(s, x + 0.24, 2.40, 5.7, 0.32, [t(seats, 14, MUTED)], page=page)
        for j, (tier, count, money) in enumerate(lines):
            yy = 2.90 + j * 0.64
            add_rect(s, x + 0.24, yy, 5.72, 0.56, PAPER_2 if j % 2 else PAPER)
            add_text(s, x + 0.36, yy, 2.4, 0.56, [t(tier, 14, TEXT)], anchor=MSO_ANCHOR.MIDDLE, page=page)
            add_text(s, x + 2.7, yy, 1.4, 0.56, [t(count, 14, TEXT, True, PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.MIDDLE, page=page)
            add_text(s, x + 4.1, yy, 1.7, 0.56, [t(money, 14, color, True, PP_ALIGN.RIGHT)], anchor=MSO_ANCHOR.MIDDLE, page=page)
    add_round(s, 0.52, 5.28, 12.30, 1.62, GOLD_WASH, radius=0.08)
    add_text(s, 0.74, 5.42, 11.9, 0.36, [t("实物支持，不计入现金", 14, GOLD_TEXT, True)], page=page)
    add_text(
        s,
        0.74,
        5.82,
        11.9,
        0.88,
        [t("茶饮、茶歇点心、讲义印刷、伴手礼。提供实物的企业列为“特别鸣谢”，只做口头鸣谢和 PPT 鸣谢。", 15, TEXT, spc=1.15)],
        page=page,
    )


def slide_companies(prs, page, total):
    s = new_slide(prs)
    header(s, "赞助落地  ·  企业", "先找会员企业，再按行业开口", page, total)
    rows = [
        ("1", "科技、人工智能企业", "联合会会员优先", "1—3 万", "企业家成长，科技与人文，高管学习席位"),
        ("2", "企业服务、咨询、律所、会计所", "听众就是客户", "1—3 万", "现场这批人，就是他们的潜在客户"),
        ("3", "金融、财富管理、私人银行", "客户维护", "1—3 万", "用来邀请和维护高净值客户"),
        ("4", "高端茶叶、文化消费品牌", "主题契合", "5,000—1 万", "与心学主题契合，也可以兼做茶歇"),
        ("5", "出版、书店、文创", "实物或小额", "实物或 5,000", "提供图书和文创伴手礼"),
    ]
    for i, (num, industry, tag, money, hook) in enumerate(rows):
        y = 1.26 + i * 1.08
        add_round(s, 0.52, y, 12.30, 0.98, CARD, LINE, 0.08)
        add_round(s, 0.70, y + 0.22, 0.54, 0.54, CINNABAR, radius=0.2)
        add_text(s, 0.70, y + 0.22, 0.54, 0.54, [t(num, 16, CREAM, True, PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.MIDDLE, page=page, pad=0.0)
        add_text(s, 1.42, y + 0.12, 6.3, 0.40, [t(industry, 16, TEXT, True)], page=page)
        add_text(s, 1.42, y + 0.52, 6.3, 0.34, [t(tag, 13, MUTED)], page=page)
        add_text(s, 7.85, y + 0.12, 4.7, 0.36, [t(money, 16, CINNABAR, True, PP_ALIGN.RIGHT)], page=page)
        add_text(s, 7.85, y + 0.50, 4.7, 0.36, [t(hook, 13, TEXT, align=PP_ALIGN.RIGHT)], page=page)


def slide_steps(prs, page, total):
    s = new_slide(prs)
    header(s, "赞助落地  ·  步骤", "A 主谈，D 配合，10 月 20 日前到账", page, total)
    steps = [
        ("10/10—12", "列出 20 家，发出一页纸", "首轮触达 20 家"),
        ("10/13—15", "面谈，重点攻 1 家战略、3 家特别", "确定主要意向"),
        ("10/16—20", "发协议、核对开票、催款", "签约并且到账"),
        ("10/21—25", "收矢量 LOGO、简介、参会名单", "物料齐全"),
        ("10/26—30", "做背景板、易拉宝、鸣谢页，进场拍照", "赞助方确认布置"),
        ("10/31", "现场兑现权益，专人拍照留证", "现场照片"),
        ("11/1—7", "发结案报告，回访，谈下一场", "续约意向"),
    ]
    for i, (when, action, output) in enumerate(steps):
        y = 1.24 + i * 0.72
        add_round(s, 0.52, y, 1.85, 0.64, INK if i == 2 else CARD, radius=0.08)
        add_text(
            s,
            0.52,
            y,
            1.85,
            0.64,
            [t(when, 12, CREAM if i == 2 else CINNABAR, True, PP_ALIGN.CENTER)],
            anchor=MSO_ANCHOR.MIDDLE,
            page=page,
            pad=0.04,
        )
        add_round(s, 2.50, y, 6.55, 0.64, CARD, LINE, 0.08)
        add_text(s, 2.64, y, 6.28, 0.64, [t(action, 14, TEXT)], anchor=MSO_ANCHOR.MIDDLE, page=page)
        add_round(s, 9.18, y, 3.64, 0.64, GOLD_WASH if i == 2 else PAPER_2, radius=0.08)
        add_text(s, 9.28, y, 3.44, 0.64, [t(output, 13, TEXT, True)], anchor=MSO_ANCHOR.MIDDLE, page=page)
    add_text(s, 0.52, 6.40, 12.3, 0.55, [t("10 月 20 日之后签约，只能保证口头鸣谢、PPT 鸣谢和公众号露出。背景板能否加标识，视制作进度。\n物料截止 10 月 25 日，逾期视为放弃相应权益。", 13, MUTED, spc=1.1)], page=page)


def slide_contract(prs, page, total):
    s = new_slide(prs)
    header(s, "赞助落地  ·  协议", "协议写六条，账分四笔", page, total)
    clauses = [
        ("1", "金额、10 月 20 日前付款，以及收款和开票主体。"),
        ("2", "权益逐项列明：席位、LOGO 位置与尺寸、鸣谢、推文次数。"),
        ("3", "物料截止 10 月 25 日，逾期视为放弃相应权益。"),
        ("4", "名称、标识和肖像以正式授权为准；不得单独用讲师肖像做宣传。"),
        ("5", "活动取消全额退款；改期可保留权益，或全额退款。"),
        ("6", "现场不得销售产品，也不得派发与活动无关的广告。"),
    ]
    for i, (num, text) in enumerate(clauses):
        y = 1.26 + i * 0.88
        add_round(s, 0.52, y, 7.55, 0.80, CARD, LINE, 0.08)
        add_text(s, 0.66, y, 0.50, 0.80, [t(num, 18, CINNABAR, True)], anchor=MSO_ANCHOR.MIDDLE, page=page)
        add_text(s, 1.20, y, 6.70, 0.80, [t(text, 13, TEXT)], anchor=MSO_ANCHOR.MIDDLE, page=page)
    books = [
        ("分开记账", "赞助、门票、其他收入各记各的"),
        ("席位另册", "赞助席位不计入门票销售张数"),
        ("只认到账", "口头意向不进入保本线"),
        ("不垫讲师费", "赞助款不用于提高讲师费"),
    ]
    add_text(s, 8.28, 1.26, 4.5, 0.36, [t("记账原则", 14, CINNABAR, True)], page=page)
    for i, (title, desc) in enumerate(books):
        y = 1.70 + i * 1.22
        add_round(s, 8.28, y, 4.54, 1.10, CARD, LINE, 0.08)
        add_text(s, 8.44, y + 0.10, 4.22, 0.36, [t(title, 15, TEXT, True)], page=page)
        add_text(s, 8.44, y + 0.48, 4.22, 0.50, [t(desc, 13, MUTED)], page=page)


def slide_report(prs, page, total):
    s = new_slide(prs)
    header(s, "赞助落地  ·  结案", "11 月 7 日前，每家一份结案报告", page, total)
    items = [
        ("01", "活动概况", "到场人数，以及按报名信息统计的听众行业构成。不泄露个人信息。"),
        ("02", "权益兑现清单", "逐项打勾，附背景板、易拉宝、鸣谢页和席位的现场照片。"),
        ("03", "传播数据", "公众号推文阅读量和转发数。"),
        ("04", "听众反馈", "满意度摘要，来自会后 3 题问卷。"),
        ("05", "下一场", "附上下一场活动的合作邀请，并记录续约意向。"),
    ]
    for i, (num, title, desc) in enumerate(items):
        y = 1.26 + i * 1.02
        add_round(s, 0.52, y, 12.30, 0.92, CARD, LINE, 0.08)
        add_text(s, 0.72, y, 0.80, 0.92, [t(num, 18, CINNABAR, True)], anchor=MSO_ANCHOR.MIDDLE, page=page)
        add_text(s, 1.60, y + 0.10, 2.5, 0.72, [t(title, 16, TEXT, True)], anchor=MSO_ANCHOR.MIDDLE, page=page)
        add_text(s, 4.20, y + 0.12, 8.3, 0.68, [t(desc, 14, TEXT)], anchor=MSO_ANCHOR.MIDDLE, page=page)


def slide_copy(prs, page, total):
    s = new_slide(prs)
    header(s, "附录  ·  可用文字", "这三段文字可以直接发出", page, total)
    add_text(s, 0.52, 1.20, 12.3, 0.32, [t("未书面确认的权益不要加进去，例如签名书、合影和闭门交流。", 13, MUTED)], page=page)

    add_round(s, 0.52, 1.58, 12.30, 1.28, CARD, LINE, 0.08)
    add_text(s, 0.70, 1.66, 12.0, 0.28, [t("海报票务区", 13, CINNABAR, True)], page=page)
    add_text(
        s,
        0.70,
        1.96,
        11.95,
        0.78,
        [t("标准席  699 元/人  ｜  早鸟 499 元（10 月 20 日前，限 60 席）\n企业团体 5 人起 599 元/人  ｜  前排席 999 元（限 20 席）\n2026 年 10 月 31 日（星期六）上午 09:00—12:00  ｜  扫码报名", 13, TEXT, spc=1.05)],
        page=page,
    )

    add_round(s, 0.52, 3.00, 12.30, 2.28, CARD, LINE, 0.08)
    add_text(s, 0.70, 3.08, 12.0, 0.26, [t("赞助邀请微信  ·  发给企业负责人", 13, CINNABAR, True)], page=page)
    add_text(
        s,
        0.70,
        3.36,
        11.95,
        1.82,
        [t("X 总好。10 月 31 日上午，我们联合复旦大学住房政策研究中心、首乾书院，\n在上海请王德峰教授讲《传习录》与阳明心学，约 200 位企业家和专业人士参加。\n想邀请贵司作为支持单位：3 万元含 4 个讲座席位和现场主展示，\n1 万元含 2 个席位，5,000 元含 1 个席位。席位可以给高管学习，也可以请重要客户。\n一页说明发您，方便的话这两天约 15 分钟聊一下。10 月 20 日前确定，\n能保证背景板和推文露出。", 13, TEXT, spc=1.05)],
        page=page,
    )

    add_round(s, 0.52, 5.42, 12.30, 1.48, CARD, LINE, 0.08)
    add_text(s, 0.70, 5.50, 12.0, 0.26, [t("报名页说明", 13, CINNABAR, True)], page=page)
    add_text(
        s,
        0.70,
        5.78,
        11.95,
        1.00,
        [t("本次为单场公开讲座，时间为 2026 年 10 月 31 日上午 09:00—12:00，\n分第一讲、第二讲两段，中间设茶歇交流。每位参会者获赠纸质讲义一份。\n课程安排、席位权益以本页公布为准。退费规则见下方。", 13, TEXT, spc=1.05)],
        page=page,
    )


def slide_now(prs, page, total):
    s = new_slide(prs)
    header(s, "本周动作", "不齐这四份，13 日不开售", page, total)
    gates = [
        ("1", "讲师费与付款方式", "A 向首乾书院要书面：金额、30% 定金、取消退还"),
        ("2", "场地与容量", "D 向场地方要书面：免费、座位数、彩排、设备"),
        ("3", "名义与肖像授权", "A 向三家主办和讲师方确认姓名、肖像、题目"),
        ("4", "收款与开票主体", "A 确认门票和赞助由谁收、谁开票、谁退款"),
    ]
    for i, (num, title, detail) in enumerate(gates):
        y = 1.26 + i * 1.05
        add_round(s, 0.52, y, 8.05, 0.95, CARD, LINE, 0.08)
        add_round(s, 0.68, y + 0.20, 0.54, 0.54, CINNABAR, radius=0.2)
        add_text(s, 0.68, y + 0.20, 0.54, 0.54, [t(num, 16, CREAM, True, PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.MIDDLE, page=page, pad=0.0)
        add_text(s, 1.40, y + 0.08, 6.95, 0.38, [t(title, 16, TEXT, True)], page=page)
        add_text(s, 1.40, y + 0.48, 6.95, 0.38, [t(detail, 13, MUTED)], page=page)
    add_round(s, 8.75, 1.26, 4.07, 4.10, INK, radius=0.08)
    add_text(s, 8.95, 1.42, 3.7, 0.36, [t("同时要完成", 14, GOLD_LINE, True)], page=page)
    side = [
        ("B", "删掉半天票，把“上下”\n改成第一讲、第二讲"),
        ("C", "报名页上线，写上四档票价和退费规则"),
        ("A、D", "列出 20 家目标企业，发出赞助邀请"),
    ]
    for i, (who, what) in enumerate(side):
        y = 1.95 + i * 1.05
        add_text(s, 8.95, y, 3.7, 0.30, [t(who, 14, GOLD_LINE, True)], page=page)
        add_text(s, 8.95, y + 0.30, 3.7, 0.64, [t(what, 13, CREAM, spc=1.05)], page=page)
    add_text(
        s,
        0.52,
        5.55,
        12.3,
        1.35,
        [t("讲师费、场地容量、上座率仍是假设。拿到书面数字后，回到收支页和讲师费页重算，不要沿用 10 万元和 200 座往下承诺。\n赞助未到账，不计入收入，也不要先把讲师费付清。", 14, MUTED, spc=1.2)],
        page=page,
    )


BUILDERS = [
    slide_cover,
    slide_decision,
    slide_gates,
    slide_tickets,
    slide_seats,
    slide_math,
    slide_fee,
    slide_roles,
    slide_schedule_a,
    slide_schedule_b,
    slide_go,
    slide_day,
    slide_risk,
    slide_parties,
    slide_audience,
    slide_hosts,
    slide_kpi,
    slide_position,
    slide_benefits,
    slide_targets,
    slide_companies,
    slide_steps,
    slide_contract,
    slide_report,
    slide_copy,
    slide_now,
]

REQUIRED = [
    "699",
    "499",
    "999",
    "162",
    "71",
    "41",
    "6 万",
    "8 万",
    "10 月 22 日",
    "10 月 20 日",
    "首乾书院",
    "复旦大学住房政策研究中心",
    "杨浦区科技企业联合会",
    "2,399",
    "2.7 万",
    "知行合一",
    "心即理",
    "全额退款",
    "学术支持单位",
    "结案报告",
    "小鹅通",
    "20:00",
    "4.5",
    "5,000",
    "易拉宝",
    "2,500",
    "群主",
    "94%",
    "619",
    "第一讲",
    "第二讲",
    "半天票",
    "30%",
    "不卖站票",
]


def build() -> Presentation:
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    prs.core_properties.title = "王德峰教授《传习录》与阳明心学讲座 · 落地方案"
    prs.core_properties.subject = "内部执行稿：落地方案、价值实施与赞助落地"
    prs.core_properties.author = "活动筹备组"
    prs.core_properties.category = "内部执行"
    prs.core_properties.language = "zh-CN"
    total = len(BUILDERS)
    for i, fn in enumerate(BUILDERS, start=1):
        fn(prs, i, total)
    return prs


def overflow_warnings() -> list[str]:
    warnings = []
    for box in BOXES:
        x, y, w, h = box["x"], box["y"], box["w"], box["h"]
        if x < -0.02 or y < -0.02 or x + w > W + 0.05 or y + h > H + 0.02:
            warnings.append(f"第{box['page']:02d}页文本框越界 x={x:.2f} y={y:.2f} {w:.2f}×{h:.2f}")
        usable_w = w - box["pad"] * 2
        usable_h = h - 0.04
        if usable_w <= 0.15 or usable_h <= 0.05:
            continue
        used = 0.0
        snippet = ""
        for para in box["paras"]:
            size = para["size"]
            text = para["text"]
            if not snippet and text.strip():
                snippet = text.strip()[:24]
            units = 0.0
            for ch in text:
                units += 0.55 if (ch.isascii() or ch.isspace()) else 1.0
            char_w = size / 72
            lines = 1 if not text else max(1, math.ceil(units * char_w / usable_w))
            line_h = size / 72 * max(para["spc"], 1) * 1.15
            used += para["before"] / 72 + lines * line_h + para["after"] / 72
        if any(p["text"].strip() for p in box["paras"]) and used > usable_h + 0.08:
            warnings.append(f"第{box['page']:02d}页可能溢出 ({w:.2f}×{h:.2f}in, 需{used:.2f}>盒{usable_h:.2f}): {snippet}")
    return warnings


def missing_phrases(prs: Presentation) -> list[str]:
    blob = "\n".join(shape.text_frame.text for slide in prs.slides for shape in slide.shapes if shape.has_text_frame)
    return [item for item in REQUIRED if item not in blob]


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    global BOXES
    BOXES = []
    prs = build()
    warnings = overflow_warnings()
    missing = missing_phrases(prs)
    prs.save(PPTX_PATH)
    print(f"{len(prs.slides)} 页  {PPTX_PATH}")
    if missing:
        print("缺少关键表述：")
        for item in missing:
            print(" -", item)
    if warnings:
        print("版面预警：")
        for item in warnings:
            print(" -", item)
    return 1 if warnings or missing else 0


if __name__ == "__main__":
    sys.exit(main())
