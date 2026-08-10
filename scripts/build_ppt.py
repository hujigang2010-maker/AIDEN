# -*- coding: utf-8 -*-
"""生成 PPT：上海人工智能产业展馆 · 对外汇报版。"""
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

import content as C

OUT = Path(__file__).resolve().parent.parent / "deliverables"
OUT.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT / "上海人工智能产业展馆_对外汇报方案.pptx"

NAVY = RGBColor(0x0B, 0x3D, 0x5C)
TEAL = RGBColor(0x0F, 0x7A, 0x6E)
TEAL_DEEP = RGBColor(0x0A, 0x4F, 0x48)
SAND = RGBColor(0xF4, 0xF7, 0xF6)
MINT = RGBColor(0xD7, 0xEB, 0xE6)
INK = RGBColor(0x1A, 0x2B, 0x2E)
GREY = RGBColor(0x5C, 0x6B, 0x70)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
AMBER = RGBColor(0xC4, 0x7B, 0x2D)


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


def add_text(slide, left, top, width, height, text, size=18, bold=False,
             color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, Emu(0))
    tf.vertical_anchor = anchor
    if isinstance(text, str):
        text = [text]
    for i, t in enumerate(text):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = t
        set_font(run, size=size, bold=bold, color=color)
    return tb


def add_bullets(slide, left, top, width, height, items, size=13, color=INK, bullet="•"):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.15
        run = p.add_run()
        run.text = f"{bullet}  {it}"
        set_font(run, size=size, color=color)
    return tb


def slide_header(slide, title, subtitle=None, page_no=None, total=None):
    add_rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.42), fill=NAVY)
    add_rect(slide, Inches(0), Inches(0.42), Inches(13.333), Inches(0.06), fill=TEAL)
    add_text(slide, Inches(0.45), Inches(0.55), Inches(10.5), Inches(0.45),
             title, size=22, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_text(slide, Inches(0.45), Inches(0.98), Inches(12), Inches(0.3),
                 subtitle, size=11, color=GREY)
    if page_no and total:
        add_text(slide, Inches(11.4), Inches(7.05), Inches(1.6), Inches(0.28),
                 f"{page_no} / {total}", size=10, color=GREY, align=PP_ALIGN.RIGHT)


def make_card(slide, left, top, width, height, title, body_items, accent=TEAL, body_size=11):
    add_rect(slide, left, top, width, height, fill=SAND)
    add_rect(slide, left, top, Inches(0.07), height, fill=accent)
    add_text(slide, left + Inches(0.18), top + Inches(0.08),
             width - Inches(0.28), Inches(0.35),
             title, size=13, bold=True, color=accent)
    add_bullets(slide, left + Inches(0.18), top + Inches(0.45),
                width - Inches(0.28), height - Inches(0.5),
                body_items, size=body_size)


def make_table(slide, left, top, width, height, headers, rows, font_size=10):
    shape = slide.shapes.add_table(len(rows) + 1, len(headers), left, top, width, height)
    table = shape.table
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        tf = cell.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h
        set_font(r, size=font_size, bold=True, color=WHITE)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = table.cell(ri + 1, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if ri % 2 == 0 else MINT
            tf = cell.text_frame
            tf.clear()
            tf.word_wrap = True
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = str(val)
            set_font(r, size=max(font_size - 1, 8), color=INK)
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
    add_rect(s, Inches(0), Inches(0), Inches(0.18), Inches(7.5), fill=TEAL)
    add_text(s, Inches(0.9), Inches(1.5), Inches(11), Inches(0.4),
             "对外汇报方案", size=14, color=MINT)
    add_text(s, Inches(0.9), Inches(2.0), Inches(11.5), Inches(1.0),
             C.PROJECT, size=30, bold=True, color=WHITE)
    add_text(s, Inches(0.9), Inches(3.2), Inches(11.5), Inches(0.6),
             f"样板对标：{C.BENCHMARK}", size=14, color=AMBER)
    add_text(s, Inches(0.9), Inches(4.0), Inches(11.5), Inches(0.8),
             [
                 "租金与装修产业生态分摊  ·  多元盈利结构  ·  支持单位协同",
                 "汇报人团队：陈晟 · 胡继刚 · 陈红苗",
             ], size=13, color=MINT)
    add_text(s, Inches(0.9), Inches(6.3), Inches(11), Inches(0.4),
             f"{C.VERSION}  |  {C.DATE_STR}  |  供合作方 / 主管部门 / 支持单位汇报",
             size=11, color=GREY)

    # 2 目录
    s, p = new()
    slide_header(s, "目录", "对外汇报结构", p, TOTAL)
    toc = [
        "01  项目概要与核心议题",
        "02  对标 CHM：一对一深化",
        "03  上海落地定位",
        "04  租金的产业生态逻辑（重点）",
        "05  装修费的分层共担（重点）",
        "06  投资与赞助结构",
        "07  支持单位",
        "08  多元盈利模式",
        "09  政策与资金匹配",
        "10  选址、内容、节奏与下一步",
    ]
    add_bullets(s, Inches(1.2), Inches(1.55), Inches(10), Inches(5.2),
                toc, size=16, bullet="▸")

    # 3 概要（无场次）
    s, p = new()
    slide_header(s, "项目概要与核心议题", "对外共识摘要", p, TOTAL)
    add_text(s, Inches(0.5), Inches(1.35), Inches(12.3), Inches(1.0),
             C.EXEC_SUMMARY, size=12, color=INK)
    make_card(s, Inches(0.5), Inches(2.5), Inches(4.0), Inches(4.2),
              "定位共识", [
                  "常设展馆，非临展",
                  "对标加州 CHM 样板",
                  "叙事主线 + 展位年更",
                  "研学与公众双客群",
              ])
    make_card(s, Inches(4.7), Inches(2.5), Inches(4.0), Inches(4.2),
              "经营共识", [
                  "先筹后建",
                  "租金阶梯 + 装修分层",
                  "启动金必须闭环",
                  "门票不作主收入",
              ], accent=NAVY)
    make_card(s, Inches(8.9), Inches(2.5), Inches(3.9), Inches(4.2),
              "本汇报核心关注", C.CORE_CONCERNS, accent=AMBER)

    # 4 CHM 画像
    s, p = new()
    slide_header(s, "样板馆画像：计算机历史博物馆（加州山景城）", C.BENCHMARK, p, TOTAL)
    rows = [[k, v] for k, v in C.CHM_PROFILE.items()]
    make_table(s, Inches(0.4), Inches(1.35), Inches(12.5), Inches(5.5),
               ["维度", "CHM 要点"], rows, font_size=11)

    # 5 一对一对照
    s, p = new()
    slide_header(s, "一对一深化：CHM → 上海方案", "学机制，不适配处做本土化", p, TOTAL)
    make_table(
        s, Inches(0.25), Inches(1.25), Inches(12.8), Inches(5.7),
        C.CHM_COMPARE[0], C.CHM_COMPARE[1:], font_size=9,
    )

    # 6 本土化要点
    s, p = new()
    slide_header(s, "上海落地定位", "承接 CHM 方法，适配国内产业与回款结构", p, TOTAL)
    add_bullets(s, Inches(0.6), Inches(1.4), Inches(12), Inches(2.8),
                C.CHM_LOCALIZE, size=14)
    make_card(s, Inches(0.5), Inches(4.3), Inches(12.3), Inches(2.3),
              "差异化要点", C.DIFFERENTIATION, accent=TEAL)

    # 7 租金
    s, p = new()
    slide_header(s, "租金：产业生态里的四级结构（重点）", C.RENT_ECOLOGY[:48] + "…", p, TOTAL)
    make_table(
        s, Inches(0.3), Inches(1.35), Inches(12.7), Inches(5.5),
        ["层级", "逻辑", "谁承担", "谈判点"],
        [[x["层级"], x["逻辑"], x["谁承担"], x["谈判点"]] for x in C.RENT_LAYERS],
        font_size=10,
    )

    # 8 装修
    s, p = new()
    slide_header(s, "装修费：壳装 / 馆装 / 专项装三层共担（重点）", C.FITOUT_ECOLOGY[:48] + "…", p, TOTAL)
    make_table(
        s, Inches(0.3), Inches(1.3), Inches(12.7), Inches(3.6),
        ["科目", "典型内容", "建议承担", "回收方式"],
        [[x["科目"], x["典型内容"], x["建议承担"], x["回收方式"]] for x in C.FITOUT_LAYERS],
        font_size=10,
    )
    add_text(s, Inches(0.5), Inches(5.1), Inches(12.3), Inches(0.35),
             "落地步骤（先测算，再谈判，后开工）", size=13, bold=True, color=TEAL)
    add_bullets(
        s, Inches(0.5), Inches(5.45), Inches(12.3), Inches(1.4),
        [f"{r[0]}. {r[1]} → {r[2]}" for r in C.RENT_FITOUT_PLAYBOOK[1:]],
        size=11,
    )

    # 9 投资
    s, p = new()
    slide_header(s, "投资与赞助结构", C.INVEST_PRINCIPLE[:50] + "…", p, TOTAL)
    make_table(
        s, Inches(0.3), Inches(1.3), Inches(12.7), Inches(5.6),
        ["路径", "出资方式", "回报机制", "优先级"],
        [[it["路径"], it["出资方式"], it["回报机制"], it["优先级"]] for it in C.INVEST_PATHS],
        font_size=9,
    )

    # 10 支持单位
    s, p = new()
    slide_header(s, "支持单位", "智库 + 市级/区级产业网络协同", p, TOTAL)
    make_table(
        s, Inches(0.3), Inches(1.3), Inches(12.7), Inches(4.2),
        C.SUPPORT_ORG_ROLES[0], C.SUPPORT_ORG_ROLES[1:], font_size=11,
    )
    add_bullets(
        s, Inches(0.5), Inches(5.7), Inches(12.3), Inches(1.2),
        [
            "拟以合作备忘录明确：联合课题、会员企业对接、活动联办、品牌互挂",
            "支持单位不替代运营主体出资义务，重点降低获客、策展与政企沟通成本",
        ],
        size=12,
    )

    # 11 盈利1
    s, p = new()
    slide_header(s, "多元盈利模式（上）", C.PROFIT_LOGIC[:52] + "…", p, TOTAL)
    make_table(
        s, Inches(0.25), Inches(1.25), Inches(12.8), Inches(5.7),
        ["收入线", "描述", "延伸玩法", "占比"],
        [
            [r["收入线"], r["描述"], r["想象力延伸"], r["占比"]]
            for r in C.REVENUE_STREAMS[:5]
        ],
        font_size=9,
    )

    # 12 盈利2
    s, p = new()
    slide_header(s, "多元盈利模式（下）", "内容、定制、场景、数据与后期输出", p, TOTAL)
    make_table(
        s, Inches(0.25), Inches(1.25), Inches(12.8), Inches(5.7),
        ["收入线", "描述", "延伸玩法", "占比"],
        [
            [r["收入线"], r["描述"], r["想象力延伸"], r["占比"]]
            for r in C.REVENUE_STREAMS[5:]
        ],
        font_size=9,
    )

    # 13 政策
    s, p = new()
    slide_header(s, "政策与资金匹配", "批复与合同说话，专项保守入账", p, TOTAL)
    make_table(
        s, Inches(0.3), Inches(1.3), Inches(12.7), Inches(5.6),
        ["政策/抓手", "对项目价值", "落地步骤", "责任方"],
        [
            [x["政策/抓手"], x["对项目价值"], x["落地步骤"], x["责任方"]]
            for x in C.POLICY_SUPPORT
        ],
        font_size=10,
    )

    # 14 选址+内容
    s, p = new()
    slide_header(s, "选址与内容结构", "一期可控 · 对标 CHM 常设+轮换", p, TOTAL)
    make_table(
        s, Inches(0.3), Inches(1.25), Inches(12.7), Inches(2.6),
        C.SITE_CANDIDATES[0], C.SITE_CANDIDATES[1:], font_size=9,
    )
    make_table(
        s, Inches(0.3), Inches(4.05), Inches(12.7), Inches(2.8),
        C.CONTENT_LAYERS[0], C.CONTENT_LAYERS[1:], font_size=9,
    )

    # 15 节奏
    s, p = new()
    slide_header(s, "组织分工与推进节奏", "先测算与对标，再谈判与闭环，后建设", p, TOTAL)
    make_table(
        s, Inches(0.3), Inches(1.25), Inches(6.1), Inches(3.6),
        C.ORG_ROLES[0], C.ORG_ROLES[1:], font_size=9,
    )
    make_table(
        s, Inches(6.6), Inches(1.25), Inches(6.4), Inches(5.5),
        C.ROADMAP[0], C.ROADMAP[1:], font_size=9,
    )

    # 16 下一步
    s, p = new()
    slide_header(s, "风险要点与下一步", "对外协同清单", p, TOTAL)
    make_table(
        s, Inches(0.3), Inches(1.25), Inches(12.7), Inches(3.0),
        C.RISKS[0], C.RISKS[1:], font_size=9,
    )
    add_text(s, Inches(0.45), Inches(4.4), Inches(12), Inches(0.35),
             "建议下一步", size=14, bold=True, color=TEAL)
    add_bullets(s, Inches(0.5), Inches(4.8), Inches(12.3), Inches(2.0),
                C.NEXT_STEPS, size=12)

    prs.save(OUT_FILE)
    print(f"已生成: {OUT_FILE}")


if __name__ == "__main__":
    build()
