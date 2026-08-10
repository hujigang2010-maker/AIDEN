# -*- coding: utf-8 -*-
"""生成 PPT：上海AI博物馆可执行落地方案（侧重投资 / 盈利 / 政策）。"""
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
OUT_FILE = OUT / "上海AI博物馆_可执行落地方案.pptx"

# 青绿+海军蓝（机构/科创，避免紫白套路）
NAVY = RGBColor(0x0B, 0x3D, 0x5C)
TEAL = RGBColor(0x0F, 0x7A, 0x6E)
TEAL_DEEP = RGBColor(0x0A, 0x4F, 0x48)
SAND = RGBColor(0xF4, 0xF7, 0xF6)
MINT = RGBColor(0xD7, 0xEB, 0xE6)
INK = RGBColor(0x1A, 0x2B, 0x2E)
GREY = RGBColor(0x5C, 0x6B, 0x70)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
AMBER = RGBColor(0xC4, 0x7B, 0x2D)
RED = RGBColor(0xA6, 0x3D, 0x2F)


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
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
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
        p.line_spacing = 1.2
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


def make_card(slide, left, top, width, height, title, body_items, accent=TEAL):
    add_rect(slide, left, top, width, height, fill=SAND)
    add_rect(slide, left, top, Inches(0.07), height, fill=accent)
    add_text(slide, left + Inches(0.18), top + Inches(0.08),
             width - Inches(0.28), Inches(0.35),
             title, size=13, bold=True, color=accent)
    add_bullets(slide, left + Inches(0.18), top + Inches(0.45),
                width - Inches(0.28), height - Inches(0.5),
                body_items, size=11)


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
            set_font(r, size=font_size - 1, color=INK)
    return table


def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    TOTAL = 14
    page = [0]

    def new():
        page[0] += 1
        return prs.slides.add_slide(blank), page[0]

    # 1 封面
    s, _ = new()
    add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), fill=NAVY)
    add_rect(s, Inches(0), Inches(0), Inches(0.18), Inches(7.5), fill=TEAL)
    add_text(s, Inches(0.9), Inches(1.6), Inches(11), Inches(0.4),
             "工作会议纪要 → 可执行落地方案", size=14, color=MINT)
    add_text(s, Inches(0.9), Inches(2.1), Inches(11.5), Inches(1.0),
             C.PROJECT, size=32, bold=True, color=WHITE)
    add_text(s, Inches(0.9), Inches(3.3), Inches(11.5), Inches(0.5),
             "核心聚焦：投资方怎么投 · 项目怎么赚 · 政策怎么落地",
             size=18, bold=True, color=AMBER)
    add_text(s, Inches(0.9), Inches(4.2), Inches(11.5), Inches(0.8),
             [
                 "基于两场会议笔记整理（2026-08-08 / 2026-08-10）",
                 "对标美国计算机历史博物馆 · WAIC会后365天常态化平台",
             ],
             size=13, color=MINT)
    add_text(s, Inches(0.9), Inches(6.3), Inches(11), Inches(0.4),
             f"{C.VERSION} 讨论稿  |  {C.DATE_STR}  |  供场地 / 资方 / 科协路演使用",
             size=11, color=GREY)

    # 2 目录
    s, p = new()
    slide_header(s, "目录", "先对齐共识，再谈投、赚、政策", p, TOTAL)
    toc = [
        "01  双场会议纪要要点与核心关注",
        "02  项目定位与差异化",
        "03  投资总原则与七条出资路径（重点）",
        "04  赞助层级与资方回报设计（重点）",
        "05  盈利逻辑与六条收入线（重点）",
        "06  政策性支持与扶持资金匹配（重点）",
        "07  场地候选与空间经济模型",
        "08  展陈分层与资源地图",
        "09  组织分工与90天推进节奏",
        "10  风险清单与下一步待办",
    ]
    add_bullets(s, Inches(1.2), Inches(1.6), Inches(10), Inches(5.2),
                toc, size=16, bullet="▸")

    # 3 纪要
    s, p = new()
    slide_header(s, "双场会议纪要要点", "场次一讨论定方向 · 场次二工作会议定推进", p, TOTAL)
    add_text(s, Inches(0.5), Inches(1.35), Inches(12.3), Inches(0.7),
             C.MEETING_SUMMARY, size=12, color=INK)
    make_card(s, Inches(0.5), Inches(2.15), Inches(4.0), Inches(4.5),
              "场次一 · 08-08", [
                  "教育部认证 + 研学切入",
                  "一二层信息化展馆",
                  "三层结构：算力/应用/移动AI",
                  "华润等赞助商思路",
                  "海外账号反向变现",
                  "运营周期 5–10 年",
              ])
    make_card(s, Inches(4.7), Inches(2.15), Inches(4.0), Inches(4.5),
              "场次二 · 08-10", [
                  "陈院长提出完整方案",
                  "免租10–20年+约5000万设备",
                  "创智汇 / 复兴岛候选",
                  "科协牵头抢IP",
                  "租金+孵化+模板输出",
                  "胡老师整理可展示版本",
              ], accent=NAVY)
    make_card(s, Inches(8.9), Inches(2.15), Inches(3.9), Inches(4.5),
              "本方案核心关注", C.CORE_CONCERNS, accent=AMBER)

    # 4 定位
    s, p = new()
    slide_header(s, "项目定位与差异化", "名称IP × 全产业链 × 会后365天", p, TOTAL)
    make_card(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(2.4),
              "一句话定位", [
                  "世界人工智能大会闭幕后的常态化AI展示与孵化平台",
                  "覆盖芯片—应用—具身全链，串联历史与最新成果",
                  "做成可向其他城市复制的标准模板",
              ])
    make_card(s, Inches(6.7), Inches(1.4), Inches(6.0), Inches(2.4),
              "与现有展厅差异", C.DIFFERENTIATION, accent=NAVY)
    make_table(
        s, Inches(0.5), Inches(4.0), Inches(12.3), Inches(2.6),
        ["维度", "现有机器人展厅", "本项目AI博物馆"],
        [
            ["范围", "机器人为主", "AI全产业链+历史溯源"],
            ["功能", "交易与展示", "展示+研学认证+孵化投资+模板输出"],
            ["内容", "国内产业为主", "中美对照 + 硅谷实时连线"],
            ["组织", "企业/园区主导", "科协牵头 + 多主体共建"],
        ],
        font_size=11,
    )

    # 5 投资原则+路径总览
    s, p = new()
    slide_header(s, "投资总原则：不自有资金站上制高点", "多主体分担 · 分层组合 · 用IP换长期权益", p, TOTAL)
    add_text(s, Inches(0.5), Inches(1.35), Inches(12.3), Inches(0.7),
             C.INVEST_PRINCIPLE, size=12, color=INK)
    rows = []
    for it in C.INVEST_PATHS:
        rows.append([it["路径"], it["出资方式"][:28] + "…", it["优先级"]])
    make_table(
        s, Inches(0.5), Inches(2.2), Inches(12.3), Inches(4.6),
        ["出资路径", "出资方式（摘要）", "优先级"],
        [[r[0], r[1].replace("…", ""), r[2]] for r in
         [[it["路径"], it["出资方式"], it["优先级"]] for it in C.INVEST_PATHS]],
        font_size=10,
    )

    # 6 投资路径细节（P0）
    s, p = new()
    slide_header(s, "投资方怎么投：P0 四条主路径", "场地钱、主赞助、政府基金、科协专项先闭环", p, TOTAL)
    p0 = [x for x in C.INVEST_PATHS if x["优先级"] == "P0"]
    positions = [
        (0.4, 1.35), (6.75, 1.35), (0.4, 4.2), (6.75, 4.2),
    ]
    for (left, top), it in zip(positions, p0):
        make_card(
            s, Inches(left), Inches(top), Inches(6.1), Inches(2.6),
            it["路径"],
            [
                f"怎么出：{it['出资方式']}",
                f"谁来出：{it['出资方画像']}",
                f"回报：{it['回报机制']}",
                f"动作：{it['落地动作']}",
            ],
        )

    # 7 赞助层级
    s, p = new()
    slide_header(s, "赞助层级与资方回报设计", "把「历史席位」产品化，方便企业立项", p, TOTAL)
    headers = C.SPONSOR_TIERS[0]
    rows = C.SPONSOR_TIERS[1:]
    make_table(s, Inches(0.4), Inches(1.4), Inches(12.5), Inches(3.6),
               headers, rows, font_size=11)
    add_bullets(
        s, Inches(0.5), Inches(5.2), Inches(12), Inches(1.6),
        [
            "华润路径：对齐其科创投入与向上汇报需求，同学通道 + 董事会材料双轨推进",
            "头部AI企业：捐赠换「行业历史席位」，比纯广告投放更易内部批准",
            "场地方：免租+设备出资，换政绩场景、物业去化与租金/招商分成",
        ],
        size=13,
    )

    # 8 盈利
    s, p = new()
    slide_header(s, "盈利逻辑：展馆不靠卖票养活，靠IP经营赚钱", C.PROFIT_LOGIC[:60] + "…", p, TOTAL)
    make_table(
        s, Inches(0.35), Inches(1.35), Inches(12.6), Inches(5.5),
        ["收入线", "稳态占比", "启动条件", "里程碑"],
        [
            [r["收入线"], r["目标占比(稳态)"], r["启动条件"], r["里程碑"]]
            for r in C.REVENUE_STREAMS
        ],
        font_size=10,
    )

    # 9 政策
    s, p = new()
    slide_header(s, "政策性支持怎么落地", "研学强制要求 × 科协牵头 × 扶持基金 × 闲置物业", p, TOTAL)
    make_table(
        s, Inches(0.3), Inches(1.35), Inches(12.7), Inches(5.5),
        ["政策/抓手", "对项目价值", "落地步骤", "责任方"],
        [
            [x["政策/抓手"], x["对项目价值"], x["落地步骤"], x["责任方"]]
            for x in C.POLICY_SUPPORT
        ],
        font_size=9,
    )

    # 10 扶持资金匹配
    s, p = new()
    slide_header(s, "扶持基金与资金类型匹配表", "申报窗口与用途一一对应，避免「有政策无材料」", p, TOTAL)
    make_table(
        s, Inches(0.4), Inches(1.4), Inches(12.5), Inches(4.2),
        C.POLICY_FUND_MATCH[0],
        C.POLICY_FUND_MATCH[1:],
        font_size=11,
    )
    add_text(
        s, Inches(0.5), Inches(5.8), Inches(12.3), Inches(0.9),
        "执行口诀：先签场地框架（免租+配套资金）→ 同步科协名义 → 再按窗口填报专项；"
        "现金资助是加法，场地与设备封闭才是项目能否开工的充要条件。",
        size=13, bold=True, color=TEAL_DEEP,
    )

    # 11 场地
    s, p = new()
    slide_header(s, "场地候选与空间经济模型", "6000–10000㎡总盘 · 约2000㎡展陈 · 其余经营反哺", p, TOTAL)
    make_table(
        s, Inches(0.3), Inches(1.35), Inches(12.7), Inches(2.8),
        C.SITE_CANDIDATES[0],
        C.SITE_CANDIDATES[1:],
        font_size=10,
    )
    make_table(
        s, Inches(0.3), Inches(4.4), Inches(12.7), Inches(2.5),
        C.UNIT_ECONOMICS[0],
        C.UNIT_ECONOMICS[1:],
        font_size=10,
    )

    # 12 展陈
    s, p = new()
    slide_header(s, "展陈分层与资源地图", "美方理论前沿 · 中方产业应用 · 实时更新", p, TOTAL)
    make_table(
        s, Inches(0.4), Inches(1.4), Inches(12.5), Inches(3.5),
        C.EXHIBIT_LAYERS[0],
        C.EXHIBIT_LAYERS[1:],
        font_size=11,
    )
    add_bullets(
        s, Inches(0.5), Inches(5.2), Inches(12), Inches(1.6),
        [
            "已筛美国核心展品约前20项；国内优先头部有收入企业",
            "陈院长可对接谷歌等资源；可联动李飞飞等海外人物直播",
            "国内头部大模型企业多数注册在上海，集中拜访效率高",
        ],
        size=13,
    )

    # 13 节奏
    s, p = new()
    slide_header(s, "组织分工与推进节奏", "先框架预算，再动线施工图", p, TOTAL)
    make_table(
        s, Inches(0.3), Inches(1.3), Inches(6.2), Inches(3.5),
        C.ORG_ROLES[0],
        C.ORG_ROLES[1:],
        font_size=9,
    )
    make_table(
        s, Inches(6.7), Inches(1.3), Inches(6.2), Inches(5.4),
        C.ROADMAP[0],
        C.ROADMAP[1:],
        font_size=9,
    )

    # 14 风险与待办
    s, p = new()
    slide_header(s, "风险清单与下一步待办", "谈判与预算是当前最大阻塞点", p, TOTAL)
    make_table(
        s, Inches(0.3), Inches(1.3), Inches(12.7), Inches(3.2),
        C.RISKS[0],
        C.RISKS[1:],
        font_size=9,
    )
    add_text(s, Inches(0.45), Inches(4.65), Inches(12), Inches(0.35),
             "会后待办（摘自纪要）", size=14, bold=True, color=TEAL)
    add_bullets(
        s, Inches(0.5), Inches(5.05), Inches(12.3), Inches(1.8),
        [f"{r[0]}：{r[1]}（{r[2]} / {r[3]}）" for r in C.TODOS],
        size=11,
    )

    prs.save(OUT_FILE)
    print(f"已生成: {OUT_FILE}")


if __name__ == "__main__":
    build()
