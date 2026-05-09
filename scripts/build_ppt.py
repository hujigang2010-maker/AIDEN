"""Generate sponsorship pitch PowerPoint deck.

Produces: 赞助方案/2026峰会赞助介绍.pptx
"""
from __future__ import annotations

import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Cm, Emu, Pt


OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "赞助方案")
OUT_PATH = os.path.normpath(os.path.join(OUT_DIR, "2026峰会赞助介绍.pptx"))

# 16:9 widescreen colors
NAVY = RGBColor(0x0E, 0x1F, 0x44)
BLUE = RGBColor(0x1F, 0x49, 0x7D)
GOLD = RGBColor(0xC9, 0xA2, 0x27)
LIGHT = RGBColor(0xF4, 0xF6, 0xFA)
GRAY = RGBColor(0x4A, 0x4A, 0x4A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT = RGBColor(0xE3, 0x4F, 0x26)


def add_blank_slide(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def add_rect(slide, left, top, width, height, fill=NAVY, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
    shape.shadow.inherit = False
    return shape


def add_text(slide, text, left, top, width, height, *, size=18, bold=False,
             color=WHITE, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             font="Microsoft YaHei"):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.margin_left = Cm(0.1)
    tf.margin_right = Cm(0.1)
    tf.margin_top = Cm(0.05)
    tf.margin_bottom = Cm(0.05)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    lines = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.name = font
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
    return tb


def add_bullet_list(slide, items, left, top, width, height, *,
                    size=16, color=GRAY, font="Microsoft YaHei",
                    line_spacing=1.3):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Cm(0.1)
    tf.margin_right = Cm(0.1)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = line_spacing
        run = p.add_run()
        run.text = "•  " + item
        run.font.name = font
        run.font.size = Pt(size)
        run.font.color.rgb = color
    return tb


def page_frame(slide, title: str, subtitle: str = "", page_no: int = 0, total: int = 0):
    SW = Cm(33.867)
    SH = Cm(19.05)
    add_rect(slide, 0, 0, SW, SH, fill=WHITE)
    # top bar
    add_rect(slide, 0, 0, SW, Cm(1.6), fill=NAVY)
    add_rect(slide, Cm(1.2), Cm(0.45), Cm(0.4), Cm(0.7), fill=GOLD)
    add_text(slide, "重构与突围 · 2026 AI 商业化峰会",
             Cm(1.8), Cm(0.35), Cm(20), Cm(0.9),
             size=14, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    if page_no:
        add_text(slide, f"P{page_no:02d} / {total:02d}",
                 Cm(30), Cm(0.35), Cm(2.5), Cm(0.9),
                 size=12, color=WHITE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    # title
    add_text(slide, title, Cm(1.5), Cm(2.0), Cm(30), Cm(1.5),
             size=28, bold=True, color=NAVY)
    if subtitle:
        add_text(slide, subtitle, Cm(1.5), Cm(3.4), Cm(30), Cm(1.0),
                 size=14, color=GRAY)
    # divider
    add_rect(slide, Cm(1.5), Cm(3.55), Cm(2.5), Cm(0.08), fill=GOLD)


def slide_cover(prs, total):
    s = add_blank_slide(prs)
    SW = Cm(33.867)
    SH = Cm(19.05)
    add_rect(s, 0, 0, SW, SH, fill=NAVY)
    # accent bars
    add_rect(s, 0, Cm(2.5), Cm(0.4), Cm(14), fill=GOLD)
    add_rect(s, Cm(33.4), Cm(2.5), Cm(0.4), Cm(14), fill=GOLD)

    add_text(s, "重构与突围", Cm(2), Cm(3.0), Cm(30), Cm(2.2),
             size=60, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, "2026 人工智能商业化落地与硬核投资破局峰会",
             Cm(2), Cm(5.5), Cm(30), Cm(1.5),
             size=26, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    add_text(s, "Sponsorship Prospectus  ·  招商赞助方案",
             Cm(2), Cm(7.4), Cm(30), Cm(1.0),
             size=18, color=WHITE, align=PP_ALIGN.CENTER)

    add_rect(s, Cm(11), Cm(9.2), Cm(12), Cm(0.05), fill=WHITE)
    add_text(s, "寻找 AI 时代的超级个体、新质资产与资本新风口",
             Cm(2), Cm(9.6), Cm(30), Cm(1.2),
             size=16, color=WHITE, align=PP_ALIGN.CENTER)

    add_text(s, "主办：北京大学经济学院上海校友会  ·  复旦大学住房政策研究中心",
             Cm(2), Cm(15.4), Cm(30), Cm(0.8),
             size=14, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, "席位倒计时：剩余 30 席深度合作 + 1 席独家总冠名  ·  截止 5 月 18 日",
             Cm(2), Cm(16.4), Cm(30), Cm(0.8),
             size=14, color=GOLD, align=PP_ALIGN.CENTER, font="Microsoft YaHei")


def slide_overview(prs, idx, total):
    s = add_blank_slide(prs)
    page_frame(s, "大会概览", "Event Overview", idx, total)

    items = [
        ("规模", "500 位高净值嘉宾"),
        ("时长", "8 小时议程 + 2.5 小时 VIP 闭门晚宴"),
        ("城市", "上海"),
        ("时间", "2026 年 5 月（具体以邀请函为准）"),
    ]
    # KPI row
    base_left = Cm(1.5)
    box_w = Cm(7.6)
    box_h = Cm(3.2)
    gap = Cm(0.4)
    for i, (k, v) in enumerate(items):
        left = base_left + (box_w + gap) * i
        add_rect(s, left, Cm(5.2), box_w, box_h, fill=LIGHT)
        add_rect(s, left, Cm(5.2), Cm(0.15), box_h, fill=GOLD)
        add_text(s, k, left + Cm(0.4), Cm(5.5), box_w - Cm(0.5), Cm(0.8),
                 size=14, color=GRAY)
        add_text(s, v, left + Cm(0.4), Cm(6.3), box_w - Cm(0.5), Cm(2.0),
                 size=20, bold=True, color=NAVY)

    add_text(s, "主办与协办", Cm(1.5), Cm(9.2), Cm(30), Cm(1.0),
             size=18, bold=True, color=NAVY)
    add_bullet_list(s, [
        "主办：北京大学经济学院上海校友会、复旦大学住房政策研究中心",
        "协办：上海市科技企业联合会  ·  上海市杨浦区科技企业联合会  ·  上海市虹口区科技企业联合会",
        "协办：中东投资联盟  ·  中国商业文化研究会长三角分会",
    ], Cm(1.5), Cm(10.2), Cm(30), Cm(5), size=15)


def slide_value_props(prs, idx, total):
    s = add_blank_slide(prs)
    page_frame(s, "为什么赞助本次峰会", "Why Sponsor", idx, total)

    cards = [
        ("01", "顶级嘉宾矩阵", [
            "腾讯云 / 阿里云 / 火山引擎",
            "白硕、夏春、孔华威、张露瑶",
            "招商银行 / 长江证券 / 金浦投资",
        ], NAVY),
        ("02", "高净值受众触达", [
            "上市公司董事长、独角兽高管",
            "一二级市场基金合伙人",
            "AI / 算力 / 大模型创始团队",
        ], BLUE),
        ("03", "全场景品牌渗透", [
            "主视觉 → 易拉宝 → 议程手册",
            "主舞台口播 → 晚宴主桌植入",
            "白皮书 → 媒体通稿 → 朋友圈",
        ], GOLD),
        ("04", "长效圈层沉淀", [
            "双校长三角校友产业联盟",
            "杨浦区科技企业联合会",
            "1V1 闭门嘉宾对接",
        ], ACCENT),
    ]
    base_left = Cm(1.5)
    box_w = Cm(7.6)
    box_h = Cm(11)
    gap = Cm(0.4)
    for i, (no, title, lines, color) in enumerate(cards):
        left = base_left + (box_w + gap) * i
        add_rect(s, left, Cm(5.0), box_w, box_h, fill=LIGHT)
        add_rect(s, left, Cm(5.0), box_w, Cm(2.0), fill=color)
        add_text(s, no, left + Cm(0.5), Cm(5.1), Cm(2), Cm(1.6),
                 size=22, bold=True, color=WHITE)
        add_text(s, title, left + Cm(2.5), Cm(5.4), box_w - Cm(2.5), Cm(1.2),
                 size=18, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
        add_bullet_list(s, lines, left + Cm(0.4), Cm(7.4), box_w - Cm(0.7), Cm(7),
                        size=13, color=GRAY)


def slide_audience(prs, idx, total):
    s = add_blank_slide(prs)
    page_frame(s, "受众画像", "Audience Profile · 500 人高净值矩阵", idx, total)

    rows = [
        ("上市公司 / 独角兽高管、董事长", "25%", NAVY),
        ("一二级市场基金合伙人、首席分析师", "20%", BLUE),
        ("AI / 算力 / 大模型创业团队创始人", "25%", GOLD),
        ("北大、复旦双校核心校友与联盟会员", "20%", ACCENT),
        ("政府、产业园区、媒体合作伙伴", "10%", GRAY),
    ]
    top0 = Cm(5.5)
    bar_h = Cm(1.4)
    gap = Cm(0.4)
    label_w = Cm(13)
    bar_full = Cm(15)
    for i, (label, pct, color) in enumerate(rows):
        top = top0 + (bar_h + gap) * i
        add_text(s, label, Cm(1.5), top, label_w, bar_h,
                 size=15, color=GRAY, anchor=MSO_ANCHOR.MIDDLE)
        # background
        add_rect(s, Cm(15), top, bar_full, bar_h, fill=LIGHT)
        pct_val = int(pct.strip("%"))
        add_rect(s, Cm(15), top, Emu(int(bar_full * pct_val / 100)), bar_h, fill=color)
        add_text(s, pct, Cm(30.4), top, Cm(2.5), bar_h,
                 size=16, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)


def slide_speakers(prs, idx, total):
    s = add_blank_slide(prs)
    page_frame(s, "核心嘉宾阵容", "Featured Speakers", idx, total)

    speakers = [
        ("白硕", "恒生研究院院长 / 前上交所总工程师"),
        ("夏春 博士", "香港国际金融学会副会长"),
        ("徐永泽", "腾讯云副总裁、华东渠道生态总经理"),
        ("吴晓东", "阿里云华东大区总经理"),
        ("邵凝光", "火山引擎华东生态渠道总经理"),
        ("姚志勇 教授", "复旦管院 / 北大经院上海校友会会长"),
        ("王维军", "泓塬资产董事长"),
        ("张露瑶 博士", "昆山杜克大学助理教授"),
        ("孔华威", "张江科创首席专家"),
        ("曲承东", "招商银行资金运营中心总经理助理"),
        ("刘胜利", "长江证券研究所金融工程首席"),
        ("饶雪莹", "金浦投资金融科技基金合伙人"),
    ]
    cols = 4
    rows = 3
    base_left = Cm(1.5)
    base_top = Cm(5.0)
    cw = Cm(7.6)
    ch = Cm(3.4)
    gap_x = Cm(0.4)
    gap_y = Cm(0.4)
    for i, (name, title) in enumerate(speakers):
        r = i // cols
        c = i % cols
        left = base_left + (cw + gap_x) * c
        top = base_top + (ch + gap_y) * r
        add_rect(s, left, top, cw, ch, fill=LIGHT)
        add_rect(s, left, top, Cm(0.15), ch, fill=GOLD)
        add_text(s, name, left + Cm(0.4), top + Cm(0.4), cw - Cm(0.5), Cm(1.0),
                 size=18, bold=True, color=NAVY)
        add_text(s, title, left + Cm(0.4), top + Cm(1.4), cw - Cm(0.5), Cm(2.0),
                 size=12, color=GRAY)


def slide_agenda(prs, idx, total):
    s = add_blank_slide(prs)
    page_frame(s, "全天议程", "Full-Day Agenda", idx, total)

    rows = [
        ("13:30–13:35", "开幕致辞 · 姚志勇教授"),
        ("13:35–14:00", "主旨演讲 1 · 白硕《大模型时代的底层架构突围与商业闭环》"),
        ("14:00–14:25", "主旨演讲 2 · 夏春博士《AI 如何改变我们的生活与投资》"),
        ("14:25–15:00", "圆桌一 · AI 硬核圆桌：技术突围与新质生产力落地"),
        ("15:00–15:30", "巅峰对话 · 云端生态与 AI 商业化\u201c天花板\u201d（腾讯/阿里/字节）"),
        ("15:30–16:20", "中场 · 陆家嘴金融城管弦乐团 · 时光音乐会"),
        ("16:20–16:40", "主旨演讲 3 · 王维军《特殊机遇/不良资产投资的新机遇》"),
        ("16:40–17:00", "主旨演讲 4 · 寇文红《从自然利率下降到 K 型经济》"),
        ("17:00–17:20", "主旨演讲 5 · 张露瑶博士《当经济学遇见 AI 与量子计算》"),
        ("17:20–17:55", "圆桌二 · 投资大圆桌：寻找新质资产风口"),
        ("17:55–18:10", "第三届 2026 人工智能商业化落地颁奖典礼"),
        ("18:10–20:30", "VIP 高端闭门晚宴（500 人圈层社交）"),
    ]
    top0 = Cm(4.8)
    row_h = Cm(0.95)
    for i, (t, e) in enumerate(rows):
        top = top0 + row_h * i
        bg = LIGHT if i % 2 == 0 else WHITE
        add_rect(s, Cm(1.5), top, Cm(30.8), row_h, fill=bg)
        add_text(s, t, Cm(1.8), top, Cm(5.5), row_h,
                 size=12, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, e, Cm(7.3), top, Cm(25), row_h,
                 size=12, color=GRAY, anchor=MSO_ANCHOR.MIDDLE)


def slide_packages_overview(prs, idx, total):
    s = add_blank_slide(prs)
    page_frame(s, "赞助方案总览", "Sponsorship Tiers", idx, total)

    tiers = [
        ("👑", "独家总冠名", "面议", "1 席", "大会冠名 + 主旨演讲 + 晚宴主桌", GOLD),
        ("💎", "钻石赞助", "￥50,000", "3 席", "圆桌席位 + 颁奖 + 5 张晚宴", NAVY),
        ("🥇", "铂金赞助", "￥30,000", "5 席", "音乐会鸣谢 + 专属展位", BLUE),
        ("🥈", "黄金赞助", "￥10,000", "8 席", "官网/大屏滚动 + 手册广告", ACCENT),
        ("🤝", "资源置换", "实物 / 服务", "限定", "用酒 / 伴手礼 / 影像 / 出行", GRAY),
        ("📍", "基础曝光", "￥1,500/位", "不限", "核心动线易拉宝 1 个", RGBColor(0x5C, 0x6B, 0x80)),
    ]
    cols = 3
    base_left = Cm(1.5)
    base_top = Cm(4.8)
    cw = Cm(10.3)
    ch = Cm(6.2)
    gx = Cm(0.4)
    gy = Cm(0.4)
    for i, (icon, name, price, slots, desc, color) in enumerate(tiers):
        r = i // cols
        c = i % cols
        left = base_left + (cw + gx) * c
        top = base_top + (ch + gy) * r
        add_rect(s, left, top, cw, ch, fill=LIGHT)
        add_rect(s, left, top, cw, Cm(1.5), fill=color)
        add_text(s, name, left + Cm(0.5), top + Cm(0.3), cw - Cm(1), Cm(1.0),
                 size=20, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, price, left + Cm(0.5), top + Cm(1.8), cw - Cm(1), Cm(1.4),
                 size=24, bold=True, color=NAVY)
        add_text(s, f"名额：{slots}", left + Cm(0.5), top + Cm(3.3), cw - Cm(1), Cm(0.8),
                 size=13, color=GRAY)
        add_text(s, desc, left + Cm(0.5), top + Cm(4.2), cw - Cm(1), Cm(2),
                 size=14, color=GRAY)


def slide_title_sponsor(prs, idx, total):
    s = add_blank_slide(prs)
    page_frame(s, "独家总冠名", "Title Sponsor · 仅 1 席", idx, total)

    add_rect(s, Cm(1.5), Cm(4.8), Cm(7), Cm(11.5), fill=GOLD)
    add_text(s, "👑", Cm(1.5), Cm(5.5), Cm(7), Cm(3),
             size=80, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, "TITLE\nSPONSOR", Cm(1.5), Cm(10), Cm(7), Cm(3),
             size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, "仅 1 席", Cm(1.5), Cm(13.8), Cm(7), Cm(1.5),
             size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    items = [
        ("品牌冠名", "大会全称升级为「【您的品牌】· 重构与突围 — 2026 AI 商业化峰会」"),
        ("主旨发声", "第一篇章 15 分钟独立主旨演讲席位"),
        ("圆桌定制", "独家指派 1 名高管进入 AI 硬核 / 投资大圆桌"),
        ("晚宴主场", "VIP 闭门晚宴主桌核心席位 3 人 + 联合主办方致祝酒辞"),
        ("全域霸屏", "主背景板顶级 Logo + 白皮书扉页 + 媒体通稿标题级露出"),
        ("资源沉淀", "双校长三角校友产业联盟战略合作伙伴永久入册"),
    ]
    top0 = Cm(5.0)
    row_h = Cm(1.85)
    label_w = Cm(4.5)
    for i, (k, v) in enumerate(items):
        top = top0 + row_h * i
        add_rect(s, Cm(9.2), top, Cm(23), Cm(1.6), fill=LIGHT)
        add_rect(s, Cm(9.2), top, Cm(0.15), Cm(1.6), fill=GOLD)
        add_text(s, k, Cm(9.5), top, label_w, Cm(1.6),
                 size=16, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, v, Cm(9.5) + label_w, top, Cm(18), Cm(1.6),
                 size=14, color=GRAY, anchor=MSO_ANCHOR.MIDDLE)


def slide_rights_matrix(prs, idx, total):
    s = add_blank_slide(prs)
    page_frame(s, "权益对比矩阵", "Rights Comparison Matrix", idx, total)

    header = ["权益项", "总冠名", "钻石", "铂金", "黄金", "基础"]
    rows = [
        ["大会冠名权", "●", "—", "—", "—", "—"],
        ["独立主旨演讲（15 min）", "●", "—", "—", "—", "—"],
        ["圆桌论坛对话席", "●", "● 限 1", "—", "—", "—"],
        ["年度颁奖授牌", "●", "●", "—", "—", "—"],
        ["音乐会环节鸣谢", "●", "●", "●", "—", "—"],
        ["主背景板 Logo 等级", "顶级", "钻石", "铂金", "黄金", "—"],
        ["议程手册广告", "扉页整版", "整版", "半版", "1/4 版", "尾页鸣谢"],
        ["白皮书署名", "封面+扉页", "1/2 版", "1/4 版", "—", "—"],
        ["VIP 晚宴入场券", "主桌 3 人", "5 张", "3 张", "1 张", "—"],
        ["重量级嘉宾 1V1", "全程定制", "2 位", "1 位", "—", "—"],
    ]
    base_left = Cm(1.5)
    base_top = Cm(4.8)
    col_w = [Cm(8), Cm(4.7), Cm(4.7), Cm(4.7), Cm(4.7), Cm(4.0)]
    row_h = Cm(0.95)

    # header
    x = base_left
    for i, h in enumerate(header):
        add_rect(s, x, base_top, col_w[i], row_h, fill=NAVY)
        add_text(s, h, x, base_top, col_w[i], row_h,
                 size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
        x += col_w[i]

    # body
    for r_idx, row in enumerate(rows):
        top = base_top + row_h * (r_idx + 1)
        bg = LIGHT if r_idx % 2 == 0 else WHITE
        x = base_left
        for c_idx, val in enumerate(row):
            add_rect(s, x, top, col_w[c_idx], row_h, fill=bg)
            color = NAVY if c_idx == 0 else (GOLD if val == "●" else GRAY)
            bold = c_idx == 0 or val == "●"
            align = PP_ALIGN.LEFT if c_idx == 0 else PP_ALIGN.CENTER
            add_text(s, val, x + (Cm(0.3) if c_idx == 0 else 0),
                     top, col_w[c_idx], row_h,
                     size=12, bold=bold, color=color, align=align,
                     anchor=MSO_ANCHOR.MIDDLE)
            x += col_w[c_idx]


def slide_barter(prs, idx, total):
    s = add_blank_slide(prs)
    page_frame(s, "资源置换合伙人", "Barter Partnerships · 零现金 · 精准置换", idx, total)

    items = [
        ("🍷", "晚宴指定用酒", "30 瓶起高端红酒 / 白酒",
         "主桌 VIP 席位 1 个 + 晚宴专属品鉴介绍 + 主背景板鸣谢"),
        ("🎁", "官方伴手礼", "全场 500 份 / VIP 精装 50 份",
         "企业定制福袋全场植入 + 定向意向客户数据回传"),
        ("📷", "独家影像服务", "双机位拍摄 + 后期精修",
         "暖场环节宣传片轮播 + 联合署名「官方影像合作伙伴」"),
        ("🚗", "官方指定出行", "VIP 接送 / 用车 ≥ 10 台",
         "「官方指定出行」标识 + 主背景板鸣谢 + 晚宴入场券 2 张"),
        ("📢", "媒体 / 渠道合作", "头部媒体通稿 / 流量包",
         "「战略媒体合作伙伴」标识 + 媒体合作专区露出"),
    ]
    top0 = Cm(4.8)
    row_h = Cm(2.1)
    for i, (icon, name, req, rights) in enumerate(items):
        top = top0 + row_h * i + Cm(0.15) * i
        add_rect(s, Cm(1.5), top, Cm(30.8), Cm(2.1), fill=LIGHT)
        add_rect(s, Cm(1.5), top, Cm(2.5), Cm(2.1), fill=NAVY)
        add_text(s, icon, Cm(1.5), top, Cm(2.5), Cm(2.1),
                 size=32, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, name, Cm(4.2), top + Cm(0.2), Cm(8), Cm(0.9),
                 size=17, bold=True, color=NAVY)
        add_text(s, "投入：" + req, Cm(4.2), top + Cm(1.1), Cm(8), Cm(0.9),
                 size=12, color=GRAY)
        add_text(s, "权益：" + rights, Cm(13), top + Cm(0.5), Cm(19), Cm(1.5),
                 size=13, color=GRAY, anchor=MSO_ANCHOR.MIDDLE)


def slide_timeline(prs, idx, total):
    s = add_blank_slide(prs)
    page_frame(s, "合作时间线", "Cooperation Timeline", idx, total)

    nodes = [
        ("即日起 – 5/18", "意向确认 · 合同 · 物料"),
        ("5/19 – 5/25", "物料制作 · 议程定稿 · 媒体预热"),
        ("5 月底 – 大会前 3 日", "主视觉 / 大屏 / 邀请函上线 · 锁定桌签"),
        ("峰会当日", "8h 全场曝光 + 500 人 VIP 晚宴"),
        ("大会后 7 日内", "媒体通稿 · 白皮书 · 数据回传"),
    ]
    base_left = Cm(1.8)
    base_top = Cm(8.5)
    box_w = Cm(6)
    box_h = Cm(4.5)
    gap = Cm(0.5)

    # connecting line
    add_rect(s, Cm(2.5), Cm(10.6), Cm(29.5), Cm(0.1), fill=NAVY)

    for i, (date, desc) in enumerate(nodes):
        left = base_left + (box_w + gap) * i
        # circle marker
        circle = s.shapes.add_shape(MSO_SHAPE.OVAL,
                                    left + box_w / 2 - Cm(0.5),
                                    Cm(10.4), Cm(1.0), Cm(1.0))
        circle.fill.solid()
        circle.fill.fore_color.rgb = GOLD
        circle.line.color.rgb = NAVY
        # number
        add_text(s, str(i + 1), left + box_w / 2 - Cm(0.5),
                 Cm(10.4), Cm(1.0), Cm(1.0),
                 size=14, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # box above (date) and below (desc)
        add_rect(s, left, base_top, box_w, Cm(1.6), fill=NAVY)
        add_text(s, date, left, base_top, box_w, Cm(1.6),
                 size=14, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        add_rect(s, left, Cm(11.8), box_w, Cm(3.2), fill=LIGHT)
        add_text(s, desc, left + Cm(0.3), Cm(12), box_w - Cm(0.6), Cm(2.8),
                 size=13, color=GRAY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def slide_cta(prs, idx, total):
    s = add_blank_slide(prs)
    SW = Cm(33.867)
    SH = Cm(19.05)
    add_rect(s, 0, 0, SW, SH, fill=NAVY)
    add_rect(s, 0, Cm(8.0), SW, Cm(0.1), fill=GOLD)

    add_text(s, "立即锁定席位", Cm(2), Cm(3.5), Cm(30), Cm(2.5),
             size=56, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, "Lock Your Seat Today", Cm(2), Cm(6.0), Cm(30), Cm(1.2),
             size=22, color=GOLD, align=PP_ALIGN.CENTER)

    add_text(s, "席位倒计时   ·   仅余 30 个深度合作席位 + 1 席独家总冠名",
             Cm(2), Cm(9.5), Cm(30), Cm(1.2),
             size=20, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, "确认截止：2026 年 5 月 18 日", Cm(2), Cm(11.0), Cm(30), Cm(1.2),
             size=20, bold=True, color=GOLD, align=PP_ALIGN.CENTER)

    box_w = Cm(9)
    box_h = Cm(3.5)
    base_left = Cm(3.4)
    base_top = Cm(13.5)
    items = [
        ("赞助合作", "组委会招商组"),
        ("联系电话", "________________"),
        ("官方邮箱", "________________"),
    ]
    gap = Cm(0.5)
    for i, (k, v) in enumerate(items):
        left = base_left + (box_w + gap) * i
        add_rect(s, left, base_top, box_w, box_h, fill=BLUE)
        add_rect(s, left, base_top, Cm(0.15), box_h, fill=GOLD)
        add_text(s, k, left + Cm(0.5), base_top + Cm(0.4), box_w - Cm(1), Cm(1),
                 size=14, color=GOLD)
        add_text(s, v, left + Cm(0.5), base_top + Cm(1.4), box_w - Cm(1), Cm(1.6),
                 size=20, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)


def main() -> None:
    prs = Presentation()
    prs.slide_width = Cm(33.867)
    prs.slide_height = Cm(19.05)

    builders = [
        slide_overview,
        slide_value_props,
        slide_audience,
        slide_speakers,
        slide_agenda,
        slide_packages_overview,
        slide_title_sponsor,
        slide_rights_matrix,
        slide_barter,
        slide_timeline,
    ]
    total = len(builders) + 2  # plus cover + cta
    slide_cover(prs, total)
    for i, b in enumerate(builders, start=1):
        b(prs, i, total)
    slide_cta(prs, total, total)

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    prs.save(OUT_PATH)
    print(f"[OK] PPT saved -> {OUT_PATH} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
