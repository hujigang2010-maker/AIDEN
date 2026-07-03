# -*- coding: utf-8 -*-
"""生成《迪拜专题沙龙——内容框架与议程》PPT。

设计要求：
- 背景以紫色为主色调；
- 主标准字体为微软雅黑（辅以金色/浅紫等辅助元素）；
- 内容组织围绕三视角：观众视角 / 热点视角 / 利益视角。
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---- 主题色（紫色系 + 金色辅助）----
PURPLE_BG = RGBColor(0x2A, 0x14, 0x4A)     # 深紫背景
PURPLE2 = RGBColor(0x3A, 0x22, 0x63)       # 次级紫
PURPLE3 = RGBColor(0x4A, 0x2E, 0x7A)       # 卡片紫
CARD = RGBColor(0x3E, 0x25, 0x6B)          # 卡片底色
GOLD = RGBColor(0xE0, 0xB9, 0x5C)          # 金色辅助
GOLD_LIGHT = RGBColor(0xF0, 0xD9, 0x9E)
VIOLET_LIGHT = RGBColor(0xC7, 0xB2, 0xF0)  # 浅紫辅助
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREY = RGBColor(0xC3, 0xB8, 0xDA)
ACCENT_PINK = RGBColor(0xE0, 0x6C, 0xB0)
ACCENT_CYAN = RGBColor(0x5E, 0xD3, 0xD3)
ACCENT_GREEN = RGBColor(0x7E, 0xD0, 0x9A)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

FONT = "微软雅黑"          # 主标准字体
FONT_ASSIST = "微软雅黑"   # 辅助沿用同族，保证跨平台一致

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
BLANK = prs.slide_layouts[6]


def add_slide():
    s = prs.slides.add_slide(BLANK)
    bg = s.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = PURPLE_BG
    return s


def box(slide, x, y, w, h, fill=None, line=None, radius=0.06):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    sh.adjustments[0] = radius
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1.2)
    sh.shadow.inherit = False
    return sh


def rect(slide, x, y, w, h, fill):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


def text(slide, x, y, w, h, runs, size=18, color=WHITE, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing=1.15,
         space_after=6, font=None):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    if isinstance(runs, str):
        runs = [runs]
    first = True
    for para in runs:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = Pt(space_after)
        if isinstance(para, tuple):
            t, opts = para
        else:
            t, opts = para, {}
        r = p.add_run()
        r.text = t
        f = r.font
        f.name = opts.get("font", font or FONT)
        f.size = Pt(opts.get("size", size))
        f.color.rgb = opts.get("color", color)
        f.bold = opts.get("bold", bold)
    return tb


def header(slide, title, subtitle=None, tag=None):
    rect(slide, Inches(0), Inches(0), Inches(0.18), SLIDE_H, GOLD)
    text(slide, Inches(0.55), Inches(0.32), Inches(11.0), Inches(0.8),
         title, size=30, color=WHITE, bold=True)
    if subtitle:
        text(slide, Inches(0.58), Inches(1.02), Inches(11.0), Inches(0.5),
             subtitle, size=15, color=GOLD_LIGHT)
    if tag:
        box(slide, Inches(11.1), Inches(0.38), Inches(1.85), Inches(0.5), fill=PURPLE3, line=GOLD)
        text(slide, Inches(11.1), Inches(0.45), Inches(1.85), Inches(0.4),
             tag, size=13, color=GOLD_LIGHT, align=PP_ALIGN.CENTER)
    rect(slide, Inches(0.55), Inches(1.55), Inches(12.2), Pt(1.6), GOLD)


def bullet_card(slide, x, y, w, h, title, lines, title_color=GOLD_LIGHT,
                body_size=13.5, title_size=16, accent=None, fill=CARD):
    box(slide, x, y, w, h, fill=fill)
    if accent:
        rect(slide, x, y, Inches(0.12), h, accent)
    text(slide, x + Inches(0.25), y + Inches(0.15), w - Inches(0.5), Inches(0.55),
         title, size=title_size, color=title_color, bold=True)
    paras = [("• " + l, {}) for l in lines]
    text(slide, x + Inches(0.25), y + Inches(0.72), w - Inches(0.5), h - Inches(0.9),
         paras, size=body_size, color=WHITE, space_after=5)


def footer(slide, page):
    text(slide, Inches(0.55), Inches(7.05), Inches(9), Inches(0.35),
         "迪拜身份规划与中东投资机会专题沙龙 · 内部策划文件", size=10, color=GREY)
    text(slide, Inches(12.3), Inches(7.05), Inches(0.7), Inches(0.35),
         str(page), size=10, color=GREY, align=PP_ALIGN.RIGHT)


page = 0


def new_slide(title=None, subtitle=None, tag=None):
    global page
    page += 1
    s = add_slide()
    if title:
        header(s, title, subtitle, tag)
    footer(s, page)
    return s


# ================== S1 封面 ==================
s = new_slide()
rect(s, Inches(0), Inches(0), SLIDE_W, SLIDE_H, PURPLE_BG)
# 顶部与底部金色装饰条
rect(s, Inches(0), Inches(0), SLIDE_W, Pt(6), GOLD)
rect(s, Inches(0), Inches(5.95), SLIDE_W, Pt(2), GOLD)
text(s, Inches(0.9), Inches(1.4), Inches(11.5), Inches(0.6),
     "系列高净值客户闭门沙龙 · 第一期策划案", size=18, color=GOLD_LIGHT)
text(s, Inches(0.9), Inches(2.05), Inches(11.8), Inches(1.9),
     [("出海新坐标：迪拜", {"size": 52, "bold": True, "color": WHITE}),
      ("身份规划 × 中东地缘 × 算力投资机会", {"size": 30, "bold": True, "color": GOLD})],
     space_after=14)
# 三视角标签
labels = [("观众视角", ACCENT_CYAN), ("热点视角", ACCENT_PINK), ("利益视角", GOLD)]
x = Inches(0.9)
for t, c in labels:
    box(s, x, Inches(4.15), Inches(2.3), Inches(0.62), fill=PURPLE2, line=c)
    text(s, x, Inches(4.26), Inches(2.3), Inches(0.42), t, size=16, color=c,
         bold=True, align=PP_ALIGN.CENTER)
    x += Inches(2.5)
text(s, Inches(0.9), Inches(5.05), Inches(11.5), Inches(0.9),
     [("15人闭门圆桌 | 一对一信任建立 | 系列化持续培育", {"size": 15, "color": GREY}),
      ("拟联合主办：中东投资联盟 · 复旦大学住房政策研究中心 · 上海市杨浦区科技企业联合会 · 国际会客厅 · 特邀 夏春先生", {"size": 13, "color": GREY})],
     space_after=6)
text(s, Inches(0.9), Inches(6.2), Inches(11.5), Inches(0.5),
     "策划版本 V2.0（三视角重构 · 紫金主题）· 2026年7月 · 上海", size=13, color=VIOLET_LIGHT)
footer(s, 1)
page = 1

# ================== S2 三视角策划总纲 ==================
s = new_slide("策划总纲：三视角驱动", "以观众为起点、以热点为引信、以利益为落点", "总 纲")
persp = [
    ("观众视角", ACCENT_CYAN, "为什么愿意来？",
     ["精准邀约：只请对的人，15人闭门、拒绝大会式喧闹",
      "低门槛入场：先讲“钱与合规”，不硬推移民、无压迫感",
      "高规格体验：学术+专家+外籍高管背书，配实时翻译",
      "稀缺感与社交价值：同层圈层链接，值得来、值得晒"]),
    ("热点视角", ACCENT_PINK, "凭什么被关注？",
     ["美股与全球资产收益：人人都在看的赚钱话题做钩子",
      "CRS信息交换落地：高净值人群当下最真实的焦虑",
      "AI算力浪潮：DeepSeek/豆包出海 + 迪拜算力洼地",
      "世界杯/油价等话题暖场，规避过期地缘硬蹭"]),
    ("利益视角", GOLD, "能赚到什么？",
     ["身份即资产：迪拜零个税、美元资产、保值+隔离",
      "捡漏机会：中东窗口期资产折价、法拍别墅7-8折",
      "产业红利：迪拜算力工厂成本洼地的投资入口",
      "对渠道方：补齐迪拜产品线、共享高净值客源"]),
]
x = Inches(0.55)
for name, c, q, lines in persp:
    box(s, x, Inches(1.8), Inches(4.0), Inches(4.7), fill=CARD)
    rect(s, x, Inches(1.8), Inches(4.0), Inches(0.62), c)
    text(s, x, Inches(1.87), Inches(4.0), Inches(0.5), name, size=18, color=PURPLE_BG,
         bold=True, align=PP_ALIGN.CENTER)
    text(s, x + Inches(0.25), Inches(2.55), Inches(3.5), Inches(0.4), q,
         size=14, color=GOLD_LIGHT, bold=True)
    paras = [("• " + l, {}) for l in lines]
    text(s, x + Inches(0.25), Inches(3.05), Inches(3.55), Inches(3.3),
         paras, size=13, color=WHITE, space_after=8)
    x += Inches(4.17)

# ================== S3 观众视角 ==================
s = new_slide("观众视角：如何让对的人愿意来", "获客的本质是“值得来 + 来了有收获 + 无压力”", "观 众")
bullet_card(s, Inches(0.55), Inches(1.8), Inches(4.0), Inches(4.7),
            "谁来（精准客群）", accent=ACCENT_CYAN, title_color=ACCENT_CYAN,
            lines=["高净值终端客户 8-10 人：可投资资产≥500万美元",
                   "渠道合作方 5-7 人：移民公司、私行、律所税务师",
                   "宁缺毋滥：每位来宾都可命名、可跟进",
                   "来源：私行转介、老客带新、家办、合作单位输送"])
bullet_card(s, Inches(4.70), Inches(1.8), Inches(4.0), Inches(4.7),
            "为何来（吸引点）", accent=GOLD, title_color=GOLD_LIGHT,
            lines=["主题钩子：美股收益 + CRS避坑，而非硬讲移民",
                   "嘉宾阵容：经济学家 + 高校研究 + 开发商外籍高管",
                   "稀缺闭门：15席限定，圈层社交价值高",
                   "即时获得感：现场做CRS自测与资产诊断"])
bullet_card(s, Inches(8.85), Inches(1.8), Inches(4.0), Inches(4.7),
            "怎么来（邀约路径）", accent=ACCENT_PINK, title_color=VIOLET_LIGHT,
            lines=["定向邀请函 + 合作单位联名背书，提升可信度",
                   "一对一电话邀约，讲清“为你解决什么问题”",
                   "老客户/渠道带客机制，给予转介绍礼遇",
                   "会前发预热资料，降低陌生感、锁定出席"])
box(s, Inches(0.55), Inches(6.58), Inches(12.35), Inches(0.5), fill=PURPLE2, line=ACCENT_CYAN, radius=0.3)
text(s, Inches(0.55), Inches(6.63), Inches(12.35), Inches(0.4),
     "观众视角一句话：把“被推销的顾虑”变成“怕错过的期待”。", size=14, color=GOLD_LIGHT,
     bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# ================== S4 热点视角 ==================
s = new_slide("热点视角：借势当下、精准卡点", "热点是流量入口，但要“借壳包装、卡准窗口”", "热 点")
hot = [
    ("美股与全球收益", ACCENT_GREEN, ["近一年美股/纳指收益复盘，全民关注的赚钱话题",
                                     "自然引出：收益在海外，如何合规申报？"]),
    ("CRS 信息交换", ACCENT_PINK, ["100+国家和地区自动交换金融账户信息已落地",
                                   "高净值人群最真实的税务焦虑，痛点即流量"]),
    ("AI 算力浪潮", ACCENT_CYAN, ["DeepSeek、豆包等中国开源模型出海热度高",
                                 "迪拜最低电价+美元结算=算力工厂洼地叙事"]),
    ("暖场轻热点", GOLD, ["世界杯、油价、金价等话题破冰，降低商业感",
                         "地缘冲突仅作背景板：3月流量红利已过，不硬蹭"]),
]
y = Inches(1.85)
for i, (t, c, lines) in enumerate(hot):
    yy = y + (Inches(1.15) + Inches(0.15)) * (i // 2)
    xx = Inches(0.55) + (Inches(6.05) + Inches(0.25)) * (i % 2)
    box(s, xx, yy, Inches(6.05), Inches(1.15), fill=CARD)
    rect(s, xx, yy, Inches(0.12), Inches(1.15), c)
    text(s, xx + Inches(0.3), yy + Inches(0.08), Inches(5.6), Inches(0.4), t,
         size=15.5, color=c, bold=True)
    paras = [("• " + l, {}) for l in lines]
    text(s, xx + Inches(0.3), yy + Inches(0.5), Inches(5.6), Inches(0.6),
         paras, size=12, color=WHITE, space_after=2, line_spacing=1.0)
box(s, Inches(0.55), Inches(5.35), Inches(12.35), Inches(1.15), fill=PURPLE2, line=GOLD)
text(s, Inches(0.85), Inches(5.45), Inches(11.8), Inches(0.4),
     "卡点原则（合规红线）", size=15, color=GOLD_LIGHT, bold=True)
text(s, Inches(0.85), Inches(5.9), Inches(11.8), Inches(0.5),
     [("• 敏感话题（美股、移民）以“全球资产配置 / 国际教育”外壳包装；不承诺避税，只讲合法规划与申报优化；热点只做引子，落点始终回到迪拜身份与投资。", {})],
     size=12.5, color=WHITE)

# ================== S5 利益视角 ==================
s = new_slide("利益视角：为参会人创造真实收益", "让投资人、渠道方、主办方都算得清这笔账", "利 益")
bullet_card(s, Inches(0.55), Inches(1.8), Inches(6.05), Inches(2.35),
            "对投资人 / 高净值客户", accent=GOLD, title_color=GOLD_LIGHT,
            lines=["身份即资产：迪拜零个人所得税、美元资产、保值抗贬值",
                   "捡漏窗口：中东冲突周期资产折价、法拍别墅7-8折机会",
                   "产业入口：迪拜算力工厂/数据中心的成本洼地投资机会",
                   "风险对冲：资产隔离 + 人身安全 + 全球身份的实用性"])
bullet_card(s, Inches(6.85), Inches(1.8), Inches(6.05), Inches(2.35),
            "对渠道合作方", accent=ACCENT_CYAN, title_color=ACCENT_CYAN,
            lines=["补齐迪拜产品线，填补现有渠道动力不足的空白",
                   "共享高净值客源，联合活动降低各自获客成本",
                   "考察团与签约的佣金/分成机制，形成长期收益",
                   "联名背书提升品牌，沉淀可持续的客户资产"])
box(s, Inches(0.55), Inches(4.35), Inches(12.35), Inches(2.15), fill=PURPLE3, line=GOLD)
text(s, Inches(0.85), Inches(4.48), Inches(11.8), Inches(0.45),
     "“收益可视化”表达法：把机会讲成可计算的账", size=16, color=GOLD_LIGHT, bold=True)
cases = [("迪拜购房+签证", "美元资产\n零个税持有\n签证续期"),
         ("法拍别墅捡漏", "7-8折入手\n打通银行-法院-业主链条"),
         ("算力资产投资", "低电价+美元结算\n应用层出海红利"),
         ("土耳其对比", "彻底隔离但\n需揭示汇率/缩水风险")]
x = Inches(0.85)
for t, d in cases:
    box(s, x, Inches(5.0), Inches(2.9), Inches(1.35), fill=CARD)
    text(s, x + Inches(0.15), Inches(5.08), Inches(2.6), Inches(0.4), t,
         size=13.5, color=GOLD, bold=True)
    for j, line in enumerate(d.split("\n")):
        text(s, x + Inches(0.15), Inches(5.46) + Inches(0.26) * j, Inches(2.6), Inches(0.3),
             line, size=11, color=WHITE)
    x += Inches(3.05)

# ================== S6 目标客群与转化路径 ==================
s = new_slide("目标客群画像与转化路径", "小而精：一场沙龙 15 人，全部可命名、可跟进", "客 群")
bullet_card(s, Inches(0.55), Inches(1.8), Inches(6.05), Inches(2.35),
            "客群 A：高净值终端客户（8-10人）", accent=GOLD,
            lines=["可投资资产 ≥ 500万美元，已有或计划配置海外资产",
                   "痛点：CRS税务追缴风险、资产透明化、子女教育与出行便利",
                   "来源：私行客户经理转介、老客户带新、家族办公室"])
bullet_card(s, Inches(6.85), Inches(1.8), Inches(6.05), Inches(2.35),
            "客群 B：渠道合作方（5-7人）", accent=ACCENT_CYAN,
            lines=["移民公司负责人：补足其迪拜产品线（现有渠道动力不足）",
                   "金融机构/私行/三方财富：需要新的客户活动抓手",
                   "律所税务师：CRS合规咨询的天然入口"])
box(s, Inches(0.55), Inches(4.45), Inches(12.35), Inches(2.05), fill=PURPLE2, line=GOLD)
text(s, Inches(0.85), Inches(4.6), Inches(12.0), Inches(0.45),
     "四步转化路径", size=16, color=GOLD_LIGHT, bold=True)
steps = [("① 沙龙触达", "热点内容引流\n建立初步信任"),
         ("② 一对一面谈", "CRS/资产诊断\n明确身份需求"),
         ("③ 迪拜考察团", "实地看房+算力园区\n开发商高管接待"),
         ("④ 签约落地", "购房+签证办理\n转介绍裂变")]
x = Inches(0.85)
for t, d in steps:
    box(s, x, Inches(5.1), Inches(2.75), Inches(1.2), fill=CARD)
    text(s, x + Inches(0.15), Inches(5.2), Inches(2.45), Inches(0.4), t,
         size=14, color=GOLD, bold=True)
    for j, line in enumerate(d.split("\n")):
        text(s, x + Inches(0.15), Inches(5.6) + Inches(0.28) * j, Inches(2.45), Inches(0.3),
             line, size=11.5, color=WHITE)
    x += Inches(3.0)

# ================== S7 内容框架总览 ==================
s = new_slide("沙龙内容框架：三大模块", "以痛点开场、以对比立信、以机会收口", "框 架")
mods = [
    ("模块一 · 破冰与痛点（30分钟）· 对应热点视角", GOLD,
     ["美股收益复盘：用真实收益数据抓注意力（热点话题包装）",
      "CRS与税务合规：信息交换机制、追缴案例、高净值人群的真实焦虑",
      "引出核心命题：为什么需要第二身份与资产隔离"]),
    ("模块二 · 身份规划方案对比（40分钟）· 对应利益视角", GOLD_LIGHT,
     ["土耳其购房入籍：40万美元门槛、护照可改名、三年后可售房",
      "土耳其风险提示：里拉贬值、房价虚高（实际常达50万+美元）、资产缩水",
      "迪拜黄金签证：购房获签、零个税、美元资产、人身安全、离岸金融中心",
      "结论导向：土耳其重“隔离”、迪拜重“实用+保值”，可组合配置"]),
    ("模块三 · 中东机遇与算力投资（40分钟）· 对应热点+利益", ACCENT_CYAN,
     ["中东地缘：冲突周期（打-停-再打）与油价联动，如何把握窗口期",
      "迪拜算力基建：中东最低电价 + 美元结算 + 政策友好 = 算力工厂洼地",
      "中美分工：美国主导底层芯片，中国应用层（DeepSeek、豆包开源模型）出海",
      "开发商高管现场分享（外籍嘉宾 + 实时翻译）：迪拜房产与园区一手信息"]),
]
y = Inches(1.85)
for title, c, lines in mods:
    box(s, Inches(0.55), y, Inches(12.35), Inches(1.55), fill=CARD)
    rect(s, Inches(0.55), y, Inches(0.12), Inches(1.55), c)
    text(s, Inches(0.9), y + Inches(0.08), Inches(11.8), Inches(0.4),
         title, size=15, color=c, bold=True)
    paras = [("• " + l, {}) for l in lines]
    text(s, Inches(0.9), y + Inches(0.45), Inches(11.8), Inches(1.05),
         paras, size=12, color=WHITE, space_after=2, line_spacing=1.05)
    y += Inches(1.68)

# ================== S8 模块一详解 ==================
s = new_slide("模块一详解：以 CRS 痛点切入", "先讲客户关心的钱，再讲客户害怕的税", "模块一")
bullet_card(s, Inches(0.55), Inches(1.85), Inches(6.05), Inches(4.5),
            "开场：美股收益与全球配置（15分钟）", accent=ACCENT_GREEN,
            lines=["近一年美股/纳指收益复盘，对比国内资产表现",
                   "抛出问题：收益在海外，申报怎么办？",
                   "世界杯、油价等热点做暖场话题，降低商业感",
                   "话术要点：不直接讲移民，先讲“钱放哪里、怎么合规”"])
bullet_card(s, Inches(6.85), Inches(1.85), Inches(6.05), Inches(4.5),
            "主体：CRS 信息交换与税务风险（15分钟）", accent=ACCENT_PINK,
            lines=["CRS机制科普：100+国家和地区自动交换金融账户信息",
                   "真实案例：境外账户被交换后收到税务核查通知的处理过程",
                   "补税与滞纳金测算演示：让风险“可计算、可感知”",
                   "过渡句：合规的终极方案，是让资产拥有一个新的身份",
                   "合规红线：不承诺避税，只讲“合法规划与申报优化”"])

# ================== S9 迪拜 vs 土耳其对比 ==================
s = new_slide("核心对比：迪拜黄金签证 vs 土耳其购房入籍", "沙龙最重要的一页：用事实对比替客户做决策", "模块二")
rows = [
    ("对比维度", "土耳其购房入籍", "迪拜黄金签证", True),
    ("投资门槛", "40万美元（评估价），实际售价常50万+美元", "购房约200万迪拉姆（约54万美元）起", False),
    ("身份性质", "直接入籍拿护照，不影响中国国籍", "10年长期居留签证（可续）", False),
    ("资产隔离", "护照可改名，实现彻底隔离（核心卖点）", "居留身份，隔离效果弱于换护照", False),
    ("房产绑定", "三年后可售房，护照仍有效", "房产出售则签证失效，需持续持有", False),
    ("主要风险", "里拉贬值、房价虚高，资产可能明显缩水", "房产流动性锁定；无入籍路径", False),
    ("资产币种", "里拉资产，汇率风险高", "美元挂钩（迪拉姆盯住美元），保值性强", False),
    ("渠道现状", "移民公司主推（房价做高、佣金高）", "佣金低、渠道动力不足 → 市场空白即机会", False),
]
y = Inches(1.8)
col_x = [Inches(0.55), Inches(2.75), Inches(7.85)]
col_w = [Inches(2.2), Inches(5.1), Inches(5.05)]
for r, (c0, c1, c2, is_head) in enumerate(rows):
    hh = Inches(0.52) if is_head else Inches(0.58)
    for ci, (cx, cw, val) in enumerate(zip(col_x, col_w, (c0, c1, c2))):
        if is_head:
            fill = GOLD
        elif ci == 0:
            fill = PURPLE3
        else:
            fill = CARD if r % 2 else PURPLE2
        rect(s, cx, y, cw, hh, fill)
        tc = PURPLE_BG if is_head else (GOLD_LIGHT if ci == 0 else WHITE)
        text(s, cx + Inches(0.12), y + Inches(0.05), cw - Inches(0.24), hh - Inches(0.08),
             val, size=13 if is_head else 11.5, color=tc,
             bold=is_head or ci == 0, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
    y += hh + Emu(20000)
text(s, Inches(0.55), y + Inches(0.06), Inches(12.3), Inches(0.5),
     "沙龙结论话术：追求“彻底隔离”选土耳其（接受资产缩水），追求“保值+实用+安全”选迪拜，高净值客户可双身份组合配置。",
     size=13, color=GOLD, bold=True)

# ================== S10 中东地缘与油价 ==================
s = new_slide("模块三详解 A：中东地缘与投资窗口", "冲突是周期性的，恐慌是布局的朋友", "模块三")
bullet_card(s, Inches(0.55), Inches(1.85), Inches(6.05), Inches(4.5),
            "地缘局势框架", accent=ACCENT_PINK,
            lines=["以伊冲突等呈周期性：打—停—再打，避免用单次事件下结论",
                   "冲突推高油价 → 海湾国家财政盈余增加 → 主权基金加大对外投资",
                   "迪拜的独特定位：中东的“安全岛”与资金避风港",
                   "内容提示：地缘只作背景板，当前热度已过、不做硬主题"])
bullet_card(s, Inches(6.85), Inches(1.85), Inches(6.05), Inches(4.5),
            "对客户的三个投资含义", accent=GOLD,
            lines=["油价上行周期 = 海湾流动性充裕期，资产价格有支撑",
                   "冲突窗口期资产折价，是逆向布局迪拜房产/基建的时点",
                   "人身与资产安全叙事：迪拜治安、法治与自由区制度背书",
                   "话术：把“乱”讲成“别人的乱、迪拜的机会”"])

# ================== S11 算力投资 ==================
s = new_slide("模块三详解 B：迪拜算力基础设施机会", "中东最低电价 × 美元结算 × 中国应用层出海", "模块三")
cards = [
    ("成本优势", ACCENT_GREEN, ["迪拜为中东电价最低地区，电力成本是算力工厂第一成本项",
                                "土地与能源政策友好，自由区税收优惠",
                                "美元结算，收入与成本币种匹配，无汇兑损耗"]),
    ("产业格局", ACCENT_CYAN, ["底层芯片仍由美国主导（英伟达等），采购与合规是关键变量",
                              "中国在应用层表现突出：DeepSeek、豆包等开源模型生态",
                              "“美国芯片 + 中国应用 + 中东能源与资本”的三方组合叙事"]),
    ("落地路径", GOLD, ["与迪拜开发商/园区合作，考察算力园区与数据中心地块",
                       "以专题内容切入：制作“迪拜算力”系列内容持续获客",
                       "远期：组织客户参与算力资产投资（份额/基建/REITs形式）"]),
]
x = Inches(0.55)
for t, c, lines in cards:
    bullet_card(s, x, Inches(1.85), Inches(4.0), Inches(4.5), t, lines, accent=c, title_color=c)
    x += Inches(4.17)

# ================== S12 沙龙议程 ==================
s = new_slide("首期沙龙议程（半日闭门场）", "建议周六下午 14:00–18:00 · 国际会客厅（拟） · 限 15 席", "议 程")
agenda = [
    ("13:30–14:00", "签到与茶歇", "一对一破冰，主办方逐一认识每位来宾", "全体工作人员"),
    ("14:00–14:10", "开场致辞", "沙龙定位与系列活动介绍，合作单位亮相", "主办方 + 中东投资联盟代表"),
    ("14:10–14:40", "模块一：美股收益与CRS合规", "美股复盘引入 → CRS机制与追缴案例 → 引出身份命题", "主讲人（主办方）"),
    ("14:40–15:20", "模块二：迪拜 vs 土耳其身份方案", "两大方案深度对比：门槛、隔离、风险、币种", "主讲人 + 移民合作方"),
    ("15:20–15:35", "茶歇交流", "工作人员定向撮合：客户 × 嘉宾一对一", "全体"),
    ("15:35–16:05", "特邀分享：全球资产配置视角", "宏观与资产配置框架下的中东机会（拟邀 夏春先生）", "夏春先生（拟）"),
    ("16:05–16:35", "模块三：中东地缘与迪拜算力投资", "冲突周期与油价 → 算力工厂逻辑 → 中国模型出海", "主讲人 + 杨浦科技企联代表"),
    ("16:35–17:05", "开发商高管圆桌", "迪拜开发商外籍高管（如埃及籍负责人）分享，配实时翻译", "开发商高管 + 翻译"),
    ("17:05–17:45", "自由交流与一对一咨询", "按意向分组：身份规划桌 / 算力投资桌 / 房产桌", "全体嘉宾"),
    ("17:45–18:00", "收尾与下期预告", "发放资料包，登记迪拜考察团意向", "主办方"),
]
y = Inches(1.78)
for i, (t, item, desc, who) in enumerate(agenda):
    fill = CARD if i % 2 else PURPLE2
    rect(s, Inches(0.55), y, Inches(1.7), Inches(0.485), fill)
    rect(s, Inches(2.25), y, Inches(2.5), Inches(0.485), fill)
    rect(s, Inches(4.75), y, Inches(5.55), Inches(0.485), fill)
    rect(s, Inches(10.3), y, Inches(2.6), Inches(0.485), fill)
    text(s, Inches(0.65), y + Inches(0.03), Inches(1.6), Inches(0.42), t,
         size=11.5, color=GOLD_LIGHT, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(2.35), y + Inches(0.03), Inches(2.4), Inches(0.42), item,
         size=11.5, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(4.85), y + Inches(0.03), Inches(5.4), Inches(0.42), desc,
         size=10.5, color=GREY, anchor=MSO_ANCHOR.MIDDLE, line_spacing=0.95)
    text(s, Inches(10.4), y + Inches(0.03), Inches(2.45), Inches(0.42), who,
         size=10.5, color=GOLD_LIGHT, anchor=MSO_ANCHOR.MIDDLE, line_spacing=0.95)
    y += Inches(0.515)

# ================== S13 系列化活动规划 ==================
s = new_slide("系列化活动规划（滚动四期）", "一场沙龙不做转化闭环，系列活动才做", "系 列")
series = [
    ("第一期 · 出海新坐标：迪拜", "身份规划 + CRS合规 + 算力机会总览（本方案）", "目标：建立信任池，筛出高意向 5 人"),
    ("第二期 · 全球身份组合配置", "土耳其/迪拜/其他方案组合策略；律所税务师联合场", "目标：一对一诊断 8 场，锁定 3 组签约意向"),
    ("第三期 · 迪拜算力与新基建专场", "算力园区项目路演；杨浦科技企联联合，邀科技企业主", "目标：形成考察团名单 10 人"),
    ("第四期 · 迪拜实地考察团", "看房 + 算力园区 + 开发商高管接待 + 自由区注册咨询", "目标：现场签约 2–3 单，形成案例素材"),
]
y = Inches(1.9)
for t, c, goal in series:
    box(s, Inches(0.55), y, Inches(12.35), Inches(1.13), fill=CARD)
    text(s, Inches(0.85), y + Inches(0.08), Inches(4.2), Inches(0.9), t,
         size=15, color=GOLD, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(5.1), y + Inches(0.08), Inches(4.6), Inches(0.95), c,
         size=12, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(9.8), y + Inches(0.08), Inches(2.9), Inches(0.95), goal,
         size=11.5, color=VIOLET_LIGHT, anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(1.26)

# ================== S14 合作单位与分工 ==================
s = new_slide("合作单位与分工（拟联络推进）", "五方资源互补：产业+学术+政企+专家+场地", "合 作")
partners = [
    ("中东投资联盟", "联合主办 / 产业资源", ["提供中东项目资源与开发商对接", "联合背书提升沙龙公信力", "输送中东投资意向客户"]),
    ("复旦大学住房政策研究中心", "学术支持", ["提供房产市场研究视角与数据背书", "拟邀研究员做“全球房产配置”客观分享", "提升内容严肃性，弱化销售感"]),
    ("上海市杨浦区科技企业联合会", "科技企业客群", ["组织科技企业主参与算力专场", "对接算力/AI企业出海需求", "政企资源与场地备选"]),
    ("夏春先生", "特邀经济学家（拟）", ["全球资产配置与宏观视角主题分享", "为“美股+CRS”模块提供专业站台", "提升活动规格与传播话题性"]),
    ("国际会客厅", "场地 / 涉外资源", ["提供高规格涉外活动场地", "协助外籍嘉宾接待与实时翻译", "涉外活动流程合规把关"]),
]
y = Inches(1.8)
for name, role, items in partners:
    box(s, Inches(0.55), y, Inches(12.35), Inches(0.95), fill=CARD)
    text(s, Inches(0.85), y + Inches(0.06), Inches(3.0), Inches(0.8), name,
         size=14.5, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(3.9), y + Inches(0.06), Inches(2.1), Inches(0.8), role,
         size=11.5, color=GOLD, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(6.1), y + Inches(0.06), Inches(6.6), Inches(0.85),
         [("· " + i, {}) for i in items],
         size=10, color=GREY, space_after=0, line_spacing=0.95, anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(1.04)

# ================== S15 执行细节与合规 ==================
s = new_slide("现场执行与风险合规要点", "小沙龙拼的是细节和分寸感", "执 行")
bullet_card(s, Inches(0.55), Inches(1.85), Inches(6.05), Inches(4.5),
            "现场执行细节", accent=ACCENT_CYAN,
            lines=["外籍嘉宾（迪拜开发商高管）全程配实时翻译，提前对稿",
                   "座位按“客户+对应顾问”穿插安排，保证一对一覆盖",
                   "PPT开场页用美股/世界杯热点，中东内容后置",
                   "资料包：对比表 + CRS自测表 + 迪拜项目册 + 考察团报名表",
                   "现场设三张主题桌：身份规划 / 算力投资 / 房产配置"])
bullet_card(s, Inches(6.85), Inches(1.85), Inches(6.05), Inches(4.5),
            "风险与合规提示", accent=ACCENT_PINK,
            lines=["敏感话题（美股、移民）借“全球资产配置”“国际教育”外壳包装",
                   "不承诺免税/避税，统一口径为“合法税务规划与申报”",
                   "土耳其方案风险（汇率、房价虚高）必须如实揭示，建立长期信任",
                   "涉外活动报备与宣传物料审核，由国际会客厅协助把关",
                   "热点营销卡点：地缘冲突话题仅作背景，避免蹭过期热点"])

# ================== S16 KPI与下一步 ==================
s = new_slide("目标 KPI 与下一步行动", "所有动作以“考察团成行、签约落地”为终点", "行 动")
kpis = [("到场人数", "15人\n（客户10 + 渠道5）"), ("一对一面谈", "≥ 8 场\n（会后一周内）"),
        ("考察团意向", "≥ 5 人\n登记报名"), ("签约转化", "2–3 单\n（系列期内）")]
x = Inches(0.55)
for t, v in kpis:
    box(s, x, Inches(1.85), Inches(2.95), Inches(1.5), fill=CARD, line=GOLD)
    text(s, x, Inches(1.98), Inches(2.95), Inches(0.4), t, size=14, color=GOLD_LIGHT,
         bold=True, align=PP_ALIGN.CENTER)
    for j, line in enumerate(v.split("\n")):
        text(s, x, Inches(2.38) + Inches(0.36) * j, Inches(2.95), Inches(0.4), line,
             size=15 if j == 0 else 11.5, color=WHITE if j == 0 else GREY,
             bold=(j == 0), align=PP_ALIGN.CENTER)
    x += Inches(3.13)
todos = [
    ("T+1周", "分别联络五家合作单位，确认联合主办意向与分工（详见Excel推进表）"),
    ("T+1周", "筛选上海本地迪拜开发商代表（外籍高管优先），确认圆桌嘉宾与翻译"),
    ("T+2周", "定稿沙龙PPT讲稿（美股+CRS开场版本），完成合规审核"),
    ("T+2周", "确认场地（国际会客厅）、日期与15人邀约名单，发出定向邀请函"),
    ("T+3周", "彩排走场：讲者对稿、翻译演练、一对一撮合分工到人"),
    ("T+4周", "第一期沙龙执行 → 会后48小时内完成全部一对一回访"),
]
y = Inches(3.75)
for t, d in todos:
    rect(s, Inches(0.55), y, Inches(1.3), Inches(0.44), GOLD)
    text(s, Inches(0.55), y + Inches(0.04), Inches(1.3), Inches(0.36), t,
         size=12, color=PURPLE_BG, bold=True, align=PP_ALIGN.CENTER)
    rect(s, Inches(1.85), y, Inches(11.05), Inches(0.44), CARD)
    text(s, Inches(2.05), y + Inches(0.04), Inches(10.7), Inches(0.36), d,
         size=12, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.53)

# ================== S17 封底 ==================
s = new_slide()
rect(s, Inches(0), Inches(3.4), SLIDE_W, Pt(2), GOLD)
text(s, Inches(0.9), Inches(2.3), Inches(11.5), Inches(1.0),
     "出海新坐标 · 迪拜", size=44, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
text(s, Inches(0.9), Inches(3.7), Inches(11.5), Inches(1.5),
     [("观众为起点 · 热点为引信 · 利益为落点", {"size": 20, "color": GOLD_LIGHT}),
      ("配套执行文件：《迪拜沙龙-执行推进表.xlsx》（三视角策略 / 议程 / 合作单位联络 / 嘉宾筛选 / 待办跟踪）", {"size": 14, "color": VIOLET_LIGHT})],
     align=PP_ALIGN.CENTER, space_after=12)

out = "/workspace/deliverables/迪拜专题沙龙-内容框架与议程.pptx"
os.makedirs(os.path.dirname(out), exist_ok=True)
prs.save(out)
print("Saved:", out, "slides:", len(prs.slides._sldIdLst))
