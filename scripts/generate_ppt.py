# -*- coding: utf-8 -*-
"""生成《迪拜专题沙龙——内容框架与议程》PPT。"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ---- 主题色（迪拜金 + 深海军蓝）----
NAVY = RGBColor(0x0B, 0x1F, 0x3A)      # 深蓝背景
NAVY2 = RGBColor(0x13, 0x2B, 0x4E)     # 次级深蓝
GOLD = RGBColor(0xC9, 0xA2, 0x4B)      # 迪拜金
GOLD_LIGHT = RGBColor(0xE3, 0xC8, 0x8A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREY = RGBColor(0xB9, 0xC4, 0xD4)
CARD = RGBColor(0x1A, 0x35, 0x5C)      # 卡片底色
ACCENT_RED = RGBColor(0xC0, 0x50, 0x46)
ACCENT_GREEN = RGBColor(0x4E, 0x9A, 0x6F)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

FONT = "Noto Sans CJK SC"

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
BLANK = prs.slide_layouts[6]


def add_slide():
    s = prs.slides.add_slide(BLANK)
    bg = s.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = NAVY
    return s


def box(slide, x, y, w, h, fill=None, line=None):
    from pptx.enum.shapes import MSO_SHAPE
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    sh.adjustments[0] = 0.06
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
    from pptx.enum.shapes import MSO_SHAPE
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


def text(slide, x, y, w, h, runs, size=18, color=WHITE, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing=1.15,
         space_after=6):
    """runs: str 或 [(text, {size,color,bold}), ...] 的段落列表"""
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
        f.name = FONT
        f.size = Pt(opts.get("size", size))
        f.color.rgb = opts.get("color", color)
        f.bold = opts.get("bold", bold)
    return tb


def header(slide, title, subtitle=None, tag=None):
    rect(slide, Inches(0), Inches(0), Inches(0.18), SLIDE_H, GOLD)
    text(slide, Inches(0.55), Inches(0.32), Inches(11.5), Inches(0.8),
         title, size=30, color=WHITE, bold=True)
    if subtitle:
        text(slide, Inches(0.58), Inches(1.02), Inches(11.5), Inches(0.5),
             subtitle, size=15, color=GOLD_LIGHT)
    if tag:
        b = box(slide, Inches(11.1), Inches(0.38), Inches(1.85), Inches(0.5), fill=CARD, line=GOLD)
        text(slide, Inches(11.1), Inches(0.44), Inches(1.85), Inches(0.4),
             tag, size=13, color=GOLD_LIGHT, align=PP_ALIGN.CENTER)
    rect(slide, Inches(0.55), Inches(1.55), Inches(12.2), Pt(1.6), GOLD)


def bullet_card(slide, x, y, w, h, title, lines, title_color=GOLD_LIGHT,
                body_size=13.5, title_size=16):
    box(slide, x, y, w, h, fill=CARD)
    text(slide, x + Inches(0.25), y + Inches(0.15), w - Inches(0.5), Inches(0.5),
         title, size=title_size, color=title_color, bold=True)
    paras = [("• " + l, {}) for l in lines]
    text(slide, x + Inches(0.25), y + Inches(0.68), w - Inches(0.5), h - Inches(0.85),
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
rect(s, Inches(0), Inches(0), SLIDE_W, SLIDE_H, NAVY)
rect(s, Inches(0), Inches(5.9), SLIDE_W, Pt(2), GOLD)
text(s, Inches(0.9), Inches(1.5), Inches(11.5), Inches(0.6),
     "系列高净值客户闭门沙龙 · 第一期策划案", size=18, color=GOLD_LIGHT)
text(s, Inches(0.9), Inches(2.15), Inches(11.8), Inches(1.9),
     [("出海新坐标：迪拜", {"size": 52, "bold": True, "color": WHITE}),
      ("身份规划 × 中东地缘 × 算力投资机会", {"size": 30, "bold": True, "color": GOLD})],
     space_after=14)
text(s, Inches(0.9), Inches(4.35), Inches(11.5), Inches(1.2),
     [("15人闭门圆桌 | 一对一信任建立 | 系列化持续培育", {"size": 17, "color": GREY}),
      ("拟联合主办：中东投资联盟 · 复旦大学住房政策研究中心 · 上海市杨浦区科技企业联合会 · 国际会客厅", {"size": 14, "color": GREY}),
      ("特邀嘉宾（拟）：夏春先生", {"size": 14, "color": GREY})],
     space_after=8)
text(s, Inches(0.9), Inches(6.15), Inches(11.5), Inches(0.5),
     "策划版本 V1.0 · 2026年7月 · 上海", size=13, color=GREY)
footer(s, 1)
page = 1

# ================== S2 策划逻辑总览 ==================
s = new_slide("策划逻辑总览", "从客户痛点出发，双主线内容，小沙龙精准转化", "总 览")
bullet_card(s, Inches(0.55), Inches(1.85), Inches(4.0), Inches(4.6), "为什么做（Why）",
            ["CRS信息交换落地，高净值客户面临税务追缴与资产透明化压力",
             "客户对“资产彻底隔离 + 人身安全 + 身份实用性”有强烈刚需",
             "中东地缘冲突推高油价，市场关注度周期性上升，存在布局窗口",
             "移民渠道普遍推土耳其（佣金高），迪拜方案存在市场空白"])
bullet_card(s, Inches(4.70), Inches(1.85), Inches(4.0), Inches(4.6), "讲什么（What）",
            ["主线一：身份规划——迪拜黄金签证 vs 土耳其购房入籍深度对比",
             "主线二：产业投资——迪拜算力基建（中东最低电价+美元结算）",
             "切入点：美股收益复盘 + CRS痛点，避免硬讲中东（热度已过）",
             "嵌入热点话题（美股、世界杯）提升现场吸引力与传播性"])
bullet_card(s, Inches(8.85), Inches(1.85), Inches(4.0), Inches(4.6), "怎么做（How）",
            ["15人左右闭门小沙龙，拒绝大型峰会（转化效率低）",
             "客群：高意向客户 + 渠道方（移民公司、金融机构）",
             "邀请迪拜开发商外籍高管（如埃及籍负责人）背书，配实时翻译",
             "系列化运营：每期一个主题，持续培育、分层转化"])

# ================== S3 目标客群与转化路径 ==================
s = new_slide("目标客群画像与转化路径", "小而精：一场沙龙 15 人，全部可命名、可跟进", "客 群")
bullet_card(s, Inches(0.55), Inches(1.85), Inches(6.05), Inches(2.4), "客群 A：高净值终端客户（8-10人）",
            ["可投资资产 ≥ 500万美元，已有或计划配置海外资产",
             "痛点：CRS税务追缴风险、资产透明化、子女教育与出行便利",
             "来源：私行客户经理转介、老客户带新、家族办公室"])
bullet_card(s, Inches(6.85), Inches(1.85), Inches(6.05), Inches(2.4), "客群 B：渠道合作方（5-7人）",
            ["移民公司负责人：补足其迪拜产品线（现有渠道动力不足）",
             "金融机构/私行/三方财富：需要新的客户活动抓手",
             "律所税务师：CRS合规咨询的天然入口"])
box(s, Inches(0.55), Inches(4.5), Inches(12.35), Inches(2.0), fill=NAVY2, line=GOLD)
text(s, Inches(0.85), Inches(4.65), Inches(12.0), Inches(0.45),
     "四步转化路径", size=16, color=GOLD_LIGHT, bold=True)
steps = [("① 沙龙触达", "热点内容引流\n建立初步信任"),
         ("② 一对一面谈", "CRS/资产诊断\n明确身份需求"),
         ("③ 迪拜考察团", "实地看房+算力园区\n开发商高管接待"),
         ("④ 签约落地", "购房+签证办理\n转介绍裂变")]
x = Inches(0.85)
for i, (t, d) in enumerate(steps):
    box(s, x, Inches(5.15), Inches(2.75), Inches(1.15), fill=CARD)
    text(s, x + Inches(0.15), Inches(5.24), Inches(2.45), Inches(0.4), t,
         size=14, color=GOLD, bold=True)
    for j, line in enumerate(d.split("\n")):
        text(s, x + Inches(0.15), Inches(5.62) + Inches(0.28) * j, Inches(2.45), Inches(0.3),
             line, size=11.5, color=WHITE)
    x += Inches(3.0)

# ================== S4 内容框架总览 ==================
s = new_slide("沙龙内容框架：三大模块", "以痛点开场、以对比立信、以机会收口", "框 架")
mods = [
    ("模块一 · 破冰与痛点（30分钟）", GOLD,
     ["美股收益复盘：用真实收益数据抓注意力（热点话题包装）",
      "CRS与税务合规：信息交换机制、追缴案例、高净值人群的真实焦虑",
      "引出核心命题：为什么需要第二身份与资产隔离"]),
    ("模块二 · 身份规划方案对比（40分钟）", GOLD_LIGHT,
     ["土耳其购房入籍：40万美元门槛、护照可改名、三年后可售房",
      "土耳其风险提示：里拉贬值、房价虚高（实际常达50万+美元）、资产缩水",
      "迪拜黄金签证：购房获签、零个税、美元资产、人身安全、离岸金融中心",
      "结论导向：土耳其重“隔离”、迪拜重“实用+保值”，可组合配置"]),
    ("模块三 · 中东机遇与算力投资（40分钟）", ACCENT_GREEN,
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
         title, size=15.5, color=c, bold=True)
    paras = [("• " + l, {}) for l in lines]
    text(s, Inches(0.9), y + Inches(0.45), Inches(11.8), Inches(1.05),
         paras, size=12, color=WHITE, space_after=2, line_spacing=1.05)
    y += Inches(1.68)

# ================== S5 模块一详解 ==================
s = new_slide("模块一详解：以 CRS 痛点切入", "先讲客户关心的钱，再讲客户害怕的税", "模块一")
bullet_card(s, Inches(0.55), Inches(1.85), Inches(6.05), Inches(4.5), "开场：美股收益与全球配置（15分钟）",
            ["近一年美股/纳指收益复盘，对比国内资产表现",
             "抛出问题：收益在海外，申报怎么办？",
             "世界杯、油价等热点做暖场话题，降低商业感",
             "话术要点：不直接讲移民，先讲“钱放哪里、怎么合规”"])
bullet_card(s, Inches(6.85), Inches(1.85), Inches(6.05), Inches(4.5), "主体：CRS 信息交换与税务风险（15分钟）",
            ["CRS机制科普：100+国家和地区自动交换金融账户信息",
             "真实案例：境外账户被交换后收到税务核查通知的处理过程",
             "补税与滞纳金测算演示：让风险“可计算、可感知”",
             "过渡句：合规的终极方案，是让资产拥有一个新的身份",
             "合规红线：不承诺避税，只讲“合法规划与申报优化”"])

# ================== S6 迪拜 vs 土耳其对比 ==================
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
    fill0 = GOLD if is_head else NAVY2
    fill = GOLD if is_head else (CARD if r % 2 else NAVY2)
    for ci, (cx, cw, val) in enumerate(zip(col_x, col_w, (c0, c1, c2))):
        rc = rect(s, cx, y, cw, hh, fill0 if ci == 0 else fill)
        tc = NAVY if is_head else (GOLD_LIGHT if ci == 0 else WHITE)
        text(s, cx + Inches(0.12), y + Inches(0.05), cw - Inches(0.24), hh - Inches(0.08),
             val, size=13 if is_head else 11.5, color=tc,
             bold=is_head or ci == 0, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
    y += hh + Emu(20000)
text(s, Inches(0.55), y + Inches(0.06), Inches(12.3), Inches(0.5),
     "沙龙结论话术：追求“彻底隔离”选土耳其（接受资产缩水），追求“保值+实用+安全”选迪拜，高净值客户可双身份组合配置。",
     size=13, color=GOLD, bold=True)

# ================== S7 中东地缘与油价 ==================
s = new_slide("模块三详解 A：中东地缘与投资窗口", "冲突是周期性的，恐慌是布局的朋友", "模块三")
bullet_card(s, Inches(0.55), Inches(1.85), Inches(6.05), Inches(4.5), "地缘局势框架",
            ["以伊冲突等呈周期性：打—停—再打，避免用单次事件下结论",
             "冲突推高油价 → 海湾国家财政盈余增加 → 主权基金加大对外投资",
             "迪拜的独特定位：中东的“安全岛”与资金避风港",
             "内容提示：地缘只作背景板，当前热度已过、不做硬主题"])
bullet_card(s, Inches(6.85), Inches(1.85), Inches(6.05), Inches(4.5), "对客户的三个投资含义",
            ["油价上行周期 = 海湾流动性充裕期，资产价格有支撑",
             "冲突窗口期资产折价，是逆向布局迪拜房产/基建的时点",
             "人身与资产安全叙事：迪拜治安、法治与自由区制度背书",
             "话术：把“乱”讲成“别人的乱、迪拜的机会”"])

# ================== S8 算力投资 ==================
s = new_slide("模块三详解 B：迪拜算力基础设施机会", "中东最低电价 × 美元结算 × 中国应用层出海", "模块三")
cards = [
    ("成本优势", ["迪拜为中东电价最低地区，电力成本是算力工厂第一成本项",
                  "土地与能源政策友好，自由区税收优惠",
                  "美元结算，收入与成本币种匹配，无汇兑损耗"]),
    ("产业格局", ["底层芯片仍由美国主导（英伟达等），采购与合规是关键变量",
                  "中国在应用层表现突出：DeepSeek、豆包等开源模型生态",
                  "“美国芯片 + 中国应用 + 中东能源与资本”的三方组合叙事"]),
    ("落地路径", ["与迪拜开发商/园区合作，考察算力园区与数据中心地块",
                  "以专题内容切入：制作“迪拜算力”系列内容持续获客",
                  "远期：组织客户参与算力资产投资（份额/基建/REITs形式）"]),
]
x = Inches(0.55)
for t, lines in cards:
    bullet_card(s, x, Inches(1.85), Inches(4.0), Inches(4.5), t, lines)
    x += Inches(4.17)

# ================== S9 沙龙议程 ==================
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
    fill = CARD if i % 2 else NAVY2
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

# ================== S10 系列化活动规划 ==================
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
         size=11.5, color=GOLD_LIGHT, anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(1.26)

# ================== S11 合作单位与分工 ==================
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
    fill = CARD
    box(s, Inches(0.55), y, Inches(12.35), Inches(0.95), fill=fill)
    text(s, Inches(0.85), y + Inches(0.06), Inches(3.0), Inches(0.8), name,
         size=14.5, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(3.9), y + Inches(0.06), Inches(2.1), Inches(0.8), role,
         size=11.5, color=GOLD, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(6.1), y + Inches(0.06), Inches(6.6), Inches(0.85),
         [("· " + " ".join(items), {})] if False else [("· " + i, {}) for i in items],
         size=10, color=GREY, space_after=0, line_spacing=0.95, anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(1.04)

# ================== S12 执行细节 ==================
s = new_slide("现场执行与风险合规要点", "小沙龙拼的是细节和分寸感", "执 行")
bullet_card(s, Inches(0.55), Inches(1.85), Inches(6.05), Inches(4.5), "现场执行细节",
            ["外籍嘉宾（迪拜开发商高管）全程配实时翻译，提前对稿",
             "座位按“客户+对应顾问”穿插安排，保证一对一覆盖",
             "PPT开场页用美股/世界杯热点，中东内容后置",
             "资料包：对比表 + CRS自测表 + 迪拜项目册 + 考察团报名表",
             "现场设三张主题桌：身份规划 / 算力投资 / 房产配置"])
bullet_card(s, Inches(6.85), Inches(1.85), Inches(6.05), Inches(4.5), "风险与合规提示",
            ["敏感话题（美股、移民）借“全球资产配置”“国际教育”外壳包装",
             "不承诺免税/避税，统一口径为“合法税务规划与申报”",
             "土耳其方案风险（汇率、房价虚高）必须如实揭示，建立长期信任",
             "涉外活动报备与宣传物料审核，由国际会客厅协助把关",
             "热点营销卡点：地缘冲突话题仅作背景，避免蹭过期热点"])

# ================== S13 KPI与下一步 ==================
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
         size=12, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    rect(s, Inches(1.85), y, Inches(11.05), Inches(0.44), CARD)
    text(s, Inches(2.05), y + Inches(0.04), Inches(10.7), Inches(0.36), d,
         size=12, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.53)

# ================== S14 封底 ==================
s = new_slide()
rect(s, Inches(0), Inches(3.4), SLIDE_W, Pt(2), GOLD)
text(s, Inches(0.9), Inches(2.3), Inches(11.5), Inches(1.0),
     "出海新坐标 · 迪拜", size=44, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
text(s, Inches(0.9), Inches(3.7), Inches(11.5), Inches(1.5),
     [("以身份为锚，以算力为帆，以信任为舟", {"size": 20, "color": GOLD_LIGHT}),
      ("配套执行文件：《迪拜沙龙-执行推进表.xlsx》（议程 / 合作单位联络 / 嘉宾筛选 / 待办跟踪）", {"size": 14, "color": GREY})],
     align=PP_ALIGN.CENTER, space_after=12)

out = "/workspace/deliverables/迪拜专题沙龙-内容框架与议程.pptx"
os.makedirs(os.path.dirname(out), exist_ok=True)
prs.save(out)
print("Saved:", out, "slides:", len(prs.slides.__iter__.__self__._sldIdLst))
