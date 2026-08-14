# -*- coding: utf-8 -*-
"""白皮书配套简报 PPT（中心研究文稿第三号）。"""
import os
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt, Emu

OUT = "/workspace/whitepaper/跨境电商与中国企业出海白皮书-简报.pptx"
LOGO = "/workspace/whitepaper/assets/logo_fudan_hprc.png"
CHARTS = "/workspace/whitepaper/assets/charts"
BLUE = RGBColor(0x0E, 0x4E, 0x9B)
RED = RGBColor(0xC8, 0x10, 0x2E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0x5B, 0x61, 0x6B)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
TEAL = RGBColor(0x1A, 0x7A, 0x6D)
LIGHT = RGBColor(0xF4, 0xF7, 0xFB)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
FONT = "微软雅黑"


def add_bg(slide, color=WHITE):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def bar(slide, y=0):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, y, prs.slide_width, Inches(0.08))
    s.fill.solid()
    s.fill.fore_color.rgb = BLUE
    s.line.fill.background()


def txt(slide, l, t, w, h, text, size=18, bold=False, color=DARK, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    return box


def footer(slide, page):
    txt(slide, 0.5, 7.15, 10, 0.28,
        "复旦大学住房政策研究中心 · 研究文稿第三号  FDU-HPRC-WP-2026-03",
        size=10, color=GRAY)
    txt(slide, 11.6, 7.15, 1.3, 0.28, str(page), size=10, color=GRAY,
        align=PP_ALIGN.RIGHT)


def card(slide, l, t, w, h, fill=LIGHT):
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(l), Inches(t), Inches(w), Inches(h))
    box.fill.solid()
    box.fill.fore_color.rgb = fill
    box.line.fill.background()
    box.adjustments[0] = 0.08
    return box


# 1 封面
s = prs.slides.add_slide(BLANK)
add_bg(s, WHITE)
bar(s)
s.shapes.add_picture(LOGO, Inches(0.7), Inches(0.4), width=Inches(6.4))
txt(s, 0.7, 2.35, 12, 0.4, "中心研究文稿 · 第三号", 16, False, GRAY)
txt(s, 0.7, 2.85, 12, 0.7, "跨境电商与中国企业出海", 34, True, BLUE)
txt(s, 0.7, 3.6, 12, 0.5, "空间重构、规则跃迁与二〇三〇展望", 22, True, BLUE)
txt(s, 0.7, 4.3, 12, 0.4, "从货物通达到系统扎根  ·  结合 2026 年最新形势", 16, False, GRAY)
txt(s, 0.7, 6.35, 12, 0.35, "FDU-HPRC-WP-2026-03    二〇二六年八月 · 上海", 14, False, GRAY)

# 2 一句话
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.6, 0.35, 12, 0.5, "2026 年的一句话判断", 24, True, BLUE)
card(s, 0.6, 1.3, 12.1, 2.4)
txt(s, 0.9, 1.7, 11.5, 1.7,
    "免税的船票已经作废。下一班船的名字叫规则、品牌、产能和空间。",
    26, True, BLUE)
txt(s, 0.7, 4.0, 12, 2.4,
    "美国已于 2025 年 8 月 29 日暂停各国 800 美元低值免税，2026 年 2 月延续。\n"
    "欧盟自 2026 年 7 月 1 日起取消 150 欧元关税豁免，2028 年转入全额征税。\n"
    "国内政策主线转为：跨境电商加海外仓“扩容升级、规范有序发展”。",
    16, False, DARK)
footer(s, 2)

# 3 五个核心判断
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.6, 0.25, 12, 0.45, "五个核心判断", 24, True, BLUE)
items = [
    "2026 年不是周期底部，而是范式切换：低值直邮的制度基础已经拆除。",
    "国内政策是扩容升级 + 规范有序，而不是继续放大流量红利。",
    "企业出海从卖货过海进入系统扎根：产能、品牌、体系同时发生。",
    "空间是被低估的战略变量：海外仓是货的住房，园区是产的住房，公寓是人的住房。",
    "2030 年的胜负手是一套可复制的全球经营系统，而不是再多一个站点。",
]
for i, t in enumerate(items):
    y = 0.95 + i * 1.12
    card(s, 0.6, y, 12.1, 1.0)
    txt(s, 0.85, y + 0.28, 11.6, 0.55, f"{i+1}.  {t}", 16, False, DARK)
footer(s, 3)

# 4 规模
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.22, 12, 0.4, "规模仍在增长，结构已经改写", 22, True, BLUE)
s.shapes.add_picture(f"{CHARTS}/chart01_cbec_scale.png", Inches(0.3), Inches(0.75), width=Inches(7.3))
txt(s, 7.8, 1.2, 5.0, 5.2,
    "海关总署：\n2025 年进出口 2.75 万亿元\n较 2020 年 +69.7%\n\n"
    "2026 年一季度 6184.6 亿元\n出口 4735.5 亿元\n进口 1449.1 亿元\n\n"
    "件数增长将受免税取消抑制，\n增量转向品牌、仓网与新兴市场。",
    16, False, DARK)
footer(s, 4)

# 5 规则时间轴
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.22, 12, 0.4, "全球规则时间轴：2025—2030", 22, True, BLUE)
s.shapes.add_picture(f"{CHARTS}/chart03_rule_timeline.png", Inches(0.35), Inches(0.75), width=Inches(12.5))
footer(s, 5)

# 6 出海四层
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.22, 12, 0.4, "企业出海：贸易、品牌、产能、体系四层叠加", 22, True, BLUE)
s.shapes.add_picture(f"{CHARTS}/chart02_outbound_layers.png", Inches(0.4), Inches(0.85), width=Inches(12.4))
footer(s, 6)

# 7 市场分层
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.22, 12, 0.4, "市场分层：欧美利润、南向增长、国内总部", 22, True, BLUE)
s.shapes.add_picture(f"{CHARTS}/chart05_market_layers.png", Inches(0.55), Inches(0.7), width=Inches(12.2))
footer(s, 7)

# 8 空间四层
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.22, 12, 0.4, "住房政策视角：货、产、人、城", 22, True, BLUE)
s.shapes.add_picture(f"{CHARTS}/chart07_space_four.png", Inches(0.2), Inches(0.7), width=Inches(7.4))
txt(s, 7.7, 1.3, 5.1, 5.2,
    "货之居所：海外仓从放货的房子，\n变成商品全生命周期住房。\n\n"
    "产之居所：海外园区是微型城市开发，\n厂房、仓储、宿舍必须一起设计。\n\n"
    "人之居所：外派公寓、属地职住、\n国内跨境团队的人才住房。\n\n"
    "城之居所：综试区空间底线，\n把决策、品牌、数据留在国内。",
    15, False, DARK)
footer(s, 8)

# 9 海外仓
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.22, 12, 0.4, "海外仓：从数量扩张转向网络与功能升级", 22, True, BLUE)
s.shapes.add_picture(f"{CHARTS}/chart04_overseas_warehouse.png", Inches(0.3), Inches(0.75), width=Inches(7.5))
txt(s, 7.95, 1.15, 4.9, 5.4,
    "政府工作报告把\n“跨境电商加海外仓”\n写成一个模式。\n\n"
    "行业样本仓已超 6200 个；\n另一口径：专注跨境电商仓\n超 1800 个、面积超 2200 万㎡。\n\n"
    "2030 年比的不是仓最多，\n而是多区域仓网 + 退货翻新\n+ 绿色循环 + 智能分拨。",
    15, False, DARK)
footer(s, 9)

# 10 四维合规
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.22, 12, 0.4, "四维合规：2030 年的市场门票", 22, True, BLUE)
s.shapes.add_picture(f"{CHARTS}/chart06_compliance.png", Inches(0.35), Inches(0.75), width=Inches(12.5))
footer(s, 10)

# 11 2030 情景
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.22, 12, 0.4, "展望 2030：三条情景", 22, True, BLUE)
s.shapes.add_picture(f"{CHARTS}/chart08_2030_scenarios.png", Inches(0.25), Inches(0.7), width=Inches(7.6))
txt(s, 8.0, 1.1, 4.9, 5.5,
    "基准 A：4.2—4.8 万亿元\n转型完成，中速增长\n\n"
    "乐观 B：约 5.5 万亿元\n规则对接 + 南向超预期\n\n"
    "承压 C：3.3—3.6 万亿元\n壁垒叠加，头部集中\n\n"
    "建议：按 A 配能力，\n按 C 做底线，按 B 做期权。",
    16, False, DARK)
footer(s, 11)

# 12 七条趋势
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.6, 0.25, 12, 0.4, "七条高确定性趋势（无论哪种情景）", 22, True, BLUE)
trends = [
    "合规是门票，不是选修课",
    "品牌替代铺货，成为分配机制",
    "“中国 + N”常态化，但 N 必须可辩护",
    "海外仓进入存量优化与功能升级",
    "人工智能从工具变成经营系统",
    "绿色与数据成为新的关税",
    "空间能力决定全球化能否闭环",
]
for i, t in enumerate(trends):
    col = i % 2
    row = i // 2
    x = 0.55 + col * 6.3
    y = 0.95 + row * 1.35
    if i == 6:
        x = 0.55
        y = 0.95 + 3 * 1.35
        card(s, x, y, 12.2, 1.15)
        txt(s, x + 0.25, y + 0.32, 11.7, 0.55, f"{i+1}.  {t}", 18, True, BLUE)
    else:
        card(s, x, y, 6.05, 1.2)
        txt(s, x + 0.25, y + 0.35, 5.6, 0.55, f"{i+1}.  {t}", 16, True, BLUE)
footer(s, 12)

# 13 政策建议
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.22, 12, 0.4, "政策建议：企业、城市、国家三端发力", 22, True, BLUE)
s.shapes.add_picture(f"{CHARTS}/chart09_policy.png", Inches(0.4), Inches(0.85), width=Inches(12.5))
footer(s, 13)

# 14 结束
s = prs.slides.add_slide(BLANK)
add_bg(s, WHITE)
bar(s)
s.shapes.add_picture(LOGO, Inches(0.7), Inches(0.45), width=Inches(6.2))
txt(s, 0.7, 2.5, 12, 1.4,
    "把货物、工厂、人才和城市的“居所”配置好，\n才能把走出去做成扎下去。",
    24, True, BLUE)
txt(s, 0.7, 4.5, 12, 0.8,
    "全文见《跨境电商与中国企业出海：空间重构、规则跃迁与二〇三〇展望》\n"
    "复旦大学住房政策研究中心  ·  FDU-HPRC-WP-2026-03",
    16, False, GRAY)
txt(s, 0.7, 6.3, 12, 0.35, "二〇二六年八月 · 上海", 14, False, GRAY)

os.makedirs("/workspace/下载版本", exist_ok=True)
prs.save(OUT)
prs.save("/workspace/下载版本/跨境电商与中国企业出海白皮书-简报.pptx")
print("saved", OUT)
