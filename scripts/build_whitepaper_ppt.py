# -*- coding: utf-8 -*-
"""白皮书配套简报 PPT（中心研究文稿第二号）。"""
import os
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

OUT = "/workspace/whitepaper/WAIC2026人工智能产业空间白皮书-简报.pptx"
LOGO = "/workspace/whitepaper/assets/brand/logo_rchp_print.png"
CHARTS = "/workspace/whitepaper/assets/charts"
BLUE = RGBColor(0x0E, 0x4E, 0x9B)
RED = RGBColor(0xC8, 0x10, 0x2E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0x5B, 0x61, 0x6B)
DARK = RGBColor(0x1A, 0x1A, 0x1A)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


def add_bg(slide, color=WHITE):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def bar(slide, y=0):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, y, prs.slide_width, Inches(0.08))
    s.fill.solid()
    s.fill.fore_color.rgb = BLUE
    s.line.fill.background()


def txt(slide, l, t, w, h, text, size=18, bold=False, color=DARK, align=PP_ALIGN.LEFT,
        font="微软雅黑"):
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
    run.font.name = font
    return box


def footer(slide, page):
    txt(slide, 0.5, 7.15, 8, 0.28,
        "复旦大学住房政策研究中心 · 研究文稿第二号  FDU-HPRC-WP-2026-02",
        size=10, color=GRAY)
    txt(slide, 11.6, 7.15, 1.3, 0.28, str(page), size=10, color=GRAY,
        align=PP_ALIGN.RIGHT)


def img(slide, path, l, t, w):
    slide.shapes.add_picture(path, Inches(l), Inches(t), width=Inches(w))


# 1 封面
s = prs.slides.add_slide(BLANK)
add_bg(s, WHITE)
bar(s)
img(s, LOGO, 0.7, 0.45, 6.2)
txt(s, 0.7, 2.5, 12, 0.5, "中心研究文稿 · 第二号", 18, False, GRAY)
txt(s, 0.7, 3.15, 12, 1.0, "WAIC2026 人工智能产业空间白皮书", 32, True, BLUE)
txt(s, 0.7, 4.3, 12, 0.5, "AI 与产业空间融合的趋势、格局与新范式", 18, False, GRAY)
txt(s, 0.7, 6.4, 12, 0.4, "FDU-HPRC-WP-2026-02    二〇二六年八月 · 上海", 14, False, GRAY)

# 2 核心判断
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.6, 0.3, 12, 0.5, "五个核心判断", 26, True, BLUE)
items = [
    "硬科技占半壁江山：机器人 24.2% + 算力 23.5% = 47.7%",
    "国产大模型并跑且更便宜：Top20 占 13 席，价差逾 4 倍",
    "青年把 AI 驯化为生活伙伴：搭子化、情绪外包、meme 化",
    "空间成为新的竞争变量：算力密度 / 场景开放度 / 空间柔性",
    "中美欧三条不可互换路径：中国优势在硬件、性价比与园区",
]
for i, t in enumerate(items):
    y = 1.15 + i * 1.05
    box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(y), Inches(12.1), Inches(0.9))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xF4, 0xF7, 0xFB)
    box.line.fill.background()
    txt(s, 0.8, y + 0.22, 11.6, 0.5, f"{i+1}.  {t}", 18, False, DARK)
footer(s, 2)

# 3 产业结构
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.25, 12, 0.45, "WAIC 2026 产业结构：硬科技占半壁江山", 22, True, BLUE)
img(s, f"{CHARTS}/chart01_exhibitor_industry.png", 0.4, 0.85, 7.4)
txt(s, 8.0, 1.4, 4.8, 5.2,
    "963 家参展商\n\n第一细分：具身智能 93 家\n第二细分：AI Agent 87 家\n\n基础大模型仅 20 家——\n增量已转向模型之上的应用\n与模型之下的硬件。",
    16, False, DARK)
footer(s, 3)

# 4 国内外对照
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.25, 12, 0.45, "中美欧三条路径：中国优势在空间与硬件", 22, True, BLUE)
img(s, f"{CHARTS}/chart22_cn_us_eu.png", 0.2, 0.8, 6.6)
img(s, f"{CHARTS}/chart10_llm_cn_vs_global.png", 6.7, 1.1, 6.3)
footer(s, 4)

# 5 国内四极
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.25, 12, 0.45, "国内四极分工：上海均衡、北京算力、广东硬件、浙江机器人", 20, True, BLUE)
img(s, f"{CHARTS}/chart16_regional_mix.png", 0.5, 0.85, 12.2)
footer(s, 5)

# 6 展馆地图
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.25, 12, 0.45, "展馆即产业地图：张江=底座，西岸=体验，世博=首发", 20, True, BLUE)
img(s, f"{CHARTS}/chart15_hall_industry.png", 0.3, 0.85, 12.6)
footer(s, 6)

# 7 大模型
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.25, 12, 0.45, "大模型格局：同样聪明，谁更便宜", 22, True, BLUE)
img(s, f"{CHARTS}/chart07_llm_top20.png", 0.2, 0.7, 6.5)
img(s, f"{CHARTS}/chart08_llm_price_perf.png", 6.7, 0.85, 6.3)
footer(s, 7)

# 8 青年观察
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.25, 12, 0.45, "青年视角：效率打底，情感与身份增值", 22, True, BLUE)
img(s, f"{CHARTS}/chart12_watcha_categories.png", 0.2, 0.8, 6.5)
img(s, f"{CHARTS}/chart13_watcha_top10.png", 6.7, 0.8, 6.3)
footer(s, 8)

# 9 具身智能
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.25, 12, 0.45, "具身智能：零售是中间地带，工业最拥挤", 22, True, BLUE)
img(s, f"{CHARTS}/chart11_embodied_scenes.png", 0.5, 1.0, 12.2)
txt(s, 0.6, 5.3, 12, 1.5,
    "银河通用便利店、穹彻药房/洗衣、擎朗具身社区、智元半导体产线、乐聚 1:1 工厂——场景成为新的竞争货币。",
    16, False, DARK)
footer(s, 9)

# 10 园区雷达
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.25, 12, 0.45, "AI 原生园区：六要素与四步转型", 22, True, BLUE)
img(s, f"{CHARTS}/chart14_park_radar.png", 0.2, 0.7, 6.4)
txt(s, 6.8, 1.3, 5.8, 5.2,
    "四步转型\n\n1. 底座化：插电即训练\n2. 场景化：公区即试验场\n3. 智能体化：用 AI 管园区\n4. 资本化：算力 REITs / 场景入股\n\n办公新范式\n人机共生单元 · 四象限场景\n柔性冗余 · 环境即服务",
    16, False, DARK)
footer(s, 10)

# 11 杨浦实证
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.25, 12, 0.45, "杨浦实证：AI 企业为大学与社群支付租金溢价", 20, True, BLUE)
img(s, f"{CHARTS}/chart19_yangpu_industry.png", 0.15, 0.8, 6.5)
img(s, f"{CHARTS}/chart20_yangpu_plates.png", 6.6, 0.7, 6.5)
footer(s, 11)

# 12 四个原型
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.25, 12, 0.45, "上海四个 AI 产业空间原型", 22, True, BLUE)
protos = [
    ("底座型 · 张江", "算力芯片馆\n114 家展商中 80 家是基础设施\n产品：液冷 / 绿电 / 超节点"),
    ("体验型 · 西岸", "终端与内容体验\n25 场论坛 + 大模型 22 家\n产品：展示层 + 消费硬件"),
    ("试验型 · 杨浦", "大创智 24 家 AI 企业\n复兴岛空间智能论坛\n产品：中试 + 场景开放协议"),
    ("制度型 · 世博", "52% 论坛 + 治理主场\n世界人工智能合作组织\n产品：首发经济与标准发布"),
]
for i, (title, body) in enumerate(protos):
    x = 0.45 + (i % 4) * 3.2
    box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(1.2), Inches(3.0), Inches(5.3))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xF4, 0xF7, 0xFB)
    box.line.fill.background()
    txt(s, x + 0.15, 1.4, 2.7, 1.0, title, 16, True, BLUE)
    txt(s, x + 0.15, 2.5, 2.7, 3.6, body, 14, False, DARK)
footer(s, 12)

# 13 建议
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.25, 12, 0.45, "给园区的三条操作清单", 24, True, BLUE)
acts = [
    ("给机器人留路", "层高 ≥4.5m 中试层、货运电梯、充电消毒间、楼宇 5G/边缘节点，写入机器人友好楼宇评级。"),
    ("给模型留电", "变压器与楼层配电按训练/推理峰值预留 30% 冗余，提供可计量绿电套餐。"),
    ("给年轻人留夜", "24 小时创作者工位、无人零售与具身社区商业——年轻人聚集处才是 AI 产品第一验证场。"),
]
for i, (t, b) in enumerate(acts):
    y = 1.2 + i * 1.7
    box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(y), Inches(12.1), Inches(1.5))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xF4, 0xF7, 0xFB)
    box.line.fill.background()
    txt(s, 0.9, y + 0.2, 11.5, 0.4, f"{i+1}. {t}", 20, True, RED)
    txt(s, 0.9, y + 0.7, 11.5, 0.6, b, 16, False, DARK)
footer(s, 13)

# 14 结束
s = prs.slides.add_slide(BLANK)
add_bg(s, WHITE)
bar(s)
txt(s, 0.8, 2.2, 12, 1.2, "空间将成为 AI 时代最诚实的计分板", 28, True, BLUE)
txt(s, 0.8, 3.6, 11.5, 1.2,
    "哪座城市、哪个园区能让智能体安全地跑起来、让年轻人愿意留下来，\n产业就会在哪里生长。",
    18, False, DARK)
txt(s, 0.8, 5.8, 12, 0.5, "复旦大学住房政策研究中心  ·  FDU-HPRC-WP-2026-02", 16, False, GRAY)

prs.save(OUT)
print("saved", OUT)
