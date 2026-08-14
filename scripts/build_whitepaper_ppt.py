# -*- coding: utf-8 -*-
"""白皮书配套简报 PPT（中心研究文稿第三号）。"""
import os
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

OUT = "/workspace/whitepaper/2026谷歌上海开发者大会白皮书-简报.pptx"
LOGO = "/workspace/whitepaper/assets/logo_fudan_hprc.png"
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
    txt(slide, 0.5, 7.15, 10, 0.28,
        "复旦大学住房政策研究中心 · 研究文稿第三号  FDU-HPRC-WP-2026-03",
        size=10, color=GRAY)
    txt(slide, 11.6, 7.15, 1.3, 0.28, str(page), size=10, color=GRAY,
        align=PP_ALIGN.RIGHT)


def img(slide, path, l, t, w):
    slide.shapes.add_picture(path, Inches(l), Inches(t), width=Inches(w))


# 1 封面
s = prs.slides.add_slide(BLANK)
add_bg(s)
bar(s)
img(s, LOGO, 0.7, 0.4, 6.2)
txt(s, 0.7, 2.45, 12, 0.45, "中心研究文稿 · 第三号", 18, False, GRAY)
txt(s, 0.7, 3.05, 12, 1.0, "2026 谷歌上海开发者大会白皮书", 30, True, BLUE)
txt(s, 0.7, 4.2, 12, 0.55, "智能体时代的中国开发者生态、出海接口与城市空间含义", 16, False, GRAY)
txt(s, 0.7, 6.35, 12, 0.4, "FDU-HPRC-WP-2026-03    二〇二六年八月 · 上海", 14, False, GRAY)

# 2 核心判断
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.6, 0.28, 12, 0.5, "五个核心判断", 26, True, BLUE)
items = [
    "上海已是 Google 在中国的开发者主场：四年三届落沪，近 2,000 名开发者齐聚世博中心",
    "技术主轴从大模型展示切到智能体交付：Android / Chrome / Cloud 全栈跑通工作流",
    "中国开发者被定位为全球规模的供给方：社区、导师、加速器、中文文档四大支柱",
    "开源端侧把 AI 拉回住宅与社区：养老防诈、急救、教育成为 Gemma 4 决赛高频场景",
    "世博片区成为全球技术发布走廊：与 WAIC 同馆相继，接口型空间需求被低估",
]
for i, t in enumerate(items):
    y = 1.05 + i * 1.1
    box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(y), Inches(12.1), Inches(0.95))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xF4, 0xF7, 0xFB)
    box.line.fill.background()
    txt(s, 0.85, y + 0.22, 11.6, 0.55, f"{i + 1}.  {t}", 16, False, DARK)
footer(s, 2)

# 3 历年举办地
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.6, 0.25, 12, 0.45, "上海主场：I/O Connect China 四年三届", 24, True, BLUE)
img(s, os.path.join(CHARTS, "chart01_history.png"), 0.7, 0.9, 12.0)
footer(s, 3)

# 4 会展时间轴
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.6, 0.25, 12, 0.45, "2026 上海：从 WAIC 到 DevFest 的人才日历", 24, True, BLUE)
img(s, os.path.join(CHARTS, "chart02_calendar.png"), 0.7, 0.95, 12.0)
footer(s, 4)

# 5 I/O 指标
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.6, 0.22, 12, 0.4, "全球底盘：I/O 2026 披露的智能体规模", 24, True, BLUE)
img(s, os.path.join(CHARTS, "chart03_io_metrics.png"), 0.55, 0.7, 12.2)
footer(s, 5)

# 6 全栈
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.6, 0.25, 12, 0.45, "全栈 AI：从 TPU 到 Spark，中国站是翻译器", 24, True, BLUE)
img(s, os.path.join(CHARTS, "chart04_fullstack.png"), 0.7, 0.9, 12.0)
footer(s, 6)

# 7 四大支柱
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.6, 0.25, 12, 0.45, "出海接口：官方四大支柱", 24, True, BLUE)
img(s, os.path.join(CHARTS, "chart06_pillars.png"), 0.6, 1.0, 12.1)
footer(s, 7)

# 8 Gemma 场景
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.6, 0.22, 12, 0.4, "为真实而构建：Gemma 4 总决赛 15 支队伍的场景", 22, True, BLUE)
img(s, os.path.join(CHARTS, "chart08_gemma_scenes.png"), 0.4, 0.7, 7.4)
img(s, os.path.join(CHARTS, "chart07_gemma_tracks.png"), 7.7, 0.85, 5.3)
footer(s, 8)

# 9 WAIC 对照
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.6, 0.22, 12, 0.4, "同城双截面：产业博览 vs 开发者接口", 24, True, BLUE)
img(s, os.path.join(CHARTS, "chart10_waic_compare.png"), 0.55, 0.7, 12.2)
footer(s, 9)

# 10 空间框架
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.6, 0.22, 12, 0.4, "城市空间：六类可操作含义", 24, True, BLUE)
img(s, os.path.join(CHARTS, "chart12_space_framework.png"), 0.55, 0.7, 12.2)
footer(s, 10)

# 11 建议
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.6, 0.28, 12, 0.5, "对上海与浦东的六条建议", 26, True, BLUE)
recs = [
    "把世博片区明确为「全球技术发布走廊」，保护档期与双语服务",
    "试点 7—30 天技术活动友好型短租，对接大会、黑客松、加速器",
    "为 GDG 等独立社区提供常设 Workshop 空间，而不是一次性场租",
    "开放养老、社区卫生、学校作为开源智能体合规试点场景",
    "出海楼宇增加「接口密度」指标：短租工位、隔音舱、国际法务",
    "与北京、深圳错位：上海巩固主场会展与国际社区，不复制总部",
]
for i, t in enumerate(recs):
    y = 1.0 + i * 0.92
    nbox = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.65), Inches(y + 0.08), Inches(0.42), Inches(0.42))
    nbox.fill.solid()
    nbox.fill.fore_color.rgb = BLUE if i % 2 == 0 else RED
    nbox.line.fill.background()
    txt(s, 0.65, y + 0.12, 0.42, 0.38, str(i + 1), 14, True, WHITE, PP_ALIGN.CENTER)
    box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.25), Inches(y), Inches(11.4), Inches(0.78))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xF4, 0xF7, 0xFB)
    box.line.fill.background()
    txt(s, 1.45, y + 0.18, 11.0, 0.48, t, 16, False, DARK)
footer(s, 11)

# 12 封底
s = prs.slides.add_slide(BLANK)
add_bg(s)
bar(s)
txt(s, 0.7, 2.3, 12, 0.5, "复旦大学住房政策研究中心", 22, True, BLUE)
txt(s, 0.7, 3.0, 12, 0.4, "Housing Policy Research Center, Fudan University", 16, False, GRAY)
txt(s, 0.7, 3.8, 12, 0.4, "研究文稿第三号  ·  FDU-HPRC-WP-2026-03", 16, False, DARK)
txt(s, 0.7, 4.5, 12, 0.8,
    "本简报为学术讨论性质，不代表复旦大学或 Google 官方立场。\n完整论证、出处与方法见白皮书正文。",
    14, False, GRAY)
txt(s, 0.7, 6.4, 12, 0.4, "二〇二六年八月 · 上海", 14, False, GRAY)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
prs.save(OUT)
print("saved", OUT)
