# -*- coding: utf-8 -*-
"""白皮书配套简报 PPT（中心研究文稿第三号）。"""
import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

OUT = "/workspace/whitepaper/2026Google开发者大会观察白皮书-简报.pptx"
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
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width,
                                   prs.slide_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def bar(slide, y=0):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, y, prs.slide_width,
                               Inches(0.08))
    s.fill.solid()
    s.fill.fore_color.rgb = BLUE
    s.line.fill.background()


def txt(slide, l, t, w, h, text, size=18, bold=False, color=DARK,
        align=PP_ALIGN.LEFT, font="微软雅黑"):
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
img(s, LOGO, 0.7, 0.45, 6.2)
txt(s, 0.7, 2.45, 12, 0.4, "中心研究文稿 · 第三号", 18, False, GRAY)
txt(s, 0.7, 3.05, 12, 0.9, "2026 Google 开发者大会观察白皮书", 30, True, BLUE)
txt(s, 0.7, 4.15, 12, 0.5, "从灵感火花到产品跨海出圈：智能体时代的全球开发者网络与城市空间含义",
    16, False, GRAY)
txt(s, 0.7, 6.35, 12, 0.4, "FDU-HPRC-WP-2026-03    二〇二六年八月 · 上海", 14,
    False, GRAY)

# 2 核心判断
s = prs.slides.add_slide(BLANK)
add_bg(s)
bar(s)
footer(s, 2)
txt(s, 0.6, 0.3, 12, 0.5, "五项核心判断", 26, True, BLUE)
points = [
    "上海站是 I/O 2026 之后全球开发者巡回的中国节点，不是孤立发布会。",
    "技术重心从更强聊天模型转向可行动的全栈智能体。",
    "“跨海出圈”把大会写成全球化生产组织问题，上海是物理接口。",
    "Gemma 4 黑客松把适老化、医疗、教育、端侧隐私带进主会场。",
    "住房与空间的新变量是开发者角色重组：会展峰值、青年租赁、一人公司、端侧入户。",
]
for i, t in enumerate(points):
    txt(s, 0.7, 1.15 + i * 1.05, 12, 0.9, f"{i+1}.  {t}", 18, False, DARK)

# 3 大会全景
s = prs.slides.add_slide(BLANK)
add_bg(s)
bar(s)
footer(s, 3)
txt(s, 0.6, 0.28, 12, 0.45, "大会全景：可核验事实", 24, True, BLUE)
facts = [
    ("名称", "2026 Google 开发者大会 / Google I/O Connect China"),
    ("时间地点", "2026 年 8 月 12—13 日 · 上海世博中心"),
    ("口号", "从灵感火花到产品跨海出圈"),
    ("公开规模", "近 2,000 名开发者（官方博客）"),
    ("四大板块", "AI · Chrome · Cloud · Android"),
    ("全球网络", "I/O Connect 巡回：柏林 / 班加罗尔 / 上海"),
]
for i, (k, v) in enumerate(facts):
    y = 1.0 + i * 0.85
    shape = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7),
                               Inches(y), Inches(2.3), Inches(0.7))
    shape.fill.solid()
    shape.fill.fore_color.rgb = BLUE
    shape.line.fill.background()
    tf = shape.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = tf.paragraphs[0].add_run()
    run.text = k
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = WHITE
    run.font.name = "微软雅黑"
    txt(s, 3.2, y + 0.12, 9.5, 0.55, v, 18, False, DARK)

# 4 日历
s = prs.slides.add_slide(BLANK)
add_bg(s)
bar(s)
footer(s, 4)
txt(s, 0.5, 0.25, 12, 0.4, "2026 年关键节点", 22, True, BLUE)
img(s, os.path.join(CHARTS, "chart01_calendar.png"), 0.5, 0.85, 12.3)

# 5 全栈
s = prs.slides.add_slide(BLANK)
add_bg(s)
bar(s)
footer(s, 5)
txt(s, 0.5, 0.22, 12, 0.4, "技术风向：全栈智能体", 22, True, BLUE)
img(s, os.path.join(CHARTS, "chart03_fullstack.png"), 0.7, 0.75, 11.8)

# 6 三层部署 + Android
s = prs.slides.add_slide(BLANK)
add_bg(s)
bar(s)
footer(s, 6)
txt(s, 0.5, 0.2, 12, 0.35, "部署三层与 Android 工具更新", 22, True, BLUE)
img(s, os.path.join(CHARTS, "chart12_deploy_layers.png"), 0.3, 0.65, 6.4)
img(s, os.path.join(CHARTS, "chart04_android.png"), 6.7, 0.7, 6.2)

# 7 黑客松
s = prs.slides.add_slide(BLANK)
add_bg(s)
bar(s)
footer(s, 7)
txt(s, 0.5, 0.2, 12, 0.35, "Gemma 4 黑客松：社会场景进入主会场", 22, True, BLUE)
img(s, os.path.join(CHARTS, "chart05_hackathon_tracks.png"), 0.2, 0.7, 5.6)
img(s, os.path.join(CHARTS, "chart06_finalist_types.png"), 6.0, 1.15, 6.9)

# 8 对照 WAIC
s = prs.slides.add_slide(BLANK)
add_bg(s)
bar(s)
footer(s, 8)
txt(s, 0.5, 0.2, 12, 0.35, "同一世博片区的两种全球功能", 22, True, BLUE)
img(s, os.path.join(CHARTS, "chart09_expo_compare.png"), 0.55, 0.7, 12.2)

# 9 角色变迁 + 出海支柱
s = prs.slides.add_slide(BLANK)
add_bg(s)
bar(s)
footer(s, 9)
txt(s, 0.5, 0.2, 12, 0.35, "从写代码到管理智能体，再接到全球规模", 22, True, BLUE)
img(s, os.path.join(CHARTS, "chart08_role_shift.png"), 0.4, 0.65, 12.4)
img(s, os.path.join(CHARTS, "chart10_four_pillars.png"), 1.3, 3.55, 10.6)

# 10 空间冲击
s = prs.slides.add_slide(BLANK)
add_bg(s)
bar(s)
footer(s, 10)
txt(s, 0.5, 0.18, 12, 0.35, "对住房与城市空间的含义", 22, True, BLUE)
img(s, os.path.join(CHARTS, "chart11_space_radar.png"), 0.15, 0.55, 6.3)
img(s, os.path.join(CHARTS, "chart14_housing_priority.png"), 6.4, 1.05, 6.6)

# 11 建议
s = prs.slides.add_slide(BLANK)
add_bg(s)
bar(s)
footer(s, 11)
txt(s, 0.6, 0.28, 12, 0.45, "六条政策建议", 24, True, BLUE)
recs = [
    "把“全球开发者接口”写成世博片区正式功能，配置实验室型会展空间。",
    "补上中价位、可合规的 7—30 天服务式公寓，承接会展峰值。",
    "青年人才住房认定增加“出海技术团队”类别。",
    "杨浦环高校承接日常开发者网络，浦东世博承接年度全球接口。",
    "启动端侧智能体入户的住宅与社区标准预研。",
    "把可被智能体调用的城市服务接口纳入数字化前瞻议题。",
]
for i, t in enumerate(recs):
    txt(s, 0.7, 0.95 + i * 0.9, 12, 0.8, f"{i+1}.  {t}", 18, False, DARK)

# 12 结束页
s = prs.slides.add_slide(BLANK)
add_bg(s)
bar(s)
img(s, LOGO, 0.7, 0.5, 6.0)
txt(s, 0.7, 2.6, 12, 0.6, "会场灯光熄灭之后", 28, True, BLUE)
txt(s, 0.7, 3.4, 11.5, 1.6,
    "人才是否住得下来，团队是否找得到日常协作的房间，社区是否接得住已经在决赛台上跑通过的智能体——这才是住房政策需要提前回答的问题。",
    18, False, DARK)
txt(s, 0.7, 5.5, 12, 0.4, "复旦大学住房政策研究中心  ·  FDU-HPRC-WP-2026-03",
    16, True, BLUE)
txt(s, 0.7, 6.1, 12, 0.4, "资料截至 2026 年 8 月 14 日公开网页", 14, False, GRAY)

prs.save(OUT)
print("saved:", OUT)
