# -*- coding: utf-8 -*-
"""白皮书配套简报 PPT（中心研究文稿第三号）。"""
import os
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt, Emu

OUT = "/workspace/whitepaper/住房产业三链融合白皮书-简报.pptx"
LOGO = "/workspace/whitepaper/assets/logo_fudan_hprc.png"
CHARTS = "/workspace/whitepaper/assets/charts"
BLUE = RGBColor(0x0E, 0x4E, 0x9B)
RED = RGBColor(0xC8, 0x10, 0x2E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0x5B, 0x61, 0x6B)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
TEAL = RGBColor(0x3A, 0xA1, 0x7E)
LIGHT = RGBColor(0xF4, 0xF7, 0xFB)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


def add_bg(slide, color=WHITE):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def bar(slide, y=0, h=0.08, color=BLUE):
    s = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, Inches(y), prs.slide_width, Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = color
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
    bar(slide, 7.42, 0.08)
    txt(slide, 0.5, 7.12, 10.2, 0.26,
        "复旦大学住房政策研究中心 · 研究文稿第三号  FDU-HPRC-WP-2026-03",
        size=10, color=GRAY)
    txt(slide, 11.5, 7.12, 1.4, 0.26, str(page), size=10, color=GRAY,
        align=PP_ALIGN.RIGHT)


def img(slide, path, l, t, w):
    slide.shapes.add_picture(path, Inches(l), Inches(t), width=Inches(w))


def card(slide, l, t, w, h, fill=LIGHT):
    s = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    s.adjustments[0] = 0.08
    return s


# 1 封面
s = prs.slides.add_slide(BLANK)
add_bg(s)
bar(s)
img(s, LOGO, 0.65, 0.4, 6.4)
txt(s, 0.7, 2.35, 12, 0.4, "中心研究文稿 · 第三号", 16, False, GRAY)
txt(s, 0.7, 2.9, 12, 0.9, "住房产业“三链融合”白皮书", 32, True, BLUE)
txt(s, 0.7, 3.9, 12, 0.45, "供应链、硬件与软件协同的 2030 图景", 20, False, GRAY)
txt(s, 0.7, 4.5, 12, 0.4, "从“造房子”到“造产品、管资产、服务生活”", 16, False, DARK)
txt(s, 0.7, 6.35, 12, 0.35,
    "FDU-HPRC-WP-2026-03    二〇二六年八月 · 上海", 14, False, GRAY)

# 2 核心判断
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.55, 0.22, 12, 0.45, "五个核心判断", 24, True, BLUE)
items = [
    "政策已把房子重新定义为可制造的工业品与可治理的数据资产",
    "供应链从项目采购清单升级为“像造汽车一样造房子”的产品体系",
    "硬件把工地、房屋与家庭同时变成可感知、可执行的物理系统",
    "软件成为三链的操作系统：一模到底、一房一档、一生可管",
    "2030 年主线：保障房先行、更新场规模化、家庭智能体化、安全档案化",
]
for i, t in enumerate(items):
    y = 1.0 + i * 1.1
    card(s, 0.55, y, 12.2, 0.95)
    txt(s, 0.75, y + 0.28, 11.8, 0.5, f"{i + 1}.  {t}", 18, False, DARK)
footer(s, 2)

# 3 三链架构
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.18, 12, 0.4, "分析框架：供应链 · 硬件 · 软件必须长在一起", 22, True, BLUE)
img(s, f"{CHARTS}/chart01_three_chain.png", 0.35, 0.7, 12.6)
footer(s, 3)

# 4 政策演进
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.18, 12, 0.4, "六年制度升级：从推广技术到培育产业链", 22, True, BLUE)
img(s, f"{CHARTS}/chart02_policy_timeline.png", 0.4, 0.7, 12.5)
footer(s, 4)

# 5 试点 + 三项制度
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.4, 0.16, 12.5, 0.38, "试点已经够用，缺的是验收、结算、保险与档案互认", 20, True, BLUE)
img(s, f"{CHARTS}/chart03_pilot_scorecard.png", 0.15, 0.65, 6.5)
img(s, f"{CHARTS}/chart04_three_systems.png", 6.7, 0.7, 6.4)
footer(s, 5)

# 6 供应链现场
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.18, 12, 0.4, "供应链现场：田林路交付 + 华南 CMC 产线", 22, True, BLUE)
img(s, f"{CHARTS}/chart05_mic_efficiency.png", 0.25, 0.7, 6.6)
txt(s, 7.1, 0.95, 5.7, 5.4,
    "2026 年 6 月 17 日\n上海徐汇田林路 65 弄交付\n全国最大模块化城市更新\n1044 户 · 2725 个混凝土模块\n\n2026 年 8 月 6 日\n华南首条 CMC 智能产线下线\n用于深圳原拆原建试点\n\n土地规则一旦写入装配式/\n模块化比例，供应链就必须\n按产品而不是按项目投资产线。",
    15, False, DARK)
footer(s, 6)

# 7 硬件
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.18, 12, 0.4, "硬件三条验收尺子：工地 · 房屋 · 家庭", 22, True, BLUE)
img(s, f"{CHARTS}/chart06_smarthome_evolution.png", 0.3, 0.7, 12.7)
footer(s, 7)

# 8 软件
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.18, 12, 0.4, "软件：从图纸工具到住房操作系统", 22, True, BLUE)
img(s, f"{CHARTS}/chart07_software_stack.png", 0.35, 0.65, 12.6)
footer(s, 8)

# 9 生命周期
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.18, 12, 0.4, "融合发生在场景：全生命周期嵌入三链", 22, True, BLUE)
img(s, f"{CHARTS}/chart08_lifecycle.png", 0.25, 0.65, 12.8)
footer(s, 9)

# 10 上海
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.18, 12, 0.4, "上海样本：超大城市的政策实验室", 22, True, BLUE)
img(s, f"{CHARTS}/chart09_shanghai_role.png", 0.4, 0.7, 12.5)
footer(s, 10)

# 11 2030
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.18, 12, 0.4, "2030 图景：四条主线（情景判断，非官方预测）", 20, True, BLUE)
img(s, f"{CHARTS}/chart10_2030_outlook.png", 0.25, 0.62, 12.8)
footer(s, 11)

# 12 好房子
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.18, 12, 0.4, "成功标志是体验指标，而不是展示馆", 22, True, BLUE)
img(s, f"{CHARTS}/chart11_good_house.png", 0.4, 0.65, 12.5)
footer(s, 12)

# 13 建议
s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.18, 12, 0.4, "政策建议：先接口、再场景、后产业", 22, True, BLUE)
img(s, f"{CHARTS}/chart12_policy_matrix.png", 0.2, 0.58, 7.3)
txt(s, 7.6, 0.85, 5.3, 5.8,
    "政府最不该做的：\n指定某一品牌机器人或音箱\n\n政府最该做的：\n规定住房作为工业品和\n数据资产时必须具备的\n接口、档案和可核查性能\n\n12 条建议覆盖：\n身份 / 互联 / 安全接口\n保障房头雁与更新工具箱\n集成商与供应链金融\nCIM 开放与三链人才",
    15, False, DARK)
footer(s, 13)

# 14 结语
s = prs.slides.add_slide(BLANK)
add_bg(s)
bar(s)
txt(s, 0.7, 1.6, 12, 0.5, "结语", 18, False, GRAY)
txt(s, 0.7, 2.2, 12, 2.2,
    "2030 年并不遥远。真正决定那一年住房面貌的，\n不是又一次概念更新，而是我们此刻是否把\n标准、档案、产线和人的接口写清楚。",
    22, True, BLUE)
txt(s, 0.7, 5.0, 12, 0.8,
    "复旦大学住房政策研究中心\nHousing Policy Research Center, Fudan University",
    16, False, GRAY)
txt(s, 0.7, 6.3, 12, 0.35, "FDU-HPRC-WP-2026-03  ·  2026 年 8 月  ·  上海",
    14, False, GRAY)

prs.save(OUT)
print("saved:", OUT)
