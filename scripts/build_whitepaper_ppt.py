# -*- coding: utf-8 -*-
"""白皮书配套简报 PPT。"""
import os
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

OUT = "/workspace/whitepaper/医疗养老与提效服务白皮书-简报.pptx"
LOGO = "/workspace/whitepaper/assets/logo_fudan_hprc.png"
CHARTS = "/workspace/whitepaper/assets/charts"
BLUE = RGBColor(0x0E, 0x4E, 0x9B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0x5B, 0x61, 0x6B)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
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
    txt(slide, 0.5, 7.15, 10.2, 0.28,
        "复旦大学住房政策研究中心 · 研究文稿第四号  FDU-HPRC-WP-2026-04",
        size=10, color=GRAY)
    txt(slide, 11.6, 7.15, 1.3, 0.28, str(page), size=10, color=GRAY, align=PP_ALIGN.RIGHT)


def card(slide, l, t, w, h, fill=LIGHT):
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(l), Inches(t), Inches(w), Inches(h))
    box.fill.solid()
    box.fill.fore_color.rgb = fill
    box.line.fill.background()
    box.adjustments[0] = 0.08
    return box


s = prs.slides.add_slide(BLANK)
add_bg(s)
bar(s)
s.shapes.add_picture(LOGO, Inches(0.7), Inches(0.4), width=Inches(6.4))
txt(s, 0.7, 2.35, 12, 0.4, "中心研究文稿 · 第四号", 16, False, GRAY)
txt(s, 0.7, 2.85, 12, 0.7, "医疗、养老与提效服务", 34, True, BLUE)
txt(s, 0.7, 3.6, 12, 0.5, "照护空间、劳动释放与二〇三〇展望", 22, True, BLUE)
txt(s, 0.7, 4.3, 12, 0.4, "把家庭隐形税改写成可及的社会化服务", 16, False, GRAY)
txt(s, 0.7, 6.35, 12, 0.35, "FDU-HPRC-WP-2026-04    二〇二六年八月 · 上海", 14, False, GRAY)

s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.6, 0.3, 12, 0.45, "一句话判断", 24, True, BLUE)
card(s, 0.6, 1.2, 12.1, 2.3)
txt(s, 0.9, 1.55, 11.5, 1.7,
    "让需要照护的人不必离开熟悉的住宅，\n让提供劳动的人不必离开岗位。", 24, True, BLUE)
txt(s, 0.7, 3.8, 12, 2.8,
    "2025 年末：60 岁及以上 3.23 亿人，占 23.0%；人均预期寿命 79.25 岁。\n"
    "十五五目标：2030 年预期寿命 80 岁，劳动生产率增速高于 GDP。\n"
    "2026 年长护险启动全国制度建设，计划 2028 年底基本全面覆盖。", 16, False, DARK)
footer(s, 2)

s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.6, 0.25, 12, 0.4, "五个核心判断", 24, True, BLUE)
items = [
    "冲击不是床位不够，而是照护进不了家、服务嵌不进社区。",
    "医疗下一步不是堆大医院，而是把连续照护留在家门口。",
    "提效首先是把家庭照护时间还给劳动力市场。",
    "长护险是 2026—2028 年最重要的制度变量。",
    "住房政策是医疗、养老、提效三类服务的总接口。",
]
for i, t in enumerate(items):
    y = 0.9 + i * 1.12
    card(s, 0.6, y, 12.1, 1.0)
    txt(s, 0.85, y + 0.28, 11.6, 0.5, f"{i+1}.  {t}", 16, False, DARK)
footer(s, 3)

s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.2, 12, 0.4, "老龄底数：长寿社会已经到来", 22, True, BLUE)
s.shapes.add_picture(f"{CHARTS}/chart01_aging.png", Inches(0.35), Inches(0.7), width=Inches(12.5))
footer(s, 4)

s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.2, 12, 0.4, "三类服务是同一道题", 22, True, BLUE)
s.shapes.add_picture(f"{CHARTS}/chart05_three_services.png", Inches(0.4), Inches(1.1), width=Inches(12.5))
footer(s, 5)

s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.2, 12, 0.4, "医疗：机构增加、床位下降、基层诊疗过半", 22, True, BLUE)
s.shapes.add_picture(f"{CHARTS}/chart03_health.png", Inches(0.3), Inches(0.7), width=Inches(7.5))
txt(s, 8.0, 1.2, 4.8, 5.2,
    "2025 年诊疗 105.8 亿人次\n基层 55.6 亿、占 52.6%\n\n"
    "县域医共体 2199 个\n家庭医生转向“有感受”\n\n"
    "医养结合约 8300 家\n下一步考核转介，而非挂牌", 16, False, DARK)
footer(s, 6)

s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.2, 12, 0.4, "养老：床位在优化，网络才是主体", 22, True, BLUE)
s.shapes.add_picture(f"{CHARTS}/chart02_beds.png", Inches(0.3), Inches(0.7), width=Inches(7.6))
txt(s, 8.05, 1.15, 4.8, 5.4,
    "约 99% 老年人居家社区养老\n\n"
    "家庭养老床位累计 57 万张\n助餐点 8 万个、日服务超 300 万\n\n"
    "从业人员 139.5 万\n持证护理员仅 47.4%\n没有人，家庭床位只是一张床", 15, False, DARK)
footer(s, 7)

s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.2, 12, 0.4, "住房是总接口", 22, True, BLUE)
s.shapes.add_picture(f"{CHARTS}/chart06_space.png", Inches(0.4), Inches(1.15), width=Inches(12.5))
footer(s, 8)

s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.2, 12, 0.4, "时间轴：2028 年险成网，2030 年服务必须进家", 22, True, BLUE)
s.shapes.add_picture(f"{CHARTS}/chart07_timeline.png", Inches(0.35), Inches(1.0), width=Inches(12.6))
footer(s, 9)

s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.2, 12, 0.4, "十五五锚点与 2030 情景", 22, True, BLUE)
s.shapes.add_picture(f"{CHARTS}/chart04_15th_plan.png", Inches(0.2), Inches(0.7), width=Inches(6.5))
s.shapes.add_picture(f"{CHARTS}/chart08_scenarios.png", Inches(6.7), Inches(0.7), width=Inches(6.3))
footer(s, 10)

s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.2, 12, 0.4, "七条高确定性趋势", 22, True, BLUE)
trends = [
    "支付从补贴走向保险",
    "床位从数量走向功能",
    "医疗从住院走向连续",
    "服务从机构走向住宅",
    "一老一小从两个口走向一张网",
    "提效从口号走向劳动时间",
    "技术从展示走向看护",
]
for i, t in enumerate(trends):
    col, row = i % 2, i // 2
    x, y = 0.55 + col * 6.3, 0.9 + row * 1.35
    if i == 6:
        card(s, 0.55, y, 12.2, 1.15)
        txt(s, 0.8, y + 0.32, 11.7, 0.5, f"{i+1}.  {t}", 18, True, BLUE)
    else:
        card(s, x, y, 6.05, 1.2)
        txt(s, x + 0.25, y + 0.35, 5.6, 0.5, f"{i+1}.  {t}", 16, True, BLUE)
footer(s, 11)

s = prs.slides.add_slide(BLANK)
bar(s)
txt(s, 0.5, 0.2, 12, 0.4, "政策建议：空间、体系、劳动、制度", 22, True, BLUE)
s.shapes.add_picture(f"{CHARTS}/chart09_policy.png", Inches(0.4), Inches(1.05), width=Inches(12.5))
footer(s, 12)

s = prs.slides.add_slide(BLANK)
add_bg(s)
bar(s)
s.shapes.add_picture(LOGO, Inches(0.7), Inches(0.45), width=Inches(6.2))
txt(s, 0.7, 2.5, 12, 1.5,
    "空间配齐了，保险和技术才有用武之地。\n住房从负担变成人力资本的基础设施。", 24, True, BLUE)
txt(s, 0.7, 4.6, 12, 0.9,
    "全文见《医疗、养老与提效服务：照护空间、劳动释放与二〇三〇展望》\n"
    "复旦大学住房政策研究中心  ·  FDU-HPRC-WP-2026-04", 16, False, GRAY)
txt(s, 0.7, 6.3, 12, 0.35, "二〇二六年八月 · 上海", 14, False, GRAY)

os.makedirs("/workspace/下载版本", exist_ok=True)
prs.save(OUT)
prs.save("/workspace/下载版本/医疗养老与提效服务白皮书-简报.pptx")
print("saved", OUT)
