# -*- coding: utf-8 -*-
"""上海新谷湾三个月试单合作方案 PPT（16:9）。

呈送对方使用。内部提醒写在备注里，不放上版面。
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pathlib import Path

PLUM = RGBColor(0x2E, 0x1F, 0x47)
PURPLE = RGBColor(0x5B, 0x3E, 0x8E)
PURPLE2 = RGBColor(0x76, 0x58, 0xA8)
LILAC = RGBColor(0xC0, 0xAE, 0xE0)
LAV = RGBColor(0xEC, 0xE6, 0xF7)
LAVED = RGBColor(0xDD, 0xD2, 0xEF)
BGT = RGBColor(0xFB, 0xFA, 0xFE)
GOLD = RGBColor(0xC1, 0x9A, 0x3A)
GREY = RGBColor(0x60, 0x5A, 0x6B)
DARK = RGBColor(0x2B, 0x25, 0x36)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
INK = RGBColor(0x3A, 0x27, 0x5C)

FONT = "Microsoft YaHei"
FOOTER = "杨浦区科技企业联合会  ·  复旦大学住房政策研究中心  ·  上海新谷湾试单  ·  2026.10"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
prs.core_properties.title = "上海新谷湾三个月试单合作方案"
prs.core_properties.author = "上海市杨浦区科技企业联合会"
prs.core_properties.subject = "入会与 90 天试单，不签招商对赌"
prs.core_properties.category = "合作方案"
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def slide():
    return prs.slides.add_slide(BLANK)


def _set_font(run, size, bold, color, font=FONT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        node = rPr.find(qn(tag))
        if node is None:
            node = rPr.makeelement(qn(tag), {})
            rPr.append(node)
        node.set("typeface", font)


def rect(s, x, y, w, h, fill=None, line=None, line_w=None):
    sp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w or 1)
    return sp


def text(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         space_after=3, line_spacing=1.0):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Pt(2)
    tf.margin_right = Pt(2)
    tf.margin_top = Pt(1)
    tf.margin_bottom = Pt(1)
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        p.space_before = Pt(0)
        p.line_spacing = line_spacing
        for txt, size, bold, color in para:
            r = p.add_run()
            r.text = txt
            _set_font(r, size, bold, color)
    return tb


def bg(s):
    rect(s, 0, 0, SW, SH, fill=BGT)


def header(s, kicker, title):
    rect(s, 0, 0, SW, Inches(1.02), fill=PLUM)
    rect(s, 0, Inches(1.02), SW, Pt(3.5), fill=GOLD)
    text(s, Inches(0.48), Inches(0.12), Inches(12.2), Inches(0.28),
         [[(kicker, 11, True, LILAC)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(0.48), Inches(0.40), Inches(12.2), Inches(0.52),
         [[(title, 26, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)


def footer(s, n, total=11):
    text(s, Inches(0.48), Inches(7.16), Inches(10.6), Inches(0.26),
         [[(FOOTER, 10, False, GREY)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(11.3), Inches(7.16), Inches(1.55), Inches(0.26),
         [[(f"{n}  /  {total}", 10, True, PURPLE)]],
         align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def notes(s, body):
    ns = s.notes_slide
    ns.notes_text_frame.text = body


def card(s, x, y, w, h, fill=WHITE):
    return rect(s, x, y, w, h, fill=fill, line=LAVED, line_w=1)


# ------------------------------------------------------------ 1 封面
s = slide()
rect(s, 0, 0, SW, SH, fill=PLUM)
rect(s, 0, 0, Inches(8.55), SH, fill=PURPLE)
rect(s, Inches(8.55), 0, Pt(5), SH, fill=GOLD)
text(s, Inches(0.62), Inches(0.48), Inches(7.6), Inches(0.7),
     [[("试单合作方案", 14, True, LILAC)],
      [("不签招商对赌  ·  先入会，再看 90 天", 13, False, LILAC)]],
     space_after=2)
text(s, Inches(0.62), Inches(1.7), Inches(7.6), Inches(1.7),
     [[("上海新谷湾", 40, True, WHITE)],
      [("三个月试单合作方案", 32, True, WHITE)]],
     space_after=0)
rect(s, Inches(0.66), Inches(3.6), Inches(1.7), Pt(3.5), fill=GOLD)
text(s, Inches(0.62), Inches(3.85), Inches(7.6), Inches(1.15),
     [[("早期项目陪看", 18, True, WHITE)],
      [("复旦周边小型路演点", 18, True, WHITE)],
      [("产线外溢与新加坡分头对接", 18, True, WHITE)]],
     space_after=4)
text(s, Inches(0.62), Inches(6.35), Inches(7.6), Inches(0.7),
     [[("上海市杨浦区科技企业联合会", 13, True, WHITE)],
      [("复旦大学住房政策研究中心    2026 年 10 月", 12, False, LILAC)]],
     space_after=2)

facts = [
    ("01", "自持可租", "约 6,000㎡ 起谈"),
    ("02", "整栋 1.67 万㎡", "小业主部分不纳入"),
    ("03", "副会长", "年费 1 万元"),
    ("04", "试单", "清单到齐后 90 天"),
]
for i, (n, a, b) in enumerate(facts):
    y = Inches(0.42) + i * Inches(1.7)
    text(s, Inches(8.9), y, Inches(4.0), Inches(0.36),
         [[(n, 12, True, GOLD)]])
    text(s, Inches(8.9), y + Inches(0.34), Inches(4.0), Inches(0.4),
         [[(a, 18, True, WHITE)]])
    text(s, Inches(8.9), y + Inches(0.78), Inches(4.0), Inches(0.55),
         [[(b, 13, False, LILAC)]])
notes(s, "呈送对方。面积是 9 月 24 日洽谈口径，以空铺清单为准。不要在开场展开机器人、改公寓、创智汇。")

# ------------------------------------------------------------ 2 范围
s = slide()
bg(s)
header(s, "范围  ·  SCOPE", "这次只做三件事")
items = [
    ("01", "早期过渡办公", "把仍在找址、付得起租金的早期科技团队，带到有权出租的房源。面积按现有开间，好项目的租金可以谈。"),
    ("02", "小型路演点", "茶馆作为联合会在复旦周边的固定小场地。每月一场，主题只做人工智能落地或早期融资。"),
    ("03", "双向通道", "企业要产线时，办公留在上海，生产走已有外地通道。新加坡选址由出海同事另开一次会。"),
]
for i, (n, t, d) in enumerate(items):
    x = Inches(0.4) + i * Inches(4.28)
    card(s, x, Inches(1.28), Inches(4.08), Inches(3.35))
    rect(s, x, Inches(1.28), Inches(4.08), Inches(0.7), fill=PURPLE if i < 2 else PLUM)
    text(s, x + Inches(0.2), Inches(1.28), Inches(3.68), Inches(0.7),
         [[(f"{n}   {t}", 16, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + Inches(0.22), Inches(2.15), Inches(3.64), Inches(2.25),
         [[(d, 15, False, DARK)]], line_spacing=1.25)

text(s, Inches(0.42), Inches(4.8), Inches(12), Inches(0.32),
     [[("同时写进确认函的四条边界", 13, True, PLUM)]])
bounds = ["不设税收平台", "不对赌出租面积", "不改公寓、不切小间", "不交付企业名录"]
for i, b in enumerate(bounds):
    x = Inches(0.4) + i * Inches(3.2)
    rect(s, x, Inches(5.22), Inches(3.05), Inches(0.62), fill=LAV)
    rect(s, x, Inches(5.22), Pt(4), Inches(0.62), fill=GOLD)
    text(s, x + Inches(0.16), Inches(5.22), Inches(2.8), Inches(0.62),
         [[(b, 14, True, PLUM)]], anchor=MSO_ANCHOR.MIDDLE)
text(s, Inches(0.42), Inches(6.05), Inches(12.4), Inches(0.85),
     [[("创智汇、中建四局另行沟通。机器人集聚不写进本单。", 13, False, GREY)],
      [("录音口径：现场带看成交大约 110 比 1。试单考核到场和陪看，不考核成交套数。", 13, False, GREY)]],
     space_after=2)
footer(s, 2)
notes(s, "内部：不要用宝龙「一个月 120 个现场客」当这单的 KPI。对方区位在杨浦最西北，物业不在自己手里。机器人、改公寓若对方再提，回答「本单不纳入」。")

# ------------------------------------------------------------ 3 资源
s = slide()
bg(s)
header(s, "资源  ·  WHAT EACH SIDE BRINGS", "双方这次各自拿出什么")
cols = [
    (PURPLE, "上海新谷湾", [
        "自持办公。单一股东，好项目可以调租",
        "茶馆和小型活动场地，试单期免费",
        "注册地址：以清单确认还能不能落",
        "成都约 80 万㎡，以及台州、海宁、绍兴飞地",
        "正在选址的新加坡点，另会对接",
    ]),
    (PLUM, "联合会 / 住房政策研究中心", [
        "白皮书和走访形成的企业、融资判断",
        "世界人工智能大会与人工智能企业库",
        "仍在投早期项目的投行关系",
        "复旦科技园，以及现有会客厅机制",
        "小型活动组织；新加坡电商与小游戏已运营两年",
    ]),
]
for i, (color, title, lines) in enumerate(cols):
    x = Inches(0.4) + i * Inches(6.45)
    card(s, x, Inches(1.28), Inches(6.2), Inches(4.85))
    rect(s, x, Inches(1.28), Inches(6.2), Inches(0.68), fill=color)
    text(s, x + Inches(0.24), Inches(1.28), Inches(5.7), Inches(0.68),
         [[(title, 18, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)
    for j, line in enumerate(lines):
        y = Inches(2.16) + j * Inches(0.72)
        rect(s, x + Inches(0.22), y, Inches(5.76), Inches(0.62), fill=LAV)
        text(s, x + Inches(0.4), y, Inches(5.45), Inches(0.62),
             [[(line, 14, False, DARK)]], anchor=MSO_ANCHOR.MIDDLE)
text(s, Inches(0.42), Inches(6.28), Inches(12.4), Inches(0.7),
     [[("楼内企业按实际入驻介绍。澳大利亚市场今年才开，不写进这 90 天的交付。", 13, False, GREY)]],
     anchor=MSO_ANCHOR.MIDDLE)
footer(s, 3)
notes(s, "内部口径：东升巨变不在这栋楼，在的是华数科技。对外不要讲成「楼里有东升巨变」。成都 80 万㎡、飞地是对方口述加公开载体，产线项目单案核实。")

# ------------------------------------------------------------ 4 合作包
s = slide()
bg(s)
header(s, "包  ·  THE PACKAGE", "先入会，再试 90 天")
steps = [
    ("1", "入会", "副会长\n年费 1 万元", "换固定小场地、月度沙龙席位，以及挂牌机会"),
    ("2", "清单", "空铺表到齐", "没有这张表，不开始推企业"),
    ("3", "90 天", "只做过程", "筛 10–15 家、每月一场、投资人陪看"),
    ("4", "复盘", "三选一", "继续、停止，或改成单场活动"),
]
for i, (n, name, mid, desc) in enumerate(steps):
    x = Inches(0.38) + i * Inches(3.24)
    card(s, x, Inches(1.32), Inches(3.08), Inches(3.55))
    rect(s, x, Inches(1.32), Inches(3.08), Inches(0.08), fill=GOLD)
    text(s, x + Inches(0.18), Inches(1.5), Inches(2.7), Inches(0.4),
         [[(f"0{n}   {name}", 16, True, PURPLE)]])
    text(s, x + Inches(0.18), Inches(2.05), Inches(2.7), Inches(1.15),
         [[(mid, 20, True, PLUM)]], line_spacing=1.05)
    text(s, x + Inches(0.18), Inches(3.35), Inches(2.7), Inches(1.25),
         [[(desc, 13, False, DARK)]], line_spacing=1.15)
rect(s, Inches(0.4), Inches(5.1), Inches(12.52), Inches(1.75), fill=PLUM)
text(s, Inches(0.65), Inches(5.25), Inches(12.05), Inches(0.4),
     [[("挂牌的三个条件，缺一就不挂", 16, True, GOLD)]])
text(s, Inches(0.65), Inches(5.72), Inches(12.05), Inches(0.9),
     [[("有真实机构主体，优先复旦科技园或现有会客厅。", 15, False, WHITE)],
      [("挂上之后有季度活动。不新设空壳机构，也不先挂一块空牌子。", 15, False, WHITE)]],
     space_after=2)
footer(s, 4)
notes(s, "入会文本只写场地和月度沙龙。招商条款只写「推荐与陪看」。不要写面积目标，不要写税收返还。")

# ------------------------------------------------------------ 5 办公
s = slide()
bg(s)
header(s, "契合 1  ·  OFFICE", "早期团队的过渡办公")
card(s, Inches(0.4), Inches(1.26), Inches(7.55), Inches(5.55))
text(s, Inches(0.62), Inches(1.4), Inches(7.1), Inches(0.4),
     [[("推过来的企业，六条都要对上", 16, True, PLUM)]])
rules = [
    ("产业", "人工智能，或相邻硬科技"),
    ("阶段", "早期。尚未被免租加投资锁定"),
    ("支付", "付得起大约 1.5–2 元/㎡/天"),
    ("面积", "对得上清单里的现有房源"),
    ("产品", "接受现在的大开间"),
    ("资本", "最好带一位能影响选址的投资人"),
]
for i, (k, v) in enumerate(rules):
    y = Inches(1.95) + i * Inches(0.75)
    rect(s, Inches(0.62), y, Inches(1.35), Inches(0.58), fill=PURPLE)
    text(s, Inches(0.62), y, Inches(1.35), Inches(0.58),
         [[(k, 14, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(2.15), y, Inches(5.5), Inches(0.58),
         [[(v, 15, False, DARK)]], anchor=MSO_ANCHOR.MIDDLE)

card(s, Inches(8.15), Inches(1.26), Inches(4.75), Inches(5.55), fill=PLUM)
text(s, Inches(8.4), Inches(1.45), Inches(4.3), Inches(0.4),
     [[("操作顺序", 16, True, GOLD)]])
order = [
    "先看空铺清单，再从库里筛人",
    "内部留下 10–15 家，整表不外发",
    "只约有权出租的那一部分",
    "看盘必须到现场，电话不算",
    "租金以清单报价为底，好项目当场谈",
]
for i, line in enumerate(order):
    y = Inches(2.1) + i * Inches(0.82)
    text(s, Inches(8.4), y, Inches(0.45), Inches(0.4),
         [[(str(i + 1), 16, True, GOLD)]])
    text(s, Inches(8.9), y, Inches(3.7), Inches(0.7),
         [[(line, 14, False, WHITE)]], anchor=MSO_ANCHOR.TOP)
footer(s, 5)
notes(s, "1.5–2 元是筛选带，不是替对方报价。自持洽谈口径在 2 元以上，好项目可降。1–4 人隔间已经改掉，预算不够改回，不能答应切小间。晶丰是对方自己的教训：长起来会被免租加投资挖走，所以每场要有早期投资人。")

# ------------------------------------------------------------ 6 路演
s = slide()
bg(s)
header(s, "契合 2  ·  SALON", "复旦周边的小型路演点")
blocks = [
    ("规模", "按茶馆实测。\n方案先按 30–80 人准备。"),
    ("频率", "试单期每月 1 场。\n90 天一共 3 场。"),
    ("主题", "只做两选一：\n人工智能落地，或早期融资。"),
    ("到场", "每场至少一位\n仍在投早期项目的机构。"),
]
for i, (t, d) in enumerate(blocks):
    x = Inches(0.4) + i * Inches(3.22)
    card(s, x, Inches(1.28), Inches(3.05), Inches(2.55))
    rect(s, x, Inches(1.28), Inches(3.05), Pt(5), fill=GOLD)
    text(s, x + Inches(0.18), Inches(1.48), Inches(2.7), Inches(0.4),
         [[(t, 16, True, PURPLE)]])
    text(s, x + Inches(0.18), Inches(2.0), Inches(2.7), Inches(1.55),
         [[(d, 15, False, DARK)]], line_spacing=1.15)

card(s, Inches(0.4), Inches(4.05), Inches(12.52), Inches(2.75))
text(s, Inches(0.64), Inches(4.2), Inches(12), Inches(0.4),
     [[("场地免费，买的是到场的人", 16, True, PLUM)]])
rows = [
    ("他们出", "场地、现场接待、开门时间"),
    ("我们出", "主题、嘉宾、企业邀约、流程"),
    ("试单期", "组织费用由联合会承担"),
    ("另计", "对外传播、直播、媒体稿"),
]
for i, (k, v) in enumerate(rows):
    col, row = i % 2, i // 2
    x = Inches(0.64) + col * Inches(6.15)
    y = Inches(4.75) + row * Inches(0.85)
    rect(s, x, y, Inches(1.4), Inches(0.62), fill=LAV)
    text(s, x, y, Inches(1.4), Inches(0.62),
         [[(k, 14, True, PURPLE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + Inches(1.55), y, Inches(4.3), Inches(0.62),
         [[(v, 15, False, DARK)]], anchor=MSO_ANCHOR.MIDDLE)
footer(s, 6)
notes(s, "600 人那场峰会不要往这栋楼排。对方自己说承载不了。30–80 人是方案假设，第一场前去茶馆量一次。")

# ------------------------------------------------------------ 7 通道
s = slide()
bg(s)
header(s, "契合 3  ·  TWO-WAY", "双向通道分开谈")
channels = [
    ("A", "上海企业要产线", "办公和研发留在上海。生产放到对方已跑过的通道：台州城投、海宁、绍兴，或成都载体。企业自己提出来，再单案对接。"),
    ("B", "成都企业来上海", "只接待本来就要来长三角的团队。不承诺从成都批量导入新谷湾。这栋楼做桥头堡，不做搬迁指标。"),
    ("C", "新加坡选址", "对方正在选点。我方电商和小游戏已在新加坡运营两年。出海同事单独开一次会，验证能不能互相落地。"),
]
for i, (n, t, d) in enumerate(channels):
    y = Inches(1.28) + i * Inches(1.55)
    card(s, Inches(0.4), y, Inches(12.52), Inches(1.42))
    rect(s, Inches(0.4), y, Inches(1.15), Inches(1.42), fill=PURPLE if i < 2 else PLUM)
    text(s, Inches(0.4), y, Inches(1.15), Inches(1.42),
         [[(n, 26, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(1.75), y + Inches(0.16), Inches(10.8), Inches(0.4),
         [[(t, 16, True, PLUM)]])
    text(s, Inches(1.75), y + Inches(0.58), Inches(10.8), Inches(0.7),
         [[(d, 14, False, DARK)]])
text(s, Inches(0.45), Inches(6.05), Inches(12.4), Inches(0.85),
     [[("A、B 不写进 90 天的过程指标。C 不写进招商安排。三件事都单案记账。", 14, True, PURPLE)]],
     anchor=MSO_ANCHOR.MIDDLE)
footer(s, 7)
notes(s, "飞地数字来自洽谈：台州城投约 4000㎡合作八年，海宁约 2000㎡三年多，绍兴也合作过。对外只说「已有通道」，具体面积等项目核实。新加坡对接人姓名没在会上留下，群里要。")

# ------------------------------------------------------------ 8 90天
s = slide()
bg(s)
header(s, "过程  ·  90 DAYS", "90 天只承诺过程")
rect(s, Inches(0.4), Inches(1.26), Inches(12.52), Inches(0.58), fill=PURPLE)
heads = [(Inches(0.5), Inches(1.7), "时间"), (Inches(2.2), Inches(3.3), "动作"),
         (Inches(5.6), Inches(5.0), "完成的样子"), (Inches(10.7), Inches(2.0), "谁负责")]
for x, w, h in heads:
    text(s, x, Inches(1.26), w, Inches(0.58),
         [[(h, 14, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)
rows = [
    ("第 1 周", "空铺清单、新加坡对接人、入会主体", "三样到齐，试单才起算", "对方填，朱震收"),
    ("第 1 周", "副会长文本", "写明场地和月度沙龙", "联合会"),
    ("第 2–4 周", "按清单筛 10–15 家", "内部表填完，整表不外发", "联合会"),
    ("每月", "一场沙龙，一位早期投资人看盘", "有到场记录，电话不算", "双方"),
    ("第 90 天", "复盘，书面三选一", "继续 / 停止 / 改成单场活动", "双方"),
]
for i, row in enumerate(rows):
    y = Inches(1.96) + i * Inches(0.9)
    fill = WHITE if i % 2 == 0 else LAV
    rect(s, Inches(0.4), y, Inches(12.52), Inches(0.82), fill=fill, line=LAVED, line_w=0.75)
    xs = [Inches(0.5), Inches(2.2), Inches(5.6), Inches(10.7)]
    ws = [Inches(1.6), Inches(3.3), Inches(4.9), Inches(2.05)]
    bolds = [True, False, False, False]
    for x, w, val, b in zip(xs, ws, row, bolds):
        text(s, x, y, w, Inches(0.82),
             [[(val, 13, b, PLUM if b else DARK)]], anchor=MSO_ANCHOR.MIDDLE)
footer(s, 8)
notes(s, "不要在这页补一个带看量指标。对方知道行业大约 110 比 1，但本单不接这个考核。")

# ------------------------------------------------------------ 9 清单
s = slide()
bg(s)
header(s, "启动门  ·  THE LIST", "没有空铺清单，不开始推企业")
text(s, Inches(0.45), Inches(1.2), Inches(12.4), Inches(0.4),
     [[("请对方按这个字段填。小业主的面积可以附在表外，不能混进可推清单。", 14, False, GREY)]])
fields = [
    ("楼层房号", "具体到间"),
    ("建筑面积", "㎡，可租面积"),
    ("权属", "自持 / 小业主"),
    ("状态", "空置，或到期日"),
    ("报价", "元/㎡/天，不含物业"),
    ("物业费", "金额，以及能否谈"),
    ("交付", "毛坯 / 精装 / 现状"),
    ("注册", "还能不能落地址"),
    ("租期", "最短租期、免租"),
    ("条件", "空调、朝向、层高"),
    ("决策", "谁能当场定租"),
    ("对接人", "姓名和电话"),
]
for i, (k, v) in enumerate(fields):
    col, row = i % 4, i // 4
    x = Inches(0.4) + col * Inches(3.22)
    y = Inches(1.75) + row * Inches(1.45)
    card(s, x, y, Inches(3.05), Inches(1.28))
    rect(s, x, y, Inches(0.1), Inches(1.28), fill=GOLD)
    text(s, x + Inches(0.28), y + Inches(0.16), Inches(2.6), Inches(0.4),
         [[(k, 16, True, PLUM)]])
    text(s, x + Inches(0.28), y + Inches(0.64), Inches(2.6), Inches(0.45),
         [[(v, 13, False, GREY)]])
footer(s, 9)
notes(s, "物业费洽谈口径约 20 元/㎡/月，停车约 700 元/月，都在开发商物业手里。表上要单列「能否谈」，不能默认可谈。")

# ------------------------------------------------------------ 10 费用
s = slide()
bg(s)
header(s, "费用  ·  WHO PAYS", "谁交付，谁付费")
rect(s, Inches(0.35), Inches(1.24), Inches(12.62), Inches(0.52), fill=PURPLE)
for x, w, h in (
    (Inches(0.45), Inches(2.1), "事项"),
    (Inches(2.6), Inches(3.5), "新谷湾"),
    (Inches(6.15), Inches(3.7), "联合会"),
    (Inches(9.95), Inches(2.85), "费用"),
):
    text(s, x, Inches(1.24), w, Inches(0.52),
         [[(h, 13, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE)
fees = [
    ("入会", "办副会长", "纳入活动点", "1 万元/年"),
    ("场地", "每月开放一场", "定主题和邀约", "场地不收费"),
    ("沙龙", "现场接待", "组织 3 场", "试单期我们承担"),
    ("陪看", "出报价，可当场定租", "筛选、预约、陪同", "不写佣金和税收"),
    ("挂牌", "出墙面和发布", "协调真实主体", "主体未定不挂"),
    ("新加坡", "出对接人", "出海同事参会", "不进本方案"),
    ("产线", "提供已有通道", "企业提出才对接", "单案单议"),
]
for i, row in enumerate(fees):
    y = Inches(1.82) + i * Inches(0.68)
    rect(s, Inches(0.35), y, Inches(12.62), Inches(0.64),
         fill=WHITE if i % 2 == 0 else LAV)
    xs = [Inches(0.45), Inches(2.6), Inches(6.15), Inches(9.95)]
    ws = [Inches(2.05), Inches(3.45), Inches(3.7), Inches(2.85)]
    for j, (x, w, val) in enumerate(zip(xs, ws, row)):
        text(s, x, y, w, Inches(0.64),
             [[(val, 13, j == 0 or j == 3, PLUM if j in (0, 3) else DARK)]],
             anchor=MSO_ANCHOR.MIDDLE)
footer(s, 10)
notes(s, "这页是给对方看的收费边界。成交佣金即使以后要谈，也不在试单文本里。避免会后变成招商对赌。")

# ------------------------------------------------------------ 11 下一步
s = slide()
bg(s)
header(s, "下一步  ·  NEXT", "群里先要两样东西")
actions = [
    ("1", "先要表，不谈框架", "空铺清单和新加坡对接人。两样到了，再约第一场看盘。"),
    ("2", "入会按副会长写", "文本写明场地和月度沙龙。招商只写推荐与陪看。"),
    ("3", "另一条线分开", "创智汇、中建四局另行沟通，不和这栋楼写在同一份确认里。"),
]
for i, (n, t, d) in enumerate(actions):
    y = Inches(1.24) + i * Inches(1.15)
    card(s, Inches(0.4), y, Inches(12.52), Inches(1.05))
    rect(s, Inches(0.4), y, Inches(0.9), Inches(1.05), fill=PURPLE)
    text(s, Inches(0.4), y, Inches(0.9), Inches(1.05),
         [[(n, 22, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(1.5), y + Inches(0.1), Inches(11.1), Inches(0.38),
         [[(t, 16, True, PLUM)]])
    text(s, Inches(1.5), y + Inches(0.5), Inches(11.1), Inches(0.42),
         [[(d, 14, False, DARK)]])

text(s, Inches(0.45), Inches(4.75), Inches(6), Inches(0.35),
     [[("当场填", 14, True, PURPLE)]])
blanks = ["新谷湾对接人", "新加坡对接人", "清单提交日", "首场沙龙窗口"]
for i, name in enumerate(blanks):
    col, row = i % 2, i // 2
    x = Inches(0.4) + col * Inches(6.45)
    y = Inches(5.18) + row * Inches(0.85)
    text(s, x, y, Inches(2.3), Inches(0.55),
         [[(name, 14, True, PLUM)]], anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x + Inches(2.3), y + Inches(0.38), Inches(3.6), Pt(1.25), fill=PURPLE2)
footer(s, 11)
notes(s, "会后三件事的原口径：朱震拉群；入会不写面积和税收；创智汇只认交情，不绑定本楼。主谈人曾在创智汇、从富湘出来，这个背景留在内部，不上版面。")


out = Path("/workspace/deliverables")
out.mkdir(parents=True, exist_ok=True)
target = out / "上海新谷湾三个月试单合作方案.pptx"
prs.save(target)
print(target)
