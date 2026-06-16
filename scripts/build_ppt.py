"""
生成《上海商办楼宇与产业园区市场深度报告》汇报版 PPT
易居研究院 × 复旦大学住房政策研究中心
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from copy import deepcopy
import os

OUTPUT_DIR = "/workspace/上海商办楼宇与产业园区市场深度报告"
os.makedirs(OUTPUT_DIR, exist_ok=True)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

SLIDE_W = prs.slide_width
SLIDE_H = prs.slide_height

# 颜色体系
NAVY = RGBColor(0x1F, 0x3A, 0x68)         # 深蓝
BLUE = RGBColor(0x2E, 0x5C, 0xA3)         # 主蓝
LIGHT_BLUE = RGBColor(0xE8, 0xF0, 0xFA)   # 浅蓝
ACCENT = RGBColor(0xC8, 0x86, 0x2D)       # 金棕
DARK = RGBColor(0x22, 0x2A, 0x35)         # 文本深色
GRAY = RGBColor(0x6E, 0x73, 0x7C)         # 文本灰
LIGHT_GRAY = RGBColor(0xF3, 0xF5, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ORANGE = RGBColor(0xE0, 0x6A, 0x2C)
GREEN = RGBColor(0x4A, 0x8C, 0x5C)


def set_text(run, text, font="微软雅黑", size=18, bold=False, color=DARK):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rpr = run._r.get_or_add_rPr()
    east = rpr.find(qn("a:ea"))
    if east is None:
        east = qn("a:ea")
        from lxml import etree
        ea = etree.SubElement(rpr, qn("a:ea"))
        ea.set("typeface", font)
    else:
        east.set("typeface", font)


def add_rect(slide, x, y, w, h, fill_color, line=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_color
    if not line:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = fill_color
    shp.shadow.inherit = False
    return shp


def add_text(slide, x, y, w, h, text, size=18, bold=False, color=DARK, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, font="微软雅黑"):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(2)
    tf.margin_right = Pt(2)
    tf.margin_top = Pt(2)
    tf.margin_bottom = Pt(2)
    tf.vertical_anchor = anchor
    if isinstance(text, list):
        for i, line in enumerate(text):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            r = p.add_run()
            set_text(r, line, font=font, size=size, bold=bold, color=color)
            p.space_after = Pt(2)
    else:
        p = tf.paragraphs[0]
        p.alignment = align
        r = p.add_run()
        set_text(r, text, font=font, size=size, bold=bold, color=color)
    return tb


def add_bullets(slide, x, y, w, h, items, size=15, color=DARK, bullet_color=BLUE,
                line_spacing=1.25, bullet_char="●"):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(2)
    tf.margin_right = Pt(2)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(4)
        p.line_spacing = line_spacing
        r1 = p.add_run()
        set_text(r1, f"{bullet_char}  ", size=size, color=bullet_color, bold=True)
        r2 = p.add_run()
        set_text(r2, item, size=size, color=color)
    return tb


def page_header(slide, chapter, title, subtitle=None):
    # 顶部装饰条
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.55), NAVY)
    add_rect(slide, 0, Inches(0.55), SLIDE_W, Inches(0.05), ACCENT)
    # 章节标签
    add_text(slide, Inches(0.4), Inches(0.05), Inches(5), Inches(0.45),
             chapter, size=14, bold=True, color=WHITE, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.MIDDLE)
    # 报告主体
    add_text(slide, Inches(8.0), Inches(0.05), Inches(5.0), Inches(0.45),
             "易居研究院 × 复旦大学住房政策研究中心", size=11,
             color=WHITE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    # 主标题
    add_text(slide, Inches(0.5), Inches(0.75), Inches(12.3), Inches(0.6),
             title, size=26, bold=True, color=NAVY,
             anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_text(slide, Inches(0.5), Inches(1.32), Inches(12.3), Inches(0.35),
                 subtitle, size=13, color=GRAY)
    # 底部
    add_rect(slide, 0, SLIDE_H - Inches(0.35), SLIDE_W, Inches(0.35), LIGHT_GRAY)
    add_text(slide, Inches(0.4), SLIDE_H - Inches(0.35), Inches(8), Inches(0.35),
             "上海商办楼宇与产业园区市场深度报告  ·  讨论稿",
             size=10, color=GRAY, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(11.8), SLIDE_H - Inches(0.35), Inches(1.2), Inches(0.35),
             "", size=10, color=GRAY, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)


def add_card(slide, x, y, w, h, title, items, accent=BLUE, title_size=14, body_size=12):
    add_rect(slide, x, y, w, h, WHITE)
    border = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    border.fill.background()
    border.line.color.rgb = accent
    border.line.width = Pt(0.75)
    border.shadow.inherit = False
    # 顶部色条
    add_rect(slide, x, y, w, Inches(0.32), accent)
    add_text(slide, x + Inches(0.12), y, w - Inches(0.2), Inches(0.32),
             title, size=title_size, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    add_bullets(slide, x + Inches(0.15), y + Inches(0.4), w - Inches(0.25), h - Inches(0.45),
                items, size=body_size, color=DARK, bullet_color=accent,
                bullet_char="•", line_spacing=1.2)


def slide_page_number(slide, num):
    add_text(slide, Inches(12.6), SLIDE_H - Inches(0.35), Inches(0.6), Inches(0.35),
             f"{num:02d}", size=10, color=GRAY, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)


def new_slide():
    return prs.slides.add_slide(prs.slide_layouts[6])  # blank


page_counter = [0]


def new_content_slide(chapter, title, subtitle=None):
    s = new_slide()
    page_header(s, chapter, title, subtitle)
    page_counter[0] += 1
    slide_page_number(s, page_counter[0])
    return s


# ============== 封面 ==============
s = new_slide()
add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY)
add_rect(s, 0, Inches(0), SLIDE_W, Inches(0.12), ACCENT)
add_rect(s, 0, SLIDE_H - Inches(0.12), SLIDE_W, Inches(0.12), ACCENT)

# 装饰大色块
deco = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0.5),
                          Inches(0.18), Inches(6.5))
deco.fill.solid()
deco.fill.fore_color.rgb = ACCENT
deco.line.fill.background()

# 副标题（上）
add_text(s, Inches(0.8), Inches(1.0), Inches(12), Inches(0.5),
         "易居研究院 × 复旦大学住房政策研究中心  联合课题组", size=18,
         color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

# 主标题
add_text(s, Inches(0.8), Inches(2.2), Inches(12), Inches(1.2),
         "上海商办楼宇与产业园区", size=52, bold=True, color=WHITE,
         anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Inches(0.8), Inches(3.4), Inches(12), Inches(1.2),
         "市场深度报告", size=58, bold=True, color=WHITE,
         anchor=MSO_ANCHOR.MIDDLE)

# 副标题
add_text(s, Inches(0.8), Inches(4.9), Inches(12), Inches(0.5),
         "全域空间供给 · 企业需求迁徙 · 产业载体运营研究", size=20,
         color=RGBColor(0xCF, 0xDD, 0xF0), anchor=MSO_ANCHOR.MIDDLE)

add_text(s, Inches(0.8), Inches(5.6), Inches(12), Inches(0.4),
         "研发方案 · 报告大纲 · 数据调用建议 · 执行指引（讨论稿）", size=14,
         color=RGBColor(0xAB, 0xC0, 0xD9), anchor=MSO_ANCHOR.MIDDLE)

# 底部联合署名
add_text(s, Inches(0.8), Inches(6.6), Inches(12), Inches(0.4),
         "联合发布 · 上海商办楼宇与产业园区市场研究课题组", size=12,
         color=RGBColor(0xC8, 0x86, 0x2D), anchor=MSO_ANCHOR.MIDDLE)


# ============== 目录 ==============
s = new_content_slide("CONTENTS", "目录 · Outline")
toc = [
    ("01", "项目定位与背景", "联合发布主体、核心定位、读者矩阵"),
    ("02", "市场发展背景", "城市更新、五个中心、需求重构"),
    ("03", "全域供给格局", "16区 + 街镇 + 楼宇 + 园区"),
    ("04", "租金空置与价格体系", "分层租金 + 空置结构 + 综合成本"),
    ("05", "入驻企业画像", "行业、规模、成长阶段、清单体系"),
    ("06", "企业迁徙与产业流动", "跨区迁徙、产业流向、吸引力指数"),
    ("07", "供需匹配与招商机会", "竞争力模型 + 三类招商机会"),
    ("08", "市场趋势预测", "租金/空置/产业/空间策略"),
    ("09", "政策建议与市场应用", "政府、园区、资产、企业"),
    ("10", "成果体系与商业化", "月报/季报/白皮书/诊断/选址"),
    ("11", "数据体系与字段设计", "六大基础数据库"),
    ("12", "核心指标体系", "五大指数"),
    ("13", "数据调用与采集指引", "五项原则、四阶段、合规要点"),
    ("14", "分工协同与首期落地", "易居 × 复旦 × 数据技术"),
]
# 左右两列布局
for i, (num, name, desc) in enumerate(toc):
    col = i // 7
    row = i % 7
    x = Inches(0.5 + col * 6.4)
    y = Inches(1.85 + row * 0.7)
    add_text(s, x, y, Inches(0.6), Inches(0.6), num, size=22, bold=True,
             color=ACCENT, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.7), y, Inches(5.5), Inches(0.4), name, size=15,
             bold=True, color=NAVY, anchor=MSO_ANCHOR.TOP)
    add_text(s, x + Inches(0.7), y + Inches(0.32), Inches(5.5), Inches(0.3),
             desc, size=11, color=GRAY)


# ============== Part 01：项目定位 ==============
s = new_content_slide("PART 01 · 项目定位", "联合发布与核心定位",
                      "上海首份面向全域商办楼宇与产业园区的供需双侧 / 企业迁徙 / 资产运营决策报告")

# 左侧：联合发布
add_rect(s, Inches(0.5), Inches(1.85), Inches(5.8), Inches(2.5), LIGHT_BLUE)
add_text(s, Inches(0.7), Inches(1.95), Inches(5.5), Inches(0.45),
         "联合发布主体", size=16, bold=True, color=NAVY)
add_text(s, Inches(0.7), Inches(2.45), Inches(5.5), Inches(0.55),
         "易居房地产研究院", size=22, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Inches(0.7), Inches(3.05), Inches(5.5), Inches(0.4),
         "×", size=18, color=ACCENT, bold=True, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Inches(0.7), Inches(3.45), Inches(5.5), Inches(0.55),
         "复旦大学住房政策研究中心", size=22, bold=True, color=NAVY,
         anchor=MSO_ANCHOR.MIDDLE)

# 右侧：核心定位
add_text(s, Inches(6.7), Inches(1.85), Inches(6.3), Inches(0.4),
         "核心定位", size=16, bold=True, color=NAVY)
add_bullets(s, Inches(6.7), Inches(2.3), Inches(6.3), Inches(2.2), [
    "全域覆盖：16区 → 街镇 → 楼宇/园区 → 企业 四级颗粒",
    "全类型载体：写字楼、产业园、科创园、孵化器、商务园",
    "供需双侧：供给底数 × 企业入驻/迁徙/扩缩租行为",
    "产业视角：集聚、流动、补链强链与新质生产力",
    "招商资管导向：政府招商 + 园区运营 + 资产策略 + 企业选址",
], size=13)

# 下方：突破性
add_text(s, Inches(0.5), Inches(4.6), Inches(12.3), Inches(0.4),
         "五大研究突破", size=16, bold=True, color=NAVY)
breakthroughs = [
    ("全域", "从核心区样本走向上海全域覆盖"),
    ("全载体", "从写字楼走向商办与产业园区全类型"),
    ("双侧", "从租金空置走向供需双侧研究"),
    ("动态", "从静态描述走向企业迁徙动态监测"),
    ("应用", "从宏观研究走向招商/资管/选址应用"),
]
for i, (k, v) in enumerate(breakthroughs):
    x = Inches(0.5 + i * 2.52)
    add_rect(s, x, Inches(5.1), Inches(2.4), Inches(1.7), WHITE)
    border = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(5.1),
                                Inches(2.4), Inches(1.7))
    border.fill.background()
    border.line.color.rgb = BLUE
    border.shadow.inherit = False
    add_rect(s, x, Inches(5.1), Inches(2.4), Inches(0.42), BLUE)
    add_text(s, x, Inches(5.1), Inches(2.4), Inches(0.42), k, size=14,
             bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.15), Inches(5.6), Inches(2.1), Inches(1.15),
             v, size=11, color=DARK, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)


# ============== Part 01 - 读者矩阵 ==============
s = new_content_slide("PART 01 · 项目定位", "五类目标读者矩阵",
                      "政府、资产、投资、企业、研究 — 一份报告对应五个使用场景")
readers = [
    ("政府与产业招商", BLUE, ["市/区招商主管部门", "街镇招商中心", "开发区/功能区管委会"],
     "空间治理 · 精准招商 · 楼宇与园区经济政策"),
    ("楼宇/园区运营方", ORANGE, ["商办楼宇业主", "产业园区运营商", "城市更新操盘方", "国企平台公司"],
     "资产定位 · 租金策略 · 招商策略 · 运营服务"),
    ("资产持有与投资", GREEN, ["REITs / 类REITs资产方", "商办投资人", "资管公司", "地产基金"],
     "竞品对标 · 估值 · 改造 · 资本化路径"),
    ("企业选址决策方", ACCENT, ["科创企业", "专精特新企业", "总部企业", "生产性服务业"],
     "区域比较 · 综合成本 · 政策匹配 · 扩缩租"),
    ("研究与咨询机构", NAVY, ["房地产研究机构", "产业咨询机构", "金融研究部门", "城市更新服务商"],
     "标准化口径 · 行业研究 · 政策咨询"),
]
for i, (name, color, items, use) in enumerate(readers):
    col = i % 3
    row = i // 3
    x = Inches(0.5 + col * 4.25)
    y = Inches(1.95 + row * 2.6)
    add_rect(s, x, y, Inches(4.0), Inches(2.4), WHITE)
    border = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, Inches(4.0), Inches(2.4))
    border.fill.background(); border.line.color.rgb = color; border.shadow.inherit = False
    add_rect(s, x, y, Inches(4.0), Inches(0.5), color)
    add_text(s, x + Inches(0.15), y, Inches(3.7), Inches(0.5),
             name, size=15, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_bullets(s, x + Inches(0.2), y + Inches(0.6), Inches(3.6), Inches(1.3),
                items, size=11, color=DARK, bullet_color=color)
    add_text(s, x + Inches(0.2), y + Inches(1.9), Inches(3.6), Inches(0.4),
             "▎ " + use, size=10, color=GRAY)


# ============== Part 02 - 市场背景 ==============
s = new_content_slide("PART 02 · 市场背景", "上海商办与产业空间逻辑正在重构",
                      "从地段 / 形象 / 总部标签 → 成本可控 / 产业协同 / 政策支持 / 人才可达")
add_card(s, Inches(0.5), Inches(1.85), Inches(4.0), Inches(2.4),
         "城市发展阶段变化", [
             "城市更新进入深水区，存量盘活成核心命题",
             "新质生产力培育对空间载体提出新要求",
             "“五个中心”推动办公/研发/总部/服务功能升级",
             "增量开发 → 存量盘活 / 产业赋能 / 精细运营",
         ], accent=NAVY)
add_card(s, Inches(4.65), Inches(1.85), Inches(4.0), Inches(2.4),
         "市场运行逻辑变化", [
             "租金竞争 → 产业服务竞争",
             "地段价值 → 综合生态价值",
             "单一办公 → 办公+研发+展示+孵化+资本",
             "楼宇招商 → 产业招商与企业生命周期服务",
         ], accent=BLUE)
add_card(s, Inches(8.8), Inches(1.85), Inches(4.0), Inches(2.4),
         "企业选址逻辑变化", [
             "综合成本可控 + 政策可达",
             "产业协同 + 上下游生态",
             "空间灵活 + 服务完善",
             "交通便利 + 人才可达",
         ], accent=ACCENT)

# 下方：传统研究的四大不足
add_rect(s, Inches(0.5), Inches(4.45), Inches(12.3), Inches(0.4), NAVY)
add_text(s, Inches(0.5), Inches(4.45), Inches(12.3), Inches(0.4),
         "  当前研究体系的四大不足  →  本报告的突破方向",
         size=14, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
issues = [
    ("重核心区，轻全域", "覆盖16区+街镇+楼宇/园区"),
    ("重资产指标，轻需求", "供需双侧 + 企业画像 + 行为研究"),
    ("重静态，轻动态", "月度/季度企业迁徙追踪"),
    ("重描述，轻落地", "招商清单 + 资管建议 + 选址方案"),
]
for i, (a, b) in enumerate(issues):
    x = Inches(0.5 + i * 3.1)
    add_rect(s, x, Inches(4.95), Inches(2.95), Inches(0.5), LIGHT_BLUE)
    add_text(s, x + Inches(0.1), Inches(4.95), Inches(2.85), Inches(0.5), a,
             size=12, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    arrow = s.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, x + Inches(1.35),
                               Inches(5.5), Inches(0.25), Inches(0.3))
    arrow.fill.solid(); arrow.fill.fore_color.rgb = ACCENT
    arrow.line.fill.background(); arrow.shadow.inherit = False
    add_rect(s, x, Inches(5.85), Inches(2.95), Inches(1.0), NAVY)
    add_text(s, x + Inches(0.15), Inches(5.85), Inches(2.65), Inches(1.0), b,
             size=12, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)


# ============== Part 03 - 全域供给格局 ==============
s = new_content_slide("PART 03 · 全域供给", "上海全域商办与产业园区供给格局",
                      "四级数据库：行政区 → 街道/镇 → 楼宇/园区 → 企业")

# 顶部统计指标盒
metrics = [
    ("商办楼宇", "全市数量盘点", BLUE),
    ("产业园区", "全市数量盘点", NAVY),
    ("总建筑面积", "存量底数", ORANGE),
    ("可租赁面积", "市场流通", ACCENT),
    ("可招商面积", "去化压力", GREEN),
    ("新增供应", "供应节奏", RGBColor(0x7C, 0x4A, 0x9E)),
]
for i, (k, v, c) in enumerate(metrics):
    x = Inches(0.5 + i * 2.13)
    add_rect(s, x, Inches(1.85), Inches(2.0), Inches(1.1), c)
    add_text(s, x, Inches(1.85), Inches(2.0), Inches(0.4), k,
             size=12, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    add_text(s, x, Inches(2.2), Inches(2.0), Inches(0.7), v,
             size=11, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)

# 16区四象限示意
add_text(s, Inches(0.5), Inches(3.2), Inches(12.3), Inches(0.4),
         "16区供给格局（按功能分组示意）", size=14, bold=True, color=NAVY)
districts = [
    ("核心商务区", BLUE, "黄浦 · 静安 · 徐汇 · 长宁 · 虹口"),
    ("浦东主力区", NAVY, "陆家嘴 · 前滩 · 张江 · 金桥 · 外高桥 · 临港"),
    ("产业新兴区", ORANGE, "杨浦 · 普陀 · 闵行 · 嘉定 · 松江"),
    ("外溢承接区", GREEN, "宝山 · 青浦 · 奉贤 · 金山 · 崇明"),
]
for i, (name, color, regions) in enumerate(districts):
    x = Inches(0.5 + i * 3.1)
    add_rect(s, x, Inches(3.7), Inches(2.95), Inches(0.5), color)
    add_text(s, x, Inches(3.7), Inches(2.95), Inches(0.5), name,
             size=13, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    add_rect(s, x, Inches(4.2), Inches(2.95), Inches(0.8), LIGHT_BLUE)
    add_text(s, x + Inches(0.1), Inches(4.2), Inches(2.75), Inches(0.8),
             regions, size=11, color=DARK, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)

# 物业类型分布
add_text(s, Inches(0.5), Inches(5.2), Inches(12.3), Inches(0.4),
         "按物业类型划分（10大类）", size=14, bold=True, color=NAVY)
prop_types = ["超甲级写字楼", "甲级写字楼", "乙级写字楼", "商务楼宇", "总部园区",
              "科创园区", "产业园区", "孵化器/众创空间", "城市更新载体", "国企平台载体"]
for i, t in enumerate(prop_types):
    col = i % 5
    row = i // 5
    x = Inches(0.5 + col * 2.5)
    y = Inches(5.65 + row * 0.55)
    add_rect(s, x, y, Inches(2.4), Inches(0.45), WHITE)
    border = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, Inches(2.4), Inches(0.45))
    border.fill.background(); border.line.color.rgb = BLUE; border.shadow.inherit = False
    add_text(s, x, y, Inches(2.4), Inches(0.45), t,
             size=11, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)


# ============== Part 03 - 重点板块 ==============
s = new_content_slide("PART 03 · 全域供给", "重点商务与产业板块清单",
                      "20大重点板块作为深度分析样本")
panels = [
    ("陆家嘴", "金融总部"), ("前滩", "总部展示"), ("张江科学城", "硬科技/生医"),
    ("临港新片区", "硬科技/国际"), ("徐汇滨江", "AI/传媒"), ("北外滩", "金融/航运"),
    ("大虹桥", "国际商务"), ("杨浦滨江", "在线新经济"), ("五角场", "科创服务"),
    ("南京西路", "高端商务"), ("人民广场", "传统商务"), ("漕河泾", "信息技术"),
    ("紫竹", "生命健康"), ("金桥", "智造/通信"), ("外高桥", "贸易/总部"),
    ("嘉定汽车城", "智能汽车"), ("G60科创走廊", "高端制造"), ("青浦西虹桥", "数贸/进博"),
    ("宝山南大智慧城", "新材料/数字"), ("奉贤东方美谷", "美丽健康"),
]
for i, (name, focus) in enumerate(panels):
    col = i % 5
    row = i // 5
    x = Inches(0.5 + col * 2.55)
    y = Inches(1.95 + row * 1.05)
    add_rect(s, x, y, Inches(2.4), Inches(0.95), WHITE)
    border = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, Inches(2.4), Inches(0.95))
    border.fill.background(); border.line.color.rgb = NAVY; border.shadow.inherit = False
    add_rect(s, x, y, Inches(2.4), Inches(0.4), NAVY)
    add_text(s, x, y, Inches(2.4), Inches(0.4), name,
             size=13, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    add_text(s, x, y + Inches(0.45), Inches(2.4), Inches(0.5), focus,
             size=11, color=DARK, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)


# ============== Part 04 - 租金空置 ==============
s = new_content_slide("PART 04 · 租金空置", "租金、空置与综合承租成本",
                      "分层租金 + 空置结构 + 综合承租成本模型")

add_card(s, Inches(0.5), Inches(1.85), Inches(4.0), Inches(2.3),
         "分层租金体系", [
             "全市平均报价/成交租金",
             "核心 / 次核心 / 产业园区 / 乙级以下",
             "16区租金 + 各物业等级租金",
             "报价 / 中介报价 / 真实成交（三类标注）",
         ], accent=BLUE)

add_card(s, Inches(4.65), Inches(1.85), Inches(4.0), Inches(2.3),
         "空置结构分析", [
             "总体空置率 + 各区/各街镇 + 各物业",
             "小/中/大面积 · 整层/多层/整栋",
             "长期空置 vs 新增空置",
             "招商可租面积 = 报告主口径",
         ], accent=ORANGE)

add_card(s, Inches(8.8), Inches(1.85), Inches(4.0), Inches(2.3),
         "综合承租成本模型", [
             "= 租金 + 物业费 + 停车 + 装修摊销",
             "  + 搬迁成本 + 通勤成本 − 政策补贴",
             "写字楼 vs 产业园区成本对比",
             "降本型迁移路径识别",
         ], accent=GREEN)

# 空置结构表
add_text(s, Inches(0.5), Inches(4.4), Inches(12.3), Inches(0.4),
         "空置面积结构 × 招商资管含义", size=14, bold=True, color=NAVY)
structures = [
    ("小面积空置 (≤300㎡)", "初创/小微团队"),
    ("中面积 (300-1000㎡)", "成长型/中型企业"),
    ("大面积 (1000-3000㎡)", "总部/研发中心"),
    ("整层空置", "整层招商溢价"),
    ("多层连续 (≥2层)", "总部/产业基地"),
    ("整栋待招商", "产业主题改造"),
    ("长期空置 (≥12月)", "重新定位或改造"),
    ("新增空置 (3月)", "短期波动信号"),
]
for i, (k, v) in enumerate(structures):
    col = i % 4
    row = i // 4
    x = Inches(0.5 + col * 3.15)
    y = Inches(4.85 + row * 0.95)
    add_rect(s, x, y, Inches(3.0), Inches(0.85), LIGHT_BLUE)
    add_text(s, x + Inches(0.1), y + Inches(0.05), Inches(2.8), Inches(0.4),
             k, size=11, bold=True, color=NAVY)
    add_text(s, x + Inches(0.1), y + Inches(0.42), Inches(2.8), Inches(0.4),
             v, size=10, color=DARK)


# ============== Part 05 - 入驻企业 ==============
s = new_content_slide("PART 05 · 企业画像", "入驻企业画像与产业需求结构",
                      "重点关注 17 大行业 × 10 类企业规模 × 9 类企业清单")

# 17大行业
add_text(s, Inches(0.5), Inches(1.85), Inches(12.3), Inches(0.4),
         "重点行业（17大产业方向）", size=14, bold=True, color=NAVY)
industries = ["人工智能", "集成电路", "生物医药", "智能驾驶", "新能源汽车", "机器人",
              "低空经济", "软件信息", "金融服务", "专业服务", "文化传媒", "数字贸易",
              "跨境电商", "高端制造", "生产性服务业", "总部经济", "专精特新"]
for i, t in enumerate(industries):
    col = i % 6
    row = i // 6
    x = Inches(0.5 + col * 2.13)
    y = Inches(2.3 + row * 0.5)
    add_rect(s, x, y, Inches(2.0), Inches(0.42), BLUE)
    add_text(s, x, y, Inches(2.0), Inches(0.42), t,
             size=11, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)

# 行业-空间矩阵
add_text(s, Inches(0.5), Inches(3.9), Inches(12.3), Inches(0.4),
         "典型行业空间需求矩阵（示例）", size=14, bold=True, color=NAVY)
headers = ["行业", "空间形态", "选址偏好", "关注因素"]
rows = [
    ["人工智能", "中大型办公+研发展示", "徐汇滨江·杨浦·张江·临港", "人才/算力/资本/政策"],
    ["生物医药", "研发办公+实验空间", "张江·临港·闵行·奉贤", "实验/审批/产业链"],
    ["智能驾驶", "办公+研发+测试", "嘉定·浦东·杨浦·临港", "测试场景/整车/政策"],
    ["专业服务", "中小型高品质办公", "黄浦·静安·陆家嘴·北外滩", "客户可达/品牌"],
    ["高端制造", "研发办公+厂办", "嘉定·松江·临港·宝山", "成本/物流/产业配套"],
]
table_x = Inches(0.5); table_y = Inches(4.35)
col_widths = [Inches(1.8), Inches(2.6), Inches(4.2), Inches(3.7)]
# Header
cx = table_x
for i, h in enumerate(headers):
    add_rect(s, cx, table_y, col_widths[i], Inches(0.4), NAVY)
    add_text(s, cx, table_y, col_widths[i], Inches(0.4), h, size=12, bold=True,
             color=WHITE, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    cx += col_widths[i]
# rows
for ri, row in enumerate(rows):
    cx = table_x
    bg = LIGHT_BLUE if ri % 2 == 0 else WHITE
    for i, val in enumerate(row):
        add_rect(s, cx, table_y + Inches(0.4 + ri * 0.42), col_widths[i],
                 Inches(0.42), bg)
        add_text(s, cx + Inches(0.1), table_y + Inches(0.4 + ri * 0.42),
                 col_widths[i] - Inches(0.15), Inches(0.42), val,
                 size=11, color=DARK, anchor=MSO_ANCHOR.MIDDLE,
                 align=PP_ALIGN.LEFT if i == 0 else PP_ALIGN.LEFT)
        cx += col_widths[i]


# ============== Part 06 - 企业迁徙 ==============
s = new_content_slide("PART 06 · 企业迁徙", "企业迁徙与产业流动趋势",
                      "8 类迁徙行为 × 多源交叉验证 × 区域吸引力指数")

# 8类迁徙
add_text(s, Inches(0.5), Inches(1.85), Inches(12.3), Inches(0.4),
         "企业迁徙类型（8类）", size=14, bold=True, color=NAVY)
mig_types = [
    ("扩张型迁移", GREEN), ("降本型迁移", ORANGE), ("总部升级型", NAVY),
    ("产业集聚型", BLUE), ("政策导向型", ACCENT), ("被动搬迁型", GRAY),
    ("缩租型迁移", RGBColor(0xA8, 0x3A, 0x3A)), ("注册地迁移办公地不变", RGBColor(0x6B, 0x4F, 0xA1)),
]
for i, (t, c) in enumerate(mig_types):
    col = i % 4
    row = i // 4
    x = Inches(0.5 + col * 3.15)
    y = Inches(2.3 + row * 0.55)
    add_rect(s, x, y, Inches(3.0), Inches(0.48), c)
    add_text(s, x, y, Inches(3.0), Inches(0.48), t, size=12, bold=True,
             color=WHITE, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)

# 产业迁徙流向
add_text(s, Inches(0.5), Inches(3.55), Inches(12.3), Inches(0.4),
         "重点产业迁徙流向", size=14, bold=True, color=NAVY)
flows = [
    ("人工智能", "→  徐汇滨江 · 杨浦 · 张江 · 临港"),
    ("智能驾驶", "→  嘉定 · 浦东 · 杨浦 · 临港"),
    ("生物医药", "→  张江 · 临港 · 闵行 · 奉贤"),
    ("集成电路", "→  张江 · 临港 · 嘉定"),
    ("金融/专业服务", "→  陆家嘴 · 北外滩 · 南京西路（稳定）"),
    ("总部企业", "→  陆家嘴 · 前滩 · 北外滩 · 徐汇滨江"),
]
for i, (k, v) in enumerate(flows):
    col = i % 2
    row = i // 2
    x = Inches(0.5 + col * 6.4)
    y = Inches(4.0 + row * 0.55)
    add_rect(s, x, y, Inches(6.2), Inches(0.48), LIGHT_BLUE)
    add_text(s, x + Inches(0.1), y, Inches(2.0), Inches(0.48),
             k, size=12, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(2.0), y, Inches(4.1), Inches(0.48),
             v, size=12, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

# 区域吸引力指数
add_rect(s, Inches(0.5), Inches(5.75), Inches(12.3), Inches(1.2), NAVY)
add_text(s, Inches(0.7), Inches(5.8), Inches(12.0), Inches(0.4),
         "区域吸引力指数 · 9 项核心指标", size=14, bold=True, color=WHITE)
indicators = ["企业净流入数量", "重点产业企业流入", "高成长企业流入",
              "租金性价比", "政策支持力度", "交通可达性",
              "产业配套成熟度", "空间供给适配", "资本与服务集聚"]
for i, ind in enumerate(indicators):
    col = i % 3
    row = i // 3
    x = Inches(0.7 + col * 4.0)
    y = Inches(6.2 + row * 0.25)
    add_text(s, x, y, Inches(4.0), Inches(0.25),
             f"●  {ind}", size=11, color=WHITE)


# ============== Part 07 - 招商机会 ==============
s = new_content_slide("PART 07 · 供需招商", "供需匹配与三类招商机会",
                      "区域机会 + 产业机会 + 企业线索 — 报告商业化价值的核心")

# 楼宇竞争力模型 - 雷达式表格
add_text(s, Inches(0.5), Inches(1.85), Inches(6.0), Inches(0.4),
         "楼宇竞争力七维模型", size=14, bold=True, color=NAVY)
dims = [
    ("区位交通", "地铁·主干路·机场高铁"),
    ("资产品质", "建筑/层高/装修/智能化"),
    ("租金表现", "报价/成交/弹性"),
    ("企业结构", "质量/集中度/龙头"),
    ("运营服务", "物业/企业/活动"),
    ("政策资源", "补贴/资质/招商"),
    ("去化能力", "空置/周期/稳定性"),
]
for i, (k, v) in enumerate(dims):
    y = Inches(2.3 + i * 0.42)
    add_rect(s, Inches(0.5), y, Inches(1.6), Inches(0.38), BLUE)
    add_text(s, Inches(0.5), y, Inches(1.6), Inches(0.38), k, size=11,
             bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    add_rect(s, Inches(2.1), y, Inches(4.4), Inches(0.38), LIGHT_BLUE)
    add_text(s, Inches(2.2), y, Inches(4.2), Inches(0.38), v, size=11,
             color=DARK, anchor=MSO_ANCHOR.MIDDLE)

# 三类招商机会
add_text(s, Inches(6.8), Inches(1.85), Inches(6.0), Inches(0.4),
         "三类招商机会", size=14, bold=True, color=NAVY)
opps = [
    ("区域机会", BLUE, ["核心区外溢承接区域", "新兴产业适配区域",
                       "总部企业导入区域", "成本型办公承接区"]),
    ("产业机会", ORANGE, ["AI产业链补链", "智能驾驶上下游",
                         "生物医药研发服务", "专精特新集聚"]),
    ("企业线索", ACCENT, ["近期融资企业", "新增招聘明显企业",
                         "扩租/迁址可能企业", "产业链上下游企业"]),
]
for i, (name, color, items) in enumerate(opps):
    x = Inches(6.8)
    y = Inches(2.3 + i * 1.55)
    add_rect(s, x, y, Inches(6.0), Inches(0.38), color)
    add_text(s, x + Inches(0.1), y, Inches(5.8), Inches(0.38), name,
             size=12, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_bullets(s, x + Inches(0.15), y + Inches(0.45), Inches(5.8),
                Inches(1.05), items, size=11, color=DARK, bullet_color=color)


# ============== Part 08 - 趋势预测 ==============
s = new_content_slide("PART 08 · 趋势预测", "市场趋势与未来机会",
                      "租金 · 空置 · 产业 · 空间策略 — 四线预测")
add_card(s, Inches(0.5), Inches(1.85), Inches(3.0), Inches(2.4),
         "租金趋势", [
             "核心商务区走势研判",
             "乙级写字楼承压幅度",
             "产业园区韧性度",
             "新兴板块成长性",
             "老旧楼宇调整压力",
         ], accent=BLUE, body_size=11)
add_card(s, Inches(3.65), Inches(1.85), Inches(3.0), Inches(2.4),
         "空置趋势", [
             "高空置板块识别",
             "新增供应压力板块",
             "去化较快板块",
             "需求支撑较强板块",
             "存量改造压力板块",
         ], accent=ORANGE, body_size=11)
add_card(s, Inches(6.8), Inches(1.85), Inches(3.0), Inches(2.4),
         "产业需求趋势", [
             "AI · 智能驾驶 · 机器人",
             "低空经济 · 生物医药",
             "集成电路 · 数字贸易",
             "跨境电商 · 绿色低碳",
             "科技服务 · 生产性服务",
         ], accent=GREEN, body_size=11)
add_card(s, Inches(9.95), Inches(1.85), Inches(2.9), Inches(2.4),
         "空间策略变化", [
             "更灵活面积",
             "总部重品牌资源",
             "成长企业关注政策",
             "硬科技要复合空间",
             "产业社区+服务",
         ], accent=ACCENT, body_size=11)

# 未来机会
add_text(s, Inches(0.5), Inches(4.45), Inches(12.3), Inches(0.4),
         "未来五大机会场域", size=14, bold=True, color=NAVY)
chances = [
    ("核心区", "总部经济 · 金融 · 专业服务 · 外资机构", NAVY),
    ("滨江区", "AI · 数字经济 · 总部展示 · 科技服务", BLUE),
    ("张江/临港", "硬科技 · 生物医药 · 高端制造 · 智能驾驶", GREEN),
    ("郊区产业空间", "成本型外溢 · 研发制造结合 · 产业链集聚", ORANGE),
    ("老旧商办更新", "主题楼宇 · 垂直产业楼宇 · 科创服务空间", ACCENT),
]
for i, (k, v, c) in enumerate(chances):
    y = Inches(4.95 + i * 0.42)
    add_rect(s, Inches(0.5), y, Inches(2.2), Inches(0.38), c)
    add_text(s, Inches(0.5), y, Inches(2.2), Inches(0.38), k,
             size=12, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    add_rect(s, Inches(2.7), y, Inches(10.1), Inches(0.38), LIGHT_BLUE)
    add_text(s, Inches(2.85), y, Inches(9.9), Inches(0.38), v,
             size=12, color=DARK, anchor=MSO_ANCHOR.MIDDLE)


# ============== Part 09 - 政策建议 ==============
s = new_content_slide("PART 09 · 政策应用", "面向四类对象的策略建议",
                      "政府 · 园区/楼宇 · 资产方 · 企业 — 一体化应用")
suggestions = [
    ("对政府部门", NAVY, [
        "建立全域商办与产业空间动态数据库",
        "以街镇为单位动态监测空置与迁徙",
        "对高空置区域实施精准招商与功能更新",
        "推动楼宇经济、园区经济与城市更新联动",
        "用企业迁徙数据优化产业政策",
    ]),
    ("对园区/楼宇运营方", BLUE, [
        "从“租赁招商”转向“产业招商”",
        "建立企业画像与招商漏斗",
        "动态监测竞品租金和空置",
        "引入企业服务、政策、资本和产业链资源",
        "对老旧载体进行场景化改造",
    ]),
    ("对资产持有方", ORANGE, [
        "重新评估资产定位",
        "建立竞品对标体系",
        "关注真实成交租金而非挂牌租金",
        "通过产业主题提升资产溢价",
        "为资产证券化 / REITs 储备数据基础",
    ]),
    ("对企业选址方", ACCENT, [
        "建立综合承租成本模型",
        "综合比较政策/产业生态/人才可达性",
        "根据企业阶段选择空间类型",
        "成长企业关注扩租弹性",
        "总部企业关注品牌与资源链接",
    ]),
]
for i, (name, color, items) in enumerate(suggestions):
    col = i % 2
    row = i // 2
    x = Inches(0.5 + col * 6.4)
    y = Inches(1.95 + row * 2.55)
    add_card(s, x, y, Inches(6.2), Inches(2.35), name, items,
             accent=color, body_size=12)


# ============== Part 10 - 成果体系 ==============
s = new_content_slide("PART 10 · 成果体系", "成果体系与商业化产品矩阵",
                      "报告 + 数据库 + 可视化 + 商业化产品 — 一份长期研究品牌")
products = [
    ("月度市场监测简报", "月", "租金/空置/迁徙/招商热点/预警", BLUE),
    ("季度深度报告", "季", "区域/产业/迁徙/资管/趋势", NAVY),
    ("年度白皮书", "年", "全域格局 + 年度趋势 + 政策建议", ACCENT),
    ("区域专项报告", "按需", "杨浦/浦东/徐汇滨江/嘉定/临港", ORANGE),
    ("楼宇/园区诊断", "按项目", "竞品/租金/空置/招商/改造", GREEN),
    ("企业选址服务", "按企业", "区域/楼宇推荐 + 政策匹配 + 成本测算",
     RGBColor(0x7C, 0x4A, 0x9E)),
]
for i, (name, freq, content, color) in enumerate(products):
    col = i % 3
    row = i // 3
    x = Inches(0.5 + col * 4.25)
    y = Inches(1.95 + row * 2.5)
    add_rect(s, x, y, Inches(4.0), Inches(2.3), WHITE)
    border = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, Inches(4.0), Inches(2.3))
    border.fill.background(); border.line.color.rgb = color; border.shadow.inherit = False
    add_rect(s, x, y, Inches(4.0), Inches(0.55), color)
    add_text(s, x + Inches(0.15), y, Inches(2.5), Inches(0.55),
             name, size=15, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(2.5), y, Inches(1.4), Inches(0.55),
             f"周期：{freq}", size=11, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.RIGHT)
    add_text(s, x + Inches(0.2), y + Inches(0.7), Inches(3.7), Inches(1.4),
             content, size=12, color=DARK, anchor=MSO_ANCHOR.MIDDLE)


# ============== Part 11 - 六大数据库 ==============
s = new_content_slide("PART 11 · 数据体系", "六大基础数据库",
                      "支撑全域、动态、可校验、可落地的研究底座")
dbs = [
    ("楼宇与园区基础库", BLUE,
     "21字段：名称/类型/区/街镇/地址/经纬度/面积/层数/竣工/物业等级/产权/运营/招商联系"),
    ("租金与空置库", ORANGE,
     "15字段：报价/成交区间/物业费/可租/空置/楼层/装修/免租/付款/租期/更新"),
    ("企业入驻库", GREEN,
     "17字段：名称/信用代码/注册资本/成立时间/注册地址/办公地址/楼宇/行业/规模/融资/资质"),
    ("企业迁徙库", ACCENT,
     "14字段：原/新地址·楼宇·街镇/时间/迁徙类型/面积变化/租金变化/行业/规模/原因推测"),
    ("产业标签库", NAVY,
     "15字段：一/二/三级行业/战新/重点产业/新质生产力/产业链/上下游/产品/客户/专利/资质"),
    ("政策与配套库", RGBColor(0x7C, 0x4A, 0x9E),
     "15字段：区域/街镇政策/租金/装修/人才补贴/税收/科技/产业基金/落户/公共平台/交通商医教"),
]
for i, (name, color, desc) in enumerate(dbs):
    col = i % 2
    row = i // 2
    x = Inches(0.5 + col * 6.4)
    y = Inches(1.95 + row * 1.55)
    add_rect(s, x, y, Inches(6.2), Inches(1.35), WHITE)
    border = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, Inches(6.2), Inches(1.35))
    border.fill.background(); border.line.color.rgb = color; border.shadow.inherit = False
    add_rect(s, x, y, Inches(6.2), Inches(0.45), color)
    add_text(s, x + Inches(0.15), y, Inches(6.0), Inches(0.45), name,
             size=14, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.2), y + Inches(0.55), Inches(5.8), Inches(0.7),
             desc, size=11, color=DARK, anchor=MSO_ANCHOR.MIDDLE)


# ============== Part 12 - 五大指数 ==============
s = new_content_slide("PART 12 · 指标体系", "核心指标体系：五大指数",
                      "提升报告专业度与可持续发布能力")
indices = [
    ("区域市场景气指数", BLUE, "判断区域冷热 · 短期波动 · 预警",
     ["租金变化", "空置率变化", "新增成交面积", "新增企业数量", "企业净流入", "去化速度"]),
    ("楼宇/园区竞争力指数", NAVY, "横向对标 · 资产诊断 · 招商资管",
     ["区位交通", "租金性价比", "空间品质", "产业集聚",
      "企业质量", "服务能力", "政策资源", "去化表现"]),
    ("企业需求热度指数", ORANGE, "需求强弱 · 增长方向 · 招商线索",
     ["新注册企业", "融资企业", "招聘岗位增长", "扩租企业", "新租企业", "重点产业活跃"]),
    ("产业集聚度指数", GREEN, "主导产业 · 集群成熟度 · 补链强链",
     ["同行业数量", "龙头数量", "上下游完整度",
      "专精特新数量", "高新数量", "协同程度"]),
    ("供需匹配度指数", ACCENT, "楼宇适配企业 · 园区适配产业 · 精准招商",
     ["面积匹配", "租金匹配", "产业匹配", "政策匹配", "使用场景匹配"]),
]
for i, (name, color, use, items) in enumerate(indices):
    col = i % 3
    row = i // 3
    x = Inches(0.5 + col * 4.25)
    y = Inches(1.95 + row * 2.55)
    add_rect(s, x, y, Inches(4.0), Inches(2.35), WHITE)
    border = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, Inches(4.0), Inches(2.35))
    border.fill.background(); border.line.color.rgb = color; border.shadow.inherit = False
    add_rect(s, x, y, Inches(4.0), Inches(0.45), color)
    add_text(s, x + Inches(0.15), y, Inches(3.7), Inches(0.45),
             name, size=13, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.15), y + Inches(0.5), Inches(3.7), Inches(0.3),
             "▎ " + use, size=10, color=GRAY)
    add_bullets(s, x + Inches(0.2), y + Inches(0.85), Inches(3.7),
                Inches(1.45), items, size=10, color=DARK, bullet_color=color)


# ============== Part 13 - 数据调用 ==============
s = new_content_slide("PART 13 · 数据调用", "数据调用、采集与合规要点",
                      "五项原则 + 四阶段 + 合规审查")
# 五原则
add_text(s, Inches(0.5), Inches(1.85), Inches(6.0), Inches(0.4),
         "五项数据调用原则", size=14, bold=True, color=NAVY)
princ = [
    ("①  公开合规", "公开/授权/合作/调研为主"),
    ("②  多源交叉", "关键字段 ≥ 2 来源验证"),
    ("③  动态更新", "建立月度更新机制"),
    ("④  分级可信", "A/B/C/D 四级可信度"),
    ("⑤  人工校验", "AI 提效 + 关键样本人工校验"),
]
for i, (k, v) in enumerate(princ):
    y = Inches(2.3 + i * 0.55)
    add_rect(s, Inches(0.5), y, Inches(1.8), Inches(0.48), NAVY)
    add_text(s, Inches(0.5), y, Inches(1.8), Inches(0.48), k, size=12,
             bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    add_rect(s, Inches(2.3), y, Inches(4.4), Inches(0.48), LIGHT_BLUE)
    add_text(s, Inches(2.45), y, Inches(4.2), Inches(0.48), v, size=11,
             color=DARK, anchor=MSO_ANCHOR.MIDDLE)

# 四阶段
add_text(s, Inches(6.9), Inches(1.85), Inches(6.0), Inches(0.4),
         "四阶段数据采集分期", size=14, bold=True, color=NAVY)
phases = [
    ("一  中心城区试点", BLUE, "黄静徐长普虹杨 7 区"),
    ("二  浦东新区补全", NAVY, "陆家嘴/张江/临港/金桥..."),
    ("三  重点产业区", ORANGE, "闵嘉松青宝奉 + 重点板块"),
    ("四  16区全域", ACCENT, "全域动态数据库"),
]
for i, (name, color, scope) in enumerate(phases):
    y = Inches(2.3 + i * 0.85)
    add_rect(s, Inches(6.9), y, Inches(2.4), Inches(0.7), color)
    add_text(s, Inches(6.9), y, Inches(2.4), Inches(0.7), name, size=13,
             bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)
    add_rect(s, Inches(9.3), y, Inches(3.5), Inches(0.7), LIGHT_BLUE)
    add_text(s, Inches(9.45), y, Inches(3.3), Inches(0.7), scope, size=12,
             color=DARK, anchor=MSO_ANCHOR.MIDDLE)

# 难点解决
add_rect(s, Inches(0.5), Inches(5.6), Inches(12.3), Inches(1.3), NAVY)
add_text(s, Inches(0.7), Inches(5.65), Inches(12.0), Inches(0.4),
         "五大难点 → 解决方案速览", size=14, bold=True, color=WHITE)
diffs = [
    "实际办公地址 → 工商+POI+招聘+官网+水牌+物业 多源交叉",
    "成交租金难获取 → 挂牌/中介/成交 三类标注，访谈积累",
    "空置口径不统一 → 招商可租口径为主，物理/隐性辅助",
    "迁徙易误判 → 至少两个信号交叉，线下确认",
    "数据合规 → 公开优先，统计汇总，发布前合规审查",
]
for i, d in enumerate(diffs):
    add_text(s, Inches(0.7), Inches(6.05 + i * 0.18), Inches(12.0), Inches(0.18),
             f"●  {d}", size=10, color=WHITE)


# ============== Part 14 - 分工与首期 ==============
s = new_content_slide("PART 14 · 分工与首期", "联合分工 × 工作机制 × 首期落地",
                      "易居 · 复旦 · 数据技术 — 三方协同、试点先行")
# 三方分工
roles = [
    ("易居研究院", BLUE, [
        "研究框架 / 统计口径",
        "市场指标 / 区域板块研究",
        "资管 / 招商策略",
        "线下访谈调研",
        "报告撰写 / 可视化 / 对外发布",
    ]),
    ("复旦大学住房政策研究中心", NAVY, [
        "城市更新与空间治理理论",
        "政策评价体系",
        "公共政策视角与合规把关",
        "课题学术背书 / 专家研讨",
        "白皮书及政策建议联合撰写",
    ]),
    ("数据技术与 AI 采集方", ACCENT, [
        "POI/AOI 楼宇园区采集",
        "工商/招聘/公告抓取与清洗",
        "企业地址识别匹配",
        "迁徙线索识别 / 标准化",
        "数据库搭建 / 接口维护",
    ]),
]
for i, (name, color, items) in enumerate(roles):
    x = Inches(0.5 + i * 4.25)
    add_card(s, x, Inches(1.85), Inches(4.0), Inches(2.4), name, items,
             accent=color, body_size=11)

# 协同机制
add_text(s, Inches(0.5), Inches(4.45), Inches(12.3), Inches(0.4),
         "四类联合工作机制", size=14, bold=True, color=NAVY)
meetings = [
    ("周度", "数据进度会", "字段/样本/异常"),
    ("双周", "研究模型评审", "口径/指标/算法"),
    ("月度", "报告选题会", "区域/产业/案例"),
    ("季度", "成果复盘会", "准确性/反馈/规划"),
]
for i, (freq, name, content) in enumerate(meetings):
    x = Inches(0.5 + i * 3.1)
    add_rect(s, x, Inches(4.9), Inches(2.95), Inches(0.5), BLUE)
    add_text(s, x, Inches(4.9), Inches(2.95), Inches(0.5),
             f"{freq} · {name}", size=12, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_rect(s, x, Inches(5.4), Inches(2.95), Inches(0.5), LIGHT_BLUE)
    add_text(s, x, Inches(5.4), Inches(2.95), Inches(0.5), content,
             size=11, color=DARK, anchor=MSO_ANCHOR.MIDDLE,
             align=PP_ALIGN.CENTER)

# 首期试点
add_rect(s, Inches(0.5), Inches(6.05), Inches(12.3), Inches(0.85), NAVY)
add_text(s, Inches(0.7), Inches(6.08), Inches(12.0), Inches(0.4),
         "首期落地 · 《上海中心城区商办楼宇与产业园区市场试点报告》",
         size=14, bold=True, color=WHITE)
add_text(s, Inches(0.7), Inches(6.45), Inches(12.0), Inches(0.4),
         "覆盖：黄浦 · 静安 · 徐汇 · 长宁 · 普陀 · 虹口 · 杨浦  ——  完成字段、口径、迁徙模型试点",
         size=12, color=WHITE)


# ============== 封底 ==============
s = new_slide()
add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY)
add_rect(s, 0, Inches(0), SLIDE_W, Inches(0.12), ACCENT)
add_rect(s, 0, SLIDE_H - Inches(0.12), SLIDE_W, Inches(0.12), ACCENT)

add_text(s, Inches(0.8), Inches(2.0), Inches(12), Inches(1.0),
         "THANKS", size=72, bold=True, color=WHITE,
         anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
add_text(s, Inches(0.8), Inches(3.2), Inches(12), Inches(0.6),
         "感谢聆听，欢迎共建", size=22,
         color=RGBColor(0xCF, 0xDD, 0xF0),
         anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)

add_text(s, Inches(0.8), Inches(4.5), Inches(12), Inches(0.5),
         "上海商办楼宇与产业园区市场研究课题组", size=18,
         color=WHITE, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
add_text(s, Inches(0.8), Inches(5.1), Inches(12), Inches(0.5),
         "易居房地产研究院  ×  复旦大学住房政策研究中心", size=16,
         color=ACCENT, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
add_text(s, Inches(0.8), Inches(6.2), Inches(12), Inches(0.4),
         "—  讨论稿 · 仅供内部研讨使用  —", size=12,
         color=RGBColor(0xAB, 0xC0, 0xD9),
         anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)


out = os.path.join(OUTPUT_DIR, "上海商办楼宇与产业园区市场深度报告_汇报版.pptx")
prs.save(out)
print(f"PPT 已生成：{out}  ·  共 {len(prs.slides)} 页")
