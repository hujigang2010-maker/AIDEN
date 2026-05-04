"""
GS · iDrive Hub 招商方案 PPT 生成脚本（商务汇报版 · 16:9）

- 中文字体：WenQuanYi Micro Hei（系统已安装）
- 主色板：
    深海蓝 #0F2D52 （主基调 / 标题底）
    智驾蓝 #1F6FEB （强调 / 链路 / 数据条）
    金 砂 #C9A24A （重点 / 数字高亮）
    云 灰 #F4F6FA （正文背景）
    炭 黑 #1B1F2A （正文字）
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pathlib import Path

# --------------------------------------------------------------------------
# 配色与字体
# --------------------------------------------------------------------------
NAVY = RGBColor(0x0F, 0x2D, 0x52)
BLUE = RGBColor(0x1F, 0x6F, 0xEB)
GOLD = RGBColor(0xC9, 0xA2, 0x4A)
CLOUD = RGBColor(0xF4, 0xF6, 0xFA)
INK = RGBColor(0x1B, 0x1F, 0x2A)
GREY = RGBColor(0x6B, 0x73, 0x80)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xD8, 0xDE, 0xE9)
GREEN = RGBColor(0x2F, 0xA3, 0x6F)
RED = RGBColor(0xD0, 0x4A, 0x4A)

CN_FONT = "WenQuanYi Micro Hei"
EN_FONT = "Calibri"


# --------------------------------------------------------------------------
# 工具函数
# --------------------------------------------------------------------------
def set_run(run, text, *, size=14, bold=False, color=INK, font_cn=CN_FONT, font_en=EN_FONT):
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font_en
    rPr = run._r.get_or_add_rPr()
    # East Asian font
    from pptx.oxml.ns import qn
    for tag in ("ea", "cs"):
        existing = rPr.find(qn(f"a:{tag}"))
        if existing is not None:
            rPr.remove(existing)
        ea = rPr.makeelement(qn(f"a:{tag}"), {"typeface": font_cn})
        rPr.append(ea)


def add_text(slide, x, y, w, h, text, *, size=14, bold=False, color=INK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, fill=None, line=None,
             italic=False):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.margin_left = Pt(6)
    tf.margin_right = Pt(6)
    tf.margin_top = Pt(2)
    tf.margin_bottom = Pt(2)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    if fill is not None:
        box.fill.solid()
        box.fill.fore_color.rgb = fill
    else:
        box.fill.background()
    if line is not None:
        box.line.color.rgb = line
        box.line.width = Pt(0.5)
    else:
        box.line.fill.background()
    lines = text.split("\n") if isinstance(text, str) else list(text)
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        set_run(run, ln, size=size, bold=bold, color=color)
        run.font.italic = italic
    return box


def add_rect(slide, x, y, w, h, *, fill=NAVY, line=None, shadow=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.5)
    if not shadow:
        shp.shadow.inherit = False
    shp.text_frame.text = ""
    return shp


def add_round(slide, x, y, w, h, text="", *, fill=BLUE, color=WHITE, size=12, bold=True, align=PP_ALIGN.CENTER):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    tf = shp.text_frame
    tf.margin_left = Pt(6)
    tf.margin_right = Pt(6)
    tf.margin_top = Pt(2)
    tf.margin_bottom = Pt(2)
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = align
    set_run(p.add_run(), text, size=size, bold=bold, color=color)
    return shp


def slide_size_emu(prs):
    return prs.slide_width, prs.slide_height


# --------------------------------------------------------------------------
# 母版：标题条 + 页码 + 页眉
# --------------------------------------------------------------------------
def add_chrome(slide, prs, *, page_no, total, phase_label="", page_title="", subtitle=""):
    sw, sh = slide_size_emu(prs)
    # 顶部色条
    add_rect(slide, 0, 0, sw, Emu(380000), fill=NAVY)
    # 阶段徽标
    if phase_label:
        add_round(slide, Inches(0.5), Inches(0.18), Inches(2.0), Inches(0.36),
                  phase_label, fill=GOLD, color=NAVY, size=11, bold=True)
    # 页面标题
    if page_title:
        add_text(slide, Inches(2.7), Inches(0.10), Inches(8.5), Inches(0.55),
                 page_title, size=22, bold=True, color=WHITE,
                 anchor=MSO_ANCHOR.MIDDLE)
    # 副标题
    if subtitle:
        add_text(slide, Inches(2.7), Inches(0.55), Inches(8.5), Inches(0.30),
                 subtitle, size=11, color=CLOUD, anchor=MSO_ANCHOR.MIDDLE)
    # 底部细条
    add_rect(slide, 0, sh - Emu(80000), sw, Emu(80000), fill=GOLD)
    # 页脚
    add_text(slide, Inches(0.5), Inches(7.05), Inches(8.0), Inches(0.30),
             "GS · iDrive Hub · 冠松静安智能驾驶研发中心 · 招商方案 v1.0",
             size=9, color=GREY)
    add_text(slide, Inches(11.5), Inches(7.05), Inches(2.0), Inches(0.30),
             f"{page_no} / {total}", size=9, color=GREY, align=PP_ALIGN.RIGHT)


# --------------------------------------------------------------------------
# 表格生成器（统一风格）
# --------------------------------------------------------------------------
def add_table(slide, x, y, w, h, header, rows, *, header_fill=NAVY, header_color=WHITE,
              zebra=(WHITE, CLOUD), header_size=11, body_size=10, col_widths=None):
    cols = len(header)
    n_rows = len(rows) + 1
    table_shape = slide.shapes.add_table(n_rows, cols, x, y, w, h)
    table = table_shape.table
    if col_widths:
        for i, cw in enumerate(col_widths):
            table.columns[i].width = cw
    # header
    for j, htxt in enumerate(header):
        cell = table.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_fill
        cell.text = ""
        tf = cell.text_frame
        tf.margin_left = Pt(4); tf.margin_right = Pt(4)
        tf.margin_top = Pt(2); tf.margin_bottom = Pt(2)
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        set_run(p.add_run(), htxt, size=header_size, bold=True, color=header_color)
    # body
    for i, row in enumerate(rows, start=1):
        bg = zebra[(i - 1) % 2]
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            cell.text = ""
            tf = cell.text_frame
            tf.margin_left = Pt(4); tf.margin_right = Pt(4)
            tf.margin_top = Pt(2); tf.margin_bottom = Pt(2)
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j > 0 else PP_ALIGN.LEFT
            set_run(p.add_run(), str(val), size=body_size, color=INK)
    return table_shape


# --------------------------------------------------------------------------
# 主流程
# --------------------------------------------------------------------------
def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank = prs.slide_layouts[6]

    # 占位计数：先写满，再回填总页数
    SLIDES = []

    def new_slide():
        s = prs.slides.add_slide(blank)
        SLIDES.append(s)
        return s

    # ============ 1. 封面 ============
    s = new_slide()
    sw, sh = slide_size_emu(prs)
    add_rect(s, 0, 0, sw, sh, fill=NAVY)
    # 装饰条
    add_rect(s, 0, Inches(5.5), sw, Inches(0.06), fill=GOLD)
    # 几何装饰
    deco = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(10.5), Inches(-2.0), Inches(8.0), Inches(8.0))
    deco.fill.solid(); deco.fill.fore_color.rgb = BLUE
    deco.line.fill.background()
    deco.shadow.inherit = False

    add_text(s, Inches(0.8), Inches(0.6), Inches(6), Inches(0.5),
             "GS · iDrive Hub", size=18, bold=True, color=GOLD)
    add_text(s, Inches(0.8), Inches(1.6), Inches(11), Inches(1.6),
             "冠松静安智能驾驶研发中心\n招商方案", size=44, bold=True, color=WHITE)
    add_text(s, Inches(0.8), Inches(3.8), Inches(11), Inches(0.6),
             "让中心城区跑通智能驾驶最后一公里", size=22, color=CLOUD, italic=True)
    add_text(s, Inches(0.8), Inches(4.6), Inches(11), Inches(0.5),
             "5 个 Phase · 8 个任务 · 一份可执行的产业园招商蓝图",
             size=14, color=CLOUD)

    add_round(s, Inches(0.8), Inches(6.2), Inches(2.4), Inches(0.45),
              "v1.0 · 招商策划阶段", fill=GOLD, color=NAVY, size=12, bold=True)
    add_text(s, Inches(3.4), Inches(6.2), Inches(8), Inches(0.45),
             "汇报对象：集团董事会 / 静安区政府 / 链主企业",
             size=11, color=CLOUD, anchor=MSO_ANCHOR.MIDDLE)

    # ============ 2. 汇报议程 ============
    s = new_slide()
    add_chrome(s, prs, page_no=2, total=0, phase_label="议程",
               page_title="汇报议程", subtitle="约 30 分钟形成完整认知")
    items = [
        ("01", "核心结论与战略定位", "Executive Summary"),
        ("02", "Phase 1 · 策略与定位", "产业研究 + 空间规划"),
        ("03", "Phase 2 · 招商执行", "链主攻坚 + 生态漏斗 + 政府关系"),
        ("04", "Phase 3 · 品牌与活动", "9 月发布会 + 年度活动 + 媒体"),
        ("05", "Phase 4 · 商业条款", "四档收费 + 合同 + 三年财务"),
        ("06", "Phase 5 · 落地推进", "12 个月甘特 + 团队分工"),
        ("07", "投决建议", "里程碑承诺与风险对冲"),
    ]
    y0 = Inches(1.3)
    for i, (no, title, sub) in enumerate(items):
        y = y0 + Inches(0.75) * i
        add_round(s, Inches(0.8), y, Inches(0.7), Inches(0.55), no,
                  fill=NAVY, color=WHITE, size=18, bold=True)
        add_text(s, Inches(1.7), y, Inches(6), Inches(0.55), title,
                 size=18, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(7.7), y, Inches(5), Inches(0.55), sub,
                 size=13, color=GREY, anchor=MSO_ANCHOR.MIDDLE)
        # 分割线
        add_rect(s, Inches(0.8), y + Inches(0.62), Inches(11.7), Emu(15000), fill=LINE)

    # ============ 3. 核心结论一页（执行摘要） ============
    s = new_slide()
    add_chrome(s, prs, page_no=3, total=0, phase_label="01 · 核心结论",
               page_title="一页看懂：项目战略与三年目标",
               subtitle="GS · iDrive Hub · 中心城区智能驾驶研发与总部首选地")
    # 顶部 4 张大卡片
    card_w = Inches(2.95); card_h = Inches(1.55); gap = Inches(0.15)
    cards = [
        ("8 万㎡", "总建面 · A~E 五栋\n可出租 5.6 万㎡", BLUE),
        ("1.2 万㎡", "户外封闭\n智驾测试区", GOLD),
        ("¥4.5–8.5", "起始租金\n元/㎡·天 (四档)", NAVY),
        ("92%", "Y3 入驻率目标\n链主≥3 / 生态≥60", GREEN),
    ]
    x = Inches(0.5)
    for i, (big, sub, color) in enumerate(cards):
        bx = x + (card_w + gap) * i
        add_rect(s, bx, Inches(1.1), card_w, card_h, fill=color)
        add_text(s, bx, Inches(1.18), card_w, Inches(0.85), big,
                 size=34, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, bx, Inches(1.95), card_w, Inches(0.65), sub,
                 size=11, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 中部：定位 / 客群 / 差异化
    add_text(s, Inches(0.5), Inches(2.95), Inches(12.3), Inches(0.4),
             "战略定位 · 三圈层客群 · 四大差异化壁垒",
             size=15, bold=True, color=NAVY)
    add_rect(s, Inches(0.5), Inches(3.40), Inches(12.3), Emu(20000), fill=GOLD)

    block_y = Inches(3.55)
    block_h = Inches(2.6)

    add_rect(s, Inches(0.5), block_y, Inches(4.0), block_h, fill=CLOUD)
    add_text(s, Inches(0.65), block_y + Inches(0.1), Inches(3.7), Inches(0.4),
             "战略定位", size=14, bold=True, color=NAVY)
    add_text(s, Inches(0.65), block_y + Inches(0.55), Inches(3.7), Inches(2.0),
             "L4 城市 NOA + 车路云一体化 + 智驾后市场\n\n研发总部 + 测试验证 + 产业服务\n三位一体园区\n\n品牌主张：让中心城区跑通\n智能驾驶最后一公里",
             size=11, color=INK)

    add_rect(s, Inches(4.65), block_y, Inches(4.0), block_h, fill=CLOUD)
    add_text(s, Inches(4.80), block_y + Inches(0.1), Inches(3.7), Inches(0.4),
             "三圈层客群", size=14, bold=True, color=NAVY)
    add_text(s, Inches(4.80), block_y + Inches(0.55), Inches(3.7), Inches(2.0),
             "内圈 · A 栋\n链主独栋 · 1–3 家总部\n\n中圈 · B/C 栋\n算法/域控/传感器/仿真\n30–50 家生态\n\n外圈 · D/E 栋 + 测试区\n服务/投资/政府/后市场",
             size=11, color=INK)

    add_rect(s, Inches(8.80), block_y, Inches(4.0), block_h, fill=NAVY)
    add_text(s, Inches(8.95), block_y + Inches(0.1), Inches(3.7), Inches(0.4),
             "四大差异化壁垒", size=14, bold=True, color=GOLD)
    add_text(s, Inches(8.95), block_y + Inches(0.55), Inches(3.7), Inches(2.0),
             "① 中心城区罕见的「独栋+测试区」组合\n\n② 静安区「一企一策」政策包\n\n③ 冠松后市场资源闭环\n   (4S/保险/二手车)\n\n④ 地铁直达 · 30min 通达虹桥/浦东",
             size=11, color=WHITE)

    # 底部：风险提示 + 关键举措
    add_text(s, Inches(0.5), Inches(6.30), Inches(12.3), Inches(0.4),
             "关键举措：5 链主 · 300 生态 · 5 中介 · 1 场 9 月发布会 · 4 档商业模式",
             size=12, bold=True, color=NAVY)

    # ============ 4. Phase 1 章节扉页 ============
    def section_cover(idx_label, title_cn, title_en, points):
        s = new_slide()
        sw, sh = slide_size_emu(prs)
        add_rect(s, 0, 0, sw, sh, fill=NAVY)
        add_rect(s, 0, Inches(3.7), sw, Emu(40000), fill=GOLD)
        add_text(s, Inches(0.8), Inches(1.0), Inches(8), Inches(0.6),
                 idx_label, size=18, bold=True, color=GOLD)
        add_text(s, Inches(0.8), Inches(1.7), Inches(11), Inches(1.3),
                 title_cn, size=44, bold=True, color=WHITE)
        add_text(s, Inches(0.8), Inches(3.0), Inches(11), Inches(0.6),
                 title_en, size=18, color=CLOUD, italic=True)
        for i, p in enumerate(points):
            y = Inches(4.2) + Inches(0.55) * i
            add_round(s, Inches(0.8), y, Inches(0.4), Inches(0.4), str(i + 1),
                      fill=GOLD, color=NAVY, size=14, bold=True)
            add_text(s, Inches(1.4), y, Inches(11), Inches(0.4), p,
                     size=15, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
        return s

    section_cover("PHASE 1",
                  "策略与定位",
                  "Strategy & Positioning · 奠基",
                  ["任务 1 · 产业定位研究与竞品对标",
                   "任务 2 · 空间功能规划（A~E 栋 + 户外测试区）"])

    # ============ 5. 任务1 · 产业链图谱 ============
    s = new_slide()
    add_chrome(s, prs, page_no=5, total=0, phase_label="Phase 1 · 任务 1",
               page_title="智能驾驶产业链六层图谱",
               subtitle="客群匹配的『地图』：从基础设施到应用运营")

    layers = [
        ("L6 应用与运营", "Robotaxi · Robobus · Robotruck · 矿区港口 · 末端配送", BLUE),
        ("L5 整车与方案", "车企 ADAS/AD 部门 · 全栈方案商（华为/Momenta/...）", NAVY),
        ("L4 软件与算法", "感知 · 规控 · 定位 · 端到端大模型 · 仿真 · 数据闭环", BLUE),
        ("L3 域控与计算", "智驾域控 · 中央计算 · 芯片（地平线/黑芝麻/英伟达）", NAVY),
        ("L2 传感器", "激光雷达 · 4D 毫米波 · 摄像头 · 高精地图 · IMU/RTK", BLUE),
        ("L1 基础设施", "路侧 RSU · V2X · 5G/卫星通信 · 高精地图测绘资质", NAVY),
    ]
    bar_x = Inches(0.7); bar_w = Inches(8.5); bar_h = Inches(0.65)
    for i, (lbl, desc, c) in enumerate(layers):
        y = Inches(1.25) + (bar_h + Inches(0.10)) * i
        add_rect(s, bar_x, y, Inches(1.6), bar_h, fill=GOLD)
        add_text(s, bar_x, y, Inches(1.6), bar_h, lbl,
                 size=12, bold=True, color=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, bar_x + Inches(1.6), y, bar_w - Inches(1.6), bar_h, fill=c)
        add_text(s, bar_x + Inches(1.7), y, bar_w - Inches(1.7), bar_h, desc,
                 size=11, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

    # 右侧三个支撑模块
    side_x = Inches(9.5); side_w = Inches(3.4)
    side = [("资本 / 孵化", "国调/中金/高瓴/园区基金"),
            ("测试 / 认证", "中汽研 · TüV · CATARC"),
            ("后市场服务", "保险 · 维修 · 改装 · 二手车 (冠松)")]
    for i, (t, d) in enumerate(side):
        y = Inches(1.6) + Inches(1.45) * i
        add_rect(s, side_x, y, side_w, Inches(1.25), fill=CLOUD, line=LINE)
        add_text(s, side_x + Inches(0.15), y + Inches(0.1), side_w - Inches(0.3), Inches(0.4),
                 t, size=14, bold=True, color=NAVY)
        add_text(s, side_x + Inches(0.15), y + Inches(0.5), side_w - Inches(0.3), Inches(0.7),
                 d, size=11, color=INK)

    add_text(s, Inches(0.7), Inches(6.4), Inches(12), Inches(0.4),
             "→ A 栋承接 L5 链主 · B 栋承接 L3/L4 · C 栋承接仿真/数据 · D 栋承接服务集群 · E 栋承接冠松后市场",
             size=12, bold=True, color=BLUE)

    # ============ 6. 任务1 · 静安/跨区竞品对标 ============
    s = new_slide()
    add_chrome(s, prs, page_no=6, total=0, phase_label="Phase 1 · 任务 1",
               page_title="静安区内 + 跨区竞品对标",
               subtitle="核心结论：中心城区独栋+测试区组合稀缺，租金锚定 6.5–8.5 元/㎡·天")
    header = ["园区", "区位", "主导产业", "租金 (元/㎡·天)", "入驻率", "政策亮点", "对我方"]
    rows = [
        ["市北高新 · 大数据基地", "静安", "大数据/AI/智算", "6.0–8.5", "~88%", "区级 5%~10% 租补 · 公寓 800 套", "无智驾产业链 / 无测试区"],
        ["静安国际中心片区", "静安", "总部金融", "9.0–13.0", "~90%", "总部经济奖励", "无产业属性 / 不可改造"],
        ["大宁国际办公", "静安", "商务/文创", "6.5–8.0", "~85%", "一般", "非产业园定位"],
        ["嘉定 · 创新中心 (新能港)", "嘉定", "智驾全产业链", "3.5–5.0", "~95%", "国家级测试场 + 牌照", "牌照不可复制 · 我方区位优"],
        ["临港 · AI 创新港", "浦东", "L4/Robotaxi", "3.0–4.5", "~80%", "装补 800–1500/㎡ · 全域开放", "政策强 · 距市区 1.5h"],
        ["张江 · IC 设计园 / AI 岛", "浦东", "芯片/AI", "6.5–8.5", "~92%", "张江专项 + IC 补贴", "算力强 · 智驾整车弱"],
        ["徐汇 · 西岸 AI 走廊", "徐汇", "AI 大模型", "8.5–11.0", "~95%", "模速空间补贴", "单价高 · 无独栋"],
    ]
    add_table(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(4.6), header, rows,
              col_widths=[Inches(2.4), Inches(0.8), Inches(1.5), Inches(1.4),
                          Inches(0.9), Inches(2.7), Inches(2.6)])
    # 结论卡
    add_rect(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.95), fill=NAVY)
    add_text(s, Inches(0.7), Inches(6.05), Inches(12.0), Inches(0.4),
             "我方差异化定价策略", size=13, bold=True, color=GOLD)
    add_text(s, Inches(0.7), Inches(6.40), Inches(12.0), Inches(0.55),
             "基础租金（6.5–8.5）+ 政策返还（区级留成 80% 三年返）+ 服务积分（测试/算力/招聘） — 等效净价低于南西核心 25%",
             size=11, color=WHITE)

    # ============ 7. 任务2 · 楼栋分配（平面 + 面积表） ============
    s = new_slide()
    add_chrome(s, prs, page_no=7, total=0, phase_label="Phase 1 · 任务 2",
               page_title="A~E 栋功能拆分 + 户外测试区",
               subtitle="独栋链主 / 生态联办 / 共享研发 / 产业服务 / 冠松总部 + 1.2 万㎡ 测试区")
    # 左：示意平面（四个栋 + 测试区）
    base_x = Inches(0.5); base_y = Inches(1.2)
    plot_w = Inches(7.0); plot_h = Inches(5.4)
    add_rect(s, base_x, base_y, plot_w, plot_h, fill=CLOUD, line=LINE)
    add_text(s, base_x, base_y + Inches(0.05), plot_w, Inches(0.3),
             "园区平面示意（北 ↑）", size=11, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    # 4 个建筑块
    bx, by = base_x + Inches(0.4), base_y + Inches(0.5)
    bw, bh = Inches(1.5), Inches(1.4)
    # D
    add_rect(s, bx, by, bw, bh, fill=NAVY)
    add_text(s, bx, by, bw, bh, "D 栋\n产业服务\n10,000 ㎡",
             size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # A
    ax = base_x + plot_w - bw - Inches(0.4)
    add_rect(s, ax, by, bw, bh, fill=GOLD)
    add_text(s, ax, by, bw, bh, "A 栋\n链主独栋\n18,000 ㎡",
             size=11, bold=True, color=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 中央广场
    cx = bx + bw + Inches(0.3); cw = ax - cx - Inches(0.3)
    add_rect(s, cx, by, cw, bh, fill=WHITE, line=LINE)
    add_text(s, cx, by, cw, bh, "中央广场 / 路演灯柱\n林荫水景",
             size=10, color=GREY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # C
    cy = by + bh + Inches(0.2)
    add_rect(s, bx, cy, bw, bh, fill=BLUE)
    add_text(s, bx, cy, bw, bh, "C 栋\n共享研发\n12,000 ㎡",
             size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # B
    add_rect(s, ax, cy, bw, bh, fill=BLUE)
    add_text(s, ax, cy, bw, bh, "B 栋\n生态联办\n16,000 ㎡",
             size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 中部连廊
    add_rect(s, cx, cy + Inches(0.5), cw, Inches(0.4), fill=CLOUD, line=LINE)
    add_text(s, cx, cy + Inches(0.5), cw, Inches(0.4), "连廊 / 食堂",
             size=10, color=GREY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 测试区 + E 栋
    ty = cy + bh + Inches(0.25)
    test_w = bw + cw + Inches(0.3)
    add_rect(s, bx, ty, test_w, Inches(1.6), fill=GREEN)
    add_text(s, bx, ty, test_w, Inches(1.6),
             "户外封闭测试区  12,000 ㎡\n城市道 / 极限场景 / 泊车 / 雨雾夜 / V2X / 监控塔",
             size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, ax, ty, bw, Inches(1.6), fill=NAVY)
    add_text(s, ax, ty, bw, Inches(1.6), "E 栋\n冠松总部\n14,000 ㎡",
             size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 右：面积表
    header = ["区域", "面积 (㎡)", "性质"]
    rows = [
        ["A 栋 链主独栋", "18,000", "出租 (可竖切 2 家)"],
        ["B 栋 生态联办", "16,000", "出租"],
        ["C 栋 共享研发", "12,000", "出租 + 工位"],
        ["D 栋 产业服务", "10,000", "出租 + 公共"],
        ["E 栋 冠松总部", "14,000", "自留"],
        ["户外测试区", "12,000", "公共/会员"],
        ["公共/连廊/绿化", "4,500", "公共"],
        ["地下停车/库房", "12,000", "部分商业"],
        ["合计建筑面积", "约 80,000", "—"],
        ["可出租净面积", "约 56,000", "—"],
    ]
    add_table(s, Inches(7.8), Inches(1.2), Inches(5.0), Inches(5.4),
              header, rows,
              col_widths=[Inches(2.0), Inches(1.4), Inches(1.6)])

    # ============ 8. Phase 2 章节扉页 ============
    section_cover("PHASE 2",
                  "招商执行（核心）",
                  "Leasing Execution · Anchor + Ecosystem + Government",
                  ["任务 3 · 链主 TOP5 攻坚",
                   "任务 4 · 300 家生态招商漏斗",
                   "任务 5 · 政府关系与政策包"])

    # ============ 9. 任务3 · 链主 TOP5 攻坚 ============
    s = new_slide()
    add_chrome(s, prs, page_no=9, total=0, phase_label="Phase 2 · 任务 3",
               page_title="链主 TOP5 攻坚作战图",
               subtitle="项目总监 + GR 总监 双人出动 · 周一例会 / 季度董事会复盘")
    header = ["#", "链主", "面积 (㎡)", "决策人 / 关键人", "当前阶段", "下一动作", "截止", "状态"]
    rows = [
        ["1", "华为车 BU", "8,000–12,000", "上海中心负责人 / GTS 上海", "T0 名片建联", "区委书记+市经信委约见", "T+30d", "推进中"],
        ["2", "百度 Apollo", "4,000–6,000", "IDG 总裁 / 华东负责人", "已发邀约函", "联合投促办赴京拜访", "T+45d", "顺利"],
        ["3", "小鹏汽车", "3,000–5,000", "李力耘 / 上海负责人", "公司层初接触", "邀请出席 9 月发布会", "T+60d", "推进中"],
        ["4", "地平线", "4,000–8,000", "余凯 / 黄畅 / 政企", "高层已建联", "邀请参观测试区 + V2X", "T+30d", "顺利"],
        ["5", "Momenta", "5,000–8,000", "曹旭东 / GR VP", "已建联，等会面", "曹 CEO + 静安区长会面", "T+45d", "推进中"],
    ]
    add_table(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(2.6),
              header, rows,
              col_widths=[Inches(0.4), Inches(1.5), Inches(1.4), Inches(2.0),
                          Inches(1.6), Inches(2.6), Inches(1.0), Inches(1.8)])

    # 中部：每家"为什么选静安、为什么选冠松"四象限
    add_text(s, Inches(0.5), Inches(4.0), Inches(12), Inches(0.4),
             "每家「一页纸」定制提案 · 四要点结构",
             size=14, bold=True, color=NAVY)
    add_rect(s, Inches(0.5), Inches(4.42), Inches(12.3), Emu(20000), fill=GOLD)

    quad_y = Inches(4.6)
    quad_h = Inches(2.3)
    quads = [
        ("决策人 / 关键人", "30 秒电梯演讲 → CEO/总裁层；GR 把区领导接口排进议程", BLUE),
        ("正缺什么", "中心城区落点 · 测试场地 · 后市场闭环 · 政企协同", NAVY),
        ("独家供给", "A 栋整栋冠名 · 户外测试区 · 静安一企一策 · 冠松 4S 流量", GOLD),
        ("报价权益", "起始租 4.5–5.5 · 免租 18–24m · 装补 1,000–1,500/㎡ · 留成 80% 三年返", GREEN),
    ]
    qw = Inches(2.95); gx = Inches(0.15)
    for i, (t, d, c) in enumerate(quads):
        x = Inches(0.5) + (qw + gx) * i
        add_rect(s, x, quad_y, qw, quad_h, fill=c)
        add_text(s, x + Inches(0.15), quad_y + Inches(0.15), qw - Inches(0.3), Inches(0.5),
                 t, size=14, bold=True, color=WHITE)
        add_text(s, x + Inches(0.15), quad_y + Inches(0.7), qw - Inches(0.3), quad_h - Inches(0.85),
                 d, size=11, color=WHITE)

    # ============ 10. 任务4 · 300 家生态漏斗 ============
    s = new_slide()
    add_chrome(s, prs, page_no=10, total=0, phase_label="Phase 2 · 任务 4",
               page_title="生态企业招商漏斗 · 300 家库",
               subtitle="5 家中介 · 9 类来源 · 端到端转化 17% · 60 家年签约目标")

    # 左：漏斗（横向条状）
    funnel = [
        ("L1 线索 (Lead)", 360, 12.0, BLUE),
        ("L2 触达 (Reached)", 290, 9.7, NAVY),
        ("L3 意向 (Intent)", 145, 4.8, BLUE),
        ("L4 谈判 (Negotiation)", 87, 2.9, NAVY),
        ("L5 签约 (Signed)", 60, 2.0, GOLD),
        ("L6 入驻 (Move-in)", 57, 1.9, GREEN),
    ]
    fy = Inches(1.25)
    for i, (lbl, n, w_inch, c) in enumerate(funnel):
        y = fy + Inches(0.65) * i
        add_rect(s, Inches(0.5), y, Inches(3.2), Inches(0.55), fill=CLOUD, line=LINE)
        add_text(s, Inches(0.55), y, Inches(3.1), Inches(0.55), lbl,
                 size=11, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, Inches(3.8), y, Inches(w_inch * 0.55), Inches(0.55), fill=c)
        add_text(s, Inches(3.8), y, Inches(w_inch * 0.55) + Inches(1.2), Inches(0.55),
                 f"  {n} 家", size=12, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

    add_text(s, Inches(0.5), Inches(5.5), Inches(10), Inches(0.4),
             "理论端到端转化（线索 → 签约）： 80% × 50% × 60% × 70% ≈ 17%",
             size=12, bold=True, color=NAVY)
    add_text(s, Inches(0.5), Inches(5.92), Inches(10), Inches(0.4),
             "60 家签约 ⇐ 360 有效线索 ⇐ 300 库 + 自拓溢出",
             size=12, color=BLUE)

    # 右：5 家中介
    header = ["中介", "侧重", "佣金档", "年度任务"]
    rows = [
        ["戴德梁行 C&W", "跨国/外资 ADAS", "首月 100% / 长租 120%", "≥ 60 家"],
        ["仲量联行 JLL", "互联网/科技搬迁", "同上", "≥ 60 家"],
        ["高力 Colliers", "C+ 腰部企业", "同上", "≥ 50 家"],
        ["世邦 CBRE", "链主 / 整层", "链主最高 150%", "≥ 40 家"],
        ["本地精品行", "中小生态/工位", "首月 80%", "≥ 90 家"],
    ]
    add_text(s, Inches(8.2), Inches(1.0), Inches(5), Inches(0.35),
             "5 家中介渠道 · 非独家 + 30 天首报", size=13, bold=True, color=NAVY)
    add_table(s, Inches(8.2), Inches(1.4), Inches(4.8), Inches(3.1),
              header, rows,
              col_widths=[Inches(1.6), Inches(1.4), Inches(1.0), Inches(0.8)])
    # 库结构小图
    add_text(s, Inches(8.2), Inches(4.7), Inches(5), Inches(0.35),
             "300 家库 · 产业链层级配比", size=13, bold=True, color=NAVY)
    bars = [("L4 算法", 60), ("L2 传感器", 50), ("L3 域控", 35), ("仿真/数据", 35),
            ("后市场", 30), ("投资/法律", 30), ("高校/科研", 30), ("L5 整车", 15),
            ("测试认证", 15)]
    by0 = Inches(5.05)
    for i, (lbl, n) in enumerate(bars):
        y = by0 + Inches(0.20) * i
        add_text(s, Inches(8.2), y, Inches(1.2), Inches(0.18), lbl, size=9, color=INK,
                 anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, Inches(9.5), y + Inches(0.04), Inches(n / 60.0 * 2.4), Inches(0.12), fill=BLUE)
        add_text(s, Inches(12.0), y, Inches(0.8), Inches(0.18), str(n), size=9, color=NAVY,
                 anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)

    # ============ 11. 任务5 · 政府关系 ============
    s = new_slide()
    add_chrome(s, prs, page_no=11, total=0, phase_label="Phase 2 · 任务 5",
               page_title="政府关系对接 · 路径地图与政策包",
               subtitle="区/市双层 · 7 步首轮汇报 · 形成「一企一策」")

    # 左：路径地图
    add_text(s, Inches(0.5), Inches(1.15), Inches(6), Inches(0.4),
             "对接路径地图（市/区双层）", size=14, bold=True, color=NAVY)
    add_rect(s, Inches(0.5), Inches(1.55), Inches(6.2), Inches(5.3), fill=CLOUD, line=LINE)
    nodes = [
        ("市委市政府", Inches(2.5), Inches(1.7), Inches(2.0), Inches(0.45), NAVY, WHITE),
        ("市经信委", Inches(0.7), Inches(2.5), Inches(1.6), Inches(0.45), BLUE, WHITE),
        ("市公安/交通委", Inches(2.5), Inches(2.5), Inches(2.0), Inches(0.45), BLUE, WHITE),
        ("市科委/发改", Inches(4.7), Inches(2.5), Inches(1.8), Inches(0.45), BLUE, WHITE),
        ("静安区委区政府", Inches(2.3), Inches(3.5), Inches(2.4), Inches(0.5), GOLD, NAVY),
        ("区投促办", Inches(0.7), Inches(4.4), Inches(1.6), Inches(0.45), NAVY, WHITE),
        ("区经委/科委", Inches(2.5), Inches(4.4), Inches(2.0), Inches(0.45), NAVY, WHITE),
        ("区财政/税务", Inches(4.7), Inches(4.4), Inches(1.8), Inches(0.45), NAVY, WHITE),
        ("区交警支队", Inches(0.7), Inches(5.1), Inches(1.6), Inches(0.45), BLUE, WHITE),
        ("区人社/房管", Inches(2.5), Inches(5.1), Inches(2.0), Inches(0.45), BLUE, WHITE),
        ("区国资/街道", Inches(4.7), Inches(5.1), Inches(1.8), Inches(0.45), BLUE, WHITE),
    ]
    for txt, x, y, w, h, fc, tc in nodes:
        add_round(s, x, y, w, h, txt, fill=fc, color=tc, size=10, bold=True)

    add_text(s, Inches(0.7), Inches(5.9), Inches(6), Inches(0.4),
             "7 步首轮汇报 · 30 天完成全覆盖", size=12, bold=True, color=NAVY)
    add_text(s, Inches(0.7), Inches(6.25), Inches(6), Inches(0.6),
             "区投促办 → 副区长 → 区委书记/区长 → 市经信委 → 市公安交警 → 市科委/发改 → 区四套班子专题会",
             size=10, color=INK)

    # 右：政策包六维
    add_text(s, Inches(7.0), Inches(1.15), Inches(6), Inches(0.4),
             "政策包 · 六维适配清单", size=14, bold=True, color=NAVY)
    pol = [
        ("牌照 / 测试", "封闭场地备案 · 1.5km 路测延伸 · L3/L4 试点联合体", BLUE),
        ("财税返还", "区级留成 80% 三年返 / 50% 后两年返", GOLD),
        ("人才 / 公寓", "链主 200 套 + 生态 300 套 · 落户绿通 80/年", NAVY),
        ("数据合规", "数交所沙盒 · 算力券最高 500 万 · 出境绿通", BLUE),
        ("产业基金", "区基金 1:1 配资 · 国调/中金/高瓴 LP 5–10 亿", GREEN),
        ("用地 / 配套", "测试区合规使用 · 周边路侧充电 · 一址多照", NAVY),
    ]
    for i, (t, d, c) in enumerate(pol):
        y = Inches(1.55) + Inches(0.85) * i
        add_rect(s, Inches(7.0), y, Inches(0.2), Inches(0.75), fill=c)
        add_rect(s, Inches(7.2), y, Inches(5.7), Inches(0.75), fill=CLOUD, line=LINE)
        add_text(s, Inches(7.35), y + Inches(0.05), Inches(5.3), Inches(0.32),
                 t, size=12, bold=True, color=NAVY)
        add_text(s, Inches(7.35), y + Inches(0.36), Inches(5.3), Inches(0.4),
                 d, size=10, color=INK)

    # ============ 12. Phase 3 章节扉页 ============
    section_cover("PHASE 3",
                  "品牌与活动",
                  "Brand & Events · Make Noise, Build Trust",
                  ["任务 6 · 9 月旗舰发布会 + 年度 10 场活动 + 媒体清单"])

    # ============ 13. 任务6 · 9月发布会 ============
    s = new_slide()
    add_chrome(s, prs, page_no=13, total=0, phase_label="Phase 3 · 任务 6",
               page_title="9 月旗舰发布会 · 「中心城区 · 智驾未来」",
               subtitle="200 人现场 + 200 人在线 · 政企签约 · 政策包发布")
    # 议程
    agenda = [
        ("13:00–13:50", "注册 / 户外测试区分批参观"),
        ("14:00–14:25", "开场 + 静安区领导致辞"),
        ("14:25–14:35", "市经信委致辞"),
        ("14:35–14:55", "主旨发布：GS · iDrive Hub 产业定位"),
        ("14:55–15:30", "链主主题演讲 ×2"),
        ("15:30–16:00", "重大签约仪式（5–8 家）"),
        ("16:00–16:20", "茶歇 + 户外实车演示（直播）"),
        ("16:20–17:30", "圆桌：城市 NOA / 数据合规 / 中心城区基础设施"),
        ("17:30–17:45", "政策包发布：「iDrive · 静安 10 条」"),
        ("17:45–18:00", "园区生态启动 + 集体合影"),
        ("18:30–20:30", "招待晚宴 + 闭门 1v1"),
    ]
    add_text(s, Inches(0.5), Inches(1.2), Inches(6), Inches(0.4),
             "议程（9 月第三周 周四 14:00–20:30）", size=14, bold=True, color=NAVY)
    for i, (t, d) in enumerate(agenda):
        y = Inches(1.65) + Inches(0.40) * i
        add_rect(s, Inches(0.5), y, Inches(1.5), Inches(0.32), fill=NAVY)
        add_text(s, Inches(0.5), y, Inches(1.5), Inches(0.32), t,
                 size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, Inches(2.05), y, Inches(4.55), Inches(0.32), fill=CLOUD, line=LINE)
        add_text(s, Inches(2.15), y, Inches(4.45), Inches(0.32), d,
                 size=10, color=INK, anchor=MSO_ANCHOR.MIDDLE)

    # 右：受邀人结构 + 倒推时间表
    add_text(s, Inches(7.0), Inches(1.2), Inches(6), Inches(0.4),
             "受邀 200 人 · 结构", size=14, bold=True, color=NAVY)
    invitees = [("政府领导", 25), ("链主 CEO", 30), ("生态代表", 60), ("投资机构", 25),
                ("高校研究院", 15), ("中介渠道", 10), ("媒体", 25), ("内部", 10)]
    for i, (lbl, n) in enumerate(invitees):
        y = Inches(1.65) + Inches(0.32) * i
        add_text(s, Inches(7.0), y, Inches(1.6), Inches(0.28), lbl,
                 size=10, color=INK, anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, Inches(8.7), y + Inches(0.05), Inches(n / 60.0 * 3.0), Inches(0.18), fill=GOLD)
        add_text(s, Inches(11.8), y, Inches(0.8), Inches(0.28), str(n),
                 size=10, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)

    add_text(s, Inches(7.0), Inches(4.3), Inches(6), Inches(0.4),
             "倒推时间表（T = 发布日）", size=14, bold=True, color=NAVY)
    timeline = ["T-90d 立项", "T-60d 物料一稿 / 邀请函", "T-45d 签约项目锁定",
                "T-30d 彩排 1", "T-21d 内容定稿", "T-14d 彩排 2 + 直播",
                "T-7d 彩排 3", "T-1d 全员彩排", "T 日 执行", "T+7d 复盘"]
    for i, t in enumerate(timeline):
        y = Inches(4.7) + Inches(0.22) * i
        add_round(s, Inches(7.0), y, Inches(0.3), Inches(0.18), "●",
                  fill=BLUE, color=WHITE, size=8, bold=True)
        add_text(s, Inches(7.4), y - Inches(0.01), Inches(5.5), Inches(0.22),
                 t, size=10, color=INK, anchor=MSO_ANCHOR.MIDDLE)

    # ============ 14. 年度活动日历 ============
    s = new_slide()
    add_chrome(s, prs, page_no=14, total=0, phase_label="Phase 3 · 任务 6",
               page_title="年度 10 场活动日历 + 媒体四层矩阵",
               subtitle="1 大会 + 4 季度沙龙 + 4 月度生态闭门会 + 1 国际峰会")
    header = ["#", "月份", "活动", "形式 / 规模", "主题方向"]
    rows = [
        ["1", "2026-05", "园区品牌发布会（启动）", "现场 80 人", "项目首秀 / 链主候选官宣"],
        ["2", "2026-06", "城市 NOA 商业化沙龙", "闭门 30 人", "链主 + 算法 + 出行"],
        ["3", "2026-07", "数据合规与算力训练营", "工作坊 50 人", "数据沙盒 + 算力券"],
        ["4", "2026-08", "智驾后市场闭门会（冠松特色）", "闭门 40 人", "保险 / 改装 / 售后"],
        ["5", "2026-09", "GS · iDrive Hub 全球招商发布会", "200 + 在线 200", "政企签约 / 政策包"],
        ["6", "2026-10", "仿真与端到端模型沙龙", "半日 60 人", "仿真生态 / 算力联训"],
        ["7", "2026-11", "智驾投融资 Demo Day", "半日 80 人", "早期算法 / 域控"],
        ["8", "2026-12", "年终招商盘点 + 入驻晚宴", "晚宴 120 人", "全年成果 / 续约"],
        ["9", "2027-01", "国际智驾产业峰会", "现场 250 人", "跨国车企 / 外资 ADAS"],
        ["10", "2027-03", "春季产业开放日 + 测试区开放周", "系列 5 天", "公众/媒体/高校"],
    ]
    add_table(s, Inches(0.5), Inches(1.2), Inches(8.0), Inches(5.3),
              header, rows,
              col_widths=[Inches(0.4), Inches(1.0), Inches(2.6), Inches(1.7), Inches(2.3)])

    # 右：媒体四层
    add_text(s, Inches(8.8), Inches(1.2), Inches(4.5), Inches(0.4),
             "媒体四层矩阵", size=14, bold=True, color=NAVY)
    layers2 = [
        ("权威 · 政策", "央视 / 人民日报 / 新华社 / 解放日报 / 上观", NAVY),
        ("财经 · 商业", "财新 / 一财 / 21 世纪 / 36Kr / 钛媒体 / 界面", BLUE),
        ("行业 · 垂类", "汽车之家 / 懂车帝 / 高工智能 / 焉知 / 36 个公众号", GOLD),
        ("国际 · 英文", "Reuters / Bloomberg / FT / Nikkei / TechCrunch", GREEN),
    ]
    for i, (t, d, c) in enumerate(layers2):
        y = Inches(1.65) + Inches(1.20) * i
        add_rect(s, Inches(8.8), y, Inches(4.5), Inches(1.05), fill=CLOUD, line=LINE)
        add_rect(s, Inches(8.8), y, Inches(0.18), Inches(1.05), fill=c)
        add_text(s, Inches(9.05), y + Inches(0.08), Inches(4.2), Inches(0.4),
                 t, size=13, bold=True, color=NAVY)
        add_text(s, Inches(9.05), y + Inches(0.45), Inches(4.2), Inches(0.6),
                 d, size=10, color=INK)

    # ============ 15. Phase 4 章节扉页 ============
    section_cover("PHASE 4",
                  "商业条款",
                  "Commercial · 4 Tiers · 3-Year Model",
                  ["任务 7 · 四档收费 + 合同框架 + 三年财务测算"])

    # ============ 16. 任务7 · 四档收费 ============
    s = new_slide()
    add_chrome(s, prs, page_no=16, total=0, phase_label="Phase 4 · 任务 7",
               page_title="四档收费方案矩阵",
               subtitle="链主月费+政策返还 / 生态月费 / 共享月费+服务佣金 / 对赌")
    cards = [
        ("甲档", "链主独栋", "月费 + 政策返还", NAVY,
         ["起始租金 4.5–5.5 元/㎡·天", "免租 18–24 个月", "装补 1,000–1,500/㎡",
          "区级税收 80% 三年返", "整栋冠名 + 测试区终身免费"]),
        ("乙档", "生态联办", "月费", BLUE,
         ["起始租金 6.5–8.0 元/㎡·天", "免租 6–12 个月", "装补 300–800/㎡",
          "合同 3+3+1 年", "测试区会员价 / 路演免费"]),
        ("丙档", "共享研发", "月费 + 服务佣金", GOLD,
         ["工位 1,800–2,800/月", "小独间 8.0–9.0 元/㎡·天", "月付制 / 免租 0–2 月",
          "服务 GMV 5%–10% 抽成", "算力券 7 折 / 共享实验室"]),
        ("丁档", "对赌", "低租 + 收益分成", GREEN,
         ["折扣租金 60%–70%", "营收 / 利润 / 估值对赌",
          "超额营收 5%–10% 返现", "园区基金 1%–3% 跟投权",
          "未达标差额补偿（双向上限）"]),
    ]
    cw = Inches(2.95); gx = Inches(0.15); cy = Inches(1.20); ch = Inches(5.4)
    for i, (lvl, name, sub, c, bullets) in enumerate(cards):
        x = Inches(0.5) + (cw + gx) * i
        add_rect(s, x, cy, cw, ch, fill=CLOUD, line=LINE)
        add_rect(s, x, cy, cw, Inches(0.95), fill=c)
        add_text(s, x, cy + Inches(0.05), cw, Inches(0.4), lvl,
                 size=18, bold=True, color=GOLD if c == NAVY else WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x, cy + Inches(0.40), cw, Inches(0.4), name,
                 size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(s, x, cy + Inches(0.65), cw, Inches(0.3), sub,
                 size=10, color=WHITE, align=PP_ALIGN.CENTER)
        for j, b in enumerate(bullets):
            by = cy + Inches(1.10) + Inches(0.65) * j
            add_round(s, x + Inches(0.15), by + Inches(0.18), Inches(0.18), Inches(0.18), "●",
                      fill=c, color=WHITE, size=8, bold=True)
            add_text(s, x + Inches(0.40), by, cw - Inches(0.5), Inches(0.55), b,
                     size=10, color=INK)

    # ============ 17. 三年财务测算 ============
    s = new_slide()
    add_chrome(s, prs, page_no=17, total=0, phase_label="Phase 4 · 任务 7",
               page_title="三年财务测算 · 关键 KPI",
               subtitle="Y3 入驻率 92% / 收入 2.5 亿 / EBITDA 由负转正 ≈ 1.25 亿")

    header = ["科目（万元）", "Y1 2026", "Y2 2027", "Y3 2028"]
    rows = [
        ["租金收入", "3,433", "7,438", "11,663"],
        ["物业费收入", "658", "1,222", "1,730"],
        ["测试区/服务/后市场", "900", "2,700", "5,600"],
        ["政策返还（净计入）", "200", "1,500", "4,500"],
        ["基金管理费", "0", "1,000", "1,500"],
        ["总收入", "5,191", "13,860", "24,993"],
        ["总成本", "9,300", "11,200", "12,500"],
        ["EBITDA", "−4,109", "+2,660", "+12,493"],
        ["EBIT", "−7,109", "−340", "+9,493"],
        ["税前利润", "−8,609", "−1,740", "+8,293"],
    ]
    add_table(s, Inches(0.5), Inches(1.2), Inches(7.0), Inches(4.5),
              header, rows,
              col_widths=[Inches(2.4), Inches(1.5), Inches(1.5), Inches(1.6)])

    # 右：KPI 三个里程碑
    kpis = [
        ("入驻率", ["35%", "65%", "92%"], BLUE),
        ("链主签约（累计）", ["1", "2", "3"], NAVY),
        ("生态签约（累计）", ["18", "45", "78"], GOLD),
        ("EBITDA 利润率", ["−79%", "+19%", "+50%"], GREEN),
    ]
    add_text(s, Inches(7.8), Inches(1.2), Inches(5.5), Inches(0.4),
             "三年关键 KPI", size=14, bold=True, color=NAVY)
    for i, (lbl, vals, c) in enumerate(kpis):
        y = Inches(1.65) + Inches(0.95) * i
        add_text(s, Inches(7.8), y + Inches(0.05), Inches(2.3), Inches(0.4),
                 lbl, size=12, bold=True, color=NAVY)
        for j, v in enumerate(vals):
            x = Inches(10.1) + Inches(1.0) * j
            add_round(s, x, y + Inches(0.02), Inches(0.85), Inches(0.55),
                      v, fill=c, color=WHITE, size=12, bold=True)
            add_text(s, x, y + Inches(0.6), Inches(0.85), Inches(0.25),
                     f"Y{j+1}", size=8, color=GREY, align=PP_ALIGN.CENTER)

    # 敏感性
    add_text(s, Inches(0.5), Inches(5.85), Inches(13), Inches(0.4),
             "敏感性（Y3 EBITDA 万元变化）：入驻 ±10% → ∓3,200/+1,400 · 租金 ±10% → ±2,100 · 政策兑现 −30% → −1,350",
             size=11, color=INK)

    # ============ 18. Phase 5 章节扉页 ============
    section_cover("PHASE 5",
                  "落地推进",
                  "Rollout · 12 Months · 4→22 Team",
                  ["任务 8 · 12 个月里程碑 + 团队扩编 + RACI 分工"])

    # ============ 19. 任务8 · 12 个月甘特 ============
    s = new_slide()
    add_chrome(s, prs, page_no=19, total=0, phase_label="Phase 5 · 任务 8",
               page_title="12 个月里程碑甘特图",
               subtitle="M1 团队就位 → M5 9 月发布会 → M12 入驻率 65%")
    months = [f"M{i+1}" for i in range(12)]
    # 表头
    chart_x = Inches(2.8); chart_y = Inches(1.2); chart_w = Inches(10.0)
    col_w = chart_w / 12.0
    for i, m in enumerate(months):
        x = chart_x + col_w * i
        add_rect(s, x, chart_y, col_w, Inches(0.35), fill=NAVY)
        add_text(s, x, chart_y, col_w, Inches(0.35), m,
                 size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 任务行
    tasks = [
        ("立项与法人主体",       0, 1, BLUE),
        ("团队组建（4→22）",     0, 12, NAVY),
        ("样板间 / 视觉",        1, 3, GOLD),
        ("户外测试区一期",       1, 5, GREEN),
        ("品牌官网 + CRM",       1, 3, BLUE),
        ("300 家库 + 中介签约",  2, 3, NAVY),
        ("首轮政府汇报",         1, 3, GOLD),
        ("政策包 v0 → v1",       3, 5, BLUE),
        ("链主 TOP5 接触",       1, 9, NAVY),
        ("链主首份 Term Sheet",  2, 2, GOLD),
        ("生态意向 30 家",       4, 4, BLUE),
        ("5 月品牌发布会",       4, 1, GREEN),
        ("9 月旗舰发布会",       8, 1, RED),
        ("「iDrive 静安 10 条」",8, 2, GOLD),
        ("入驻率 50%",           10, 1, GREEN),
        ("入驻率 65%",           11, 1, GREEN),
    ]
    row_h = Inches(0.28)
    for i, (lbl, start, dur, c) in enumerate(tasks):
        y = chart_y + Inches(0.4) + (row_h + Inches(0.03)) * i
        add_text(s, Inches(0.4), y, Inches(2.35), row_h, lbl,
                 size=9, color=INK, anchor=MSO_ANCHOR.MIDDLE)
        # 行底色
        add_rect(s, chart_x, y, chart_w, row_h, fill=CLOUD)
        # 任务条
        bx = chart_x + col_w * start
        bw = col_w * dur
        add_rect(s, bx + Emu(20000), y + Emu(20000), bw - Emu(40000), row_h - Emu(40000), fill=c)

    # ============ 20. 任务8 · 团队 + RACI ============
    s = new_slide()
    add_chrome(s, prs, page_no=20, total=0, phase_label="Phase 5 · 任务 8",
               page_title="4 人核心团队 → 22 人稳态 + RACI 分工",
               subtitle="启动 M1–3：4 人 / 扩张 M4–6：14 人 / 稳态 M7–12：22 人")
    # 左：4 人核心
    add_text(s, Inches(0.5), Inches(1.15), Inches(6), Inches(0.4),
             "4 人核心团队（启动期）", size=14, bold=True, color=NAVY)
    roles = [
        ("项目总监 GM", "全面统筹 / 链主谈判 / 重大客户", NAVY),
        ("GR 总监", "政府关系 / 政策包 / 合规", BLUE),
        ("招商总监", "渠道 / 漏斗 / 生态签约", GOLD),
        ("运营总监", "物业 / IT / 测试区 / 活动", GREEN),
    ]
    for i, (t, d, c) in enumerate(roles):
        y = Inches(1.55) + Inches(0.95) * i
        add_rect(s, Inches(0.5), y, Inches(6.0), Inches(0.85), fill=CLOUD, line=LINE)
        add_rect(s, Inches(0.5), y, Inches(0.18), Inches(0.85), fill=c)
        add_text(s, Inches(0.7), y + Inches(0.08), Inches(5.7), Inches(0.4),
                 t, size=14, bold=True, color=NAVY)
        add_text(s, Inches(0.7), y + Inches(0.45), Inches(5.7), Inches(0.4),
                 d, size=11, color=INK)

    add_text(s, Inches(0.5), Inches(5.6), Inches(6), Inches(0.4),
             "扩编节奏：4 → 14 → 22 人", size=12, bold=True, color=NAVY)
    add_text(s, Inches(0.5), Inches(5.95), Inches(6), Inches(0.6),
             "招商 1→5→8 · GR 1→2→2 · 运营 1→3→5 · 市场 0→1→3 · 法务/财务/HR 0→3→3",
             size=10, color=INK)

    # 右：RACI 表
    add_text(s, Inches(7.0), Inches(1.15), Inches(6), Inches(0.4),
             "RACI 关键事项分工", size=14, bold=True, color=NAVY)
    header = ["关键事项", "项目总监", "GR", "招商", "运营", "董事长", "法务"]
    rows = [
        ["链主谈判", "R", "C", "C", "I", "A", "C"],
        ["政府汇报 / 政策包", "C", "R", "I", "I", "A", "R"],
        ["中介渠道 / 漏斗例会", "A", "I", "R", "I", "I", "C"],
        ["9 月发布会 / 测试区", "A", "C", "C", "R", "C", "I"],
        ["合同 / 对赌", "A", "C", "C", "I", "A", "R"],
        ["财务模型 / 预算", "A", "I", "C", "C", "A", "I"],
        ["数据合规", "C", "C", "C", "C", "A", "R"],
    ]
    add_table(s, Inches(7.0), Inches(1.55), Inches(6.0), Inches(4.0),
              header, rows,
              col_widths=[Inches(2.0), Inches(0.7), Inches(0.6), Inches(0.7),
                          Inches(0.7), Inches(0.8), Inches(0.6)],
              header_size=10, body_size=10)
    add_text(s, Inches(7.0), Inches(5.65), Inches(6.0), Inches(0.5),
             "R = Responsible · A = Accountable · C = Consulted · I = Informed",
             size=10, italic=True, color=GREY)

    # ============ 21. 投决建议（Call to Action） ============
    s = new_slide()
    add_chrome(s, prs, page_no=21, total=0, phase_label="07 · 投决建议",
               page_title="投决建议 · 里程碑承诺",
               subtitle="资源到位的前提下，T+30/90/180 三个节点的硬承诺")

    # 三个节点大卡
    nodes2 = [
        ("T + 30 天", "团队就位 + 政府首轮汇报",
         ["4 人核心入职", "区投促办 + 副区长 + 市经信委首轮汇报", "5 家中介签约"], BLUE),
        ("T + 90 天", "链主首份 Term Sheet",
         ["TOP5 全部进入会面阶段", "≥ 1 家链主签 Term Sheet", "300 家库 + CRM 上线"], GOLD),
        ("T + 180 天", "9 月发布会 + 政策包定稿",
         ["发布会 200 人到场 + 5 家签约", "「iDrive 静安 10 条」发布", "签约 ≥ 30 家 / 入驻率 ≥ 30%"], GREEN),
    ]
    for i, (t, sub, bullets, c) in enumerate(nodes2):
        x = Inches(0.5) + Inches(4.30) * i
        add_rect(s, x, Inches(1.2), Inches(4.10), Inches(4.5), fill=CLOUD, line=LINE)
        add_rect(s, x, Inches(1.2), Inches(4.10), Inches(0.95), fill=c)
        add_text(s, x, Inches(1.25), Inches(4.10), Inches(0.5), t,
                 size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(s, x, Inches(1.70), Inches(4.10), Inches(0.4), sub,
                 size=12, color=WHITE, align=PP_ALIGN.CENTER)
        for j, b in enumerate(bullets):
            y = Inches(2.35) + Inches(0.65) * j
            add_round(s, x + Inches(0.20), y + Inches(0.15), Inches(0.25), Inches(0.25),
                      str(j + 1), fill=c, color=WHITE, size=10, bold=True)
            add_text(s, x + Inches(0.55), y + Inches(0.05), Inches(3.40), Inches(0.55),
                     b, size=12, color=INK)

    # 底部 CTA
    add_rect(s, Inches(0.5), Inches(5.95), Inches(12.3), Inches(1.0), fill=NAVY)
    add_text(s, Inches(0.7), Inches(6.0), Inches(12.0), Inches(0.45),
             "请董事会审议", size=14, bold=True, color=GOLD)
    add_text(s, Inches(0.7), Inches(6.40), Inches(12.0), Inches(0.55),
             "① 启动预算（Y0–Y1）  ② 4 人核心团队招聘授权  ③ 静安区一企一策政府专班 ④ 9 月发布会预算  ⑤ 链主谈判授权区间",
             size=12, color=WHITE)

    # ============ 22. 致谢 ============
    s = new_slide()
    sw, sh = slide_size_emu(prs)
    add_rect(s, 0, 0, sw, sh, fill=NAVY)
    add_rect(s, 0, Inches(3.6), sw, Emu(40000), fill=GOLD)
    add_text(s, Inches(0.8), Inches(2.3), Inches(11), Inches(1.4),
             "GS · iDrive Hub", size=44, bold=True, color=GOLD,
             align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.8), Inches(3.8), Inches(11), Inches(0.7),
             "让中心城区跑通智能驾驶最后一公里",
             size=22, color=WHITE, italic=True, align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.8), Inches(5.0), Inches(11), Inches(0.5),
             "Q & A · 谢 谢", size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.8), Inches(6.4), Inches(11), Inches(0.4),
             "冠松集团 · GS iDrive Hub 项目组",
             size=12, color=CLOUD, align=PP_ALIGN.CENTER)

    # --------------------------------------------------------------------
    # 回填总页数
    # --------------------------------------------------------------------
    total = len(SLIDES)
    for idx, sl in enumerate(SLIDES, start=1):
        for shape in sl.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.text and run.text.strip().endswith(" / 0"):
                        run.text = run.text.replace(" / 0", f" / {total}")

    # 输出
    out = Path(__file__).resolve().parent.parent / "docs" / "deck" / \
        "GS-iDrive-Hub-招商方案.pptx"
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    print(f"✓ Deck written: {out}  ({total} slides)")
    return out


if __name__ == "__main__":
    build()
