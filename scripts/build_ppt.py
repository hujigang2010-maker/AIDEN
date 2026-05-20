"""Generate the Greentown China 100,000 RMB sponsorship PPT deck."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from copy import deepcopy
from lxml import etree


GREEN = RGBColor(0x00, 0x6B, 0x3F)
DARK_GREEN = RGBColor(0x00, 0x4A, 0x2C)
LIGHT_GREEN = RGBColor(0xD8, 0xE8, 0xDF)
GOLD = RGBColor(0xC8, 0xA2, 0x5B)
DARK = RGBColor(0x1F, 0x1F, 0x1F)
GREY = RGBColor(0x66, 0x66, 0x66)
LIGHT_GREY = RGBColor(0xF2, 0xF2, 0xF2)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def add_bg(slide, color=WHITE):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    bg.shadow.inherit = False
    return bg


def add_rect(slide, x, y, w, h, fill, line=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
    shp.shadow.inherit = False
    return shp


def add_text(slide, x, y, w, h, text, size=14, bold=False, color=DARK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font="微软雅黑"):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.word_wrap = True
    tf.vertical_anchor = anchor

    if isinstance(text, str):
        lines = text.split("\n")
    else:
        lines = text

    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        r.font.name = font
        rPr = r._r.get_or_add_rPr()
        ea = etree.SubElement(rPr, qn("a:ea"))
        ea.set("typeface", font)
        cs = etree.SubElement(rPr, qn("a:cs"))
        cs.set("typeface", font)
    return tb


def add_title_bar(slide, title, subtitle=None):
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.85), GREEN)
    add_rect(slide, 0, Inches(0.85), SLIDE_W, Inches(0.05), GOLD)
    add_text(slide, Inches(0.5), Inches(0.12), Inches(10), Inches(0.6),
             title, size=22, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_text(slide, Inches(0.5), Inches(0.95), Inches(12), Inches(0.4),
                 subtitle, size=11, color=GREY)


def add_footer(slide, page_no, total=12):
    add_text(slide, Inches(0.5), Inches(7.05), Inches(8), Inches(0.3),
             "绿城中国 · 绿城·潮鸣外滩  |  2026 AI 商业化峰会 · 晚宴冠名战略合作伙伴 · 10 万元赞助专项设计方案",
             size=9, color=GREY)
    add_text(slide, Inches(11.5), Inches(7.05), Inches(1.4), Inches(0.3),
             f"{page_no:02d} / {total:02d}", size=9, color=GREY, align=PP_ALIGN.RIGHT)


def add_table(slide, x, y, w, h, headers, rows, col_widths_ratio=None,
              header_fill=GREEN, header_color=WHITE, body_size=10, header_size=11):
    n_cols = len(headers)
    n_rows = len(rows) + 1
    tbl_shape = slide.shapes.add_table(n_rows, n_cols, x, y, w, h)
    tbl = tbl_shape.table

    if col_widths_ratio:
        total = sum(col_widths_ratio)
        for i, r in enumerate(col_widths_ratio):
            tbl.columns[i].width = int(w * r / total)

    for c, hdr in enumerate(headers):
        cell = tbl.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_fill
        cell.margin_left = Inches(0.08)
        cell.margin_right = Inches(0.08)
        cell.margin_top = Inches(0.04)
        cell.margin_bottom = Inches(0.04)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.word_wrap = True
        tf.text = ""
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = hdr
        run.font.bold = True
        run.font.size = Pt(header_size)
        run.font.color.rgb = header_color
        run.font.name = "微软雅黑"
        rPr = run._r.get_or_add_rPr()
        ea = etree.SubElement(rPr, qn("a:ea"))
        ea.set("typeface", "微软雅黑")

    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            cell = tbl.cell(r_idx + 1, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if r_idx % 2 == 0 else LIGHT_GREY
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            cell.margin_top = Inches(0.04)
            cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            tf.text = ""
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = str(val)
            run.font.size = Pt(body_size)
            run.font.color.rgb = DARK
            run.font.name = "微软雅黑"
            rPr = run._r.get_or_add_rPr()
            ea = etree.SubElement(rPr, qn("a:ea"))
            ea.set("typeface", "微软雅黑")
    return tbl


def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    blank = prs.slide_layouts[6]

    total = 12

    # --- Slide 1: Cover ---
    s = prs.slides.add_slide(blank)
    add_bg(s, DARK_GREEN)
    add_rect(s, 0, Inches(3.0), SLIDE_W, Inches(0.04), GOLD)
    add_text(s, Inches(0.6), Inches(0.7), Inches(8), Inches(0.5),
             "GREENTOWN CHINA · 潮鸣外滩", size=14, color=GOLD, bold=True)
    add_text(s, Inches(0.6), Inches(1.4), Inches(12), Inches(1.4),
             "晚宴冠名战略合作伙伴", size=44, bold=True, color=WHITE)
    add_text(s, Inches(0.6), Inches(2.3), Inches(12), Inches(1.2),
             "绿城中国 | 绿城·潮鸣外滩", size=34, bold=True, color=WHITE)
    add_text(s, Inches(0.6), Inches(3.3), Inches(12), Inches(0.6),
             "10 万元赞助专项设计方案 · Sponsorship Plan", size=18, color=GOLD)
    add_text(s, Inches(0.6), Inches(4.2), Inches(12), Inches(1.6),
             [
                 "重构与突围 · 2026 人工智能商业化落地与硬核投资破局峰会",
                 "主办：北京大学经济学院上海校友会 · 复旦大学住房政策研究中心",
                 "时间：2026 年 5 月 22 日   地点：上海·北外滩·一滴水   规模：500+ 高净值嘉宾",
             ],
             size=14, color=WHITE)
    add_text(s, Inches(0.6), Inches(6.5), Inches(12), Inches(0.4),
             "提报日期：2026 年 5 月", size=11, color=GOLD)

    # --- Slide 2: 合作背景 ---
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title_bar(s, "01 · 合作背景与目标", "Cooperation Background & Goals")
    add_text(s, Inches(0.5), Inches(1.4), Inches(12.3), Inches(0.6),
             "在「重构与突围 · 2026 AI 商业化峰会」500 人高净值客群中，完成绿城·潮鸣外滩的全链路品牌渗透。",
             size=14, bold=True, color=DARK)

    add_rect(s, Inches(0.5), Inches(2.1), Inches(3.0), Inches(4.5), LIGHT_GREEN)
    add_text(s, Inches(0.65), Inches(2.25), Inches(2.7), Inches(0.5),
             "客群匹配", size=14, bold=True, color=GREEN)
    add_text(s, Inches(0.65), Inches(2.75), Inches(2.7), Inches(3.7),
             [
                 "· 25% 上市公司 / 独角兽高管",
                 "· 20% 一二级市场基金合伙人",
                 "· 25% AI / 算力 / 大模型创始团队",
                 "· 20% 北大、复旦双校核心校友",
                 "· 10% 政府、媒体合作伙伴",
                 "",
                 "500 + 高净值嘉宾，与潮鸣",
                 "外滩 285/310 户型核心",
                 "客群高度吻合。",
             ],
             size=11.5, color=DARK)

    add_rect(s, Inches(3.7), Inches(2.1), Inches(3.0), Inches(4.5), LIGHT_GREEN)
    add_text(s, Inches(3.85), Inches(2.25), Inches(2.7), Inches(0.5),
             "曝光场景", size=14, bold=True, color=GREEN)
    add_text(s, Inches(3.85), Inches(2.75), Inches(2.7), Inches(3.7),
             [
                 "· 主会场 8h 大屏轮播",
                 "· 主背景板钻石级 logo",
                 "· 议程手册 / 白皮书署名",
                 "· 500 份手拎袋 + 折页",
                 "· 晚宴 KV / 桌卡 / 菜单冠名",
                 "· 川总专场 15min 宣讲",
                 "· 主持人多轮口播",
                 "· 媒体通稿 + 九宫格",
             ],
             size=11.5, color=DARK)

    add_rect(s, Inches(6.9), Inches(2.1), Inches(3.0), Inches(4.5), LIGHT_GREEN)
    add_text(s, Inches(7.05), Inches(2.25), Inches(2.7), Inches(0.5),
             "项目到访", size=14, bold=True, color=GREEN)
    add_text(s, Inches(7.05), Inches(2.75), Inches(2.7), Inches(3.7),
             [
                 "· 论坛结束后主持口播引导",
                 "· 华为尊界接驳 8–10 辆",
                 "· 考斯特补位",
                 "· 项目售楼处沙盘 + 样板间",
                 "· 案场尊享接待",
                 "· 现场销售 1V1 跟进",
                 "· 单次到访体验 ≥ 60min",
             ],
             size=11.5, color=DARK)

    add_rect(s, Inches(10.1), Inches(2.1), Inches(3.0), Inches(4.5), LIGHT_GREEN)
    add_text(s, Inches(10.25), Inches(2.25), Inches(2.7), Inches(0.5),
             "长效沉淀", size=14, bold=True, color=GREEN)
    add_text(s, Inches(10.25), Inches(2.75), Inches(2.7), Inches(3.7),
             [
                 "· 双校长三角校友产业联盟",
                 "  战略合作伙伴永久入册",
                 "· 大会回顾视频片头鸣谢",
                 "· 白皮书扉页联合署名",
                 "· 后续活动优先合作权",
                 "· 嘉宾意向客户数据回传",
             ],
             size=11.5, color=DARK)

    add_footer(s, 2, total)

    # --- Slide 3: 10万定制权益总览 ---
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title_bar(s, "02 · 10 万元晚宴冠名 · 定制权益总览",
                  "在钻石赞助 5 万基础上叠加「晚宴冠名 + 项目专场宣讲 + 尊界接驳到访」三大独家权益")
    add_table(
        s, Inches(0.4), Inches(1.5), Inches(12.5), Inches(5.2),
        headers=["模块", "核心权益", "形式 / 物料", "责任方"],
        rows=[
            ["A 晚宴冠名", "「晚宴冠名战略合作伙伴」全程统称权", "晚宴 KV、席卡、桌卡、菜单 logo 植入；主持人开场+结束口播", "主办 + 活动公司"],
            ["B 主会场植入", "主背景板 logo + 议程手册整版 + 宣传片轮播 + 手拎袋折页", "钻石级 logo / 议程整版 / 500 份手拎袋 / 户型折页", "主办 + 印刷单位"],
            ["C 现场展位", "电梯口 + 展场内 2 处专属品牌展位，各派销售 2 人", "易拉宝、户型 KT 板、纸质物料、品牌桌台轻包装", "绿城 + 活动公司"],
            ["D 项目专场宣讲", "晚宴开场前 15 分钟，川总「潮鸣外滩」专场 PPT 演讲", "PPT、话筒、音响、LED 大屏；主持人引荐", "绿城 + 主办"],
            ["E 接驳到访", "华为尊界 8–10 辆 + 考斯特补位，定向引导到案场", "尊界车队 + 引导员 + 案场接待", "绿城"],
            ["F 论坛后口播", "圆桌结束后主持人口播赞助方并引导前往项目", "定制口播稿 + 现场引导员", "主办"],
            ["G 媒体宣发", "媒体通稿、九宫格、白皮书署名、回顾视频片头鸣谢", "通稿、朋友圈九宫格、回顾短片", "主办媒体组"],
            ["H 长效圈层", "双校长三角校友产业联盟战略合作伙伴永久入册", "联盟通讯录 + 牌匾交接", "主办"],
        ],
        col_widths_ratio=[1.4, 3.3, 4.2, 1.8],
        body_size=10.5,
    )
    add_footer(s, 3, total)

    # --- Slide 4: 与原赞助级别对比 ---
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title_bar(s, "03 · 权益对比矩阵", "与原峰会赞助级别的差异定位 · 10 万 = 钻石 + 晚宴冠名 + 项目专场")
    add_table(
        s, Inches(0.4), Inches(1.5), Inches(12.5), Inches(5.2),
        headers=["权益项", "黄金 1 万", "铂金 3 万", "钻石 5 万", "10 万 晚宴冠名（本方案）"],
        rows=[
            ["大会主背景板 logo 等级", "黄金", "铂金", "钻石", "钻石 +「晚宴冠名」副标位"],
            ["议程手册广告", "1/4 版", "半版", "整版", "整版 + 项目折页夹页"],
            ["白皮书署名", "—", "1/4 版", "1/2 版", "扉页联合署名"],
            ["VIP 晚宴入场券", "1 张", "3 张", "5 张", "主桌 + 销售/接待 共 8 张"],
            ["圆桌席位", "—", "—", "限 1", "—（资源集中投入晚宴）"],
            ["晚宴冠名权", "—", "—", "—", "● 独家"],
            ["晚宴前项目专场 15min 宣讲", "—", "—", "—", "● 独家（川总）"],
            ["晚宴桌卡 / 菜单 logo", "—", "—", "—", "● 全场"],
            ["主持人口播", "—", "鸣谢", "鸣谢", "开场 + 结束 + 引导到访 3 段"],
            ["现场展位", "—", "—", "—", "2 处（电梯口 + 展场内）"],
            ["项目到访接驳", "—", "—", "—", "尊界 8–10 辆 + 考斯特"],
            ["双校友联盟入册", "—", "—", "永久", "永久 · 战略级"],
        ],
        col_widths_ratio=[3.5, 1.4, 1.4, 1.6, 3.6],
        body_size=10,
    )
    add_footer(s, 4, total)

    # --- Slide 5: 晚宴植入位 ---
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title_bar(s, "04 · 晚宴视觉植入位 & 责任分工",
                  "席卡 + 晚宴 KV 由主办加上；桌卡 / 菜单由活动公司印刷植入")
    add_table(
        s, Inches(0.4), Inches(1.5), Inches(12.5), Inches(5.2),
        headers=["序号", "物料", "植入方式", "尺寸建议", "责任方"],
        rows=[
            ["1", "晚宴 KV（背板/LED）", "「晚宴冠名战略合作伙伴 绿城中国 | 绿城·潮鸣外滩」横向锁屏", "请场地方提供准确像素比；首选 1920×1080 或 LED 条屏 6144×1080", "主办（绿城提供素材）"],
            ["2", "席卡", "卡面右下角 logo + 一行品牌主张", "90×55mm 双面，竖式或横式以场地最终方案为准", "主办"],
            ["3", "桌卡（桌号牌）", "桌号 + 「绿城·潮鸣外滩 · 晚宴冠名」字样", "A5 双面折立 148×210mm，或亚克力 200×150mm", "活动公司 / 印刷"],
            ["4", "菜单", "封面整版 logo + 内页页脚 logo", "对开 A4 210×285mm 或单页 285×210mm", "活动公司 / 印刷"],
            ["5", "晚宴 LED 轮播", "30s 项目宣传片 + 5 张静帧（5s/张）循环", "16:9 1920×1080，H.264 .mp4 ≤ 200MB", "绿城出片，主办上屏"],
            ["6", "主持人口播", "开宴鸣谢 / 川总宣讲引荐 / 散场再次鸣谢 共 3 段", "每段约 30 秒", "主办 + 绿城共写稿"],
            ["7", "桌花卡 / 餐巾纸条（可选）", "项目主 slogan 烫金字条", "150×30mm 条形", "活动公司"],
        ],
        col_widths_ratio=[0.4, 1.8, 3.6, 4.0, 1.8],
        body_size=10,
    )
    add_footer(s, 5, total)

    # --- Slide 6: 物料尺寸建议 ---
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title_bar(s, "05 · 视频 & PPT & 印刷物料尺寸建议",
                  "待场地方最终像素比确认后微调；主会场已确认 16:9 无问题")
    add_table(
        s, Inches(0.4), Inches(1.5), Inches(12.5), Inches(5.2),
        headers=["素材类型", "推荐比例", "推荐分辨率", "时长 / 帧率", "备注"],
        rows=[
            ["主会场宣传片（轮播）", "16:9", "1920×1080（4K 备 3840×2160）", "≤ 60s，25/30 fps", "已确认主会场 16:9"],
            ["晚宴 LED 宣传片", "16:9（条屏另出）", "1920×1080 或 6144×1080", "30–60s 循环", "请场地方提供像素比"],
            ["川总专场宣讲 PPT", "16:9", "1920×1080 (40×22.5cm)", "≤ 25 张，控制 15min", "封面/封底/沙盘 3D 页预留"],
            ["晚宴 KV 主视觉", "依场地像素比", "建议先 1920×1080 主稿，再延展", "静帧 PNG/JPG + AI 源", "RGB 色域"],
            ["项目 logo", "矢量", "AI / EPS / SVG（无白底）", "—", "另出 PNG 透明底 4 套尺寸"],
            ["户型折页（285/310）", "印刷品", "297×210mm 三折页，出血 3mm", "CMYK 300 dpi", "印刷单位统一排版"],
            ["项目易拉宝", "印刷品", "800×2000mm，出血 5mm", "CMYK 150 dpi", "电梯口 + 展场内各 1 组"],
            ["朋友圈九宫格", "1:1", "1080×1080 PNG × 9", "—", "媒体组统一发布"],
        ],
        col_widths_ratio=[2.4, 2.0, 3.2, 1.8, 2.2],
        body_size=10.5,
    )
    add_footer(s, 6, total)

    # --- Slide 7: 现场展位 ---
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title_bar(s, "06 · 现场展位与动线设计",
                  "保底易拉宝 + 项目纸质物料；设计时间允许则升级背景墙")

    # Left card
    add_rect(s, Inches(0.5), Inches(1.5), Inches(6.1), Inches(5.0), LIGHT_GREEN)
    add_rect(s, Inches(0.5), Inches(1.5), Inches(6.1), Inches(0.55), GREEN)
    add_text(s, Inches(0.65), Inches(1.55), Inches(5.8), Inches(0.45),
             "展位 1 · 电梯口（嘉宾抵达第一触点）", size=14, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(0.7), Inches(2.2), Inches(5.7), Inches(4.2),
             [
                 "· 品牌桌台 1 张（轻包装）",
                 "· 易拉宝 ×2（项目主形象 + 户型亮点）",
                 "· 户型折页/纸质物料展示架",
                 "· 285/310 户型 KT 板",
                 "· 销售 2 人（统一品牌服装、胸卡）",
                 "· 二维码 / 加微登记表",
                 "",
                 "目标：第一时间触达入场嘉宾，",
                 "完成项目认知建立 & 意向客户登记。",
             ],
             size=12, color=DARK)

    # Right card
    add_rect(s, Inches(6.8), Inches(1.5), Inches(6.1), Inches(5.0), LIGHT_GREEN)
    add_rect(s, Inches(6.8), Inches(1.5), Inches(6.1), Inches(0.55), GREEN)
    add_text(s, Inches(6.95), Inches(1.55), Inches(5.8), Inches(0.45),
             "展位 2 · 展场内（主会场前厅近茶歇区）", size=14, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(7.0), Inches(2.2), Inches(5.7), Inches(4.2),
             [
                 "· 洽谈圆桌 1 组（轻包装）",
                 "· 易拉宝 ×2 + 285/310 户型 KT 板",
                 "· 茶歇/休息时段重点拦截",
                 "· 销售 2 人现场 1V1 沟通",
                 "· 主会场屏幕同步轮播宣传片",
                 "· 现场预约「项目到访」",
                 "",
                 "目标：承接会议中场休息流量，",
                 "完成意向客户深度沟通 & 预约到访。",
             ],
             size=12, color=DARK)

    add_text(s, Inches(0.5), Inches(6.6), Inches(12.5), Inches(0.4),
             "现场盘点：保底交付 4 个易拉宝 + 2 张品牌桌台轻包装 + 项目纸质物料；时间允许升级背景挂幔/KT 板背景墙。",
             size=11, color=GREY)
    add_footer(s, 7, total)

    # --- Slide 8: 项目参观时间冲突 ---
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title_bar(s, "07 · 项目参观与晚宴衔接 · 时间冲突解决方案",
                  "测算往返车程 + 案场参观 ≥ 60 分钟，与现议程 45 分钟以上冲突")
    add_table(
        s, Inches(0.4), Inches(1.5), Inches(12.5), Inches(5.0),
        headers=["方案", "做法", "优势", "风险 / 协调点"],
        rows=[
            ["方案 A · 推荐", "「会后专场」：18:10–20:30 晚宴照常进行；20:30 散场后由尊界车队定向接送意向嘉宾至项目案场夜场参观（灯光秀+样板间），再分送回酒店", "不打断议程；项目夜场氛围有记忆点；尊界利用率高；意向客户更聚焦", "需提前确认意向名单；案场夜间接待人员到位"],
            ["方案 B", "议程内插入「项目专场参观」：17:55 颁奖结束后拉长为 90 分钟参观窗口，晚宴延后至 19:30 开始", "参观时间充裕", "影响主办原议程及晚宴餐厅档期；需主办协调"],
            ["方案 C", "「云参观 + 邀约到访」：晚宴期间用 3D VR 沙盘 + 川总 15min 现场宣讲完成线上参观，现场登记次日专车到访", "完全不影响议程；可批量收集意向", "现场参观感弱；接待成本后置"],
        ],
        col_widths_ratio=[1.6, 4.5, 3.4, 3.0],
        body_size=10.5,
    )
    add_text(s, Inches(0.4), Inches(6.55), Inches(12.5), Inches(0.4),
             "建议首选方案 A；最终时间窗以主办、活动公司、绿城三方 5 月 19 日前确认为准，绿城据此安排尊界与案场接待。",
             size=11, color=GREEN, bold=True)
    add_footer(s, 8, total)

    # --- Slide 9: 晚宴流程脚本 ---
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title_bar(s, "08 · 晚宴现场流程脚本", "Dinner Run-of-Show · 建议版")
    add_table(
        s, Inches(0.4), Inches(1.4), Inches(12.5), Inches(5.4),
        headers=["时间", "环节", "内容 / 物料", "音视频需求"],
        rows=[
            ["18:00–18:10", "嘉宾入场", "晚宴 KV 点亮 · 席卡引位 · 项目宣传片轮播（低音乐）", "LED 16:9；BGM"],
            ["18:10–18:15", "主持人开场 + 冠名鸣谢", "口播「晚宴冠名战略合作伙伴 绿城中国 | 绿城·潮鸣外滩」", "话筒×2"],
            ["18:15–18:30", "川总·绿城·潮鸣外滩专场宣讲", "项目 PPT 15min + Q&A 5min", "1080P 16:9 PPT；话筒；切换器"],
            ["18:30–18:35", "联合祝酒辞", "主办 + 绿城联合致酒辞", "话筒×2"],
            ["18:35–19:30", "正餐第一轮", "项目宣传片循环；菜单/桌卡 logo 植入", "宣传片循环；BGM"],
            ["19:30–20:00", "圈层社交 + 校友联盟介绍", "联盟介绍片；绿城战略合作伙伴永久入册牌匾交接", "LED；摄影"],
            ["20:00–20:25", "正餐第二轮 / 自由交流", "销售场内 1V1 沟通", "BGM"],
            ["20:25–20:30", "结束鸣谢 & 引导到访", "再次口播；引导至尊界车队（方案 A 启动）", "话筒×1"],
        ],
        col_widths_ratio=[1.4, 2.4, 5.6, 2.0],
        body_size=10.5,
    )
    add_footer(s, 9, total)

    # --- Slide 10: 媒体宣发 ---
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title_bar(s, "09 · 媒体宣发 & 长效价值", "Media Plan & Long-term Value")

    items = [
        ("媒体通稿", "标题或副标题级出现「绿城·潮鸣外滩」+ 晚宴冠名身份；主流财经/地产头部媒体 ≥ 5 家"),
        ("朋友圈九宫格", "9 张视觉，至少 3 张含项目 logo / KV；文案口径与绿城品牌部对齐"),
        ("白皮书署名", "扉页联合署名「晚宴冠名战略合作伙伴 绿城中国 · 绿城·潮鸣外滩」"),
        ("回顾视频", "片头 3s 鸣谢 + 片尾 logo 墙重点位 + 川总宣讲/参观画面"),
        ("校友联盟", "双校长三角校友产业联盟战略合作伙伴永久入册"),
        ("数据回传", "现场扫码/意向客户名单回传绿城销售线"),
    ]
    for i, (title, desc) in enumerate(items):
        col = i % 3
        row = i // 3
        x = Inches(0.4 + col * 4.25)
        y = Inches(1.6 + row * 2.55)
        add_rect(s, x, y, Inches(4.0), Inches(2.35), LIGHT_GREEN)
        add_rect(s, x, y, Inches(4.0), Inches(0.5), GREEN)
        add_text(s, x + Inches(0.15), y + Inches(0.05), Inches(3.7), Inches(0.4),
                 title, size=13, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + Inches(0.2), y + Inches(0.6), Inches(3.6), Inches(1.7),
                 desc, size=11.5, color=DARK)
    add_footer(s, 10, total)

    # --- Slide 11: 时间节点 ---
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title_bar(s, "10 · 时间节点 Timeline", "由于设计制作时间仅 1 天，关键截止节点务必前置")
    add_table(
        s, Inches(0.4), Inches(1.5), Inches(12.5), Inches(5.0),
        headers=["日期", "节点", "绿城方", "主办 / 活动公司"],
        rows=[
            ["5/18", "确认合作框架及金额", "盖章合同回传", "出具盖章合同"],
            ["5/19", "物料源文件交付", "logo（AI/EPS/PNG）· 项目宣传片 · 川总 PPT · 户型折页源文件 · 品牌简介", "晚宴 KV / 桌卡 / 菜单 / 席卡设计稿初稿"],
            ["5/20", "设计稿确认", "确认一轮，提出修改意见", "完成印刷物料终稿"],
            ["5/21", "印刷下单 & 物流", "—", "桌卡、菜单、易拉宝、户型折页全部下印"],
            ["5/22 上午", "现场布展", "销售 4 人到场对接展位包装", "活动公司完成展位、KV、LED 调试"],
            ["5/22 全天", "大会执行", "执行 + 陪同 + 案场对接", "执行 + 主持口播 + 宣传片轮播"],
            ["5/22 20:30+", "项目参观（方案 A）", "尊界 8–10 辆 + 考斯特补位 · 案场接待", "现场引导员 · 名单对接"],
            ["5/23–5/29", "宣发 & 回执", "同步绿城自媒体矩阵传播", "通稿/九宫格发布；7 日内出回执"],
        ],
        col_widths_ratio=[1.2, 1.8, 4.6, 4.0],
        body_size=10.5,
    )
    add_footer(s, 11, total)

    # --- Slide 12: 投入产出 & 对接人 ---
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title_bar(s, "11 · 投入产出 & 对接矩阵", "Investment & Contact")

    add_rect(s, Inches(0.4), Inches(1.5), Inches(6.1), Inches(5.0), LIGHT_GREEN)
    add_rect(s, Inches(0.4), Inches(1.5), Inches(6.1), Inches(0.55), GREEN)
    add_text(s, Inches(0.55), Inches(1.55), Inches(5.8), Inches(0.45),
             "投入产出概要", size=14, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(0.6), Inches(2.2), Inches(5.8), Inches(4.2),
             [
                 "赞助金额：人民币 100,000 元",
                 "（壹拾万元整）",
                 "",
                 "覆盖人群：500+ 高净值嘉宾",
                 "  · 直接触达 ≥ 500 人",
                 "  · 媒体二次曝光 ≥ 50w+",
                 "  · 项目到访意向客户 30–60 组",
                 "",
                 "付款方式：合同签订后 5 个工作日",
                 "一次性付款至主办指定账户",
                 "",
                 "发票内容：会议服务费 / 赞助费",
                 "（增值税普通发票或专用发票）",
             ],
             size=12, color=DARK)

    add_rect(s, Inches(6.8), Inches(1.5), Inches(6.1), Inches(5.0), LIGHT_GREEN)
    add_rect(s, Inches(6.8), Inches(1.5), Inches(6.1), Inches(0.55), GREEN)
    add_text(s, Inches(6.95), Inches(1.55), Inches(5.8), Inches(0.45),
             "对接矩阵", size=14, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(7.0), Inches(2.2), Inches(5.8), Inches(4.2),
             [
                 "绿城方·总对接：Aiden（方案/付款/合同）",
                 "绿城方·项目宣讲：川总团队",
                 "绿城方·现场销售：销售负责人（4 人）",
                 "",
                 "主办·组委会招商：13262607888",
                 "主办·现场执行：wangsheng（KV/印刷/流程）",
                 "活动公司·印刷：印刷负责人（桌卡/菜单）",
                 "",
                 "关键确认事项：",
                 "  ① 晚宴场地像素比 / 屏幕规格",
                 "  ② 主办对接的印刷单位 logo 植入",
                 "  ③ 项目参观方案（推荐方案 A）",
             ],
             size=12, color=DARK)

    add_footer(s, 12, total)

    out = "/workspace/deliverables/绿城中国-潮鸣外滩-10万元晚宴冠名赞助专项设计方案.pptx"
    prs.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    build()
