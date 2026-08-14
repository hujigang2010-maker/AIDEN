# -*- coding: utf-8 -*-
"""
生成《2026年具身智能机器人白皮书》。

署名：复旦大学住房政策研究中心（CHPS）
运行：python3 scripts/build_whitepaper.py
输出：dist/复旦大学住房政策研究中心_2026年具身智能机器人白皮书.docx
"""

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
DIST = ROOT / "dist"
DIST.mkdir(parents=True, exist_ok=True)
OUT = DIST / "复旦大学住房政策研究中心_2026年具身智能机器人白皮书.docx"

NAVY = RGBColor(0x0A, 0x2F, 0x6B)
DARK = RGBColor(0x1B, 0x2A, 0x3A)
GRAY = RGBColor(0x5A, 0x64, 0x6E)
ACCENT = RGBColor(0x8B, 0x1E, 0x3F)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TABLE_HEADER_BG = "0A2F6B"
TABLE_ALT_BG = "F4F6F9"


def set_run_font(run, size=12, bold=False, color=DARK, font="宋体", east_asia=None, italic=False):
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.name = font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:ascii"), "Times New Roman" if font in ("宋体", "楷体", "黑体") else font)
    rFonts.set(qn("w:hAnsi"), "Times New Roman" if font in ("宋体", "楷体", "黑体") else font)
    rFonts.set(qn("w:eastAsia"), east_asia or font)


def set_cell_shading(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_borders(cell, color="C5CDD6"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        tcBorders.append(el)
    tcPr.append(tcBorders)


def add_page_number(paragraph):
    run1 = paragraph.add_run("— ")
    set_run_font(run1, size=9, color=GRAY, font="宋体")
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run = paragraph.add_run()
    set_run_font(run, size=9, color=GRAY, font="Times New Roman")
    r = run._r
    r.append(fld_begin)
    r.append(instr)
    r.append(fld_sep)
    t = OxmlElement("w:t")
    t.text = "1"
    r.append(t)
    r.append(fld_end)
    run2 = paragraph.add_run(" —")
    set_run_font(run2, size=9, color=GRAY, font="宋体")


def setup_header_footer(doc):
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.6)
    section.bottom_margin = Cm(2.4)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.6)
    section.header_distance = Cm(1.2)
    section.footer_distance = Cm(1.2)
    section.different_first_page_header_footer = True

    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = hp.add_run("复旦大学住房政策研究中心  ·  2026年具身智能机器人白皮书")
    set_run_font(run, size=9, color=NAVY, font="宋体")

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_page_number(fp)


def add_cover_line(doc, text, size=12, bold=False, color=DARK, font="宋体", space_after=6, space_before=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color, font=font)
    return p


def add_title(doc, text, size=22):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(10)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=True, color=NAVY, font="黑体")
    return p


def add_heading_cn(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18 if level == 1 else 12)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    size = 16 if level == 1 else 13
    color = NAVY if level == 1 else ACCENT
    run = p.add_run(text)
    set_run_font(run, size=size, bold=True, color=color, font="黑体")
    return p


def add_body(doc, text, first_indent=True, bold=False, size=12):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    if first_indent:
        p.paragraph_format.first_line_indent = Cm(0.74)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=DARK, font="宋体")
    return p


def add_quote(doc, text, source=""):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.left_indent = Cm(0.74)
    p.paragraph_format.right_indent = Cm(0.5)
    run = p.add_run(text)
    set_run_font(run, size=11, italic=True, color=NAVY, font="楷体")
    if source:
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(10)
        p2.paragraph_format.right_indent = Cm(0.5)
        r2 = p2.add_run(source)
        set_run_font(r2, size=10, color=GRAY, font="楷体")


def add_bullets(doc, items):
    for item in items:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.left_indent = Cm(0.74)
        run = p.add_run("●  " + item)
        set_run_font(run, size=12, color=DARK, font="宋体")


def add_caption(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(10)
    run = p.add_run(text)
    set_run_font(run, size=10, color=GRAY, font="楷体")
    return p


def add_table(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        set_run_font(run, size=10, bold=True, color=WHITE, font="黑体")
        set_cell_shading(cell, TABLE_HEADER_BG)
        set_cell_borders(cell, "0A2F6B")
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx > 0 else WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(str(val))
            set_run_font(run, size=10, color=DARK, font="宋体")
            if r_idx % 2 == 1:
                set_cell_shading(cell, TABLE_ALT_BG)
            set_cell_borders(cell)
    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Cm(w)
    return table


def add_page_break(doc):
    doc.add_page_break()


def build_cover(doc):
    for _ in range(2):
        doc.add_paragraph()
    add_cover_line(doc, "复旦大学住房政策研究中心", size=16, bold=True, color=NAVY, font="黑体", space_after=4)
    add_cover_line(doc, "Center for Housing Policy Studies, Fudan University", size=11, color=GRAY, font="Times New Roman", space_after=18)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(16)
    run = p.add_run("CHPS 研 究 报 告  ·  2026 年 第 1 号")
    set_run_font(run, size=12, bold=True, color=ACCENT, font="黑体")

    add_cover_line(doc, "2026 年具身智能机器人白皮书", size=26, bold=True, color=NAVY, font="黑体", space_before=12, space_after=10)
    add_cover_line(
        doc,
        "从展会现场到居住空间\n产业跃迁、住房政策含义与 2030 展望",
        size=16,
        bold=True,
        color=DARK,
        font="楷体",
        space_after=18,
    )
    add_cover_line(
        doc,
        "Embodied Intelligent Robotics 2026\nFrom Exhibition Floors to Living Spaces: Industry Shift, Housing Policy Implications, and the 2030 Horizon",
        size=11,
        color=GRAY,
        font="Times New Roman",
        space_after=28,
    )

    add_cover_line(doc, "基于世界人工智能大会（WAIC 2026）及同期国内外展会公开报道的综合研究", size=11, color=DARK, font="宋体", space_after=8)
    add_cover_line(doc, "2026 年 8 月", size=14, bold=True, color=NAVY, font="黑体", space_before=24, space_after=6)
    add_cover_line(doc, "上海", size=12, color=DARK, font="宋体", space_after=6)

    add_page_break(doc)


def build_statement(doc):
    add_heading_cn(doc, "说  明", level=1)
    add_body(
        doc,
        "本白皮书由复旦大学住房政策研究中心（Center for Housing Policy Studies, CHPS）组织撰写。中心成立于 2009 年，挂靠复旦大学管理学院，坚持学术性、中立性与公益性原则，长期围绕住房保障、公共住房政策、住房市场调控与住房金融开展研究。本次研究将观察对象从“住房本身”延伸至正在进入住房、社区与康养空间的具身智能机器人，意在回答一个对住房政策越来越无法回避的问题：当机器人开始成为高频出现的“住户”和“服务者”，居住空间、社区配套、适老化改造与住房公共服务体系应当如何提前准备。",
    )
    add_body(
        doc,
        "白皮书所依据的事实材料，主要来自 2026 年 1 月至 8 月公开新闻报道、展会官方信息、政府部门通知与产业观察文章，重点覆盖 2026 年世界人工智能大会（WAIC，7 月 17 日至 20 日，上海）、国际消费电子展（CES 2026）、汉诺威工业博览会、北京亦庄人形机器人半程马拉松，以及即将于 8 月 19 日至 23 日举行的世界机器人大会等。文中企业产品、订单规模、融资数据等均转引自公开报道，供政策讨论与学术研究参考，不构成对任何企业或投资标的的背书。",
    )
    add_body(
        doc,
        "具身智能仍处早期。产业热度、展会密度与真实入户能力之间存在明显时滞。本中心选择在 2026 年这一“量产元年”发布白皮书，正是因为住房政策的调整周期往往长于技术迭代周期：若等到机器人大规模进入家庭再讨论户型、公区、物业与数据治理，公共部门将处于被动。展望 2030 年，不是为了追逐技术叙事，而是为住房制度预留足够的适应窗口。",
    )
    add_body(doc, "文责自负。欢迎学界、产业界与政策部门批评指正。", first_indent=True)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(12)
    run = p.add_run("复旦大学住房政策研究中心\n2026 年 8 月 14 日")
    set_run_font(run, size=12, color=DARK, font="楷体")
    add_page_break(doc)


def build_toc(doc):
    add_heading_cn(doc, "目  录", level=1)
    items = [
        "摘要：2026 年的五个判断与 2030 年的一条主线",
        "一、导论：住房政策为何必须正视具身智能",
        "二、2026 年产业坐标：从技术叙事转向产业叙事",
        "三、WAIC 2026 深度观察：五个结构性趋势",
        "四、全球与国内展会扫描：一年之内的现场证据",
        "五、场景落地路径：工厂先行、康养跟进、家庭后置",
        "六、住房、社区与城市：当机器人成为“住户”",
        "七、展望 2030：四个情景与关键变量",
        "八、政策建议",
        "附录一  2026 年重要展会、赛事与政策节点",
        "附录二  主要公开资料来源",
        "附录三  术语简释",
    ]
    for item in items:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.5
        run = p.add_run(item)
        set_run_font(run, size=12, color=DARK, font="宋体")
    add_page_break(doc)


def build_abstract(doc):
    add_heading_cn(doc, "摘要：2026 年的五个判断与 2030 年的一条主线", level=1)
    add_body(
        doc,
        "2026 年是具身智能从“实验室 Demo”转向“可交付作业”的转折之年。这一判断并非来自抽象推演，而是来自一整年密集展会与公开报道所呈现的现场证据。7 月 17 日至 20 日在上海举行的世界人工智能大会，以“智能伙伴 共创未来”为主题，展览面积首次突破 10 万平方米，1100 余家企业参展，4486 项展品中超 300 款全球首发。具身智能首次升格为与智算并列的两大核心赛道，独立成馆，超 200 家企业、208 款终端、逾 300 台真机同台。对照 2024 年大会机器人专区仅 25 款人形机器人、2025 年具身智能初为焦点，三年密度跃迁本身就是产业加速的注脚。",
    )
    add_body(doc, "本白皮书将现场观察收敛为五个判断：", first_indent=True)
    add_bullets(
        doc,
        [
            "评价尺度已经迁移。展台上武术与舞蹈大幅减少，拆垛、上料、装配、巡检、药房分拣成为主流；指标从自由度数、单动作成功率转向连续作业时长、异常恢复率与场景适配成本。",
            "产业链第一次被完整摊开。整机厂之外，力触觉传感器、灵巧手、数据采集与仿真评测、跨本体开发平台以独立身份登场，“卖铲子”的基础设施企业估值与关注度显著抬升。",
            "数据升格为独立产业层。高质量真实物理交互数据缺口巨大，多城万平方米级采集中心、无本体可穿戴采集、仿真引擎三条路径并行，头部团队目标从千万小时冲向亿小时级。",
            "落地顺序已经收敛：工业与仓储先行、商业服务与政务场景跟进、康养机构作为家庭之前的过渡、真正规模化入户仍需数年。智元联合创始人彭志辉公开预判，规模化入户至少还要五年。",
            "住房与社区不再是技术叙事的“远景附录”，而正在成为数据飞轮、适老化改造与公共服务供给的交汇点。机器人友好住宅、社区机器人配置率、居家数据治理，应纳入住房政策议程。",
        ],
    )
    add_body(
        doc,
        "展望 2030 年，本中心不以“全面普及人形保姆”为基准情景，而以“分层渗透”为基准：工业与物流场景基本完成规模化部署；康养机构与社区驿站形成可复制的人机协同服务包；家庭端以订阅服务、远程兜底和轻量化陪伴产品为主，全自主家务机器人仍集中于支付能力较强、户型条件较好的家庭。住房政策的任务，不是等待技术成熟后再被动适应，而是在 2026—2030 年窗口期内，把机器人可达性、社区服务设施、适老化智能改造、居住数据权利与住房公平，写成可执行的制度安排。",
    )
    add_body(
        doc,
        "关键词：具身智能；人形机器人；世界人工智能大会；住房政策；适老化改造；智慧社区；2030 展望",
        first_indent=False,
    )


def build_ch1(doc):
    add_heading_cn(doc, "一、导论：住房政策为何必须正视具身智能", level=1)
    add_heading_cn(doc, "（一）问题的提出", level=2)
    add_body(
        doc,
        "住房政策研究通常处理土地、金融、保障、市场调控与居住福利。机器人长期被归入工业政策或科技产业政策。2026 年的现场变化，使这一学科边界开始松动。展会上的机器人不再只是工厂里的机械臂，它们被要求折叠衣物、整理客厅、在养老驿站迎宾问询、在社区完成巡检与配送。家庭、社区与机构养老空间，正在从“技术终局的修辞”变成“可参观、可试用、可签约”的场景。",
    )
    add_body(
        doc,
        "住房是最复杂的非结构化环境之一。户型差异、家具摆放、地面高差、门宽、光线、宠物、儿童与老人的随机行为，使家庭成为具身智能最难、也最有价值的训练场。产业界已有企业联合家政平台进入上千户家庭采集数据；也有企业以每月数千元的租赁模式试水。对住房政策而言，这意味着：住宅不再只是人的容器，也可能成为机器人的工作场所、数据生产车间和公共服务的新接口。",
    )
    add_heading_cn(doc, "（二）研究视角与方法", level=2)
    add_body(
        doc,
        "本白皮书采取“展会民族志 + 公开报道综合 + 住房政策推演”的方法。展会是产业自我呈现的舞台，存在营销放大，但连续对比 2024、2025、2026 三年的展示内容、参展主体与评价话术，仍能识别结构性位移：从“能不能动”到“能不能稳定干活”，从整机崇拜到数据与零部件基础设施，从工厂专用到生活服务延展。住房政策推演则把这些技术事实翻译为户型规范、社区配套、保障房与老旧小区更新、物业服务与数据治理等可操作议题。",
    )
    add_heading_cn(doc, "（三）“十五五”与住房民生的交汇", level=2)
    add_body(
        doc,
        "2026 年是“十五五”开局之年。《中华人民共和国国民经济和社会发展第十五个五年规划纲要》将具身智能列为未来产业，要求统筹布局实训场，推进虚实融合训练，攻关本体与核心零部件，加速人形机器人等产品升级和应用落地；同时在“创造美好数智生活”一节中，明确拓展养老助残、智能家居、智慧社区等融合应用。住房政策研究不能把这两段文字读成互不相干的产业条款与民生条款——它们将在同一套住宅和社区里相遇。",
    )


def build_ch2(doc):
    add_heading_cn(doc, "二、2026 年产业坐标：从技术叙事转向产业叙事", level=1)
    add_heading_cn(doc, "（一）量产元年的含义", level=2)
    add_body(
        doc,
        "业界将 2026 年称为具身智能或人形机器人的“量产元年”“应用元年”。需要澄清的是：量产元年并不等于普及元年。它首先意味着：本体硬件开始收敛、供应链可以按周交付、垂直场景出现连续数月的试运行、千台级订单不再只是新闻标题。因时机器人公开称其灵巧手 2025 年出货量突破 1 万台；傲意科技董事副总裁张琦接受采访时表示，去年订单约“大几千台”，今年预计达数万台、上半年已破万台。灵巧手作为独立品类跑出整机厂体系，是量产叙事中最硬的证据之一。",
    )
    add_body(
        doc,
        "整机侧，智元远征系列被报道为可在工厂完成上料、装配、巡检的量产级全尺寸人形方案；乐聚在 WAIC 以 1:1 复刻产线展示拆垛与上料，称方案已在长三角多家工厂试运行 3 到 4 个月；星动纪元称物流分拣方案已与顺丰、中国邮政在全国十余个物流中心常态运营。这些案例的共同特征不是“表演更精彩”，而是“离开展馆后还能继续跑”。",
    )
    add_heading_cn(doc, "（二）资本热度与用户冷静并存", level=2)
    add_body(
        doc,
        "IT桔子数据显示，2026 年上半年国内具身智能赛道融资规模达 935 亿元，同比约增 5 倍，前二十家企业包揽约六成资金。与此同时，家庭用户的决策逻辑非常朴素：能不能干活、值不值这个价、能不能替代人工、安不安全。多家头部从业者承认，家庭端尚未到大规模买单阶段。资本用十年终局定价，消费者用当下体验打分，两套评价体系并行，是颠覆性技术产业化初期的常态，也是住房政策需要保持清醒的原因——不能把展会热度直接写成“机器人即将标配入户”的规划语言。",
    )
    add_heading_cn(doc, "（三）中国供给与全球舞台", level=2)
    add_body(
        doc,
        "2026 年 1 月的 CES 上，中国企业在人形与具身赛道的展位密度显著抬升，被媒体概括为“中国时刻”：北京人形机器人创新中心携“具身天工”系列强调全自主分拣；智元、傅利叶、众擎、擎朗等展示工业、陪护与服务矩阵；卧安机器人全球首发家庭保姆机器人 onero，现场演示无遥操完成识别衣物、打开洗衣机、分类投放的流程。4 月汉诺威工博会上，人形机器人展商约 15 家，Agile Robots 的 Agile ONE、中联重科 Robot Ops 等把“物理 AI”明确写成工业系统的一部分。中国企业在消费电子展上的声量和在工业展上的系统集成能力，共同构成 2026 年的全球坐标：竞争已经从“是否做人形”转向“如何进入产线、仓配、康养和家庭”。",
    )


def build_ch3(doc):
    add_heading_cn(doc, "三、WAIC 2026 深度观察：五个结构性趋势", level=1)
    add_body(
        doc,
        "WAIC 2026 于 7 月 20 日收官。官方与主流媒体披露的规模数据高度一致：三地四馆、十万平方米、1100 余家企业、现场观众超 40 万人次、177 个采购团组、意向采购约 203.6 亿元。具身智能独立展馆（世博展览馆 H3）成为本届最受关注的物理空间之一。参展具身企业由上届 80 多家增至 200 多家，真机逾 300 台且几乎全部动态运行。智元联合创始人、总裁兼首席技术官彭志辉的概括被多家媒体引用：关注点不再是机器人能不能跑、跳舞、表演，而是能否进入工厂和商业服务场景、长时间稳定工作，以及跨任务泛化能力如何。",
    )
    add_quote(
        doc,
        "“机器人已经告别才艺表演，转而回答一个更朴素的问题：你能在真实场景里稳定干活吗？”",
        "——综合 WAIC 2026 收官观察报道",
    )

    add_heading_cn(doc, "（一）力触觉传感进入铺量阶段", level=2)
    add_body(
        doc,
        "过去力传感器是夹爪末端的一颗器件，2026 年正沿关节、手掌、指尖乃至全身铺设。具身大模型要做精细操作，必须有高密度、低延时的力触觉输入；量产成本与体积压缩，使“全身铺设”开始具备工程可行性。蓝点触控展示精度达 0.1%FS、支持 10kHz 响应的六维力与关节扭矩传感器，并称广东基地建有全自动产线，年设计产能关节力传感器 100 万套、六维力 20 万套，交付周期 2 至 3 周。他山科技发布动态触觉芯片绿宝石 E10A 与全手传感器模组，称视触融合检测工作站作业成功率达 99.8% 以上。戴盟等厂商演示仅凭触觉分辨血糖试纸是否空包，用以说明视觉在透明、遮挡、反光场景的局限。对住房场景而言，力触觉是机器人能否在狭小厨房、卫生间、床边安全接触老人与易碎物的前提，而不仅仅是工业装配的精度问题。",
    )

    add_heading_cn(doc, "（二）灵巧手从“能动”到“能用、能卖”", level=2)
    add_body(
        doc,
        "连杆、腱绳、直驱三条技术路线在同一展馆全系列摆开，争论焦点从“哪条更优”转向产线能力、成本曲线和交付节奏。灵心巧手展出“灵巧手组装灵巧手”全自动产线，并将灵巧手作为独立操作智能模块出售；因时机器人现场摆出六大系列，强调全栈自研自产。知行机器人展示可左右手自适应切换的五指手与工业级三指手。强脑科技 Revo3 在全掌布置阵列触觉、指尖融合视触觉，同时坦言防水、宽温域与百万次寿命仍难比人手。住房政策视角下，灵巧手价格下探决定了“家务机器人”会先进入哪一类住房：若末端执行器仍是高成本模块，则家庭渗透将高度依赖租赁、共享和社区集中部署，而不是户户购买。",
    )

    add_heading_cn(doc, "（三）数据从内部工序升格为独立产业层", level=2)
    add_body(
        doc,
        "与语言模型百亿小时级语料相比，高质量真实物理交互数据极度稀缺。光轮智能首次参加 WAIC，被报道为全球首个具身数据独角兽，2026 年估值突破 150 亿元，其联合创始人杨海波称数据与评测需求较去年增长百倍至千倍，从数千小时跃至数百万小时量级。博登智能在宁波、湖州、马鞍山布局超 3 万平方米创新中心，称年产 50 万小时真机数据与百万小时级 Ego 数据。穹彻智能展示无本体采集系统 RoboPocket，并称依托 47 个城市十万小时级家庭场景数据训练。简智机器人强调有效数据必须覆盖头、手、身与全维度感知。对住房研究最敏感的一点是：家庭正在被当作数据矿。若缺乏知情同意、数据最小化与公共监管，所谓“入户训练”可能演变为对私人居住空间的系统性采集。",
    )

    add_heading_cn(doc, "（四）具身大模型走向长程自主与跨本体泛化", level=2)
    add_body(
        doc,
        "2025 年常见“叠了一件衣服”的单技能展示；2026 年的追问变成“一条指令能串联多少步”“换一台本体还能不能用”“连续工作十余小时稳不稳”。千寻智能展示 Moz1 接收单条指令后整理客厅，将物品归位冰箱、送碗入洗碗机，被随机抛出纸巾后数秒内重规划；并称已在宁德时代产线完成高压测试插头插接，成功率 99% 以上。原力灵机联合阶跃星辰发起 6 台机器人 15 小时协同拼装 3.5 米长城积木的挑战。RoboScience 演示 30 秒更换不同品牌灵巧手仍可抓取。VLA 与世界模型的融合架构，正在取代非此即彼的路线之争。住房场景需要的正是这种长程能力：老人说“我渴了”，机器人要完成定位水杯、避开杂物、评估防滑与温度、送至手边并确认饮用安全，而不是只执行“抓杯子”一个原子动作。",
    )

    add_heading_cn(doc, "（五）量产、开发平台与商业闭环成为新竞争面", level=2)
    add_body(
        doc,
        "当本体和模型足够多之后，瓶颈转为：不同品牌如何共享运动能力、终端客户如何无需工程师即可部署、如何把一台样机变成一千台可交付产品。加速进化发布搭载 NVIDIA Thor 的 Booster T2 开发平台；桥介数物展示跨机型运动能力平台 RoboCraft AI；数字华夏发布跨本体场景大脑 RoboEase；卓益得仿生机器人 Moya 据报道完成千台量产签约，覆盖康养、汽车、通讯等场景。商业模式上，企业越来越多地出售“场景评估—方案定制—部署—运营”的全栈服务，而不是单台硬件。对住房和物业行业，这意味着未来采购对象可能是服务商而非设备商，合同形态更接近设施管理与长期照护服务，而不是一次性家电购买。",
    )


def build_ch4(doc):
    add_heading_cn(doc, "四、全球与国内展会扫描：一年之内的现场证据", level=1)
    add_body(
        doc,
        "仅看 WAIC 容易把上海的热度误认为全球同步。把 2026 年前八个月的主要展会与赛事连成一条时间线，可以更清楚地看到：工业现场、消费展会、城市级路测和即将到来的世界机器人大会，正在用不同语言讲述同一件事——具身智能进入部署态。",
    )

    add_heading_cn(doc, "（一）CES 2026：家庭叙事的全球首秀窗口", level=2)
    add_body(
        doc,
        "1 月 6 日至 9 日，拉斯维加斯 CES 被多家中文媒体称为人形机器人的“科技春晚”。与 WAIC 侧重工业作业不同，CES 更集中地贩售“进入日常生活”的想象。卧安 onero 以家庭保姆定位完成洗衣全流程自主演示；越疆发布消费级四足 Rover X1，主打家庭陪伴与管家；LG 推出 CLOiD 智能家居助手，双臂五指手与自家家电生态绑定；傅利叶海外系统展示陪伴交互。CES 的意义在于国际舆论场：它让“中国供给 + 家庭场景”成为一种可被全球观众直接观看的产品语言。但展会家庭是布景，真实住宅是迷宫。CES 证明需求叙事成立，并不证明 2026 年的普通住房已经具备接纳条件。",
    )

    add_heading_cn(doc, "（二）汉诺威工博会：物理 AI 被写成工业系统", level=2)
    add_body(
        doc,
        "4 月 20 日至 24 日的汉诺威工业博览会，主题词是工业 AI。主办方介绍约 15 家企业展示人形系统。Agile Robots 将 Agile ONE 定义为工厂地板上的一员，与机械臂、移动机器人共同组成生产系统，而不是舞台上的独立明星；西门子用制鞋柔性产线演示自主打包机器人与人形机器人协同；中联重科全球首发 Robot Ops，把 DataOps、DevOps 与 AgentOps 写成可复用的工程平台，并演示轮式人形与物流移动机器人协同分拣。汉诺威给出的政策启示是：欧洲产业界并不把人形视为消费电子，而视为补齐“难以标准化工序”的柔性劳动力。中国若只在消费展会上讲家庭故事、在工业展会上却缺少系统集成，住房端的机器人也难以获得可靠的运维体系。",
    )

    add_heading_cn(doc, "（三）北京亦庄半马：城市本身成为测试场", level=2)
    add_body(
        doc,
        "4 月 19 日，2026 北京亦庄半程马拉松暨人形机器人半程马拉松举行，超百支机器人赛队与约 1.2 万名人类选手共跑。媒体报道称荣耀机器人以 50 分 26 秒完赛夺冠。比成绩更重要的是制度设计：亦庄以“具身智能十条”支持购销与场景开放，投用北京市首个人形机器人中试验证平台，启动京津冀首个具身智能超级工厂建设，实施社会实验计划，开放制造、康养、酒店商超等场景，并建成全球首个智慧康养机器人养老驿站，同步授牌“居家适老化改造公共样板间”。这是目前国内最接近“住房—社区—产业”闭环的地方实践：机器人不是关在园区里的展品，而是被允许出现在街道、驿站和样板间里。",
    )

    add_heading_cn(doc, "（四）世界机器人大会前夜：人机共生被写成大会主题", level=2)
    add_body(
        doc,
        "本白皮书完稿时，2026 世界机器人大会（WRC）将于 8 月 19 日至 23 日在北京北人亦创国际会展中心举行，主题为“人机共生，产需共融”。官方信息显示参展企业约 300 余家、展品超 2000 件、首发新品突破 150 件，设智创、智合、智造、智趣四馆，应用场景明确包括生产制造、仓储物流、餐饮零售、医疗康养、安全生产与应急救援。大会还将启动“全球机器人应用探索计划”，匹配量产人形、四足与灵巧手供创新团队免费试用。对住房政策观察者而言，WRC 的价值在于把“康养”和“人机共生”从论坛词汇写成展区目录——这意味着公共服务部门已经成为明确的需求侧。",
    )

    add_heading_cn(doc, "（五）下半年国内工业与科技展：量产叙事将继续被核验", level=2)
    add_body(
        doc,
        "第 26 届中国国际工业博览会定于 10 月 12 日至 16 日在国家会展中心（上海）举行，机器人展规划面积超 6 万平方米，主题指向“具身智能·智领未来”，强调告别概念展示、面向工业级长期运行。11 月 26 日至 28 日，第二十八届高交会将在深圳举办人工智能与机器人双馆，并设“具身智能演练场”。下半年这两场展会，将检验 WAIC 上宣布的量产承诺是否转化为可参观的产线与可下单的交期。住房与物业部门应把工博会、高交会视为采购与标准学习的窗口，而不是科技新闻的背景板。",
    )

    add_caption(doc, "表 1  2026 年与具身智能密切相关的主要展会与赛事")
    add_table(
        doc,
        ["时间", "活动", "地点", "与本白皮书相关的观察要点"],
        [
            ["1 月 6–9 日", "CES 2026", "拉斯维加斯", "家庭保姆、陪伴与家电生态；中国企业供给密度显著上升"],
            ["4 月 19 日", "亦庄人形机器人半马", "北京经开区", "城市级路测；康养驿站与适老样板间同步出现"],
            ["4 月 20–24 日", "汉诺威工博会", "德国汉诺威", "工业 AI 与多机协同；人形被写入产线系统"],
            ["7 月 17–20 日", "WAIC 2026", "上海", "具身独立成馆；从炫技到连续作业；数据基础设施登台"],
            ["8 月 19–23 日", "世界机器人大会", "北京亦庄", "人机共生；医疗康养列入展区；应用探索计划"],
            ["10 月 12–16 日", "中国工博会", "上海", "工业量产与核心零部件；长期运行稳定性"],
            ["11 月 26–28 日", "高交会", "深圳", "具身演练场；粤港澳产业链与消费电子接口"],
        ],
        col_widths=[3.2, 3.6, 3.0, 5.8],
    )


def build_ch5(doc):
    add_heading_cn(doc, "五、场景落地路径：工厂先行、康养跟进、家庭后置", level=1)
    add_heading_cn(doc, "（一）已经收敛的顺序", level=2)
    add_body(
        doc,
        "彭志辉在 WAIC 期间重申智元“沿途下蛋”策略：工厂率先落地，商业服务（迎宾、导购、接待）其次，家庭最后，前提是安全性、泛化能力和成本门槛得到解决。宇树科技 CMO 王其鑫亦表示，终极目标是服务千家万户，但必须先历经工业、商业、康养打磨。这一顺序与住房政策的风险排序一致：工厂是受控环境，社区与机构是半受控环境，家庭是高风险、高隐私、高差异环境。政策若倒置顺序，鼓励尚未稳定的全自主家务机器人直接进入保障房或高龄独居家庭，将把技术风险转嫁给最脆弱的居住者。",
    )
    add_heading_cn(doc, "（二）工业与物流：连续作业被当场演示", level=2)
    add_body(
        doc,
        "WAIC 现场，乐聚复刻工厂拆垛上料；极智嘉以人形、移动机器人与拣选工作站编队作业；智元与均普智能展示芯片上料、装盒、转运超长流程；它石智航 1:1 复刻线束工厂环形流水线；微亿智造称多臂质检已在中、美、德车企工厂运行。物流侧，星动纪元称已在邮政与快递枢纽常态分拣。工业场景的政策含义主要不在住房，而在产业空间：机器人超级工厂、中试平台和实训场将改变园区用地结构，并可能通过高技能就业推高周边住房需求。",
    )
    add_heading_cn(doc, "（三）康养：家庭之前最重要的过渡带", level=2)
    add_body(
        doc,
        "傅利叶在 WAIC 展示“具身之家”居家陪伴全链路验证 Demo，同时明确这不是量产家用产品；其康养策略是医院和专业机构切入，再向社区与家庭延伸，并用轮式高负载双臂机器人 GRW 覆盖搬运与照护中的体力环节。北京亦庄养老驿站把迎宾机器人、按摩理疗与适老监测放进同一空间。工信部、民政部开展的智能养老服务机器人结对攻关与场景应用试点（2025—2027 年）要求：居家产品不少于 200 户、200 台套验证；社区和机构类不少于 20 个社区或 20 家机构。森丽康“小丽”等产品已进入多地养老院、社区和家庭试点。本中心认为，康养机构是 2026—2028 年最应该加大公共投入的场景：它能积累照护数据、训练护理人员与机器人的分工，并为后续入户提供安全标准，而不是用“机器人替代护工”的口号压缩人力。",
    )
    add_heading_cn(doc, "（四）家庭：两条路径，一个终点", level=2)
    add_body(
        doc,
        "家庭赛道在 2026 年明显分化。一条路径押注劳动价值：破壳机器人目标是先做成 10 件家务；自变量联合 58 到家进入超 1000 户，核心目的是换取真实家庭行为与环境数据；未来不远机器人以约 3000 元/月先租后买，截至 2026 年 7 月覆盖 500 余家庭、累计服务超 5 万小时。另一条路径押注情绪价值：优必选 U1 定位下一代智能终端；松延动力“小布米”切入儿童陪伴；越疆“鹿萌”、节卡、数字华夏、心言集团“巴布”等强调陪伴、教育与主动交互。IDC 被引述称 2026 年国内教育陪伴机器人市场规模将突破 10 亿美元。两条路径短期分流，长期应汇合为“能干活也能陪伴”。住房政策不宜只为“全能保姆”预留空间，而应同时规范轻量化陪伴产品进入儿童房与老年房所涉及的内容安全、摄像头与语音数据问题。",
    )


def build_ch6(doc):
    add_heading_cn(doc, "六、住房、社区与城市：当机器人成为“住户”", level=1)
    add_heading_cn(doc, "（一）空间设计：机器人可达性应成为新的适老指标", level=2)
    add_body(
        doc,
        "中房网等行业媒体 2026 年 6 月已提出：养老助残机器人要求入户门净宽不小于 80 厘米、床边通道预留 40—45 厘米；配送与清洁机器人依赖平层地面、无障碍坡道和电梯轿厢尺寸。“机器人可达性”可能成为户型与公区设计的重要指标，影响老旧小区更新与新盘溢价。这与既有无障碍设计高度同向：服务机器人的通行需求，在很大程度上就是轮椅、助行器和护理人员的通行需求。把机器人友好写成与无障碍、适老化同一套规范，可以避免另起炉灶，也有助于公共资金形成合力。",
    )
    add_body(
        doc,
        "新房层面，应在住宅设计规范与商品房预售条件中逐步明确：户内连续通行宽度、门槛与高差、插座与充电舱位、网络与算力接口、电梯轿厢承重与开门宽度、快递与垃圾机器人的楼栋停靠点。存量房层面，老旧小区更新和居家适老化改造，应将“机器人可进入、可充电、可回传必要状态”作为可选项，而不是强制每户购买机器人。保障性住房与公租房尤其需要防止“智能溢价”把技术福利只留给高价商品房。",
    )
    add_heading_cn(doc, "（二）住宅作为训练场：30 万套户型的数字孪生启示", level=2)
    add_body(
        doc,
        "开源项目将约 30 万套中国真实住宅户型和 5000 个可交互三维家庭场景用于机器人训练，说明中国住房存量的空间多样性，本身就是具身智能的战略数据资产。这对住房研究有两层含义。第一，户型标准化与家具模数化，不仅影响居住品质，也影响机器人泛化成本——过度个性化的精装可能提高服务机器人部署费用。第二，数字户型一旦用于训练，就涉及测绘数据、室内布局与生活痕迹的敏感信息，住房和城乡建设、数据、公安等部门需要明确：哪些住宅数据可以开源训练，哪些必须脱敏或禁止出境。",
    )
    add_heading_cn(doc, "（三）社区服务配套与区域住房价值", level=2)
    add_body(
        doc,
        "机器人对住房市场的影响，至少通过两条通道发生。一是社区服务配套：率先配置巡检、配送、康养驿站机器人的小区，可能形成对时间贫困家庭和老年家庭的吸引力，从而获得一定溢价。二是产业人口集聚：机器人研发、制造、运维总部所在城市及周边，高技能就业将支撑住房需求韧性。本中心建议，城市住房发展规划增加观察指标：社区机器人配置率、康养机器人驿站覆盖率、机器人产业就业与租赁住房匹配度。同时必须警惕替代效应：若基层物业、保洁、配送岗位被大规模替代，部分租赁住房的需求结构会变化，社会政策要同步安排转岗与再就业，而不能只计算“效率红利”。",
    )
    add_heading_cn(doc, "（四）支付体系：把机器人写入适老化“场景包”", level=2)
    add_body(
        doc,
        "2026 年适老化产品政策出现一个关键变化：支付开始按场景而不是按监管分类发钱。西安等地将外骨骼助力机器人、电动轮椅与智能马桶、智能门锁列入同一张以旧换新或补贴清单；北京、上海亦将外骨骼等纳入补贴档。深圳居家适老化“焕新”将助行机器人、喂饭机器人、情感陪伴装置等纳入，按实际售价 30%、单人累计最高 1 万元补贴。长护险、国补、适老改造补贴开始交叉覆盖同一批老人。住房政策应主动与民政、工信、医保对表，把“住宅内可安装、可通行、可充电”作为补贴硬件的前置条件，避免出现“补了设备却进不了门、充不了电、连不上物业平台”的浪费。",
    )
    add_heading_cn(doc, "（五）居住数据权利与人机共居安全", level=2)
    add_body(
        doc,
        "家庭机器人必然携带摄像头、麦克风、深度传感器和地图构建能力。它看到的是冰箱里的药品、卧室的行动轨迹、儿童的学习内容和老人的健康状态。现行《个人信息保护法》提供了原则框架，但“具身设备在私人住宅中的持续感知”仍缺少专门规范。本中心主张四条底线：默认本地处理与最短存储；录制状态必须可见、可一键关闭；向厂商回传须单独同意并允许撤回；禁止将保障性住房、公办养老机构中的居住数据用于商业模型训练，除非另有匿名化与伦理审查。安全方面，家庭场景应强制要求碰撞力限制、防跌倒误伤、远程人工接管和“一键停机”，并明确物业与厂商的责任分际。",
    )


def build_ch7(doc):
    add_heading_cn(doc, "七、展望 2030：四个情景与关键变量", level=1)
    add_body(
        doc,
        "展望不是预言。本中心给出 2030 年四个情景，便于政策部门做压力测试。基准情景概率最高；政策的价值在于提高基准情景的质量、压低失序情景的概率。",
    )
    add_heading_cn(doc, "（一）情景 A：分层渗透（基准）", level=2)
    add_body(
        doc,
        "到 2030 年，工业与仓储物流中的具身机器人成为常见生产力工具，千台级、万台级部署在沿海制造业和枢纽物流中心不再稀奇。康养机构与社区驿站形成相对稳定的“机器人 + 护理员”班组，负责转运、巡视、陪伴和夜间看护辅助，但关键护理决策仍由人承担。家庭端，10% 左右有支付能力、户型条件较好的城镇家庭使用租赁或购买的服务机器人完成收纳、地面清洁、简单备餐辅助；更多家庭接触的是社区共享机器人和情感陪伴产品。住房设计规范完成一轮修订，新建全龄友好住宅普遍满足机器人与轮椅通行；存量小区通过适老化改造部分达标。这是本中心认为应当作为规划基准的情景。",
    )
    add_heading_cn(doc, "（二）情景 B：加速入户（乐观）", level=2)
    add_body(
        doc,
        "若 2027—2028 年世界模型与力控出现超预期突破，灵巧手成本快速下降，保险与远程接管体系成熟，则 2030 年家庭渗透可能明显高于基准。商品房营销将“机器人管家”写成标配，物业费结构改变，对家政与月嫂的替代效应显现。风险在于：住房差异被技术放大，老旧小区与保障房若未同步改造，会形成“智能居住鸿沟”。乐观情景下，公共部门更要提前做公平性安排，而不是放任市场溢价。",
    )
    add_heading_cn(doc, "（三）情景 C：工业落地、家庭延宕（保守）", level=2)
    add_body(
        doc,
        "若家庭安全事故、隐私争议或成本下降不及预期，产业将在工业端兑现，家庭端长期停留在展会与订阅试点。对住房政策而言，这并非最坏结果，但会造成“规划过度”：若现在按全面入户去改建所有住宅，可能形成无效投资。因此本白皮书主张“新房按可达性预留、存量按需求改造、机构优先部署”，而不是一刀切智能化。",
    )
    add_heading_cn(doc, "（四）情景 D：失序扩散（需要防范）", level=2)
    add_body(
        doc,
        "失序情景的特征是：未成熟产品在补贴刺激下进入高龄独居家庭；数据在住宅中被过度采集；物业、厂商、家属责任不清导致事故后无人负责；基层服务岗位快速流失而转岗政策缺位。2030 年若出现这一情景，受损的不仅是产业声誉，更是公众对智慧养老和住房公共服务的信任。防范之道是把安全标准、数据权利、责任保险和就业过渡写成 2026—2027 年即可启动的制度，而不是 2030 年的愿景。",
    )
    add_caption(doc, "表 2  面向 2030 年的关键变量")
    add_table(
        doc,
        ["变量", "2026 年状态（据公开报道）", "对 2030 年住房场景的含义"],
        [
            ["连续作业与异常恢复", "工厂演示连续流程，家庭仍易卡住", "达不到稳定阈值则家庭只能租赁+远程兜底"],
            ["力触觉与灵巧手成本", "开始铺量，价格有下探趋势", "决定家务机器人是购买品还是社区共享品"],
            ["家庭实景数据治理", "入户采集已出现，专门规范不足", "治理滞后将抑制正当应用、放大隐私风险"],
            ["康养机构验证", "试点启动，驿站与结对攻关推进", "机构跑通是入户的安全闸门"],
            ["住宅与公区适老改造", "补贴清单开始纳入机器人相关产品", "门宽、高差、电梯、充电位成为硬约束"],
            ["商业模式", "硬件销售与全栈运营并存", "物业和养老服务商可能成为真正买家"],
        ],
        col_widths=[3.4, 5.4, 6.8],
    )
    add_heading_cn(doc, "（五）2030 年住房政策应达到的状态", level=2)
    add_body(
        doc,
        "无论技术走得快或慢，到 2030 年，一套稳健的住房政策应当做到：新建住宅的无障碍与机器人可达性基本并轨；重点城市的社区养老服务设施具备人机协同条件；保障性住房和老旧小区改造中的智能投入可测量、可审计、不排斥低收入家庭；居住空间数据采集有法可依；房地产和物业评价增加“社区机器人服务能力”维度，但禁止以机器人概念进行误导性销售。技术可以迟到，规则不应缺席。",
    )


def build_ch8(doc):
    add_heading_cn(doc, "八、政策建议", level=1)
    add_heading_cn(doc, "（一）把“机器人可达性”并入无障碍与适老化标准", level=2)
    add_body(
        doc,
        "建议住房和城乡建设部门在修订住宅设计规范、无障碍设计规范和居家适老化改造指南时，增加通行宽度、地面高差、电梯轿厢、充电与停靠、网络接口等条款，明确其同时服务于轮椅使用者与服务机器人。新房一次性做到位，成本远低于未来拆改。",
    )
    add_heading_cn(doc, "（二）机构与社区优先，家庭试点严控安全", level=2)
    add_body(
        doc,
        "公共资金应优先支持养老机构、社区驿站、保障房小区的共享服务机器人，而不是补贴私人购买未经验证的全自主家务机器人。家庭试点必须满足：连续运行验证、强制保险、可见的录制状态、一键停机和远程人工接管。对标工信部、民政部结对攻关的户数与机构数要求，形成可检查的验收指标。",
    )
    add_heading_cn(doc, "（三）建立住宅场景数据与伦理规则", level=2)
    add_body(
        doc,
        "建议网信、住建、民政联合出台《家庭与社区服务机器人数据指引》，明确默认本地化、最短存储、禁止从保障性住房和公办机构批量抽取可识别生活数据用于商业训练。支持建设公共、脱敏的中国住宅空间仿真数据集，降低企业对私人住宅的采集依赖。",
    )
    add_heading_cn(doc, "（四）重构物业、养老服务与住房金融的支付接口", level=2)
    add_body(
        doc,
        "允许符合条件的机器人服务费进入物业费或社区养老服务包，探索与长护险、适老化补贴的衔接。住房金融方面，对加装充电、通行改造和社区机器人停靠设施的老旧小区更新项目给予绿色或适老金融支持。禁止将“未来机器人入户”作为期房溢价的主要卖点，除非设施已经建成并通过验收。",
    )
    add_heading_cn(doc, "（五）以城市为单元开展社会实验，而不是以楼盘为单元炒作", level=2)
    add_body(
        doc,
        "亦庄经验表明，政策、中试、场景开放、驿站和赛事可以组成闭环。建议上海、深圳、杭州等城市在既有智慧社区和养老服务体系上，划定若干街道开展“人机共居”社会实验，重点评估事故率、老人接受度、护理员工作强度变化、物业成本与住房满意度，而不是评估“新闻曝光量”。实验结果应公开，供住房保障和房地产调控参考。",
    )
    add_heading_cn(doc, "（六）就业过渡与住房需求监测同步", level=2)
    add_body(
        doc,
        "机器人在保洁、配送、基层物业岗位的替代可能改变部分租赁需求。建议人社与住建建立联合监测：机器人部署密度、相关岗位工资与流动性、周边租赁成交。对受影响劳动者，优先提供机器人运维、远程接管、适老服务等新岗位培训，使“人机共生”不落成“人被挤出社区”。",
    )


def build_appendix(doc):
    add_heading_cn(doc, "附录一  2026 年重要展会、赛事与政策节点", level=1)
    add_table(
        doc,
        ["日期", "节点", "要点"],
        [
            ["2025-06", "工信部、民政部结对攻关通知", "智能养老服务机器人家庭/社区/机构验证要求"],
            ["2026-01-06", "CES 2026 开幕", "卧安 onero、具身天工、Care-bot 等家庭与工业产品亮相"],
            ["2026-04-19", "亦庄人形机器人半马", "城市级测试；康养驿站与适老样板间"],
            ["2026-04-20", "汉诺威工博会", "工业 AI；Agile ONE、Robot Ops"],
            ["2026-07-17", "WAIC 2026 开幕", "主题“智能伙伴 共创未来”；具身独立成馆"],
            ["2026-07-20", "WAIC 闭幕", "1100 余家企业；意向采购约 203.6 亿元"],
            ["2026-08-19", "世界机器人大会", "主题“人机共生，产需共融”（白皮书完稿时即将举行）"],
            ["2026-10-12", "中国工博会（上海）", "机器人展聚焦量产与长期运行"],
            ["2026-11-26", "高交会（深圳）", "人工智能与机器人双馆、具身演练场"],
        ],
        col_widths=[3.2, 5.2, 7.2],
    )

    add_heading_cn(doc, "附录二  主要公开资料来源", level=1)
    add_body(doc, "以下来源用于事实核对与现场描述，引用时已做综合转述。网址以检索时为准。", first_indent=True)
    sources = [
        "世界人工智能大会相关报道：《WAIC 2026 收官：我们看到具身智能正在发生五个结构性趋势》（腾讯新闻，2026-07-20）；《WAIC 2026 现场，机器人扎堆“干活”》（新浪财经 / 孙小程，2026-07-19）；《从“炫技”到“实干”：WAIC 2026 见证 AI 进入物理世界》（央广网，2026-07-29）；《智能伙伴 共创未来》大会综述；香港 01《机械人不在舞台表演，而是进厂“打工”》。",
        "CES 2026：《具身智能迎来中国时刻》（东方财富网等综合报道，2026-01）；卧安机器人官方《onero 人形具身机器人全球首发》；北京人形机器人创新中心 CES 报道。",
        "汉诺威工博会：Hannover Messe 官方新闻稿（2026-04）；Agile Robots《Humanoid Agile ONE embodies Physical AI》；中联重科 Robot Ops 全球首发通稿（PR Newswire，2026-04-22）。",
        "北京亦庄：经开区官网关于 2026 半马、具身智能十条、养老驿站与中试平台的报道；21 世纪经济报道相关综述。",
        "世界机器人大会：大会官网展商手册与同期活动；北京市政府英文新闻稿 World Robot Conference 2026 to Start on August 19。",
        "家庭与康养：《用户不买单，资本狂下注，家用机器人到底赌什么？》（投中网 / 亿欧网，2026-08-11）；《WAIC 2026 观察：康养赛道的工程化样本》（前沿在线，2026-07-24）；工信部、民政部结对攻关通知；新华社《瞭望》人工智能拓展家政服务报道。",
        "住房交叉：中房网《当机器人成为“住户”》（2026-06-24）；适老化产品与补贴政策公开报道；《中华人民共和国国民经济和社会发展第十五个五年规划纲要》具身智能与数智生活相关章节。",
    ]
    for i, s in enumerate(sources, 1):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.left_indent = Cm(0.74)
        p.paragraph_format.first_line_indent = Cm(-0.74)
        run = p.add_run(f"[{i}] {s}")
        set_run_font(run, size=10.5, color=DARK, font="宋体")

    add_heading_cn(doc, "附录三  术语简释", level=1)
    terms = [
        ("具身智能", "使人工智能获得身体，能在物理世界中感知、决策并执行动作的技术范式，人形、轮式双臂、四足等均是其本体形态。"),
        ("VLA", "Vision-Language-Action，视觉—语言—动作模型，将看懂、听懂与动手放在同一套策略中。"),
        ("世界模型", "对物理世界动态进行预测与模拟的模型，用于规划、纠错和仿真训练。"),
        ("灵巧手", "高自由度末端执行器，用于抓取、装配、家务等精细操作，正成为独立产品品类。"),
        ("机器人可达性", "本白皮书使用的住房政策概念，指住宅户内与公区允许服务机器人安全通行、停靠、充电与作业的空间条件。"),
        ("沿途下蛋", "产业界对按技术成熟度匹配场景的策略表述：工厂—商服—康养—家庭依次落地。"),
    ]
    for name, desc in terms:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.5
        r1 = p.add_run(name + "  ")
        set_run_font(r1, size=12, bold=True, color=NAVY, font="黑体")
        r2 = p.add_run(desc)
        set_run_font(r2, size=12, color=DARK, font="宋体")

    add_heading_cn(doc, "结  语", level=1)
    add_body(
        doc,
        "2026 年的展会告诉我们：具身智能已经从“有没有”转向“好不好用”。2030 年的住房将告诉我们：这项技术究竟是扩大了居住福利，还是制造了新的空间不平等。复旦大学住房政策研究中心愿意把后一个问题，持续放在公共讨论的桌面上。机器人可以进厂打工，也可以进社区助老；但住房首先是人的权利。任何智能身体进入住宅，都应当使人更有尊严地居住，而不是使人变成数据采集对象或被溢价排斥的住户。",
    )
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(18)
    run = p.add_run("（全文完）")
    set_run_font(run, size=12, color=GRAY, font="楷体")


def build():
    doc = Document()
    setup_header_footer(doc)
    build_cover(doc)
    build_statement(doc)
    build_toc(doc)
    build_abstract(doc)
    build_ch1(doc)
    build_ch2(doc)
    build_ch3(doc)
    build_ch4(doc)
    build_ch5(doc)
    build_ch6(doc)
    build_ch7(doc)
    build_ch8(doc)
    build_appendix(doc)
    doc.save(OUT)
    print(f"已生成：{OUT}")
    print(f"文件大小：{OUT.stat().st_size / 1024:.1f} KB")


if __name__ == "__main__":
    build()
