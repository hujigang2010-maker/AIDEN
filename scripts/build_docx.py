"""
森马（上海）国际运营中心 · 项目介绍与解读 Word 生成器
Semir Global Headquarter — Project Introduction & Interpretation
"""

from __future__ import annotations
import os
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ---------- Theme ----------
PRIMARY  = RGBColor(0xC2, 0x18, 0x4B)   # 森马磁红
ACCENT   = RGBColor(0x55, 0x2A, 0xA0)   # 紫
ACCENT2  = RGBColor(0x06, 0x6B, 0x7C)   # 青
INK      = RGBColor(0x1A, 0x1A, 0x24)
TEXT     = RGBColor(0x2A, 0x2A, 0x36)
DIM      = RGBColor(0x55, 0x57, 0x6E)
MUTE     = RGBColor(0x80, 0x83, 0x99)
BG_SOFT  = 'F6F5FA'
BG_SOFT2 = 'FDF2F7'
BG_SOFT3 = 'F0F8FA'
RULE     = 'D9D7E4'

CN = '思源宋体'
CN2 = '思源黑体 CN'
EN = 'Inter'


def set_run_cn(run, font=CN2):
    """Ensure CJK font is also applied."""
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font)
    rFonts.set(qn('w:ascii'), font)
    rFonts.set(qn('w:hAnsi'), font)


def add_styled_run(p, text, *, size=11, color=TEXT, bold=False, italic=False, font=CN2):
    run = p.add_run(text)
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    set_run_cn(run, font)
    return run


def add_para(doc, text='', *, size=11, color=TEXT, bold=False, italic=False, font=CN2,
             align=WD_ALIGN_PARAGRAPH.LEFT, before=4, after=4, line_spacing=1.6,
             first_line_indent=None, left_indent=None):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line_spacing
    if first_line_indent is not None:
        pf.first_line_indent = Cm(first_line_indent)
    if left_indent is not None:
        pf.left_indent = Cm(left_indent)
    if text:
        add_styled_run(p, text, size=size, color=color, bold=bold,
                       italic=italic, font=font)
    return p


def shade_cell(cell, hex_color: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def set_cell_border(cell, *, top=None, bottom=None, left=None, right=None, color=RULE, size=4):
    """size in 1/8 pt units."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for side, want in (('top', top), ('bottom', bottom), ('left', left), ('right', right)):
        if want is None:
            continue
        el = tcBorders.find(qn(f'w:{side}'))
        if el is None:
            el = OxmlElement(f'w:{side}')
            tcBorders.append(el)
        el.set(qn('w:val'), 'single' if want else 'nil')
        el.set(qn('w:sz'), str(size))
        el.set(qn('w:color'), color)


def add_horizontal_rule(doc, color=PRIMARY, width_pts=1.5):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    pPr.append(pBdr)
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(int(width_pts * 8)))
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), f'{color[0]:02X}{color[1]:02X}{color[2]:02X}')
    pBdr.append(bottom)
    pPr_spacing = OxmlElement('w:spacing')
    pPr_spacing.set(qn('w:before'), '120')
    pPr_spacing.set(qn('w:after'), '120')
    pPr.append(pPr_spacing)


def add_callout(doc, title_text, body_text, *, fill=BG_SOFT2, accent=PRIMARY):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = True
    cell = table.cell(0, 0)
    shade_cell(cell, fill)
    set_cell_border(cell, left=True, color=f'{accent[0]:02X}{accent[1]:02X}{accent[2]:02X}', size=24)
    set_cell_border(cell, top=False, bottom=False, right=False)

    # title
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    add_styled_run(p, title_text, size=11, color=accent, bold=True)

    # body
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_after = Pt(2)
    p2.paragraph_format.line_spacing = 1.6
    add_styled_run(p2, body_text, size=10.5, color=TEXT)

    # spacer
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def add_h1(doc, num, cn, en):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(24)
    p.paragraph_format.space_after = Pt(4)
    add_styled_run(p, f'{num}  ', size=22, color=PRIMARY, bold=True, font=EN)
    add_styled_run(p, cn, size=20, color=INK, bold=True)
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_after = Pt(8)
    add_styled_run(p2, en, size=10, color=MUTE, font=EN)
    set_run_cn(p2.runs[0], font=EN)


def add_h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(4)
    add_styled_run(p, '◆ ', size=13, color=PRIMARY, bold=True)
    add_styled_run(p, text, size=14, color=INK, bold=True)


def add_h3(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    add_styled_run(p, text, size=12, color=ACCENT, bold=True)


def add_bullet(doc, label, body):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.line_spacing = 1.7
    add_styled_run(p, '· ', size=11, color=PRIMARY, bold=True)
    if label:
        add_styled_run(p, label, size=11, color=INK, bold=True)
        add_styled_run(p, '  ', size=11)
    add_styled_run(p, body, size=11, color=TEXT)


def add_table_header(table, headers, *, fill='2A1633', text_color=RGBColor(0xFF, 0xFF, 0xFF)):
    row = table.rows[0]
    for i, cell in enumerate(row.cells):
        cell.text = ''
        shade_cell(cell, fill)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_styled_run(p, headers[i], size=10.5, color=text_color, bold=True)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


# ============================================================
def build():
    doc = Document()

    # Page setup
    for section in doc.sections:
        section.top_margin = Cm(2.4)
        section.bottom_margin = Cm(2.4)
        section.left_margin = Cm(2.6)
        section.right_margin = Cm(2.6)

    # Default style
    style = doc.styles['Normal']
    style.font.name = CN2
    style.font.size = Pt(11)
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), CN2)
    rFonts.set(qn('w:ascii'), CN2)
    rFonts.set(qn('w:hAnsi'), CN2)

    # =========  封面区  =========
    add_para(doc, 'SEMIR  ·  SHANGHAI  ·  MINHANG  ·  WUJING',
             size=9, color=MUTE, font=EN, align=WD_ALIGN_PARAGRAPH.CENTER,
             before=24, after=12, line_spacing=1.2)

    add_para(doc, '森马（上海）国际运营中心',
             size=28, color=INK, bold=True,
             align=WD_ALIGN_PARAGRAPH.CENTER,
             before=4, after=4, line_spacing=1.2)
    add_para(doc, '项目介绍与解读',
             size=18, color=PRIMARY, bold=True,
             align=WD_ALIGN_PARAGRAPH.CENTER, after=12)
    add_para(doc, "Semir Group's Global Headquarter — Introduction & Interpretation",
             size=11, color=DIM, font=EN, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
    add_horizontal_rule(doc, color=PRIMARY, width_pts=1.2)

    add_para(doc,
             '本文档基于《森马（上海）国际运营中心·项目总体介绍》提案 PPT，'
             '系统整理项目的战略定位、宏观区位、产业规划、商业落位与运营合作等核心信息，'
             '并附以专业视角的解读与价值评估，便于决策者、合作伙伴及内部团队快速把握项目全貌。',
             size=11, color=TEXT, line_spacing=1.85, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=8, after=8)

    # Key facts table
    add_h2(doc, '项目一图速读 · Project at a Glance')
    facts = [
        ('项目名称', '森马（上海）国际运营中心 · Semir Group\'s Global Headquarter'),
        ('项目区位', '上海市闵行区吴泾镇 · 元江路 · 15 号线元江路 TOD'),
        ('战略定位', '南部科创走廊新兴载体 · 大零号湾示范枢纽 · 长三角融合标杆'),
        ('产业方向', '科技潮玩产业策源高地 · IP 创制 / AI 设计 / 直播电商 / 选品仓储'),
        ('商业定位', 'Z·世代潮玩社交主场 · 潮玩元宇宙 · 青年引力场'),
        ('建筑规模', '总建筑面积 22 万㎡ · 商业建筑面积 5.2 万㎡ · 商业车位 1,500+'),
        ('楼宇构成', '6 栋楼宇 · 零售 / 总部 / Livehouse / 酒店 / 艺术 / 直播 / 集群 / 餐饮'),
        ('TOD 流量', '15 号线元江路站日均客流 5–7 万人次'),
        ('辐射人口', '15 分钟车行覆盖 24 万居住人口 + 12 万产业办公人口'),
        ('产业牌照', '中国百货协会潮玩次元商业专委会 · AI 潮玩产业基地（中国动漫集团）'),
    ]
    tbl = doc.add_table(rows=len(facts), cols=2)
    tbl.autofit = False
    tbl.columns[0].width = Cm(3.6)
    tbl.columns[1].width = Cm(12.2)
    for i, (k, v) in enumerate(facts):
        c0, c1 = tbl.rows[i].cells
        c0.width = Cm(3.6)
        c1.width = Cm(12.2)
        shade_cell(c0, BG_SOFT2)
        shade_cell(c1, 'FFFFFF')
        for c in (c0, c1):
            set_cell_border(c, top=True, bottom=True, left=False, right=False, color=RULE)
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(4)
        p0.paragraph_format.space_after = Pt(4)
        add_styled_run(p0, k, size=10.5, color=PRIMARY, bold=True)
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(4)
        p1.paragraph_format.space_after = Pt(4)
        p1.paragraph_format.line_spacing = 1.5
        add_styled_run(p1, v, size=10.5, color=TEXT)

    add_para(doc, '', after=0)

    # 目录
    add_h2(doc, '目  录 · Contents')
    toc_items = [
        ('一', '项目背景与战略意义', 'Background & Strategic Significance'),
        ('二', '项目概况：区位、交通与建筑', 'Overview: Location · Traffic · Architecture'),
        ('三', '产业规划解读', 'Industry Planning Interpretation'),
        ('四', '商业规划解读', 'Commercial Planning Interpretation'),
        ('五', '运营与合作策略', 'Operation & Cooperation Strategy'),
        ('六', '项目价值评估与机会分析', 'Value Assessment & Opportunities'),
        ('七', '风险提示与建议', 'Risks & Recommendations'),
        ('八', '结语', 'Closing Remarks'),
    ]
    for num, cn, en in toc_items:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Cm(0.6)
        add_styled_run(p, f'{num}、', size=11, color=PRIMARY, bold=True)
        add_styled_run(p, cn, size=11, color=INK, bold=True)
        add_styled_run(p, f'   {en}', size=9.5, color=MUTE, font=EN)

    doc.add_page_break()

    # ===================  第一章  ===================
    add_h1(doc, '一', '项目背景与战略意义', 'Chapter 1 · Background & Strategic Significance')

    add_h2(doc, '1.1  时代背景')
    add_para(doc,
             '"潮玩经济"是近年来中国消费市场最具增长力的新兴赛道之一。以盲盒、手办、二次元 IP、'
             '国潮文创、数字潮玩为代表的潮玩品类，正在快速汇聚 Z 世代的注意力与可支配收入，'
             '相关产业链涵盖 IP 孵化、设计研发、AI 创作、智能制造、仓储分销、电商直播、'
             '展演活动、社交体验等多个环节，呈现出"内容驱动 + 体验消费 + 跨界联名"的复合型业态。',
             size=11, color=TEXT, line_spacing=1.85, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_para(doc,
             '与此同时，上海正深入实施"南北转型"战略，闵行作为南部科创中心核心承载区，'
             '在"大零号湾"科创策源功能区的牵引下，正从传统制造业重镇加速向"科技 + 时尚 + 文创"'
             '的复合型创新城区跃迁。元江路—剑川路地区中心被列为闵行五大中心之一，'
             '能级比肩漕河泾、张江，是上海未来产城融合的关键节点。',
             size=11, color=TEXT, line_spacing=1.85, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_h2(doc, '1.2  森马的战略转型')
    add_para(doc,
             '森马集团从单一服装品牌运营商，逐步演化为"时尚 + 文创 + 产业地产"的综合性集团。'
             '本项目是森马继续深化产城融合战略、布局新一代潮玩产业的关键一子——'
             '通过自有总部 + 商业 + 产业园 + 配套酒店的"产城合一"模式，'
             '在闵行吴泾打造一个面向 Z·世代的国际化潮玩产业枢纽。',
             size=11, color=TEXT, line_spacing=1.85, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_h2(doc, '1.3  三大战略身份解读')
    add_callout(doc,
                '① 南部科创走廊的"新兴载体" — IMPORTANT CARRIER',
                '项目地处闵行南部科创主轴（元江路），是从"零号湾"创新策源向南拓展、'
                '产业承接落地的关键容器。其"新兴"二字意味着承担差异化角色：'
                '不与紫竹高新区、临港浦江争夺硬科技与先进制造，而是聚焦"潮玩 + AI + IP"等'
                '与时尚消费高度耦合的内容型科技赛道。',
                fill=BG_SOFT2, accent=PRIMARY)
    add_callout(doc,
                '② 大零号湾科创区的"示范枢纽" — INNOVATION NODE',
                '"枢纽"意味着既是流量节点（15 号线 TOD 日均 5–7 万人次），'
                '也是要素枢纽（汇集高校设计资源、央企资源、品牌资源、AI 技术资源）。'
                '"示范"则强调可复制性：一旦模式跑通，可向上海其他特色小镇、'
                '乃至长三角同类项目输出标准。',
                fill=BG_SOFT, accent=ACCENT)
    add_callout(doc,
                '③ 长三角"一核三带"融合标杆 — SUPER PLATFORM',
                '"一核"指上海大都市核心圈，"三带"涵盖沪宁、沪杭、沿江沿海创新走廊。'
                '项目作为"科技 + 时尚 + 文创"的融合性平台，凭借森马的全国品牌网络与产业链，'
                '具备成为长三角潮玩与文创跨城市协同节点的潜质。',
                fill=BG_SOFT3, accent=ACCENT2)

    add_h2(doc, '1.4  小结：从"地产项目"到"产业策源地"')
    add_para(doc,
             '本项目区别于普通商办综合体的核心，在于其"产业策源"属性：'
             '不只是提供空间载体，更主动构建"产业 + 商业 + 内容 + 流量"的循环。'
             '从战略意义上看，项目所承担的角色更接近一个"产业平台 + 城市文化客厅"的复合体，'
             '其成功与否将直接影响闵行南部科创走廊的潮玩 / 文创赛道占位，'
             '并为森马集团从"品牌商"向"生态运营商"的角色升级提供关键试验场。',
             size=11, color=TEXT, line_spacing=1.85, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # ===================  第二章  ===================
    add_h1(doc, '二', '项目概况：区位、交通与建筑', 'Chapter 2 · Overview')

    add_h2(doc, '2.1  宏观区位')
    add_para(doc,
             '项目位于上海市闵行区吴泾镇元江路主干道沿线，处于"元江路—剑川路"地区中心的核心位置。'
             '该区域在上海整体规划中具备三层属性：',
             size=11, color=TEXT, line_spacing=1.85, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_bullet(doc, '行政属性：', '闵行五大中心之一，与梅陇、七宝共筑闵行 TOD 金三角；')
    add_bullet(doc, '产业属性：', '元宇宙、未来能源、智慧医疗、人工智能、低空经济五大产业鼎立；')
    add_bullet(doc, '功能属性：', '南部科创走廊功能轴 + 东西交通主廊 + 产城融合纽带。')

    add_para(doc,
             '上海南部科创中心已形成八大主要产业功能区，与本项目形成层次清晰的协同结构：',
             size=11, color=TEXT, before=8, after=6, line_spacing=1.85,
             first_line_indent=0.74, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    tbl = doc.add_table(rows=9, cols=2)
    tbl.autofit = False
    tbl.columns[0].width = Cm(7.0)
    tbl.columns[1].width = Cm(8.8)
    add_table_header(tbl, ['功能区', '定位'])
    zones = [
        ('"零号湾"创新创业集聚区', '创新策源区'),
        ('紫竹国家高新技术产业开发区', '高新产业承载区'),
        ('临港浦江国际科技城', '高新产业承载区'),
        ('向阳工业互联网基地', '高新产业承载区'),
        ('闵行经济技术开发区', '先进制造业承载区'),
        ('莘庄工业区', '先进制造业承载区'),
        ('上海航天产业基地', '战略产业承载区'),
        ('马桥人工智能创新试验区', '战略产业承载区'),
    ]
    for i, (a, b) in enumerate(zones, start=1):
        c0, c1 = tbl.rows[i].cells
        c0.width = Cm(7.0); c1.width = Cm(8.8)
        for c in (c0, c1):
            set_cell_border(c, top=True, bottom=True, color=RULE)
            shade_cell(c, 'FFFFFF' if i % 2 else BG_SOFT)
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        add_styled_run(c0.paragraphs[0], a, size=10.5, color=TEXT)
        c0.paragraphs[0].paragraph_format.space_before = Pt(3)
        c0.paragraphs[0].paragraph_format.space_after = Pt(3)
        add_styled_run(c1.paragraphs[0], b, size=10.5, color=PRIMARY, bold=True)
        c1.paragraphs[0].paragraph_format.space_before = Pt(3)
        c1.paragraphs[0].paragraph_format.space_after = Pt(3)

    add_para(doc, '', after=4)
    add_callout(doc, '解读', 
                '本项目在该结构中扮演"文创 + 科技融合"的差异化角色，与硬科技承载区形成功能互补，'
                '而非内卷竞争。这种"内容型科创"卡位有助于项目获得政策、人才、流量的多方倾斜。',
                fill=BG_SOFT, accent=ACCENT)

    add_h2(doc, '2.2  交通辐射')
    add_para(doc,
             '元江路是闵行南部一条东西向主干道，西起昆阳北路、东至龙吴路，全长约 11.5 公里，'
             '横贯吴泾、颛桥、马桥镇及莘庄工业区。项目周边交通可达性如下：',
             size=11, color=TEXT, line_spacing=1.85, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    tbl = doc.add_table(rows=7, cols=3)
    tbl.autofit = False
    tbl.columns[0].width = Cm(4.2)
    tbl.columns[1].width = Cm(3.0)
    tbl.columns[2].width = Cm(8.6)
    add_table_header(tbl, ['目的地', '距离', '车程 / 备注'])
    rows = [
        ('紫竹高新区',         '5 km',  '车程约 10 min'),
        ('申嘉湖高速入口',     '4 km',  '车程约 15 min'),
        ('浦东国际机场',       '36 km', '车程约 1.5 h · 机场快线 1 h 内可达'),
        ('虹桥国际机场',       '17 km', '车程约 1 h · 机场快线 45 min 内可达'),
        ('大学城',             '4 km',  '车程约 10 min · 地铁可直达'),
        ('15 号线元江路站 TOD','—',    '日均客流 5–7 万人次'),
    ]
    for i, r in enumerate(rows, start=1):
        for j, v in enumerate(r):
            c = tbl.rows[i].cells[j]
            c.width = (Cm(4.2), Cm(3.0), Cm(8.6))[j]
            shade_cell(c, 'FFFFFF' if i % 2 else BG_SOFT)
            set_cell_border(c, top=True, bottom=True, color=RULE)
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            if j == 1:
                add_styled_run(p, v, size=10.5, color=PRIMARY, bold=True, font=EN)
            else:
                add_styled_run(p, v, size=10.5, color=TEXT)

    add_para(doc, '', after=4)
    add_callout(doc, '解读',
                '15 号线元江路站直连市区与南部科创主轴，5–7 万人次的日均吞吐量为项目提供了'
                '"消费 + 通勤 + 旅游"的三重客流基础，是项目能够支撑潮玩商业、'
                '艺术活动、直播电商等多业态的关键流量基础设施。',
                fill=BG_SOFT2, accent=PRIMARY)

    add_h2(doc, '2.3  立足周边：人口与产业')
    add_para(doc,
             '项目腹地保利住宅项目已落成，远处塘湾镇和别墅区人口稳定。3 km 范围内聚集了大量产业园和企业办公项目，'
             '生活服务配套方面逐步完善：',
             size=11, color=TEXT, line_spacing=1.85, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_bullet(doc, '居住人口：', '15 分钟车行覆盖 24 万；')
    add_bullet(doc, '办公人口：', '15 分钟车行覆盖 12 万；')
    add_bullet(doc, '产业坐标：',
              '森马产业、衣恋、普洛斯、雅诗兰黛、莲谷科技园、保利光合跃城、'
              '颛桥科技绿洲、拉夏贝尔、舜江集团总部 1 号、闵行物流园、'
              '中建产研、首开塘湾基地等。')

    add_h2(doc, '2.4  六栋楼宇 · 一座潮玩之城')
    add_para(doc,
             '项目由 6 栋楼组成，总建筑面积约 22 万㎡，商业建筑面积约 5.2 万㎡，'
             '商业停车位 1,500+。各楼宇功能呈现"垂直分层 + 主题集群"的结构：',
             size=11, color=TEXT, line_spacing=1.85, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    tbl = doc.add_table(rows=7, cols=2)
    tbl.autofit = False
    tbl.columns[0].width = Cm(2.4)
    tbl.columns[1].width = Cm(13.4)
    add_table_header(tbl, ['楼号', '功能 / 程序'])
    bldg = [
        ('1#', '1–4F 零售；5F 及以上 森马总部办公'),
        ('2#', '整栋 二次元主题 Livehouse / 秀场'),
        ('3#', '1–4F 休闲运动、萌宠空间；5F 及以上 酒店'),
        ('4#', '1–3F 潮玩艺术中心；4F 直播中心；5F 及以上 潮玩产业集群'),
        ('5#', '1–4F 动漫书店、休闲娱乐；5F 及以上 潮玩产业集群'),
        ('6#', '1–5F 品质生活、特色餐饮、服务配套、商务宴请'),
    ]
    for i, (a, b) in enumerate(bldg, start=1):
        c0, c1 = tbl.rows[i].cells
        c0.width = Cm(2.4); c1.width = Cm(13.4)
        for c in (c0, c1):
            set_cell_border(c, top=True, bottom=True, color=RULE)
            shade_cell(c, 'FFFFFF' if i % 2 else BG_SOFT)
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.space_before = Pt(3); p0.paragraph_format.space_after = Pt(3)
        add_styled_run(p0, a, size=11, color=PRIMARY, bold=True, font=EN)
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(3); p1.paragraph_format.space_after = Pt(3)
        add_styled_run(p1, b, size=10.5, color=TEXT)

    add_para(doc, '', after=4)
    add_callout(doc, '解读',
                '这一构成的关键设计在于"垂直业态 + 楼宇专业化"——零售 / 演艺 / 休闲 / 艺术 / 内容 / 餐饮分楼承担，'
                '既避免业态相互干扰，又让每栋楼具备独立的 IP 与运营节奏。'
                '4# 与 5# 高层"潮玩产业集群"则与底层潮玩商业形成"楼上做内容、楼下做体验"的闭环，'
                '是项目区别于传统商业综合体的核心差异点。',
                fill=BG_SOFT3, accent=ACCENT2)

    # ===================  第三章  ===================
    add_h1(doc, '三', '产业规划解读', 'Chapter 3 · Industry Planning')

    add_h2(doc, '3.1  产业定位：科技潮玩产业策源高地')
    add_para(doc,
             '项目以"科技潮玩产业策源高地"为统一定位，确立三大战略目标与五大优势聚焦：',
             size=11, color=TEXT, line_spacing=1.85, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_bullet(doc, '三大战略目标：', '国际化产业枢纽 · 品牌运营高地 · IP 创制中心；')
    add_bullet(doc, '五大优势聚焦：', '创意集聚 · 源头孵化 · 场景体验 · 生态复合 · 集约选品；')
    add_bullet(doc, '驱动机制：',
              '"产业转化 + 转化产业"双轮驱动 · "投资驱动 + 市场驱动"双管齐下。')

    add_callout(doc, '解读 · 两组"双"的含义',
                '"产业转化 + 转化产业"——前者指把高校 / 央企 / 协会的研究成果与 IP 资源在项目内转化为产品；'
                '后者指通过项目反向重塑潮玩产业的供给侧（设计标准、选品流程、AI 工具链）。\n'
                '"投资驱动 + 市场驱动"——前者指依靠森马自有资金、政府引导基金、机构合作设立产业基础；'
                '后者指通过商业流量、活动 IP、直播电商等市场化手段反哺产业生态。'
                '这种"政府 - 资本 - 市场"三角驱动是项目可持续运营的关键。',
                fill=BG_SOFT2, accent=PRIMARY)

    add_h2(doc, '3.2  五层产业金字塔')
    add_para(doc,
             '项目按层次构建一体化潮玩产业生态体系，比例与角色如下：',
             size=11, color=TEXT, line_spacing=1.85, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    tbl = doc.add_table(rows=6, cols=4)
    tbl.autofit = False
    tbl.columns[0].width = Cm(1.6)
    tbl.columns[1].width = Cm(5.4)
    tbl.columns[2].width = Cm(2.8)
    tbl.columns[3].width = Cm(6.0)
    add_table_header(tbl, ['占比', '类型', '角色', '配置'])
    layers = [
        ('10%', '头部央企 · 行业协会', '导向',  '3 个 · 2,000 ㎡'),
        ('10%', '共享配套服务体系',    '吸附点','3 个 · 2,000 ㎡'),
        ('20%', '中型潮玩运营企业',    '基础',  '4–6 个 · 5,000 ㎡'),
        ('40%', '小型潮玩运营企业',    '骨架',  '30 个 · 200–500 ㎡'),
        ('20%', '中小型潮玩服务机构',  '血肉',  '15 个 · 200–500 ㎡'),
    ]
    for i, r in enumerate(layers, start=1):
        for j, v in enumerate(r):
            c = tbl.rows[i].cells[j]
            c.width = (Cm(1.6), Cm(5.4), Cm(2.8), Cm(6.0))[j]
            shade_cell(c, 'FFFFFF' if i % 2 else BG_SOFT)
            set_cell_border(c, top=True, bottom=True, color=RULE)
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            if j == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                add_styled_run(p, v, size=11, color=PRIMARY, bold=True, font=EN)
            elif j == 2:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                add_styled_run(p, v, size=10.5, color=ACCENT, bold=True)
            else:
                add_styled_run(p, v, size=10.5, color=TEXT)

    add_para(doc, '', after=4)
    add_callout(doc, '解读 · 金字塔背后的逻辑',
                '该结构呈"小头 + 实腰 + 长尾"形态，比单纯"二八结构"更接近真实的潮玩产业生态。'
                '头部 10% 起信用背书与政策对接作用；20% 共享配套是"产业土壤"；'
                '中型 + 小型企业（合计 60%）承担实际产能；最末端 20% 服务机构承担"血肉补充"，'
                '形成 200–500㎡ 大量灵活工位的高密度入孵环境。'
                '这一比例与企业数量（3+3+5+30+15 ≈ 56 家）的设定，'
                '体现了项目对潮玩产业"创意密集型 + 长尾化 + 小批量多频次"特性的精准回应。',
                fill=BG_SOFT, accent=ACCENT)

    add_h2(doc, '3.3  双产业牌照')
    add_para(doc,
             '项目已确认与中国百货协会和中国动漫集团在本项目共同设立两大产业牌照：',
             size=11, color=TEXT, line_spacing=1.85, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_bullet(doc, '潮玩次元商业专委会：',
              '4 月起中百协潮玩次元商业专委会（筹）开始筹备，6 月在济南理事会会议正式挂牌成立，'
              '业务往来已正常开展；')
    add_bullet(doc, 'AI 潮玩产业基地：',
              '合作资源涵盖中国动漫集团、上海交大设计学院、闵行科协、森马集团等。'
              '未来活动包括"全国潮玩设计技能大赛""国漫·潮游集""动漫新品及跨界产品首发会"等。')

    add_callout(doc, '解读 · 牌照的真正价值',
                '在潮玩这种新兴产业，"标准制定权"和"行业话语权"是稀缺的战略资产。'
                '通过中百协（商业渠道权威）+ 中国动漫集团（内容与 IP 权威）双牌照，'
                '项目将获得：① 行业活动的优先承办权；② 行业标准与统计数据的合作话语权；'
                '③ 对接全国会员单位、设计院校、动漫机构的天然通道；'
                '④ 政府层面的政策对接信用增信。这是普通商业项目难以复制的"软基础设施"。',
                fill=BG_SOFT2, accent=PRIMARY)

    add_h2(doc, '3.4  六大产业配套平台')
    add_para(doc,
             '为支撑潮玩产业生态，项目自建六大产业配套平台：',
             size=11, color=TEXT, line_spacing=1.85, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    services = [
        ('选品中心',         '1# 4F',  '华东首个 IP 潮玩选品 + 仓储式销售空间，助力客户批发选品。'),
        ('代运营物流中心',   '森马二期仓库', '联动骏耀科技，提供智能仓储 + 物流代运营服务。'),
        ('共享直播中心',     '4# 4F',  '潮玩电商直播间，联动绮丽少女女团直播，服务产业销售需求。'),
        ('AI 共享设计中心', '4# 5F',  '联动高校人才资源，引入 AI 科技，共建 AI 潮玩设计中心。'),
        ('AI 共享打样 / DIY', '4# 5F', '助力客户快速打样，提高面世效率。'),
        ('潮玩产业展厅',     '5# 5F',  '聚焦品牌 IP 叙事，强化行业交流与渠道拓展。'),
    ]
    for n, loc, d in services:
        add_bullet(doc, f'{n}（{loc}）：', d)

    add_callout(doc, '解读 · 完整的"潮玩产业服务总线"',
                '这六个平台覆盖了潮玩从"创意—设计—打样—试销—展演—分销"的完整链条，'
                '具有显著的全产业链整合属性。尤其是 AI 共享设计 + AI 打样的组合，'
                '将传统 3–6 个月的产品周期压缩到数周量级，极大降低中小潮玩创业团队的门槛。'
                '这是项目最具技术含量与未来想象空间的核心环节。',
                fill=BG_SOFT3, accent=ACCENT2)

    add_h2(doc, '3.5  三轨并行的产业服务体系')
    tbl = doc.add_table(rows=4, cols=2)
    tbl.autofit = False
    tbl.columns[0].width = Cm(3.6)
    tbl.columns[1].width = Cm(12.2)
    add_table_header(tbl, ['模式', '内容'])
    triple = [
        ('企业主导自营',
         '保障配套服务提供能力：自持综合服务中心、园区餐饮、商务办公、'
         '中小型设计工作室、人才与财务服务中心。'),
        ('绿色通道协助',
         '发挥行政力量的能动性，申请街道政府开启绿色通道：'
         '行政审批、企业代办、申报协调等。'),
        ('机构专业运营',
         '引入专业潮玩产业服务机构，提供客户招聘、人力资源管理、人才培训、'
         '财务服务、科技金融等专业服务。'),
    ]
    for i, (a, b) in enumerate(triple, start=1):
        c0, c1 = tbl.rows[i].cells
        c0.width = Cm(3.6); c1.width = Cm(12.2)
        for c in (c0, c1):
            set_cell_border(c, top=True, bottom=True, color=RULE)
            shade_cell(c, 'FFFFFF' if i % 2 else BG_SOFT)
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(3); p0.paragraph_format.space_after = Pt(3)
        add_styled_run(p0, a, size=11, color=PRIMARY, bold=True)
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(3); p1.paragraph_format.space_after = Pt(3)
        p1.paragraph_format.line_spacing = 1.6
        add_styled_run(p1, b, size=10.5, color=TEXT)

    # ===================  第四章  ===================
    add_h1(doc, '四', '商业规划解读', 'Chapter 4 · Commercial Planning')

    add_h2(doc, '4.1  商业定位：Z·世代潮玩社交主场')
    add_para(doc,
             '项目提出三层递进式商业定位：',
             size=11, color=TEXT, line_spacing=1.85, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_bullet(doc, '功能层：', 'Z·世代潮玩社交主场——核心客群锚定 15–30 岁泛潮玩与二次元客群；')
    add_bullet(doc, '体验层：', '潮玩元宇宙 · 青年引力场——通过 IP、艺术、直播、Livehouse 营造可沉浸的潮玩世界；')
    add_bullet(doc, '社群层：', '潮玩艺术爱好者聚集地——形成具有圈层归属感的内容社区。')

    add_h2(doc, '4.2  楼层业态垂直落位')
    add_para(doc,
             '商业部分按"垂直主题集群"组织，六栋楼分别承担差异化角色，主力业态以"★"标记：',
             size=11, color=TEXT, line_spacing=1.85, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    tbl = doc.add_table(rows=8, cols=7)
    tbl.autofit = False
    widths = [Cm(1.8), Cm(2.3), Cm(2.3), Cm(2.3), Cm(2.3), Cm(2.3), Cm(2.3)]
    for i, w in enumerate(widths):
        tbl.columns[i].width = w
    add_table_header(tbl, ['楼层', '1#', '2#', '3#', '4#', '5#', '6#'])
    rows = [
        ('6F+', '森马总部办公', '—', '酒店', '潮玩产业办公', '潮玩产业办公', '—'),
        ('5F',  '森马总部办公', '—', '酒店', '潮玩产业办公', '潮玩产业办公', '商务宴请'),
        ('4F',  '★ 旗舰零售',   '二次元 Live', '休闲运动', '★ 直播基地', '动漫书店', '特色餐饮'),
        ('3F',  '★ 旗舰零售',   '二次元 Live', '萌宠空间', '★ 潮玩艺术', '动漫书店', '品质生活'),
        ('2F',  '★ 旗舰零售',   '二次元 Live', '休闲运动', '★ 潮玩艺术', 'IP 潮玩零售', '服务配套'),
        ('1F',  '★ IP 潮玩街区（六栋贯通）', '', '', '', '', ''),
        ('B1/B2', '停车场', '', '', '', '', ''),
    ]
    for i, r in enumerate(rows, start=1):
        for j, v in enumerate(r):
            c = tbl.rows[i].cells[j]
            c.width = widths[j]
            set_cell_border(c, top=True, bottom=True, color=RULE)
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            is_star = v.startswith('★')
            shade_cell(c, BG_SOFT2 if is_star else ('FFFFFF' if i % 2 else BG_SOFT))
            if j == 0:
                add_styled_run(p, v, size=10.5, color=PRIMARY, bold=True, font=EN)
            elif is_star:
                add_styled_run(p, v, size=9.5, color=PRIMARY, bold=True)
            elif v == '—':
                add_styled_run(p, v, size=10.5, color=MUTE)
            else:
                add_styled_run(p, v, size=9.5, color=TEXT)

    add_para(doc, '', after=4)
    add_callout(doc, '解读 · 商业空间的"叙事流"',
                '观察这张表可以发现一个隐性的"消费叙事流"：'
                '1F 全栋贯通的 IP 潮玩街区作为"入口与流量层"，吸引泛客群入场；'
                '2–4F 形成"主力业态层"，由旗舰零售、潮玩艺术、直播基地、二次元 Livehouse 等'
                '深度内容空间组成；5F 以上则是"产业层"，将商业流量与产业研发紧密耦合。'
                '这种"流量层—体验层—产业层"的垂直叙事结构是项目商业策划的最大创新点。',
                fill=BG_SOFT, accent=ACCENT)

    add_h2(doc, '4.3  五大主力业态')
    anchors = [
        ('动漫潮玩谷主题街区', '~3,000 ㎡',
         '南上海首个动漫潮玩主题商业街区，二次元品牌形象展示中心为意向品牌。'),
        ('IP 潮玩选品 & 仓储式零售中心', '~5,000 ㎡',
         '华东首个 IP 潮玩选品 + 仓储式零售中心，意向品牌：秋子 ACG 超级贩卖仓、'
         '轻语有品潮玩仓储超市、宏腾玩具选品中心。'),
        ('潮玩艺术中心', '2,000 ㎡（4# 1–3F）',
         '南上海首个特色艺术文化体验空间，依托特色空间和强展陈能力打造先锋潮流艺术地标。'),
        ('森马展厅 & 二次元 Livehouse', '700 ㎡',
         '国内首个二次元主题 Livehouse，意向品牌：绮丽少女。'),
        ('动漫主题书店', '1,500 ㎡',
         '南上海首个动漫主题书店，集全球动漫书籍展售、签售活动、主题咖啡于一体。'
         '意向品牌：超级悦沢、樱漫书店。'),
    ]
    for n, area, d in anchors:
        add_bullet(doc, f'{n}（{area}）：', d)

    add_callout(doc, '解读 · 多个"首个"的意义',
                '主力业态密集使用了"南上海首个 / 华东首个 / 国内首个"等定位语，'
                '本质上是用"首创性"换取流量稀缺性。在潮玩商业供给同质化日益严重的当下，'
                '这种由"地理首发权"+"业态首发权"组合形成的认知壁垒，'
                '是项目能在开业初期快速形成话题与客流的关键。',
                fill=BG_SOFT2, accent=PRIMARY)

    # ===================  第五章  ===================
    add_h1(doc, '五', '运营与合作策略', 'Chapter 5 · Operations & Cooperation')

    add_h2(doc, '5.1  三大旗舰活动')
    add_para(doc,
             '项目以"中心广场 + 4#"为活动主舞台，规划三类年度旗舰活动：',
             size=11, color=TEXT, line_spacing=1.85, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_bullet(doc, '潮玩集市（特色集会）：',
              '面向泛潮玩与二次元垂类客群，开展 IP 艺术展、次元市集、嘉年华、动漫展会、'
              '二次元路演、品牌出海交流会等活动，以垂类热度延展广域客群。')
    add_bullet(doc, '动漫新品及跨界产品首发会（产品首发）：',
              '聚焦年度全新动漫内容、原创 IP 形象、潮玩手办、周边衍生与跨界联名首发，'
              '集现场体验、互动打卡、合作签约于一体。')
    add_bullet(doc, '全国潮玩设计大赛（主题大赛）：',
              '围绕原创 IP 形象、手办盲盒、国潮文创、数字潮玩等方向展开创作，'
              '合作单位包括中国动漫集团、上海交大设计学院等。')

    add_callout(doc, '解读 · 活动—产业—商业三位一体',
                '这三类活动构成了项目的"内容飞轮"：大赛输出创意池 → 首发会将创意转化为产品 → '
                '集市完成最终的市场试销与扩散。三者形成的闭环，让项目从一个静态的物理空间，'
                '升级为一个具有"赛事 + 节庆 + 商务"三种节律的活内容平台。',
                fill=BG_SOFT3, accent=ACCENT2)

    add_h2(doc, '5.2  合作伙伴矩阵')
    tbl = doc.add_table(rows=4, cols=2)
    tbl.autofit = False
    tbl.columns[0].width = Cm(3.6)
    tbl.columns[1].width = Cm(12.2)
    add_table_header(tbl, ['类别', '意向 / 拟合作品牌'])
    matrix = [
        ('IP 潮玩 / 二次元',
         '绮丽少女、秋子 ACG 超级贩卖仓、轻语有品、宏腾玩具、超级悦沢、樱漫书店等'),
        ('产业合作机构',
         '中国百货协会、中国动漫集团、上海交大设计学院、闵行科协、森马集团、骏耀科技'),
        ('周边产业生态',
         '森马产业、衣恋、普洛斯、雅诗兰黛、保利、中建产研等'),
    ]
    for i, (a, b) in enumerate(matrix, start=1):
        c0, c1 = tbl.rows[i].cells
        c0.width = Cm(3.6); c1.width = Cm(12.2)
        for c in (c0, c1):
            set_cell_border(c, top=True, bottom=True, color=RULE)
            shade_cell(c, 'FFFFFF' if i % 2 else BG_SOFT)
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(3); p0.paragraph_format.space_after = Pt(3)
        add_styled_run(p0, a, size=11, color=PRIMARY, bold=True)
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(3); p1.paragraph_format.space_after = Pt(3)
        p1.paragraph_format.line_spacing = 1.6
        add_styled_run(p1, b, size=10.5, color=TEXT)

    # ===================  第六章  ===================
    add_h1(doc, '六', '项目价值评估与机会分析', 'Chapter 6 · Value Assessment')

    add_h2(doc, '6.1  项目核心价值六维评估')
    val = [
        ('战略卡位价值',  '南部科创走廊 + 大零号湾 + 长三角融合三位一体的战略身份，'
                          '决定项目天然具有"政策 + 资金 + 人才 + 流量"四重资源加持的可能。'),
        ('交通流量价值',  '15 号线元江路 TOD 日均 5–7 万人次的基础客流，'
                          '叠加 24 万居住 + 12 万办公人口的近域消费基数。'),
        ('物业规模价值',  '22 万㎡ 总建筑、5.2 万㎡ 商业、1,500+ 车位的体量，'
                          '足以承载多业态、多品牌、多 IP 并存。'),
        ('产业牌照价值',  '中百协与中国动漫集团双牌照的稀缺性，'
                          '是项目长期具备行业话语权的核心资产。'),
        ('技术赋能价值',  'AI 设计 + AI 打样 + 智能仓储 + 共享直播的全链路 AI 基础设施，'
                          '是潮玩中小企业最稀缺的能力。'),
        ('运营内容价值',  '集市 + 首发 + 大赛构成的"内容飞轮"，'
                          '为项目持续提供新闻、社交媒体话题、客流脉冲。'),
    ]
    for n, d in val:
        add_bullet(doc, f'{n}：', d)

    add_h2(doc, '6.2  机会窗口')
    add_para(doc,
             '从赛道、政策、城市三个维度审视，项目所处的时间窗口具备如下机会：',
             size=11, color=TEXT, line_spacing=1.85, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_bullet(doc, '赛道机会：', '中国潮玩市场仍处于"内容侧出海 + 国潮回流"的双向放大期，'
                                  '原创 IP 与 AI 工具尚有较大渗透空间；')
    add_bullet(doc, '政策机会：', '上海"南北转型"+"科技时尚特色小镇"双重政策窗口；')
    add_bullet(doc, '城市机会：', '15 号线全线贯通带来的元江路片区价值重估；')
    add_bullet(doc, '内容机会：', '二次元 Livehouse、潮玩艺术展、AI 共创等新内容空间需求强劲。')

    # ===================  第七章  ===================
    add_h1(doc, '七', '风险提示与建议', 'Chapter 7 · Risks & Recommendations')

    add_h2(doc, '7.1  风险提示')
    risks = [
        ('品牌招商风险',
         '潮玩头部品牌相对集中，意向品牌（秋子 ACG、绮丽少女、超级悦沢、樱漫等）'
         '签约转化与开业节奏需提前管控，避免业态空窗。'),
        ('内容运营风险',
         '集市 / 首发 / 大赛对内容策划与组织能力要求高，'
         '若缺少专业团队，活动易陷入"开业期热闹、长尾期冷清"的常见困境。'),
        ('客群匹配风险',
         '吴泾片区基础客群以居住与办公为主，与 Z 世代潮玩主力客群（高校 / 跨区潮流人群）'
         '存在一定错位，需依赖 15 号线 + 大学城联动持续引流。'),
        ('产业孵化风险',
         '60% 的中小型企业 / 服务机构存活率与产值贡献波动较大，'
         '需要明确的退出与汰换机制以及阶段性产业基金支持。'),
    ]
    for n, d in risks:
        add_bullet(doc, f'{n}：', d)

    add_h2(doc, '7.2  落地建议')
    add_bullet(doc, '锚定主力 IP：',
              '尽快锁定 2–3 个具有全国知名度的潮玩 / 二次元主力品牌作为首店招商主轴，'
              '为后续中小品牌跟投建立信号；')
    add_bullet(doc, '建立产业基金：',
              '依托双牌照与森马背景，联合区级 / 街道引导基金，'
              '设立"潮玩 + AI"种子轮基金，绑定入驻企业；')
    add_bullet(doc, '强化高校联动：',
              '与上海交大设计学院、闵行科协建立稳定的人才输送与赛事合作机制，'
              '形成"学校—竞赛—孵化器—产品"的人才闭环；')
    add_bullet(doc, '内容日历化：',
              '将三大旗舰活动 + 月度小型活动形成年度内容日历，'
              '提前 6 个月对外释放，构筑"潮玩界年度盛事"的认知；')
    add_bullet(doc, '数据资产化：',
              '把选品中心 / 直播中心 / 大赛报名沉淀的设计 / 销售 / 客流数据'
              '统一打通，形成项目专属的"潮玩产业数据底座"。')

    # ===================  结语  ===================
    add_h1(doc, '八', '结语', 'Chapter 8 · Closing Remarks')
    add_para(doc,
             '森马（上海）国际运营中心，本质上是一次面向 Z·世代的"产业 + 商业 + 文化"复合实验。'
             '从战略卡位到金字塔产业生态，从垂直业态到内容飞轮，'
             '项目展现出超越传统综合体的系统性思考与平台化野心。',
             size=11, color=TEXT, line_spacing=1.95, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_para(doc,
             '如果说传统商业综合体提供的是"场"，森马国际运营中心试图提供的，'
             '则是一个让 IP、AI、艺术、品牌、青年、生活在同一场域共生的"潮玩元宇宙"。'
             '若三大旗舰活动如期落地、双牌照充分激活、AI 配套全链路打通，'
             '项目将不仅成为闵行南部的城市新名片，更有望成为长三角潮玩产业的"枢纽级"标的。',
             size=11, color=TEXT, line_spacing=1.95, first_line_indent=0.74,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_horizontal_rule(doc, color=PRIMARY, width_pts=1.0)
    add_para(doc, '— END —',
             size=11, color=MUTE, font=EN,
             align=WD_ALIGN_PARAGRAPH.CENTER, before=8, after=4)
    add_para(doc, "森马（上海）国际运营中心 · Semir Group's Global Headquarter",
             size=10, color=MUTE, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
    add_para(doc, '本文档基于公开提案信息整理与解读，仅供内部研究与合作沟通使用。',
             size=9, color=MUTE, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)

    out_path = os.path.join(os.path.dirname(__file__), '..', 'docs',
                            '森马国际运营中心-项目介绍与解读.docx')
    out_path = os.path.normpath(out_path)
    doc.save(out_path)
    print(f'DOCX saved: {out_path}')


if __name__ == '__main__':
    build()
