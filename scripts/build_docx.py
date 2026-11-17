# -*- coding: utf-8 -*-
"""生成可提交学院讨论的合作框架 Word。"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from coop_data import (
    AI_PLAN,
    ANNUAL_GOALS,
    BUSINESSES,
    CHANNELS,
    CLIENT_ACTIONS,
    CLIENTS,
    COMPETITORS,
    DECISION_FACTORS,
    EXISTING_AIGC,
    FUNNEL,
    INTERVIEW_FIELDS,
    KPI,
    META,
    MILESTONES,
    MODULES,
    MOU_ITEMS,
    NEAR_TERM_A,
    NEAR_TERM_B,
    PRICE_LADDER,
    PRIORITY_VISITS,
    PRODUCTS,
    PROJECTS_2027,
    QUARTERS,
    RISKS,
    ROLES,
    ROADMAP,
    ROUTES,
    SCALE_EXAMPLE,
    SEGMENTS,
    SOURCES,
    VISIT_STATUS,
    VISIT_TIERS,
    VISITS,
)

CONTENT_CM = 17.2
NAVY = "13294B"
CRIMSON = "8C1D2C"
INK = "1C2430"
MUTED = "5E6A78"
GOLD = "8A6A32"


def _rfonts(r, font):
    r.font.name = font
    rpr = r._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), font)


def add_run(p, text, size=12, bold=False, color=INK, font="宋体", italic=False):
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    r.font.color.rgb = RGBColor.from_string(color)
    _rfonts(r, font)
    return r


def add_marked(p, text, size=12, color=INK, font="宋体"):
    parts = text.split("**")
    for i, part in enumerate(parts):
        if part:
            add_run(p, part, size=size, bold=i % 2 == 1, color=color, font=font)


def _shade(cell, fill):
    tc = cell._tc
    tcpr = tc.get_or_add_tcPr()
    shd = tcpr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcpr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def _set_cell_border(cell, color="E4DDD2"):
    tc = cell._tc
    tcpr = tc.get_or_add_tcPr()
    tc_borders = tcpr.find(qn("w:tcBorders"))
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tcpr.append(tc_borders)
    for edge in ("top", "left", "bottom", "right"):
        el = tc_borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            tc_borders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)


def _cell_margins(cell):
    tc = cell._tc
    tcpr = tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for edge, val in (("top", "60"), ("left", "80"), ("bottom", "60"), ("right", "80")):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:w"), val)
        node.set(qn("w:type"), "dxa")
        mar.append(node)
    tcpr.append(mar)


def _widths(table, widths):
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl = table._tbl
    tblpr = tbl.tblPr
    tblw = tblpr.find(qn("w:tblW"))
    if tblw is None:
        tblw = OxmlElement("w:tblW")
        tblpr.append(tblw)
    total = sum(widths)
    tblw.set(qn("w:w"), str(total))
    tblw.set(qn("w:type"), "dxa")
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tblpr.addnext(grid)
    for child in list(grid):
        grid.remove(child)
    for w in widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        grid.append(gc)
    for row in table.rows:
        tr = row._tr
        trpr = tr.get_or_add_trPr()
        if trpr.find(qn("w:cantSplit")) is None:
            trpr.append(OxmlElement("w:cantSplit"))
        for i, cell in enumerate(row.cells):
            tc = cell._tc
            tcpr = tc.get_or_add_tcPr()
            tcw = tcpr.find(qn("w:tcW"))
            if tcw is None:
                tcw = OxmlElement("w:tcW")
                tcpr.append(tcw)
            tcw.set(qn("w:w"), str(widths[i]))
            tcw.set(qn("w:type"), "dxa")


def _header_row(row):
    trpr = row._tr.get_or_add_trPr()
    if trpr.find(qn("w:tblHeader")) is None:
        trpr.append(OxmlElement("w:tblHeader"))


def cm_to_dxa(values):
    return [int(v * 567) for v in values]


def prevent_row_split_ok():
    return None


def set_keep(p):
    p.paragraph_format.keep_with_next = True


def style_run_font(style, font, size, color, bold):
    style.font.name = font
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.color.rgb = RGBColor.from_string(color)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), font)


def setup(doc):
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.9)
    section.right_margin = Cm(1.9)
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(1.9)
    section.header_distance = Cm(0.7)
    section.footer_distance = Cm(0.5)
    section.different_first_page_header_footer = True

    normal = doc.styles["Normal"]
    style_run_font(normal, "宋体", 12, INK, False)
    normal.paragraph_format.line_spacing = 1.25
    normal.paragraph_format.space_after = Pt(6)

    for name, size, before, after, color in (
        ("Heading 1", 16, 16, 8, CRIMSON),
        ("Heading 2", 14, 14, 6, NAVY),
        ("Heading 3", 12, 10, 4, NAVY),
    ):
        st = doc.styles[name]
        style_run_font(st, "黑体", size, color, True)
        st.paragraph_format.space_before = Pt(before)
        st.paragraph_format.space_after = Pt(after)
        st.paragraph_format.line_spacing = 1.15
        st.paragraph_format.keep_with_next = True
        st.paragraph_format.first_line_indent = Cm(0)

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    hp.paragraph_format.space_after = Pt(2)
    add_run(hp, f"{META['party_a']}  ×  {META['party_b']}", 9, False, MUTED, "微软雅黑")
    add_run(hp, "          合作洽谈稿", 9, False, CRIMSON, "微软雅黑")
    ppr = hp._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "C4A35A")
    pbdr.append(bottom)
    ppr.append(pbdr)

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_run(fp, "非正式招生简章  ·  未经审批不得对外承诺品牌、证书、学费和收益", 8, False, MUTED, "微软雅黑")
    add_run(fp, "          ", 8, False, MUTED, "微软雅黑")
    _page_field(fp)

    core = doc.core_properties
    core.title = "复旦大学终身教育学院与上海市杨浦区科技企业联合会2026—2027年战略合作框架方案"
    core.author = META["compiler"]
    core.subject = META["nature"]
    core.category = "合作洽谈稿"
    core.comments = "产品、价格、证书和分成均未获学院批准。"


def _page_field(paragraph):
    def make_run():
        node = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        rfonts = OxmlElement("w:rFonts")
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            rfonts.set(qn(attr), "微软雅黑")
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), "16")
        color = OxmlElement("w:color")
        color.set(qn("w:val"), MUTED)
        rpr.append(rfonts)
        rpr.append(sz)
        rpr.append(color)
        node.append(rpr)
        return node

    begin_run = make_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin_run.append(begin)

    instr_run = make_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    instr_run.append(instr)

    end_run = make_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    end_run.append(end)

    paragraph._p.append(begin_run)
    paragraph._p.append(instr_run)
    paragraph._p.append(end_run)


def h1(doc, text):
    p = doc.add_heading(text, level=1)
    for run in p.runs:
        _rfonts(run, "黑体")
        run.font.color.rgb = RGBColor.from_string(CRIMSON)
    return p


def h2(doc, text):
    p = doc.add_heading(text, level=2)
    for run in p.runs:
        _rfonts(run, "黑体")
        run.font.color.rgb = RGBColor.from_string(NAVY)
    return p


def h3(doc, text):
    p = doc.add_heading(text, level=3)
    for run in p.runs:
        _rfonts(run, "黑体")
        run.font.color.rgb = RGBColor.from_string(NAVY)
    return p


def body(doc, text, indent=True, size=12, space=8):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(space)
    pf.line_spacing = 1.25
    if indent:
        pf.first_line_indent = Cm(0.74)
    add_marked(p, text, size=size)
    return p


def bullet(doc, text, level=0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.75 + level * 0.5)
    pf.first_line_indent = Cm(-0.35)
    pf.space_after = Pt(3)
    pf.space_before = Pt(0)
    pf.line_spacing = 1.2
    add_run(p, "·  ", 12, True, CRIMSON, "宋体")
    add_marked(p, text, size=12)
    return p


def caption(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    add_run(p, text, 10, True, NAVY, "微软雅黑")
    return p


def note(doc, text):
    widths = cm_to_dxa([CONTENT_CM])
    table = doc.add_table(rows=1, cols=1)
    _widths(table, widths)
    cell = table.cell(0, 0)
    _shade(cell, "F8F1F2")
    _set_cell_border(cell, "E4C8CE")
    _cell_margins(cell)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    add_marked(p, text, size=10.5, color="5C2430")
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(4)
    spacer.paragraph_format.space_before = Pt(2)
    return table


def add_table(doc, headers, rows, widths_cm, font=9.5, center_cols=()):
    scale = CONTENT_CM / sum(widths_cm)
    widths = cm_to_dxa([w * scale for w in widths_cm])
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    _widths(table, widths)
    _header_row(table.rows[0])
    for i, text in enumerate(headers):
        cell = table.rows[0].cells[i]
        _shade(cell, NAVY)
        _set_cell_border(cell, NAVY)
        _cell_margins(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        add_run(p, text, font, True, "FFFFFF", "微软雅黑")
    for r, row in enumerate(rows):
        for c, text in enumerate(row):
            cell = table.rows[r + 1].cells[c]
            _shade(cell, "F8F6F2" if r % 2 else "FFFFFF")
            _set_cell_border(cell, "E4DDD2")
            _cell_margins(cell)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c in center_cols else WD_ALIGN_PARAGRAPH.LEFT
            add_run(p, str(text), font, False, INK, "微软雅黑")
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(2)
    spacer.paragraph_format.space_after = Pt(4)
    return table


def page_break(doc):
    doc.add_page_break()


def cover(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(36)
    add_run(p, "合作洽谈稿", 12, True, CRIMSON, "微软雅黑")
    add_run(p, "    |    " + META["doc_date"], 12, False, MUTED, "微软雅黑")

    t = doc.add_paragraph()
    t.paragraph_format.space_before = Pt(28)
    t.paragraph_format.space_after = Pt(0)
    t.paragraph_format.line_spacing = 1.15
    add_run(t, "复旦大学终身教育学院", 26, True, NAVY, "黑体")

    t2 = doc.add_paragraph()
    t2.paragraph_format.space_before = Pt(4)
    t2.paragraph_format.space_after = Pt(0)
    add_run(t2, "与上海市杨浦区科技企业联合会", 22, True, NAVY, "黑体")

    t3 = doc.add_paragraph()
    t3.paragraph_format.space_before = Pt(16)
    t3.paragraph_format.space_after = Pt(8)
    add_run(t3, "2026—2027年战略合作框架方案", 20, True, CRIMSON, "黑体")

    line = doc.add_paragraph()
    line.paragraph_format.space_after = Pt(12)
    ppr = line._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "18")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "C4A35A")
    pbdr.append(bottom)
    ppr.append(pbdr)

    body(doc, "以学院的教育、师资和培训交付为核心，以联合会的产业资源、企业客户、市场渠道和活动运营为支撑，共同建设面向科技企业、央国企、政府园区及企业家的非学历教育与产业交流平台。", indent=False, size=12)
    body(doc, "委托培训形成业务基础，产业参访形成差异，科创人才短训形成增长空间。", indent=False, size=12)

    meta_rows = [
        ("文件性质", "合作洽谈稿。供学院内部讨论，不是招生简章，也不构成学校授权。"),
        ("编制单位", META["compiler"]),
        ("对接单位", META["party_a"] + "（拟）"),
        ("编制日期", META["compiled"]),
        ("主要依据", "2026年11月6日两份洽谈纪要，2026年11月10日科创培训项目交流。"),
        ("建议期限", "首期一年，非排他试点。备忘录写意向，专项合同写费用和交付。"),
    ]
    add_table(doc, ["项目", "说明"], meta_rows, [3.2, 14.0], font=10.5)
    note(doc, "本文件中的课程名称、价格、班额、分成和证书，均为建议或测算。正式合作以学校审批、合作备忘录和项目实施合同为准。未经审批，不得对外承诺招生、证书、品牌或收益。")


def guide(doc):
    h1(doc, "使用说明")
    body(doc, "2026年11月6日、11月10日的三份洽谈，已经把合作方向说清楚了。本方案把会议共识和联合会的实施建议放在同一份文件里，方便学院一次看完。")
    h2(doc, "三类内容")
    bullet(doc, "**会议共识。** 合作定位、学院与联合会的能力互补、企业培训和参访是近期抓手、先签备忘录再推进项目。")
    bullet(doc, "**实施建议。** 六项业务、三类合同模式、三门优先产品、研修班参数、2027年项目表和管理机制。这些内容需要双方确认后才能执行。")
    bullet(doc, "**尚未核准。** 内部管理费比例、税务和票据安排、个人账户、学历学位、校友权益、校名和标识的对外使用。这些事项不写入正式合作文件。")
    h2(doc, "四份材料如何对应")
    body(doc, "联合会建议下次沟通同时带上三项成果：合作备忘录草案、企业参访资源菜单、人工智能与科创企业家研修班方案。本 Word 是完整文本；同目录的演示文稿用于当面汇报；工作表用于筛选客户、核实参访和重算假设。")
    bullet(doc, "正文第一至十一章，是合作框架。")
    bullet(doc, "附件一至附件四，分别回答有什么参访资源、有什么客户、研修班如何招生、2027年怎么排。")
    bullet(doc, "工作表中名为“内部跟进”的表含待填写的联系人和金额，提交学院时不要附上。")
    h2(doc, "建议的读法")
    body(doc, "时间有限时，先看第七章的三个试点、第十一章的三个谈判重点，以及附件一的十二个优先参访项目。其余章节供会后核对边界、报价和审批路径。")


def contents(doc):
    h1(doc, "目录")
    items = [
        "一、合作定位与总体目标",
        "二、合作主体及职责分工",
        "三、六大重点合作业务",
        "四、三类商业合作模式",
        "五、联合运营与项目管理机制",
        "六、2026—2027年合作实施计划",
        "七、首批优先推进项目",
        "八、合作备忘录建议明确的事项",
        "九、2027年度合作成效评价",
        "十、近期双方工作安排",
        "十一、三次洽谈中最值得把握的谈判重点",
        "附件一  科技企业标杆参访资源清单",
        "附件二  首批企业培训客户及需求线索",
        "附件三  人工智能企业家研修班市场需求分析",
        "附件四  2027年度拟合作项目计划",
        "资料依据",
    ]
    for i, item in enumerate(items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.space_before = Pt(1)
        add_run(p, f"{i:02d}    ", 12, True, CRIMSON, "微软雅黑")
        add_run(p, item, 12, False, NAVY, "微软雅黑")
    body(doc, "打开文件后可使用 Word 或 WPS 的导航窗格跳转。页眉页脚从第二页起显示。", indent=False, size=10.5)


def framework(doc):
    h1(doc, "一、合作定位与总体目标")
    h2(doc, "1. 合作定位")
    body(doc, "双方拟建立长期、稳定、互利的合作关系，围绕科技创新、产业升级、企业人才培养和产教融合，探索“高校教育资源 + 产业组织平台 + 市场化运营”。")
    body(doc, "学院提供课程体系、专家师资、教育管理和培训交付。联合会提供产业生态、企业渠道、区域政府资源和活动运营。双方共同覆盖人才培养、产业参访、企业赋能、智库研究和科技交流。")
    h2(doc, "2. 合作目标")
    bullet(doc, "**近期，2026年底。** 建立对接机制，签署合作备忘录，完成企业参访资源库和首批课程方案。")
    bullet(doc, "**中期，2027年。** 推动企业委托培训、人工智能创新人才研修班和产业参访落地。")
    bullet(doc, "**长期，2027年以后。** 形成具备持续招生能力和产业特色的科创人才培养机制，再扩展长三角及全国的企业、园区和政府委托项目。")
    h2(doc, "3. 合作原则")
    body(doc, "优势互补、专业分工、项目化运作、依法合规、市场导向、成果共享。具体合作以学校审批、双方正式协议和项目实施合同为准。")

    h1(doc, "二、合作主体及职责分工")
    h2(doc, "甲方：复旦大学终身教育学院（拟）")
    for item in (
        "教育培训项目的论证、立项及校内审批。",
        "课程体系设计、教学质量管理与师资组织。",
        "校内教学资源、场地及学员服务协调。",
        "经批准的结业考核与证书发放。",
        "政府、国央企委托培训项目的教学组织和交付。",
        "合作项目的教育品牌和教学标准管理。",
    ):
        bullet(doc, item)
    h2(doc, "乙方：上海市杨浦区科技企业联合会（拟）")
    for item in (
        "科技企业、金融机构、产业园区及政府客户拓展。",
        "企业培训需求收集、市场调研与项目引荐。",
        "社群渠道、论坛活动和公开课市场组织。",
        "企业参访资源开发、商务协调与活动服务。",
        "产业专家、企业家及应用场景资源对接。",
        "项目宣传推广、客户维护及经批准的配套服务。",
    ):
        bullet(doc, item)
    h2(doc, "项目运营合作方")
    body(doc, "具体项目可以根据需要引入专业课程策划、招生推广、班务服务、活动执行或企业参访机构。赵海等已有课程产品和运营经验的人员，可作为后续课程开发或项目运营的候选合作人员，不当然取得校方岗位或项目授权。")
    note(doc, "基本分工：学院主导教学，联合会侧重市场与产业资源。学院方面由刘老师协调院内对接，具体授权和负责人以正式指定为准。联合会方面由执行会长牵头，协调朱总及相关工作人员。运营事项在项目协议里写明。")

    h1(doc, "三、六大重点合作业务")
    body(doc, "六项业务同时规划，不同时铺开。最高优先级里先做委托培训、研修班和参访。特色项目等主线成熟后再培育。")
    add_table(
        doc,
        ["序号", "业务", "对象", "优先级", "产品形态"],
        [b for b in BUSINESSES],
        [1.4, 2.4, 4.6, 3.2, 5.6],
        font=9,
        center_cols=(0, 3),
    )

    h2(doc, "方向一：央国企及科技企业委托培训")
    body(doc, "**优先级最高。** 面向央国企、金融机构、科技企业、专精特新企业和产业园区，做定制化培训。")
    body(doc, "重点主题包括：生成式人工智能与企业数字化转型；企业战略管理、公司治理与组织创新；科技成果转化、技术商业化与资本运作；产业投资、并购重组与融资实务；先进制造、具身智能与新质生产力；国央企科技创新与产业转型。")
    body(doc, "联合会对接企业培训部门并带回需求。学院按需求设计课程、组织师资、实施培训。双方按实际承担的服务签订项目协议。优先了解的客户包括国泰海通、国药控股、银行保险机构、央国企和杨浦区重点科技企业。")
    body(doc, "建议产品为 3—5 天企业高管培训班、专题研修班和定制化企业内训。")

    h2(doc, "方向二：科创企业家与人工智能人才研修班")
    body(doc, "**优先级最高。** 联合开发面向科技企业创始人、董事长、总经理、CTO 及核心管理团队的非学历短期课程。建议首发名称用《" + AI_PLAN["name"] + "》，最终名称由学院审批确定。")
    body(doc, "课程覆盖人工智能产业趋势，大模型、智能体和企业应用，科技创新与成果转化，创业融资、估值与资本市场，公司战略、组织与人才，产业政策、园区与企业成长，标杆企业现场教学，以及企业家交流。")
    body(doc, f"建议学制 {AI_PLAN['months']}，每班 {AI_PLAN['class_size']}，周末集中授课，并安排专题讲座、企业参访和实践交流。定价可参照洽谈中提出的数千元至 2 万元区间，在测算教学成本、管理费用和招生规模后确定。详细市场分析见附件三。")
    note(doc, "结业证书和非学历培训相关身份权益，依据学校正式批准的项目规则确定。不得承诺学历学位，也不得承诺超出校方规定的校友权益。")

    h2(doc, "方向三：标杆企业参访与长三角产业游学")
    body(doc, "**优先级最高，建议作为最先启动的产品。** 11月10日的交流已经提出，由联合会整理参访资源菜单；符合价格体系后，可以开展委托合作。这是目前最具体、最容易先做起来的动作。")
    body(doc, "联合会建设资源库，学院按培训项目需要安排现场教学。产品包括半日参访、一日专题和 3—5 天长三角产业研学，并可研究外籍企业家及国际商务团组。上海—苏州—南京—杭州等线路作为组合建议。完整目录、状态和服务标准见附件一。")
    note(doc, "名单中的企业是拟对接清单。接待档期、人数、费用和讲解，都以企业重新确认的结果为准。")

    h2(doc, "方向四：公开课、产业论坛与市场化招生")
    body(doc, "**优先级高。** 利用联合会现有论坛、企业社群和活动组织能力，为学院非学历项目组织市场。路径是：公开课引流，专题沙龙，课程咨询，正式报名，培训交付，学员持续服务。")
    bullet(doc, "每年联合策划 1—2 场有产业影响力的科创教育交流活动。")
    bullet(doc, "每季度开展人工智能、产业投资、企业管理等主题公开课或沙龙。")
    bullet(doc, "通过行业协会、园区、企业客户和校友资源拓展合规招生渠道。")
    body(doc, "论坛可以和课程推广放在同一场活动里。报名、学费收取和招生宣传遵守学校规定。学院负责课程产品和教学审核。联合会在取得授权后，提供市场推广、活动组织和渠道服务。")

    h2(doc, "方向五：地方政府委托培训与产业咨询")
    body(doc, "**优先级中高。** 围绕地方产业人才培养、科技招商、产业升级和企业治理，对接政府部门、产业园区和驻沪办事机构。")
    for item in (
        "地方科技型企业经营管理人才培训。",
        "专精特新企业创新能力提升培训。",
        "新质生产力和人工智能产业专题研修。",
        "地方招商引资与产业投资能力培训。",
        "国企改革与科技成果转化培训。",
        "园区企业经营诊断及产业发展专题咨询。",
    ):
        bullet(doc, item)
    body(doc, "可探索“政府组织企业、学院提供培训、联合会协调产业资源”。研究咨询必须写明承担主体和工作范围。涉及采购、招投标的，依法履行程序，不预设财政补贴。")

    h2(doc, "方向六：银发经济、国学文化与特色研修")
    body(doc, "**中期培育。** 主线业务成熟后，再做具有复旦人文特色和产业前景的专题。")
    bullet(doc, "**银发经济与智慧康养。** 养老产业经营管理、长期护理保险、智慧养老、医疗健康服务创新。对象是养老机构管理者、保险机构及康养企业。")
    bullet(doc, "**国学哲学与企业家领导力。** 把中国哲学、管理思想、组织文化和企业家成长做成非学历讲座或文化研修。")
    bullet(doc, "**国际交流。** 在现有国际企业参访和海外游学经验上，研究经批准的国际研修与交流。")
    note(doc, "国学研究院等机构化设想应作为单独议题研究，不认定为本次合作已经确定的事项。")

    h1(doc, "四、三类商业合作模式")
    h2(doc, "模式 A：企业或政府直接委托学院")
    body(doc, "适用于具备直接签约条件的央国企、政府机关、金融机构及其他客户。这是优先考虑的模式。")
    bullet(doc, "委托单位与校方授权主体签署合同。")
    bullet(doc, "学院负责教学、项目交付和统一收费管理。")
    bullet(doc, "联合会承担经书面约定的市场调研、产业参访或运营服务。")
    bullet(doc, "服务费依据实际工作、交付成果、合同约定及相关审批支付。")
    h2(doc, "模式 B：联合会整合需求并委托学院")
    body(doc, "适用于联合会具备独立组织需求、合法签约资质和真实采购关系的企业培训项目。")
    bullet(doc, "联合会与客户签订服务合同。")
    bullet(doc, "教学部分按学校规定采购或委托学院实施。")
    bullet(doc, "联合会负责自己承担的市场、会务、产业协调和配套服务。")
    bullet(doc, "双方根据真实服务内容分别核算、签约和结算。")
    note(doc, "模式 B 以学校允许相应合作主体和合同结构为前提。不采用虚构委托主体、虚假服务项目规避学校审批。")
    h2(doc, "模式 C：联合项目开发与专业服务采购")
    body(doc, "适用于双方联合研发的短训、公开课和产业参访。")
    bullet(doc, "立项前明确课程和市场定位。")
    bullet(doc, "分别核算教学、招生服务、场地、参访、物料和宣传成本。")
    bullet(doc, "按实际服务投入及风险承担确定合同价格。")
    bullet(doc, "由合法授权主体收取相应款项，按合同验收结算。")
    bullet(doc, "项目盈利安排须经双方财务和相关职能部门认可。")
    note(doc, "不把学院内部管理费比例，或会议中的分成设想，写成已经批准的合作分成。")

    h1(doc, "五、联合运营与项目管理机制")
    bullet(doc, "**联合工作小组。** 双方各指定一名项目负责人，并配置市场、教学、产业资源和财务对接人员。")
    bullet(doc, "**项目负责人制度。** 每个项目走完需求确认、立项审批、成本预算、实施交付、验收结算和效果复盘。")
    bullet(doc, "**项目共享清单。** 包括企业客户、课程目录、师资方向、参访资源、招生渠道和项目台账。客户个人信息、商业秘密和学校内部资料，遵守保密、授权和数据保护要求。")
    bullet(doc, "**月度沟通。** 每月一次业务推进会，按季度复盘，决定新增课程、重点客户和资源投入。")
    bullet(doc, "**品牌授权与质量管理。** 课程宣传、复旦大学名称与标识、证书、校友身份表述和校园参访，均以正式授权及管理规定为准。")

    h1(doc, "六、2026—2027年合作实施计划")
    add_table(doc, ["阶段", "重点任务", "预期成果"], ROADMAP, [3.2, 6.8, 7.2], font=9.5)
    body(doc, "以上时间为建议排期。项目立项、招生情况和学校审批进度优先于这张表。")

    h1(doc, "七、首批优先推进项目")
    body(doc, "先集中资源做三个项目。六项业务的其余部分进入年度计划，不在第一轮同时启动。")
    h3(doc, "项目一：科技企业标杆参访服务")
    body(doc, "联合会整理首批 10—20 家候选机构的参访信息，核实接待能力、服务费用和预约条件。学院结合现有委托培训项目试用。")
    h3(doc, "项目二：央国企委托培训项目")
    body(doc, "优先对接国泰海通、国药控股，或已建立联系的金融机构培训部门。以客户的具体需求为基础，由学院形成课程建议方案。")
    h3(doc, "项目三：人工智能与科技创新企业领军人才研修班")
    body(doc, "学院牵头形成课程方案、师资结构和预算。联合会测试企业家的招生需求。完成校内立项和正式招生审批后，再实施首期培训。")
    note(doc, "三家重点客户里，只要有一家形成真实培训需求，备忘录和后续业务就有了落脚点。不建议一次性联系全部二十二个对象。")

    h1(doc, "八、合作备忘录建议明确的事项")
    body(doc, "正式备忘录建议至少写明下列事项。首期合作期限设为一年，以非排他方式试点。备忘录明确合作意向；费用、品牌、招生承诺和业务交付，通过后续专项合同落实。")
    for i, item in enumerate(MOU_ITEMS, 1):
        bullet(doc, f"{i}. {item}")

    h1(doc, "九、2027年度合作成效评价")
    body(doc, "试点期同时看项目、客户、教学、资金、资源和品牌，不以签约金额作为唯一标准。")
    add_table(doc, ["指标", "看什么"], KPI, [3.2, 14.0], font=10.5)

    h1(doc, "十、近期双方工作安排")
    h2(doc, "联合会近期提交")
    for item in NEAR_TERM_B:
        bullet(doc, item)
    body(doc, "本方案及同套演示文稿、工作表，就是这四份材料的合订稿。")
    h2(doc, "学院近期协调")
    for item in NEAR_TERM_A:
        bullet(doc, item)
    body(doc, "双方完成资料交换后，再确定正式备忘录文本和第一批试点项目。")

    h1(doc, "十一、三次洽谈中最值得把握的谈判重点")
    h2(doc, "第一，联合会要证明的是转化能力")
    body(doc, "学院目前有相对成熟的 B 端委托培训交付体系，同时在探索 C 端招生和新产品。联合会的价值，是市场开拓、客户转化和产业组织。只介绍几家企业，支撑不起长期合作。")
    h2(doc, "第二，参访菜单是最近的行动")
    body(doc, "11月10日，学院明确要求整理参访资源菜单，并表示符合价格体系后可以开展委托合作。这是三份纪要里最容易落实的近期行动，不宜只停留在战略合作的表述上。")
    h2(doc, "第三，未核准的财务和证书安排不进正式文件")
    body(doc, "原始谈话里提到的内部管理费、税务减免、合同路径和证书权益，都还没有正式核准。尤其是以不同票据、个人账户或费用拆分调整利润的想法，不写入正式合作文件。")
    note(doc, "可执行的方式是：真实服务、真实合同、依法开票、校方审批。下次与刘老师及学院负责人沟通时，带上备忘录草案、参访资源菜单和研修班产品方案。这三份材料分别对应资源、产品和执行。")


def annex_visits(doc):
    page_break(doc)
    h1(doc, "附件一  科技企业标杆参访资源清单")
    body(doc, f"编制单位：{META['compiler']}。合作对接单位：{META['party_a']}（拟）。编制日期：{META['compiled']}。本附件是 2026—2027 年度产业参访与现场教学的候选目录，第一版。")
    body(doc, "清单服务学院的政府委托培训、央国企培训、科技企业高管研修和国际交流。现场教学采用企业实地参访、专题讲解、产业交流、案例教学和管理层座谈，形成“理论教学 + 产业现场 + 案例研讨 + 实践交流”。")
    h2(doc, "资源状态")
    for key, text in VISIT_STATUS.items():
        bullet(doc, f"**{key}。** {text}。")
    note(doc, "所有企业的接待档期、人数、费用和讲解服务，以实际确认结果为准。本清单不是接待承诺函，也不是学校授权文件。")
    caption(doc, "表 1  候选参访资源")
    add_table(
        doc,
        ["序号", "产业", "企业或平台", "建议主题", "状态"],
        [(a, b, c, d, e) for a, b, c, d, e in VISITS],
        [1.3, 3.0, 4.2, 6.5, 1.6],
        font=8,
        center_cols=(0, 4),
    )
    body(doc, "单家企业建议 1.5—2 小时。每半天安排 1 家核心企业或展示平台。机器人方向优先组织“张江机器人谷 + 代表企业”。银发方向可与傅利叶智能等企业组成“养老服务现场 + 康复设备展示”。")
    body(doc, "医疗和研发企业按许可范围安排。研发实验室、生产环境、知识产权和商业秘密区域，不列为默认开放内容。厂区、生产线和保密设施，只在企业批准范围内实施。")
    body(doc, "杭州“六小龙”、理想汽车常熟相关资源和蔚来合肥相关资源，可用于设计 3—5 天专题研学。此前交流提到过部分历史参访经验，对接名单和接待条件仍须逐一核实。")
    h2(doc, "首批十二个重点项目")
    body(doc, "综合课程匹配、产业代表性和开发难度，建议先核实下表。第一批不等于已经获得参访确认。")
    add_table(doc, ["批次", "参访资源", "首选教学方向"], PRIORITY_VISITS, [2.4, 6.4, 8.4], font=10, center_cols=(0,))
    h2(doc, "五条精品线路")
    add_table(doc, ["线路", "主题", "组合", "适合对象"], ROUTES, [2.0, 3.6, 6.2, 5.4], font=9)
    body(doc, "线路是教学组合建议。接待顺序和交通，按实际预约确定。")
    h2(doc, "三档服务与报价")
    add_table(doc, ["服务类型", "建议时长", "建议配置"], VISIT_TIERS, [3.2, 3.4, 10.6], font=10)
    body(doc, "费用按企业接待成本、讲解、人数、交通和会务确定。联合会向学院提供项目制报价，分开列示企业接待费用、交通会务支出和资源协调服务费用，并提供服务清单、合同和合规结算材料。已经打进培训合同的参访，不重复计算为独立收入。")
    h2(doc, "实施步骤")
    bullet(doc, "双方确认首批 12 个资源及课程适配方向。")
    bullet(doc, "联合会逐项核实接待窗口、开放时间、人数、费用和具体服务。")
    bullet(doc, "形成带企业简介、参访模块和报价的《首批科技企业现场教学项目手册》。")
    bullet(doc, "选择 1—2 个学院既有培训项目试点，再按交付情况修订服务标准。")
    h3(doc, "联合会")
    for item in (
        "建立并动态更新企业参访资源库。",
        "按学院需求联系企业，核实接待条件。",
        "提供项目介绍、接待方案和费用报价。",
        "协调企业接待、行程衔接和经批准的服务事项。",
        "开发长三角产业研学和特色现场教学产品。",
    ):
        bullet(doc, item)
    h3(doc, "学院")
    for item in (
        "提供培训项目需求及学员群体信息。",
        "明确教学目标、计划人数和预算。",
        "评估参访内容与课程是否匹配。",
        "负责培训项目的整体教学组织和质量管理。",
        "按实际需求确认委托或采购安排。",
    ):
        bullet(doc, item)


def annex_clients(doc):
    page_break(doc)
    h1(doc, "附件二  首批企业培训客户及需求线索")
    body(doc, f"编制单位：{META['compiler']}。编制日期：{META['compiled']}。本表根据前期洽谈、已有企业联系和拟开展的培训业务整理，供后续访谈、课程设计和项目合作使用。")
    note(doc, "客户需求、预算、规模和采购意向均待确认。未获得的信息填“待确认”。行业里常见的培训主题和估算预算，不能写成客户已经确认的采购意向。本表故意不填写未经确认的负责人姓名、电话和预算。内部执行再使用工作表中的“内部跟进”，该表不要随本文件外发。")
    body(doc, "优先级分三档。P0 是会议中已经提到可以直接认识培训部门，或有较明确引荐条件的对象。P1 有企业或行业渠道，仍需核实决策人和真实需求。P2 具有行业匹配性，尚未形成具体项目。政府、园区和行业组织更多是项目组织方、推荐渠道或潜在采购方，不宜全部写成已经到手的培训客户。")
    caption(doc, "表 2  客户与渠道线索")
    add_table(
        doc,
        ["序号", "对象", "类型", "关系基础", "拟匹配需求", "优先级"],
        CLIENTS,
        [1.2, 3.3, 2.2, 3.8, 4.5, 1.6],
        font=7.5,
        center_cols=(0, 5),
    )
    h2(doc, "第一轮只推进三家")
    body(doc, "国泰海通、国药控股和中远相关企业，是会议里明确提到与培训部门有联系的线索，放在最前面。招商银行、交通银行、地方政府和产业园区作为第二梯队。")
    bullet(doc, "**国泰海通证券。** 先问 2027 年培训计划、预算周期、科技金融和 AI 培训需求。")
    bullet(doc, "**国药控股。** 围绕医药经营管理、组织发展和数字化转型，了解集团或所属单位是否有班次。")
    bullet(doc, "**中远相关企业。** 先核实对应的是集团、下属公司还是培训机构，再确定专题。主体不清，不设计具体方案。")
    body(doc, "建议 2026 年 11 月至 12 月完成第一轮沟通，2027 年第一季度按反馈选择试点。三家培训负责人可以先引荐给刘老师。")
    h2(doc, "首轮动作")
    add_table(doc, ["客户或渠道", "首轮事项", "分工", "目标成果"], CLIENT_ACTIONS, [3.2, 5.4, 4.4, 4.2], font=9)
    h2(doc, "可对应的课程产品")
    body(doc, "下表的周期是初步设计参数。授课、招生规模、报价和证书，由学院按项目确定。")
    add_table(doc, ["产品", "目标客户", "周期", "组织形式"], PRODUCTS, [5.6, 4.4, 2.6, 4.6], font=9)
    h2(doc, "访谈至少采集这些字段")
    add_table(doc, ["字段", "填写要求"], INTERVIEW_FIELDS, [3.4, 13.8], font=10)
    h2(doc, "首期三个结果")
    bullet(doc, "形成一批真实培训需求。优先拿到国泰海通、国药控股和中远相关企业的部门、主题、人数和采购周期。")
    bullet(doc, "形成一门可以讨论招生的标准课程，重点是人工智能与科创企业家研修班。")
    bullet(doc, "形成一条可重复的委托培训流程：需求收集、课程设计、项目审批、签约采购、教学交付、费用结算。")
    body(doc, "2026 年底前，双方共同筛选首批项目，并形成 2027 年度合作计划。所有事项以客户实际需求、学院审批和正式合同为准。")


def annex_ai(doc):
    page_break(doc)
    h1(doc, "附件三  人工智能企业家研修班市场需求分析")
    body(doc, f"项目暂定名：《{AI_PLAN['name']}》。编制单位：{META['compiler']}。编制日期：{META['compiled']}。本附件是市场调研和招生可行性分析，第一版，不是招生简章。")
    note(doc, f"核心建议：首期定位为“企业家 AI 战略与落地实践”。小班 {AI_PLAN['class_size']}，学制 {AI_PLAN['months']}，企业定向招生为主，公开课和私域为辅。建议测试价 {AI_PLAN['price']:,} 元/人，不是学院核准收费。")
    h2(doc, "1. 为什么现在做，以及不做什么")
    body(doc, "企业管理者对人工智能的学习需求，正在从工具使用转向战略决策、组织变革和商业价值。2025年8月，国务院印发《关于深入实施“人工智能+”行动的意见》，推动人工智能与科技、产业融合。麦肯锡 2026 年全球人工智能调查显示，企业应用已经比较普遍，但组织层面的财务回报仍然不稳定。该调查是全球样本，不能直接当成上海企业的培训需求。")
    body(doc, "因此，面向企业家的课程应讨论如何把技术变成可以执行、衡量和持续改进的经营方案。")
    h3(doc, "学院已经有一门 AI 课")
    body(doc, f"学院 2025 年已开设《{EXISTING_AIGC['name']}》第三期、第四期，项目编号 {EXISTING_AIGC['codes']}。公开页面显示，学费 {EXISTING_AIGC['price_note']}，每期 {EXISTING_AIGC['seats']} 个席位。{EXISTING_AIGC['days']}。上课地点为{EXISTING_AIGC['campus']}。完成课时的学员按该项目规则获得非学历教育结业证书。页面由{EXISTING_AIGC['office']}发布。")
    body(doc, "这说明学院不是从零开始做人工智能培训，已经具备课程、师资和招生交付基础。新合作如果再做一轮工具和提示词培训，会和现有产品重叠，也缺少更高价格的理由。")
    body(doc, "新项目可以走三条路，具体由学院选择：与现有技能课形成“基础应用 + 企业管理进阶”；在学校认可的条件下，与既有项目团队一起做企业家定制班；或者用产业资源和现场教学做新的进阶项目。三条路都要避免校内重复建设。")
    h2(doc, "2. 双方各自拿得出什么")
    body(doc, "学院的优势，是校内外专家和跨学科课程、政府与大型企业委托培训经验、已经开出的人工智能课程，以及教学、班务和交付体系。学院也在探索面向社会招生的短期产品和渠道。")
    body(doc, "联合会的优势，是科技企业和产业组织联系、历史累计约 2 万人的私域社群、每年数百人规模的科技产业论坛、企业参访和沙龙，以及与部分央国企、金融机构、科技企业培训部门的联系。会员、园区和行业组织可以用来开发生源。")
    note(doc, "约 2 万人是历史累计规模，不是当前有效覆盖人数，更不是报名意向人数。招生测算不使用这个总数。")
    h2(doc, "3. 向谁招生")
    body(doc, "首期主要寻找有经营决策权、能提出真实业务课题、并且具备培训预算的管理者。零基础软件学习者和普通职场工具使用者，不作为主要对象，以免课程分散。")
    add_table(doc, ["目标客户", "关注问题", "课程方向", "价值"], SEGMENTS, [4.4, 5.0, 4.2, 3.6], font=8.5)
    h3(doc, "五类需要访谈验证的需求")
    bullet(doc, "**AI 战略与投资决策。** 技术路线、行业趋势、投资边界和商业价值。")
    bullet(doc, "**应用与经营改善。** 营销、客服、销售、知识库、流程自动化、研发管理中的投入产出。")
    bullet(doc, "**组织与管理。** AI 转型路线、职责、项目管理和效果衡量。")
    bullet(doc, "**融资与资本运作。** 产业投资、估值、融资规划、技术商业化和资本市场。")
    bullet(doc, "**产业交流。** 企业案例、标杆参访和学员之间的交流。")
    body(doc, "这五类来自政策、公开市场资料和现有客户结构，尚未经过联合会会员企业的专项问卷。")
    add_table(doc, ["决策因素", "访谈要确认的问题"], DECISION_FACTORS, [3.4, 13.8], font=10)
    h2(doc, "4. 公开价格只能作参照")
    add_table(doc, ["项目及主体", "学制", "公开学费", "使用注意"], COMPETITORS, [6.4, 3.2, 2.8, 4.8], font=8)
    body(doc, "课程长度、交付深度和证书主体不同，价格不能直接比较，也不代表实际招生人数或成交价。短期工具课已经有明确的公开价格。如果新课只讲趋势、提示词和工具演示，就没有明显的差异。")
    body(doc, "建议用六件事建立差异：企业经营问题诊断；高校师资与企业导师结合；AI、战略、产业投资和企业治理放在同一门课；高质量标杆参访；每位学员形成可讨论、可实施的企业 AI 应用方案；按学校规定建立班级学习和后续交流。")
    body(doc, f"复旦项目页面：{EXISTING_AIGC['url']}")
    h2(doc, "5. 六个渠道，先做定向")
    bullet(doc, "**会员及重点企业，优先级最高。** 筛选创始人、董事长、CEO、CTO 和管理层。方式包括定向邀请、需求访谈、专场说明会、团体报名和活动后跟进。建议先做 30 家企业的一对一访谈。")
    bullet(doc, "**私域社群，优先级最高。** 按职业和产业重新分类，重点识别科技企业创始人及高管、传统产业经营者、投资与产业基金、园区及国央企管理人员。路径是内容分享、主动咨询、需求访谈、正式报名。")
    bullet(doc, "**论坛与公开课，优先级高。** 建议主题为《人工智能时代：企业家如何实现从技术认知到商业价值创造》。规模 100—200 人，优先邀请经过筛选的企业负责人。大型论坛的报名转化要用实际数据验证。")
    bullet(doc, "**央国企和金融机构，优先级高。** 结合国泰海通、国药控股、中远相关企业及银行培训部门，探索定制班、团体报名和“课程 + 参访”。采购周期、预算和决策链通过访谈确认。")
    bullet(doc, "**园区、驻沪办和地方组织，优先级中高。** 包括长三角园区、地方科技企业协会、驻沪办事机构，以及地方工信、科技和产业人才渠道。政府出资培训履行实际采购和预算程序。")
    bullet(doc, "**学院现有客户，优先级高，须校方协调。** 学院评估既有企业客户、学员和官方招生平台能否用于新项目。联合会补充外部渠道。不得未经授权使用学校客户信息，也不得宣传尚未批准的课程。")
    h3(doc, "25 人的渠道配置")
    body(doc, "下表是内部配置目标，不是各渠道已经验证的转化能力。先取得企业培训需求，再做集中推广。")
    channel_rows = [(name, f"{n}人", f"{int(ratio * 100)}%") for name, n, ratio in CHANNELS]
    channel_rows.append(("合计", f"{AI_PLAN['target_n']}人", "100%"))
    add_table(doc, ["渠道", "拟招生", "占比"], channel_rows, [9.0, 4.0, 4.2], font=10, center_cols=(1, 2))
    h2(doc, "6. 产品方案")
    body(doc, f"论证名称采用《{AI_PLAN['name']}》。备选名称包括《{AI_PLAN['alt_names'][0]}》和《{AI_PLAN['alt_names'][1]}》。")
    body(doc, f"招生对象为科技企业创始人、董事长、总经理、核心高管、CTO、投资机构负责人，以及传统企业数字化转型负责人。首期 {AI_PLAN['class_size']}。")
    body(doc, f"学制 {AI_PLAN['months']}，周末集中。初步安排 {AI_PLAN['teach_days']} 天课堂和 {AI_PLAN['visit_days']} 天现场教学，合计 {AI_PLAN['total_days']} 天。实际学时按学院规定核定。")
    add_table(doc, ["模块", "主要内容", "天数"], [(a, b, f"{c}天") for a, b, c in MODULES], [4.6, 9.6, 3.0], font=10, center_cols=(2,))
    body(doc, "每位学员原则上完成一份《企业 AI 应用机会与实施计划》，包括当前经营问题、可选场景、数据和流程条件、预期成本收益与风险、90 天行动计划。讨论尽量使用真实企业案例。不承诺培训后必然实现降本、增收或融资。")
    body(doc, "是否颁发非学历教育结业证书、证书署名及相关权益，以学院正式批准方案为准。")
    h2(doc, "7. 定价与一套明确标注的假设")
    add_table(doc, ["产品层级", "建议价格", "定位"], PRICE_LADDER, [4.8, 4.4, 8.0], font=10)
    body(doc, f"标准产品 {AI_PLAN['price']:,} 元的依据，是课程长于现有短训，增加了战略、投资、项目实践和现场教学，同时与十万元以上的长期管理课程保持区分。它是市场测试建议价。")
    body(doc, "按此单价，20 人理论学费 336,000 元，25 人 420,000 元，30 人 504,000 元。这是报名学费，尚未扣除学校管理费用、教学成本、市场费用和税费。")
    body(doc, f"为了看价格是否承受得住，可以暂时假设：学费的 {int(AI_PLAN['disposable_ratio'] * 100)}% 用于覆盖直接支出及相关服务；每班教学组织、讲师、场地、参访、市场和运营合并为 {AI_PLAN['cost_assumption']:,} 元；先不考虑人数变化带来的成本变化，以及其他税费和分摊。")
    body(doc, "在这组假设下，20 人扣除成本后余额约 13,000 元，25 人约 55,000 元，30 人约 97,000 元，理论平衡点约为 19 人。工作表把单价、成本和比例做成可修改的输入，修改后自动重算。")
    note(doc, "50% 和 155,000 元不是学院确认的收费、管理费或成本。扣除成本后的余额不是任何一方的可分配利润。实际结算和服务采购，按照学校审批制度和合法合同执行。正式立项前，学院应重新测算师资、场地、招生服务、税费和其他支出。")
    h2(doc, "8. 转化路径与时间")
    body(doc, "用高意向客户优先的漏斗管理。下表是目标情景，不代表手中已经有这么多客户，也不是测得的转化率。")
    add_table(doc, ["环节", "试点目标"], [(a, f"{b}人") for a, b in FUNNEL], [12.0, 5.2], font=10, center_cols=(1,))
    body(doc, "正式立项前，先完成不少于 30 位目标企业管理者的需求访谈，检验内容、价格、时间和真实报名意向。联系信息使用已经授权的渠道。正式报名和收费通过校方批准的渠道进行。")
    bullet(doc, "**2026年11月至12月。** 完成首批 30 位访谈，整理企业与社群分层，形成课程、师资方向和模块初稿，与学院沟通和既有 AIGC 项目的关系，讨论初步预算。成果是调研简报、课程方案和目标客户清单。")
    bullet(doc, "**2027年1月至2月。** 取得学校要求的审批和宣传授权，明确招生规则、收费、授课和证书说明，确定教师、企业导师和可接待的参访资源，准备课程说明和报名渠道。")
    bullet(doc, "**2027年2月至4月。** 定向说明会、团体招生和渠道推广，完成审核及缴费。按招生成果和预算决定是否开班，并完成教学和现场教学。未审批前不以学校名义收费或承诺开班。")
    bullet(doc, "**2027年下半年。** 按首期结果决定是否做第二期、行业专项班或企业定制班。制造业 AI、金融科技、银发经济可以延伸，但以需求和课程效果为依据。")
    h2(doc, "9. 风险")
    add_table(doc, ["风险", "应对"], RISKS, [3.2, 14.0], font=9.5)
    h2(doc, "10. 本附件的结论")
    bullet(doc, "内容以企业 AI 应用和战略管理为核心，与学院既有技能班互补。")
    bullet(doc, "招生以企业定向为第一渠道，先开发重点科技企业、企业家社群和央国企培训客户。")
    bullet(doc, f"以 {AI_PLAN['target_n']} 人、{AI_PLAN['price']:,} 元/人作为初步测算基准，市场验证和学院审批后再调整。")
    bullet(doc, "用产业参访和学员自己的经营课题增强课程，让学习成果落在业务上。")
    bullet(doc, "用第一期的招生、交付和成本决定是否复制。先形成成熟模式。")
    body(doc, "合作原则是：联合开发市场，学院主导教学，产业资源协同，项目合规运营。2026 年底前完成需求调研、课程初稿、资源清单和合作框架；2027 年按正式审批和实际招生推进首期。")


def annex_plan(doc):
    page_break(doc)
    h1(doc, "附件四  2027年度拟合作项目计划")
    body(doc, f"计划周期：2027年1月至12月。编制单位：{META['compiler']}。编制日期：{META['compiled']}。本附件把合作收成四项核心落地业务、三项重点培育业务、三项储备探索业务。")
    body(doc, "2027 年优先做出实际项目，形成稳定的培训客户、标准化课程、产业参访服务和合作结算机制。不同时成立多个平台或机构。")
    note(doc, "表中的数量都是建议性工作目标，不代表已有订单、预算或校方承诺。")
    add_table(doc, ["指标", "2027年建议目标"], ANNUAL_GOALS, [4.4, 12.8], font=10.5)
    h2(doc, "项目总表")
    add_table(
        doc,
        ["序号", "项目", "对象", "时间", "目标", "层级"],
        PROJECTS_2027,
        [1.2, 4.6, 4.2, 1.8, 3.4, 1.6],
        font=8,
        center_cols=(0, 3, 5),
    )
    h2(doc, "四项核心业务")
    h3(doc, "项目 01  央国企及金融机构定制培训")
    body(doc, "按企业真实需求，为央国企、金融机构和大型企业提供管理人员专题培训。重点了解国泰海通、国药控股、中远相关企业及所属培训机构、招商银行上海分行、交通银行相关机构和上海地区城市商业银行。")
    body(doc, "课程方向包括人工智能赋能经营管理、金融科技与数字化转型、科技创新与成果转化、企业战略与组织创新、科技金融与产业投资。建议 3—5 天定制。2027 年完成 5 家以上需求访谈，争取落地 2—3 个项目。学院负责课程、师资和交付；联合会负责引荐、需求协调和协议约定的配套服务。")
    h3(doc, "项目 02  人工智能与科技创新企业领军人才高级研修班")
    body(doc, f"建议学制 {AI_PLAN['months']}，班额 {AI_PLAN['class_size']}，初步价格测试 {AI_PLAN['price']:,} 元/人。首期拟安排在 2027 年第二季度；招生、审批和交付符合预期后，再争取第四季度第二期。名称、学费、证书和课程权益须经学校批准。")
    h3(doc, "项目 03  科技企业标杆参访与现场教学")
    body(doc, "把联合会的企业资源变成学院可以选择、采购和交付的现场教学。方向覆盖人工智能与数字经济、机器人与具身智能、生物医药、高端制造与新能源汽车、科技金融、银发经济。首期重点协调腾讯、美团、B站、张江机器人谷、张江药谷、国泰海通及相关制造企业。2027 年争取完成 6—10 批独立参访或现场教学。每个项目写明接待主体、教学内容、人数、时长、报价和预约条件。")
    h3(doc, "项目 04  科创公开课与产业交流")
    body(doc, "把产业活动和人才培养连起来。主题可包括人工智能时代的企业经营变革、科技成果转化与资本市场、具身智能与未来产业、AI 赋能传统产业、科创企业家的创新管理。2027 年开展 4—6 场，其中 1 场作为规模较大的年度活动。宣传、报名、收费和校名使用，按项目性质履行审批。")
    h2(doc, "三项重点培育")
    bullet(doc, "**地方政府与园区产业人才培训。** 通过驻沪办、行业协会、园区和既有政府培训渠道拓展。争取 1—2 个项目，优先采用政府或园区依法委托。")
    bullet(doc, "**科技金融与企业投融资专题。** 面向创始人、投资机构和金融机构，完成 1 套课程方案，并争取 1 次试点。")
    bullet(doc, "**银发经济与智慧康养。** 围绕养老产业、长期护理保险、智慧康养、康复机器人和机构经营，完成 1 个专题并争取试点，教学与现场参访结合。")
    h2(doc, "三项储备")
    bullet(doc, "**科技企业微咨询与成果转化。** 经营诊断、AI 场景评估和产业课题。完成 2—3 项实际需求诊断，再评估后续合作。")
    bullet(doc, "**长三角产业研学与国际交流。** 形成上海、杭州、苏州、南京线路及国际团组的 1—2 套可执行方案。")
    bullet(doc, "**中国哲学与企业家人文领导力。** 整合哲学、国学和企业管理，形成 1 套经过讨论的课程初稿。")
    body(doc, "储备方向不作为 2027 年必须产生收入的考核任务。")
    h2(doc, "季度安排")
    add_table(doc, ["季度", "核心工作", "重点成果"], QUARTERS, [3.0, 7.2, 7.0], font=9)
    h2(doc, "规模示例")
    body(doc, "下面的收入只说明业务量级，方便讨论首批项目大概落在什么范围。它不是已经确认的营业收入，不是合作目标，也不是联合会可以获得的收益。收入归属、成本、学校管理费用、税费和各方服务费，由项目合同和财务审批确定。")
    scale_rows = []
    total = 0
    for name, qty, price, unit in SCALE_EXAMPLE:
        amount = qty * price
        total += amount
        scale_rows.append((name, f"{qty}{unit}", f"{price:,}元", f"{amount:,}元"))
    scale_rows.append(("合计", "—", "—", f"{total:,}元"))
    add_table(doc, ["收费项目", "示例数量", "单价假设", "理论收入"], scale_rows, [5.0, 3.4, 4.0, 4.8], font=10, center_cols=(1, 2, 3))
    body(doc, "测算假设各项目独立收费。已经打包在培训项目中的参访不再重复计入。若首期成功，再评估研修班第二期、行业专项班和其他储备业务。")
    h2(doc, "职责与节点")
    add_table(doc, ["事项", "学院（拟）", "联合会"], ROLES, [4.0, 6.6, 6.6], font=9, center_cols=(1, 2))
    add_table(doc, ["事项", "牵头", "节点"], MILESTONES, [7.2, 4.6, 5.4], font=10)
    body(doc, "节点以正式合作和学校立项、审批为前提。联合会负责市场开发和产业资源协调，不自动取得学校培训业务代理权。品牌授权、合作主体、收入归属、证书和招生权限，均须经过正式审批。")
    h2(doc, "每个实质项目都要有的六份资料")
    for item in (
        "项目需求及客户信息表。",
        "课程或参访服务方案。",
        "经审核的项目预算及报价。",
        "责任分工及实施进度表。",
        "正式合同及验收结算资料。",
        "项目效果反馈及复盘记录。",
    ):
        bullet(doc, item)
    h2(doc, "实施顺序")
    bullet(doc, "**先抓订单。** 从国泰海通、国药控股、中远相关企业和银行培训资源切入，形成真实需求。")
    bullet(doc, "**再抓产品。** 以人工智能企业家研修班为核心，做出能够稳定招生的非学历产品。联合会先验证约 25 名目标学员的报名潜力，不先投入大规模广告和会务。")
    bullet(doc, "**然后抓交付。** 把参访做成学院培训项目可以持续采购的现场教学，并把价格、联系人、人数和课程内容写成可选用的服务目录。")
    bullet(doc, "**最后抓拓展。** 前三项成熟后，再做地方政府培训、科技金融、银发经济和跨区域研学。")
    body(doc, "年度核心，是推动合作从资源对接进入正式签约、实际教学交付和可持续的客户开发。本计划不构成对任何项目、收益、招生人数或学校授权的承诺。")


def sources(doc):
    page_break(doc)
    h1(doc, "资料依据")
    for i, src in enumerate(SOURCES, 1):
        bullet(doc, f"{i}. {src}")
    body(doc, "上海交通大学相关课程页面供学院复核，不作为本方案的定价依据。麦肯锡调查只用于说明全球企业在组织回报上的普遍困难。")
    note(doc, "编制日期为 2026年11月17日。此后若学院公开课程、审批制度或客户联系发生变化，以新的书面材料为准，并同步修订工作表中的对应行。")


def build(path: Path):
    doc = Document()
    setup(doc)
    cover(doc)
    page_break(doc)
    guide(doc)
    contents(doc)
    page_break(doc)
    framework(doc)
    annex_visits(doc)
    annex_clients(doc)
    annex_ai(doc)
    annex_plan(doc)
    sources(doc)
    doc.save(path)


if __name__ == "__main__":
    out = Path(__file__).resolve().parents[1] / "output"
    out.mkdir(exist_ok=True)
    target = out / "复旦终身教育学院与杨浦科联合作框架方案.docx"
    build(target)
    print(target)
