# -*- coding: utf-8 -*-
"""生成《泰隆银行·科创企业成长伙伴计划（2027）》商务沟通稿 Word。"""

import os
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import partner2027_data as D

NAVY = RGBColor(0x0C, 0x1E, 0x33)
GOLD = RGBColor(0xA6, 0x84, 0x4A)
INK = RGBColor(0x1B, 0x28, 0x36)
MUTED = RGBColor(0x5C, 0x68, 0x74)
RED = RGBColor(0x8E, 0x2F, 0x2F)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "微软雅黑"
CONTENT_CM = 16.6
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "deliverables", "泰隆银行科创企业成长伙伴计划2027_商务沟通稿.docx")


def set_run(run, size, bold=False, color=INK, font=FONT):
    run.font.name = font
    run.bold = bold
    run.font.size = Pt(size)
    run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)
    rFonts.set(qn("w:eastAsia"), font)
    rFonts.set(qn("w:cs"), font)


def shade_cell(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def set_cell_width(cell, cm):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcW")):
        tcPr.remove(old)
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"), str(int(cm * 567)))
    tcW.set(qn("w:type"), "dxa")
    tcPr.append(tcW)


def set_cell_margins(cell, top=60, bottom=60, left=80, right=80):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.find(qn("w:tcMar"))
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for m, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = tcMar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tcMar.append(node)
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")


def valign(cell, val="center"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    v = tcPr.find(qn("w:vAlign"))
    if v is None:
        v = OxmlElement("w:vAlign")
        tcPr.append(v)
    v.set(qn("w:val"), val)


def prevent_row_split(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    cant = OxmlElement("w:cantSplit")
    trPr.append(cant)


def write_cell(cell, text, size=9, bold=False, color=INK, fill=None, align="left"):
    if fill:
        shade_cell(cell, fill)
    valign(cell, "center")
    set_cell_margins(cell)
    p = cell.paragraphs[0]
    p.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }[align]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.05
    # clear existing runs
    if p.runs:
        for r in p.runs:
            r.text = ""
    run = p.add_run(text)
    set_run(run, size, bold, color)
    return p


def set_table_widths(table, widths):
    table.autofit = False
    table.allow_autofit = False
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    total = sum(widths)
    tblW.set(qn("w:w"), str(int(total * 567)))
    tblW.set(qn("w:type"), "dxa")
    tblLayout = tblPr.find(qn("w:tblLayout"))
    if tblLayout is None:
        tblLayout = OxmlElement("w:tblLayout")
        tblPr.append(tblLayout)
    tblLayout.set(qn("w:type"), "fixed")
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for child in list(grid):
            grid.remove(child)
    else:
        grid = OxmlElement("w:tblGrid")
        tblPr.addnext(grid)
    for w in widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(w * 567)))
        grid.append(gc)
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            set_cell_width(cell, widths[i])


def add_table(doc, headers, data, widths, font=9, header_fill="0C1E33"):
    table = doc.add_table(rows=1 + len(data), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    set_table_widths(table, widths)
    for i, h in enumerate(headers):
        write_cell(table.rows[0].cells[i], h, size=font, bold=True, color=WHITE, fill=header_fill, align="center")
    prevent_row_split(table.rows[0])
    trPr = table.rows[0]._tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    trPr.append(tblHeader)
    for r, row in enumerate(data):
        fill = "F6F3EE" if r % 2 == 0 else "FFFCF8"
        prevent_row_split(table.rows[r + 1])
        for c, val in enumerate(row):
            align = "center" if c == 0 and len(str(val)) <= 6 else "left"
            write_cell(table.rows[r + 1].cells[c], str(val), size=font, color=INK, fill=fill, align=align)
    doc.add_paragraph()
    return table


def add_p(doc, text, size=10.5, bold=False, color=INK, space_before=0, space_after=6, align="left", first_line=None):
    p = doc.add_paragraph()
    p.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }[align]
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if first_line:
        p.paragraph_format.first_line_indent = Cm(first_line)
    run = p.add_run(text)
    set_run(run, size, bold, color)
    return p


def add_h(doc, text, level):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16 if level == 1 else 12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.1
    p.paragraph_format.keep_with_next = True
    size = {1: 16, 2: 13.5, 3: 12}[level]
    run = p.add_run(text)
    set_run(run, size, True, NAVY if level < 3 else RGBColor(0x1A, 0x3A, 0x5C))
    if level == 1:
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "8")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "A6844A")
        pBdr.append(bottom)
        pPr.append(pBdr)
    return p


def add_bullets(doc, items, size=10.5):
    for item in items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.75)
        p.paragraph_format.first_line_indent = Cm(-0.4)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.12
        run = p.add_run("·  " + item)
        set_run(run, size, False, INK)


def callout(doc, title, body, fill="F4ECDC"):
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    set_table_widths(table, [CONTENT_CM])
    cell = table.cell(0, 0)
    shade_cell(cell, fill)
    set_cell_margins(cell, 80, 80, 120, 120)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(title)
    set_run(r, 10.5, True, NAVY)
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(2)
    p2.paragraph_format.space_after = Pt(0)
    p2.paragraph_format.line_spacing = 1.12
    r2 = p2.add_run(body)
    set_run(r2, 10.5, False, INK)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def add_page_field(paragraph, kind):
    run = paragraph.add_run()
    set_run(run, 9, False, MUTED)
    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    run._r.append(fld1)
    run2 = paragraph.add_run()
    set_run(run2, 9, False, MUTED)
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {kind} "
    run2._r.append(instr)
    run3 = paragraph.add_run()
    set_run(run3, 9, False, MUTED)
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "end")
    run3._r.append(fld2)


def setup_section(section, with_header):
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(1.8)
    section.header_distance = Cm(0.8)
    section.footer_distance = Cm(0.6)
    if not with_header:
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False
        section.header.paragraphs[0].text = ""
        section.footer.paragraphs[0].text = ""
        return
    section.header.is_linked_to_previous = False
    hp = section.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    hp.paragraph_format.space_after = Pt(2)
    r = hp.add_run("泰隆银行 · 科创企业成长伙伴计划（2027）    商务沟通稿")
    set_run(r, 9, False, MUTED)
    pPr = hp._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "A6844A")
    pBdr.append(bottom)
    pPr.append(pBdr)

    section.footer.is_linked_to_previous = False
    fp = section.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = fp.add_run("洽谈稿 · 预算、接待与品牌授权均待正式确认 · 第 ")
    set_run(r, 9, False, MUTED)
    add_page_field(fp, "PAGE")
    r2 = fp.add_run(" 页")
    set_run(r2, 9, False, MUTED)


def build():
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = INK
    rPr = normal.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), FONT)
    rFonts.set(qn("w:hAnsi"), FONT)
    rFonts.set(qn("w:eastAsia"), FONT)
    pf = normal.paragraph_format
    pf.line_spacing = 1.15
    pf.space_after = Pt(6)

    core = doc.core_properties
    core.title = D.TITLE
    core.subject = "商务沟通稿"
    core.author = "上海市杨浦区科技企业联合会"
    core.category = "洽谈稿"
    core.comments = "依据2026-11-06与2026-11-09两次洽谈整理，预算与接待待确认。"

    setup_section(doc.sections[0], with_header=False)

    add_p(doc, "商务沟通稿", 12, True, GOLD, space_after=2)
    add_p(doc, D.DATE, 10.5, False, MUTED, space_after=10)
    add_p(doc, "泰隆银行", 14, True, GOLD, space_after=0)
    add_p(doc, "科创企业成长伙伴计划", 26, True, NAVY, space_after=0)
    add_p(doc, "2027", 18, True, GOLD, space_after=8)
    add_p(doc, D.THESIS, 12, False, INK, space_after=10)

    meta = [
        ("建议合作单位", D.PARTNERS),
        ("拟联动单位", D.LINKED),
        ("文件性质", D.NATURE),
        ("主要依据", "2026年11月6日企业服务合作洽谈；2026年11月9日泰隆银行、复旦科技园及相关方交流"),
        ("配套文件", "同名简报（PPT）、资源排期与预算表（Excel）"),
    ]
    add_table(doc, ["项目", "内容"], meta, [3.4, 13.2], font=10)

    callout(
        doc,
        "提交前请先读这一段",
        "本稿把两次会议里已经讨论的事项，整理成可以交给行长和相关单位的文本。"
        "年度排期、活动指标和金额是建议方案。11月6日谈到的参访经费“十几万至二十几万元”是口头讨论，不是已批准预算。"
        "11月18日约100人科创路演，以及茶歇、会务、场地支持，需要先核实活动是否举行、费用是否获批。"
        "名单中的企业尚未确认2027年接待。复旦科技园及学校单位可以参与资源协同；联合主办、品牌、课程和证书必须另行授权。"
        "本稿不承诺招生录取或任何特殊入学权益，也不建议按贷款、开户或转介结果向个人支付佣金。",
    )

    # body section
    new_sec = doc.add_section()
    setup_section(new_sec, with_header=True)
    sectPr = new_sec._sectPr
    pgNumType = OxmlElement("w:pgNumType")
    pgNumType.set(qn("w:start"), "1")
    sectPr.append(pgNumType)

    add_h(doc, "一、这份稿怎么用", 1)
    add_p(doc, "三份文件分工不同，建议一起提交。", first_line=0.74)
    add_bullets(doc, [
        "PPT用于当面沟通，把判断、架构、报价和近期决定讲清楚。",
        "本稿（Word）是可以留档的沟通文本，正文谈合作机制，两份附件分别谈参访资源和路演计划。",
        "Excel是可计算、可筛选的工作表，包含资源状态、年历、预算公式和会后跟踪空表。",
    ])
    add_p(doc, "阅读顺序建议先看第二章的一页纸，再看费用和近三个月安排。附件供资源核验和排期讨论，不需要在第一次会议上逐条宣读。")

    add_h(doc, "二、合作共识（一页纸草案）", 1)
    add_p(doc, "银行此前希望用一页纸把合作内容说清，并分阶段开票。这个做法可以保留。一页纸解决“做什么”，合同和采购流程解决“谁签、怎么付、怎么验收”。")
    consensus = [
        ("项目名称", "泰隆银行·科创企业成长伙伴计划"),
        ("合作架构", "1个联合服务平台 + 企业、教育、金融3类资源 + 6项业务"),
        ("平台构成", "泰隆银行与上海市杨浦区科技企业联合会联合服务；复旦科技园、高校继续教育及区域园区按授权联动"),
        ("服务闭环", D.LOOP),
        ("2026年试点", "核实11月18日路演及支持费用；12月完成1场20—30人标杆参访"),
        ("2027年参访", "原则上每月1场，全年12场，年度服务240—360人次"),
        ("2027年路演", "4场重点产业路演 + 4场小型专项对接，作为独立项目"),
        ("第一轮经费", "参访年度建议15万—20万元；路演年度建议14万元。均待审批，按实核定"),
        ("近期确认", "预算量级、首场参访、承办与结算、年度计划、客户服务与资源权益"),
        ("明确不做", "未授权不使用高校品牌；不承诺入学；不对个人按结果支付佣金"),
    ]
    add_table(doc, ["事项", "草案表述"], consensus, [3.4, 13.2], font=10)
    add_p(doc, "如双方认可上述方向，可在下表签署沟通意见。签署不代表预算已经批准，也不代替采购合同。", size=10.5)
    sign = doc.add_table(rows=5, cols=3)
    sign.style = "Table Grid"
    set_table_widths(sign, [3.2, 6.7, 6.7])
    headers = ["", "泰隆银行", "项目执行方"]
    labels = ["确认人", "职务", "日期", "沟通意见"]
    for i, h in enumerate(headers):
        write_cell(sign.rows[0].cells[i], h, 10, True, WHITE, "0C1E33", "center")
    for r, lab in enumerate(labels, 1):
        write_cell(sign.rows[r].cells[0], lab, 10, True, NAVY, "F6F3EE", "center")
        for c in (1, 2):
            write_cell(sign.rows[r].cells[c], " ", 10, False, INK, "FFFCF8")
            # give signature rows some height via paragraph spacing
            sign.rows[r].cells[c].paragraphs[0].paragraph_format.space_before = Pt(8)
            sign.rows[r].cells[c].paragraphs[0].paragraph_format.space_after = Pt(8)
    doc.add_paragraph()

    add_h(doc, "三、两次洽谈与合作判断", 1)
    add_h(doc, "3.1  2026年11月6日：企业服务合作洽谈", 2)
    add_p(doc, "这次沟通里，参访已经被放到可执行的位置。泰隆银行提出，由项目执行方负责企业参访的资源清单、预算和组织执行。年度经费曾谈到十几万至二十几万元。这是口头讨论的量级，后续仍要走银行审批，本稿不把它写成已批预算。")
    add_p(doc, "同一次会议还谈到客户子女的高校研学。这个需求可以单列成家庭研学产品，范围限于校园文化、科普课程和创新体验。招生录取和任何特殊入学权益不在合作内容里。")
    add_h(doc, "3.2  2026年11月9日：泰隆银行、复旦科技园及相关方交流", 2)
    add_p(doc, "三方交流把合作从单次参访扩到了科技园和银行国际业务。泰隆银行虹口支行介绍了国际业务，以及航运、进口贸易相关客户基础。这是后文“科技企业出海”场次的依据，具体产品仍由银行核定。")
    add_p(doc, "会议还提到，11月18日拟举办一场约100人的科创路演。泰隆方面考虑支持茶歇、会务和场地费用，并设置企业对接环节。这是距离现在最近的试点，但有两件事必须先核实：活动是否仍按计划举行，支持费用是否已经获批。在核实完成前，对外材料只写“拟举办、支持在考虑中”。")
    add_h(doc, "3.3  合作判断", 2)
    add_p(doc, "如果合作只停留在帮忙联系一家企业、赞助一场活动或介绍若干客户，执行方没有可重复交付的产品，银行也没有稳定的企业服务场景。更合适的做法，是争取一个有年度经费、可以按项目验收的综合服务机制。")
    add_p(doc, "机制的顺序是：参访负责把企业家带到现场，科技金融负责形成银行可跟进的需求，高校和复旦相关资源负责拉开与普通商务活动的差别，培训、产业对接和投融资负责把单次活动变成下一场合作。")

    add_h(doc, "四、总体定位与合作架构", 1)
    add_h(doc, "4.1  项目名称与目标", 2)
    add_p(doc, "建议项目名称使用“泰隆银行·科创企业成长伙伴计划”。名称把银行放在前面，是因为客户经营和金融服务的责任在银行；“成长伙伴”对应的是陪伴企业，而不是以到场人数作为唯一成果。")
    add_p(doc, "年度参访希望建立五样东西：一套可以更新的参访资源库，一批有真实交流需求的企业家，一个银行与科技企业、园区的常态互动机制，一套需求收集和评价办法，以及可以延伸到路演、培训和跨境金融的合作机会。")
    add_h(doc, "4.2  1个平台 + 3类资源 + 6项业务", 2)
    add_table(
        doc,
        ["层次", "构成", "在合作中的作用"],
        [
            ["联合服务平台", "泰隆银行 × 上海市杨浦区科技企业联合会", "对银行和会员企业的总接口"],
            ["拟联动", "复旦科技园、高校继续教育及科研机构、区域园区", "按授权提供场地、项目、课程和企业"],
            ["企业资源", "商协会、园区、科技企业", "参访对象、路演项目和沙龙嘉宾"],
            ["教育资源", "高校专家、课程、校友", "研修、研学和成果转化"],
            ["金融资源", "银行、投资机构、产业基金", "债权服务、股权交流和产业资本"],
        ],
        [3.2, 7.2, 6.2],
        font=9.5,
    )
    add_table(
        doc,
        ["序号", "业务", "在2027年的位置"],
        [[f"{i+1:02d}", name, "基础产品" if i == 0 else ("独立扩展" if i == 1 else "按预算择机")]
         for i, name in enumerate(D.BUSINESSES)],
        [2.0, 4.6, 10.0],
        font=10,
    )
    add_p(doc, "六项业务对应后文六个方向。2026年年底到2027年第一季度，只把标杆参访和科技金融路演作为必须推进的两项，其余四项进入储备，避免第一次就把合作盘子做得过大。")

    add_h(doc, "五、六个合作方向", 1)
    add_table(
        doc,
        ["方向", "内容", "优先级", "阶段安排"],
        [[d["name"], d["detail"], d["priority"], d["window"]] for d in D.DIRECTIONS],
        [3.6, 6.6, 1.8, 4.6],
        font=9,
    )
    add_h(doc, "5.1  标杆参访：先做成年度服务", 2)
    add_p(doc, "参访是最先能够签约、也最快能够形成服务收入的产品。建议固定三个系列，便于银行向客户解释，也便于执行方提前储备接待资源。")
    add_table(doc, ["系列", "优先资源"], D.SERIES, [5.2, 11.4], font=10)
    add_p(doc, "标准场次为每月1场、20—30名董事长、总经理或核心高管、约半天。流程、资源分级和2027年排期见附件一。执行方负责企业对接、策划和组织；银行负责目标客户和活动采购；复旦科技园和联合会只做各自授权范围内的宣传和邀约。")
    add_p(doc, "参访资源应成为执行方的长期产品。单次无偿帮银行联系企业，无法覆盖约访、方案和现场成本，也不能积累可复用的接待关系。")
    add_h(doc, "5.2  科技金融路演：形成优质企业入口", 2)
    add_p(doc, "建议每季度安排重点场，全年采用“4场重点产业路演 + 4场小型专项对接”。重点场为6—8家企业路演、投资机构点评、银行融资服务和一对一洽谈；为了保证交流时间，一般控制在5—6家。小型场为3—4个项目、20—30人闭门沟通。")
    add_p(doc, "这一产品和执行方已有的投资机构、高校成果转化资源，以及泰隆银行的债权融资能力放在同一场活动里。它不替代银行尽调，也不对投资结果作承诺。全年安排、专题和预算见附件二。")
    add_p(doc, "11月18日的约100人路演，是检验这套方法的最近机会。确认活动仍举行、且支持费用获批之后，再把银行的企业对接环节写进正式议程。")
    add_h(doc, "5.3  企业服务沙龙：持续的小规模获客", 2)
    add_p(doc, "在参访和路演之间，每月可以有一场10—20人的闭门沙龙。银行不必在开场推销贷款。企业先把经营问题提出来，再由客户经理判断是否进入金融服务。建议主题包括：")
    add_bullets(doc, D.SALON_TOPICS)
    add_h(doc, "5.4  复旦研学与企业家培训：两条业务线", 2)
    add_p(doc, "这一方向可以与正在推进的复旦终身教育学院合作衔接，但建议分成两条独立业务线，避免把企业家研修和客户家庭活动混成一个招生项目。")
    add_table(doc, ["对象", "产品内容"], EDU_ROWS(), [4.2, 12.4], font=10)
    add_p(doc, "高校场地、课程、教师和证书权益分别确认。任何材料都不写招生录取或特殊入学安排。11月6日会议里银行对客户子女高校研学表现出兴趣，因此第二条业务线有需求基础，但仍须等学校授权后再对外邀约。")
    add_h(doc, "5.5  企业出海与年度品牌活动", 2)
    add_p(doc, "出海场次建议与长三角制造业、人工智能出海企业和外贸公司结合，主题定为“科技企业出海与跨境金融服务专场”。企业若有结算、外汇和供应链融资需求，这场活动的业务价值会高于普通讲座。金融产品名称和适用条件由银行核定，执行方不代为承诺额度或费率。")
    add_p(doc, "2027年的品牌活动可以有两类：一类是新策划的科创企业年度大会或产业论坛，另一类是复旦科技园、科技企业联合会既有活动中的金融服务合作。泰隆银行可以按实际预算成为金融服务合作伙伴或赞助方。大型赞助不作为这一轮合作的主要收入来源。")

    add_h(doc, "六、角色、分工与需要确认的事项", 1)
    add_p(doc, "建议项目执行方的角色是科创企业综合服务的发起和运营合作方。可以反复使用的能力有三项：科技企业与商协会的组织能力，高校及复旦相关资源的协调能力，科技项目与投资机构的连接能力。")
    add_table(doc, ["主体", "负责事项"], D.ROLES, [4.8, 11.8], font=10)
    add_p(doc, "实际主办、承办、协办的名称，以授权文件和当次活动确认文件为准。对外海报、手册和证书在授权完成前，只使用已经同意的单位名称。")
    add_h(doc, "6.1  请行长明确的五件事", 2)
    add_table(
        doc,
        ["事项", "希望确认的内容"],
        D.FIVE,
        [4.4, 12.2],
        font=10,
    )
    add_h(doc, "6.2  会前核实清单", 2)
    add_p(doc, "下面八项是把五件事拆成可以追踪的动作，状态以2026年11月17日为准。")
    add_table(
        doc,
        ["事项", "当前状态", "建议确认方", "时间"],
        D.CONFIRM,
        [4.4, 5.4, 3.6, 3.2],
        font=8.5,
    )

    add_h(doc, "七、费用、报价与结算", 1)
    add_p(doc, "费用分三档，供谈判使用。它不是泰隆银行已经确认的采购价。实际金额按人数、参访企业收费、专家、交通和场地核定，并包含合同约定的税费口径。")
    add_table(doc, ["合作模式", "建议费用", "收益方式"], D.PRICING, [4.6, 5.0, 7.0], font=10)
    add_h(doc, "7.1  第一轮怎么报", 2)
    add_p(doc, "建议先提交参访年度合作15万—20万元，把路演年度14万元作为第二项独立方案。这样与行长优先把经费用于企业参访的意向一致，也避免第一次报价超过口头讨论的量级太多。")
    add_p(doc, "15万—20万元不是“12场 × 5,000—10,000元”的简单乘积。5,000—10,000元/场只覆盖策划、协调和组织执行服务费。企业接待费、交通、专家和场地按单场实报。年度金额同时包含全年资源维护。12场服务费本身约6万—12万元，年度建议高于这个区间，是为了把按实发生的活动成本一并纳入银行可审批的盘子。")
    add_p(doc, "路演标准版为14万元：重点路演4场 × 2万元，小型对接4场 × 1.2万元，年度协调与复盘1.2万元。若园区提供免费场地或合作方分担会务，银行采购额相应下调。参访预算与路演预算分别测算。若银行只设一个企业服务年度预算，应重新汇总和压缩，不能把两套数字相加后当作已经同意的总额。")
    add_h(doc, "7.2  结算原则", 2)
    add_bullets(doc, [
        "先单场试点、验收、结算，再过渡到年度框架下的分项目采购。",
        "可以沿用一页纸明确合作内容；承办主体、费用标准、发票、审批和结算符合银行采购及财务要求。",
        "不按贷款审批金额、开户金额或客户转介结果向个人支付佣金。",
        "客户开户、授信和资金业务的转化，由银行自行统计和评价。",
    ])
    add_p(doc, "倾向的合作节奏是：2026年试点，2027年年度框架。试点场次用于验证接待质量和客户反馈，不在试点前锁定全年12场的具体企业。")

    add_h(doc, "八、2026年11月至2027年第一季度", 1)
    add_table(doc, ["时间", "核心任务", "具体成果"], D.TIMELINE, [4.2, 4.6, 7.8], font=9.5)
    add_p(doc, "活动结束后按7日、30日、90日跟踪。7日内完成总结、反馈和需求分类；30日内完成重点企业回访；90日内复盘满意度和下一步需求。考核看有效企业数、需求条数、银行对接数、投资与产业交流数、满意度和授权情况。人数只作为规模记录。")
    add_h(doc, "8.1  建议的推进顺序", 2)
    add_bullets(doc, [
        "第一步，确认11月18日科创路演是否举行，以及赞助或会务支持是否获批。",
        "第二步，落实1—2个标杆参访，优先京东、美团、傅里叶和张江机器人相关资源。",
        "第三步，以参访为基础项目、路演为扩展项目，签订年度框架，写明预算、权益和结算。",
    ])
    add_p(doc, "经营上的顺序同样明确：先把参访做成稳定的年度服务，再延伸到投融资路演、复旦终身教育和企业出海。这三块的客单价和业务关联度更高，但都依赖参访和路演先积累企业。")

    add_h(doc, "附件一    2027年度科技企业标杆参访资源清单", 1)
    add_p(doc, "泰隆银行·科创企业成长伙伴计划", 11, True, GOLD, space_after=2)
    add_p(doc, "文件性质：资源储备与年度合作洽谈稿。编制日期：2026年11月17日。", 10, False, MUTED)
    add_h(doc, "A1  项目定位与年度目标", 2)
    add_p(doc, "依托科技企业联合会、复旦科技园、高校校友网络和科技产业资源，与泰隆银行一起做标杆参访、企业家交流、产业链对接和科技金融服务。")
    add_table(doc, ["指标", "建议值"], D.VISIT_TARGETS, [4.2, 12.4], font=10)
    add_p(doc, "重点领域：人工智能、具身智能、数字经济、高端装备、新能源汽车、科技消费、航运与企业出海。")
    add_h(doc, "A2  资源状态怎么读", 2)
    add_table(
        doc,
        ["类别", "含义", "本稿中的使用方式"],
        [
            ["A类 · 重点推进", "洽谈中已有对接线索，或可协助约访", "优先联系，仍须确认接待"],
            ["B类 · 历史参访资源", "此前有组织参访的经历", "可尝试重新协调，不视为延续档期"],
            ["C类 · 拓展储备", "有主题价值，接待条件尚未开发", "用于替换和补位，不排入已承诺场次"],
        ],
        [4.0, 6.2, 6.4],
        font=10,
    )
    add_h(doc, "A3  A类：重点推进资源", 2)
    a_rows = [[f"{v['no']:02d}", v["org"], v["theme"], v["status"], v["month"]]
              for v in D.VISITS if v["cat"] == "A"]
    add_table(doc, ["序号", "单位 / 平台", "核心主题", "资源状态", "建议月份"], a_rows,
              [1.3, 4.9, 4.6, 3.6, 2.2], font=8.5)
    add_h(doc, "A4  B类：具有历史参访经验的资源", 2)
    add_p(doc, "下表来自既往洽谈和参访经历。它不代表相关企业已经同意承接2027年活动。接待地点、人数、费用和内容都要重新确认。")
    b_rows = [[f"{v['no']:02d}", v["org"], v["theme"], v["track"], v["month"]]
              for v in D.VISITS if v["cat"] == "B"]
    add_table(doc, ["序号", "单位", "建议主题", "产业方向", "建议月份"], b_rows,
              [1.4, 4.6, 4.8, 3.4, 2.4], font=8)
    add_h(doc, "A5  C类：建议拓展的产业方向", 2)
    add_table(doc, ["产业领域", "拟开发资源", "参访价值", "年历位置"], D.VISIT_DIRS,
              [3.8, 4.8, 5.2, 2.8], font=8.5)
    add_h(doc, "A6  2027年参访排期", 2)
    add_p(doc, "排期是储备方案。1月、6月、8月出现两家优先单位时，以先确认接待的一家为准，另一家转为储备或并场。魔力社区、魔丸熊、腾讯、霸王茶姬未单列进月历，可按报名和预算替换插入。")
    add_table(doc, ["月份", "参访主题", "优先资源"], D.VISIT_CALENDAR, [2.2, 6.4, 8.0], font=10)
    add_h(doc, "A7  标准流程、分工与经费", 2)
    add_table(doc, ["时间", "内容"], D.VISIT_FLOW, [4.0, 12.6], font=10)
    add_p(doc, "泰隆银行负责客户筛选与邀约、金融服务、客户经理、采购和费用审批。上海市杨浦区科技企业联合会及项目执行方负责资源协调、方案、行程、现场和总结。复旦科技园及其他机构按项目参与邀约和园区协同。")
    add_p(doc, "经费上，先按单场报价、验收和结算，再建立年度合作。建议2026年12月落实一场试点，2027年第一季度转入每月一场。每场预算单列企业接待费，不把服务费说成全包价。")
    add_h(doc, "A8  预期成果", 2)
    add_bullets(doc, [
        "一套可以按季度更新的科技企业参访资源库。",
        "一批带着经营和交流需求来的科技企业家。",
        "泰隆银行与科技企业、产业园区的常态互动。",
        "活动执行、需求收集和后续服务的记录方法。",
        "能够延伸到路演、培训、产业协同和跨境金融的场次。",
    ])
    add_p(doc, "项目定位：让企业参访成为科技企业交流的入口，让泰隆银行成为企业成长过程中的长期服务伙伴。", bold=True)

    add_h(doc, "附件二    2027年度科创企业投融资路演计划", 1)
    add_p(doc, "泰隆银行·科创企业成长伙伴计划", 11, True, GOLD, space_after=2)
    add_p(doc, "规划年度：2027年。文件性质：年度合作洽谈稿。本附件与参访预算分别测算。", 10, False, MUTED)
    add_h(doc, "B1  总体目标", 2)
    add_p(doc, "围绕人工智能、具身智能、生物医药、先进制造、新能源和企业出海，做常态化路演和产业服务。方法是“产业筛选 + 项目路演 + 资本对接 + 银行服务 + 后续跟踪”。下列数字是组织工作目标，不是已经拿到的项目、投资或授信。")
    add_table(doc, ["指标", "年度计划"], D.ROADSHOW_TARGETS, [6.0, 10.6], font=10)
    add_h(doc, "B2  八场安排", 2)
    rs_rows = [[m, n, "重点路演" if k == "重点" else "小型对接", s, f] for m, n, k, s, f in D.ROADSHOWS]
    add_table(doc, ["时间", "活动名称", "类型", "规模", "重点内容"], rs_rows,
              [1.6, 5.6, 2.2, 2.2, 5.0], font=8)
    add_p(doc, "6月、9月和12月用于企业复访、重点项目推进、活动复盘和下一年度储备，不强制增加路演场次。")
    add_h(doc, "B3  四个重点专题", 2)
    for i, t in enumerate(D.THEMES, 1):
        add_h(doc, f"专题{i}  {t['name']}", 3)
        add_p(doc, "拟定名称：" + t["title"])
        add_p(doc, "重点赛道：" + t["tracks"])
        add_p(doc, "拟邀对象：" + t["guests"])
        add_p(doc, "合作价值：" + t["value"])
    add_h(doc, "B4  标准流程与筛选", 2)
    add_p(doc, "下面是一场60—80人重点路演的参考议程。小型对接会压缩为3—4个项目和更长的闭门讨论。")
    add_table(doc, ["时间", "内容"], D.ROADSHOW_FLOW, [4.0, 12.6], font=10)
    add_p(doc, "拟参与企业原则上满足下列条件中的若干项，不要求一次具备全部条件：")
    add_bullets(doc, D.SCREEN)
    add_table(doc, ["项目库", "放入的企业"], D.POOLS, [4.4, 12.2], font=10)
    add_p(doc, "企业是否符合银行授信条件或基金投资标准，由相关机构独立尽调后决定。执行方提供的是真实、完整的项目介绍和基本经营资料，不代替尽调结论。")
    add_h(doc, "B5  分工", 2)
    add_p(doc, "泰隆银行提供金融服务团队和客户经理，参与筛选有真实经营与金融需求的企业，负责银行客户邀约、金融咨询和后续服务，并在获批预算内承担约定的活动服务费。尽调与授信按银行自身流程办理。")
    add_p(doc, "杨浦区科技企业联合会及项目执行方负责总体策划和产业主题，组织企业、商协会和产业专家，协调投资机构、园区和产业资源，完成受托的执行、邀约和对接，并提交活动总结、需求反馈和非敏感跟踪材料。")
    add_p(doc, "复旦科技园及其他机构按具体活动协助提供企业和成果转化项目，参与推荐，并探索与现有创业大赛、成果转化活动联动。是否列为联合主办或协办，以正式授权为准。")
    add_h(doc, "B6  年度预算建议", 2)
    add_p(doc, "优先采用“单场采购 + 年度框架统筹”。标准版测算如下。")
    budget_rows = []
    total = 0
    for name, qty, price, note in D.ROADSHOW_BUDGET:
        amount = qty * price
        total += amount
        budget_rows.append([name, str(qty), f"{price:,.0f}", f"{amount:,.0f}", note])
    budget_rows.append(["年度建议合计", "—", "—", f"{total:,.0f}", "与参访预算分别列支"])
    add_table(doc, ["项目", "数量", "单场预算（元）", "年度预算（元）", "说明"], budget_rows,
              [4.6, 1.6, 3.0, 3.0, 4.4], font=8.5)
    add_p(doc, "费用覆盖会务组织、场地、物料、执行、企业邀约、项目筛选和成果整理。税费以采购清单和合同为准。部分场次若由园区免费提供场地，或由合作机构承担部分会务，银行采购预算相应降低。")
    add_h(doc, "B7  成果、评估与实施步骤", 2)
    add_table(doc, ["节点", "完成内容"], D.TRACK, [3.0, 13.6], font=10)
    add_p(doc, "主要记录：有效参会企业及机构数量，实际路演项目数量，融资与经营需求条数，银行金融服务对接数量，投资机构及产业合作交流数量，企业满意度和持续参与情况。客户信息的交换和使用须取得适当授权。")
    add_table(
        doc,
        ["阶段", "时间", "工作"],
        D.PHASES,
        [2.6, 3.6, 10.4],
        font=9.5,
    )
    add_p(doc, "合作愿景是逐步形成“发现科技企业、对接产业需求、提供债权融资服务、连接股权资本、组织企业家交流”的服务生态，使泰隆银行的服务从单一产品延伸到企业成长过程中的持续陪伴。这一段是方向，不是对2027年业务结果的承诺。")

    add_h(doc, "附则", 1)
    add_bullets(doc, [
        "两份附件中的企业与合作模式，主要依据2026年11月6日、11月9日会议纪要及既往参访经历整理。",
        "年度排期、活动指标和预算是本次建议，不是已经获批的银行安排。",
        "对外宣传、联合主办、课程和证书，在相关单位书面授权前不使用其名称作主办方。",
        "本稿与PPT、Excel如有数字出入，以Excel中的公式表为准，并在提交前核对一次。",
        "本稿自2026年11月17日起作为洽谈文本使用，内容随接待确认和预算批复更新。",
    ])
    add_p(doc, "上海市杨浦区科技企业联合会  ·  项目执行方", 10.5, True, NAVY, space_before=8)
    add_p(doc, "承办主体以最终采购合同载明的单位为准。", 10, False, MUTED)

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT)
    print(OUT)


def EDU_ROWS():
    return D.EDU_LINES


if __name__ == "__main__":
    build()
