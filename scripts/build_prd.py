#!/usr/bin/env python3
"""生成《合作执行需求说明书（PRD）》。

正文依据 2026年11月14日的合作原则沟通整理。
脚本只固定表述和版式，不新增保证、费用或到场义务。
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "deliverables"
FONT = "微软雅黑"

PINE = "0F5C58"
PINE_RGB = (0x0F, 0x5C, 0x58)
INK_RGB = (0x1B, 0x3A, 0x4B)
GOLD = "8C6A32"
GOLD_RGB = (0x8C, 0x6A, 0x32)
BODY = (0x24, 0x32, 0x38)
MUTED = (0x5E, 0x6A, 0x70)
WHITE = (0xFF, 0xFF, 0xFF)
LINE = "D5DDD8"
ALT = "F4F8F7"
PAPER = "F7F3EA"
CREAM = "FBF8F2"
RED_SOFT = "F8EFE8"

CONTENT_WIDTH = 17.8

MUST_PHRASES = [
    "不将一方此前对第三方作出的承诺转移给另一方承担",
    "不承担大会资金、场地费用、招商指标、嘉宾到场数量及活动举办效果等保证责任",
    "也不为任何第三方作出的承诺兜底",
    "不能作为我方必须完成的指标",
    "更不能形成对赌、保底或未达标赔偿责任",
    "不承诺任何具体赞助金额",
    "也不承担招商保底任务",
    "未经我方事先书面同意，不得要求我个人或复旦大学住房政策研究中心承担任何费用",
    "须另行协商确认",
    "不因高校及其他机构参与协办、资源协调而转移",
    "原则上自2026年11月15日起，不再随意调整组织架构",
    "不再随意增加主办、承办及支持单位",
    "辰山、水乡客厅、艺术学院",
    "不对嘉宾到场数量、招商金额及最终活动效果提供保证",
    "任何一方均可停止尚未开展的合作事项",
    "不得继续未经授权使用另一方或相关机构名义",
    "上海星智耀科技有限公司",
    "尽最大努力",
    "尽合理努力",
    "全力推动",
    "沉默不构成接受",
]


def set_run_font(run, size: float, bold: bool = False, color: tuple | None = None) -> None:
    run.font.name = FONT
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = RGBColor(*(color or BODY))
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        r_fonts.set(qn(attr), FONT)


def shade(cell, color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    node = OxmlElement("w:shd")
    node.set(qn("w:val"), "clear")
    node.set(qn("w:color"), "auto")
    node.set(qn("w:fill"), color)
    tc_pr.append(node)


def set_cell_margins(cell, top="50", bottom="50", left="80", right="80") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = OxmlElement("w:tcMar")
    for edge, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:w"), value)
        node.set(qn("w:type"), "dxa")
        tc_mar.append(node)
    tc_pr.append(tc_mar)
    valign = OxmlElement("w:vAlign")
    valign.set(qn("w:val"), "center")
    tc_pr.append(valign)


def set_table_widths(table, widths: list[float]) -> None:
    table.autofit = False
    table.allow_autofit = False
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    total = int(sum(widths) * 567)
    tbl_w.set(qn("w:w"), str(total))
    tbl_w.set(qn("w:type"), "dxa")
    grid = table._tbl.find(qn("w:tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tbl_pr.addnext(grid)
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(int(width * 567)))
        grid.append(col)
    for row in table.rows:
        tr_pr = row._tr.get_or_add_trPr()
        tr_pr.append(OxmlElement("w:cantSplit"))
        for index, cell in enumerate(row.cells):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(int(widths[index] * 567)))
            tc_w.set(qn("w:type"), "dxa")


def set_table_borders(table, color: str = LINE, sz: str = "4") -> None:
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), sz)
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), color)
        borders.append(node)
    tbl_pr = table._tbl.tblPr
    old = tbl_pr.find(qn("w:tblBorders"))
    if old is not None:
        tbl_pr.remove(old)
    tbl_pr.append(borders)


def set_row_height(row, centimeters: float) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tr_height = tr_pr.find(qn("w:trHeight"))
    if tr_height is None:
        tr_height = OxmlElement("w:trHeight")
        tr_pr.append(tr_height)
    tr_height.set(qn("w:val"), str(int(centimeters * 567)))
    tr_height.set(qn("w:hRule"), "atLeast")


def align_cell_top(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    valign = tc_pr.find(qn("w:vAlign"))
    if valign is None:
        valign = OxmlElement("w:vAlign")
        tc_pr.append(valign)
    valign.set(qn("w:val"), "top")


def write_lines(cell, lines: list[str], size: float = 9, bold: bool = False,
                color: tuple | None = None, fill: str | None = None,
                align: str = "left", space: float = 1) -> None:
    cell.text = ""
    for index, line in enumerate(lines):
        paragraph = cell.paragraphs[0] if index == 0 else cell.add_paragraph()
        paragraph.alignment = {
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
        }.get(align, WD_ALIGN_PARAGRAPH.LEFT)
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(space if index < len(lines) - 1 else 0)
        paragraph.paragraph_format.line_spacing = 1.08
        set_run_font(paragraph.add_run(line), size, bold, color)
    if fill:
        shade(cell, fill)
    set_cell_margins(cell)


def add_table(doc, headers: list[str], rows: list[list[str]], widths: list[float],
              size: float = 9, header_fill: str = PINE) -> None:
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    set_table_borders(table)
    set_table_widths(table, widths)
    for index, header in enumerate(headers):
        write_lines(
            table.rows[0].cells[index],
            [header],
            size=8.5,
            bold=True,
            color=WHITE,
            fill=header_fill,
            align="center",
        )
    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    mark = OxmlElement("w:tblHeader")
    mark.set(qn("w:val"), "true")
    tr_pr.append(mark)
    for row_index, row in enumerate(rows):
        fill = ALT if row_index % 2 == 0 else "FFFFFF"
        for col_index, value in enumerate(row):
            write_lines(
                table.rows[row_index + 1].cells[col_index],
                value.split("\n"),
                size=size,
                fill=fill,
                align="left",
            )
    gap(doc, 4)
    return table


def gap(doc, after: float = 6) -> None:
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = 1.0


def add_text(doc, text: str, size: float = 10.5, bold: bool = False,
             color: tuple | None = None, align: str = "left",
             before: float = 0, after: float = 6, keep: bool = False) -> None:
    paragraph = doc.add_paragraph()
    paragraph.alignment = {
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    paragraph.paragraph_format.line_spacing = 1.15
    paragraph.paragraph_format.keep_with_next = keep
    set_run_font(paragraph.add_run(text), size, bold, color)


def add_runs(doc, parts: list[tuple[str, bool]], size: float = 10.5,
             color: tuple | None = None, before: float = 0, after: float = 6,
             keep: bool = False) -> None:
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = 1.15
    paragraph.paragraph_format.keep_with_next = keep
    for text, bold in parts:
        set_run_font(paragraph.add_run(text), size, bold, color)


def add_heading(doc, text: str, size: float = 14) -> None:
    add_text(doc, text, size=size, bold=True, color=PINE_RGB, before=12, after=4, keep=True)


def add_rule(doc) -> None:
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(2)
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), PINE)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def callout(doc, title: str, lines: list[str], fill: str = PAPER, title_color=GOLD_RGB) -> None:
    table = doc.add_table(rows=1, cols=1)
    set_table_borders(table, color="E4D7C3", sz="4")
    set_table_widths(table, [CONTENT_WIDTH])
    cell = table.cell(0, 0)
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(2)
    paragraph.paragraph_format.keep_with_next = True
    set_run_font(paragraph.add_run(title), 10, True, title_color)
    for line in lines:
        item = cell.add_paragraph()
        item.paragraph_format.space_before = Pt(1)
        item.paragraph_format.space_after = Pt(1)
        item.paragraph_format.line_spacing = 1.12
        set_run_font(item.add_run(line), 10, False, INK_RGB)
    shade(cell, fill)
    set_cell_margins(cell, top="70", bottom="70", left="120", right="120")
    gap(doc, 6)


def add_page_field(paragraph) -> None:
    run = paragraph.add_run()
    set_run_font(run, 8, color=MUTED)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)


def setup(doc: Document) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(10.5)
    r_pr = normal.element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        r_fonts.set(qn(attr), FONT)

    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.7)
    section.bottom_margin = Cm(1.55)
    section.left_margin = Cm(1.6)
    section.right_margin = Cm(1.6)
    section.header_distance = Cm(0.45)
    section.footer_distance = Cm(0.35)

    header = section.header.paragraphs[0]
    header.paragraph_format.space_after = Pt(2)
    set_run_font(header.add_run("合作执行需求说明书 PRD"), 8, True, PINE_RGB)
    set_run_font(header.add_run("　　V1.0  ·  待书面确认  ·  2026年11月14日"), 8, False, MUTED)
    p_pr = header._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), PINE)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(footer.add_run("未书面确认前，沉默不构成接受新增义务　　第 "), 8, color=MUTED)
    add_page_field(footer)
    set_run_font(footer.add_run(" 页"), 8, color=MUTED)

    core = doc.core_properties
    core.author = "复旦大学住房政策研究中心相关工作对接代表"
    core.title = "2026银发科技与具身智能产业创新峰会合作执行需求说明书"
    core.subject = "合作原则确认 PRD V1.0"
    core.category = "合作确认"
    core.comments = "依据2026年11月14日合作原则整理。不构成结果保证，不替代授权文件。"


REQUIREMENTS = [
    {
        "id": "REQ-01",
        "title": "双方合作定位",
        "paragraphs": [
            "本次合作以共同推动大会成功举办为目标。",
            "我负责协调高校专家、养老科技、人工智能、机器人、金融投资等相关嘉宾及产业资源。你负责对接组委会、场地、会务组织及既有合作单位。",
            "双方各司其职、相互配合，但不将一方此前对第三方作出的承诺转移给另一方承担。",
        ],
        "accept": "对外口径和内部分工都按本条执行。已对第三方作出的承诺，仍由原承诺方自己承担。",
    },
    {
        "id": "REQ-02",
        "title": "复旦大学住房政策研究中心的责任边界",
        "paragraphs": [
            "我以复旦大学住房政策研究中心相关工作对接代表的身份参与筹备，在实际授权范围内协调学术和产业资源。",
            "中心的正式参与、对外署名及邀请函使用等，均须符合相应授权要求。",
            "复旦大学及复旦大学住房政策研究中心不承担大会资金、场地费用、招商指标、嘉宾到场数量及活动举办效果等保证责任，也不为任何第三方作出的承诺兜底。",
        ],
        "accept": "使用中心或复旦大学名义之前，先核对授权。邀请函、议程和合作说明中不写入本条已经排除的保证责任。",
    },
    {
        "id": "REQ-03",
        "title": "嘉宾邀请及三分之一目标",
        "paragraphs": [
            "目前详细议程及邀请函已经准备完成，但相关嘉宾尚未正式启动集中邀请。",
            "我愿意充分发挥个人和相关平台的资源优势，尽最大努力争取高质量嘉宾参会。嘉宾最终能否出席，取决于个人意愿、工作安排及档期等多方面因素。",
            "议程中约三分之一嘉宾的目标，可以作为双方努力争取的方向，但不能作为我方必须完成的指标，更不能形成对赌、保底或未达标赔偿责任。",
        ],
        "accept": "集中邀请在最终议程和邀请函书面确认后，按我方安排的方式、渠道和节奏推进。完成与否看邀请是否已经按此安排开展，不以到场人数计算。",
    },
    {
        "id": "REQ-04",
        "title": "赞助、招商及展位费用",
        "paragraphs": [
            "我可以协助对接相关企业、金融机构和潜在赞助方，但不承诺任何具体赞助金额，也不承担招商保底任务。",
            "对于展位、场地及其他活动成本，由实际经营或签约责任主体与组委会协商落实。",
            "未经我方事先书面同意，不得要求我个人或复旦大学住房政策研究中心承担任何费用。涉及新增费用或商业收益分配的，须另行协商确认。",
        ],
        "accept": "赞助沟通只写协助对接。任何费用落到个人或中心之前，都有一份事先的书面同意；没有这份同意，该项费用就不成立。",
    },
    {
        "id": "REQ-05",
        "title": "主办方及组织架构",
        "paragraphs": [
            "充分尊重上海星智耀科技有限公司作为主办单位的既有安排，以及你前期协调形成的合作关系。",
            "主办单位应承担其依法依约应承担的组织、运营和相关商务责任，不因高校及其他机构参与协办、资源协调而转移。",
            "现有指导、主办、承办、协办、支持及学术支持单位名单，在相关单位同意的前提下，以双方最终确认的版本为准。原则上自2026年11月15日起，不再随意调整组织架构。",
        ],
        "accept": "2026年11月15日前确认唯一对外名单。此后再改，须双方书面确认，并取得相关单位同意。本文不代替主办单位作出承诺。",
    },
    {
        "id": "REQ-06",
        "title": "后续增补单位与嘉宾",
        "paragraphs": [
            "如现有嘉宾邀请未达到预期，可以继续增补与大会主题高度相关的养老科技、人工智能、机器人、智慧医疗、科技服务及产业投资机构。",
            "既有辰山、水乡客厅、艺术学院等合作资源应得到合理尊重和妥善安排。",
            "新增单位原则上以嘉宾或产业合作方身份参与，不再随意增加主办、承办及支持单位，避免影响大会整体定位和邀请工作的连续性。",
        ],
        "accept": "增补对象与大会主题直接相关，身份默认是嘉宾或产业合作方。辰山、水乡客厅、艺术学院的既有安排不被无说明地拿掉。",
    },
    {
        "id": "REQ-07",
        "title": "工作决策及沟通机制",
        "paragraphs": [
            "双方确认最终议程、邀请函和组织架构后，各自负责相应资源协调工作。",
            "我负责的嘉宾邀请方式、渠道和沟通节奏，由我根据实际情况自主安排；你负责组委会及相关会务协调。",
            "重大事项提前沟通，具体执行以书面确认的信息为准，尽量减少反复修改方案和重复性沟通。",
        ],
        "accept": "执行时能指出对应的书面确认。组织架构、对外名义、费用、保底指标和对外议程版本，都按重大事项提前沟通。",
    },
    {
        "id": "REQ-08",
        "title": "合作风险及退出机制",
        "paragraphs": [
            "双方均应尽合理努力推动大会举办，但不对嘉宾到场数量、招商金额及最终活动效果提供保证。",
            "如果最终嘉宾、赞助或合作资源没有达到预期，双方应协商调整方案，不因此追究另一方未作出保证的结果责任。",
            "如后续出现新的保底要求、费用承担或超出授权的责任要求，且双方无法达成一致，任何一方均可停止尚未开展的合作事项，并妥善处理已形成的对外沟通，不得继续未经授权使用另一方或相关机构名义。",
        ],
        "accept": "未达预期时先协商调整方案。若触发退出，停止的是尚未开展的事项，并同步处理已经发出的名义和沟通，不单方面把机构名称留在后续材料里。",
    },
]


def build(path: Path) -> None:
    doc = Document()
    setup(doc)

    add_text(doc, "2026上海国际智慧康养产业博览会", 11, color=GOLD_RGB, align="center", before=6, after=2)
    add_text(doc, "合作执行需求说明书", 22, True, PINE_RGB, "center", after=1)
    add_text(doc, "AI重塑养老未来  ·  2026银发科技与具身智能产业创新峰会", 12, True, INK_RGB, "center", after=2)
    add_text(doc, "PRD  V1.0　　供双方一次性确认合作原则", 10.5, color=MUTED, align="center", after=2)
    add_rule(doc)
    add_text(
        doc,
        "2026年11月28日大会　　提出日期：2026年11月14日　　文件编号：PRD-YK-20261114",
        9,
        color=MUTED,
        align="center",
        before=4,
        after=8,
    )

    add_table(
        doc,
        ["项目", "内容"],
        [
            ["致", "赵同学"],
            ["提出", "复旦大学住房政策研究中心相关工作对接代表"],
            ["主办单位", "上海星智耀科技有限公司。本文尊重其既有安排，不代其作承诺。"],
            ["文件性质", "合作执行需求说明书。用于固定工作边界，不是赞助协议，不是担保文件，也不替代授权、组委会文件或主办单位合同。"],
            ["确认方式", "回复“确认 REQ-01 至 REQ-08”，或按编号写下保留意见。沉默不构成接受。"],
        ],
        [3.2, 14.6],
        size=9.5,
    )

    add_heading(doc, "一、为什么现在写成这一份", 13)
    add_text(
        doc,
        "结合2026年11月14日上午的沟通，合作原则需要一次性写清。截至当日，距离11月28日大会还有14天。我愿意发挥自身及相关平台的资源优势，全力推动大会；同时把各自做什么、不保证什么、什么必须另纸确认，一次说完，避免后面反复改口径。",
    )
    add_text(
        doc,
        "用语按这个意思理解：“全力推动”“尽最大努力”“尽合理努力”，都是指在授权、对方意愿和档期允许的范围内主动协调。三句话都不是对嘉宾到场、赞助金额或活动效果的结果保证。",
    )

    add_heading(doc, "二、八项原则速览", 13)
    add_text(doc, "下表供快速对照。完整表述和验收方式见第三节，两处不一致时以第三节全文为准。", size=10.5, after=4)
    add_table(
        doc,
        ["编号", "原则", "确认后的做法"],
        [
            ["REQ-01", "合作定位", "各做各的资源协调。一方对第三方的既有承诺，不转给另一方。"],
            ["REQ-02", "中心边界", "只在实际授权内协调。复旦大学及中心不担保资金、场地、招商、到场和效果，也不兜底。"],
            ["REQ-03", "嘉宾与三分之一", "尽最大努力邀请高质量嘉宾。约三分之一是努力方向，不是指标、对赌或赔偿责任。"],
            ["REQ-04", "赞助与费用", "可协助对接，不承诺金额，不接招商保底。费用须事先书面同意，收益分配另议。"],
            ["REQ-05", "主办与架构", "主办责任留在主办单位。名单以双方最终确认版为准，11月15日起不随意调整。"],
            ["REQ-06", "增补嘉宾与单位", "未达预期可以增补相关机构。新增默认是嘉宾或产业合作，不随意加主办、承办、支持。"],
            ["REQ-07", "决策与沟通", "议程、邀请函、组织架构确认后分头执行。重大事项提前沟通，执行以书面为准。"],
            ["REQ-08", "风险与退出", "尽合理努力，不对结果作保证。新的保底、费用或越权责任谈不拢，可以停止未开展事项。"],
        ],
        [2.2, 3.6, 12.0],
        size=9,
    )

    add_heading(doc, "三、需求说明", 13)
    add_text(doc, "每条先写确认后的规则，再写怎样算双方对齐。规则原文如下，不再另作缩写。", after=2)

    for item in REQUIREMENTS:
        add_text(
            doc,
            f"{item['id']}　　{item['title']}",
            size=12,
            bold=True,
            color=INK_RGB,
            before=10,
            after=3,
            keep=True,
        )
        for paragraph in item["paragraphs"]:
            add_text(doc, paragraph, size=10.5, before=0, after=3, keep=True)
        callout(doc, "怎样算对齐", [item["accept"]], fill=ALT, title_color=PINE_RGB)

    add_heading(doc, "四、谁负责什么", 13)
    add_text(doc, "分工只解决“谁去协调”。它不把主办单位依法依约应承担的组织、运营和商务责任转移给高校、中心或个人。", after=4)
    add_table(
        doc,
        ["角色", "负责", "不负责"],
        [
            [
                "对接代表\n（我方）",
                "在实际授权范围内，协调高校专家及养老科技、人工智能、机器人、金融投资等相关嘉宾和产业资源。\n嘉宾邀请的方式、渠道和节奏，由我方按实际情况自主安排。",
                "不承接你方或第三方此前作出的承诺。\n不保证到场人数、赞助金额和活动效果。\n未经事先书面同意，不承担费用。",
            ],
            [
                "赵同学\n（你方）",
                "对接组委会、场地、会务组织及既有合作单位。\n负责组委会及相关会务协调。",
                "不把组委会、主办单位或其他第三方的承诺改写成我方任务。\n不单方面要求我方或中心承担费用、保底和越权责任。",
            ],
            [
                "上海星智耀\n科技有限公司",
                "作为主办单位，承担其依法依约应承担的组织、运营和相关商务责任。",
                "上述责任不因高校及其他机构参与协办、资源协调而转移。\n本文不代主办单位签署或承诺。",
            ],
            [
                "复旦大学及\n住房政策研究中心",
                "正式参与、对外署名和邀请函使用，仅在相应授权范围内发生。",
                "不承担大会资金、场地费用、招商指标、嘉宾到场数量及活动举办效果等保证责任。\n不为任何第三方作出的承诺兜底。",
            ],
        ],
        [3.3, 7.4, 7.1],
        size=8.5,
    )

    add_heading(doc, "五、重大事项", 13)
    add_text(doc, "下列事项提前沟通，没有书面确认就不进入执行。", after=4)
    add_table(
        doc,
        ["事项", "书面确认前"],
        [
            ["调整指导、主办、承办、协办、支持或学术支持单位", "维持已确认版本。11月15日起不随意调整。"],
            ["对外使用复旦大学、中心或其他单位名义", "不落款、不署名、不把名义写入新的邀请和宣传。"],
            ["新增费用、商业收益分配、保底任务或量化指标", "不发生。三分之一嘉宾继续只作为努力方向。"],
            ["更换对外使用的议程或邀请函版本", "不发出未确认的版本。"],
            ["把新增单位写入主办、承办或支持", "新增单位维持嘉宾或产业合作方身份。"],
        ],
        [8.2, 9.6],
        size=9.5,
    )

    add_heading(doc, "六、时间", 13)
    add_table(
        doc,
        ["时点", "事项", "完成的含义"],
        [
            ["2026年11月14日", "提交本确认稿", "八项原则一次给出。对方尚未回复前，不视为已经接受。"],
            ["2026年11月15日起", "组织架构停止随意调整", "没有新的书面确认，就不再改名单。若当天仍未选定唯一版本，底稿A和底稿B都不当作终稿，也不另起第三套对外名单。"],
            ["书面确认之后", "分头推进", "我方按自主节奏开展嘉宾邀请。你方推进组委会和会务。执行只认书面确认过的议程、邀请函和组织架构。"],
            ["邀请未达预期时", "协商增补或调整", "可以增补主题高度相关的嘉宾和产业合作方。不把未达预期写成一方的结果责任。"],
            ["2026年11月28日", "大会举行", "按书面确认的方案举办。到场数量、招商金额和活动效果仍不设保证。"],
        ],
        [3.6, 4.4, 9.8],
        size=9,
    )

    add_heading(doc, "七、确认前需要拍板的五件事", 13)
    add_text(
        doc,
        "这五件事还没有在本文中被决定。未确认前，按“确认前怎么做”一列执行。两套组织架构和两个时段都来自已经写好的工作底稿，此处并列，是为了在11月15日前收成一套对外版本。",
        after=4,
    )
    add_table(
        doc,
        ["编号", "待确认", "时限", "确认前怎么做"],
        [
            ["OI-01", "组织架构的唯一对外版本", "2026年11月15日前", "不把底稿A或底稿B单独当作终稿。中心名称出现在底稿中，不等于署名和承办身份已经授权。"],
            ["OI-02", "对外议程用哪一个时段", "启动集中邀请前", "13:00—16:30 与 12:30—17:30 两套底稿都不单独发出。地点口径同为上海展览中心中央大厅会议室C，仍以组委会最终确认为准。"],
            ["OI-03", "中心正式参与、署名和邀请函使用的授权范围", "使用名义前", "不用中心或复旦大学名义落款，也不把中心写成已经确认的承办单位。"],
            ["OI-04", "水乡客厅的具体安排", "议程锁定前", "作为既有合作资源妥善安排。不默认写入主办、承办或支持。"],
            ["OI-05", "是否存在新增费用或商业收益分配", "该事项发生前", "没有另行书面确认，就不设立费用，也不做分配。"],
        ],
        [1.8, 4.6, 3.4, 8.0],
        size=8.5,
    )

    add_text(doc, "OI-01 底稿A：邀请函文本目前使用的名单", size=11, bold=True, color=INK_RGB, before=4, after=3, keep=True)
    add_table(
        doc,
        ["层级", "单位"],
        [
            ["指导单位", "上海现代服务业联合会、上海市养老服务和老龄产业协会、国家技术转移东部中心"],
            ["主办单位", "上海星智耀科技有限公司"],
            ["承办单位", "上海市杨浦区科技企业联合会、复旦大学住房政策研究中心"],
            ["协办单位", "长三角一体化示范区新发展建设有限公司"],
            ["支持单位", "上海辰山国家植物园、璞跃中国、上海视觉艺术学院"],
            ["学术支持单位", "复旦大学老龄研究院"],
        ],
        [3.4, 14.4],
        size=9,
    )
    add_text(doc, "OI-01 底稿B：分论坛执行议程首页目前使用的名单", size=11, bold=True, color=INK_RGB, before=2, after=3, keep=True)
    add_table(
        doc,
        ["层级", "单位"],
        [
            ["大会主办", "上海现代服务业联合会、上海市养老服务和老龄产业协会、上海丰之米会展有限公司、中国工商银行股份有限公司上海市分行"],
            ["分论坛主办", "上海星智耀科技有限公司"],
            ["联合承办", "上海市杨浦区科技企业联合会、复旦大学住房政策研究中心"],
            ["拟邀学术支持", "复旦大学老龄研究院"],
        ],
        [3.4, 14.4],
        size=9,
    )
    callout(
        doc,
        "两套底稿的差异，确认时需要收束",
        [
            "底稿A将上海星智耀科技有限公司列为主办单位，将相关协会等列为指导单位，并单列协办和支持单位。",
            "底稿B将协会、上海丰之米会展有限公司、中国工商银行股份有限公司上海市分行列为大会主办，将星智耀列为分论坛主办，首页未列协办和支持单位。",
            "两套底稿都出现了复旦大学住房政策研究中心。这只说明底稿写过这个名称，不说明正式参与、对外署名和邀请函使用已经获得授权。",
            "辰山与上海视觉艺术学院在底稿A中为支持单位。水乡客厅未进入两套单位名单，仍按 REQ-06 作为既有合作资源安排。",
            "对外只使用双方最终确认、且相关单位已经同意的一套名单。",
        ],
    )

    add_heading(doc, "八、请这样回复", 13)
    add_text(doc, "可以整页回复，也可以只回复编号。建议直接使用下面两种写法之一。", after=3)
    callout(
        doc,
        "回复示例",
        [
            "确认 REQ-01 至 REQ-08。OI-01 采用（底稿A / 底稿B / 另附名单）。OI-02 采用（13:00—16:30 / 12:30—17:30）。OI-03、OI-04、OI-05 另复。",
            "或：确认 REQ-01、REQ-02、REQ-03、REQ-04、REQ-06、REQ-07、REQ-08。REQ-05 的保留意见是：……。",
        ],
        fill=CREAM,
        title_color=INK_RGB,
    )
    add_text(
        doc,
        "未逐条回复的条款，不因对方继续推进会务而自动生效。与本文冲突的口头安排，在书面确认前不作为执行依据。",
    )

    add_heading(doc, "九、确认栏", 13)
    add_text(doc, "签字表示确认 REQ-01 至 REQ-08 的规则。OI-01 至 OI-05 仍以单独的书面回复为准，不因在本栏签字而被视为已经选定。", after=4)
    signature = add_table(
        doc,
        ["确认方", "确认内容", "签字", "日期"],
        [
            ["复旦大学住房政策研究中心相关工作对接代表", "提出本文。确认范围限于实际授权内的资源协调，不含机构保证责任。", "", ""],
            ["赵同学", "确认所收到的条款编号。保留意见写在编号之后。", "", ""],
        ],
        [5.2, 7.6, 2.6, 2.4],
        size=9,
    )
    for row in signature.rows[1:]:
        set_row_height(row, 1.45)
        for cell in row.cells:
            align_cell_top(cell)
    add_text(
        doc,
        "上海星智耀科技有限公司、组委会及相关单位不是本确认栏的签署方。需要其承担的事项，仍按其自身授权和签约文件办理。",
        size=9.5,
        color=MUTED,
        before=2,
    )

    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


def validate(path: Path) -> None:
    document = Document(path)
    text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                text += "\n" + cell.text
    missing = [phrase for phrase in MUST_PHRASES if phrase not in text]
    if missing:
        raise SystemExit("缺少关键表述：\n" + "\n".join(missing))
    forbidden = ["保证三分之一嘉宾到场", "赞助金额保底为", "中心承担场地费用", "未达标应当赔偿"]
    hit = [phrase for phrase in forbidden if phrase in text]
    if hit:
        raise SystemExit("出现不应存在的保证表述：\n" + "\n".join(hit))
    if len(document.tables) < 10:
        raise SystemExit(f"表格数量异常：{len(document.tables)}")
    print(f"已生成 {path}")
    print(f"段落 {len(document.paragraphs)}，表格 {len(document.tables)}，关键表述 {len(MUST_PHRASES)} 条均在文中")


def main() -> None:
    destination = OUT / "2026银发科技峰会_合作执行需求说明书PRD.docx"
    build(destination)
    validate(destination)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(exc, file=sys.stderr)
        raise
