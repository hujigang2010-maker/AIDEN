#!/usr/bin/env python3
"""在嘉宾均可由复旦大学请到的前提下，生成主旨分享与圆桌的最优议程。"""

import shutil
import subprocess
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "deliverables"
FONT = "微软雅黑"
PINE = "0F5C58"
PINE_RGB = (0x0F, 0x5C, 0x58)
INK_RGB = (0x1B, 0x3A, 0x4B)
GOLD_RGB = (0x8C, 0x6A, 0x32)
ALT = "F4F8F7"
LINE = "D5DDD8"

ORGANIZERS = [
    ("大会主办", "上海现代服务业联合会"),
    ("大会主办", "上海市养老服务和老龄产业协会"),
    ("大会主办", "上海丰之米会展有限公司"),
    ("大会主办", "中国工商银行股份有限公司上海市分行"),
    ("分论坛主办", "上海星智耀科技有限公司"),
    ("联合承办", "上海市杨浦区科技企业联合会"),
    ("联合承办", "复旦大学住房政策研究中心"),
    ("学术支持", "复旦大学老龄研究院"),
]

KEYNOTES = [
    ("13:18—13:38", "彭希哲", "复旦大学老龄研究院院长、文科资深教授", "《“十五五”银发经济与长寿社会发展的新机遇》"),
    ("13:38—14:00", "顾捷", "傅利叶智能创始人、CEO", "《从康复机器人到具身智能：AI如何改变未来养老》"),
    ("14:00—14:18", "邱锡鹏", "复旦大学教授、上海创智学院全时导师", "《从大模型到智能体：AI康养的下一代交互技术》"),
    ("14:18—14:36", "王学集", "涂鸦智能创始人兼CEO", "AI Home与未来智慧养老家庭"),
]

PANEL = [
    ("第一席", "胡继刚", "高校成果转化、智慧居住与资本协同"),
    ("第二席", "邓泰华　智元机器人创始人、董事长兼CEO", "人形机器人产业化与进院条件"),
    ("第三席", "胡湛　复旦大学老龄研究院副院长、银龄未来实验室主任", "真实需求与AI向善"),
    ("第四席", "邱建伟　泰康之家首席执行官", "机构采购、试点与运营"),
    ("第五席", "钱鹏鹤　移远通信董事长兼CEO", "智能感知与养老物联网络"),
    ("第六席", "顾建忠　上海银行董事长", "银发金融、设备融资与采购支付"),
]


def validate() -> None:
    keynote_names = [row[1] for row in KEYNOTES]
    panel_names = [row[1] for row in PANEL]
    assert len(keynote_names) == len(set(keynote_names))
    assert len(panel_names) == 6
    assert len(panel_names) == len(set(panel_names))
    assert not set(keynote_names) & set(panel_names)
    assert KEYNOTES[0][0].startswith("13:18")
    assert KEYNOTES[-1][0].endswith("14:36")


def set_run_font(run, size: float, bold: bool = False, color: tuple | None = None) -> None:
    run.font.name = FONT
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
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


def set_cell_margins(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = OxmlElement("w:tcMar")
    for edge, value in (("top", "28"), ("left", "70"), ("bottom", "28"), ("right", "70")):
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
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(int(sum(widths) * 567)))
    tbl_w.set(qn("w:type"), "dxa")
    grid = tbl.find(qn("w:tblGrid"))
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
        for index, cell in enumerate(row.cells):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(int(widths[index] * 567)))
            tc_w.set(qn("w:type"), "dxa")


def set_table_borders(table) -> None:
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "4")
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), LINE)
        borders.append(node)
    table._tbl.tblPr.append(borders)


def write_cell(cell, text: str, size: float = 10.5, bold: bool = False, color: tuple | None = None,
               fill: str | None = None, align: str = "left") -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if align == "center" else WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(1)
    paragraph.paragraph_format.space_after = Pt(1)
    paragraph.paragraph_format.line_spacing = 1.08
    set_run_font(paragraph.add_run(text), size, bold, color)
    if fill:
        shade(cell, fill)
    set_cell_margins(cell)


def add_table(doc, headers: list[str], rows: list[list[str]], widths: list[float]) -> None:
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    set_table_borders(table)
    set_table_widths(table, widths)
    for index, header in enumerate(headers):
        write_cell(table.rows[0].cells[index], header, size=10, bold=True, color=(255, 255, 255), fill=PINE, align="center")
    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    header_mark = OxmlElement("w:tblHeader")
    header_mark.set(qn("w:val"), "true")
    tr_pr.append(header_mark)
    for row_index, row in enumerate(rows):
        fill = ALT if row_index % 2 == 0 else "FFFFFF"
        for col_index, value in enumerate(row):
            write_cell(
                table.rows[row_index + 1].cells[col_index],
                value,
                size=10.5,
                fill=fill,
                align="center" if col_index == 0 else "left",
            )


def add_line(doc, text: str, size: float = 12, bold: bool = False, color: tuple | None = None,
             align: str = "left", before: float = 0, after: float = 4, keep: bool = False):
    paragraph = doc.add_paragraph()
    paragraph.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
    }[align]
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = 1.15
    paragraph.paragraph_format.keep_with_next = keep
    set_run_font(paragraph.add_run(text), size, bold, color)
    return paragraph


def build(path: Path) -> None:
    doc = Document()
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
    section.top_margin = Cm(1.35)
    section.bottom_margin = Cm(1.2)
    section.left_margin = Cm(1.7)
    section.right_margin = Cm(1.7)

    add_line(doc, "2026上海国际智慧康养产业博览会", 12, color=GOLD_RGB, align="center", after=2)
    add_line(doc, "AI重塑养老未来", 18, True, PINE_RGB, "center", after=0)
    add_line(doc, "2026银发科技与具身智能产业创新峰会", 13, True, INK_RGB, "center", after=1)
    add_line(doc, "主旨分享与圆桌议程", 14, True, PINE_RGB, "center", after=1)
    add_line(doc, "2026年11月28日（周六）　上海展览中心中央大厅会议室C", 11, align="center", after=4)

    add_line(doc, "主办方", 16, True, PINE_RGB, before=2, after=3)
    grouped: dict[str, list[str]] = {}
    for role, name in ORGANIZERS:
        grouped.setdefault(role, []).append(name)
    for role, names in grouped.items():
        add_line(doc, f"{role}：{'、'.join(names)}", 11, after=1)
    add_line(doc, "嘉宾由复旦大学出面邀请。", 11, before=1, after=2)

    add_line(doc, "主旨分享", 15, True, PINE_RGB, before=2, after=1, keep=True)
    add_line(doc, "13:18—14:36", 11, color=GOLD_RGB, after=3, keep=True)
    add_table(
        doc,
        ["时间", "嘉宾", "职务", "题目"],
        [list(row) for row in KEYNOTES],
        [3.0, 2.4, 5.6, 6.6],
    )

    add_line(doc, "圆桌", 15, True, PINE_RGB, before=4, after=1, keep=True)
    add_line(doc, "15:00—15:45", 11, color=GOLD_RGB, after=1, keep=True)
    add_line(doc, "机器人进入养老院：技术成熟了，商业模式成熟了吗？", 12, True, INK_RGB, after=1, keep=True)
    add_line(doc, "主持：徐洁平", 11, after=3, keep=True)
    add_table(
        doc,
        ["席位", "嘉宾", "讨论重点"],
        [[seat, name, focus] for seat, name, focus in PANEL],
        [2.2, 8.2, 7.2],
    )

    core = doc.core_properties
    core.title = "2026银发科技与具身智能产业创新峰会主旨与圆桌议程"
    core.subject = "嘉宾均可出席时的最优主旨与圆桌"
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


def export_pdf(docx_path: Path) -> None:
    soffice = shutil.which("soffice")
    if not soffice:
        return
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(docx_path.parent), str(docx_path)],
        check=True,
    )


def main() -> None:
    validate()
    docx_path = OUT / "2026上海国际智慧康养产业博览会_主旨与圆桌最优议程.docx"
    build(docx_path)
    export_pdf(docx_path)
    print(docx_path)


if __name__ == "__main__":
    main()
