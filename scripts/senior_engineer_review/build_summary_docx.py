"""把 P8—P11 工作总结底稿生成可编辑的 DOCX（宋体 / 宋体加粗，表头黑体）。

字号、行距、缩进与 PDF 版用同一份 layout.py；不含任何印章或签名图片，
页脚带“非官方”标识。DOCX 用页面边框近似外框：末页边框不会像 PDF 那样在正文结束处收口。
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

import layout as L
import textcheck
import typeset_summary

ALIGN = {
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def _set_font(run, style: L.Style, bold: bool | None = None) -> None:
    name = L.DOCX_FONT[style.family]
    run.font.name = name
    run.font.size = Pt(style.size)
    run.font.bold = style.bold if bold is None else bold
    fonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        fonts.set(qn(attr), name)


def _format(paragraph, style: L.Style, keep_with_next: bool = False) -> None:
    fmt = paragraph.paragraph_format
    fmt.alignment = ALIGN[style.align]
    fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    fmt.line_spacing = Pt(style.leading)
    fmt.space_before = Pt(style.before)
    fmt.space_after = Pt(style.after)
    fmt.keep_with_next = keep_with_next
    fmt.widow_control = True
    if style.indent_em:
        # 用“字符”为单位的首行缩进，改字号后仍是两个字。
        ind = paragraph._p.get_or_add_pPr().get_or_add_ind()
        ind.set(qn("w:firstLineChars"), str(int(style.indent_em * 100)))
        ind.set(qn("w:firstLine"), str(int(style.indent_em * style.size * 20)))


def _paragraph(container, name: str, runs: list[tuple[str, str]], keep_with_next: bool = False):
    style = L.STYLES[name]
    paragraph = container.add_paragraph()
    _format(paragraph, style, keep_with_next)
    for flag, text in runs:
        _set_font(paragraph.add_run(text), style, bold=True if flag == "b" else None)
    return paragraph


def _borders(parent, tag: str, edges: tuple[str, ...], size_eighths: int, extra: dict | None = None) -> None:
    borders = OxmlElement(tag)
    for key, value in (extra or {}).items():
        borders.set(qn(key), value)
    for edge in edges:
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), str(size_eighths))
        element.set(qn("w:space"), "0" if tag != "w:pgBorders" else "8")
        element.set(qn("w:color"), "333333")
        borders.append(element)
    parent.append(borders)


def _table(document, block: dict) -> None:
    style = L.STYLES["table"]
    table = document.add_table(rows=1 + len(block["rows"]), cols=len(block["header"]))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    properties = table._tbl.tblPr
    _borders(properties, "w:tblBorders", ("top", "left", "bottom", "right", "insideH", "insideV"), int(L.TABLE_LINE_WIDTH * 8))
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    properties.append(layout)
    widths = [L.TEXT_WIDTH * w / 100 for w in block["widths"]]
    # 列宽要同时写进表格网格和每个单元格，Word 和 LibreOffice 才都按它排。
    table_width = OxmlElement("w:tblW")
    table_width.set(qn("w:w"), str(round(L.TEXT_WIDTH * 20)))
    table_width.set(qn("w:type"), "dxa")
    for old in properties.findall(qn("w:tblW")):
        properties.remove(old)
    properties.append(table_width)
    margins = OxmlElement("w:tblCellMar")
    for edge, value in (("top", L.TABLE_CELL_PAD_Y), ("left", L.TABLE_CELL_PAD_X), ("bottom", L.TABLE_CELL_PAD_Y), ("right", L.TABLE_CELL_PAD_X)):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:w"), str(round(value * 20)))
        element.set(qn("w:type"), "dxa")
        margins.append(element)
    properties.append(margins)
    for column, width in zip(table._tbl.tblGrid.findall(qn("w:gridCol")), widths, strict=True):
        column.set(qn("w:w"), str(round(width * 20)))
    cell_style = L.Style(style.family, style.size, False, style.leading, align="center")
    for r, row in enumerate([block["header"], *block["rows"]]):
        tr_pr = table.rows[r]._tr.get_or_add_trPr()
        tr_pr.append(OxmlElement("w:cantSplit"))
        if r == 0:
            tr_pr.append(OxmlElement("w:tblHeader"))
        for c, text in enumerate(row):
            cell = table.cell(r, c)
            cell.width = Pt(widths[c])
            cell.vertical_alignment = 1
            paragraph = cell.paragraphs[0]
            _format(paragraph, cell_style)
            paragraph.paragraph_format.alignment = ALIGN["center" if r == 0 or block["align"][c] == "c" else "left"]
            _set_font(paragraph.add_run(text), cell_style, bold=r == 0)
    spacer = document.add_paragraph()
    spacer.paragraph_format.space_after = Pt(0)
    spacer.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    spacer.paragraph_format.line_spacing = Pt(style.after + 1)


def build(content: dict, out_docx: Path) -> None:
    document = Document()
    section = document.sections[0]
    section.page_width, section.page_height = Pt(L.PAGE_W), Pt(L.PAGE_H)
    section.left_margin = Pt(L.TEXT_LEFT)
    section.right_margin = Pt(L.PAGE_W - L.TEXT_RIGHT)
    section.top_margin = Pt(L.FRAME_TOP + L.PAD_TOP)
    section.bottom_margin = Pt(L.PAGE_H - L.FRAME_BOTTOM + L.PAD_BOTTOM)
    section.header_distance = Pt(8)
    section.footer_distance = Pt(10)
    _borders(
        section._sectPr,
        "w:pgBorders",
        ("top", "left", "bottom", "right"),
        int(L.FRAME_LINE_WIDTH * 8),
        {"w:offsetFrom": "text"},
    )

    normal = document.styles["Normal"]
    normal.font.name = L.DOCX_FONT["song"]
    normal.font.size = Pt(L.STYLES["p"].size)
    normal.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), L.DOCX_FONT["song"])

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    mark = footer.add_run(L.MARK_FOOTER)
    _set_font(mark, L.Style("hei", L.MARK_FOOTER_SIZE, True, L.MARK_FOOTER_SIZE))
    mark.font.color.rgb = RGBColor(*L.MARK_COLOR)

    _paragraph(document, "form_title", [("", content["form_title"])])
    _paragraph(document, "form_note", [("", content["form_note"])])
    previous = ""
    for block in content["blocks"]:
        kind = block["type"]
        if kind in ("title", "h1", "h2"):
            _paragraph(document, kind, [("", block["text"])], keep_with_next=True)
        elif kind == "note":
            _paragraph(document, kind, [("", block["text"])])
        elif kind == "p":
            runs = [(flag, text) for flag, text in block["runs"]]
            lead = all(flag == "b" for flag, _ in runs) or previous == "h2"
            _paragraph(document, "p", runs, keep_with_next=lead)
        elif kind == "table":
            _table(document, block)
        else:
            raise ValueError(f"未知的块类型：{kind}")
        previous = kind

    sign = _paragraph(document, "sign", [("", typeset_summary.sign_line(content))])
    sign.paragraph_format.space_before = Pt(L.SIGN_GAP)
    _paragraph(document, "sign", [("", content["sign"]["date"] + "　")])

    document.core_properties.title = "个人专业技术工作总结（非官方排版审阅版，可编辑）"
    document.core_properties.subject = "内部排版审阅用，不具官方效力；不含任何印章或电子签名。"
    document.core_properties.author = ""
    document.save(out_docx)


def check(content: dict, docx_path: Path) -> dict:
    """重新读出 DOCX 正文，与底稿逐字比对，并统计实际使用的字体。"""
    document = Document(docx_path)
    body = document.element.body
    texts, fonts = [], set()
    for element in body.iter(qn("w:t")):
        texts.append(element.text or "")
    for element in body.iter(qn("w:rFonts")):
        fonts.add(element.get(qn("w:eastAsia")))
    expected = textcheck.normalize("".join(typeset_summary.ordered_texts(content)))
    diff = textcheck.char_diff(expected, textcheck.normalize("".join(texts)))
    pictures = len(list(body.iter(qn("w:drawing")))) + len(list(body.iter(qn("w:pict"))))
    return {"text_vs_content": diff, "east_asia_fonts": sorted(f for f in fonts if f), "pictures": pictures}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--content", required=True, type=Path)
    parser.add_argument("--out", required=True, type=Path)
    args = parser.parse_args()
    content = json.loads(args.content.read_text(encoding="utf-8"))
    build(content, args.out)
    report = check(content, args.out)
    print(json.dumps(report, ensure_ascii=False, indent=1))
    if report["text_vs_content"]["missing"] or report["text_vs_content"]["extra"] or report["pictures"]:
        raise SystemExit("DOCX 与底稿不一致，或混入了图片")


if __name__ == "__main__":
    main()
