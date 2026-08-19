#!/usr/bin/env python3
"""生成给律师阅卷用的完整经过说明（Markdown + Word）。不是律师函。"""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from content import LAWYER_BRIEF, LAWYER_SECTIONS

CHINESE_FONT = "宋体"
HEADING_FONT = "黑体"
ACCENT = RGBColor(0x0B, 0x2F, 0x5B)
MUTED = RGBColor(0x5C, 0x6B, 0x7A)
RED = RGBColor(0xA6, 0x3D, 0x2F)

MD_NAME = "给律师的事故完整经过说明_2026-08-19.md"
DOCX_NAME = "青岛红枫路交通事故_给律师的完整经过说明_20260819.docx"


def _set_run_font(run, font_name: str, size: float, bold: bool = False, color: RGBColor | None = None) -> None:
    run.font.name = font_name
    run.font.size = Pt(size)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), font_name)
    rfonts.set(qn("w:ascii"), font_name)
    rfonts.set(qn("w:hAnsi"), font_name)
    rfonts.set(qn("w:cs"), font_name)


def _add_paragraph(
    doc: Document,
    text: str = "",
    *,
    font: str = CHINESE_FONT,
    size: float = 12,
    bold: bool = False,
    alignment=WD_ALIGN_PARAGRAPH.LEFT,
    first_line_indent: float | None = None,
    space_after: float = 6,
    space_before: float = 0,
    line_spacing: float = 1.5,
    color: RGBColor | None = None,
):
    p = doc.add_paragraph()
    p.alignment = alignment
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing = line_spacing
    if first_line_indent is not None:
        pf.first_line_indent = Cm(first_line_indent)
    if text:
        run = p.add_run(text)
        _set_run_font(run, font, size, bold=bold, color=color)
    return p


def render_markdown() -> str:
    b = LAWYER_BRIEF
    lines = [
        f"# {b['title']}",
        "",
        b["subtitle"],
        "",
        f"> {b['file_note']}",
        "",
        "## 请律师先做这几件事",
        "",
    ]
    for i, item in enumerate(b["ask"], 1):
        lines.append(f"{i}. {item}")
    lines.append("")
    for sec in LAWYER_SECTIONS:
        lines.append(f"## {sec['h']}")
        lines.append("")
        for p in sec["paras"]:
            lines.append(p)
            lines.append("")
    lines.append("## 十、现有材料与仍缺文件")
    lines.append("")
    lines.append("已经可以交给律师的：")
    lines.append("")
    for x in b["have"]:
        lines.append(f"- {x}")
    lines.append("")
    lines.append("请指导家属尽快补齐的：")
    lines.append("")
    for x in b["need"]:
        lines.append(f"- {x}")
    lines.append("")
    lines.append("重新生成：`python3 scripts/build_all.py`。3D 示范动画留作待用，不是证据。")
    lines.append("")
    return "\n".join(lines)


def build_markdown(output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(render_markdown(), encoding="utf-8")


def build_document(output_path: Path) -> None:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(2.4)
        section.bottom_margin = Cm(2.4)
        section.left_margin = Cm(2.6)
        section.right_margin = Cm(2.6)
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)

    normal_style = doc.styles["Normal"]
    normal_style.font.name = CHINESE_FONT
    normal_style.font.size = Pt(12)
    rpr = normal_style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), CHINESE_FONT)

    b = LAWYER_BRIEF
    _add_paragraph(
        doc,
        b["title"],
        font=HEADING_FONT,
        size=18,
        bold=True,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=6,
        line_spacing=1.3,
        color=ACCENT,
    )
    _add_paragraph(
        doc,
        b["subtitle"],
        font=HEADING_FONT,
        size=11,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=8,
        line_spacing=1.3,
        color=MUTED,
    )
    _add_paragraph(doc, b["file_note"], bold=True, color=RED, space_after=10)

    _add_paragraph(doc, "请律师先做这几件事", font=HEADING_FONT, size=14, bold=True, space_before=8, space_after=8, color=ACCENT, line_spacing=1.3)
    for item in b["ask"]:
        _add_paragraph(doc, f"• {item}", first_line_indent=0.37, space_after=3, line_spacing=1.4)

    for sec in LAWYER_SECTIONS:
        _add_paragraph(doc, sec["h"], font=HEADING_FONT, size=14, bold=True, space_before=14, space_after=8, color=ACCENT, line_spacing=1.3)
        for p in sec["paras"]:
            _add_paragraph(doc, p, first_line_indent=0.74, space_after=6)

    _add_paragraph(doc, "十、现有材料与仍缺文件", font=HEADING_FONT, size=14, bold=True, space_before=14, space_after=8, color=ACCENT, line_spacing=1.3)
    _add_paragraph(doc, "已经可以交给律师的：", bold=True, space_after=4)
    for x in b["have"]:
        _add_paragraph(doc, f"• {x}", first_line_indent=0.37, space_after=3, line_spacing=1.4)
    _add_paragraph(doc, "请指导家属尽快补齐的：", bold=True, space_before=8, space_after=4)
    for x in b["need"]:
        _add_paragraph(doc, f"• {x}", first_line_indent=0.37, space_after=3, line_spacing=1.4)

    _add_paragraph(
        doc,
        "本文根据 2026 年 8 月 14–19 日材料整理。重新生成：python3 scripts/build_all.py。3D 示范动画留作待用，不是证据。",
        size=10.5,
        color=MUTED,
        space_before=12,
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    out = root / "deliverables"
    build_markdown(out / MD_NAME)
    build_document(out / DOCX_NAME)
    print(f"Wrote {out / MD_NAME}")
    print(f"Wrote {out / DOCX_NAME}")


if __name__ == "__main__":
    main()
