"""把 P8—P11 工作总结的结构化底稿排成 4 页 PDF（仅文字与内表格，不含外框和水印）。

流程：底稿 JSON -> HTML/CSS -> 无头 Chrome 打印 -> 4 页 A4 PDF。
外框、原表水印和“非官方”标识由 build_review_pdf.py 在合成阶段叠加。

用法：
    python3 typeset_summary.py --content summary_p8_p11.json --font-dir ./fonts --out summary_pages.pdf
"""

from __future__ import annotations

import argparse
import html
import json
from pathlib import Path

import pymupdf

import layout as L
import prepare_fonts
from chrome_pdf import print_to_pdf


SHORT_PARAGRAPH_CHARS = int(L.TEXT_WIDTH / L.STYLES["p"].size) * 3


def _font_css(style: L.Style) -> str:
    family, weight = L.PDF_FONT[(style.family, style.bold)]
    return f'font-family:"{family}";font-weight:{weight};font-size:{style.size}pt;line-height:{style.leading}pt;'


def _block_css(name: str, extra: str = "") -> str:
    s = L.STYLES[name]
    return (
        f".{name}{{{_font_css(s)}margin:{s.before}pt 0 {s.after}pt 0;"
        f"text-indent:{s.indent_em}em;text-align:{s.align};{extra}}}"
    )


def build_css() -> str:
    table = L.STYLES["table"]
    bold_family, bold_weight = L.PDF_FONT[("song", True)]
    margin_right = L.PAGE_W - L.TEXT_RIGHT
    margin_bottom = L.PAGE_H - (L.FRAME_BOTTOM - L.PAD_BOTTOM)
    head_h = L.FORM_HEAD_BOTTOM - L.FRAME_TOP_FIRST
    return "\n".join(
        [
            f"@page{{size:{L.PAGE_W}pt {L.PAGE_H}pt;"
            f"margin:{L.FRAME_TOP + L.PAD_TOP}pt {margin_right}pt {margin_bottom}pt {L.TEXT_LEFT}pt;}}",
            f"@page :first{{margin-top:{L.FRAME_TOP_FIRST}pt;}}",
            "html,body{margin:0;padding:0;}",
            f"body{{color:{L.TEXT_COLOR};line-break:strict;orphans:2;widows:2;"
            "font-kerning:normal;font-synthesis:none;}",
            f".form-head{{height:{head_h}pt;margin-bottom:{L.PAD_TOP}pt;box-sizing:border-box;"
            "padding-top:5pt;overflow:hidden;}",
            _block_css("form_title", "letter-spacing:1.5pt;"),
            _block_css("form_note", "margin-top:3pt;white-space:nowrap;"),
            _block_css("title", "break-after:avoid;"),
            _block_css("h1", "break-after:avoid;"),
            _block_css("h2", "break-after:avoid;"),
            _block_css("p", "text-wrap:pretty;"),
            _block_css("note", "break-before:avoid;"),
            f'.p b{{font-family:"{bold_family}";font-weight:{bold_weight};}}',
            ".lead{break-after:avoid;}",
            ".short{break-inside:avoid;}",
            f"table{{{_font_css(table)}width:100%;border-collapse:collapse;table-layout:fixed;"
            f"margin:{table.before}pt 0 {table.after}pt 0;text-indent:0;}}",
            f"th,td{{border:{L.TABLE_LINE_WIDTH}pt solid rgb(51,51,51);"
            f"padding:{L.TABLE_CELL_PAD_Y}pt {L.TABLE_CELL_PAD_X}pt;vertical-align:middle;"
            "overflow-wrap:anywhere;}",
            f'th{{font-family:"{bold_family}";font-weight:{bold_weight};text-align:center;}}',
            "tr{break-inside:avoid;}",
            "thead{display:table-header-group;}",
            "td.c{text-align:center;}td.l{text-align:left;}",
            _block_css("sign"),
            f".sign-block{{margin-top:{L.SIGN_GAP + L.LAST_FRAME_PAD}pt;break-inside:avoid;break-before:avoid;}}",
            ".sign-date{padding-right:1em;}",
        ]
    )


def _esc(text: str) -> str:
    return html.escape(text, quote=False)


def _table_html(block: dict) -> str:
    cols = "".join(f'<col style="width:{w}%">' for w in block["widths"])
    head = "".join(f"<th>{_esc(h)}</th>" for h in block["header"])
    rows = []
    for row in block["rows"]:
        cells = "".join(
            f'<td class="{a}">{_esc(cell)}</td>' for a, cell in zip(block["align"], row, strict=True)
        )
        rows.append(f"<tr>{cells}</tr>")
    return (
        f'<table id="{block["id"]}"><colgroup>{cols}</colgroup>'
        f"<thead><tr>{head}</tr></thead><tbody>{''.join(rows)}</tbody></table>"
    )


def sign_line(content: dict) -> str:
    return f'{content["sign"]["label"]}　{"_" * 21}'


def ordered_texts(content: dict) -> list[str]:
    """底稿里会被排进版面的文字，按版面先后排列；表格按行、行内按列。"""
    pieces = [content["form_title"], content["form_note"]]
    for block in content["blocks"]:
        if block["type"] == "p":
            pieces.append("".join(text for _, text in block["runs"]))
        elif block["type"] == "table":
            pieces.extend(block["header"])
            pieces.extend(cell for row in block["rows"] for cell in row)
        else:
            pieces.append(block["text"])
    pieces.extend([sign_line(content), content["sign"]["date"]])
    return pieces


def check_glyph_coverage(content: dict, font_dir: Path) -> None:
    text = "".join(ordered_texts(content))
    for path in prepare_fonts.prepare(font_dir).values():
        missing = prepare_fonts.missing_chars(path, text)
        if missing:
            raise RuntimeError(f"{path.name} 缺少字形：{''.join(missing)}")


def build_html(content: dict, font_dir: Path) -> str:
    parts = [
        '<div class="form-head">'
        f'<div class="form_title">{_esc(content["form_title"])}</div>'
        f'<div class="form_note">{_esc(content["form_note"])}</div>'
        "</div>"
    ]
    previous_kind = ""
    for block in content["blocks"]:
        kind = block["type"]
        if kind in ("title", "h1", "h2", "note"):
            parts.append(f'<div class="{kind}">{_esc(block["text"])}</div>')
        elif kind == "p":
            runs = block["runs"]
            inner = "".join(f"<b>{_esc(t)}</b>" if flag == "b" else _esc(t) for flag, t in runs)
            classes = ["p"]
            # 纯加粗的引导行、以及二级标题后的第一段，都要和下一段留在同一页，
            # 这样页末不会只剩“标题 + 一行”。
            if all(flag == "b" for flag, _ in runs) or previous_kind == "h2":
                classes.append("lead")
            # 三行以内的段落不跨页，避免页首或页末只剩一行。
            if sum(len(t) for _, t in runs) <= SHORT_PARAGRAPH_CHARS:
                classes.append("short")
            parts.append(f'<div class="{" ".join(classes)}">{inner}</div>')
        elif kind == "table":
            parts.append(_table_html(block))
        else:
            raise ValueError(f"未知的块类型：{kind}")
        previous_kind = kind
    parts.append(
        '<div class="sign-block">'
        f'<div class="sign sign-line">{_esc(sign_line(content))}</div>'
        f'<div class="sign sign-date">{_esc(content["sign"]["date"])}</div>'
        "</div>"
    )
    fonts = prepare_fonts.font_face_css(font_dir, L.PDF_FONT)
    return (
        '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">'
        f"<style>{fonts}\n{build_css()}</style></head>"
        f"<body>{''.join(parts)}</body></html>"
    )


def _spans(page: pymupdf.Page) -> list[dict]:
    return [
        s
        for b in page.get_text("dict")["blocks"]
        if b["type"] == 0
        for line in b["lines"]
        for s in line["spans"]
        if s["text"].strip()
    ]


def measure(pdf_path: Path, sign_label: str) -> dict:
    """量出每页文字的四边，并定出末页收框位置。"""
    doc = pymupdf.open(pdf_path)
    pages = []
    for page in doc:
        spans = _spans(page)
        pages.append(
            {
                "top": round(min(s["bbox"][1] for s in spans), 2),
                "bottom": round(max(s["bbox"][3] for s in spans), 2),
                "left": round(min(s["bbox"][0] for s in spans), 2),
                "right": round(max(s["bbox"][2] for s in spans), 2),
            }
        )
    last = doc[-1]
    hits = last.search_for(sign_label)
    if not hits:
        raise RuntimeError("末页没有找到落款：正文溢出，或落款被挤到了下一页")
    sign_top = min(r.y0 for r in hits)
    body_bottom = max(s["bbox"][3] for s in _spans(last) if s["bbox"][3] < sign_top)
    return {
        "page_count": doc.page_count,
        "pages": pages,
        "last_body_bottom": round(body_bottom, 2),
        "sign_top": round(sign_top, 2),
        "last_frame_bottom": round(body_bottom + L.LAST_FRAME_PAD, 2),
    }


def typeset(content: dict, out_pdf: Path, font_dir: Path) -> dict:
    check_glyph_coverage(content, font_dir)
    print_to_pdf(build_html(content, font_dir), out_pdf)
    metrics = measure(out_pdf, content["sign"]["label"])
    if metrics["page_count"] != L.SUMMARY_PAGE_COUNT:
        raise RuntimeError(
            f"工作总结排出了 {metrics['page_count']} 页，必须正好 {L.SUMMARY_PAGE_COUNT} 页才能保持 14 页页序"
        )
    if metrics["sign_top"] - metrics["last_frame_bottom"] < 6:
        raise RuntimeError("落款与末页外框之间的间距不足")
    return metrics


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--content", required=True, type=Path)
    parser.add_argument("--font-dir", required=True, type=Path, help="prepare_fonts.py 的字体目录")
    parser.add_argument("--out", required=True, type=Path)
    parser.add_argument("--html-out", type=Path, help="同时保存排版用的 HTML，便于检查")
    args = parser.parse_args()
    content = json.loads(args.content.read_text(encoding="utf-8"))
    if args.html_out:
        args.html_out.write_text(build_html(content, args.font_dir), encoding="utf-8")
    metrics = typeset(content, args.out, args.font_dir)
    print(json.dumps(metrics, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
