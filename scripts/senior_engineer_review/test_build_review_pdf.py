"""用完全虚构的底稿检验组装流程（不含任何真实申报内容）。

运行：python3 test_build_review_pdf.py --font-dir <字体目录>
虚构底稿模仿真实底稿的结构：P3—P12 每页是“正文表单对象 + 共用水印表单对象”，
P13 带一个签名域。
"""

from __future__ import annotations

import argparse
import json
import tempfile
from pathlib import Path

import pikepdf
import pymupdf

import build_review_pdf as B
import layout as L
import overlays
import textcheck
import typeset_summary

SENTENCE = "Synthetic review text for layout checks only; it carries no real application content. "


def synthetic_content(paragraphs: int) -> dict:
    blocks: list[dict] = [{"type": "title", "text": "Synthetic Summary"}]
    for i in range(paragraphs):
        if i % 6 == 0:
            blocks.append({"type": "h1", "text": f"Section {i // 6 + 1}"})
        if i % 3 == 0:
            blocks.append({"type": "h2", "text": f"{i // 3 + 1}. Item heading"})
        blocks.append({"type": "p", "runs": [["b", "Lead: "], ["", SENTENCE * 4]]})
    blocks.append(
        {
            "type": "table",
            "id": "t",
            "widths": [30, 70],
            "align": ["c", "l"],
            "header": ["Key", "Value"],
            "rows": [["A", "alpha"], ["B", "beta"]],
        }
    )
    return {
        "form_title": "Synthetic Form Heading",
        "form_note": "(synthetic note)",
        "blocks": blocks,
        "sign": {"label": "Signature", "date": "Y    M    D"},
    }


def four_page_content(font_dir: Path, work: Path) -> dict:
    for paragraphs in range(30, 70):
        content = synthetic_content(paragraphs)
        try:
            typeset_summary.typeset(content, work / "probe.pdf", font_dir)
        except RuntimeError:
            continue
        return content
    raise AssertionError("找不到正好排成 4 页的虚构内容")


def _text_stream(lines: list[str]) -> bytes:
    ops = [b"1 1 1 rg 0 0 595.28 841.89 re f 0 0 0 rg BT /F1 6 Tf"]
    y = 800.0
    for line in lines:
        ops.append(f"1 0 0 1 40 {y:.2f} Tm ({line}) Tj".encode("ascii"))
        y -= 7
    ops.append(b"ET")
    return b"\n".join(ops)


def _p3_stream() -> bytes:
    ops = [b"0.75 0 0 -0.75 0 841.89 cm BT /F1 12 Tf"]
    for y in B.P3_LABEL_YS:
        ops.append(f"1 0 0 -1 {B.P3_LABEL_X} {y} Tm (L) Tj".encode("ascii"))
    ops.append(b"1 0 0 -1 200 300 Tm (untouched) Tj ET")
    return b"\n".join(ops)


def synthetic_base(content: dict, out_pdf: Path) -> None:
    # 全角空格只是落款的占位，比对时本来就不计入。
    text = "".join(c for c in "".join(typeset_summary.ordered_texts(content)) if c.isascii())
    chunks = [
        text[i : i + 150].replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
        for i in range(0, len(text), 150)
    ]
    per_page = -(-len(chunks) // L.SUMMARY_PAGE_COUNT)

    pdf = pikepdf.new()
    font = pdf.make_indirect(pikepdf.Dictionary(Type=pikepdf.Name.Font, Subtype=pikepdf.Name.Type1, BaseFont=pikepdf.Name.Helvetica))
    resources = pikepdf.Dictionary(Font=pikepdf.Dictionary(F1=font))
    bbox = pikepdf.Array([0, 0, L.PAGE_W, L.PAGE_H])

    def form(data: bytes, res: pikepdf.Dictionary) -> pikepdf.Object:
        stream = pdf.make_stream(data)
        stream["/Type"] = pikepdf.Name.XObject
        stream["/Subtype"] = pikepdf.Name.Form
        stream["/BBox"] = bbox
        stream["/Resources"] = res
        return pdf.make_indirect(stream)

    watermark = form(b"0.9 g 100 400 300 20 re f", pikepdf.Dictionary())
    for number in range(1, L.PAGE_COUNT + 1):
        page = pdf.add_blank_page(page_size=(L.PAGE_W, L.PAGE_H))
        if number == B.P3_PAGE:
            data = _p3_stream()
        elif L.SUMMARY_FIRST_PAGE <= number < L.SUMMARY_FIRST_PAGE + L.SUMMARY_PAGE_COUNT:
            index = number - L.SUMMARY_FIRST_PAGE
            data = _text_stream(chunks[index * per_page : (index + 1) * per_page])
        else:
            data = _text_stream([f"page {number} stays as it is"])
        if 3 <= number <= 12:
            inner = form(data, resources)
            outer = form(b"q /fullpage Do Q", pikepdf.Dictionary(XObject=pikepdf.Dictionary(fullpage=inner)))
            page.obj["/Resources"] = pikepdf.Dictionary(XObject=pikepdf.Dictionary(fzFrm0=outer, fzFrm1=watermark))
            page.obj["/Contents"] = pikepdf.Array(
                [pdf.make_stream(b"q /fzFrm0 Do Q"), pdf.make_stream(b"q /fzFrm1 Do Q")]
            )
        else:
            page.obj["/Resources"] = resources
            page.obj["/Contents"] = pdf.make_stream(data)

    widget = pdf.make_indirect(
        pikepdf.Dictionary(
            Type=pikepdf.Name.Annot,
            Subtype=pikepdf.Name.Widget,
            FT=pikepdf.Name.Sig,
            T=pikepdf.String("Signature1"),
            Rect=pikepdf.Array([387.5, 169.13, 503.5, 285.13]),
            F=132,
            P=pdf.pages[12].obj,
        )
    )
    pdf.pages[12].obj["/Annots"] = pikepdf.Array([widget])
    pdf.Root["/AcroForm"] = pikepdf.Dictionary(Fields=pikepdf.Array([widget]), SigFlags=3)
    pdf.save(out_pdf)


def run(font_dir: Path) -> None:
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp)
        content = four_page_content(font_dir, work)
        base = work / "base.pdf"
        synthetic_base(content, base)
        summary_pdf, marks_pdf, out = work / "summary.pdf", work / "marks.pdf", work / "out.pdf"
        metrics = typeset_summary.typeset(content, summary_pdf, font_dir)
        overlays.build_marks_pdf(marks_pdf, font_dir)

        report = B.build_full(base, content, summary_pdf, metrics, marks_pdf, out, "fzFrm1", skip_text_gate=False)
        assert report["p3_labels"] == {"applied": True, "moved_operators": 8, "shift_pt": 50.62}, report
        assert report["signature_fields_kept"] == 1

        before, after = pymupdf.open(base), pymupdf.open(out)
        assert after.page_count == L.PAGE_COUNT

        # P3：8 个标签整体下移 50.62pt，其余文字不动。
        def tops(doc: pymupdf.Document, needle: str) -> list[float]:
            return sorted(round(r.y0, 1) for r in doc[B.P3_PAGE - 1].search_for(needle))

        moved = [b - a for a, b in zip(tops(before, "L"), tops(after, "L"), strict=True)]
        assert len(moved) == 8 and all(abs(d - 50.6) < 0.2 for d in moved), moved
        assert tops(before, "untouched") == tops(after, "untouched")

        # P8—P11：文字与底稿逐字一致，且每个字只出现一遍（旧文字层已不在页面上）。
        expected = textcheck.normalize("".join(typeset_summary.ordered_texts(content)))
        first = L.SUMMARY_FIRST_PAGE - 1
        actual = "".join(textcheck.body_text(after[first + i]) for i in range(L.SUMMARY_PAGE_COUNT))
        diff = textcheck.char_diff(expected, actual)
        assert not diff["missing"] and not diff["extra"], diff

        # 每页都有非官方标识；未改动的页面正文一字不差。
        for number in range(1, L.PAGE_COUNT + 1):
            page_text = textcheck.normalize(after[number - 1].get_text())
            assert textcheck.normalize(L.MARK_FOOTER) in page_text, number
            if number not in (B.P3_PAGE, *report["summary_pages_replaced"]):
                assert textcheck.body_text(after[number - 1]) == textcheck.body_text(before[number - 1]), number

        # 签名域原样保留，水印对象仍被 P8—P11 引用。
        with pikepdf.open(out) as pdf:
            assert len(pdf.Root.AcroForm.Fields) == 1
            assert pdf.pages[12].obj.Annots[0].FT == "/Sig"
            for i in range(L.SUMMARY_PAGE_COUNT):
                assert "/fzFrm1" in pdf.pages[first + i].obj.Resources.XObject.keys()
                assert "/fzFrm0" not in pdf.pages[first + i].obj.Resources.XObject.keys()

        # 底稿文字与重排底稿不一致时必须拒绝出件。
        tampered = json.loads(json.dumps(content))
        tampered["blocks"][0]["text"] += " changed"
        try:
            B.build_full(base, tampered, summary_pdf, metrics, marks_pdf, work / "x.pdf", "fzFrm1", skip_text_gate=False)
        except RuntimeError as error:
            assert "不一致" in str(error)
        else:
            raise AssertionError("文字不一致时没有拦住")
    print("通过：组装流程的 6 组检查全部成立")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--font-dir", required=True, type=Path)
    run(parser.parse_args().font_dir)
