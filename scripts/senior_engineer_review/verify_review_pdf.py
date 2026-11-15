"""对审阅版 PDF 做技术核验，输出 JSON、Markdown 报告和全页渲染图。

检查项：页数与页面尺寸；字体是否全部嵌入、有无 Type 3；缺字（.notdef / U+FFFD）；
重复文字层与不可见文字；四个渲染引擎的逐页一致性；工作总结文字与底稿是否逐字一致；
外框与表格线；签名域的结构状态；给了底稿时，逐页比对未改动区域的像素。
"""

from __future__ import annotations

import argparse
import io
import json
import shutil
import subprocess
import tempfile
from collections import Counter
from pathlib import Path

import numpy as np
import pikepdf
import pymupdf
import pypdfium2
from PIL import Image, ImageFilter

import layout as L
import textcheck
import typeset_summary

DPI = 100
ENGINES = ("mupdf", "poppler", "pdfium", "ghostscript")


def _run(cmd: list[str]) -> str:
    return subprocess.run(cmd, check=False, capture_output=True, text=True).stdout


def fonts_report(pdf: Path) -> dict:
    rows = []
    for line in _run(["pdffonts", str(pdf)]).splitlines()[2:]:
        parts = line.split()
        if len(parts) < 7:
            continue
        emb, sub, uni = parts[-5], parts[-4], parts[-3]
        name = parts[0]
        kind = " ".join(parts[1:-6]) if len(parts) > 7 else parts[1]
        rows.append({"name": name, "type": kind, "embedded": emb == "yes", "subset": sub == "yes", "to_unicode": uni == "yes"})
    return {
        "fonts": rows,
        "not_embedded": sorted({r["name"] for r in rows if not r["embedded"]}),
        "type3": sorted({r["name"] for r in rows if "Type 3" in r["type"]}),
        "without_to_unicode": sorted({r["name"] for r in rows if not r["to_unicode"]}),
    }


def text_layer_report(doc: pymupdf.Document) -> list[dict]:
    pages = []
    for page in doc:
        notdef = replacement = invisible = total = 0
        seen: Counter = Counter()
        for trace in page.get_texttrace():
            hidden = trace.get("type") == 3 or trace.get("opacity", 1) == 0
            for code, glyph, origin, _bbox in trace["chars"]:
                if chr(code).isspace():
                    continue
                total += 1
                notdef += glyph == 0
                replacement += code == 0xFFFD
                invisible += hidden
                seen[(code, round(origin[0], 0), round(origin[1], 0))] += 1
        duplicates = sum(n - 1 for n in seen.values() if n > 1)
        pages.append(
            {
                "page": page.number + 1,
                "chars": total,
                "notdef_glyphs": notdef,
                "replacement_chars": replacement,
                "invisible_chars": invisible,
                "duplicate_chars": duplicates,
            }
        )
    return pages


def _gray(image: Image.Image) -> np.ndarray:
    return np.asarray(image.convert("L").filter(ImageFilter.GaussianBlur(1.2)), dtype=np.int16)


def _coarse(image: Image.Image) -> np.ndarray:
    """缩到四分之一：笔画粗细、抗锯齿的差别被平均掉，只剩文字和线条落在哪里。"""
    small = image.convert("L").resize((image.width // 4, image.height // 4), Image.BOX)
    return np.asarray(small, dtype=np.float64).ravel()


def render(pdf: Path, engine: str, out_dir: Path) -> list[Image.Image]:
    if engine == "mupdf":
        doc = pymupdf.open(pdf)
        return [Image.open(io.BytesIO(p.get_pixmap(dpi=DPI, alpha=False).tobytes("png"))).convert("RGB") for p in doc]
    if engine == "pdfium":
        doc = pypdfium2.PdfDocument(str(pdf))
        return [doc[i].render(scale=DPI / 72).to_pil().convert("RGB") for i in range(len(doc))]
    prefix = out_dir / engine
    if engine == "poppler":
        subprocess.run(["pdftoppm", "-r", str(DPI), "-png", str(pdf), str(prefix)], check=True)
    else:
        subprocess.run(
            ["gs", "-q", "-dNOPAUSE", "-dBATCH", "-sDEVICE=png16m", f"-r{DPI}", "-dTextAlphaBits=4",
             "-dGraphicsAlphaBits=4", f"-sOutputFile={prefix}-%02d.png", str(pdf)],
            check=True,
        )
    return [Image.open(p).convert("RGB") for p in sorted(out_dir.glob(f"{engine}-*.png"))]


def cross_engine_report(pdf: Path, render_dir: Path | None) -> dict:
    """各引擎渲染结果先轻度模糊再比较，滤掉抗锯齿差异，只留下真正的版面差异。"""
    with tempfile.TemporaryDirectory() as tmp:
        images = {engine: render(pdf, engine, Path(tmp)) for engine in ENGINES}
    reference = images["mupdf"]
    if render_dir:
        render_dir.mkdir(parents=True, exist_ok=True)
        for number, image in enumerate(reference, 1):
            image.save(render_dir / f"page-{number:02d}.png")
    pages = []
    for index, ref in enumerate(reference):
        base, base_coarse = _gray(ref), _coarse(ref)
        row = {"page": index + 1, "layout_correlation": {}, "differing_pixel_ratio": {}}
        for engine in ENGINES[1:]:
            image = images[engine][index]
            if image.size != ref.size:
                image = image.resize(ref.size)
            row["layout_correlation"][engine] = round(float(np.corrcoef(base_coarse, _coarse(image))[0, 1]), 4)
            row["differing_pixel_ratio"][engine] = round(float((np.abs(_gray(image) - base) > 96).mean()), 5)
        pages.append(row)
    return {
        "dpi": DPI,
        "reference": "mupdf",
        "pages": pages,
        "worst_differing_pixel_ratio": max(max(p["differing_pixel_ratio"].values()) for p in pages),
        "lowest_layout_correlation": min(min(p["layout_correlation"].values()) for p in pages),
    }


def contact_sheet(render_dir: Path, out_png: Path, columns: int) -> None:
    files = sorted(render_dir.glob("page-*.png"))
    thumbs = [Image.open(f).convert("RGB") for f in files]
    width, height = thumbs[0].size
    scale = 0.5
    tw, th = int(width * scale), int(height * scale)
    rows = -(-len(thumbs) // columns)
    sheet = Image.new("RGB", (columns * (tw + 8) + 8, rows * (th + 8) + 8), (200, 200, 200))
    for i, thumb in enumerate(thumbs):
        sheet.paste(thumb.resize((tw, th), Image.LANCZOS), (8 + (i % columns) * (tw + 8), 8 + (i // columns) * (th + 8)))
    sheet.save(out_png)


def summary_report(doc: pymupdf.Document, content: dict, first_page: int) -> dict:
    expected = textcheck.normalize("".join(typeset_summary.ordered_texts(content)))
    pages = [doc[first_page - 1 + i] for i in range(L.SUMMARY_PAGE_COUNT)]
    actual = "".join(textcheck.body_text(p) for p in pages)
    report = {"text_vs_content": textcheck.char_diff(expected, actual), "pages": []}
    for offset, page in enumerate(pages):
        sizes: Counter = Counter()
        families: Counter = Counter()
        for block in page.get_text("dict")["blocks"]:
            if block["type"] != 0:
                continue
            for line in block["lines"]:
                if abs(line["dir"][1]) > 0.01 or line["bbox"][1] > L.FRAME_BOTTOM:
                    continue
                for span in line["spans"]:
                    count = len(span["text"].strip())
                    sizes[round(span["size"], 2)] += count
                    families[span["font"]] += count
        strokes = sorted({round(d["width"], 2) for d in page.get_drawings() if d.get("width") and d["type"] in ("s", "fs")})
        rects = [d["rect"] for d in page.get_drawings() if d["type"] == "s" and d["rect"].width > 400]
        frame = max(rects, key=lambda r: r.height) if rects else None
        lines = textcheck.body_lines(page)
        inside = [line for line in lines if frame is None or line["bbox"][3] <= frame.y1 + 0.5]
        report["pages"].append(
            {
                "page": first_page + offset,
                "font_sizes_pt": dict(sorted(sizes.items())),
                "fonts": dict(families),
                "stroke_widths_pt": strokes,
                "frame": [round(v, 2) for v in frame] if frame else None,
                "text_left": round(min(line["bbox"][0] for line in lines), 2),
                "text_right": round(max(line["bbox"][2] for line in lines), 2),
                "text_bottom_inside_frame": round(max(line["bbox"][3] for line in inside), 2),
            }
        )
    return report


def signature_report(pdf: Path) -> dict:
    size = pdf.stat().st_size
    fields = []
    with pikepdf.open(pdf) as doc:
        for field in doc.Root.get("/AcroForm", {}).get("/Fields", []):
            value = field.get("/V")
            entry = {"name": str(field.get("/T", "")), "type": str(field.get("/FT", "")), "has_value": value is not None}
            if value is not None and "/ByteRange" in value:
                byte_range = [int(v) for v in value.ByteRange]
                entry["byte_range"] = byte_range
                entry["byte_range_fits_file"] = byte_range[2] + byte_range[3] <= size
                entry["covers_whole_file"] = byte_range[2] + byte_range[3] == size
            fields.append(entry)
    pdfsig = _run(["pdfsig", "-nocert", str(pdf)]) if shutil.which("pdfsig") else ""
    return {
        "fields": fields,
        "pdfsig_lines": [line.strip() for line in pdfsig.splitlines() if "Signature Validation" in line or "Total document signed" in line],
        "conclusion": "本文件经过重排和再保存，其中的签名域（如有）不构成有效电子签章。",
    }


def base_compare_report(pdf: Path, base: Path, marks: Path, changed_pages: set[int]) -> list[dict]:
    """未改动的页面：除“非官方标识”覆盖到的像素外，应与底稿逐像素一致。"""
    out_doc, base_doc, marks_doc = pymupdf.open(pdf), pymupdf.open(base), pymupdf.open(marks)

    def pixels(page: pymupdf.Page) -> np.ndarray:
        pix = page.get_pixmap(dpi=DPI, alpha=False, annots=False)
        return np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3).astype(np.int16)

    mark = (pixels(marks_doc[0]).min(axis=2) < 250)
    mark = np.asarray(Image.fromarray(mark).filter(ImageFilter.MaxFilter(5)))
    rows = []
    for index in range(min(out_doc.page_count, base_doc.page_count)):
        diff = np.abs(pixels(out_doc[index]) - pixels(base_doc[index])).max(axis=2)
        outside = diff[~mark]
        rows.append(
            {
                "page": index + 1,
                "expected_to_change": index + 1 in changed_pages,
                "max_diff_outside_marks": int(outside.max()),
                "changed_pixels_outside_marks": int((outside > 8).sum()),
            }
        )
    return rows


def to_markdown(report: dict) -> str:
    out = ["# 技术核验报告", "", f"- 文件：`{report['file']}`", f"- 页数：{report['page_count']}；页面尺寸：{report['page_sizes']}", ""]
    fonts = report["fonts"]
    out += ["## 字体", "", "| 字体 | 类型 | 嵌入 | 子集 | ToUnicode |", "|---|---|---|---|---|"]
    out += [
        f"| {f['name']} | {f['type']} | {'是' if f['embedded'] else '否'} | {'是' if f['subset'] else '否'} | {'是' if f['to_unicode'] else '否'} |"
        for f in fonts["fonts"]
    ]
    out += ["", f"- 未嵌入：{fonts['not_embedded'] or '无'}；Type 3：{fonts['type3'] or '无'}；缺 ToUnicode：{fonts['without_to_unicode'] or '无'}", ""]
    out += ["## 文字层", "", "| 页 | 字符数 | 缺字(.notdef) | U+FFFD | 不可见文字 | 同位置重复字符 |", "|---|---|---|---|---|---|"]
    out += [
        f"| {p['page']} | {p['chars']} | {p['notdef_glyphs']} | {p['replacement_chars']} | {p['invisible_chars']} | {p['duplicate_chars']} |"
        for p in report["text_layer"]
    ]
    cross = report["cross_engine"]
    out += ["", f"## 跨引擎渲染（{cross['dpi']} dpi，以 MuPDF 为基准）", "", "| 页 | 版面相关系数 Poppler | PDFium | Ghostscript | 明显差异像素比 Poppler | PDFium | Ghostscript |", "|---|---|---|---|---|---|---|"]
    for p in cross["pages"]:
        corr, diff = p["layout_correlation"], p["differing_pixel_ratio"]
        out.append(
            f"| {p['page']} | {corr['poppler']:.4f} | {corr['pdfium']:.4f} | {corr['ghostscript']:.4f} | "
            f"{diff['poppler']:.5f} | {diff['pdfium']:.5f} | {diff['ghostscript']:.5f} |"
        )
    out += [
        "",
        f"- 最低版面相关系数：{cross['lowest_layout_correlation']:.4f}（1 为完全一致；缺字、错位或换行不同会明显拉低）",
        f"- 最大明显差异像素比：{cross['worst_differing_pixel_ratio']:.5f}（轻度模糊后灰度差超过 96 的像素占比，用来排除抗锯齿差别）",
        "",
    ]
    if "summary" in report:
        s = report["summary"]
        d = s["text_vs_content"]
        out += [
            "## 工作总结页（P8—P11）",
            "",
            f"- 与文字底稿比对：底稿 {d['expected_chars']} 字，页面 {d['actual_chars']} 字；缺 {d['missing'] or '无'}，多 {d['extra'] or '无'}；顺序一致：{'是' if d['same_sequence'] else '否（表格读序不同属正常）'}",
            "",
            "| 页 | 字号分布(pt:字数) | 线宽(pt) | 外框(x0,y0,x1,y1) | 文字左/右边 | 框内文字底边 |",
            "|---|---|---|---|---|---|",
        ]
        out += [
            f"| {p['page']} | {p['font_sizes_pt']} | {p['stroke_widths_pt']} | {p['frame']} | {p['text_left']} / {p['text_right']} | {p['text_bottom_inside_frame']} |"
            for p in s["pages"]
        ]
        out.append("")
    sig = report["signatures"]
    out += ["## 签名域", "", f"- 签名域数量：{len(sig['fields'])}"]
    out += [f"  - {f['name']}：{json.dumps({k: v for k, v in f.items() if k != 'name'}, ensure_ascii=False)}" for f in sig["fields"]]
    out += [f"- pdfsig：{sig['pdfsig_lines'] or '无签名可验'}", f"- 结论：{sig['conclusion']}", ""]
    if "base_compare" in report:
        out += ["## 与底稿逐页像素比对（不含标识覆盖区域）", "", "| 页 | 预期有改动 | 最大色差 | 变化像素数 |", "|---|---|---|---|"]
        out += [
            f"| {p['page']} | {'是' if p['expected_to_change'] else '否'} | {p['max_diff_outside_marks']} | {p['changed_pixels_outside_marks']} |"
            for p in report["base_compare"]
        ]
        out.append("")
    return "\n".join(out)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--content", type=Path, help="P8—P11 文字底稿 JSON")
    parser.add_argument("--summary-first-page", type=int, default=L.SUMMARY_FIRST_PAGE, help="工作总结在本文件中的起始页")
    parser.add_argument("--base", type=Path, help="14 页底稿 PDF，用于比对未改动页面")
    parser.add_argument("--marks", type=Path, help="标识页 PDF（与 --base 一起用）")
    parser.add_argument("--render-dir", type=Path, help="保存全页渲染图的目录")
    parser.add_argument("--contact-sheet", type=Path, help="全页缩略总览图")
    parser.add_argument("--columns", type=int, default=7)
    parser.add_argument("--out-json", type=Path)
    parser.add_argument("--out-md", type=Path)
    args = parser.parse_args()

    doc = pymupdf.open(args.pdf)
    report: dict = {
        "file": args.pdf.name,
        "page_count": doc.page_count,
        "page_sizes": sorted({f"{p.rect.width:.2f}x{p.rect.height:.2f}" for p in doc}),
        "fonts": fonts_report(args.pdf),
        "text_layer": text_layer_report(doc),
        "cross_engine": cross_engine_report(args.pdf, args.render_dir),
        "signatures": signature_report(args.pdf),
    }
    if args.content:
        content = json.loads(args.content.read_text(encoding="utf-8"))
        report["summary"] = summary_report(doc, content, args.summary_first_page)
    if args.base and args.marks:
        changed = {3, *range(L.SUMMARY_FIRST_PAGE, L.SUMMARY_FIRST_PAGE + L.SUMMARY_PAGE_COUNT)}
        report["base_compare"] = base_compare_report(args.pdf, args.base, args.marks, changed)
    if args.contact_sheet and args.render_dir:
        contact_sheet(args.render_dir, args.contact_sheet, args.columns)
    if args.out_json:
        args.out_json.write_text(json.dumps(report, ensure_ascii=False, indent=1), "utf-8")
    markdown = to_markdown(report)
    if args.out_md:
        args.out_md.write_text(markdown, "utf-8")
    print(markdown)


if __name__ == "__main__":
    main()
