"""组装“非官方排版审阅版”PDF。

两种用法：

1. 完整 14 页：以用户提供的 14 页 PDF 为底，P8—P11 换成重排后的工作总结，
   修正 P3 两个竖排栏目标签的位置，其余页面内容流不动；签名域原样保留，
   不压平、不新建、不复制任何印章或签名；每页叠加“非官方”标识。
2. --sample：只输出 P8—P11 四页样张（不带平台水印），用于没有底稿时先看版式。

底稿里 P8—P11 的文字必须和重排底稿逐字一致，否则直接报错退出——
本工具只改版式，不改内容。
"""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path

import pikepdf
import pymupdf

import layout as L
import overlays
import textcheck
import typeset_summary

# P3“学历情况”“奖惩情况”两个竖排标签在底稿中的位置（Skia 坐标，y 向下，1 单位 = 0.75pt）。
# 两组标签各偏上约 50.6pt，没有落在所属单元格的中线上。
P3_LABEL_X = 90.5
P3_LABEL_YS = (
    413.123275,
    431.723275,
    450.323275,
    468.923275,
    577.623275,
    596.223275,
    614.823275,
    633.423275,
)
P3_LABEL_SHIFT = 67.5
P3_PAGE = 3

SUMMARY_XOBJECT = pikepdf.Name("/RvSummary")
TITLE_SUFFIX = "（非官方排版审阅版）"
SUBJECT = "内部排版审阅用，不具官方效力；页面上的印章外观不等于有效电子签章，正式用途以江苏省官方归档件为准。"


def _close(a: float, b: float, tol: float = 1e-3) -> bool:
    return math.isclose(float(a), b, abs_tol=tol)


def _nested_forms(xobject: pikepdf.Object, seen: set) -> list[pikepdf.Object]:
    """按引用关系走遍一个表单对象及其内嵌的表单对象。"""
    key = xobject.objgen
    if key in seen:
        return []
    seen.add(key)
    found = [xobject]
    children = xobject.get("/Resources", {}).get("/XObject", {})
    for name in children.keys():
        child = children[name]
        if child.get("/Subtype") == "/Form":
            found.extend(_nested_forms(child, seen))
    return found


def fix_p3_labels(pdf: pikepdf.Pdf) -> dict:
    """把 P3 两个竖排标签移到各自单元格的中线上。只改 8 个 Tm 的纵坐标。"""
    page = pdf.pages[P3_PAGE - 1]
    xobjects = page.obj.get("/Resources", {}).get("/XObject", {})
    seen: set = set()
    forms = [f for name in xobjects.keys() for f in _nested_forms(xobjects[name], seen)]
    for form in forms:
        instructions = pikepdf.parse_content_stream(form)
        hits = [
            i
            for i, (operands, operator) in enumerate(instructions)
            if str(operator) == "Tm"
            and _close(operands[4], P3_LABEL_X)
            and any(_close(operands[5], y) for y in P3_LABEL_YS)
        ]
        if not hits:
            continue
        if len(hits) != len(P3_LABEL_YS):
            raise RuntimeError(f"P3 标签只匹配到 {len(hits)} 处，应为 {len(P3_LABEL_YS)} 处，底稿结构与预期不符")
        other_users = [
            n + 1
            for n, other in enumerate(pdf.pages)
            if n != P3_PAGE - 1
            and form.objgen
            in {
                f.objgen
                for name in other.obj.get("/Resources", {}).get("/XObject", {}).keys()
                for f in _nested_forms(other.obj.Resources.XObject[name], set())
            }
        ]
        if other_users:
            raise RuntimeError(f"P3 的内容流同时被第 {other_users} 页使用，不能直接修改")
        for i in hits:
            operands, operator = instructions[i]
            moved = list(operands)
            moved[5] = float(operands[5]) + P3_LABEL_SHIFT
            instructions[i] = pikepdf.ContentStreamInstruction(moved, operator)
        form.write(pikepdf.unparse_content_stream(instructions))
        return {"applied": True, "moved_operators": len(hits), "shift_pt": round(P3_LABEL_SHIFT * 0.75, 2)}
    return {"applied": False, "reason": "P3 中没有找到预期位置的竖排标签，未作任何改动"}


def _summary_page_content(index: int, last_frame_bottom: float, watermark: pikepdf.Name | None) -> bytes:
    parts = [b"q ", bytes(SUMMARY_XOBJECT), b" Do Q\nq\n", overlays.frame_stream(index, last_frame_bottom), b"Q\n"]
    if watermark is not None:
        parts += [b"q ", bytes(watermark), b" Do Q\n"]
    return b"".join(parts)


def _base_text_gate(base_path: Path, content: dict) -> dict:
    """底稿 P8—P11 的文字必须与重排底稿逐字一致（按字符多重集）。"""
    doc = pymupdf.open(base_path)
    first = L.SUMMARY_FIRST_PAGE - 1
    actual = "".join(textcheck.body_text(doc[first + i]) for i in range(L.SUMMARY_PAGE_COUNT))
    expected = textcheck.normalize("".join(typeset_summary.ordered_texts(content)))
    return textcheck.char_diff(expected, actual)


def _set_metadata(pdf: pikepdf.Pdf, title: str) -> None:
    with pdf.open_metadata(set_pikepdf_as_editor=False) as meta:
        meta["dc:title"] = title
        meta["dc:description"] = SUBJECT
        meta["pdf:Producer"] = "senior_engineer_review（非官方排版工具）"
    pdf.docinfo["/Title"] = title
    pdf.docinfo["/Subject"] = SUBJECT
    pdf.docinfo["/Producer"] = "senior_engineer_review（非官方排版工具）"


def _overlay_marks(pdf: pikepdf.Pdf, marks_pdf: Path) -> None:
    with pikepdf.open(marks_pdf) as marks:
        source = marks.pages[0].as_form_xobject()
        # 立刻取出内容再关闭标识文件，避免保存时回头去读一个已关闭的文件。
        form = pdf.make_stream(source.read_bytes())
        for key in ("/Type", "/Subtype", "/BBox", "/Resources", "/Group", "/Matrix"):
            if key in source:
                form[key] = pdf.copy_foreign(marks.make_indirect(source[key]))
    for page in pdf.pages:
        page.add_overlay(form, pikepdf.Rectangle(0, 0, L.PAGE_W, L.PAGE_H))


def build_sample(summary_pdf: Path, metrics: dict, marks_pdf: Path, out_pdf: Path) -> dict:
    # 表单对象的内容是保存时才读取的，所以不能在排版件自身上改页面，要另起一个文件。
    summary = pikepdf.open(summary_pdf)
    pdf = pikepdf.new()
    for index, source in enumerate(summary.pages):
        form = pdf.make_indirect(pdf.copy_foreign(source.as_form_xobject()))
        page = pdf.add_blank_page(page_size=(L.PAGE_W, L.PAGE_H))
        page.obj["/Resources"] = pikepdf.Dictionary(XObject=pikepdf.Dictionary({str(SUMMARY_XOBJECT): form}))
        page.obj["/Contents"] = pdf.make_stream(
            _summary_page_content(index, metrics["last_frame_bottom"], watermark=None)
        )
    _overlay_marks(pdf, marks_pdf)
    _set_metadata(pdf, f"个人专业技术工作总结 P8—P11 样张{TITLE_SUFFIX}")
    pdf.save(out_pdf, object_stream_mode=pikepdf.ObjectStreamMode.disable, compress_streams=True)
    return {"mode": "sample", "pages": len(pdf.pages)}


def build_full(
    base_pdf: Path,
    content: dict,
    summary_pdf: Path,
    metrics: dict,
    marks_pdf: Path,
    out_pdf: Path,
    watermark_name: str,
    skip_text_gate: bool,
) -> dict:
    report: dict = {"mode": "full"}
    gate = _base_text_gate(base_pdf, content)
    report["base_text_gate"] = gate
    if (gate["missing"] or gate["extra"]) and not skip_text_gate:
        raise RuntimeError(
            "底稿 P8—P11 的文字与重排底稿不一致，已停止："
            f"缺 {gate['missing']}，多 {gate['extra']}。请先核对底稿，不要带着差异出件。"
        )

    pdf = pikepdf.open(base_pdf)
    if len(pdf.pages) != L.PAGE_COUNT:
        raise RuntimeError(f"底稿有 {len(pdf.pages)} 页，应为 {L.PAGE_COUNT} 页")

    fields_before = [f.objgen for f in pdf.Root.get("/AcroForm", {}).get("/Fields", [])]
    report["p3_labels"] = fix_p3_labels(pdf)

    watermark = pikepdf.Name("/" + watermark_name)
    summary = pikepdf.open(summary_pdf)
    for index in range(L.SUMMARY_PAGE_COUNT):
        page = pdf.pages[L.SUMMARY_FIRST_PAGE - 1 + index]
        old_xobjects = page.obj.get("/Resources", {}).get("/XObject", {})
        if str(watermark) not in old_xobjects.keys():
            raise RuntimeError(
                f"第 {L.SUMMARY_FIRST_PAGE + index} 页没有名为 {watermark_name} 的水印对象，"
                "无法原样保留水印；请用 --watermark-xobject 指定正确的名字"
            )
        form = pdf.copy_foreign(summary.pages[index].as_form_xobject())
        page.obj["/Resources"] = pikepdf.Dictionary(
            XObject=pikepdf.Dictionary(
                {str(SUMMARY_XOBJECT): pdf.make_indirect(form), str(watermark): old_xobjects[str(watermark)]}
            )
        )
        page.obj["/Contents"] = pdf.make_stream(
            _summary_page_content(index, metrics["last_frame_bottom"], watermark=watermark)
        )
    report["summary_pages_replaced"] = list(
        range(L.SUMMARY_FIRST_PAGE, L.SUMMARY_FIRST_PAGE + L.SUMMARY_PAGE_COUNT)
    )

    _overlay_marks(pdf, marks_pdf)
    title = str(pdf.docinfo.get("/Title", "")).strip() or "专业技术资格评审申报表"
    if TITLE_SUFFIX not in title:
        title += TITLE_SUFFIX
    _set_metadata(pdf, title)

    fields_after = [f.objgen for f in pdf.Root.get("/AcroForm", {}).get("/Fields", [])]
    if fields_before != fields_after:
        raise RuntimeError("签名域列表发生了变化，已停止")
    report["signature_fields_kept"] = len(fields_after)

    pdf.save(out_pdf, object_stream_mode=pikepdf.ObjectStreamMode.disable, compress_streams=True)
    report["pages"] = len(pdf.pages)
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--content", required=True, type=Path, help="P8—P11 文字底稿 JSON")
    parser.add_argument("--font-dir", required=True, type=Path)
    parser.add_argument("--work-dir", required=True, type=Path, help="中间文件目录")
    parser.add_argument("--out", required=True, type=Path)
    parser.add_argument("--base", type=Path, help="14 页底稿 PDF；不给则必须用 --sample")
    parser.add_argument("--sample", action="store_true", help="只输出 P8—P11 四页样张")
    parser.add_argument("--watermark-xobject", default="fzFrm1", help="底稿页面资源里水印表单对象的名字")
    parser.add_argument(
        "--skip-text-gate",
        action="store_true",
        help="仅供用无关底稿测试组装流程；正式出件不得使用",
    )
    args = parser.parse_args()
    if bool(args.base) == args.sample:
        parser.error("--base 与 --sample 必须二选一")

    args.work_dir.mkdir(parents=True, exist_ok=True)
    content = json.loads(args.content.read_text(encoding="utf-8"))
    summary_pdf = args.work_dir / "summary_pages.pdf"
    marks_pdf = args.work_dir / "marks.pdf"
    metrics = typeset_summary.typeset(content, summary_pdf, args.font_dir)
    overlays.build_marks_pdf(marks_pdf, args.font_dir)

    if args.sample:
        report = build_sample(summary_pdf, metrics, marks_pdf, args.out)
    else:
        report = build_full(
            args.base,
            content,
            summary_pdf,
            metrics,
            marks_pdf,
            args.out,
            args.watermark_xobject,
            args.skip_text_gate,
        )
    report["summary_metrics"] = metrics
    (args.work_dir / "build_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=1), "utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
