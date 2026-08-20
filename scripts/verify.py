# -*- coding: utf-8 -*-
"""校验交付物：文件存在、可打开、十条课题齐全。"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from openpyxl import load_workbook
from pptx import Presentation
from docx import Document

import content as C

ROOT = Path(__file__).resolve().parents[1]
EXPORTS = ROOT / "exports"

FILES = [
    "AIDEN_西电产学研课题介绍.pptx",
    "AIDEN_西电产学研课题介绍.docx",
    "AIDEN_西电课题选题与编组表.xlsx",
    "微信群发布稿.txt",
    "微信群发布稿.md",
    "课题一页总览.html",
]


def fail(msg: str) -> None:
    print(f"FAIL  {msg}")
    sys.exit(1)


def ok(msg: str) -> None:
    print(f"OK    {msg}")


def main() -> None:
    ids = [t["id"] for t in C.TOPICS]
    if len(ids) != 10 or len(set(ids)) != 10:
        fail("content.TOPICS 必须恰好 10 个不重复编号")
    ok("十条课题编号完整")

    for name in FILES:
        path = EXPORTS / name
        if not path.exists() or path.stat().st_size < 200:
            fail(f"缺少或过小：{path}")
        ok(f"存在 {name}  ({path.stat().st_size} bytes)")

    pptx = EXPORTS / FILES[0]
    docx = EXPORTS / FILES[1]
    xlsx = EXPORTS / FILES[2]
    for zpath in (pptx, docx, xlsx):
        if not zipfile.is_zipfile(zpath):
            fail(f"不是有效 Office 压缩包：{zpath.name}")
        ok(f"zip 结构正常 {zpath.name}")

    prs = Presentation(pptx)
    if len(prs.slides) != 16:
        fail(f"PPT 应为 16 页，实际 {len(prs.slides)}")
    if prs.slide_width.inches < 13 or prs.slide_height.inches < 7:
        fail("PPT 应为 16:9 宽屏")
    ppt_text = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                ppt_text.append(shape.text_frame.text)
    blob = "\n".join(ppt_text)
    for tid in ids:
        if tid not in blob:
            fail(f"PPT 未出现 {tid}")
    if C.RELEASE_DATE.split("（")[0] not in blob and "8月24日" not in blob:
        fail("PPT 未写明下周一发布日")
    ok("PPT 16 页、16:9、含全部课题编号与发布日")

    doc = Document(docx)
    doc_text = "\n".join(p.text for p in doc.paragraphs)
    for tid in ids:
        if tid not in doc_text:
            fail(f"Word 未出现 {tid}")
    if "已经开工" not in doc_text:
        fail("Word 缺少「已经开工」口径")
    ok("Word 含十条课题与开工口径")

    wb = load_workbook(xlsx)
    need_sheets = ["使用说明", "课题总览", "技能匹配", "选题意向", "项目组编组", "时间表", "群发短稿"]
    for name in need_sheets:
        if name not in wb.sheetnames:
            fail(f"Excel 缺表：{name}")
    overview = wb["课题总览"]
    found = {row[0].value for row in overview.iter_rows(min_row=2, max_col=1) if row[0].value}
    if set(ids) - found:
        fail(f"课题总览缺编号：{set(ids) - found}")
    intent = wb["选题意向"]
    if intent.max_row < 20:
        fail("选题意向行数不足，学生没法填")
    ok("Excel 七张表、十条编号、选题意向可填")

    wechat = (EXPORTS / "微信群发布稿.txt").read_text(encoding="utf-8")
    html = (EXPORTS / "课题一页总览.html").read_text(encoding="utf-8")
    for blob_name, blob in (("微信稿", wechat), ("HTML", html)):
        for tid in ids:
            if tid not in blob:
                fail(f"{blob_name} 未出现 {tid}")
        if "第一志愿" not in blob:
            fail(f"{blob_name} 未说明志愿填法")
    if "<html" not in html.lower():
        fail("HTML 结构不完整")
    ok("微信稿与 HTML 含十条课题")
    print("全部校验通过")


if __name__ == "__main__":
    main()
