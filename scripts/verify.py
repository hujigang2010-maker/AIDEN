# -*- coding: utf-8 -*-
"""校验给对方的汇报 PPT 和协议。"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from pptx import Presentation

ROOT = Path(__file__).resolve().parent.parent / "output"
PPT = ROOT / "港大经管上海中心_联合合作汇报.pptx"
AGR = ROOT / "港大经管上海中心_联合策划服务合作协议.docx"

# 汇报版不应出现的内部腔、分析腔、套话
BANNED = [
    "口径",
    "听懂",
    "不是我们编",
    "层 A",
    "层A",
    "粘性",
    "抓手",
    "赋能",
    "闭环",
    "SOP",
    "KPI",
    "灰色",
    "收口",
    "对齐",
    "颗粒度",
    "赛道",
    "对方需求",
    "建议稿",
    "V1.1",
]


def fail(msg: str) -> None:
    print(f"FAIL  {msg}")
    sys.exit(1)


def ok(msg: str) -> None:
    print(f"OK    {msg}")


def ppt_text(path: Path) -> str:
    prs = Presentation(path)
    chunks = []
    for slide in prs.slides:
        for shp in slide.shapes:
            if shp.has_text_frame:
                chunks.append(shp.text_frame.text)
            if shp.has_table:
                for row in shp.table.rows:
                    for cell in row.cells:
                        chunks.append(cell.text)
    return "\n".join(chunks)


def doc_text(path: Path) -> str:
    doc = Document(path)
    parts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return "\n".join(parts)


def main() -> None:
    if not PPT.exists() or PPT.stat().st_size < 1000:
        fail(f"PPT 缺失：{PPT}")
    if not AGR.exists() or AGR.stat().st_size < 1000:
        fail(f"协议 缺失：{AGR}")
    ok("两份汇报文件存在")

    prs = Presentation(PPT)
    if len(prs.slides) != 11:
        fail(f"PPT 应为 11 页，实际 {len(prs.slides)}")
    ok("PPT 11 页")

    t = ppt_text(PPT) + "\n" + doc_text(AGR)
    for w in BANNED:
        if w in t:
            fail(f"汇报版出现不宜对外的用语：{w}")
    ok("未使用内部分析用语")

    for needle in ("88,000", "捌万捌仟", "潘嘉琰", "出海", "主任", "一次付清", "EMBA", "不另收费"):
        if needle not in t:
            fail(f"缺少必要内容：{needle}")
    if "学费" not in t:
        fail("协议或 PPT 未写明不含学费分成")
    ppt = ppt_text(PPT)
    agr = doc_text(AGR)
    if "EMBA" not in ppt or "EMBA" not in agr:
        fail("PPT 与协议均须写明 EMBA 招生")
    if "一整包" not in ppt and "一次合作" not in agr:
        fail("须写明招生、出海课、请主任到场合在一次合作")
    ok("费用、EMBA招生、出海专题课、请主任到场已合并写入")
    print("ALL PASSED")


if __name__ == "__main__":
    main()
