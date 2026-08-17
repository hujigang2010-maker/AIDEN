#!/usr/bin/env python3
"""校验三份交付物结构与关键结论是否写入。"""

from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

from openpyxl import load_workbook
from pptx import Presentation
from docx import Document

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "deliverables"

MUST = [
    "胡**",
    "64",
    "胫骨远端",
    "腓骨近端",
    "半脱位",
    "伤残尚未",
    "内外固定",
    "跟骨骨刺",
]


def docx_text(path: Path) -> str:
    doc = Document(path)
    parts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return "\n".join(parts)


def ppt_text(path: Path) -> str:
    prs = Presentation(path)
    chunks = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                chunks.append(shape.text_frame.text)
    return "\n".join(chunks)


def main() -> None:
    docx = OUT / "青岛红枫路交通事故_伤情与处理备忘录_20260817.docx"
    xlsx = OUT / "青岛红枫路交通事故_伤情伤残与行动表_20260817.xlsx"
    pptx = OUT / "青岛红枫路交通事故_伤情与伤残简报_20260817.pptx"
    for p in (docx, xlsx, pptx):
        assert p.exists() and p.stat().st_size > 2000, p

    text = docx_text(docx)
    print("DOCX 段落+表格字符", len(text), "文件字节", docx.stat().st_size)
    for k in MUST:
        assert k in text, f"Word 缺少：{k}"

    wb = load_workbook(xlsx)
    sheets = wb.sheetnames
    print("XLSX 工作表", sheets)
    need = ["伤情鉴定", "伤残情况", "外伤与退变对照", "费用台账", "72小时待办", "沟通口径卡", "责任与程序"]
    for n in need:
        assert n in sheets, n
    ws = wb["伤情鉴定"]
    joined = "\n".join(str(c.value or "") for row in ws.iter_rows(max_row=8) for c in row)
    assert "10008056847" in joined and "10008059257" in joined
    dis = wb["伤残情况"]
    dis_text = "\n".join(str(c.value or "") for row in dis.iter_rows(max_row=10) for c in row)
    assert "尚未" in dis_text or "不能" in dis_text

    prs = Presentation(pptx)
    print("PPT 页数", len(prs.slides))
    assert len(prs.slides) == 10
    ptext = ppt_text(pptx)
    for k in ("伤残尚未鉴定", "胫骨远端", "禁止说"):
        assert k in ptext.replace("\n", "") or k in ptext, k

    print("校验通过")


if __name__ == "__main__":
    main()
