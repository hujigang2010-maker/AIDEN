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
    "人民医院",
    "保险公司",
    "美团",
    "康复期",
    "首选",
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


def pdf_text(path: Path) -> str:
    try:
        from pypdf import PdfReader
    except ImportError:
        return ""
    reader = PdfReader(str(path))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def main() -> None:
    docx = OUT / "青岛红枫路交通事故_伤情与处理备忘录_20260817.docx"
    xlsx = OUT / "青岛红枫路交通事故_伤情伤残与行动表_20260817.xlsx"
    pptx = OUT / "青岛红枫路交通事故_伤情与伤残简报_20260817.pptx"
    comm_docx = OUT / "青岛红枫路交通事故_肇事方沟通方案与体检分析_20260817.docx"
    comm_pdf = OUT / "青岛红枫路交通事故_肇事方沟通方案与体检分析_20260817.pdf"
    comm_xlsx = OUT / "青岛红枫路交通事故_肇事方沟通流程表_20260817.xlsx"
    for p in (docx, xlsx, pptx, comm_docx, comm_pdf, comm_xlsx):
        assert p.exists() and p.stat().st_size > 2000, p

    text = docx_text(docx)
    print("DOCX 段落+表格字符", len(text), "文件字节", docx.stat().st_size)
    for k in MUST:
        assert k in text, f"Word 缺少：{k}"

    wb = load_workbook(xlsx)
    sheets = wb.sheetnames
    print("XLSX 工作表", sheets)
    need = ["伤情鉴定", "伤残情况", "外伤与退变对照", "转院报销", "费用台账", "72小时待办", "沟通口径卡", "责任与程序"]
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

    ctext = docx_text(comm_docx)
    print("沟通方案 DOCX 字符", len(ctext), "字节", comm_docx.stat().st_size)
    for k in ("胫骨远端", "腓骨近端", "跟骨骨刺", "开场", "律师函", "方案甲", "12", "人民医院", "美团"):
        assert k in ctext, f"沟通方案 Word 缺少：{k}"
    assert "不要请诉讼律师" in ctext or "现在不要请诉讼律师" in ctext

    cwb = load_workbook(comm_xlsx)
    print("沟通流程 XLSX 工作表", cwb.sheetnames)
    for n in ("怎么用", "12步沟通顺序", "对方一句话怎么接", "律师决策", "三套方案", "影像三份对照"):
        assert n in cwb.sheetnames, n
    flow = cwb["12步沟通顺序"]
    flow_text = "\n".join(str(c.value or "") for row in flow.iter_rows(max_row=16) for c in row)
    assert "未开始" in flow_text and "肇事女骑手" in flow_text

    ptext_c = pdf_text(comm_pdf)
    print("沟通方案 PDF 抽取字符", len(ptext_c), "字节", comm_pdf.stat().st_size)
    assert comm_pdf.stat().st_size > 8000
    if ptext_c:
        for k in ("胫骨远端", "方案甲", "律师", "开场", "10008056847", "胡继刚"):
            assert k in ptext_c.replace(" ", ""), f"沟通方案 PDF 缺少：{k}"

    html = OUT / "hongfeng-guide.html"
    html_cn = OUT / "青岛红枫路交通事故_处理总览.html"
    for p in (html, html_cn):
        assert p.exists() and p.stat().st_size > 8000, p
        htext = p.read_text(encoding="utf-8")
        assert "<script" not in htext
        for k in ("青岛红枫路", "胫骨远端", "人民医院", "10008056847", "美团"):
            assert k in htext, f"网页缺少：{k}"
    print("HTML 字节", html.stat().st_size)

    print("校验通过")


if __name__ == "__main__":
    main()
