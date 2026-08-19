#!/usr/bin/env python3
"""一次生成 Word / Excel / PPT / PDF。"""

from pathlib import Path

from generate_comm_docx import build_comm_document
from generate_comm_pdf import build_comm_pdf
from generate_comm_xlsx import build_comm_workbook
from generate_docx import build_document
from generate_guide_html import main as build_guide_html
from generate_lawyer_brief import build_document as build_lawyer_brief
from generate_lawyer_brief import build_markdown as build_lawyer_md
from generate_lawyer_brief import DOCX_NAME as LAWYER_DOCX
from generate_lawyer_brief import MD_NAME as LAWYER_MD
from generate_pptx import build_ppt
from generate_xlsx import build_workbook

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "deliverables"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    docx = OUT / "青岛红枫路交通事故_伤情与处理备忘录_20260817.docx"
    xlsx = OUT / "青岛红枫路交通事故_伤情伤残与行动表_20260817.xlsx"
    pptx = OUT / "青岛红枫路交通事故_伤情与伤残简报_20260817.pptx"
    comm_docx = OUT / "青岛红枫路交通事故_肇事方沟通方案与体检分析_20260817.docx"
    comm_pdf = OUT / "青岛红枫路交通事故_肇事方沟通方案与体检分析_20260817.pdf"
    comm_xlsx = OUT / "青岛红枫路交通事故_肇事方沟通流程表_20260817.xlsx"
    build_document(docx)
    build_workbook(xlsx)
    build_ppt(pptx)
    build_comm_document(comm_docx)
    build_comm_pdf(comm_pdf)
    build_comm_workbook(comm_xlsx)
    lawyer_md = OUT / LAWYER_MD
    lawyer_docx = OUT / LAWYER_DOCX
    build_lawyer_md(lawyer_md)
    build_lawyer_brief(lawyer_docx)
    build_guide_html()
    print(f"已生成：\n  {docx}\n  {xlsx}\n  {pptx}\n  {comm_docx}\n  {comm_pdf}\n  {comm_xlsx}\n  {lawyer_md}\n  {lawyer_docx}")


if __name__ == "__main__":
    main()
