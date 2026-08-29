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
from generate_lawyer_brief import build_pdf as build_lawyer_pdf
from generate_lawyer_brief import DOCX_NAME as LAWYER_DOCX
from generate_lawyer_brief import MD_NAME as LAWYER_MD
from generate_lawyer_brief import PDF_NAME as LAWYER_PDF
from generate_lawyer_illustrated import DOCX_NAME as LAWYER_ILLU_DOCX
from generate_lawyer_illustrated import PDF_NAME as LAWYER_ILLU_PDF
from generate_lawyer_illustrated import build_document as build_lawyer_illustrated
from generate_lawyer_illustrated import build_pdf as build_lawyer_illustrated_pdf
from generate_lawyer_illustrated import extract_photos
from generate_liu_note import DOCX_NAME as LIU_DOCX
from generate_liu_note import MD_NAME as LIU_MD
from generate_liu_note import PDF_NAME as LIU_PDF
from generate_liu_note import build_document as build_liu_note
from generate_liu_note import build_markdown as build_liu_md
from generate_liu_note import build_pdf as build_liu_pdf
from generate_mom_liu_talk import DOCX_NAME as MOM_DOCX
from generate_mom_liu_talk import MD_NAME as MOM_MD
from generate_mom_liu_talk import PDF_NAME as MOM_PDF
from generate_mom_liu_talk import build_document as build_mom_card
from generate_mom_liu_talk import build_markdown as build_mom_md
from generate_mom_liu_talk import build_pdf as build_mom_pdf
from generate_pptx import build_ppt
from generate_solution import DOCX_NAME as SOL_DOCX
from generate_solution import XLSX_NAME as SOL_XLSX
from generate_solution import build_document as build_solution_doc
from generate_solution import build_workbook as build_solution_wb
from generate_xlsx import build_workbook
from pack_all import build_zip

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "deliverables"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    docx = OUT / "青岛抚顺路和哈尔滨路路口交通事故_伤情与处理备忘录_20260817.docx"
    xlsx = OUT / "青岛抚顺路和哈尔滨路路口交通事故_伤情伤残与行动表_20260817.xlsx"
    pptx = OUT / "青岛抚顺路和哈尔滨路路口交通事故_伤情与伤残简报_20260817.pptx"
    comm_docx = OUT / "青岛抚顺路和哈尔滨路路口交通事故_肇事方沟通方案与体检分析_20260817.docx"
    comm_pdf = OUT / "青岛抚顺路和哈尔滨路路口交通事故_肇事方沟通方案与体检分析_20260817.pdf"
    comm_xlsx = OUT / "青岛抚顺路和哈尔滨路路口交通事故_肇事方沟通流程表_20260817.xlsx"
    build_document(docx)
    build_workbook(xlsx)
    build_ppt(pptx)
    build_comm_document(comm_docx)
    build_comm_pdf(comm_pdf)
    build_comm_workbook(comm_xlsx)
    lawyer_md = OUT / LAWYER_MD
    lawyer_docx = OUT / LAWYER_DOCX
    lawyer_pdf = OUT / LAWYER_PDF
    lawyer_illu_docx = OUT / LAWYER_ILLU_DOCX
    lawyer_illu_pdf = OUT / LAWYER_ILLU_PDF
    build_lawyer_md(lawyer_md)
    build_lawyer_brief(lawyer_docx)
    build_lawyer_pdf(lawyer_pdf)
    extract_photos()
    build_lawyer_illustrated(lawyer_illu_docx)
    build_lawyer_illustrated_pdf(lawyer_illu_pdf)
    liu_md = OUT / LIU_MD
    liu_docx = OUT / LIU_DOCX
    liu_pdf = OUT / LIU_PDF
    build_liu_md(liu_md)
    build_liu_note(liu_docx)
    build_liu_pdf(liu_pdf)
    mom_md = OUT / MOM_MD
    mom_docx = OUT / MOM_DOCX
    mom_pdf = OUT / MOM_PDF
    build_mom_md(mom_md)
    build_mom_card(mom_docx)
    build_mom_pdf(mom_pdf)
    sol_docx = OUT / SOL_DOCX
    sol_xlsx = OUT / SOL_XLSX
    build_solution_doc(sol_docx)
    build_solution_wb(sol_xlsx)
    build_guide_html()
    archive = build_zip()
    print(
        f"已生成：\n  {docx}\n  {xlsx}\n  {pptx}\n  {comm_docx}\n  {comm_pdf}\n  {comm_xlsx}\n"
        f"  {lawyer_md}\n  {lawyer_docx}\n  {lawyer_pdf}\n  {lawyer_illu_docx}\n  {lawyer_illu_pdf}\n"
        f"  {liu_md}\n  {liu_docx}\n  {liu_pdf}\n  {mom_md}\n  {mom_docx}\n  {mom_pdf}\n"
        f"  {sol_docx}\n  {sol_xlsx}\n  {archive}"
    )


if __name__ == "__main__":
    main()
