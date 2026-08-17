#!/usr/bin/env python3
"""一次生成 Word / Excel / PPT。"""

from pathlib import Path

from generate_docx import build_document
from generate_pptx import build_ppt
from generate_xlsx import build_workbook

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "deliverables"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    docx = OUT / "青岛红枫路交通事故_伤情与处理备忘录_20260817.docx"
    xlsx = OUT / "青岛红枫路交通事故_伤情伤残与行动表_20260817.xlsx"
    pptx = OUT / "青岛红枫路交通事故_伤情与伤残简报_20260817.pptx"
    build_document(docx)
    build_workbook(xlsx)
    build_ppt(pptx)
    print(f"已生成：\n  {docx}\n  {xlsx}\n  {pptx}")


if __name__ == "__main__":
    main()
