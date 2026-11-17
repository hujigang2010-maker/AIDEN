# -*- coding: utf-8 -*-
"""一次生成合作框架 Word、洽谈 PPT 和工作表。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_docx import build as build_docx
from build_ppt import build as build_ppt
from build_xlsx import build as build_xlsx

OUT = Path(__file__).resolve().parents[1] / "output"


def main():
    OUT.mkdir(exist_ok=True)
    docx = OUT / "复旦终身教育学院与杨浦科联合作框架方案.docx"
    pptx = OUT / "复旦终身教育学院与杨浦科联合作洽谈.pptx"
    xlsx = OUT / "复旦终身教育学院与杨浦科联合作工作表.xlsx"
    build_docx(docx)
    build_ppt(pptx)
    build_xlsx(xlsx)
    for path in (docx, pptx, xlsx):
        print(f"{path.name}\t{path.stat().st_size}")


if __name__ == "__main__":
    main()
