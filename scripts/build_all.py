# -*- coding: utf-8 -*-
"""一次生成 PPT / Word / Excel 三件套。"""
from pathlib import Path

import build_docx
import build_ppt
import build_xlsx

ROOT = Path(__file__).resolve().parent.parent


def main():
    ppt = build_ppt.build()
    doc = build_docx.build()
    xls = build_xlsx.build()
    print("---")
    for p in (ppt, doc, xls):
        size = p.stat().st_size
        print(f"{p.relative_to(ROOT)}  {size} bytes")


if __name__ == "__main__":
    main()
