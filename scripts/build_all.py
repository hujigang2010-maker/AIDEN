#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一次生成 Word / PPT / Excel 三份交付物。"""

from pathlib import Path

import build_proposal_docx as docx_mod
import build_proposal_pptx as pptx_mod
import build_proposal_xlsx as xlsx_mod


def main():
    out = Path(__file__).resolve().parents[1] / "output"
    out.mkdir(parents=True, exist_ok=True)
    p1 = pptx_mod.build()
    p2 = docx_mod.build()
    p3 = xlsx_mod.build()
    print("完成：")
    for p in (p1, p2, p3):
        print(" -", p, p.stat().st_size, "bytes")


if __name__ == "__main__":
    main()
