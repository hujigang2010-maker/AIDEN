# -*- coding: utf-8 -*-
"""一次生成综合方案 PPT 与配套工作簿。"""

import os
import sys

from build_pptx import main as build_ppt
from build_xlsx import main as build_xlsx


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_dir = os.path.join(root, "deliverables")
    os.makedirs(out_dir, exist_ok=True)
    sys.argv = [
        "build_pptx.py",
        os.path.join(out_dir, "潮起潮落TIDELINE-北外滩尚9一滴水-综合合作方案.pptx"),
    ]
    build_ppt()
    sys.argv = [
        "build_xlsx.py",
        os.path.join(out_dir, "潮起潮落TIDELINE-北外滩尚9一滴水-方案附表.xlsx"),
    ]
    build_xlsx()


if __name__ == "__main__":
    main()
