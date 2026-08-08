#!/usr/bin/env python3
"""一键生成《接手失败项目的通用框架》全部交付物。"""

from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_takeover_framework_ppt import build_ppt
from build_takeover_framework_docx import build_docx
from build_takeover_framework_excel import build_excel


def main():
    out = ROOT / "deliverables"
    out.mkdir(parents=True, exist_ok=True)
    build_ppt(out / "接手失败项目的通用框架.pptx")
    build_docx(out / "接手失败项目的通用框架_操作手册.docx")
    build_excel(out / "接手失败项目的通用框架_执行清单.xlsx")
    print("全部交付物已生成 →", out)


if __name__ == "__main__":
    main()
