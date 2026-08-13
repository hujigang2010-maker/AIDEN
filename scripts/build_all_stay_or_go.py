#!/usr/bin/env python3
"""一次生成《发展与去留》全部交付物。"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_stay_or_go_docx import build_docx
from build_stay_or_go_excel import build_excel
from build_stay_or_go_ppt import build_ppt


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    out = root / "deliverables"
    out.mkdir(parents=True, exist_ok=True)
    build_ppt(out / "发展与去留_2026中美判断备忘录.pptx")
    build_docx(out / "发展与去留_2026中美判断备忘录.docx")
    build_excel(out / "发展与去留_2026中美判断备忘录_决策表.xlsx")
    print("全部交付物已生成。")


if __name__ == "__main__":
    main()
