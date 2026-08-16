#!/usr/bin/env python3
"""一键生成 WorkBuddy 银行引荐判断包（Word + PPT + Excel）。"""

from pathlib import Path

from generate_workbuddy_intro_docx import build_document
from generate_workbuddy_intro_excel import build_workbook
from generate_workbuddy_intro_ppt import build_ppt


def main() -> None:
    root = Path(__file__).resolve().parents[1] / "deliverables"
    root.mkdir(parents=True, exist_ok=True)
    build_document(root / "WorkBuddy银行引荐判断备忘录_杨行长约见_20260816.docx")
    build_ppt(root / "WorkBuddy银行引荐判断_杨行长约见简报_20260816.pptx")
    build_workbook(root / "WorkBuddy银行引荐_三家银行对照与十天行动表_20260816.xlsx")


if __name__ == "__main__":
    main()
