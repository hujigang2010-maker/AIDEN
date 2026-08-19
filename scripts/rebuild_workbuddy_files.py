#!/usr/bin/env python3
"""重新生成全部 WorkBuddy 交付物，并做 Office 兼容修复。"""

from __future__ import annotations

import runpy
from pathlib import Path

from office_compat import repair_office_file

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = Path(__file__).resolve().parent
DELIVERABLES = ROOT / "deliverables"


def main() -> None:
    DELIVERABLES.mkdir(parents=True, exist_ok=True)
    generators = [
        "generate_tailong_workbuddy_report_ppt.py",
        "generate_tailong_workbuddy_report_docx.py",
        "generate_tailong_workbuddy_v3_ppt.py",
        "generate_tailong_workbuddy_v3_excel.py",
        "generate_workbuddy_intro_ppt.py",
        "generate_workbuddy_intro_docx.py",
        "generate_workbuddy_intro_excel.py",
    ]
    for name in generators:
        print(f"==> {name}")
        runpy.run_path(str(SCRIPTS / name), run_name="__main__")

    for path in sorted(DELIVERABLES.glob("*")):
        if path.suffix.lower() in {".pptx", ".docx", ".xlsx"}:
            repair_office_file(path, path)
            print(f"已兼容修复：{path.name}")


if __name__ == "__main__":
    main()
