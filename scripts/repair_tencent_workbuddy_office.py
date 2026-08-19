#!/usr/bin/env python3
"""修复腾讯 WorkBuddy 生成、Office/WPS 提示损坏或打不开的 pptx/docx/xlsx。

用法：
  python3 scripts/repair_tencent_workbuddy_office.py 损坏的文件.pptx
  python3 scripts/repair_tencent_workbuddy_office.py 某个目录 --inplace
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from office_compat import repair_office_file

OFFICE_SUFFIXES = {".pptx", ".docx", ".xlsx"}


def collect(paths: list[Path]) -> list[Path]:
    files: list[Path] = []
    for path in paths:
        if path.is_dir():
            for child in sorted(path.rglob("*")):
                if child.is_file() and child.suffix.lower() in OFFICE_SUFFIXES:
                    files.append(child)
        elif path.is_file():
            files.append(path)
        else:
            raise FileNotFoundError(path)
    return files


def main() -> int:
    parser = argparse.ArgumentParser(description="修复腾讯 WorkBuddy 生成的 Office 文件")
    parser.add_argument("inputs", nargs="+", type=Path, help="文件或目录")
    parser.add_argument(
        "--inplace",
        action="store_true",
        help="原地覆盖；默认写到同目录 *_已修复.ext",
    )
    args = parser.parse_args()

    files = collect(args.inputs)
    if not files:
        print("没有找到 pptx/docx/xlsx", file=sys.stderr)
        return 1

    for src in files:
        dst = src if args.inplace else None
        out = repair_office_file(src, dst)
        print(f"已修复：{src.name} -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
