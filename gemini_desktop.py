#!/usr/bin/env python3
"""兼容把 AIDEN 克隆到 ~/gemini-desktop 后，直接运行仓库根目录的本文件。"""

from __future__ import annotations

import runpy
import sys
from pathlib import Path

TARGET = Path(__file__).resolve().parent / "gemini-desktop" / "gemini_desktop.py"

if not TARGET.is_file():
    sys.stderr.write(
        "找不到 gemini-desktop/gemini_desktop.py。\n"
        "请切换到包含该目录的功能分支，然后执行：\n"
        "  python3 gemini-desktop/gemini_desktop.py --install\n"
    )
    raise SystemExit(1)

sys.argv[0] = str(TARGET)
runpy.run_path(str(TARGET), run_name="__main__")
