# -*- coding: utf-8 -*-
"""一次生成 Word 手册、Excel 模型和 PPT 简报。"""

from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))

import model_core
import build_docx
import build_xlsx
import build_ppt


def main():
    model_core.verify()
    paths = [build_docx.build(), build_xlsx.build(), build_ppt.build()]
    for path in paths:
        print(path)


if __name__ == "__main__":
    main()
