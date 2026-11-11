"""生成康博会分论坛全套交付文件。"""

from __future__ import annotations

import build_excel
import build_exec_word
import build_internal_ppt
import build_sponsor_ppt
import build_sponsor_word


def main():
    paths = [
        build_internal_ppt.build(),
        build_exec_word.build(),
        build_excel.build(),
        build_sponsor_word.build(),
        build_sponsor_ppt.build(),
    ]
    for path in paths:
        print(path)


if __name__ == "__main__":
    main()
