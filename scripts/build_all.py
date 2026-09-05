# -*- coding: utf-8 -*-
"""一次生成全部对外交付物。"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from generate_agreement import build as build_agreement
from generate_excel import build as build_excel
from generate_md import build as build_md
from generate_onepager import build as build_onepager
from generate_ppt import build as build_ppt


def main() -> None:
    out = Path(__file__).resolve().parent.parent / "output"
    out.mkdir(parents=True, exist_ok=True)
    files = [
        build_ppt(),
        build_agreement(),
        build_excel(),
        build_onepager(),
        *build_md(),
    ]
    print("全部交付物已生成：")
    for p in files:
        print(f"  {p}")


if __name__ == "__main__":
    main()
