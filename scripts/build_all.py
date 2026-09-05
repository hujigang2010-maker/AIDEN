# -*- coding: utf-8 -*-
"""生成给对方看的汇报 PPT 和协议。"""

from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from generate_report_agreement import build as build_agreement
from generate_report_ppt import build as build_ppt


def main() -> None:
    files = [build_ppt(), build_agreement()]
    print("给对方的汇报版已生成：")
    for p in files:
        print(f"  {p}")


if __name__ == "__main__":
    main()
