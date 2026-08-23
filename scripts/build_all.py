# -*- coding: utf-8 -*-
"""一次性生成全部交付物。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_docx import build as build_docx
from build_ics import build as build_ics
from build_poster import build as build_poster
from build_ppt import build as build_ppt
from build_wechat import build as build_wechat
from build_xlsx import build as build_xlsx


def main():
    build_wechat()
    build_ppt()
    build_docx()
    build_xlsx()
    build_ics()
    build_poster()
    print("全部交付物已生成到 exports/")


if __name__ == "__main__":
    main()
