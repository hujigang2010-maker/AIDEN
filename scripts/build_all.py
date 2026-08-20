# -*- coding: utf-8 -*-
"""一次性生成 PPT / Word / Excel / 微信稿 / HTML。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_docx import build as build_docx
from build_html import build as build_html
from build_ppt import build as build_ppt
from build_wechat import build as build_wechat
from build_xlsx import build as build_xlsx


def main():
    build_wechat()
    build_ppt()
    build_docx()
    build_xlsx()
    build_html()
    print("全部交付物已生成到 exports/")


if __name__ == "__main__":
    main()
