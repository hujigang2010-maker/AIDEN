# -*- coding: utf-8 -*-
"""一次性生成对外汇报版 PPT / Word / Excel。"""
from build_ppt import build as build_ppt
from build_xlsx import build as build_xlsx
from build_docx import build as build_docx


def main():
    build_ppt()
    build_xlsx()
    build_docx()
    print("全部对外汇报交付物已生成到 deliverables/")


if __name__ == "__main__":
    main()
