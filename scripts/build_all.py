# -*- coding: utf-8 -*-
"""一次性生成 PPT / Excel / Word。"""
from build_ppt import build as build_ppt
from build_xlsx import build as build_xlsx
from build_docx import build as build_docx


def main():
    build_ppt()
    build_xlsx()
    build_docx()
    print("全部交付物已生成到 deliverables/")


if __name__ == "__main__":
    main()
