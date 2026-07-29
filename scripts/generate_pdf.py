# -*- coding: utf-8 -*-
"""将网页版宣传册导出为 PDF。"""
from pathlib import Path

from weasyprint import HTML

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "deliverables" / "网页版宣传册_复旦杨浦科创生态共建计划.html"
OUT = ROOT / "deliverables" / "复旦杨浦科创生态共建计划_宣传册.pdf"


def build():
    HTML(filename=str(SRC)).write_pdf(str(OUT))
    print(f"已生成：{OUT}")


if __name__ == "__main__":
    build()
