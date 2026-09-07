#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""校验简历交付物：文件存在、关键表述齐全、PDF 页数正确。"""

from __future__ import annotations

import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output"

KEYWORDS = [
    "胡继刚",
    "产业合作",
    "招商",
    "AI",
    "出海",
    "北欧创新国际会客厅",
    "靖江印象城",
    "三个代表案例",
    "复旦大学住房政策研究中心",
]


def test_files_exist():
    required = [
        "胡继刚-简历-优化版.docx",
        "胡继刚-简历-优化版.pdf",
        "胡继刚-简历-优化版.html",
        "胡继刚-简历-一页精华.pdf",
        "胡继刚-简历-一页精华.html",
        "简历优化说明.md",
    ]
    missing = [name for name in required if not (OUT / name).exists()]
    assert not missing, f"缺少文件：{missing}"


def test_docx_is_valid():
    path = OUT / "胡继刚-简历-优化版.docx"
    assert zipfile.is_zipfile(path), "docx 不是有效 zip"
    with zipfile.ZipFile(path) as zf:
        assert "word/document.xml" in zf.namelist()
        xml = zf.read("word/document.xml").decode("utf-8")
    for key in ("胡继刚", "招商", "出海"):
        assert key in xml, f"Word 中缺少：{key}"


def test_html_keywords():
    html = (OUT / "胡继刚-简历-优化版.html").read_text(encoding="utf-8")
    for key in KEYWORDS:
        assert key in html, f"HTML 中缺少：{key}"
    assert "220283" not in html
    assert "爱国路" not in html
    assert "马喜艳" not in html
    assert "不写入现任" in html
    assert "李祥" in html


def test_pdf_pages():
    try:
        import pymupdf
    except ImportError:
        print("跳过 PDF 页数检查（未安装 pymupdf）")
        return
    full = pymupdf.open(OUT / "胡继刚-简历-优化版.pdf")
    brief = pymupdf.open(OUT / "胡继刚-简历-一页精华.pdf")
    assert full.page_count == 2, f"优化版应为 2 页，实际 {full.page_count}"
    assert brief.page_count == 1, f"精华版应为 1 页，实际 {brief.page_count}"
    full_text = "".join(p.get_text() for p in full)
    for key in KEYWORDS:
        assert key in full_text, f"PDF 中缺少：{key}"
    assert "新城控股" in full_text
    assert "中南" in full_text
    assert "不写入现任" in full_text
    assert "李祥" in full_text


def main():
    test_files_exist()
    test_docx_is_valid()
    test_html_keywords()
    test_pdf_pages()
    print("简历交付物校验通过")


if __name__ == "__main__":
    main()
