#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""校验简历交付物：文件存在、关键表述齐全、PDF 页数与分页正确。"""

from __future__ import annotations

import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output"

KEYWORDS = [
    "胡继刚",
    "复旦大学住房政策研究中心",
    "投资副总经理",
    "在职攻读",
    "城市更新",
    "租赁",
    "靖江印象城",
    "金坛理想城",
    "新城控股",
    "中南",
    "2011–2021",
]

FORBIDDEN = [
    "复合型专家",
    "业绩卓著",
    "三个代表案例",
    "220283",
    "爱国路",
    "马喜艳",
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
    for key in ("胡继刚", "在职攻读", "城市更新", "靖江印象城"):
        assert key in xml, f"Word 中缺少：{key}"
    for bad in ("复合型专家", "业绩卓著", "马喜艳"):
        assert bad not in xml, f"Word 中不应出现：{bad}"


def _zh_pack(html: str) -> str:
    start = html.find('id="pack-zh"')
    end = html.find('id="pack-en"')
    assert start != -1 and end != -1 and end > start
    return html[start:end]


def test_html_keywords():
    html = (OUT / "胡继刚-简历-优化版.html").read_text(encoding="utf-8")
    zh = _zh_pack(html)
    for key in KEYWORDS:
        assert key in zh, f"中文稿缺少：{key}"
    for bad in FORBIDDEN:
        assert bad not in zh, f"中文稿不应出现：{bad}"
    assert "不写入现任" in zh
    assert "李祥" in zh
    assert "Hu Jigang" in html
    assert "导出 PDF" in html
    assert "优化说明" in html
    assert "秘书长 / 副教授级高级工程师" not in zh
    assert "副教授级高级工程师为职称" in zh


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
    assert abs(full[0].rect.width - 595.3) < 8, f"优化版应为 A4 宽，实际 {full[0].rect.width}"
    assert abs(full[0].rect.height - 841.9) < 8, f"优化版应为 A4 高，实际 {full[0].rect.height}"
    page1 = full[0].get_text()
    page2 = full[1].get_text()
    full_text = page1 + page2
    for key in KEYWORDS:
        assert key in full_text, f"PDF 中缺少：{key}"
    assert "万科" in page1, "第一页应有万科证明"
    assert "复旦大学住房政策研究中心" in page1, "第一页应有现任复旦职务"
    assert "新城控股" in page2, "第二页应有新城"
    assert "中南" in page2, "第二页应有中南"
    assert "代表项目" in page2, "第二页应有项目表"
    assert "在职攻读" in page2, "第二页应标明 MBA 在职攻读"
    assert "不写入现任" in full_text
    assert "李祥" in full_text
    for bad in FORBIDDEN:
        assert bad not in full_text, f"PDF 中不应出现：{bad}"
    brief_text = brief[0].get_text()
    assert "胡继刚" in brief_text
    assert "在职攻读" in brief_text
    assert "李祥" in brief_text


def main():
    test_files_exist()
    test_docx_is_valid()
    test_html_keywords()
    test_pdf_pages()
    print("简历交付物校验通过")


if __name__ == "__main__":
    main()
