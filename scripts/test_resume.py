#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""校验两页履历：口径、关键词、禁词、A4 页数。"""

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
    "产业招商",
    "智能制造",
    "已签约产业合作",
    "创智汇",
    "东方枢纽",
    "华为汽车",
    "森马产业园",
    "方案拟定",
    "国资",
    "主办方代表发言",
    "北欧创新国际会客厅",
    "靖江印象城",
    "金坛理想城",
    "新城控股",
    "中南",
    "江阴白鹭湾",
    "PMP",
    "房屋建筑工程人工智能",
    "2011–2021",
    "存量空间",
    "上海山东省商会",
    "面试官",
    "副教授级高级工程师为职称",
]

FORBIDDEN_PRINT = [
    "复合型专家",
    "业绩卓著",
    "220283",
    "爱国路",
    "马喜艳",
    "待核",
    "待补",
    "李祥",
    "引进华为",
    "波士顿科学",
    "创智天地",
    "城市更新研究会",
]


def test_files_exist():
    required = [
        "胡继刚-简历.docx",
        "胡继刚-简历.pdf",
        "胡继刚-简历.html",
        "排版说明.md",
    ]
    missing = [name for name in required if not (OUT / name).exists()]
    assert not missing, f"缺少文件：{missing}"


def test_docx_is_valid():
    path = OUT / "胡继刚-简历.docx"
    assert zipfile.is_zipfile(path), "docx 不是有效 zip"
    with zipfile.ZipFile(path) as zf:
        assert "word/document.xml" in zf.namelist()
        xml = zf.read("word/document.xml").decode("utf-8")
    for key in ("胡继刚", "在职攻读", "城市更新", "产业招商", "智能制造", "创智汇", "东方枢纽", "方案拟定", "北欧创新国际会客厅", "PMP"):
        assert key in xml, f"Word 中缺少：{key}"
    for bad in FORBIDDEN_PRINT:
        assert bad not in xml, f"Word 中不应出现：{bad}"
    assert "秘书长 / 副教授级高级工程师" not in xml
    assert "引进华为" not in xml


def test_html_keywords():
    html = (OUT / "胡继刚-简历.html").read_text(encoding="utf-8")
    for key in KEYWORDS:
        assert key in html, f"稿件缺少：{key}"
    for bad in FORBIDDEN_PRINT:
        assert bad not in html, f"稿件不应出现：{bad}"
    assert "现任复旦大学住房政策研究中心秘书长" in html
    assert "求职方向" in html
    assert "闵行森马产业园" in html
    assert "AI专家" not in html
    assert "服务100家" not in html


def test_pdf_pages():
    import pymupdf

    full = pymupdf.open(OUT / "胡继刚-简历.pdf")
    assert full.page_count == 2, f"应为 2 页，实际 {full.page_count}"
    assert abs(full[0].rect.width - 595.3) < 8, f"应为 A4 宽，实际 {full[0].rect.width}"
    assert abs(full[0].rect.height - 841.9) < 8, f"应为 A4 高，实际 {full[0].rect.height}"
    page1 = full[0].get_text()
    page2 = full[1].get_text()
    full_text = page1 + page2
    for key in KEYWORDS:
        assert key in full_text, f"PDF 中缺少：{key}"
    assert "万科" in page1
    assert "复旦大学住房政策研究中心" in page1
    assert "创智汇" in page1
    assert "东方枢纽" in page1
    assert "华为汽车" in page1
    assert "森马产业园" in page1
    assert "方案拟定" in page1
    assert "主办方代表发言" in page1
    assert "北欧创新国际会客厅" in page1
    assert "新城控股" in page2
    assert "中南" in page2
    assert "江阴白鹭湾" in page2
    assert "代表项目" in page2
    assert "在职攻读" in page2
    assert "PMP" in page2
    for bad in FORBIDDEN_PRINT:
        assert bad not in full_text, f"PDF 中不应出现：{bad}"
    # 抽查无缺字框
    assert "□" not in full_text
    assert "\ufffd" not in full_text


def test_docx_two_pages():
    """若本机有 LibreOffice，校验 Word 亦为两页。"""
    import shutil
    import subprocess
    import tempfile

    if not shutil.which("soffice"):
        print("跳过 Word 页数检查（未安装 LibreOffice）")
        return
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, str(OUT / "胡继刚-简历.docx")],
            check=True,
            timeout=60,
        )
        import pymupdf

        pdfs = list(Path(tmp).glob("*.pdf"))
        assert pdfs, "Word 转 PDF 失败"
        doc = pymupdf.open(pdfs[0])
        assert doc.page_count == 2, f"Word 应为 2 页，实际 {doc.page_count}"
        text = "".join(page.get_text() for page in doc)
        assert "万科" in text and "中南" in text and "创智汇" in text


def main():
    test_files_exist()
    test_docx_is_valid()
    test_html_keywords()
    test_pdf_pages()
    test_docx_two_pages()
    print("简历交付物校验通过")


if __name__ == "__main__":
    main()
