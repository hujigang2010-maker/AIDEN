#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""校验 0921 修订稿：新联系方式、学历口径、现职课题、万科/新城/中南履历，以及 A4 页数。"""

from __future__ import annotations

import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output"

REQUIRED = [
    "胡继刚",
    "13262607888",
    "hujigang2010@gmail.com",
    "hu13262607888",
    "复旦大学硕士研究生",
    "复旦大学不动产资产管理协会创始理事长",
    "不是技术研发岗",
    "新质产业空间",
    "北欧创新国际会客厅",
    "杨浦宝龙旭辉广场环创中心",
    "宝龙商办不是雇主",
    "人民网",
    "新华社",
    "中新社",
    "央广网",
    "上海市人民政府",
    "万科企业股份有限公司",
    "镇江大港金域蓝湾",
    "配建安置房",
    "万纬嘉兴平湖",
    "盐城",
    "台州",
    "集团首进山东",
    "同股同权",
    "非已签约落地",
    "新城控股集团股份有限公司",
    "中南建设集团",
]

FORBIDDEN = [
    "18678408669",
    "262782809@qq.com",
    "hu262782809",
    "复旦 MBA",
    "复旦MBA",
    "工商管理硕士",
    "不动产资产管理协会秘书长",
    "住房保障、租赁",
    "220283",
    "爱国路",
    "马喜艳",
]


def _norm(text: str) -> str:
    return text.replace("\n", "").replace(" ", "")


def test_files_exist():
    required = [
        "胡继刚-简历-优化版.docx",
        "胡继刚-简历-优化版.pdf",
        "胡继刚-简历-优化版0921.pdf",
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
    for key in REQUIRED:
        assert key in xml, f"Word 中缺少：{key}"
    for bad in FORBIDDEN:
        assert bad not in xml, f"Word 中不应出现：{bad}"


def test_html_keywords():
    html = (OUT / "胡继刚-简历-优化版.html").read_text(encoding="utf-8")
    brief = (OUT / "胡继刚-简历-一页精华.html").read_text(encoding="utf-8")
    for key in REQUIRED:
        assert key in html, f"完整版 HTML 缺少：{key}"
    for bad in FORBIDDEN:
        assert bad not in html, f"完整版 HTML 不应出现：{bad}"
        assert bad not in brief, f"精华版 HTML 不应出现：{bad}"
    for key in (
        "13262607888",
        "hujigang2010@gmail.com",
        "hu13262607888",
        "复旦大学硕士研究生",
        "不是技术研发岗",
        "宝龙商办不是雇主",
        "新华社",
        "同股同权",
    ):
        assert key in brief, f"精华版 HTML 缺少：{key}"


def test_pdf_pages():
    try:
        import pymupdf
    except ImportError:
        print("跳过 PDF 页数检查（未安装 pymupdf）")
        return
    full = pymupdf.open(OUT / "胡继刚-简历-优化版.pdf")
    dated = pymupdf.open(OUT / "胡继刚-简历-优化版0921.pdf")
    brief = pymupdf.open(OUT / "胡继刚-简历-一页精华.pdf")
    assert full.page_count == 2, f"优化版应为 2 页，实际 {full.page_count}"
    assert dated.page_count == 2, f"0921 副本应为 2 页，实际 {dated.page_count}"
    assert brief.page_count == 1, f"精华版应为 1 页，实际 {brief.page_count}"
    assert abs(full[0].rect.width - 595.3) < 8, f"优化版应为 A4 宽，实际 {full[0].rect.width}"
    assert abs(full[0].rect.height - 841.9) < 8, f"优化版应为 A4 高，实际 {full[0].rect.height}"
    page1 = full[0].get_text()
    page2 = full[1].get_text()
    full_text = page1 + page2
    for key in REQUIRED:
        assert key in full_text, f"PDF 中缺少：{key}"
    for bad in FORBIDDEN:
        assert bad not in full_text, f"PDF 中不应出现：{bad}"
    assert "复旦大学住房政策研究中心" in page1
    assert "万科企业股份有限公司" in page1
    assert "镇江大港金域蓝湾" in page1
    assert "新城控股" in page2
    assert "中南建设" in page2
    assert "复旦大学硕士研究生" in page2
    assert "复旦大学不动产资产管理协会创始理事长" in page2
    brief_text = brief[0].get_text()
    for key in ("13262607888", "hujigang2010@gmail.com", "复旦大学硕士研究生", "宝龙商办不是雇主", "新华社"):
        assert key in brief_text, f"精华版 PDF 缺少：{key}"
    for bad in FORBIDDEN:
        assert bad not in brief_text, f"精华版 PDF 不应出现：{bad}"


def test_pdf_not_clipped():
    """关键长句必须完整出现，避免 overflow:hidden 裁切。"""
    try:
        import pymupdf
    except ImportError:
        return
    full = pymupdf.open(OUT / "胡继刚-简历-优化版.pdf")
    text = _norm("".join(p.get_text() for p in full))
    must = [
        "关注产业如何用空间、政策、资本落地城市",
        "落户杨浦宝龙旭辉广场环创中心",
        "推动进入盐城、台州、常州金坛",
        "系集团首进山东的重资产项目",
        "推进中，非已签约落地",
    ]
    for key in must:
        assert _norm(key) in text, f"PDF 可能被裁切，缺少完整句：{key}"


def main():
    test_files_exist()
    test_docx_is_valid()
    test_html_keywords()
    test_pdf_pages()
    test_pdf_not_clipped()
    print("简历交付物校验通过")


if __name__ == "__main__":
    main()
