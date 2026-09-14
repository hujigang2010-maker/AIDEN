#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""校验第八版：两版优胜劣汰。印刷面无待核、无引进华为。页数为 A4 两页。"""

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
    "新质产业空间",
    "国资合作",
    "主办方代表发言",
    "北欧创新国际会客厅",
    "靖江印象城",
    "金坛理想城",
    "新城控股",
    "中南",
    "江阴白鹭湾",
    "PMP",
    "苏中县市",
    "房屋建筑工程人工智能",
    "2011–2021",
    "求职意向",
    "土木本科",
    "锦天城",
    "人工智能商业化",
    "能力切片",
    "IELTS 6.0",
]

FORBIDDEN_PRINT = [
    "复合型专家",
    "业绩卓著",
    "三个代表案例",
    "220283",
    "爱国路",
    "马喜艳",
    "待核",
    "待补",
    "原稿口径",
    "原稿未列总价",
    "告诉我即可",
    "李祥",
    "登台战略合作",
    "山东商会",
    "面试官",
    "引进华为",
    "波士顿科学",
    "创智天地",
    "超级孵化器",
]


def test_files_exist():
    required = [
        "胡继刚-简历-优化版.docx",
        "胡继刚-简历-优化版.pdf",
        "胡继刚-简历-优化版.html",
        "胡继刚-简历-一页精华.pdf",
        "胡继刚-简历-一页精华.html",
        "简历优化说明.md",
        "2026已签约产业合作-面试口径.md",
    ]
    missing = [name for name in required if not (OUT / name).exists()]
    assert not missing, f"缺少文件：{missing}"


def test_docx_is_valid():
    path = OUT / "胡继刚-简历-优化版.docx"
    assert zipfile.is_zipfile(path), "docx 不是有效 zip"
    with zipfile.ZipFile(path) as zf:
        assert "word/document.xml" in zf.namelist()
        xml = zf.read("word/document.xml").decode("utf-8")
    for key in ("胡继刚", "在职攻读", "城市更新", "产业招商", "智能制造", "创智汇", "东方枢纽", "方案拟定", "北欧创新国际会客厅", "PMP"):
        assert key in xml, f"Word 中缺少：{key}"
    for bad in FORBIDDEN_PRINT:
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
    for bad in FORBIDDEN_PRINT:
        assert bad not in zh, f"中文稿不应出现：{bad}"
    assert "产业招商" in zh
    assert "智能制造" in zh
    assert "已签约产业合作" in zh
    assert "创智汇" in zh
    assert "东方枢纽" in zh
    assert "方案拟定" in zh
    assert "锦天城" in zh
    assert "人工智能商业化" in zh
    assert "求职意向" in zh
    assert "土木本科" in zh
    assert "能力切片" in zh
    assert "创始理事长" in zh
    assert "不是雇主" in zh
    assert "PMP" in zh
    assert "苏中县市" in zh
    assert "房屋建筑工程人工智能" in zh
    assert "Hu Jigang" in html
    assert "导出 PDF" in html
    assert "优化说明" in html
    assert "秘书长 / 副教授级高级工程师" not in zh
    assert html.find("现任复旦大学住房政策研究中心秘书长") != -1
    assert "参加" not in zh
    assert "登台战略合作" not in zh
    assert "AI专家" not in zh
    assert "服务100家" not in zh
    assert "三条证明" not in zh
    assert zh.find("产业招商") < zh.find("如何落地城市")
    notes = html[html.find('id="pack-notes"'):]
    assert "李祥" in notes


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
    full_text = (page1 + page2).replace("\n", "")
    page1f = page1.replace("\n", "")
    page2f = page2.replace("\n", "")
    for key in KEYWORDS:
        assert key in full_text, f"PDF 中缺少：{key}"
    assert "万科" in page1f
    assert "复旦大学住房政策研究中心" in page1f
    assert "产业招商" in page1f
    assert "智能制造" in page1f
    assert "创智汇" in page1f
    assert "东方枢纽" in page1f
    assert "华为汽车" in page1f
    assert "森马产业园" in page1f
    assert "方案拟定" in page1f
    assert "主办方代表发言" in page1f
    assert "北欧创新国际会客厅" in page1f
    assert "锦天城" in page1f
    assert "人工智能商业化" in page1f
    assert "求职意向" in page1f
    assert "土木本科" in page1f
    assert "能力切片" in page1f
    assert "三条证明" not in page1f
    assert "新城控股" in page2f
    assert "中南" in page2f
    assert "江阴白鹭湾" in page2f
    assert "代表项目" in page2f
    assert "在职攻读" in page2f
    assert "PMP" in page2f
    assert "房屋建筑工程人工智能" in page2f
    vanke = page1f[page1f.find("万科企业"):]
    assert vanke.find("镇江") < vanke.find("靖江印象城")
    for bad in FORBIDDEN_PRINT:
        assert bad not in full_text, f"PDF 中不应出现：{bad}"
    brief_text = brief[0].get_text().replace("\n", "")
    assert "胡继刚" in brief_text
    assert "创智汇" in brief_text
    assert "产业招商" in brief_text
    assert "在职攻读" in brief_text
    assert "万科" in brief_text
    for bad in FORBIDDEN_PRINT:
        assert bad not in brief_text, f"精华版不应出现：{bad}"


def test_interview_brief():
    text = (OUT / "2026已签约产业合作-面试口径.md").read_text(encoding="utf-8")
    for key in ("甲方抬头", "签约日", "华为", "143", "133", "定位招商", "引进华为", "盖章"):
        assert key in text, f"面试口径缺少：{key}"


def main():
    test_files_exist()
    test_docx_is_valid()
    test_html_keywords()
    test_pdf_pages()
    test_interview_brief()
    print("简历交付物校验通过")


if __name__ == "__main__":
    main()
