#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""核对申报表按原表上下排列，并已改参加工作时间、删去两家单位。"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / "output" / "专业技术资格评审申报表.xlsx"
HTML = ROOT / "output" / "专业技术资格评审申报表.html"

REMOVED = ("上海贝迪创建科技有限公司", "上海宝龙商办企业管理有限公司")


def main() -> None:
    subprocess.run([sys.executable, str(ROOT / "scripts" / "generate_application_tables.py")], check=True)
    wb = load_workbook(XLSX)
    assert wb.sheetnames == ["申报表"], wb.sheetnames
    ws = wb["申报表"]
    texts = [str(value) for row in ws.iter_rows(values_only=True) for value in row if value not in (None, "")]
    blob = "\n".join(texts)

    for name in REMOVED:
        assert name not in blob, name
    assert "2011年06月27日" in blob
    for row in ws.iter_rows(values_only=True):
        values = [str(value) for value in row if value not in (None, "")]
        if any("参加工作年月" in value for value in values):
            assert "2011年06月27日" in values
            assert "2015年09月" not in "".join(values)
            break
    else:
        raise AssertionError("未找到参加工作年月")

    import importlib.util

    spec = importlib.util.spec_from_file_location("tables", ROOT / "scripts" / "generate_application_tables.py")
    tables = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(tables)
    jobs = tables.JOBS
    assert len(jobs) == 16
    assert jobs[0][1] == "江苏省建筑装饰设计研究院有限公司"
    assert jobs[-1][1] == "上海睿泽股权投资管理有限公司"
    for period, org, _dept, _post, _work, title, short in jobs:
        assert org in blob
        assert title in blob
        assert short in blob
        assert period.split("\n")[0] in blob

    assert texts.count("江苏省建筑装饰设计研究院有限公司") >= 3
    assert "主持" in blob
    assert "合著" in blob and "独著" in blob
    assert "已具备" in blob
    for title in ("基  本  情  况", "工  作  经  历", "任现职以来考核情况", "评审、登记备案情况"):
        assert title in blob

    page = HTML.read_text(encoding="utf-8")
    for name in REMOVED:
        assert name not in page
    assert "2011年06月27日" in page
    assert 'id="basic"' in page and 'id="jobs"' in page and 'id="summary"' in page
    assert page.index("基本情况") < page.index('id="jobs"') < page.index("任现职后主要专业技术工作业绩登记")
    assert page.index('id="jobs"') < page.index('id="summary"') < page.index("任现职以来考核情况")
    jobs_html = page.split('id="jobs"', 1)[1].split('id="performance"', 1)[0]
    assert jobs_html.index("江苏省建筑装饰设计研究院有限公司") < jobs_html.index("上海睿泽股权投资管理有限公司")
    assert jobs_html.count("<tr>") == 17
    summary_html = page.split('id="summary"', 1)[1].split('id="review"', 1)[0]
    assert "上海贝迪" not in summary_html
    assert summary_html.count("<tr>") == 17
    from docx import Document
    from docx.oxml.ns import qn

    word = Document(ROOT / "output" / "专业技术资格评审申报表.docx")
    paragraphs = [item.text for item in word.paragraphs if item.text.strip()]
    assert paragraphs.index("基  本  情  况") < paragraphs.index("工  作  经  历") < paragraphs.index("任现职以来考核情况")
    pieces = list(paragraphs)
    for table in word.tables:
        for tr in table._tbl.tr_lst:
            for tc in tr.tc_lst:
                pieces.append("".join(node.text or "" for node in tc.iter(qn("w:t"))))
    blob_docx = "\n".join(pieces)
    for name in REMOVED:
        assert name not in blob_docx
    assert "2011年06月27日" in blob_docx
    assert "参加工作年月" in blob_docx
    assert blob_docx.index("江苏省建筑装饰设计研究院有限公司") < blob_docx.index("上海睿泽股权投资管理有限公司")
    assert "已具备" in blob_docx
    job_table = word.tables[4]
    job_orgs = []
    for tr in job_table._tbl.tr_lst[1:]:
        texts = ["".join(node.text or "" for node in tc.iter(qn("w:t"))) for tc in tr.tc_lst]
        job_orgs.append(texts[1])
    assert len(job_orgs) == 16
    assert job_orgs[0].startswith("江苏省建筑装饰设计研究院有限公司")
    assert job_orgs[-1].startswith("上海睿泽股权投资管理有限公司")
    print(f"通过：Word 与表格上下排列，工作经历 {len(jobs)} 段，参加工作年月为 2011年06月27日")


if __name__ == "__main__":
    main()
