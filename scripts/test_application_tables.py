#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""核对申报表整理结果：栏目、行数和关键原文。"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / "output" / "专业技术资格评审申报表.xlsx"
HTML = ROOT / "output" / "专业技术资格评审申报表.html"


def rows(ws):
    data = []
    for excel_row in ws.iter_rows(min_row=4, values_only=True):
        if any(cell not in (None, "") for cell in excel_row):
            data.append(tuple("" if cell is None else str(cell) for cell in excel_row))
    return data


def main() -> None:
    subprocess.run([sys.executable, str(ROOT / "scripts" / "generate_application_tables.py")], check=True)
    wb = load_workbook(XLSX)
    expected = [
        "说明",
        "封面",
        "个人承诺",
        "基本情况",
        "学历",
        "奖惩",
        "继续教育",
        "工作经历",
        "工作业绩",
        "论文论著",
        "总结项目",
        "总结补充",
        "考核与评审",
        "待核查",
    ]
    assert wb.sheetnames == expected, wb.sheetnames

    cover = {item: value for item, value in rows(wb["封面"])}
    assert cover["单位"] == "江苏省建筑装饰设计研究院有限公司"
    assert cover["编号"] == "2506170651940505"
    assert cover["填表时间"] == "2025年08月12日"

    basic = {item: value for item, value in rows(wb["基本情况"])}
    assert basic["公民身份号码（社会保障号）"] == "220283198704250614"
    assert basic["参加工作年月"] == "2015年09月"
    assert basic["专业技术工作年限"] == "9年"
    assert basic["准入资格、时间及证书号码"] == "（原表未填）"

    edu = rows(wb["学历"])
    assert len(edu) == 3
    assert edu[0][0] == "复旦大学" and edu[0][4] == "于2021年01月13日"
    assert edu[2][3] == "其他" and "2016年07月04日" in edu[2][4]

    awards = rows(wb["奖惩"])
    assert [item[0] for item in awards] == [
        "住房政策研究创新贡献奖",
        "城市更新卓越实践奖",
        "可持续社区营造标杆项目奖",
    ]

    study = rows(wb["继续教育"])
    assert len(study) == 3
    assert "240" in study[0][3] and "16.5" in study[2][3]

    jobs = rows(wb["工作经历"])
    assert len(jobs) == 18, len(jobs)
    assert jobs[0][2] == "上海贝迪创建科技有限公司"
    assert jobs[0][1] == "2026年05月至今"
    assert jobs[0][6] == "高级工程师"
    assert jobs[-1][2] == "上海睿泽股权投资管理有限公司"
    assert jobs[-1][4] == "技术岗"
    assert jobs[-1][6] == "资料员、施工员"
    assert "空档 2015-11" in str(wb["工作经历"]["A2"].value)

    performance = rows(wb["工作业绩"])
    assert len(performance) == 3
    assert all(item[6] == "主持" and item[7] == "待核查" for item in performance)

    papers = rows(wb["论文论著"])
    assert [item[4] for item in papers] == ["合著", "合著", "独著"]
    assert papers[0][1].startswith("铝膜板")

    projects = rows(wb["总结项目"])
    assert len(projects) == 7
    assert projects[1][6] == ""
    assert projects[1][1] == "新城香溢紫郡"
    assert sum(1 for item in projects if item[6] == "待核查") == 6
    assert projects[-1][0] == "6" and "东台" in projects[-1][1]

    review = { (part, item): value for part, item, value in rows(wb["考核与评审"]) }
    assert review[("年度考核", "考核情况")] == "无"
    assert review[("评议组", "赞成票")] == "3"
    assert review[("评审委员会", "赞成票")] == "27"
    assert "已具备" in review[("评审委员会", "评审意见")]
    assert review[("备注", "联系电话")] == "18678408669"
    assert review[("登记备案", "日期")] == "2025年12月09日"

    pending = rows(wb["待核查"])
    assert len(pending) == 9
    assert all("待核查" in item[2] for item in pending)

    page = HTML.read_text(encoding="utf-8")
    assert 'name="viewport"' in page
    for token in (
        "上海贝迪创建科技有限公司",
        "220283198704250614",
        "待核查",
        "新城香溢紫郡",
        "不是 PDF",
    ):
        assert token in page, token

    print(f"通过：{len(wb.sheetnames)} 张表，工作经历 {len(jobs)} 行，待核查 {len(pending)} 条")


if __name__ == "__main__":
    main()
