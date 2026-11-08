#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""核对职业履历与核心项目整理稿没有串入证件号，并保留申报表原句。"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
from generate_career_brief import (  # noqa: E402
    DOCX_NAME,
    FORBIDDEN,
    FORM_ROWS,
    HTML_NAME,
    OUT_DIR,
    PROJECTS,
    PROOF_ROWS,
    SUMMARY_ROWS,
    XLSX_NAME,
    main,
)


def text_of_docx(path: Path) -> str:
    doc = Document(path)
    parts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            parts.extend(cell.text for cell in row.cells)
    return "\n".join(parts)


def main_test() -> None:
    main()
    html = (OUT_DIR / HTML_NAME).read_text(encoding="utf-8")
    docx_text = text_of_docx(OUT_DIR / DOCX_NAME)
    wb = load_workbook(OUT_DIR / XLSX_NAME)
    combined = html + docx_text
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=True):
            combined += "\n" + " ".join("" if value is None else str(value) for value in row)

    required = [
        "江苏中南建设集团股份有限公司",
        "施工员/资料员",
        "上海万科企业有限公司",
        "金科地产集团有限公司华东区域公司",
        "万科（上海）实业发展有限公司",
        "上海睿泽股权投资管理有限公司",
        "上海苏科企业管理有限公司",
        "钢结构施工技术管理",
        "中南海湾新城·金石国际广场（2011.07–2015.08）",
        "超高层五星级酒店公寓（26万㎡，120m）",
        "施工技术管理",
        "控制结构偏移误差＜3mm",
        "节约工期45天，获江苏省优质工程奖",
        "新城香溢紫郡",
        "降低地基处理成本23%",
        "合肥金科博翠天辰",
        "项目获合肥市绿色建筑示范工程认证",
        "系统能效比提升至4.8",
        "上海万科天空之城",
        "获LEED金级认证",
        "ZL202020536480.4",
        "建筑工程管理用功能脚架",
        "杭州湾融创文旅城水路乐园项目（海盐）",
        "地震模拟性能提升35%",
        "苏州绿地江南映月花园",
        "施工碳排放降低28%",
        "绿地苏州东太湖旅游度假区",
        "累计节约项目成本2,800万元",
        "东台市养老服务中心地下车库装饰工程",
        "减少返工损失150万元",
        "偏差超5%自动预警",
        "来源是一览表，不是本轮.doc正文",
    ]
    missing = [item for item in required if item not in combined]
    if missing:
        raise SystemExit("缺少原句：\n" + "\n".join(missing))
    leaked = [item for item in FORBIDDEN if item in combined]
    if leaked:
        raise SystemExit("不应出现：\n" + "\n".join(leaked))
    if len(SUMMARY_ROWS) != 7 or len(FORM_ROWS) != 8 or len(PROOF_ROWS) != 15 or len(PROJECTS) != 8:
        raise SystemExit("行数不对")
    sheets = set(wb.sheetnames)
    expected = {"整理说明", "工作总结履历", "申报表工作经历", "项目证明对照", "核心项目"}
    if sheets != expected:
        raise SystemExit(f"工作表不符：{sheets}")
    proof_sheet = wb["项目证明对照"]
    proof_values = [proof_sheet.cell(row, 2).value for row in range(4, 19)]
    if proof_values[0] != "江苏中南建设集团股份有限公司":
        raise SystemExit(f"项目证明顺序不对：{proof_values[0]}")
    if proof_values[-1] != "上海徐汇绿地商业管理有限公司":
        raise SystemExit(f"末行单位不对：{proof_values[-1]}")
    print("ok", len(combined))


if __name__ == "__main__":
    main_test()
