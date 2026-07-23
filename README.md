# 复兴岛·全球创客岛收官答卷大会

**主呈现：PPT + Excel**

面向杨浦区复兴岛领导政绩呈现的大型活动策划：  
科创出海 × 人工智能 × 具身智能国际大会（含国际会议厅 / 会客厅揭牌）。

## 建议举办日期

**2026 年 9 月 12 日（星期六 · 农历八月初二）**

- 9 月 15 日收官节点前、最靠近的开市 / 挂匾黄道吉日
- 备选：9 月 9 日（周三）；规避：9 月 10 日（杨公忌日）

## 主呈现文件

| 文件 | 用途 |
| --- | --- |
| `deliverables/…策划方案.pptx` | **汇报主文件**（16 页：一页总览、择日、议程表、预算 KPI、决策清单） |
| `deliverables/…执行计划表.xlsx` | **执行台账**（15 张表：总览仪表盘 + 议程/嘉宾/揭牌/倒排/预算/风险等） |

Word 文稿仍可生成，但汇报与落地以 PPT + Excel 为准。

## 重新生成

```bash
pip install python-pptx python-docx openpyxl
python3 scripts/build_ppt.py    # PPT 主呈现
python3 scripts/build_xlsx.py   # Excel 执行台账
python3 scripts/build_docx.py   # 可选：完整文字版
```

结构化内容集中在 `scripts/content.py`。
