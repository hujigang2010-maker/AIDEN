# 临港 × WeTest · 中日数字经济跨国发展与高质量交付闭门会

> 政策护航（临港新片区） × 质量守护（WeTest） —— 一站式中日出海/入华解决方案沙龙
>
> 本仓库为一份完整的活动策划交付物，**核心抓手是「会后转化」**：从 75 人邀请池 → 25 人到场 → 3 单签约（T+90 内），目标 ROI ≥ 5×。

## 📦 交付物清单（`deliverables/`）

| 类型 | 文件 | 用途 |
| --- | --- | --- |
| Word | `临港x WeTest 中日闭门会_活动策划方案_V1.0.docx` | 详细策划方案（13 章 + 2 附录），用于与临港/合作伙伴对齐方案、内部立项汇报 |
| PPT | `临港x WeTest 中日闭门会_宣讲与策划PPT_V1.0.pptx` | 25 页演示稿，用于内部立项、对外宣讲、现场展示 |
| Excel | `临港x WeTest 中日闭门会_转化运营管理表_V1.0.xlsx` | 14 个 Sheet，含 Dashboard / 邀约 / 现场 / 转化漏斗 / SOP / KPI / 预算 / 风险 / 复盘 |

## 🎯 转化主线

```
邀请池 75  →  到场 25  →  强意向 15  →  商机 9 (T+7)  →  POC 5 (T+30)  →  签单 3 (T+90)
```

* 会前：议题征集 + 白皮书 + TOP10 1V1 预热
* 会中：8 个转化抓手（体验角、限时权益券、圆桌 1V1 承接、闭门通道等）
* 会后：48h 黄金期 SOP + ABC 分级 + 7-30-60-90 节奏

## 🛠️ 重新生成方式

```bash
pip install python-docx python-pptx openpyxl
python3 scripts/build_word.py
python3 scripts/build_ppt.py
python3 scripts/build_excel.py
```

## 📂 仓库结构

```
deliverables/    # 三份最终交付文件
scripts/         # 生成脚本（可改参数 / 颜色 / 内容后重新生成）
README.md        # 本说明
```
