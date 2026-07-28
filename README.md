# 学生实习赋能计划 · 可行性论证方案

面向高中生、大学生、毕业生三类人群的实习服务体系可行性论证材料。

## 交付物

| 文件 | 说明 |
|------|------|
| `exports/学生实习赋能计划_可行性论证方案.pptx` | 论证 PPT（背景、人群、产品、定价、盈利、落地、风险） |
| `exports/学生实习赋能计划_可行性论证数据表.xlsx` | 配套 Excel（企业清单、战略取舍、定价、测算、计划、KPI） |
| `exports/实习推送_企业名录与收费流程说明.pptx` | 专项说明：上海 AI/具身智能名录、收费项目、推送成本与流程 |

## 重新生成

```bash
python3 scripts/build_internship_feasibility_ppt.py
python3 scripts/build_internship_feasibility_excel.py
python3 scripts/build_internship_process_ppt.py
```

依赖：`python-pptx`、`openpyxl`（通常已预装，若缺失可 `pip3 install python-pptx openpyxl`）。

## 方案要点

- **人群**：高中生（申请向）· 大学生（履历品牌）· 毕业生（就业转化）
- **产品**：只做高价值商业实习（上海 AI / 具身智能 / 大厂通道）+ 就业转化包
- **已取消**：199 / 699 / 1,680「开证明」低价档（费用低、成本与风险高）
- **定价**：商业标准 5,480；进阶通道 11,800；就业包 9,800；推荐信仅项目内加购 +1,500
- **试点**：建议 6 个月验证，基准情景约 200 人、营收约 165 万、毛利率约 38%
