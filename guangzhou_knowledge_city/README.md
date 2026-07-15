# 广州知识城“全球自贸365街区”项目 — 提资清单、服务建议书与报价

提供方:复旦大学住房政策研究中心、上海市杨浦区科技企业联合会

## 交付文件(`deliverables/`)

| 文件 | 格式 | 内容 |
| --- | --- | --- |
| 广州知识城全球自贸365街区_服务建议书.docx | Word | 项目概要、服务范围(八大模块)、成果交付、提资清单、服务报价、团队介绍 |
| 广州知识城全球自贸365街区_提资清单与报价.xlsx | Excel | 三个工作表:项目概要 / 提资清单(12项) / 报价单(四阶段合计118万元) |
| 广州知识城全球自贸365街区_项目概要.pptx | PPT | 8页:封面、概念、产业方向、下阶段重点、成果交付、提资清单摘要、服务报价、结尾 |

## 重新生成

```bash
pip install python-pptx python-docx openpyxl
cd scripts
python3 generate_word.py
python3 generate_excel.py
python3 generate_ppt.py
```

文案统一维护在 `scripts/content.py`,修改后重新运行三个生成脚本即可保持三种载体内容一致。
