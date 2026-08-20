# AIDEN × 西电产学研课题介绍

给西安电子科技大学老师与经过初筛的同学用的课题发布材料。十条题目都来自 AIDEN 仓库里**已经开工**的真实任务，供下周一正式发出后选题组队。

## 交付物

| 文件 | 用途 |
|------|------|
| `exports/AIDEN_西电产学研课题介绍.docx` | 完整课题介绍，老师留底、同学细读 |
| `exports/AIDEN_西电产学研课题介绍.pptx` | 宣讲 15 分钟 |
| `exports/AIDEN_西电课题选题与编组表.xlsx` | 学生志愿、老师编组 |
| `exports/AIDEN_西电课题_发群一页稿.md` | 可直接发群的一页稿（Markdown） |
| `exports/AIDEN_西电课题_发群一页稿.txt` | 同上，纯文本，适合微信粘贴 |

## 重新生成

```bash
python3 -m pip install python-pptx python-docx openpyxl
python3 scripts/build_all.py
python3 scripts/verify.py
```

依赖：`python-pptx`、`python-docx`、`openpyxl`。

## 口径（不要改偏）

- 只写已开工课题，不编「假实习」。
- 进组是做事，不开空证明。
- 远程为主，上海可短期驻场。
- 编组由西电老师确认；AIDEN 侧不在学生群单独招生。
- 计划发布日：2026-08-24（周一）。

## 十条课题

T01 MP Typer 排版引擎 · T02 复旦链接小程序 · T03 楼宇数据可视化 · T04 活动报名 MCP · T05 TicNote 知识工作室 · T06 企业 AI 文档工程 · T07 具身智能与产业空间 · T08 住房学区数据工具 · T09 AIDEN 多智能体工作流（主线） · T10 产业白皮书工作坊。
