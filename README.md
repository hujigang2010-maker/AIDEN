# AIDEN

AIDEN 是一个多任务工作区。`main` 分支存放共享配置；具体应用和文档交付物位于各自的 `cursor/*` 功能分支。

## 本分支

`cursor/dongsheng-fusion-90day-plan-7f33` 生成 **东昇聚变 · 政府事务与上海产业落地 90 天工作设想** 三件套：

| 文件 | 用途 |
| --- | --- |
| `deliverables/东昇聚变_政府事务与上海产业落地_90天工作设想.pptx` | 高管汇报（16 页） |
| `deliverables/东昇聚变_政府事务与上海产业落地_90天工作设想.docx` | 书面工作设想 |
| `deliverables/东昇聚变_政府事务与上海产业落地_90天工作台账.xlsx` | 分周计划、作战图、选址与现场核验台账 |

口径：补位不抢位、先懂装置再对外、重资产看场域与财政。岗位主责是政府事务与产业落地参谋，不是专职拿地。

## 生成

```bash
python3 scripts/build_all.py
python3 scripts/verify_deliverables.py
```

依赖：`python-pptx`、`python-docx`、`openpyxl`。

## 显示语言

本仓库默认使用**简体中文**。首次打开项目时，请安装推荐的中文语言包，并在命令面板中选择 **Configure Display Language → 中文(简体)**，然后重启 Cursor。

更多说明见 [AGENTS.md](./AGENTS.md)。
