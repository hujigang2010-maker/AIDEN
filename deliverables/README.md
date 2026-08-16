# WorkBuddy 银行引荐判断包

内部判断材料（2026-08-16）：要不要现在引荐中国银行、上海银行，以及下周五（8月21日）如何约见泰隆上海分行杨行长。

## 结论

可以引荐，但不要本周并行约见。先把杨行长这场主线走实；上海银行做第二主线预热，中国银行只做探索。不要用新银行催腾讯。

## 交付文件

| 文件 | 说明 |
|------|------|
| `WorkBuddy银行引荐判断备忘录_杨行长约见_20260816.docx` | 完整判断备忘录 |
| `WorkBuddy银行引荐判断_杨行长约见简报_20260816.pptx` | 9 页决策简报 |
| `WorkBuddy银行引荐_三家银行对照与十天行动表_20260816.xlsx` | 对照表、十天节奏、约见准备、口径卡、会后决策树 |

## 重新生成

```bash
python3 scripts/build_workbuddy_intro.py
```

或分别运行：

```bash
python3 scripts/generate_workbuddy_intro_docx.py
python3 scripts/generate_workbuddy_intro_ppt.py
python3 scripts/generate_workbuddy_intro_excel.py
```

依赖：`python-pptx`、`python-docx`、`openpyxl`。
