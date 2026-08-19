# WorkBuddy 文件修复

腾讯 WorkBuddy 生成的 PPT / Word / Excel，在 PowerPoint、WPS 里常会提示**文件损坏、无法打开**。根因是底层 python-pptx 把 16:9 幻灯片仍标记成 `type="screen4x3"`，尺寸与预设不一致。

本分支已按 Office 兼容规则重新生成并修复全部交付物。

## 已修复文件

| 文件 | 说明 |
|------|------|
| `泰隆银行x腾讯云WorkBuddy_战略合作实施方案_V3.0.pptx` | V3.0 决策汇报 PPT（25 页） |
| `泰隆银行x腾讯云WorkBuddy_落地执行作战表_V3.0.xlsx` | V3.0 作战表 |
| `泰隆银行上海分行与腾讯WorkBuddy合作推进情况汇报_20260803.pptx` | 行长汇报 5 页 PPT |
| `泰隆银行上海分行与腾讯WorkBuddy合作推进情况汇报_20260803.docx` | 行长汇报 Word |
| `WorkBuddy银行引荐判断_杨行长约见简报_20260816.pptx` | 杨行长约见 9 页简报 |
| `WorkBuddy银行引荐判断备忘录_杨行长约见_20260816.docx` | 引荐判断备忘录 |
| `WorkBuddy银行引荐_三家银行对照与十天行动表_20260816.xlsx` | 对照与行动表 |

## 自己修复一份 WorkBuddy 导出文件

把打不开的 pptx / docx / xlsx 放到本仓库后执行：

```bash
python3 scripts/repair_tencent_workbuddy_office.py 你的文件.pptx
```

会在同目录生成 `你的文件_已修复.pptx`。加 `--inplace` 则覆盖原文件。

重新生成全部交付物：

```bash
python3 scripts/rebuild_workbuddy_files.py
```
