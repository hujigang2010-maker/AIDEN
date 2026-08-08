# 接手失败项目的通用框架

把一次「救火」沉淀成可迁移的作战手册：**先全量摸底 → 再按自己的逻辑重构 → 完成后及时转向**。

适用于 **AI 项目管理**、**产品接手**、技术债抢救、失败业务盘活等场景。

## 交付物

| 文件 | 说明 |
|------|------|
| `deliverables/接手失败项目的通用框架.pptx` | 12 页方法论汇报 PPT |
| `deliverables/接手失败项目的通用框架_操作手册.docx` | 可归档/培训的操作手册 |
| `deliverables/接手失败项目的通用框架_执行清单.xlsx` | 摸底 / 杀留重写 / 完成转向可填写清单 |

## 框架三步

1. **全量摸底**：只读优先，输出真相图（可复用 / 需改造 / 冻结 / 待查）。
2. **按自己的逻辑重构**：重定成功定义，用杀/留/重写压缩到一条主路径。
3. **完成后及时转向**：提前写下完成标准，到期离开清理态，沉淀 playbook。

## 重新生成

```bash
pip install python-pptx python-docx openpyxl
python3 scripts/build_all_takeover_framework.py
```

也可单独运行：

```bash
python3 scripts/build_takeover_framework_ppt.py
python3 scripts/build_takeover_framework_docx.py
python3 scripts/build_takeover_framework_excel.py
```
