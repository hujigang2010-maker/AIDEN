# 前世叙事的功能分析框架（PLN-F）

本分支沉淀一套 **AI 可用** 的前世叙事功能分析框架，回答：不同文化里的前世故事分别在承担什么功能。

## 四维功能

1. **心理补偿**（F1）
2. **关系解释**（F2）
3. **身份建构**（F3）
4. **道德约束**（F4）

## 交付物

| 路径 | 说明 |
|------|------|
| `output/前世叙事的功能分析框架.pptx` | 汇报用 PPT |
| `framework/PLN-F.md` | 完整框架说明（可作 RAG / 提示词上下文） |
| `framework/pln_f.schema.json` | 结构化标注 Schema |
| `framework/example_annotation.json` | 标注示例 |
| `scripts/generate_past_life_framework_ppt.py` | PPT 生成脚本 |

## 生成 PPT

```bash
python3 scripts/generate_past_life_framework_ppt.py
```

## 使用姿态

功能分析描述叙事「完成了什么社会—心理工作」，**不做真伪裁决**。
