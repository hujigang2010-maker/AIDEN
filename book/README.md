# 《以新为鉴——存量时代的城市生存法则》书稿

以新加坡为镜，按“供中国借鉴”的模式整理新加坡在就业、经济周期、住房、养老医疗、人口和 AI 转型方面的公开资料；结构参照《以日为鉴：衰退时代生存指南》（五篇十五章），扩展为“序章 + 六篇十八章 + 后记 + 结语 + 三个附录”。

## 目录结构

- `src/`：分篇 Markdown 源稿（按文件名顺序合并）
- `assets/charts/`：10 张数据图表（由 `scripts/build_sg_charts.py` 生成）
- `../deliverables/以新为鉴_存量时代的城市生存法则.md`：合并后的完整书稿
- `../deliverables/以新为鉴_存量时代的城市生存法则.docx`：Word 版书稿（含封面、目录、图表、页码）

## 每章固定结构

数据看板 → 正文（新加坡做法）→ 民生切面（合成家庭画像）→ 周期视角（康波/经济周期）→ 本章反思 → 供中国借鉴（政策层 / 城市层 / 个人家庭层）

## 重新生成

```bash
pip install python-docx matplotlib
python3 scripts/build_sg_charts.py
python3 scripts/build_sg_book.py
```

数据检索时间为 2026 年 10 月，来源见书稿附录 C。
