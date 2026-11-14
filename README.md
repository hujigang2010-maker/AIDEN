# 2026银发科技峰会合作执行需求说明书（PRD）

本分支把 2026年11月14日上午沟通形成的八项合作原则，整理成一份可直接确认的需求说明书。

活动是 2026上海国际智慧康养产业博览会中的分论坛：AI重塑养老未来——2026银发科技与具身智能产业创新峰会，大会日期为 2026年11月28日。

## 文件

- `deliverables/2026银发科技峰会_合作执行需求说明书PRD.docx`：确认稿正文
- `deliverables/2026银发科技峰会_合作执行需求说明书PRD.pdf`：同一份文件的阅读版

文档固定的是分工和责任边界。嘉宾到场、赞助金额和活动效果都不写成保证。复旦大学及复旦大学住房政策研究中心的正式署名，仍以实际授权为准。

## 重新生成

```bash
python3 scripts/build_prd.py
```

脚本使用 `python-docx`。本机装有 LibreOffice 时，可用下面的命令导出 PDF：

```bash
soffice --headless --convert-to pdf --outdir deliverables deliverables/2026银发科技峰会_合作执行需求说明书PRD.docx
```
