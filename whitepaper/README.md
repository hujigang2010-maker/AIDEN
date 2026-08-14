# 2026 Google 开发者大会观察白皮书

复旦大学住房政策研究中心 · 研究文稿第三号（FDU-HPRC-WP-2026-03）

## 下载版（推荐）

见目录 `下载版本/`：

- Word 白皮书（.docx）
- PDF 白皮书（.pdf，若环境可转换）
- Markdown 文字稿（.md）
- 简报 PPT（.pptx）

## 源文件

- `2026Google开发者大会观察白皮书.md` / `.docx`
- `2026Google开发者大会观察白皮书-简报.pptx`
- `assets/logo_fudan_hprc.png`：中心中英文组合标识
- `assets/charts/`：14 幅说明图表
- `data/`：大会基本信息与 Gemma 4 黑客松入围名单快照

## 重新生成

```bash
python3 scripts/build_hprc_logo.py
python3 scripts/build_whitepaper_charts.py
python3 scripts/build_whitepaper_docx.py
python3 scripts/build_whitepaper_ppt.py
```

正文依据 2026 年 8 月 12—13 日上海世博中心 Google I/O Connect China 的公开新闻与官方博客整理，并讨论其对人才住房、会展空间与产业空间的含义。
