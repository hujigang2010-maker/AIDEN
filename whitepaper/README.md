# 2026 谷歌上海开发者大会白皮书

复旦大学住房政策研究中心 · 研究文稿第三号（FDU-HPRC-WP-2026-03）

以 2026 年 8 月 12—13 日上海世博中心举行的 Google I/O Connect China 为观察对象，整理公开新闻与官方材料，讨论智能体时代的中国开发者生态、出海接口及其城市空间含义。与研究文稿第二号《WAIC2026 人工智能产业空间白皮书》形成同城双截面对照。

## 下载版（推荐）

见目录 `下载版本/`：

- Word 白皮书（.docx）
- PDF 白皮书（.pdf）
- Markdown 文字稿（.md）
- 简报 PPT（.pptx）

## 源文件

- `2026谷歌上海开发者大会白皮书.md` / `.docx`
- `2026谷歌上海开发者大会白皮书-简报.pptx`
- `assets/logo_fudan_hprc.png`：中心中英文组合标识
- `assets/charts/`：12 幅图表
- `data/`：历年举办地、I/O 2026 指标、Gemma 4 入围名单快照

## 重新生成

```bash
python3 scripts/build_hprc_logo.py
python3 scripts/build_whitepaper_charts.py
python3 scripts/build_whitepaper_docx.py
python3 scripts/build_whitepaper_ppt.py
python3 scripts/build_whitepaper_pdf.py
```
