# 住房产业“三链融合”白皮书

复旦大学住房政策研究中心 · 研究文稿第三号（FDU-HPRC-WP-2026-03）

面向 2030 年，讨论住房如何从“土地 + 钢筋混凝土”重构为**供应链、硬件与软件**一体化产业。

## 下载版（推荐）

见目录 `下载版本/`：

- Word 白皮书（.docx）
- PDF 白皮书（.pdf）
- Markdown 文字稿（.md）
- 简报 PPT（.pptx）

## 源文件

- `住房产业三链融合白皮书.md` / `.docx` / `.pdf`
- `住房产业三链融合白皮书-简报.pptx`
- `assets/logo_fudan_hprc.png`：中心中英文组合标识
- `assets/charts/`：12 幅研究图表

## 重新生成

```bash
python3 scripts/build_hprc_logo.py
python3 scripts/build_whitepaper_charts.py
python3 scripts/build_whitepaper_docx.py
python3 scripts/build_whitepaper_ppt.py
soffice --headless --convert-to pdf --outdir whitepaper whitepaper/住房产业三链融合白皮书.docx
```
