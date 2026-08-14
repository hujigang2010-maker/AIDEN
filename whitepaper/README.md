# WAIC2026 人工智能产业空间白皮书

复旦大学住房政策研究中心 · 研究文稿第二号（FDU-HPRC-WP-2026-02）

## 精排版简报（每页配图，推荐对外输出）

- `下载版本/WAIC2026人工智能产业空间白皮书-精排版.pdf`
- `下载版本/WAIC2026人工智能产业空间白皮书-精排版.pptx`

22 页，16:9，每页一张摄影图；图表页为顶部横图 + 数据图。

重新生成：

```bash
python3 scripts/build_designed_deck.py
```

见目录 `下载版本/`：

- Word 白皮书（.docx）
- PDF 白皮书（.pdf）
- Markdown 文字稿（.md）
- 简报 PPT（.pptx）

## 源文件

- `WAIC2026人工智能产业空间白皮书.md` / `.docx`
- `WAIC2026人工智能产业空间白皮书-简报.pptx`
- `assets/logo_fudan_hprc.png`：中心中英文组合标识
- `assets/charts/`：22 幅统计图表
- `data/`：xsct.ai、watcha.cn、杨浦园区样本数据快照

## 重新生成

```bash
python3 scripts/build_hprc_logo.py
python3 scripts/build_whitepaper_charts.py
python3 scripts/build_whitepaper_docx.py
python3 scripts/build_whitepaper_ppt.py
```
