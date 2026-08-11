# WAIC2026 人工智能产业空间白皮书

复旦大学住房政策研究中心 · 研究文稿第二号（FDU-HPRC-WP-2026-02）

## 交付物

- `WAIC2026人工智能产业空间白皮书.md`：Markdown 文字稿（含全部图表引用）
- `WAIC2026人工智能产业空间白皮书.docx`：Word 版本（含封面 logo、图表与表格排版）
- `assets/logo_fudan_hprc.png`：中心中英文组合标识
- `assets/charts/`：全套 14 幅统计图表
- `data/`：xsct.ai 大模型榜单与 watcha.cn 观猹产品数据快照（2026-08 抓取）

## 数据来源

1. WAIC2026 全量资源整合总表.xlsx（963 家参展商、175 场论坛、4262 家品牌库等）
2. xsct.ai（XSCT Bench 大模型评测榜单）
3. watcha.cn（观猹 AI 产品社区公开接口）

## 重新生成

```bash
pip install openpyxl python-docx matplotlib pillow
python3 scripts/build_hprc_logo.py          # 生成中心 logo
python3 scripts/build_whitepaper_charts.py  # 生成 14 幅图表
python3 scripts/build_whitepaper_docx.py    # 由 MD 生成 Word 版本
```

图表脚本依赖 Noto Sans CJK 中文字体（`fonts-noto-cjk`）。
