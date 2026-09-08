# 胡继刚履历

两页 A4 专家履历（Word + PDF），面向产业招商 / 城市更新 / 存量空间 / 国资平台。

## 生成

```bash
python3 scripts/generate_resume.py
```

产出：

- `output/胡继刚-简历.docx`
- `output/胡继刚-简历.pdf`
- `output/胡继刚-简历.html`（浏览器预览，可再导出 PDF）

校验：

```bash
python3 scripts/test_resume.py
```

需要 `python-docx`、`pymupdf`，以及本机 `google-chrome`（用于 HTML 转 PDF）。中文字体为 Noto Serif SC / Noto Sans SC，位于 `scripts/fonts/`。
