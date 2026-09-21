# AIDEN · 胡继刚简历（优化版 0921 修订）

本分支按《胡继刚-简历-优化版0921.pdf》与 2026-09-21 修改建议，生成两页完整版与一页精华版。

## 生成

```bash
python3 scripts/generate_resume.py
python3 scripts/test_resume.py
```

依赖：`python-docx`、本机 `google-chrome`（无头打印 PDF）。校验可选 `pymupdf`。

## 交付物

- `output/胡继刚-简历-优化版.pdf` / `.docx` / `.html`
- `output/胡继刚-简历-优化版0921.pdf`（与来稿文件名对齐的副本）
- `output/胡继刚-简历-一页精华.pdf` / `.html`
- `output/简历优化说明.md`
