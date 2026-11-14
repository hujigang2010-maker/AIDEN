# 合作原则

来稿原文的排版稿。正文不增字、不减字。

## 文件

- `deliverables/合作原则.docx`
- `deliverables/合作原则.pdf`

## 重新生成

```bash
python3 scripts/build_prd.py
```

脚本使用 `python-docx`。本机装有 LibreOffice 时，可用下面的命令导出 PDF：

```bash
soffice --headless --convert-to pdf --outdir deliverables deliverables/合作原则.docx
```
