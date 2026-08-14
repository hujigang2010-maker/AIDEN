# -*- coding: utf-8 -*-
"""将白皮书 Markdown 转为 PDF（嵌入中文字体）。"""
import os
import re
import markdown
from xhtml2pdf import pisa

BASE = "/workspace/whitepaper"
MD = os.path.join(BASE, "2026谷歌上海开发者大会白皮书.md")
OUT = os.path.join(BASE, "2026谷歌上海开发者大会白皮书.pdf")
FONT = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"

raw = open(MD, encoding="utf-8").read()
# 保留 logo 图片，去掉居中 HTML 外壳
raw = raw.replace('<p align="center"><img src="assets/logo_fudan_hprc.png" alt="复旦大学住房政策研究中心" width="520"></p>',
                  f'<p style="text-align:center"><img src="{BASE}/assets/logo_fudan_hprc.png" width="420"/></p>')
# 让相对图片路径可被 xhtml2pdf 找到
html_body = markdown.markdown(raw, extensions=["tables", "nl2br"])
html_body = html_body.replace('src="assets/', f'src="{BASE}/assets/')

html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<style>
@font-face {{
  font-family: "wqy";
  src: url("{FONT}");
}}
@page {{
  size: A4;
  margin: 2.2cm 2cm 2.2cm 2cm;
}}
body {{ font-family: wqy; font-size: 11pt; color: #222; line-height: 1.45; }}
h1 {{ color: #0E4E9B; font-size: 18pt; }}
h2 {{ color: #0E4E9B; font-size: 14pt; }}
h3 {{ color: #0E4E9B; font-size: 12pt; }}
img {{ max-width: 16cm; }}
table {{ border-collapse: collapse; width: 100%; font-size: 9pt; margin: 8px 0; }}
th, td {{ border: 1px solid #C9CFD8; padding: 4px 6px; }}
th {{ background: #DCE6F1; }}
blockquote {{ color: #0E4E9B; border-left: 3px solid #0E4E9B; padding-left: 8px; }}
</style>
</head>
<body>
{html_body}
</body>
</html>
"""

with open(OUT, "wb") as f:
    status = pisa.CreatePDF(html, dest=f, encoding="utf-8")
print("pdf errors:", status.err)
print("saved", OUT, "size", os.path.getsize(OUT))
