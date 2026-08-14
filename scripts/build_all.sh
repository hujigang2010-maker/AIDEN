#!/usr/bin/env bash
# 生成住房产业三链融合白皮书的全部交付物
set -euo pipefail
cd /workspace
python3 scripts/build_hprc_logo.py
python3 scripts/build_whitepaper_charts.py
python3 scripts/build_whitepaper_docx.py
python3 scripts/build_whitepaper_ppt.py
mkdir -p whitepaper/下载版本
soffice --headless --convert-to pdf --outdir whitepaper \
  "whitepaper/住房产业三链融合白皮书.docx"
cp "whitepaper/住房产业三链融合白皮书.md" "whitepaper/下载版本/"
cp "whitepaper/住房产业三链融合白皮书.docx" "whitepaper/下载版本/"
cp "whitepaper/住房产业三链融合白皮书.pdf" "whitepaper/下载版本/"
cp "whitepaper/住房产业三链融合白皮书-简报.pptx" "whitepaper/下载版本/"
echo "all deliverables ready"
