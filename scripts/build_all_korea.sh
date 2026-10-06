#!/usr/bin/env bash
# 一键生成《以韩为鉴》全部交付物：配图、Word 书稿、全文 Markdown、数据底稿 Excel
set -euo pipefail
cd "$(dirname "$0")/.."
python3 scripts/build_korea_charts.py
python3 scripts/build_korea_book.py
python3 scripts/build_korea_data_excel.py
ls -lh output/
