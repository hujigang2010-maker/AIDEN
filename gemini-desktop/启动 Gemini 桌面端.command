#!/bin/bash
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
if [ -x "$HERE/bin/gemini-desktop" ]; then
  exec "$HERE/bin/gemini-desktop" "$@"
fi
if [ -f "$HERE/gemini_desktop.py" ]; then
  exec python3 "$HERE/gemini_desktop.py" "$@"
fi
if [ -f "$HERE/gemini-desktop/gemini_desktop.py" ]; then
  exec python3 "$HERE/gemini-desktop/gemini_desktop.py" "$@"
fi
echo "找不到 Gemini 桌面端主程序。请先运行：python3 gemini-desktop/gemini_desktop.py --install"
read -r _
exit 1
