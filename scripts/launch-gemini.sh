#!/usr/bin/env bash
# 启动修复后的 Gemini 桌面端。
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
ELECTRON_BIN="${ROOT}/node_modules/.bin/electron"

if [[ ! -x "${ELECTRON_BIN}" ]]; then
  echo "正在安装依赖…"
  (cd "${ROOT}" && npm install)
fi

if [[ ! -f "${ROOT}/dist/main.js" ]]; then
  (cd "${ROOT}" && npm run build)
fi

cd "${ROOT}"
exec "${ELECTRON_BIN}" "${ROOT}"
