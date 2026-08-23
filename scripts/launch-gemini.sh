#!/usr/bin/env bash
# 启动 Gemini 桌面窗口。
# 默认走系统 Chrome 的 --app 模式：顶层窗口 + 复用已有 Google 登录态，
# 避免 Electron iframe / 独立分区导致「授权成功却登不进去」。
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
GEMINI_URL="${GEMINI_URL:-https://gemini.google.com/app}"
CHROME_BIN="${CHROME_BIN:-}"

if [[ -z "${CHROME_BIN}" ]]; then
  for candidate in google-chrome-stable google-chrome chromium-browser chromium; do
    if command -v "${candidate}" >/dev/null 2>&1; then
      CHROME_BIN="$(command -v "${candidate}")"
      break
    fi
  done
fi

if [[ -n "${CHROME_BIN}" && "${GEMINI_LAUNCHER:-chrome}" != "electron" ]]; then
  PROFILE_DIR="${GEMINI_CHROME_PROFILE:-${HOME}/.config/gemini-desktop-chrome}"
  mkdir -p "${PROFILE_DIR}"
  exec "${CHROME_BIN}" \
    --app="${GEMINI_URL}" \
    --user-data-dir="${PROFILE_DIR}" \
    --no-first-run \
    --no-default-browser-check \
    --no-sandbox \
    --test-type \
    --disable-dev-shm-usage \
    --use-gl=angle \
    --use-angle=swiftshader-webgl \
    --password-store=basic \
    --class=gemini-desktop \
    --window-size=1280,840
fi

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
