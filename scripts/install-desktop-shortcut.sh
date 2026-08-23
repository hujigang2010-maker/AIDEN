#!/usr/bin/env bash
# 在当前 Linux 桌面安装 Gemini 快捷方式。
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
LAUNCHER="${ROOT}/scripts/launch-gemini.sh"
ICON_SRC="/usr/share/icons/WhiteSur/apps/scalable/gemini.svg"
ICON_DST="${HOME}/.local/share/icons/gemini-desktop.svg"
APP_DIR="${HOME}/.local/share/applications"
DESKTOP_DIR="${HOME}/Desktop"
ENTRY="${APP_DIR}/gemini-desktop.desktop"

chmod +x "${LAUNCHER}"
mkdir -p "${APP_DIR}" "${HOME}/.local/share/icons" "${DESKTOP_DIR}"
if [[ -f "${ICON_SRC}" ]]; then
  cp "${ICON_SRC}" "${ICON_DST}"
fi

ICON_LINE="${ICON_DST}"
if [[ ! -f "${ICON_LINE}" ]]; then
  ICON_LINE="applications-internet"
fi

cat > "${ENTRY}" <<EOF
[Desktop Entry]
Type=Application
Version=1.0
Name=Gemini
Name[zh_CN]=Gemini 桌面端
Comment=Google Gemini 桌面客户端（顶层窗口登录）
Comment[zh_CN]=Google Gemini 桌面客户端（顶层窗口登录）
Exec=${LAUNCHER}
Icon=${ICON_LINE}
Terminal=false
Categories=Network;Office;
StartupNotify=true
StartupWMClass=gemini-desktop
EOF

chmod +x "${ENTRY}"
cp "${ENTRY}" "${DESKTOP_DIR}/gemini-desktop.desktop"
chmod +x "${DESKTOP_DIR}/gemini-desktop.desktop"

if command -v update-desktop-database >/dev/null 2>&1; then
  update-desktop-database "${APP_DIR}" || true
fi

echo "已安装桌面快捷方式："
echo "  ${ENTRY}"
echo "  ${DESKTOP_DIR}/gemini-desktop.desktop"
