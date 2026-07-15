#!/usr/bin/env bash
set -euo pipefail

CURSOR_SERVER_DATA="${CURSOR_SERVER_DATA:-$HOME/.cursor-server/data}"
LANGUAGE_PACK_ID="ms-ceintl.vscode-language-pack-zh-hans"
LOCALE="zh-cn"

mkdir -p "$CURSOR_SERVER_DATA/User" "$CURSOR_SERVER_DATA/Machine"

write_json() {
  local file="$1"
  printf '{\n  "locale": "%s"\n}\n' "$LOCALE" > "$file"
}

write_json "$CURSOR_SERVER_DATA/User/argv.json"
write_json "$CURSOR_SERVER_DATA/Machine/argv.json"
printf '{\n  "locale": "%s"\n}\n' "$LOCALE" > "$CURSOR_SERVER_DATA/User/settings.json"

if [ -d "$HOME/.cursor-server/extensions" ]; then
  ext_dir=$(ls -d "$HOME/.cursor-server/extensions/${LANGUAGE_PACK_ID}-"* 2>/dev/null | head -1 || true)
  if [ -n "${ext_dir:-}" ]; then
    echo "已检测到中文语言包: $ext_dir"
  else
    echo "警告: 未检测到中文语言包，请在 Cursor 扩展市场安装 $LANGUAGE_PACK_ID"
  fi
fi

echo "已将显示语言配置为: $LOCALE"
echo "请在本机 Cursor 中执行: Configure Display Language -> 中文(简体) -> 重启"
