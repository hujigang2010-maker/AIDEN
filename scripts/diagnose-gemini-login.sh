#!/usr/bin/env bash
# 诊断 Gemini 桌面登录：区分「浏览器授权成功但应用没会话」和网络故障。
set -euo pipefail

echo "== Gemini 登录诊断 =="
echo "时间: $(date -Is)"
echo

check_url() {
  local name="$1"
  local url="$2"
  local code
  code="$(curl -sS -o /dev/null -w '%{http_code}' --max-time 15 -A 'Mozilla/5.0' "${url}" || echo fail)"
  local reach="不可达"
  if [[ "${code}" =~ ^[2345][0-9][0-9]$ ]]; then
    reach="可达"
  fi
  printf '%-28s %s (%s)\n' "${name}" "${code}" "${reach}"
}

echo "-- 网络 --"
check_url "gemini.google.com" "https://gemini.google.com/app"
check_url "accounts.google.com" "https://accounts.google.com/"
check_url "oauth2.googleapis.com" "https://oauth2.googleapis.com/"
echo

echo "-- 代理环境变量 --"
printf 'HTTP_PROXY=%s\n' "${HTTP_PROXY:-<空>}"
printf 'HTTPS_PROXY=%s\n' "${HTTPS_PROXY:-<空>}"
printf 'ALL_PROXY=%s\n' "${ALL_PROXY:-<空>}"
echo

echo "-- 本机包装器 --"
if [[ -x "$(dirname "$0")/../node_modules/.bin/electron" ]]; then
  echo "已安装本仓库 Electron 桌面端"
else
  echo "尚未 npm install"
fi

if [[ -d "/Applications/Gemini.app" ]]; then
  echo "检测到官方 macOS Gemini.app"
  echo "若浏览器授权成功但 App 仍未登录，多半是应用内 token 换票没走系统代理。"
  echo "请改用本仓库 scripts/launch-gemini.sh，或为官方 App 打开 Clash TUN。"
fi

echo
echo "结论提示："
echo "1. 上面三个地址都能返回 2xx/3xx，说明网络和 Google 授权入口可用。"
echo "2. 旧包装器把 Gemini 放进 iframe，授权 cookie 写不回应用，会表现为一直无法登录。"
echo "3. 请使用本仓库的顶层窗口客户端：npm run dev 或 scripts/launch-gemini.sh"
