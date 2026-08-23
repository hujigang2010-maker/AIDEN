# Gemini 桌面端登录修复

官方原生 Gemini 桌面应用只提供 **macOS（Apple Silicon，macOS 15+）** 和 **Windows**。
这台 Linux 云桌面里常见的 Chrome PWA / Electron 套壳会出现：

- 网络和 Google 授权都已通过
- 应用窗口却仍停在「Sign in」，或跳到 `gemini.google/app` 的 **404**

根因是登录回跳写错了主机名（`gemini.google`，缺少 `.com`），或登录发生在另一个
Chrome 窗口，会话没有回到桌面应用。

本目录提供一个用 **本机 Google Chrome `--app` 模式** 打开的桌面端：

1. 强制走 `accounts.google.com/ServiceLogin?continue=https://gemini.google.com/app`
2. 用独立配置目录保存 Cookie，登录一次后下次直接进入 Gemini
3. 用扩展把错误的 `gemini.google/...` 纠正到 `gemini.google.com/...`

## 安装

在仓库根目录执行：

```bash
python3 gemini-desktop/gemini_desktop.py --install
```

或：

```bash
bash gemini-desktop/install.sh
```

安装后可从应用菜单或 Dock 打开 **Gemini 桌面端**。

## 使用

```bash
~/.local/bin/gemini-desktop              # 启动
~/.local/bin/gemini-desktop-fix-login    # 授权成功但仍未登录时强制回跳
python3 gemini-desktop/gemini_desktop.py --reset-profile   # 清空损坏会话后重登
```

首次打开会进入 Google 登录页。请在**同一个应用窗口**里完成登录，不要改用旁边的
普通 Chrome 窗口。登录成功后会自动回到 `https://gemini.google.com/app`。

## 自检

```bash
python3 gemini-desktop/tests/test_gemini_desktop.py
```
