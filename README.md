# Gemini 桌面端（登录修复）

官方或第三方 Gemini 桌面包装器经常出现这种症状：

- 应用已经装好
- 网络能打开 Google
- 浏览器里授权登录也成功
- 回到桌面端却始终进不了已登录状态

根因通常不是账号本身，而是**登录发生在 iframe / 外部浏览器里，会话写不回应用窗口**。旧实现还会剥离 `X-Frame-Options`，把 `https://gemini.google.com/app` 嵌进页面；Google 登录页拒绝这种第三方框架，cookie 也变成第三方上下文，授权成功也没用。

本分支改成：

1. **默认用系统 Chrome `--app` 顶层窗口**打开 Gemini，登录发生在真实 Chrome 里
2. 独立应用配置目录，不把授权丢到外部标签后再丢失会话
3. Electron 备选路径同样改为顶层窗口 + `persist:gemini` + Chrome User-Agent
4. Linux 云桌面补齐 `--disable-dev-shm-usage` 与 SwiftShader，避免空白窗口

## 运行

```bash
npm install
npm test
./scripts/launch-gemini.sh
```

强制走 Electron：

```bash
GEMINI_LAUNCHER=electron npm run dev
```

Linux 桌面快捷方式：

```bash
chmod +x scripts/*.sh
./scripts/install-desktop-shortcut.sh
```

## 诊断

```bash
./scripts/diagnose-gemini-login.sh
```

若你用的是官方 macOS `Gemini.app`，浏览器授权成功但 App 仍未登录，多半是应用内 `oauth2.googleapis.com` 换票没走系统代理。可改用本客户端，或给官方 App 打开 Clash / 系统代理的 TUN 模式。
