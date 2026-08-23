# Gemini 桌面端（登录修复）

官方或第三方 Gemini 桌面包装器经常出现这种症状：

- 应用已经装好
- 网络能打开 Google
- 浏览器里授权登录也成功
- 回到桌面端却始终进不了已登录状态

根因通常不是账号本身，而是**登录发生在 iframe / 外部浏览器里，会话写不回应用窗口**。旧实现还会剥离 `X-Frame-Options`，把 `https://gemini.google.com/app` 嵌进页面；Google 登录页拒绝这种第三方框架，cookie 也变成第三方上下文，授权成功也没用。

本分支改成：

1. **顶层窗口**直接打开 Gemini，不再 iframe
2. 会话固定在 `persist:gemini` 分区，刷新后仍保持登录
3. 登录弹窗留在**同一分区**，禁止把 OAuth 丢给系统浏览器
4. 使用 Chrome 148 User-Agent，去掉 Electron 指纹，避免 Google 直接 403

## 运行

```bash
npm install
npm test
npm run dev
```

Linux 桌面快捷方式：

```bash
chmod +x scripts/*.sh
./scripts/install-desktop-shortcut.sh
```

之后可从桌面图标或 `scripts/launch-gemini.sh` 启动。

## 诊断

```bash
./scripts/diagnose-gemini-login.sh
```

若你用的是官方 macOS `Gemini.app`，浏览器授权成功但 App 仍未登录，多半是应用内 `oauth2.googleapis.com` 换票没走系统代理。可改用本客户端，或给官方 App 打开 Clash / 系统代理的 TUN 模式。
