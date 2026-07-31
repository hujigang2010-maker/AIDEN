# TicNote Studio

面向 TicNote 录音笔 / 云端知识库的本地网页应用：实时语音转写、音频转文字、网页剪藏、知识库时间线、智能体整理，以及 TicNote API 同步与自动转写。

## 能做什么

1. **知识库时间线**：每天采集的录音、网页、云端文件按日期沉淀在左侧。
2. **连接智能体**：纪要官 / 结构官 / 顿悟官 / 研究员 / 播客制片 / 翻译官；研究员可对接 TicNote Deep Research。
3. **实时提问与翻译**：对当前文稿或整库提问；支持本地粗译或 AI / TicNote 翻译。
4. **网页内容转化**：粘贴 URL，本地代理抓取正文并入库。
5. **音频转文字** + **导出全部**：上传音频后一键转写；一键导出总结/导图/顿悟/深研/播客 Markdown、SVG 与原文。

## 关于 TicNote API Key

本工具**不能代替你生成官方 TicNote Key**。请到官网自行创建后粘贴进右侧栏：

1. 打开 [https://www.ticnote.cn](https://www.ticnote.cn) 并登录  
2. 头像 → 个人中心 → **TicNote Key** / API 密钥  
3. 新建并复制（形如 `tncn_sk_…` 或 `tnovs_sk_…`）  
4. 在本应用右侧粘贴 → 点「连接」→ 选择项目或填入 chatId/projectId →「同步并自动转写」

你给的项目链接中的 `chatId=2078873748446011394` 可直接填入「projectId / chatId」输入框。

> AppKey 与 Token 只保存在本机浏览器 `localStorage`，不会上传到第三方（除你配置的 TicNote / AI 端点外）。

## 运行方式

```bash
cd ticnote-studio
python3 server.py
```

浏览器打开：<http://127.0.0.1:8765>

- `server.py`：静态页面 + **TicNote API 代理**（绕过浏览器 CORS）+ 网页抓取 + 可选 AI 转发  
- 纯离线采集（麦克风听写 / 本地模板）：也可直接用浏览器打开 `index.html`（TicNote 同步仍需本地服务）

可选环境变量：`HOST`、`PORT`（默认 `127.0.0.1:8765`）。

## 推荐流程

1. 启动本地服务并打开页面  
2. （可选）右侧填入 TicNote AppKey → 连接 → 同步项目  
3. 或：麦克风实时转写 / 上传音频点「🎧 音频转文字」/ 抓取网页  
4. 下方切换模板整理；需要时点「📦 导出全部」  
5. 在「智能体」「提问/翻译」里继续加工知识库  

## 文件

| 文件 | 说明 |
|------|------|
| `index.html` | 界面 |
| `app.js` | 前端逻辑（本地模板、知识库、同步 UI） |
| `server.py` | 本地服务与 TicNote 代理 |
| `common.py` | AppKey → API Base URL 路由（来自官方 Skill） |
| `TICNOTE-APIKEY.md` | 官方 Key 获取说明摘录 |

## 说明

- 实时听写依赖浏览器 Web Speech API，建议 Chrome / Edge。  
- 上传音频文件：配置了 OpenAI 兼容端点则走云端转写；否则「播放并实时听写」。  
- TicNote 云端转写/深研受账号套餐与 VIP 权限影响，若返回 `no_rights` 需在官网侧开通。  
- 本地模板为零成本离线整理；AI 与 TicNote 为可选增强。
