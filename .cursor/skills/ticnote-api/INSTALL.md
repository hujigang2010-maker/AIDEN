# TicNote API Skill 已安装

- 技能目录：`.cursor/skills/ticnote-api/`
- 官方包：https://voice-recorder-cdn.ticnote.com/ticnote-web-prd/2026-03/ticnote-api.zip

## 配置 API Key

1. 打开 https://www.ticnote.cn 并登录
2. 头像 → 个人中心 →「TicNote Key」
3. 新建并复制 AppKey（形如 `tncn_sk_…`）
4. 在对话里把 AppKey 发给助手完成配置（勿提交到 Git）

验证：

```bash
python3 .cursor/skills/ticnote-api/scripts/get_token.py --appkey "你的AppKey"
```
