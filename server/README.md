# 企业信息代理（电话号码补充）

> 用于把【企业明细】中的 `电话` 列从空白补充为真实注册电话。
> 调用 **天眼查** 或 **企查查** 开放 API。

## 为什么需要代理？

浏览器**不能直接**调用天眼查/企查查 API：
1. **CORS 限制** — 两家都没开放浏览器直连
2. **Key 暴露** — API Key 写在前端就等于公开

本目录的 `qcc-proxy.js` 是一个**零依赖**的 Node.js 中转服务：
浏览器 → `http://localhost:3001/api/enrich?company=XXX` → 代理带 Key 查天眼查 → 返回 JSON。

## 启动 3 步

### 1. 申请开放平台 API（任选其一）

| 平台 | 申请地址 | 说明 |
| --- | --- | --- |
| 天眼查开放平台（**推荐**） | <https://open.tianyancha.com> | 接口最简洁,Bearer Token |
| 企查查开放平台 | <https://openapi.qcc.com> | 需要 AppKey + SecretKey |

### 2. 设置环境变量

天眼查：
```bash
export TYC_TOKEN="你申请到的 Bearer Token"
```

企查查（备选）：
```bash
export QCC_KEY="你的 AppKey"
export QCC_SECRET="你的 SecretKey"
```

> 两者都设置时,代理优先用天眼查,失败再回落到企查查。

### 3. 启动代理

```bash
node server/qcc-proxy.js
```

输出：
```
╭─────────────────────────────────────────────────────────╮
│  企业信息代理已启动 → http://localhost:3001
├─────────────────────────────────────────────────────────┤
│  天眼查 (TYC_TOKEN):   ✅ 已配置
│  企查查 (QCC_KEY):     ❌ 未设置
╰─────────────────────────────────────────────────────────╯
```

### 4. 在网页里配置

打开 `index.html`,右上角【**⚙ API 配置**】 → 代理地址填 `http://localhost:3001` → 保存。

回到【**企业明细**】 Tab,点【**🔄 同步电话/企业信息**】即可批量补全可见行的电话、注册资本、法人、统一信用代码等。

## API 路由

| 路由 | 说明 |
| --- | --- |
| `GET /health` | 健康检查 + 配置状态 |
| `GET /api/enrich?company=XX` | 按企业名查询（推荐） |
| `GET /api/search?company=XX` | 同上 |

返回示例：
```json
{
  "source": "tianyancha",
  "name": "上海合生创展商业地产管理有限公司",
  "phone": "021-65010888",
  "legalPerson": "张三",
  "regCapital": "5000 万元",
  "regTime": "2008-09-12",
  "regNumber": "91310110xxxxxxxx",
  "city": "上海",
  "businessScope": "商业地产经营管理 …",
  "url": null
}
```

## 部署到生产

- **本地长期跑**：建议用 `pm2 start qcc-proxy.js --name qcc` 守护
- **公网部署**：放到 Vercel / Cloudflare Workers / 阿里云函数计算 / 自有服务器,记得：
  1. 把环境变量改为平台密钥管理
  2. 加访问控制（IP 白名单 / API Key 校验）—— 否则别人可以白嫖你的天眼查额度
- **节流**：天眼查接口有调用配额。代理已经做了 5 秒超时,前端默认 1 秒一条逐次发起。

## 免责声明

数据所有权归原平台。商业用途请先取得 天眼查 / 企查查 的商业授权。
