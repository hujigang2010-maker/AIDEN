# 杨浦五角场 5km 商办竞品地图

基于**高德地图 JS API** 的商办竞品调研可视化网页 — 顶部总览看板、左侧交互地图（支持**卫星图**）、右侧【项目概览】/【企业明细】双 Tab，地图与表格全双向联动。可通过 **天眼查 / 企查查** 一键补充全部 171 家企业的电话、法人、注册资本、统一信用代码。

> **基准点：** 上海市五角场创新创业学院（杨浦区国权路18号）
> **半径：** 5 km · 分层：1 km / 1–2 km / 2–3 km / 3–5 km
> **数据：** 39 个项目 · **171** 家企业

---

## 功能一览

| 模块 | 说明 |
| --- | --- |
| 总览看板 | 项目总数 + 4 圈层项目数,点击直接做距离筛选 |
| 交互地图 | 基准点呼吸光晕 + 4 圈层 + 业态色彩 marker + 企业数 badge |
| 🛰️ 标准/卫星切换 | 地图工具栏一键切换 标准底图 / 卫星 + 路网 |
| Tab：项目概览 | 序号 / 项目 / 业态 / 距离 / 🚗 驾车 / 🚇 公交,与地图双向联动 |
| Tab：企业明细 | 12 列：项目+企业 / 地址 / 距离 / 体量 / 租金 / 免租期 / 剩余租期 / 行业 / 📞 电话 / 物业费 / 性质 |
| 📞 电话录入 | 表格中点击电话格手动录入,自动持久化到 localStorage |
| 🔄 批量补全 | 通过 天眼查 API 一次性补全所有 171 家企业的工商信息 |
| ⏳ 剩余租期 | 由 `leaseStart` + `leaseTerm` 自动计算,3 色 pill |
| 搜索 | 项目名 / 企业名 / 地址 / 行业 / 标签 / 性质 / 电话 全部命中 |
| CSV 导出 | 完整 25 列（含 法人/注册资本/统一信用代码/注册时间/登记状态 等） |
| 响应式 | 1280px 以下自动切换上下布局 |

---

## 快速开始

```bash
# 静态页（必备 - 直接看地图 + 表格）
python -m http.server 8000   # 浏览器访问 http://localhost:8000
# 或直接双击 index.html
```

## 📞 一键补全所有 171 家企业的工商信息

### ⚠️ 先决条件：必须在**本地笔记本 / 办公网络**运行

> **天眼查 WAF 默认封禁所有云服务器 IP**（AWS / 阿里云海外 / Cursor Cloud Agent 等),
> 在云上跑会返回 `HTTP 418` + WAF 拦截页面。
> 在你自己的笔记本上跑就完全没问题（家庭宽带 / 公司办公网都行）。

### 第 1 步 — 把 token 放进 `.env`（不会被提交到仓库）

```bash
cp .env.example .env
# 编辑 .env, 填入:
# TYC_TOKEN=<你的天眼查 token>
```

> `.gitignore` 已经把 `.env` 排除,绝不会推到 GitHub。

### 第 2 步 — 一次性批量补全

```bash
node server/enrich-all.js
```

输出示例（本地正常网络）：
```
╭─────────────────────────────────────────────╮
│  天眼查批量补全 · 共 171 家企业
│  Token: 735d7ff4...5062
│  节流: 1200ms/条
╰─────────────────────────────────────────────╯
  已有: 0    待补: 171

[  1/171] 0✅ 0❌ 0∅  ETA 3 分 25 秒   招商银行 杨浦支行
[  2/171] 1✅ 0❌ 0∅  ETA 3 分 22 秒   平安人寿 上海北分
…
[171/171] 156✅ 8❌ 7∅  ETA 0 秒

  完成 → 成功 156 · 失败 8 · 无匹配 7
  结果已写入: data/enriched-tenants.json
```

特性：
- **断点续传**：重跑只补还没成功的,已经拿到的不会再请求
- **节流**：默认 1.2 秒/条,可用 `THROTTLE_MS=2000 node server/enrich-all.js` 调慢
- **每 10 条 flush 到磁盘**：意外中断也保留进度
- **智能检测 WAF/Token 失效**：提前退出并打印明确错误

### 第 3 步 — 刷新网页

刷新后,顶部 badge 会从 `⚪ 工商数据未补全` 变为 `✅ 工商已补全 156 家 · 2026-05-28 11:32:05`。
企业明细的【📞 电话】列自动填充,详情面板里会显示法人、注册资本等信息。
CSV 导出会带上完整的 25 列工商字段。

### 备选：实时调用代理（边用边查）

也可以启动代理服务,在网页里随时点【🔄 同步电话】按钮：
```bash
node server/qcc-proxy.js   # http://localhost:3001
```
然后页面右上角【⚙ API 配置】→ 填代理地址。详见 [`server/README.md`](./server/README.md)。

---

## 安全说明

| 文件 | 是否提交 git | 说明 |
| --- | --- | --- |
| `.env` | ❌ 不提交（已 gitignored） | 含真实 token |
| `.env.example` | ✅ 提交 | 模板,无真实 token |
| `data/enriched-tenants.json` | ✅ 提交 | 仅含企业公开工商信息（这些本来就是工商局公开的） |
| `server/qcc-proxy.js` | ✅ 提交 | 代码,不含密钥 |
| `server/enrich-all.js` | ✅ 提交 | 代码,不含密钥 |

如果你想团队共享补全结果而不让其他人看到你的 token：把 `data/enriched-tenants.json` 推到仓库,别人 clone 后**不用再跑 token**,网页加载时自动读取该文件。这就是 build-once-serve-everywhere 模式。

---

## ⏳ 剩余租期是怎么算的？

每个企业记录包含：
```js
{ leaseStart: '2023-06-01', leaseTerm: 5 }
```

页面加载时算出 `leaseEnd = leaseStart + leaseTerm`,再 `(end - now)` 推剩余月数：

| 状态 | 触发条件 | 颜色 |
| --- | --- | --- |
| 在租 | 剩余 > 6 月 | 🟢 绿 |
| 即将到期 | 0 < 剩余 ≤ 6 月 | 🟠 橙 |
| 已到期 | 剩余 ≤ 0 | 🔴 红 |
| 未知 | 缺少租约字段 | 灰 |

CSV 中包含 `租约开始 / 期限 / 到期日期 / 剩余(月)` 完整列。

---

## 目录结构

```
.
├── index.html                          # 入口页
├── css/style.css
├── js/app.js                           # 地图 / Tab / 卫星 / 租期 / 工商合并
├── data/
│   ├── competitors.js                  # 主竞品 + 内嵌企业
│   ├── extra-tenants.js                # 重点园区扩充企业（按 parkId 自动合并）
│   └── enriched-tenants.json           # ★ 工商补全结果（由 enrich-all.js 生成）
├── server/
│   ├── qcc-proxy.js                    # 长驻代理（边用边查）
│   ├── enrich-all.js                   # ★ 一次性批量补全 CLI
│   └── README.md                       # 代理部署说明
├── .env.example                        # token 模板
├── .gitignore                          # 排除 .env / node_modules
└── README.md
```

---

## 数据结构

```js
// data/competitors.js
{
  id: 'A01', name: '合生国际广场',
  category: '商办写字楼',
  address: '上海市杨浦区淞沪路18号',
  lng, lat, drive, transit, developer, area, year, subs, tags, note,
  propertyFee: 28,                  // 物业费 元/㎡/月
  nature: '民营 · 甲级写字楼',       // 性质说明
  tenants: [
    {
      name: '招商银行 杨浦支行',
      industry: '金融 / 银行',
      area: 3200, rent: 5.8, rentFreeMonths: 3,
      leaseStart: '2023-04-15', leaseTerm: 5,
      dealYear: 2023
      // phone / legalPerson / regCapital / regNumber / regTime …
      // 由 enrich-all.js 自动补上
    }
  ]
}
```

```js
// data/enriched-tenants.json
{
  "_meta": { "source": "天眼查 open.api", "enrichedAt": "...", "count": 156 },
  "tenants": {
    "A01|招商银行 杨浦支行": {
      "queryName": "招商银行 杨浦支行",
      "tycName":   "招商银行股份有限公司上海杨浦支行",
      "phone":     "021-65XXXXXX",
      "legalPerson":"张三",
      "regCapital":"5000万元",
      "regNumber": "91310110xxxxxxxx",
      "regTime":   "2008-09-12",
      "regStatus": "存续"
    }
  }
}
```

启动时 `app.js`：
1. 加载 `competitors.js` + `extra-tenants.js` → 拼成 171 家
2. fetch `enriched-tenants.json` → 命中的 key 把 phone/法人/注册资本… 合并进 tenant
3. 渲染。表格、InfoWindow、CSV 导出自动用到这些字段

---

## CSV 导出（企业明细 Tab,25 列）

| 字段 | 来源 |
| --- | --- |
| 序号 / 项目名称 / 企业名称 | 数据集 |
| **工商登记名** | 天眼查（`tycName`） |
| 地址 / 距离(米) / 距离圈层 | 计算 |
| 办公体量 / 成交租金 / 免租期 | 数据集 |
| 租约开始 / 期限 / 到期日期 / 剩余(月) | 计算 |
| 主力客户行业 | 数据集 |
| **电话** | 手动录入 + 天眼查 |
| **法人代表 / 注册资本 / 统一信用代码 / 注册时间 / 登记状态** | 天眼查 |
| 物业费 / 性质 / 成交年份 / 备注 | 数据集 |

---

## 已知限制

- **WAF 阻断**：天眼查防火墙封禁绝大多数云服务器 IP,**必须在本地笔记本 / 办公网络运行** `enrich-all.js`
- **API 配额**：天眼查每个 token 都有日 / 月调用次数限制,补全 171 家约消耗 171 次配额
- **商业授权**：工商信息商业用途请取得 天眼查 / 企查查 的商业授权
