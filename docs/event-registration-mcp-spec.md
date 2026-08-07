# 活动报名 MCP 规格：工具清单与 Provider 优先级

> 状态：设计稿（不写实现代码）  
> 目标刚需：**创建报名 → 产出可分享链接 → 回流报名名单**  
> 原则：统一工具面 + 可插拔 Provider；先跑通开放能力强的平台，互动吧预留接口。

---

## 1. 产品边界

### 做

- Agent / 运营一句话创建报名页，立刻拿到链接
- 按活动拉取名单、关闭或调整报名
- 多平台切换时，对话习惯不变

### 不做（首期）

- 签到墙、抽奖、直播、公域推荐流
- 复杂票务财务（退款规则、分账）
- 各平台完整功能镜像

### 成功标准（MVP）

1. `create_registration` 成功返回 `registration_url`
2. `list_registrants` 能拿到至少姓名/手机/提交时间
3. 默认 Provider 可在无人工点后台的情况下完成 1–2
4. `provider=hudongba` 即便暂不可自动发布，也能输出「待发布草稿 + 人工确认步骤」，不阻塞统一工具面

---

## 2. 统一数据模型

```ts
type RegistrationEvent = {
  id: string                 // 内部 ID：{provider}:{native_id}
  provider: ProviderId
  title: string
  start_at?: string          // ISO 8601
  end_at?: string
  location?: string
  capacity?: number
  status: "draft" | "open" | "closed" | "needs_manual_publish"
  registration_url?: string
  admin_url?: string
  fields: FieldDef[]         // 报名表字段
  native: Record<string, unknown>  // 平台原始 ID / token
}

type FieldDef = {
  key: string
  label: string
  type: "text" | "phone" | "email" | "single_select" | "multi_select" | "date" | "number"
  required?: boolean
  options?: string[]
}

type Registrant = {
  id: string
  submitted_at: string
  fields: Record<string, string | number | boolean | string[]>
  channel?: string           // URL 传参 / 分销来源
}
```

默认报名字段建议（各 Provider 映射）：

| key | label | type | required |
|---|---|---|---|
| name | 姓名 | text | yes |
| phone | 手机号 | phone | yes |
| company | 单位/公司 | text | no |
| title | 职位 | text | no |
| remark | 备注 | text | no |

---

## 3. MCP 工具清单

对外只暴露统一工具；`provider` 可选，缺省走配置默认值。

### P0（MVP 必做）

#### `create_registration`

- **用途**：创建并尽量发布报名，返回链接
- **入参**：
  - `title`（必填）
  - `start_at` / `end_at` / `location` / `capacity`（可选）
  - `description`（可选，纯文本或 Markdown）
  - `fields`（可选；缺省用默认字段）
  - `provider`（可选）
  - `channel_defaults`（可选，如 `utm_source`）
  - `auto_publish`（默认 `true`；互动吧若无法自动发布则降级）
- **出参**：`RegistrationEvent`（至少含 `id`、`status`、`registration_url` 或人工步骤）
- **失败策略**：
  - 金数据/飞书：抛可重试错误
  - 互动吧：返回 `status=needs_manual_publish` + `checklist[]`，不视为硬失败

#### `get_registration_link`

- **用途**：按活动 ID 取最新报名链接 / 短链 / 嵌入代码
- **入参**：`event_id`；可选 `with_channel`（追加渠道参数）
- **出参**：`{ registration_url, embed_html?, qr_hint? }`

#### `list_registrants`

- **用途**：分页拉取报名名单
- **入参**：`event_id`；`cursor?`；`limit?`（默认 50，最大 200）；`since?`
- **出参**：`{ items: Registrant[], next_cursor? }`

#### `close_registration`

- **用途**：停止接收新报名（或设为满员/下线）
- **入参**：`event_id`
- **出参**：更新后的 `RegistrationEvent`

### P1（闭环增强）

| 工具 | 用途 |
|---|---|
| `update_registration` | 改标题、时间、名额、字段 |
| `get_registration` | 查单个活动元数据与状态 |
| `export_registrants` | 导出 CSV / 写入飞书表 |
| `set_webhook` | 配置报名提交回调（若 Provider 支持） |
| `list_registrations` | 列出本账号近期活动 |

### P2（有规模再做）

| 工具 | 用途 |
|---|---|
| `create_channel_link` | 分销/渠道专属链接 |
| `get_checkin_code` | 签到码/核销入口 |
| `duplicate_registration` | 从历史活动复制 |
| `send_reminder` | 短信/微信提醒（强依赖平台能力） |

### 明确不做进 MCP 的能力

- 海报智能设计、公域推荐、售票营销玩法
- 直播、上墙、抽奖
- 任意平台后台「点哪都行」的全量代理

---

## 4. Provider 优先级

评分维度（每项 1–5）：**开放成熟度、能否程序化出链接、国内传播匹配、凭证获取成本、与刚需契合度**。

| 优先级 | Provider ID | 开放成熟度 | 程序化出链接 | 国内传播 | 凭证成本 | 建议角色 |
|---|---|---|---|---|---|---|
| **1** | `jinshuju` | 5 | 5 | 3 | 2 | **首期默认实现** |
| **2** | `feishu` | 5 | 4 | 3 | 2 | **内部/团队场景并行** |
| **3** | `hudongba` | 2 | 2 | 5 | 4 | **主传播预留；半自动** |
| **4** | `huodongxing` | 2 | 3 | 4 | 4 | 备选；客户指定再接 |
| **5** | `eventbrite` | 5 | 5 | 1 | 2 | 海外/演示/对照实现 |

### 4.1 `jinshuju`（金数据）— 先跑通

- **理由**：官方已有 API v1、Webhook、表单嵌入、**官方 MCP**；最接近「创建 → 链接 → 名单」。
- **映射**：
  - 创建表单 → `create_registration`
  - 表单公开 URL → `registration_url`
  - 表单数据列表 / Webhook → `list_registrants` / 实时同步
  - URL 传参 → 渠道追踪
- **首期策略**：
  - 方案 A（最快）：直接代理/封装官方 MCP，统一改名为上述工具
  - 方案 B：自研薄封装调 API v1，便于多 Provider 同构
- **凭证**：API Key / Token（环境变量 `JINSHUJU_*`）
- **风险**：公域曝光不如互动吧；收费套餐限制需确认「程序化创建表单」权限

### 4.2 `feishu`（飞书多维表格表单）— 内部刚需

- **理由**：Open API 成熟；可建表/字段、开启表单分享拿 `shared_url`；名单天然进多维表，利于运营。
- **映射**：
  - Base + Table + Form + `shared=true` → 报名链接
  - 记录列表 → 报名名单
- **适用**：内部沙龙、校友会、园区活动；链接主要在企业微信/飞书/微信群传播
- **凭证**：`FEISHU_APP_ID` / `FEISHU_APP_SECRET` + 目标 `app_token`
- **风险**：对外「活动品牌页」弱；外部分享权限需设 `anyone_editable`

### 4.3 `hudongba`（互动吧）— 预留接口

- **理由**：当前主传播渠道；公开开发者文档弱，企业对接门槛高。
- **首期实现形态（半自动）**：
  1. Agent 生成标准化活动草稿（标题、时间、地点、字段、文案）
  2. 输出人工发布清单 + 建议粘贴内容
  3. 人工发布后，用 `register_external_link(event_id, url)`（P1 辅助工具）回填链接
  4. 名单：优先人工导出后 `import_registrants`；若后续拿到官方 API 再替换为全自动
- **接口预留**（实现层）：

```ts
interface EventRegistrationProvider {
  id: ProviderId
  capabilities: {
    create: boolean
    publish: boolean
    shareLink: boolean
    listRegistrants: boolean
    close: boolean
    webhook: boolean
  }
  createRegistration(input): Promise<RegistrationEvent>
  getRegistrationLink(nativeId, opts?): Promise<{ registration_url: string }>
  listRegistrants(nativeId, page): Promise<{ items: Registrant[]; next_cursor?: string }>
  closeRegistration(nativeId): Promise<RegistrationEvent>
}
```

- `hudongba` 首期：`create/publish/shareLink/list` 多为 `false` 或降级实现；但必须实现同一接口，避免日后双轨。

### 4.4 `huodongxing`（活动行）— 第二梯队

- **已确认公开能力**：开放平台页侧重「嵌入代码」把报名嵌进自有站点。
- **评估**：有同类报名链接能力，但「程序化创建活动并回传 URL」文档不完整，不宜首期投入。
- **触发接入条件**（满足任一即可排期）：
  - 客户强制要求活动行链接
  - 拿到可用的企业 API / 稳定嵌入工作流
  - 互动吧对接长期卡住且活动行更易谈下来

### 4.5 `eventbrite` — 对照实现（可选）

- 完整 REST：创建活动 → 票种 → publish → URL。
- 用作架构冒烟与海外场景；国内传播价值低，不进默认 Provider。

---

## 5. 推荐落地顺序

```text
Phase 0  本文档评审通过
Phase 1  统一 MCP skeleton + jinshuju Provider（真链路基线）
Phase 2  feishu Provider（内部活动双轨）
Phase 3  hudongba 半自动 Provider（草稿 + 回填链接 + 名单导入）
Phase 4  Webhook → 飞书表/通知；export_registrants
Phase 5  条件成熟后再接 huodongxing / 官方互动吧 API
```

### Phase 1 验收用例

> 「帮我创建一个下周三 19:00 的 AI 沙龙报名，地点张江，限额 40 人，要姓名和手机。」

期望：

1. 返回可点击 `registration_url`
2. 用该链接提交一条测试报名
3. `list_registrants` 能看到该条

### Phase 3 验收用例（互动吧）

> 同上，但 `provider=hudongba`

期望：

1. 返回结构化草稿 + 人工发布步骤（非空链接硬失败）
2. 人工发布后回填 URL，后续 `get_registration_link` 正常
3. 导入/同步名单后 `list_registrants` 可用

---

## 6. 配置与安全

| 项 | 说明 |
|---|---|
| 默认 Provider | 环境变量 `EVENT_REG_DEFAULT_PROVIDER=jinshuju` |
| 密钥 | 仅环境变量 / 密钥管理，禁止写入仓库 |
| 权限最小化 | 金数据/飞书应用只开表单与数据读写 |
| 审计 | 每次 `create_*` 记 `who/when/provider/event_id/url` |
| 隐私 | 名单含手机号；导出与日志需脱敏选项 |

---

## 7. 决策建议（给业务方）

1. **对外传播继续用互动吧**，但自动化先别赌官方 API。
2. **自动化出链先用金数据（或飞书）跑通**，证明 Agent 工作流。
3. **活动行可以进路线图，但排在互动吧预留之后**，除非客户点名。
4. **不要为每个平台单独做一个 MCP**；一个 MCP、多个 Provider。
5. 若短期只要「能发链接收名单」，金数据官方 MCP 可直接试用；我们这套规格的价值在于：**统一工具名 + 预留互动吧 + 可换平台**。

---

## 8. 下一步（实现时）

确认本规格后，最小原型建议目录：

```text
packages/event-registration-mcp/
  src/
    index.ts              # MCP server 入口
    tools/*.ts            # 统一工具
    providers/
      types.ts
      jinshuju.ts
      feishu.ts
      hudongba.ts         # 半自动 stub
    model.ts
  README.md
```

首期只接线：`create_registration` / `get_registration_link` / `list_registrants` / `close_registration`，默认 `jinshuju`，`hudongba` 返回草稿清单。
