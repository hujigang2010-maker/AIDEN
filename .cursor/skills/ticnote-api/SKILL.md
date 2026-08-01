---
name: ticnote-api
description: |
  TicNote API integration: API Key获取说明、Token认证、文件上传、知识库项目列表查询。
  Triggers:
  - "ticnote", "TicNote"
  - "获取apikey", "获取api key", "get api key"
  - "获取token", "get token", "ticnote token"
  - "上传文件到ticnote", "upload to ticnote"
  - "知识库列表", "项目列表", "knowledge base list"
  - "ticnote api", "ticnote接口"
---

# TicNote API Integration

TicNote 平台的 API 集成指南，覆盖认证流程和核心接口操作。

## 基础信息

| 项目 | 值 |
|------|-----|
| **认证方式** | Bearer Token |
| **Token 有效期** | 24 小时 |
| **请求格式** | JSON（文件上传除外） |

> ⚠️ 以上信息为占位内容，待补充真实 API 地址和参数。

## Workflow

根据用户需求，执行对应操作：

> **凭证缺失时的行为规则：** 当用户请求任何需要 Token 的操作（如获取项目列表、上传文件等），但当前上下文中没有 AppKey / Token 时，必须在回复中**同时**提供以下两部分内容：
> 1. 请求用户提供 AppKey
> 2. **附上获取 AppKey 的完整指引**（见下方「获取 API Key」章节），帮助还没有 AppKey 的用户自行获取

1. **获取 API Key** → 提供 TicNote 平台操作说明（参考 [API Key 获取指南](references/get-apikey.md)）
2. **获取 Token** → 使用 AppKey 调用 `/api/p1/appkey/login` 接口获取 Bearer Token
3. **查询知识库项目列表** → 使用 Token 调用 `/api/v2/file-index/chats` 接口获取项目/文件夹列表
4. **获取项目下文件列表** → 使用 Token 调用 `GET /api/v1/file-index/file-tree?rootId={projectId}` 获取项目下所有文件
5. **上传文件** → 使用 Token 调用文件上传接口（三步：签名 → COS 上传 → 注册知识库）
6. **查看文件详情 / 轮询转写** → `GET /api/v2/file-index/file-detail/{recordId}`
7. **提交转写任务** → `POST /api/v1/task/transcribe/commit`
8. **重新总结（音频）** → `POST /api/v1/task/resummary/commit`
9. **文件总结（非音频）** → `POST /api/project/project/summary/local_file`
10. **翻译** → `POST /api/v1/translate`
11. **分享文件** → `POST /api/share/{audio|localFile}`，访问 `GET /api/share/{shareType}/{shareCode}`
12. **生成播客** → `POST /api/podcast/generate`
13. **Deep Research** → `POST /api/v1/deep/research/query`
14. **用户设置** → `GET/PUT /api/v1/user/setting`
15. **知识库文件管理** → 删除、重命名、复制、移动

### 1. 获取 API Key（说明性）

引导用户在 TicNote 平台获取 API Key，详细步骤见 `references/get-apikey.md`。

### 2. 获取 Token

**接口：** `POST /api/p1/appkey/login`

通过 AppKey 登录获取 Bearer Token，无需额外认证。

```bash
curl -X POST <BASE_URL>/api/p1/appkey/login \
  -H "Content-Type: application/json" \
  -d '{"appkey": "<YOUR_APPKEY>"}'
```

**响应示例：**
```json
{
  "code": 0,
  "data": {
    "token": "eyJhbGciOiJIUzI1NiIs...",
    "firstLogin": false,
    "tokenType": "Bearer"
  }
}
```

> 后续请求需在 Header 中携带：`Authorization: Bearer <token>`

执行方式：
```bash
python ${SKILL_PATH}/scripts/get_token.py --appkey "<APPKEY>"
```

**错误响应示例：**
```json
{
  "code": 11865,
  "msg": "Appkey not found"
}
```

当接口返回非 0 的 `code` 时，说明请求失败，需要将 `msg` 内容反馈给用户并附上详细说明：

| code | msg | 说明 |
|------|-----|------|
| 11865 | Appkey not found | AppKey 无效或不存在。请检查：1）AppKey 是否拼写正确、完整（以 `tnovs_sk_`、`tnovs_sit_sk_`、`tncn_sk_` 或 `tncn_sit_sk_` 开头）；2）是否已在 TicNote 平台成功获取过 AppKey；3）如确认无误但仍报错，需联系管理员确认该 AppKey 是否已被删除或重置。 |
| 11864 | User not found | AppKey 对应的用户不存在。可能是账号已注销或从未注册，请确认关联的账号（邮箱/手机号）状态正常。 |
| 11866 | Account associated with appkey not found | AppKey 对应的账号在系统中不存在。可能是账号已被注销，请使用有效账号重新获取 AppKey。 |

### 3. 获取知识库项目/文件夹列表

**接口：** `GET /api/v2/file-index/chats`

获取当前用户知识库中的项目（文件夹）列表，支持按名称模糊搜索。

**请求参数：**

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| query | string | 否 | 搜索关键词，按项目名称模糊匹配 |

```bash
curl -X GET "<BASE_URL>/api/v2/file-index/chats?query=<KEYWORD>" \
  -H "Authorization: Bearer <TOKEN>"
```

**响应示例：**
```json
{
  "chats": [
    {
      "id": "chat_001",
      "name": "研发文档库",
      "type": "private_virtual",
      "chat_type": "virtual_employee",
      "is_group": false,
      "has_agent": true,
      "project_id": "10086",
      "project_name": "研发文档库",
      "projectInfo": {
        "id": "10086",
        "name": "研发文档库",
        "icon": "📁",
        "color": "#4A90D9",
        "recordType": 1,
        "fileNum": 42
      },
      "participants": [
        {
          "id": "1001",
          "name": "用户A",
          "type": "human",
          "isCurrentUser": true,
          "role": "owner"
        }
      ],
      "agent_count": 1,
      "createdAt": "2026-01-15T08:00:00Z",
      "updatedAt": "2026-03-10T12:00:00Z",
      "lastMessageAt": "2026-03-12T09:30:00Z"
    }
  ]
}
```

**响应字段说明：**

| 字段 | 说明 |
|------|------|
| `chats[].name` | 项目/文件夹名称 |
| `chats[].projectInfo.fileNum` | 文件数量 |

> ⚠️ 输出规则：列表仅展示**序号、名称、文件数**，不对用户展示 `project_id`、`chat_type` 等内部字段。

执行方式：
```bash
python ${SKILL_PATH}/scripts/list_projects.py --token "<TOKEN>" --appkey "<APPKEY>" [--query "<KEYWORD>"]
```

### 4. 获取项目下文件列表

**接口：** `GET /api/v1/file-index/file-tree`

获取指定项目（文件夹）下的所有文件列表，返回文件树结构。

**请求参数（Query）：**

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| rootId | long | 是 | 项目 ID（即 `project_id`，来自项目列表接口） |

```bash
curl -X GET "<BASE_URL>/api/v1/file-index/file-tree?rootId=<PROJECT_ID>" \
  -H "Authorization: Bearer <TOKEN>"
```

**响应示例：**
```json
{
  "success": true,
  "fileTree": [
    {
      "id": "2020733843988295682",
      "fileId": "2020733843988295681",
      "fileType": "agent_file",
      "name": "西安天气推送-2026-02-09.html",
      "type": "file",
      "path": "西安天气推送-2026-02-09.html",
      "subRemark": "{\"transcodeStatus\":\"suc\",\"notRead\":1,\"originSuffix\":\"agent_file\"}",
      "children": []
    }
  ]
}
```

**响应字段说明：**

| 字段 | 说明 |
|------|------|
| `fileTree` | 文件树数组 |
| `fileTree[].id` | 记录 ID（recordId），用于文件详情查询 |
| `fileTree[].fileId` | 文件 ID |
| `fileTree[].fileType` | 文件类型（如 `agent_file`、`upload_recording`、`recording_file`、`pdf`、`docx` 等） |
| `fileTree[].name` | 文件名 |
| `fileTree[].type` | 节点类型（`file` 或 `directory`） |
| `fileTree[].path` | 文件路径 |
| `fileTree[].subRemark` | 附加信息 JSON（含 `transcodeStatus`、`summaryId`、`transcribeId`、`deepResearchStatus` 等） |
| `fileTree[].children` | 子节点（目录时包含子文件） |

执行方式：
```bash
python ${SKILL_PATH}/scripts/list_files.py --token "<TOKEN>" --appkey "<APPKEY>" --root-id "<PROJECT_ID>"
```

> ⚠️ 输出规则：列表仅展示**序号、名称**，不对用户展示 `fileType`、`id` 等内部字段。

### 5. 上传文件到知识库

上传文件分三步：获取签名 → PUT 上传 COS → 注册知识库。

**目录选择流程：** 当用户未指定 `parentId`（目标目录）时：
1. 先调用「获取项目列表」接口（见第 3 节），将项目列表展示给用户
2. 用户选择某个项目 → 使用该 `project_id` 作为 `parentId` 执行上传
3. 用户未选择任何项目、直接要求继续 → 以 `parentId` 为空发起上传请求（后端会自动上传到用户的默认目录）

**上传前总结确认（重要）：** 用户请求上传文件时，在**执行上传之前**必须先询问是否需要上传完成后一并触发文件总结：
- 提示示例：「准备上传 xxx 到 xxx 项目，上传完成后是否需要一并触发文件总结？」
- 用户确认需要总结 → 上传完成后自动调用对应总结接口（音频用第 8 节，非音频用第 9 节），无需二次确认
- 用户明确拒绝 → 仅上传，不触发总结
- 用户未明确回复（如只说"上传"）→ 先问再传，不要跳过这个确认步骤

**fileId 确认机制（重要）：** 上传脚本注册成功后，会自动通过文件列表接口确认真实的 `fileId`，并输出 `confirmed_file_id` 字段。后续操作（总结、转写等）**必须使用 `confirmed_file_id`**，而非注册接口直接返回的 `fileId`（两者可能不一致）。

执行方式：
```bash
# 完整上传（指定目标目录）
python ${SKILL_PATH}/scripts/upload_file.py --token "<TOKEN>" --appkey "<APPKEY>" --file "/path/to/file" --parent-id "<PROJECT_ID>"

# 未指定目录时，脚本会自动列出项目让用户选择
python ${SKILL_PATH}/scripts/upload_file.py --token "<TOKEN>" --appkey "<APPKEY>" --file "/path/to/file"
```

> ⚠️ 输出规则：项目选择列表仅展示**序号、名称、文件数**，不对用户展示 `project_id` 等内部字段。

#### 步骤一：获取 COS 上传签名

**接口：** `GET /api/v1/tencent/oss/apply/token`

通过后端签名服务获取腾讯云 COS 上传授权，无需本地安装 COS SDK。

> 参考前端实现：`docs/ufile_tencent.js` → `getUFileToken()`，`docs/UploadTool.ts` → `_tokenUrl`

**请求参数（Query）：**

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| method | string | 是 | 与实际上传方式一致：普通上传 `PUT`，分片初始化 `POST` |
| bucket | string | 是 | COS Bucket 名称，如 `tc-nj-ticnote-1324023246` |
| key | string | 是 | 文件在 COS 上的完整路径（自动 URL 编码） |
| content_md5 | string | 否 | 文件内容 MD5（前端 SparkMD5 计算，可为空） |
| contentType | string | 否 | 文件 MIME 类型，如 `audio/wav`、`application/pdf` |
| date | string | 否 | 可为空 |

```bash
curl -X GET "<BASE_URL>/api/v1/tencent/oss/apply/token?method=PUT&bucket=tc-nj-ticnote-1324023246&key=ticnote-web-prd%2F2026-03%2Fabc123def456.pdf&content_md5=&contentType=application/pdf&date=" \
  -H "Authorization: Bearer <TOKEN>"
```

**响应：** JSON 格式，签名在 `.data` 字段（普通上传）或 `.auth` 字段（分片上传）。

#### 步骤二：PUT 上传文件到腾讯云 COS

使用步骤一返回的签名，直接通过 HTTP PUT 将文件上传到 COS，无需 cos-python-sdk-v5。

> 参考前端实现：`docs/ufile_tencent.js` → `uploadFile()`

**COS 上传地址：** `https://{bucket}.cos.{region}.myqcloud.com/{key}`

**请求 Headers：**
| Header | 说明 |
|--------|------|
| Authorization | 步骤一返回的签名字符串 |
| Content-Type | 文件 MIME 类型 |
| Content-MD5 | 文件内容 MD5（与签名时一致） |

```bash
curl -X PUT "https://{bucket}.cos.{region}.myqcloud.com/{key}" \
  -H "Authorization: <签名字符串>" \
  -H "Content-Type: application/pdf" \
  -H "Content-MD5: <文件MD5>" \
  --data-binary @file.pdf
```

**文件存储路径（key）规则：** `{env}/{YYYY-MM}/{uuid}.{ext}`
- `env`：根据 AppKey 判断——AppKey 中包含 `sit` 则为 `ticnote-web-sit`，否则为 `ticnote-web-prd`
- `YYYY-MM`：当前年月，如 `2026-03`
- `uuid`：随机生成的 UUID（32 位 hex，不含连字符）
- `ext`：文件扩展名

> 参考前端实现：`docs/UploadTool.ts` → `_joinFileName()` 使用 `genorateUuid()`

**COS 配置（根据 AppKey 自动选择）：**

| AppKey 前缀 | 区域 | Bucket | Region | CDN |
|---|---|---|---|---|
| `tncn_*` | 国内 | `tc-nj-ticnote-1324023246` | `ap-nanjing` | `https://cdn.ticnote.cn` |
| `tnovs_*` | 海外 | `voice-recorder-1308581983` | `na-siliconvalley` | `https://voice-recorder-cdn.ticnote.com` |

> 国内 sit/prd 共用同一个 bucket，海外 sit/prd 亦然。env 前缀（`ticnote-web-sit` / `ticnote-web-prd`）区分文件归属。

**支持的文件类型：** PDF、TXT、DOC/DOCX、XLS/XLSX、PPT/PPTX、MD、CSV、HTML、MP3、WAV、MP4、图片等。

#### 步骤三：注册文件到知识库

**接口：** `POST /api/v1/knowledge/upload`

将 COS 上传成功后的文件 CDN URL 注册到知识库指定目录。

> 注册时使用 **CDN URL**（`{cdn_domain}/{key}`），与前端 `UploadTool.ts` 一致。

**请求参数：**

| 参数 | 类型 | 位置 | 必填 | 说明 |
|------|------|------|------|------|
| parentId | long | query | 否 | 父目录 ID（即 project_id），为空时上传到根目录 |

**请求体（JSON 数组）：**

```json
[
  {
    "fileName": "文件名.pdf",
    "fileType": "pdf",
    "fileUrl": "https://cdn.ticnote.cn/ticnote-web-prd/2026-03/uuid.pdf",
    "recordType": 1
  }
]
```

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| fileName | string | 是 | 文件名（含扩展名） |
| fileType | string | 是 | 文件类型（扩展名，如 `pdf`、`docx`、`mp3`） |
| fileUrl | string | 是 | 文件在 COS 上的完整 URL |
| recordType | int | 否 | 记录类型 |
| fileId | long | 否 | 文件 ID（已有文件时使用） |
| recordTime | long | 否 | 录制时间戳（音频类文件） |

```bash
curl -X POST "<BASE_URL>/api/v1/knowledge/upload?parentId=<PROJECT_ID>" \
  -H "Authorization: Bearer <TOKEN>" \
  -H "Content-Type: application/json" \
  -d '[{"fileName":"report.pdf","fileType":"pdf","fileUrl":"https://cdn.example.com/ticnote-web-prd/2026-03/abc123.pdf"}]'
```

**响应示例：**
```json
{
  "code": 0,
  "data": {
    "totalCount": 1,
    "successCount": 1,
    "failedCount": 0,
    "successFiles": [
      {
        "recordId": 2031647298439811075,
        "fileId": 98765,
        "fileName": "report.pdf",
        "fileType": "pdf",
        "fileUrl": "https://cdn.ticnote.cn/ticnote-web-prd/2026-03/uuid.pdf",
        "status": 1,
        "type": 2
      }
    ],
    "failedFiles": []
  }
}
```

> ⚠️ **关键字段：** `successFiles[].recordId` 是后续查询文件详情（轮询转写状态）的必要参数。

**successFiles 字段说明：**

| 字段 | 说明 |
|------|------|
| recordId | 知识库记录 ID（用于 file-detail 查询） |
| fileId | 文件 ID |
| fileName | 文件名 |
| fileType | 文件类型 |
| fileUrl | 文件 URL |
| status | 处理状态（见状态码表） |
| type | 文件类型（0=录音目录, 1=目录, 2=文件） |

**failedFiles 字段说明：**

| 字段 | 说明 |
|------|------|
| fileName | 文件名 |
| fileType | 文件类型 |
| fileUrl | 文件 URL |
| failureReason | 失败原因 |
| errorCode | 错误代码 |

### 6. 查看文件详情 / 轮询转写状态

**接口：** `GET /api/v2/file-index/file-detail/{recordId}`

两种使用场景：
1. **查看文件内容** — 获取单个文件的详细信息（转写文本、总结、元信息等）
2. **轮询转写状态** — 上传音视频文件后，定期轮询等待转写完成

**请求参数：**

| 参数 | 类型 | 位置 | 必填 | 说明 |
|------|------|------|------|------|
| recordId | long | path | 是 | 知识库记录 ID（来自上传响应的 `successFiles[].recordId`） |

```bash
curl -X GET "<BASE_URL>/api/v2/file-index/file-detail/<RECORD_ID>" \
  -H "Authorization: Bearer <TOKEN>"
```

**响应示例（音视频文件）：**
```json
{
  "code": 0,
  "data": {
    "recordId": 2031647298439811075,
    "fileId": 98765,
    "fileName": "meeting-recording.mp3",
    "title": "meeting-recording.mp3",
    "fileType": "mp3",
    "type": 2,
    "isVoice": true,
    "status": 2,
    "transcodeStatus": "suc",
    "fileUrl": "https://cdn.ticnote.cn/...",
    "formatUrl": "https://cdn.ticnote.cn/...",
    "duration": 3600,
    "language": "zh",
    "transcribeId": 123456,
    "transcribeJson": "{ ... 转写内容 JSON ... }",
    "summaryId": 789012,
    "summaryJson": "{ ... 总结内容 JSON ... }",
    "dataVersion": "v2",
    "deepResearchStatus": 0,
    "updateTime": "2026-03-12T12:00:00Z",
    "owner": { "id": "1001", "name": "用户A" }
  }
}
```

**响应示例（文档文件）：**
```json
{
  "code": 0,
  "data": {
    "recordId": 2031647298439811075,
    "fileId": 98765,
    "fileName": "report.pdf",
    "title": "report.pdf",
    "fileType": "pdf",
    "type": 2,
    "isVoice": false,
    "status": 2,
    "transcodeStatus": "suc",
    "fileUrl": "https://cdn.ticnote.cn/...",
    "transcribeJson": "{ ... 文件预览内容 ... }",
    "deepResearchStatus": 0,
    "updateTime": "2026-03-12T12:00:00Z",
    "owner": { "id": "1001", "name": "用户A" }
  }
}
```

**关键字段说明：**

| 字段 | 说明 |
|------|------|
| `recordId` | 知识库记录 ID |
| `fileId` | 文件 ID |
| `isVoice` | 是否为音视频文件 |
| `status` | 处理状态（见下方状态码表） |
| `transcodeStatus` | 转码状态（见下方状态码表） |
| `transcribeJson` | 转写/预览内容（JSON 字符串） |
| `summaryJson` | 总结内容（JSON 字符串，仅音视频） |
| `duration` | 文件时长-秒（仅音视频） |
| `language` | 识别语言（仅音视频） |
| `deepResearchStatus` | 深度研究状态：0=未开始, 1=已完成 |
| `dprSessionId` | Deep Research 会话 ID — 文件创建时自动生成，用于 Deep Research 接口的 `sessionId` 参数 |

**status（处理状态）枚举：**

| code | 名称 | 说明 |
|------|------|------|
| -1 | RECORDING | 录音中 |
| 0 | PENDING | 待处理 |
| 1 | PROCESSING | 处理中 |
| 2 | COMPLETED | 已完成 |
| 3 | FAILED | 失败 |
| 4 | TRANSSUC | 转录成功 |
| 5 | SUMMARYSUC | 总结成功 |

**transcodeStatus（转码状态）枚举：**

| 值 | 说明 |
|------|------|
| `null` | 未开始 / 等待中 |
| `ing` | 转码中 |
| `suc` | 转码成功 |
| `fail` | 转码失败 |
| `no_rights` | 无权限（需要 VIP） |

**轮询策略（音视频文件转写）：**

上传音视频文件后，后端会自动触发转码和转写。需要轮询判断完成的条件：
1. `transcodeStatus == "suc"` — 转码完成
2. `status >= 2` — 处理完成（COMPLETED / TRANSSUC / SUMMARYSUC）
3. `transcribeJson` 不为空 — 转写内容已生成

推荐轮询参数：
- 间隔：5 秒
- 超时：10 分钟（600 秒）
- 提前退出：`transcodeStatus == "fail"` 或 `status == 3`（FAILED）

**需要转码的文件类型：** MP3, WAV, MP4, MOV, M4A, CAF, AVI, RMVB, OPUS, AAC（即 `MediaFileTypeEnum` 中 `uploadNeedTranscode=true` 的类型）

执行方式：
```bash
# 查看文件详情（单次查询）
python ${SKILL_PATH}/scripts/file_detail.py --token "<TOKEN>" --appkey "<APPKEY>" --record-id "<RECORD_ID>"

# 轮询模式（等待转写完成）
python ${SKILL_PATH}/scripts/file_detail.py --token "<TOKEN>" --appkey "<APPKEY>" --record-id "<RECORD_ID>" --poll [--interval 5] [--timeout 600]
```

> ⚠️ 输出规则：仅展示**文件名、状态、转码状态、时长、语言、转写/总结内容预览**，不对用户展示 `recordId`、`fileId`、`fileType`、`dprSessionId` 等内部字段。

### 7. 提交转写任务

**接口：** `POST /api/v1/task/transcribe/commit`

上传音视频文件后，手动提交转写任务（自动转写关闭时使用）。

**请求体：**
```json
{
  "fileId": 98765,
  "language": "zh",
  "model": "qwen-max-latest",
  "hasSpeakers": true,
  "detailLevel": "more_details",
  "template": "",
  "templateCustomize": ""
}
```

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| fileId | long | 是 | 文件 ID（来自上传响应 `successFiles[].fileId`） |
| language | string | 否 | 语言代码，如 `zh`、`en` |
| model | string | 否 | 转写模型 |
| hasSpeakers | boolean | 否 | 是否区分说话人，默认 false |
| detailLevel | string | 否 | 详细程度 |
| template | string | 否 | 录音笔选用的模板 |
| templateCustomize | string | 否 | 用户自定义模板（优先级高于 template） |

**响应示例：**
```json
{"code": 0, "data": {"transcribeTaskId": "1876161722453798913"}}
```

执行方式：
```bash
python ${SKILL_PATH}/scripts/transcribe_commit.py --token "<TOKEN>" --appkey "<APPKEY>" --file-id <FILE_ID> [--language zh] [--model "qwen-max-latest"] [--has-speakers] [--detail-level "more_details"]
```

### 8. 重新总结（音频文件）

**接口：** `POST /api/v1/task/resummary/commit`

对已有转写结果重新生成总结。

**请求体：**
```json
{
  "fileId": 98765,
  "model": "o4-mini",
  "detailLevel": "more_details",
  "lang": "zh",
  "hasSpeakers": false,
  "template": "",
  "templateCustomize": ""
}
```

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| fileId | long | 是 | 文件 ID |
| model | string | 否 | 总结模型，如 `qwen-max-latest`、`o4-mini` |
| detailLevel | string | 否 | 详细程度，如 `more_details` |
| lang | string | 否 | 输入文本的语言 |
| hasSpeakers | boolean | 否 | 是否区分说话人 |
| template | string | 否 | 模板 |
| templateCustomize | string | 否 | 自定义模板 |

执行方式：
```bash
python ${SKILL_PATH}/scripts/resummary_commit.py --token "<TOKEN>" --appkey "<APPKEY>" --file-id <FILE_ID> [--model "o4-mini"] [--detail-level "more_details"] [--lang "zh"] [--has-speakers]
```

### 9. 文件总结（非音频文件）

**接口：** `POST /api/project/project/summary/local_file`

对文档类文件（PDF、DOC 等）触发总结。也可用于重新总结。

**请求参数（Query）：**

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| taskId | long | 是 | 文件 ID（即 fileId） |
| model | string | 否 | 总结模型，如 `qwen-max-latest`、`o4-mini` |
| detailLevel | string | 否 | 详细程度，如 `more_details` |

```bash
curl -X POST "<BASE_URL>/api/project/project/summary/local_file?taskId=<FILE_ID>&model=qwen-max-latest&detailLevel=more_details" \
  -H "Authorization: Bearer <TOKEN>"
```

执行方式：
```bash
python ${SKILL_PATH}/scripts/summary_local_file.py --token "<TOKEN>" --appkey "<APPKEY>" --file-id <FILE_ID> [--model "qwen-max-latest"] [--detail-level "more_details"]
```

### 10. 翻译

**接口：** `POST /api/v1/translate`

对已有转写内容创建翻译任务。

**请求体：**
```json
{
  "transcribeId": 123456,
  "targetLanguage": "en"
}
```

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| transcribeId | long | 是 | 转写 ID（来自 file-detail 响应的 `transcribeId`） |
| targetLanguage | string | 是 | 目标语言代码，如 `en`、`zh`、`ja` |

**响应示例：**
```json
{"code": 0, "data": {"transcribeTaskId": "1876161722453798913"}}
```

执行方式：
```bash
python ${SKILL_PATH}/scripts/translate.py --token "<TOKEN>" --appkey "<APPKEY>" --transcribe-id <TRANSCRIBE_ID> --target-language "en"
```

### 11. 分享文件

#### 创建分享

**接口：** `POST /api/share/{shareType}`

| shareType | 说明 |
|-----------|------|
| `audio` | 音频文件分享 |
| `localFile` | 非音频文件分享 |

**请求体：** JSON 字符串（内容根据分享类型不同）

**响应：** 返回分享码（shareCode）
```json
{"code": 0, "data": "<SHARE_CODE>"}
```

#### 访问分享

**接口：** `GET /api/share/{shareType}/{shareCode}`

通过分享码获取分享内容。

```bash
curl -X GET "<BASE_URL>/api/share/audio/<SHARE_CODE>"
```

执行方式：
```bash
# 创建分享
python ${SKILL_PATH}/scripts/share.py create --token "<TOKEN>" --appkey "<APPKEY>" --share-type audio --data '<JSON>'

# 访问分享
python ${SKILL_PATH}/scripts/share.py get --appkey "<APPKEY>" --share-type audio --share-code "<SHARE_CODE>"
```

### 12. 生成播客

**接口：** `POST /api/podcast/generate`

基于文件的总结内容生成播客音频。

**请求体：**
```json
{
  "summaryId": 789012,
  "localFileId": null
}
```

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| summaryId | long | 条件 | 总结 ID（音频文件使用） |
| localFileId | long | 条件 | 本地文件 ID（文档文件使用） |

> 二选一：音频文件传 `summaryId`，文档文件传 `localFileId`。

执行方式：
```bash
# 音频文件
python ${SKILL_PATH}/scripts/podcast_generate.py --token "<TOKEN>" --appkey "<APPKEY>" --summary-id <SUMMARY_ID>

# 文档文件
python ${SKILL_PATH}/scripts/podcast_generate.py --token "<TOKEN>" --appkey "<APPKEY>" --local-file-id <LOCAL_FILE_ID>
```

### 13. Deep Research

**接口：** `POST /api/v1/deep/research/query`

基于文件内容发起深度研究。核心参数 `sessionId` 需从文件详情接口获取。

**请求体：**
```json
{
  "sessionId": "2032367876959596548",
  "sessionType": 9,
  "question": "分析这份文档的核心观点",
  "msgId": 1773394072812,
  "outline": "- 核心观点\n- 论证逻辑",
  "source": 3
}
```

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| sessionId | string | 是 | 会话 ID — 从 `GET /api/v2/file-index/file-detail/{recordId}` 返回的 `dprSessionId` 字段获取 |
| sessionType | int | 是 | 会话类型（见下方枚举表） |
| question | string | 是 | 用户输入的研究问题 |
| msgId | long | 否 | 消息 ID — 后端会用 `System.currentTimeMillis()` 覆盖，**传任意值即可** |
| outline | string | 否 | 研究大纲 — 可手动填写或由前端 AI 预生成 |
| source | int | 否 | 来源标识 — 后端不做校验，仅用于统计追踪 |

**sessionType 枚举（ChatTypeEnum）：**

| 值 | 枚举名 | 说明 | sessionId 来源 |
|---|---|---|---|
| 5 | DEEP_RESEARCH_REPORT | 项目级 Deep Research | projectId 下的 sessionId |
| 6 | FILE_DP_RESEARCH | **音频文件** Deep Research | `data.dprSessionId`（音频文件详情） |
| 9 | LOCAL_FILE_RESEARCH | **非音频文件** Deep Research | `data.dprSessionId`（文档文件详情） |

> 完整 ChatTypeEnum 还包括：0=ASK_AI_ALL, 1=ASK_AI, 2=CHAT_WITH_SHADOW, 3=AHA_MOMENTS, 4=RANDOM_THOUGHT, 7=FILE_AHA_MOMENTS, 8=WEB_CHAT

**获取 sessionId 的步骤：**

```
1. 调用 file-detail 接口：GET /api/v2/file-index/file-detail/{recordId}
2. 从响应中获取：
   ├── 音频文件 → data.dprSessionId  (sessionType=6)
   └── 文档文件 → data.dprSessionId  (sessionType=9)
3. 判断 isVoice 字段确定 sessionType
```

**响应：** 返回研究结果
```json
{"code": 0, "data": 123456}
```

执行方式：
```bash
# 方式一：传 recordId，自动获取 dprSessionId 和 sessionType
python ${SKILL_PATH}/scripts/deep_research.py --token "<TOKEN>" --appkey "<APPKEY>" --record-id <RECORD_ID> --question "分析这份文档的核心观点"

# 方式二：直接传 sessionId + sessionType
python ${SKILL_PATH}/scripts/deep_research.py --token "<TOKEN>" --appkey "<APPKEY>" --session-id <DPR_SESSION_ID> --session-type 9 --question "分析这份文档的核心观点"
```

### 14. 用户设置（自动转写）

#### 获取设置

**接口：** `GET /api/v1/user/setting`

```bash
curl -X GET "<BASE_URL>/api/v1/user/setting" \
  -H "Authorization: Bearer <TOKEN>"
```

**响应示例：**
```json
{
  "code": 0,
  "data": {
    "autoTranscribeStyle": { "auto": true },
    "languageStyle": { "language": "zh" },
    "transcribeLanguage": { "language": "zh" }
  }
}
```

#### 保存设置

**接口：** `PUT /api/v1/user/setting`

```json
{
  "autoTranscribeStyle": { "auto": true },
  "transcribeLanguage": { "language": "zh" }
}
```

执行方式：
```bash
# 获取设置
python ${SKILL_PATH}/scripts/user_setting.py get --token "<TOKEN>" --appkey "<APPKEY>"

# 保存设置
python ${SKILL_PATH}/scripts/user_setting.py put --token "<TOKEN>" --appkey "<APPKEY>" --data '{"autoTranscribeStyle":{"auto":true}}'
```

### 15. 知识库文件管理

#### 批量删除

**接口：** `POST /api/v1/knowledge/delete/batch`

**请求体：** recordId 数组
```json
[2031647298439811075, 2031647298439811076]
```

**响应：** 返回删除数量
```json
{"code": 0, "data": 2}
```

#### 文件重命名

**接口：** `PUT /api/v1/knowledge/edit/{recordId}`

**请求体：**
```json
{
  "title": "新文件名",
  "color": "#4A90D9",
  "icon": "📄"
}
```

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| title | string | 否 | 新标题/文件名 |
| color | string | 否 | 颜色 |
| icon | string | 否 | 图标 |

#### 文件复制

**接口：** `POST /api/v1/knowledge/copyTo/{targetParentId}`

将文件复制到目标目录。

**路径参数：** `targetParentId` — 目标目录的 recordId

**请求体：** recordId 数组
```json
[2031647298439811075]
```

**响应：** 返回复制后的文件列表
```json
{"code": 0, "data": [{"recordId": ..., "title": ..., ...}]}
```

#### 文件移动

**接口：** `POST /api/v1/knowledge/moveTo/{targetParentId}`

将文件移动到目标目录。

**路径参数：** `targetParentId` — 目标目录的 recordId

**请求体：** recordId 数组
```json
[2031647298439811075]
```

**响应：** 返回移动数量
```json
{"code": 0, "data": 1}
```

执行方式：
```bash
# 批量删除
python ${SKILL_PATH}/scripts/knowledge_manage.py delete --token "<TOKEN>" --appkey "<APPKEY>" --record-ids <ID1> <ID2>

# 重命名
python ${SKILL_PATH}/scripts/knowledge_manage.py rename --token "<TOKEN>" --appkey "<APPKEY>" --record-id <RECORD_ID> --title "新文件名"

# 复制
python ${SKILL_PATH}/scripts/knowledge_manage.py copy --token "<TOKEN>" --appkey "<APPKEY>" --target-parent-id <TARGET_ID> --record-ids <ID1>

# 移动
python ${SKILL_PATH}/scripts/knowledge_manage.py move --token "<TOKEN>" --appkey "<APPKEY>" --target-parent-id <TARGET_ID> --record-ids <ID1>
```

## 脚本路径说明

`${SKILL_PATH}` = `.codebanana/.skills/ticnote-api`，使用 `run_terminal_cmd` 执行所有脚本。

| # | 功能 | 脚本 |
|---|------|------|
| 2 | 获取 Token | `scripts/get_token.py` |
| 3 | 知识库项目列表 | `scripts/list_projects.py` |
| 4 | 项目下文件列表 | `scripts/list_files.py` |
| 5 | 上传文件 | `scripts/upload_file.py` |
| 6 | 文件详情/轮询 | `scripts/file_detail.py` |
| 7 | 提交转写任务 | `scripts/transcribe_commit.py` |
| 8 | 重新总结（音频） | `scripts/resummary_commit.py` |
| 9 | 文件总结（非音频） | `scripts/summary_local_file.py` |
| 10 | 翻译 | `scripts/translate.py` |
| 11 | 分享文件 | `scripts/share.py` |
| 12 | 生成播客 | `scripts/podcast_generate.py` |
| 13 | Deep Research | `scripts/deep_research.py` |
| 14 | 用户设置 | `scripts/user_setting.py` |
| 15 | 知识库文件管理 | `scripts/knowledge_manage.py` |

## 错误处理

### HTTP 层错误

| HTTP 状态码 | 含义 | 处理方式 |
|-------------|------|----------|
| 401 | Token 无效或过期 | 重新获取 Token |
| 403 | API Key 权限不足 | 提示用户检查 Key 权限 |
| 413 | 文件过大 | 提示文件大小限制 |
| 429 | 请求频率超限 | 等待后重试 |

### 业务层错误（HTTP 200 但 code ≠ 0/200）

| 业务 code | 接口 | 含义 | 处理方式 |
|-----------|------|------|----------|
| 503 | `/api/v1/knowledge/upload` | 无权限写入目标项目 | 当前 AppKey 对应的账号不是该项目的 owner。**处理流程：** 1）告知用户无权限写入该项目；2）列出当前账号拥有的项目供用户重新选择；3）或提示用户更换对应 owner 的 AppKey |
| 11865 | `/api/p1/appkey/login` | Appkey not found | 见第 2 节错误码表 |
| 11864 | `/api/p1/appkey/login` | User not found | 见第 2 节错误码表 |
| 11866 | `/api/p1/appkey/login` | Account not found | 见第 2 节错误码表 |

> **注意：** 业务层 503 ≠ HTTP 503。前者是 HTTP 200 返回的 JSON 中 `code: 503`，表示权限不足（如上传文件到非 owner 项目）；后者是真正的服务不可用。脚本已自动识别并给出明确提示。

## Reference Files

- **API Key 获取指南** (`references/get-apikey.md`) — TicNote 平台上获取 API Key 的详细步骤
