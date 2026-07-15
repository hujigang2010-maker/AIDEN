# 投资领域交流探讨 · 2026-05-17 整理输出

本目录基于 2026-05-17 的"投资领域交流探讨"会议纪要整理生成。

## 一、文件清单

| 文件 | 说明 |
|---|---|
| `01_说话人1_要点摘要.md` | **说话人 1** 的要点摘要——聚焦关注股票、原因、方法论 |
| `02_说话人1_思维导图.md` | **说话人 1** 的思维导图（Markmap 大纲 + Mermaid mindmap 双格式） |
| `03_说话人1_投资观点.pptx` | **说话人 1** 的 PPT（15 页，宽屏 16:9，深蓝金主题；含「三安光电三重打击」案例深读） |
| `04_其他说话人_交流纪要.md` | 其余说话人的独立纪要；与说话人 1 的交织/补充已以「📎 与说话人 1 交集」标注 |
| `05_说话人1_学生听课注释.md` | **学生听课式注释本**——对说话人 1 的每个论点做课堂笔记式批注（含三安光电三重打击深读、反方观点、风险提示、待查项、复习卡片） |
| `投资纪要_2026-05-17.zip` | 上述所有文件的一键打包，便于离线下载 |

## 二、下载方式

### 方式 1 · 一键下载整包（推荐）
- **GitHub Web**：进入 PR 页面 → `outputs/` 目录 → 点击 `投资纪要_2026-05-17.zip` → 右上角 **Download raw file**
- **直链**（替换为你实际的 PR 分支）：
  ```
  https://raw.githubusercontent.com/<owner>/<repo>/cursor/meeting-summary-investment-aa5a/outputs/投资纪要_2026-05-17.zip
  ```
- 本仓库的直链：
  ```
  https://github.com/hujigang2010-maker/AIDEN/raw/cursor/meeting-summary-investment-aa5a/outputs/投资纪要_2026-05-17.zip
  ```

### 方式 2 · 单文件下载
在 PR 中点开任意文件，右上角点击 **Download raw file**，或使用直链：
```
https://github.com/hujigang2010-maker/AIDEN/raw/cursor/meeting-summary-investment-aa5a/outputs/<文件名>
```
PPT 用此方式直接下载 `.pptx` 二进制。

### 方式 3 · 命令行下载（git）
```bash
git clone -b cursor/meeting-summary-investment-aa5a https://github.com/hujigang2010-maker/AIDEN.git
cd AIDEN/outputs
```

### 方式 4 · 命令行下载（wget / curl 单文件）
```bash
# 整包
wget https://github.com/hujigang2010-maker/AIDEN/raw/cursor/meeting-summary-investment-aa5a/outputs/投资纪要_2026-05-17.zip

# 单 PPT
curl -L -o 说话人1_投资观点.pptx \
  https://github.com/hujigang2010-maker/AIDEN/raw/cursor/meeting-summary-investment-aa5a/outputs/03_说话人1_投资观点.pptx
```

### 方式 5 · 一次性重新打包
```bash
cd outputs
zip -r 投资纪要_2026-05-17.zip . -x "*.zip"
```

## 三、文档使用建议

1. 先看 `01_要点摘要.md`：把握"说话人 1"的全部观点骨架（10 分钟）。
2. 看 `02_思维导图.md`：在脑中建立"主线 + 七大赛道 + 重点股票"的网状结构（5 分钟）。
3. 打开 `03_投资观点.pptx`：用于汇报 / 团队分享 / 复习（5 分钟扫一遍）。
4. 深读 `05_学生听课注释.md`：补全你需要的产业知识、风险提示、待查项；这是**真正能上场用**的版本。
5. 最后看 `04_其他说话人_交流纪要.md`：补全与会者的合作、AI 工具实战、行业生态等"配菜"。

## 四、思维导图渲染说明

- **Markmap**：复制 `02_说话人1_思维导图.md` 中 "A. Markmap" 部分到 https://markmap.js.org/repl
- **Mermaid**：在 GitHub / Typora / Obsidian / Notion 等支持的环境直接渲染。

## 五、重新生成 PPT

```bash
python3 scripts/build_ppt.py
```
依赖：`python-pptx`（`pip install python-pptx`）。

---

**免责声明**：本目录的所有内容仅供个人学习与团队内部交流，**不构成任何投资建议**。涉及个股的判断保留了演讲者原意，并在 `05_学生听课注释.md` 中以学生视角做了风险与反方观点的补充，请独立判断。
