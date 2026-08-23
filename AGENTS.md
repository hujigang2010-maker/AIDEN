# AIDEN

## Cursor Cloud 开发说明

### 显示语言（中文）

本仓库已配置为默认使用简体中文：

- `.vscode/extensions.json` 推荐安装 **Chinese (Simplified) Language Pack**
- `.vscode/settings.json` 将 `locale` 设为 `zh-cn`
- `.cursor/rules/zh-cn.mdc` 要求 AI 使用简体中文交流

若界面仍为英文，请在 Cursor 中执行：

1. 按 `Ctrl+Shift+P`（macOS：`Cmd+Shift+P`）打开命令面板
2. 输入并选择 **Configure Display Language**
3. 选择 **中文(简体)** / `zh-cn`
4. 按提示重启 Cursor

### 仓库结构（重要）

- `main` 分支是占位分支，主要存放共享配置，本身没有可运行的应用。
- 实际交付物在各自的 `cursor/*` 功能分支上，每个分支对应一个独立产品或文档任务。
- 要运行某个项目，请切换到对应分支（例如 `git checkout` 或 `git worktree`）。

### 两类交付物

1. **`mp-typer` 网页应用**（分支 `cursor/wechat-mp-editor-b201`）：纯前端 React 19 + Vite SPA，用于生成微信公众号内联样式 HTML 的 Markdown 编辑器。无需后端、数据库或环境变量。
2. **文档生成脚本**（大多数其他 `cursor/*` 分支）：使用 `python-pptx`、`python-docx`、`openpyxl`、`lxml` 的 Python 3 脚本，输出 PPTX/DOCX/XLSX/PNG 文件。这些是一次性批处理脚本，不是长期运行的服务。

### 网页应用（mp-typer）运行方式

- 包管理器为 **npm**（存在 `package-lock.json`）。进入分支目录后先执行 `npm install`。
- 开发服务器：`npm run dev`，Vite 监听 `http://localhost:5173/`。
- 代码检查：`npm run lint`（ESLint flat config）
- 构建：`npm run build`（`tsc -b && vite build`）
- 当前未定义自动化测试。
- 可选构建环境变量：`BASE_PATH`（子路径部署，例如 `/AIDEN/`）、`SINGLE_FILE=1`（生成单文件 `index.html`）。

### Python 文档生成分支

- 所需库（`python-pptx`、`python-docx`、`openpyxl`、`lxml`）通常已由启动脚本预装，可直接运行 `python3 scripts/<name>.py`。
- 仓库中没有 `requirements.txt`，依赖通过脚本导入隐式声明。

### 在 `main` 分支上运行其他分支的应用

- 可使用链接工作树，例如：
 `git worktree add /workspace/.worktree-mp-typer origin/cursor/wechat-mp-editor-b201`
- 然后在该目录中执行 `npm install` 和 `npm run dev`。

### Gemini 桌面端（分支 `cursor/gemini-desktop-login-fix-474b`）

- 纯 Electron 包装器，直接以顶层窗口打开 `https://gemini.google.com/app`。
- 先执行 `npm install`，再 `npm test`、`npm run dev`。
- 登录必须在应用窗口内完成；不要把 OAuth 丢到外部浏览器。
- Linux 快捷方式：`./scripts/install-desktop-shortcut.sh`。
