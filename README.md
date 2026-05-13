# MP Typer · 公众号排版助手

一个纯前端的微信公众号 Markdown 排版工具，支持实时预览、多套主题、代码高亮，并可以**一键复制富文本**到微信公众号后台直接发布。

> 微信公众号编辑器会过滤 `class` 但保留 `style` 内联样式，因此本项目所有样式都会被内联到生成的 HTML 中，确保粘贴到公众号后样式不会丢失。

## ✨ 功能特性

- 📝 **Markdown 编辑器**：左侧编辑、右侧实时预览
- 🎨 **多套主题**：默认蓝 / 优雅绿 / 极客橙 / 少女粉，覆盖资讯、生活、科技、情感类公众号
- 🛠️ **快捷工具栏**：一键插入标题、加粗、引用、列表、链接、图片、代码块、表格
- ⌨️ **键盘快捷键**：`Ctrl/Cmd+B` 加粗、`Ctrl/Cmd+I` 斜体、`Tab` 缩进
- 🌈 **代码高亮**：内置 30+ 常见语言，颜色样式自动内联（公众号也能正常显示）
- 📋 **一键复制**：
  - 「复制到公众号」：复制带富文本的 HTML，直接粘贴到公众号正文
  - 「复制 HTML」：复制内联样式的 HTML 源码
- 💾 **自动保存**：内容和主题选择会保存在 `localStorage`，刷新不丢失
- 📁 **导入导出**：支持导入 `.md` 文件、导出当前文档
- 🛡️ **XSS 防护**：使用 DOMPurify 净化预览内容
- 📱 **响应式**：移动端自动切换为上下布局

## 🚀 快速开始

### 本地开发

```bash
npm install
npm run dev          # 打开 http://localhost:5173
```

### 本地构建并预览

```bash
npm run build        # 产物输出到 dist/
npm run preview      # 预览 dist/ 在 http://localhost:4173
```

### 部署到 GitHub Pages（自动）

项目已自带 `.github/workflows/deploy.yml`。把代码推到 `main` 分支后：

1. 打开 GitHub 仓库 → **Settings → Pages**
2. 「Build and deployment」选择 **GitHub Actions**
3. 等待 Actions 跑完，会得到一个 `https://<your-name>.github.io/<repo>/` 的访问地址

部署到子路径时，Vite 的 `base` 已在构建时通过环境变量 `BASE_PATH` 注入，workflow 会自动设置为 `/<repo>/`。

### 部署到根路径（如 Netlify / Vercel / 自有 nginx）

直接 `npm run build`，把 `dist/` 上传或托管即可。

### 部署到任意子路径

```bash
BASE_PATH=/my-sub-path/ npm run build
```

## 📖 使用说明

1. 在左侧编辑器写 Markdown，右侧实时看到公众号样式预览
2. 顶部「主题」下拉切换不同的视觉风格
3. 写完后点击右上角 **「复制到公众号」**
4. 打开微信公众平台 → 新建图文 → 在正文区域直接 `Ctrl/Cmd + V` 粘贴
5. 图片：插入 `![alt](https://图片URL)` 即可，建议先把图片上传到公众号图床或可访问的图床

> ⚠️ 公众号粘贴小贴士：建议在 PC 端的微信公众平台后台粘贴，使用 Chrome / Edge 等现代浏览器，效果最佳。

## 🏗️ 技术栈

- [Vite](https://vitejs.dev/) + [React 19](https://react.dev/) + TypeScript
- [marked](https://marked.js.org/) — Markdown 解析
- [highlight.js](https://highlightjs.org/) — 代码高亮
- [DOMPurify](https://github.com/cure53/DOMPurify) — HTML 净化

## 📂 项目结构

```
src/
├── components/
│   ├── Editor.tsx        # Markdown 文本编辑器
│   ├── Toolbar.tsx       # 编辑工具栏
│   └── Preview.tsx       # 预览面板
├── themes/               # 主题样式（全部使用内联样式）
│   ├── default.ts
│   ├── elegant.ts
│   ├── tech.ts
│   ├── pink.ts
│   └── types.ts
├── utils/
│   ├── render.ts         # Markdown → 内联样式 HTML
│   ├── inlineHljs.ts     # 把代码高亮 class 转成内联 style
│   └── clipboard.ts      # 富文本复制
├── App.tsx
├── App.css
└── sample.ts             # 示例内容
```

## 🎨 自定义主题

新增主题只需在 `src/themes/` 下新建一个文件，按照 `Theme` 接口定义所有标签的样式，
再在 `src/themes/index.ts` 中导出。所有样式必须是**可序列化的 CSS 属性键值对**，
项目会在渲染时拼接为 `style="..."` 内联到 HTML 上，以适配公众号编辑器。

## 📜 License

MIT
