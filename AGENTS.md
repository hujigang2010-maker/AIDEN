# AIDEN

## Cursor Cloud specific instructions

### Repository shape (important, non-obvious)
- The `main` branch is an (almost) empty placeholder — it contains only `README.md`. There is **nothing to run on `main`**.
- Real deliverables live on per-task `cursor/*` feature branches. Each branch is an independent product/output. To work on or run something, check out the relevant branch (e.g. via `git checkout` or a `git worktree`).

### Two kinds of deliverables
1. **`mp-typer` web app** (branch `cursor/wechat-mp-editor-b201`): a pure front-end React 19 + Vite SPA — a Markdown editor that produces inline-styled HTML for WeChat Official Accounts. No backend, database, or env vars required.
2. **Document-generation scripts** (most other `cursor/*` branches): Python 3 scripts using `python-pptx`, `python-docx`, `openpyxl`, and `lxml` that emit PPTX/DOCX/XLSX/PNG files. These are one-off batch scripts, not long-running services, and often write outputs to hard-coded `/workspace/...` paths.

### Web app (mp-typer) — run / lint / build
- Package manager is **npm** (`package-lock.json` present). Run `npm install` in the branch checkout before anything else.
- Dev server: `npm run dev` → Vite on `http://localhost:5173/`. This is the entire runtime; nothing else is needed.
- Lint: `npm run lint` (ESLint flat config). Build: `npm run build` (`tsc -b && vite build`).
- No automated test suite is defined (no `test` script / framework).
- Optional build-time env vars: `BASE_PATH` (sub-path deploy, e.g. `/AIDEN/`) and `SINGLE_FILE=1` (self-contained `index.html`).

### Python document-generation branches
- Required libraries (`python-pptx`, `python-docx`, `openpyxl`, `lxml`) are preinstalled by the startup update script, so `python3 scripts/<name>.py` works without extra setup. There is no `requirements.txt`; dependencies are implicit imports.

### Running an app while staying on `main`
- Because runnable code is on other branches, the cleanest way to run/demo without leaving your working branch is a linked worktree, e.g. `git worktree add /workspace/.worktree-mp-typer origin/cursor/wechat-mp-editor-b201`, then `npm install` + `npm run dev` inside it.
