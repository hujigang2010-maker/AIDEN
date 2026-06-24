# AGENTS.md

## Cursor Cloud specific instructions

### Repository layout (important)
This repo is a multi-branch artifact repository. The `main` branch is an empty
placeholder (`README.md` contains only `# AIDEN`) — there is **no application
code, `package.json`, or service on `main`**. Each `cursor/*` feature branch is
an independent deliverable; most are static document/content bundles (`.md`,
`.docx`, `.xlsx`, `.pptx`, images) with nothing to run.

The only runnable application lives on branch `cursor/wechat-mp-editor-b201`:
**MP Typer (公众号排版助手)** — a pure front-end Markdown editor (Vite + React 19 +
TypeScript) that renders inline-styled HTML for pasting into WeChat Official
Accounts. There is no backend, database, or external dependency; it is a 100%
client-side SPA.

### Toolchain
- Node.js 22 + npm (preinstalled). The app uses Vite 8 and React 19.

### Running / building / linting the app (on `cursor/wechat-mp-editor-b201`)
Commands are defined in that branch's `package.json` `scripts`:
- Install deps: `npm install`
- Dev server: `npm run dev` → serves at `http://localhost:5173/`
- Build: `npm run build` (`tsc -b && vite build`, output in `dist/`)
- Lint: `npm run lint` (ESLint)
- Preview built output: `npm run preview`
- There is **no test command/framework** defined (no test runner in scripts/deps).

### Notes / gotchas
- Because `main` has no `package.json`, dependency install must be guarded; the
  startup update script only runs `npm install` when a `package.json` is present
  on the checked-out branch.
- To run the app without leaving your current branch, you can use a git worktree,
  e.g. `git worktree add /home/ubuntu/app-worktree cursor/wechat-mp-editor-b201`
  (worktrees cannot be created directly under `/`; use a path under `/home` or
  `/workspace`).
