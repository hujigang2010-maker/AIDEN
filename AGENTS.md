# AIDEN

## Cursor Cloud specific instructions

### Repository shape (important)

This repository is a collection of **independent, one-off projects, each living on its
own `cursor/*` branch**. The `main` branch is intentionally almost empty (just this file
and `README.md`) — there is **no single application on `main`**. Do not expect a
`package.json` or build tooling on `main`.

Branch contents fall into a few categories:

- **Document / deck deliverables** (most branches): Markdown, `.docx`, `.pptx`, images,
  `.zip`, etc. These have no build step — open the files directly.
- **Static HTML sites**: e.g. `cursor/wujiaochang-competitor-map-8ada`,
  `cursor/semir-international-ops-center-site-c921`,
  `cursor/shanghai-commercial-report-cc11`. Serve with any static server
  (`python3 -m http.server`) from the branch root.
- **Node / Vite web app**: `cursor/wechat-mp-editor-b201` ("MP Typer", a WeChat Official
  Account Markdown editor). This is the only branch with a `package.json`, and its Vite
  build is what is published to the `gh-pages` branch / GitHub Pages.

### Working with a specific project branch

Because each project is on a different branch, the cleanest way to work on one without
disturbing `main` is a git worktree:

```bash
git fetch origin <branch>
git worktree add ../work-<name> origin/<branch>
cd ../work-<name>
```

### Running the Vite app (`cursor/wechat-mp-editor-b201`)

Node 22 + npm are preinstalled. From a worktree of that branch:

```bash
npm install
npm run dev        # http://localhost:5173  (add -- --host to expose)
npm run lint       # eslint
npm run build      # tsc -b && vite build  -> dist/
npm run preview    # serve dist/ on http://localhost:4173
```

Build flavors (see that branch's README): `BASE_PATH=/AIDEN/ npm run build` for a
GitHub Pages sub-path, and `SINGLE_FILE=1 npm run build` for a self-contained
`dist/index.html`.

The app is pure front-end (no backend, no env vars, no secrets) and persists content to
`localStorage`.
