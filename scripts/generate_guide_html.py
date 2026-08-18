#!/usr/bin/env python3
"""生成手机浏览器可直接打开的普通 HTML：不用 React、不用 ES module、不用 JS。"""

from __future__ import annotations

import html
import json
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "deliverables"
OUTS = [
    OUT_DIR / "hongfeng-guide.html",
    OUT_DIR / "青岛红枫路交通事故_处理总览.html",
]
WEB = ROOT / "web"


def load_data() -> dict:
    script = (
        "import { writeFileSync } from 'fs';"
        "import { META, TREE, WEEKS, RESULTS } from './src/data.ts';"
        "writeFileSync('/tmp/guide-data.json', JSON.stringify({META,TREE,WEEKS,RESULTS}));"
    )
    subprocess.run(
        ["node", "--experimental-strip-types", "-e", script],
        cwd=WEB,
        check=True,
    )
    return json.loads(Path("/tmp/guide-data.json").read_text(encoding="utf-8"))


def e(s: object) -> str:
    return html.escape(str(s or ""), quote=True)


def triple(card: dict) -> str:
    parts = [
        f'<div><h4>做什么</h4><p>{e(card.get("do"))}</p></div>',
        f'<div><h4>怎么推进</h4><p>{e(card.get("how"))}</p></div>',
        f'<div><h4>会得到什么</h4><p>{e(card.get("result"))}</p></div>',
    ]
    if card.get("fail"):
        parts.append(
            f'<div class="fail"><h4>过不了就</h4><p>{e(card.get("fail"))}</p></div>'
        )
    return f'<div class="triple">{"".join(parts)}</div>'


def tree_node(n: dict, open_ids: set[str], depth: int = 0) -> str:
    kids = n.get("children") or []
    tone = e(n.get("tone") or "plain")
    kicker = f'<b>{e(n.get("kicker"))}</b> ' if n.get("kicker") else ""
    body = f'<div class="card inner">{triple(n["card"])}</div>'
    if kids:
        body += '<ul class="kids">' + "".join(
            tree_node(c, open_ids, depth + 1) for c in kids
        ) + "</ul>"
    opened = " open" if n["id"] in open_ids or depth <= 1 else ""
    return (
        f'<li class="node {tone}">'
        f"<details{opened}>"
        f'<summary>{kicker}{e(n["title"])}</summary>'
        f"{body}"
        f"</details></li>"
    )


def collect_open(n: dict, open_ids: set[str], depth: int = 0) -> None:
    if depth <= 1 or n["id"] in {"weeks", "hospital"}:
        open_ids.add(n["id"])
    for c in n.get("children") or []:
        collect_open(c, open_ids, depth + 1)


def weeks_html(weeks: list[dict]) -> str:
    blocks = []
    for w in weeks:
        now = " now" if w.get("phase") == "now" else ""
        items = []
        for it in w["items"]:
            items.append(
                f'<div class="item"><h3>{e(it["title"])}</h3>'
                f'<p class="meta">{e(it["owner"])}</p>'
                f'{triple(it["card"])}</div>'
            )
        blocks.append(
            f'<article class="card week{now}">'
            f'<h2>{e(w["label"])}</h2>'
            f'<p class="meta">{e(w["dates"])}</p>'
            f'<p class="goal">目标：{e(w["goal"])}</p>'
            f'{"".join(items)}</article>'
        )
    return "".join(blocks)


def results_html(results: list[dict]) -> str:
    blocks = []
    for r in results:
        lis = "".join(f"<li>{e(x)}</li>" for x in r["items"])
        blocks.append(
            f'<article class="card"><h2>{e(r["title"])}</h2><ul>{lis}</ul></article>'
        )
    return "".join(blocks)


def build(data: dict) -> str:
    meta = data["META"]
    open_ids: set[str] = set()
    collect_open(data["TREE"], open_ids)
    pins = "".join(f"<li>{e(p)}</li>" for p in meta["pins"])
    tree = (
        '<ul class="kids root">'
        + tree_node(data["TREE"], open_ids)
        + "</ul>"
    )
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<title>{e(meta["title"])}</title>
<style>
:root {{
  --ink:#1c1915; --paper:#f4efe4; --panel:#fffaf2; --line:#d7ccb8;
  --navy:#1e3a5f; --brick:#9a3412; --gold:#b45309; --muted:#6b6256;
}}
* {{ box-sizing:border-box; }}
body {{
  margin:0;
  background:radial-gradient(900px 420px at 8% -8%, #fde8c8 0%, transparent 55%), var(--paper);
  color:var(--ink);
  font-family:-apple-system,BlinkMacSystemFont,"PingFang SC","Hiragino Sans GB","Noto Sans SC","Microsoft YaHei",sans-serif;
  line-height:1.55;
}}
.page {{ max-width:920px; margin:0 auto; padding:20px 16px 48px; }}
h1 {{ color:var(--navy); font-size:1.4rem; line-height:1.35; margin:0 0 8px; }}
h2 {{ color:var(--navy); margin:0 0 8px; font-size:1.15rem; }}
h3 {{ margin:0 0 4px; font-size:1rem; }}
h4 {{ margin:0 0 4px; font-size:.75rem; color:var(--navy); letter-spacing:.08em; }}
p {{ margin:0 0 8px; }}
.eyebrow {{ color:var(--gold); font-weight:700; letter-spacing:.1em; font-size:.75rem; margin:0 0 8px; }}
.sub,.meta,.hint {{ color:var(--muted); }}
.pins {{ display:grid; gap:8px; padding:0; margin:14px 0 0; list-style:none; }}
.pins li {{ background:var(--panel); border:1px solid var(--line); border-left:4px solid var(--navy); padding:10px 12px; }}
.tabs {{
  display:flex; gap:8px; margin:20px 0 14px; position:sticky; top:0;
  background:rgba(244,239,228,.96); padding:10px 0; z-index:5;
}}
.tabs label {{
  flex:1; text-align:center; border:1px solid var(--navy); background:#fff; color:var(--navy);
  padding:10px 6px; border-radius:999px; font-size:.9rem; cursor:pointer;
}}
.sr {{ position:absolute; left:-9999px; }}
#tab-tree:checked ~ .tabs label[for="tab-tree"],
#tab-weeks:checked ~ .tabs label[for="tab-weeks"],
#tab-end:checked ~ .tabs label[for="tab-end"] {{
  background:var(--navy); color:#fff;
}}
.panel {{ display:none; }}
#tab-tree:checked ~ #panel-tree,
#tab-weeks:checked ~ #panel-weeks,
#tab-end:checked ~ #panel-end {{ display:block; }}
.card {{ background:var(--panel); border:1px solid var(--line); border-radius:14px; padding:14px; margin:0 0 12px; }}
.card.inner {{ margin:10px 0 0; }}
.kids {{ margin:0; padding:0 0 0 8px; list-style:none; }}
.kids.root {{ padding:0; }}
.node {{ margin:8px 0; }}
details {{ border:1px solid var(--line); border-radius:12px; background:#fff; padding:4px 8px 8px; }}
summary {{ cursor:pointer; padding:10px 8px; font-weight:600; }}
.now > details {{ background:#fff1e6; }}
.ok > details {{ background:#e8f3ec; }}
.warn > details, .stop > details {{ background:#fde8e4; }}
.money > details {{ background:#f7ecd3; }}
.triple {{ display:grid; gap:8px; }}
.triple > div {{ border-top:1px dashed var(--line); padding-top:8px; }}
.fail {{ background:#fde8e4; border-radius:10px; padding:8px 10px; }}
.week.now {{ outline:3px solid var(--brick); outline-offset:2px; }}
.goal {{ background:#e8eef6; padding:8px 10px; border-radius:10px; font-weight:600; }}
.item {{ border-top:1px solid var(--line); margin-top:10px; padding-top:8px; }}
footer {{ margin-top:24px; color:var(--muted); font-size:.82rem; text-align:center; }}
.note {{ background:#fff; border:1px dashed var(--gold); border-radius:12px; padding:10px 12px; margin:0 0 16px; font-size:.9rem; }}
</style>
</head>
<body>
<div class="page">
  <p class="note">这是普通网页，点开就能看。不要用微信内置浏览器；请用 Safari、Chrome 或系统浏览器。点树枝左侧小三角可展开。</p>
  <p class="eyebrow">内部处理总览 · 点树枝展开</p>
  <h1>{e(meta["title"])}</h1>
  <p class="sub">{e(meta["subtitle"])}</p>
  <p class="meta">{e(meta["date"])}</p>
  <ul class="pins">{pins}</ul>

  <input class="sr" type="radio" name="tab" id="tab-tree" checked>
  <input class="sr" type="radio" name="tab" id="tab-weeks">
  <input class="sr" type="radio" name="tab" id="tab-end">
  <nav class="tabs">
    <label for="tab-tree">1. 思维树</label>
    <label for="tab-weeks">2. 每周做什么</label>
    <label for="tab-end">3. 会得到什么</label>
  </nav>
  <section class="panel" id="panel-tree">
    <p class="hint">先点树枝标题展开，里面就是这一枝的做 / 推 / 果。</p>
    {tree}
  </section>
  <section class="panel" id="panel-weeks">
    <p class="hint">本周红框。每一项都拆成做、推、果。</p>
    {weeks_html(data["WEEKS"])}
  </section>
  <section class="panel" id="panel-end">
    <p class="hint">过程结果可以本周勾；钱的结果要等认定书和评残，不要现在对外报价。</p>
    {results_html(data["RESULTS"])}
  </section>
  <footer>本文是家属内部梳理，不是医学鉴定、不是司法鉴定、不是律师函。对外以医院已审核报告、发票原件、交警认定书为准。</footer>
</div>
</body>
</html>
"""


def main() -> None:
    data = load_data()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    html_text = build(data)
    assert "青岛红枫路" in html_text
    assert "<script" not in html_text
    assert "胫骨远端" in html_text
    assert "人民医院" in html_text
    for out in OUTS:
        out.write_text(html_text, encoding="utf-8")
        print("wrote", out, "bytes", out.stat().st_size)


if __name__ == "__main__":
    main()
