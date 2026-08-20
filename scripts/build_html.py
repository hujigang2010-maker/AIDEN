# -*- coding: utf-8 -*-
"""生成浏览器可打开的一页总览 HTML。"""
import sys
from html import escape
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import content as C

OUT = Path(__file__).resolve().parents[1] / "exports"
OUT.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT / "课题一页总览.html"


def cards():
    chunks = []
    for t in C.TOPICS:
        chunks.append(
            f"""
<article class="card">
  <div class="id">{escape(t['id'])}</div>
  <h3>{escape(t['name'])}</h3>
  <p class="meta">{escape(t['track'])} · {escape(t['status'])} · {escape(t['quota'])} · {escape(t['weeks'])}</p>
  <p>{escape(t['one_liner'])}</p>
  <p class="muted">适合：{escape(t['major'])}</p>
</article>"""
        )
    return "\n".join(chunks)


def track_html():
    parts = []
    for tr in C.TRACKS:
        ids = "、".join(tr["ids"])
        parts.append(
            f"<div class='track'><h3>{escape(tr['name'])}</h3>"
            f"<p>{escape(tr['blurb'])}</p><p class='ids'>{escape(ids)}</p></div>"
        )
    return "\n".join(parts)


def principles():
    return "\n".join(f"<li>{escape(x)}</li>" for x in C.PRINCIPLES)


def build():
    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{escape(C.TITLE)}</title>
  <style>
    :root {{
      --navy: #0B3D5C;
      --teal: #1A7A6D;
      --gold: #C4A35A;
      --ink: #1A2A33;
      --muted: #5A6A72;
      --bg: #F4F8F7;
      --card: #ffffff;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
      color: var(--ink);
      background: var(--bg);
      line-height: 1.55;
    }}
    header {{
      background: linear-gradient(120deg, #082E45, #0B3D5C 55%, #1A7A6D);
      color: #fff;
      padding: 48px 8vw 40px;
    }}
    header .kicker {{ color: #D0E8E4; letter-spacing: .08em; font-size: 14px; }}
    h1 {{ margin: 12px 0 8px; font-size: 32px; font-weight: 700; }}
    header .sub {{ color: #E8F3F1; font-size: 18px; }}
    header .meta {{ margin-top: 18px; color: #C4A35A; font-size: 14px; }}
    main {{ padding: 28px 8vw 64px; max-width: 1200px; margin: 0 auto; }}
    h2 {{ color: var(--navy); border-left: 4px solid var(--teal); padding-left: 12px; }}
    .tracks {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; }}
    .track, .card {{
      background: var(--card);
      border-radius: 12px;
      padding: 16px 18px;
      box-shadow: 0 1px 0 rgba(11,61,92,.06);
    }}
    .grid {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 14px; }}
    .card .id {{
      display: inline-block;
      background: var(--navy);
      color: #fff;
      font-weight: 700;
      padding: 2px 8px;
      border-radius: 6px;
      font-size: 13px;
    }}
    .card h3 {{ margin: 8px 0 6px; font-size: 16px; color: var(--navy); }}
    .meta {{ color: var(--teal); font-size: 13px; margin: 0 0 8px; }}
    .muted {{ color: var(--muted); font-size: 13px; }}
    ul {{ padding-left: 1.2em; }}
    footer {{
      padding: 24px 8vw 40px;
      color: var(--muted);
      font-size: 13px;
    }}
    @media (max-width: 900px) {{
      .tracks, .grid {{ grid-template-columns: 1fr; }}
      h1 {{ font-size: 26px; }}
    }}
  </style>
</head>
<body>
  <header>
    <div class="kicker">{escape(C.SCHOOL)} · {escape(C.VERSION)} · 计划 {escape(C.RELEASE_DATE)} 发布</div>
    <h1>{escape(C.TITLE)}</h1>
    <p class="sub">{escape(C.SUBTITLE)}</p>
    <p class="meta">{escape(C.ORG_LINE)}<br/>对接：{escape(C.MENTOR)} · {escape(C.MENTOR_TITLE)}</p>
  </header>
  <main>
    <h2>先选赛道</h2>
    <div class="tracks">{track_html()}</div>
    <h2>十条已开工课题</h2>
    <div class="grid">{cards()}</div>
    <h2>原则</h2>
    <ul>{principles()}</ul>
    <h2>怎么报名</h2>
    <p>每人 1 个第一志愿 + 1 个备选，写 80 字以内「我能贡献什么」。老师编组后拉项目组工作群。</p>
    <p class="muted">{escape(C.WECHAT_CONTACT_NOTE)}</p>
  </main>
  <footer>{escape(C.DATE_STR)} · 配套 PPT / Word / Excel / 微信群发布稿 在同一 exports 目录</footer>
</body>
</html>
"""
    OUT_FILE.write_text(html, encoding="utf-8")
    print(f"已生成 {OUT_FILE}")
    return OUT_FILE


if __name__ == "__main__":
    build()
