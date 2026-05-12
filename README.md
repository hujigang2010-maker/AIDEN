# 森马（上海）国际运营中心 · Semir Group's Global Headquarter

> A single-page project showcase for **森马国际运营中心** — the upcoming "Z·世代潮玩社交主场" anchored next to Shanghai Metro Line 15 元江路 TOD in Minhang's 大零号湾 innovation corridor.

This site is a static, self-contained reimagining of the project's 28-slide investment / partnership deck. It walks visitors through the four chapters of the original presentation:

| #  | Chapter        | English                  |
| -- | -------------- | ------------------------ |
| 01 | 项目概况篇      | Project Overview         |
| 02 | 产业规划篇      | Industry Planning        |
| 03 | 商业规划篇      | Commercial Planning      |
| 04 | 运营合作篇      | Operations & Cooperation |

## Highlights

- **Hero stage** featuring the project's three strategic pillars (新兴载体 · 示范枢纽 · 融合标杆).
- **Macro location & traffic** — distance/time to 紫竹高新区, 申嘉湖高速, 浦东/虹桥机场, 大学城 and the 15 号线 元江路 TOD (5–7 万人次/日).
- **Surroundings stats** — 24 万 居住人口 · 12 万 产业办公人口 within a 15-minute drive.
- **Buildings 1–6** detailed floor-by-floor program (零售 / 总部 / Livehouse / 酒店 / 潮玩艺术中心 / 直播 / 动漫书店 / 餐饮).
- **Industry pyramid** visualising the 10/10/20/40/20 % distribution of 头部央企 → 共享配套 → 中型 → 小型 → 服务机构.
- **Dual licenses**: 潮玩次元商业专委会 + AI 潮玩产业基地.
- **6 产业配套平台** — 选品中心 / 代运营物流 / 共享直播 / AI 共享设计 / AI 共享打样 / 潮玩产业展厅.
- **5 主力业态** — 动漫潮玩谷街区 / IP 潮玩选品 & 仓储零售 / 潮玩艺术中心 / 二次元 Livehouse / 动漫主题书店.
- **3 旗舰活动** — 潮玩集市 / 动漫新品首发会 / 全国潮玩设计大赛.
- **Partner matrix** grouped by IP 潮玩/二次元, 产业合作机构, 周边产业生态.

## Design

- Dark, "潮玩元宇宙" aesthetic with magenta / violet / cyan gradient accents.
- Animated glow orbs, grid background, mouse-parallax in the hero.
- Scroll-revealed cards, animated counters, sticky chapter rail (right side).
- Fully responsive down to mobile (`<720 px`).
- No build step, no dependencies — only Google Fonts (Noto Sans SC + Inter).

## Project Structure

```
.
├── index.html                # Single-page layout (all four chapters)
├── assets/
│   ├── css/styles.css        # Theme, layout, components, responsive rules
│   └── js/main.js            # Nav scroll, chapter rail, reveals, counters, parallax
├── scripts/
│   ├── build_ppt.py          # Generator: 创意 24 页 PPT
│   └── build_docx.py         # Generator: 项目介绍与解读 Word
├── docs/
│   ├── README.md             # Document folder index
│   ├── 森马国际运营中心-项目介绍.pptx           # Creative presentation deck
│   └── 森马国际运营中心-项目介绍与解读.docx     # Interpretation Word doc
└── README.md
```

## Documents (PPT + Word)

In addition to the web showcase, the `docs/` folder ships two distributable assets generated from `scripts/`:

- **`森马国际运营中心-项目介绍.pptx`** — 24-slide 16:9 deck in a dark "潮玩元宇宙" style (cover, contents, 4 chapter dividers, content slides, core-value summary, vision quote, THANKS).
- **`森马国际运营中心-项目介绍与解读.docx`** — long-form Word document with 8 chapters: Background, Overview, Industry, Commercial, Operations, Value Assessment, Risks & Recommendations, Closing — each major section accompanied by a colored "解读 / Interpretation" callout.

Both files are regenerated with:

```bash
pip install python-pptx python-docx
python3 scripts/build_ppt.py
python3 scripts/build_docx.py
```

## Run Locally

The site is purely static — open `index.html` directly, or serve it with any static file server:

```bash
# Python 3
python3 -m http.server 8080

# Node.js (npx)
npx serve .
```

Then visit <http://localhost:8080>.

## Browser Support

Targets modern evergreen browsers (Chromium, Firefox, Safari, Edge). Uses CSS custom properties, `backdrop-filter`, `IntersectionObserver`, and `prefers-reduced-motion` for accessibility.
