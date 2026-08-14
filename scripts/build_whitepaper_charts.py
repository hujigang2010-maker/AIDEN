# -*- coding: utf-8 -*-
"""《2026 谷歌上海开发者大会白皮书》图表生成脚本。

数据来源见 whitepaper/data/ 与正文附录。图表仅基于公开报道可核验口径绘制，
不虚构未公布的分会场人数、营收或未披露的议程席次。
"""
import json
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import FancyBboxPatch

FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
]
FONT_FAMILY = "WenQuanYi Micro Hei"
for f in FONT_CANDIDATES:
    if os.path.exists(f):
        font_manager.fontManager.addfont(f)
        if "wqy" in f:
            FONT_FAMILY = "WenQuanYi Micro Hei"
            break
plt.rcParams.update({
    "font.family": FONT_FAMILY,
    "axes.unicode_minus": False,
    "figure.dpi": 150,
    "savefig.dpi": 150,
    "axes.edgecolor": "#C9CFD8",
    "axes.linewidth": 0.8,
    "axes.titlesize": 14,
    "axes.titleweight": "bold",
    "axes.labelsize": 11,
    "xtick.labelsize": 10,
    "ytick.labelsize": 10,
    "legend.fontsize": 10,
})
BLUE = "#0E4E9B"
RED = "#C8102E"
GRAY = "#8A93A3"
PALETTE = ["#0E4E9B", "#2E7CC3", "#5FA8DC", "#93C6E8", "#C8102E",
           "#E8734A", "#F2A65A", "#6B7B8C", "#3AA17E", "#8B6BB8"]
OUT = "/workspace/whitepaper/assets/charts"
DATA = "/workspace/whitepaper/data"
os.makedirs(OUT, exist_ok=True)


def note(fig, text):
    fig.text(0.01, 0.012, text, fontsize=8.5, color=GRAY)


def save(fig, name):
    fig.savefig(os.path.join(OUT, name), bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("saved", name)


# ============ 图1 历年举办地 ============
hist = json.load(open(os.path.join(DATA, "io_connect_history.json"), encoding="utf-8"))
fig, ax = plt.subplots(figsize=(10.2, 3.6))
xs = [e["year"] for e in hist["editions"]]
colors = [BLUE if e["city"] == "上海" else RED for e in hist["editions"]]
ax.scatter(xs, [1] * len(xs), s=420, c=colors, zorder=3)
ax.plot(xs, [1] * len(xs), color="#C9CFD8", lw=3, zorder=1)
for e in hist["editions"]:
    ax.text(e["year"], 1.18, e["city"], ha="center", fontsize=13, color=BLUE, fontweight="bold")
    ax.text(e["year"], 0.78, e["dates"], ha="center", fontsize=9, color=GRAY)
    ax.text(e["year"], 0.62, e["note"], ha="center", fontsize=8.5, color="#5B616B")
ax.set_xlim(2022.4, 2026.6)
ax.set_ylim(0.45, 1.45)
ax.axis("off")
ax.set_title("图1  Google I/O Connect China 历年举办地（2023—2026）")
note(fig, "数据来源：动点科技、腾讯新闻、Google for Developers、IT之家、Google 黑板报公开报道。红色为北京站。")
save(fig, "chart01_history.png")


# ============ 图2 2026 上海会展时间轴 ============
events = [
    ("WAIC 2026", "7.17–20", "世博中心等四馆", 0),
    ("I/O Connect China", "8.12–13", "上海世博中心", 1),
    ("出海加速器开营", "9月上旬", "线上线下结合", 2),
    ("GDG DevFest 上海", "11月初", "上海·约3000人", 3),
]
fig, ax = plt.subplots(figsize=(10.2, 3.8))
ax.plot([0, 3], [0.5, 0.5], color="#C9CFD8", lw=4, zorder=1)
for name, date, loc, x in events:
    c = RED if x == 1 else BLUE
    ax.scatter([x], [0.5], s=280, c=c, zorder=3)
    ax.text(x, 0.92, name, ha="center", fontsize=12, color=c, fontweight="bold")
    ax.text(x, 0.72, date, ha="center", fontsize=10, color="#333")
    ax.text(x, 0.18, loc, ha="center", fontsize=9, color=GRAY)
ax.set_xlim(-0.45, 3.45)
ax.set_ylim(-0.05, 1.2)
ax.axis("off")
ax.set_title("图2  2026 年上海开发者/AI 会展时间轴")
note(fig, "数据来源：WAIC 官方日程、IT之家、Google for Startups Accelerator China、GDG Shanghai。")
save(fig, "chart02_calendar.png")


# ============ 图3 I/O 2026 关键指标 ============
metrics = json.load(open(os.path.join(DATA, "io2026_metrics.json"), encoding="utf-8"))
fig, axes = plt.subplots(2, 4, figsize=(11.2, 5.2))
for ax, m in zip(axes.ravel(), metrics["metrics"]):
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    rect = FancyBboxPatch((0.06, 0.08), 0.88, 0.84, boxstyle="round,pad=0.02,rounding_size=0.08",
                          facecolor="#F4F7FB", edgecolor="#D5DEE8", linewidth=1)
    ax.add_patch(rect)
    v = m["value"]
    label = f"{v:,.0f}" if v >= 100 else f"{v:g}"
    ax.text(0.5, 0.68, label, ha="center", va="center", fontsize=20,
            color=BLUE, fontweight="bold")
    ax.text(0.5, 0.42, m["unit"], ha="center", fontsize=11, color=RED)
    ax.text(0.5, 0.24, m["name"], ha="center", fontsize=10, color="#1A1A1A")
fig.suptitle("图3  Google I/O 2026 主旨演讲披露的关键规模指标", fontsize=14, fontweight="bold", y=0.98)
note(fig, "数据来源：Sundar Pichai, I/O 2026 opening keynote, 2026-05-19。资本开支取 180–190 billion USD 区间中值，折合约 1,850 亿美元。")
plt.tight_layout(rect=[0, 0.04, 1, 0.94])
save(fig, "chart03_io_metrics.png")


# ============ 图4 全栈 AI 四层 ============
layers = [
    ("产品与平台层", "Search / Gemini App / Android / Chrome / Workspace / Play"),
    ("模型与智能体层", "Gemini 3.5 · Gemini Omni · Gemma 4 · Spark · Antigravity"),
    ("开发者工具层", "Android Studio · Firebase AI Logic · Cloud Run · ADK · AI Studio"),
    ("算力与硅基层", "TPU 8t（训练）· TPU 8i（推理）· 跨站点超百万 TPU 训练集群"),
]
fig, ax = plt.subplots(figsize=(10.4, 4.6))
ax.set_xlim(0, 10)
ax.set_ylim(0, 5)
ax.axis("off")
for i, (title, body) in enumerate(layers):
    y = 3.7 - i * 1.05
    c = PALETTE[i]
    rect = FancyBboxPatch((0.4, y), 9.2, 0.9, boxstyle="round,pad=0.02,rounding_size=0.08",
                          facecolor=c, edgecolor="none")
    ax.add_patch(rect)
    ax.text(0.7, y + 0.55, title, color="white", fontsize=13, fontweight="bold", va="center")
    ax.text(0.7, y + 0.25, body, color="white", fontsize=10, va="center")
ax.set_title("图4  Google 全栈 AI 架构（I/O 2026 → I/O Connect China 的技术底盘）")
note(fig, "整理自 I/O 2026 主旨演讲、Android 开发者博客、Google Cloud I/O 要点与中国站官方表述。")
save(fig, "chart04_fullstack.png")


# ============ 图5 中国站四大板块 ============
fig, ax = plt.subplots(figsize=(8.6, 4.8))
blocks = ["AI / 智能体", "Android", "Chrome", "Cloud"]
vals = [40, 25, 15, 20]  # 示意权重：报名文案四板块并列，AI 为叙事主轴
explode = (0.04, 0, 0, 0)
wedges, texts, autotexts = ax.pie(
    vals, labels=blocks, autopct=lambda p: f"{p:.0f}%",
    colors=PALETTE[:4], explode=explode, startangle=90,
    pctdistance=0.62, textprops={"fontsize": 11})
for t in autotexts:
    t.set_color("white")
    t.set_fontweight("bold")
ax.set_title("图5  中国站公开议程的四大技术板块（叙事权重示意）")
note(fig, "说明：官方报名文案并列 AI、Android、Chrome、Cloud；饼图按主旨叙事权重示意，非官方席次统计。")
save(fig, "chart05_tracks.png")


# ============ 图6 四大支柱 ============
pillars = [
    ("本土社区连接", "GDG / DevFest\n北上深办公室网络"),
    ("技术导师指导", "GDE / Office Hours\n现场工作坊与专家对话"),
    ("出海创业加速", "Google for Startups\nAccelerator China"),
    ("中文技术文档", "本地化文档与\n开发者中文站点"),
]
fig, ax = plt.subplots(figsize=(10.4, 3.6))
ax.set_xlim(0, 10)
ax.set_ylim(0, 3.2)
ax.axis("off")
for i, (t, b) in enumerate(pillars):
    x = 0.35 + i * 2.45
    rect = FancyBboxPatch((x, 0.45), 2.25, 2.2, boxstyle="round,pad=0.04,rounding_size=0.1",
                          facecolor="#F4F7FB", edgecolor=BLUE, linewidth=1.4)
    ax.add_patch(rect)
    ax.scatter([x + 1.12], [2.22], s=520, c=BLUE, zorder=3)
    ax.text(x + 1.12, 2.22, str(i + 1), ha="center", va="center", color="white",
            fontsize=14, fontweight="bold")
    ax.text(x + 1.12, 1.55, t, ha="center", fontsize=12, color=BLUE, fontweight="bold")
    ax.text(x + 1.12, 0.95, b, ha="center", fontsize=9.5, color="#5B616B")
ax.set_title("图6  Google 连接中国创新与全球规模的四大支柱")
note(fig, "数据来源：Google 黑板报《AI 全栈赋能，Google 助力中国开发者实现全球可持续增长》（2026-08-12）。")
save(fig, "chart06_pillars.png")


# ============ 图7 Gemma 四赛道 ============
fig, ax = plt.subplots(figsize=(9.6, 4.2))
tracks = ["A AI Agent", "B 多模态", "C 端侧 AI", "D 社会向善"]
# 按项目名称/官网描述做场景归类后的赛道示意计数（见 gemma4_finalists.json）
counts = [6, 1, 3, 5]
bars = ax.barh(tracks[::-1], counts[::-1], color=[PALETTE[i] for i in range(3, -1, -1)], height=0.55)
for b, v in zip(bars, counts[::-1]):
    ax.text(b.get_width() + 0.08, b.get_y() + b.get_height() / 2,
            f"{v} 支（示意归类）", va="center", fontsize=10)
ax.set_xlim(0, 8)
ax.set_xlabel("入围项目数（按公开项目名与赛道定义归类）")
ax.set_title("图7  Gemma 4 开发者大赛总决赛入围项目的赛道结构")
note(fig, "入围名单来自 hackathon.googdg.cn；赛道归属由课题组按项目名称与官方赛道定义归类，非官方公布分赛道名单。")
save(fig, "chart07_gemma_tracks.png")


# ============ 图8 入围项目场景 ============
from collections import Counter
finalists = json.load(open(os.path.join(DATA, "gemma4_finalists.json"), encoding="utf-8"))["finalists"]
cnt = Counter(x["scene"] for x in finalists)
items = sorted(cnt.items(), key=lambda kv: kv[1])
fig, ax = plt.subplots(figsize=(9.4, 4.4))
names = [k for k, _ in items]
vals = [v for _, v in items]
colors = [RED if n in ("养老", "医疗", "社会向善") else BLUE for n in names]
bars = ax.barh(names, vals, color=colors, height=0.58)
for b, v in zip(bars, vals):
    ax.text(b.get_width() + 0.05, b.get_y() + b.get_height() / 2, str(v), va="center")
ax.set_xlim(0, 4.2)
ax.set_title("图8  Gemma 4 总决赛入围项目的应用场景分布（n=15）")
note(fig, "场景标签由课题组根据官网项目名称归纳。红色强调与住房/社区公共服务更近的养老、医疗与社会向善方向。")
save(fig, "chart08_gemma_scenes.png")


# ============ 图9 加速器时间轴 ============
fig, ax = plt.subplots(figsize=(10.6, 3.4))
steps = [
    ("4.23", "报名开始"),
    ("6.14", "报名截止"),
    ("5–7月", "线上面试"),
    ("8月上旬", "确认入营"),
    ("9月上旬", "开营"),
    ("9–12月", "课程辅导"),
    ("2027.1/3", "展示日"),
]
ax.plot(range(len(steps)), [1] * len(steps), color="#C9CFD8", lw=3, zorder=1)
for i, (d, n) in enumerate(steps):
    c = RED if i == 3 else BLUE
    ax.scatter([i], [1], s=220, c=c, zorder=3)
    ax.text(i, 1.22, d, ha="center", fontsize=10, color=c, fontweight="bold")
    ax.text(i, 0.72, n, ha="center", fontsize=9.5, color="#333")
ax.set_xlim(-0.5, 6.5)
ax.set_ylim(0.5, 1.45)
ax.axis("off")
ax.set_title("图9  2026 谷歌出海创业加速器关键节点")
note(fig, "数据来源：startup.googlecnapps.cn/accelerator/ 与 GFSA CN 选拔条款（2026）。红色节点与 I/O Connect China 同期。")
save(fig, "chart09_accelerator.png")


# ============ 图10 WAIC vs I/O Connect 对照 ============
fig, ax = plt.subplots(figsize=(10.6, 4.8))
ax.axis("off")
ax.set_xlim(0, 10)
ax.set_ylim(0, 6)
headers = [(0.4, "维度"), (2.4, "WAIC 2026"), (6.2, "I/O Connect China 2026")]
for x, t in headers:
    ax.text(x, 5.5, t, fontsize=12, color=BLUE, fontweight="bold")
rows = [
    ("时间 / 地点", "7.17–20 · 世博四馆", "8.12–13 · 世博中心"),
    ("规格", "全球治理 + 产业博览", "开发者技术连接站"),
    ("规模口径", "963 家参展 / 175 场论坛", "近 2,000 名开发者"),
    ("技术主轴", "具身智能 + 大模型落地", "智能体全栈 + 出海工具链"),
    ("空间含义", "园区/厂房/算力密度", "会展枢纽 + 人才与出海接口"),
    ("本中心文稿", "FDU-HPRC-WP-2026-02", "FDU-HPRC-WP-2026-03"),
]
for i, (a, b, c) in enumerate(rows):
    y = 4.7 - i * 0.72
    bg = "#F4F7FB" if i % 2 == 0 else "white"
    ax.add_patch(FancyBboxPatch((0.3, y - 0.22), 9.4, 0.62,
                                boxstyle="round,pad=0.01,rounding_size=0.04",
                                facecolor=bg, edgecolor="none"))
    ax.text(0.45, y + 0.08, a, fontsize=10.5, color="#1A1A1A", fontweight="bold")
    ax.text(2.4, y + 0.08, b, fontsize=10.5, color="#333")
    ax.text(6.2, y + 0.08, c, fontsize=10.5, color="#333")
ax.set_title("图10  同城相继的两场大会：产业博览 vs 开发者接口")
note(fig, "WAIC 口径见本中心研究文稿第二号；I/O Connect 口径见 Google 黑板报与 IT之家。")
save(fig, "chart10_waic_compare.png")


# ============ 图11 三城网络 ============
fig, ax = plt.subplots(figsize=(9.2, 4.4))
cities = ["北京\n（总部/政策/算力叙事）", "上海\n（会展主场/出海接口）", "深圳\n（硬件/端侧/制造）"]
roles = [3.2, 4.5, 3.0]
bars = ax.bar(["北京", "上海", "深圳"], [3.2, 4.5, 3.0], color=[PALETTE[1], RED, PALETTE[8]], width=0.55)
ax.set_ylim(0, 5.5)
ax.set_ylabel("城市角色强度（研究示意，1—5）")
for b, t in zip(bars, ["半决赛+Office", "大会主场+半决赛+Office", "半决赛+Office"]):
    ax.text(b.get_x() + b.get_width() / 2, b.get_height() + 0.15, t, ha="center", fontsize=9, color=GRAY)
ax.set_title("图11  Google 在中国的北—上—深开发者城市网络")
note(fig, "角色强度为课题组根据公开活动落点（大会主场、Gemma 半决赛、Google 办公室）所作示意评分，非官方指数。")
save(fig, "chart11_three_cities.png")


# ============ 图12 空间含义框架 ============
fig, ax = plt.subplots(figsize=(10.4, 4.8))
ax.set_xlim(0, 10)
ax.set_ylim(0, 5)
ax.axis("off")
boxes = [
    (0.4, 2.6, "会展空间", "世博中心连续承接\nWAIC 与 I/O Connect\n成为全球技术发布枢纽"),
    (3.5, 2.6, "人才住房", "年轻开发者、外籍参会者\n短期住宿 + 中长期租赁\n社区配套与职住平衡"),
    (6.6, 2.6, "出海办公", "国内总部 + 海外主体\n的双注册空间需求\n灵活工位与跨境时差"),
    (0.4, 0.35, "端侧与楼宇", "Edge AI / Android XR\n进入家庭与办公室内\n对户型与弱电提出新要求"),
    (3.5, 0.35, "社区智能", "养老防诈、急救、教育\n智能体嵌入既有社区\n而非另起数据中心"),
    (6.6, 0.35, "政策接口", "会展经济、人才公寓\n开源社区空间、跨境\n数据与住房制度衔接"),
]
for x, y, t, b in boxes:
    ax.add_patch(FancyBboxPatch((x, y), 2.9, 2.05, boxstyle="round,pad=0.04,rounding_size=0.08",
                                facecolor="#F4F7FB", edgecolor=BLUE, linewidth=1.2))
    ax.text(x + 1.45, y + 1.55, t, ha="center", fontsize=12, color=BLUE, fontweight="bold")
    ax.text(x + 1.45, y + 0.75, b, ha="center", fontsize=9, color="#5B616B")
ax.set_title("图12  开发者大会映射到城市空间的六类含义")
note(fig, "本图为课题组分析框架，用于把技术大会翻译为住房与城市政策可操作的观察维度。")
save(fig, "chart12_space_framework.png")
