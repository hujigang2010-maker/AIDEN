# -*- coding: utf-8 -*-
"""《WAIC2026 人工智能产业空间白皮书》全套图表生成脚本。

数据来源：
- WAIC2026_全量资源整合总表.xlsx（参展商、行业、论坛、具身智能场景等）
- xsct.ai（XSCT Bench 大模型评测榜单，2026-08 抓取，whitepaper/data/xsct_models.json）
- watcha.cn 观猹（AI 产品社区数据，2026-08 抓取，whitepaper/data/watcha_*.json）
"""
import json
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

# ---------- 全局样式 ----------
FONT_CANDIDATES = [
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
]
FONT_FAMILY = "WenQuanYi Micro Hei"
for f in FONT_CANDIDATES:
    if os.path.exists(f):
        font_manager.fontManager.addfont(f)
        if "NotoSansCJK" in f:
            FONT_FAMILY = "Noto Sans CJK SC"
        elif "wqy" in f:
            FONT_FAMILY = "WenQuanYi Micro Hei"
plt.rcParams.update({
    "font.family": FONT_FAMILY,
    "axes.unicode_minus": False,
    "figure.dpi": 150,
    "savefig.dpi": 150,
    "axes.edgecolor": "#C9CFD8",
    "axes.linewidth": 0.8,
    "axes.titlesize": 15,
    "axes.titleweight": "bold",
    "axes.labelsize": 11,
    "xtick.labelsize": 10,
    "ytick.labelsize": 10,
    "legend.fontsize": 10,
})
BLUE = "#0E4E9B"      # 复旦蓝
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
    fig.savefig(os.path.join(OUT, name), bbox_inches="tight",
                facecolor="white")
    plt.close(fig)
    print("saved", name)


# ============ 图1 参展商八大类分布 ============
cats = [("机器人与智能硬件", 233), ("算力·芯片·基础设施", 226),
        ("大模型与生成式AI", 137), ("行业AI应用", 114),
        ("AI技术与算法", 89), ("企业服务与营销", 86),
        ("机构与平台", 67), ("智能网联汽车", 11)]
fig, ax = plt.subplots(figsize=(9, 4.8))
names = [c[0] for c in cats][::-1]
vals = [c[1] for c in cats][::-1]
colors = [RED if "机器人" in n or "算力" in n else BLUE for n in names]
bars = ax.barh(names, vals, color=colors, height=0.62)
for b, v in zip(bars, vals):
    ax.text(b.get_width() + 3, b.get_y() + b.get_height() / 2,
            f"{v} 家（{v/963*100:.1f}%）", va="center", fontsize=10)
ax.set_xlim(0, 285)
ax.set_title("图 1  WAIC 2026 参展商八大类行业分布（共 963 家）")
ax.spines[["top", "right"]].set_visible(False)
note(fig, "数据来源：WAIC2026 全量资源整合总表·行业统计（红色为占比前二的\u201c硬科技\u201d大类，合计 47.7%）")
save(fig, "chart01_exhibitor_industry.png")

# ============ 图2 细分领域 Top15 ============
subs = [("具身智能/人形机器人", 93), ("AI Agent/智能应用", 87),
        ("智能算力/数据中心/云", 78), ("工业机器人/智能制造", 68),
        ("企业软件/SaaS", 52), ("半导体/集成电路", 40),
        ("AI芯片/算力芯片", 39), ("科研院所/高校/行业组织", 39),
        ("数据智能/数据服务", 33), ("传感器/光电器件", 32),
        ("通信/网络/5G", 30), ("消费电子/智能硬件", 27),
        ("医疗健康/生物医药", 26), ("智能语音/NLP/翻译", 22),
        ("服务/消费机器人", 21)]
fig, ax = plt.subplots(figsize=(9, 6))
names = [s[0] for s in subs][::-1]
vals = [s[1] for s in subs][::-1]
colors = [RED if i >= 13 else BLUE for i in range(15)]
bars = ax.barh(names, vals, color=colors, height=0.62)
for b, v in zip(bars, vals):
    ax.text(b.get_width() + 1, b.get_y() + b.get_height() / 2,
            f"{v}", va="center", fontsize=10)
ax.set_xlim(0, 105)
ax.set_title("图 2  参展商细分领域 Top 15（39 个细分领域）")
ax.spines[["top", "right"]].set_visible(False)
note(fig, "数据来源：WAIC2026 全量资源整合总表·行业统计（红色为前两大细分：具身智能与 AI Agent）")
save(fig, "chart02_subsector_top15.png")

# ============ 图3 论坛赛道分布 ============
tracks = [("综合论坛", 45), ("大模型与AI基础", 35), ("产业与工业智能化", 32),
          ("算力与AI芯片", 22), ("教育与人才发展", 9), ("治理标准与政策", 7),
          ("前沿科技与探索", 5), ("机器人与具身智能", 5), ("金融与科技投资", 4),
          ("内容创意与AIGC", 4), ("能源与可持续发展", 3), ("医疗与生命科学", 3),
          ("女性与多元发展", 1)]
fig, ax = plt.subplots(figsize=(9, 5.6))
names = [t[0] for t in tracks][::-1]
vals = [t[1] for t in tracks][::-1]
colors = [RED if n in ("大模型与AI基础", "产业与工业智能化", "算力与AI芯片")
          else BLUE for n in names]
bars = ax.barh(names, vals, color=colors, height=0.62)
for b, v in zip(bars, vals):
    ax.text(b.get_width() + 0.5, b.get_y() + b.get_height() / 2,
            f"{v} 场（{v/175*100:.1f}%）", va="center", fontsize=9.5)
ax.set_xlim(0, 55)
ax.set_title("图 3  WAIC 2026 论坛赛道分布（共 175 场）")
ax.spines[["top", "right"]].set_visible(False)
note(fig, "数据来源：WAIC2026 全量资源整合总表·赛道分类统计（红色为\u201c技术—产业—底座\u201d三大主线赛道）")
save(fig, "chart03_forum_tracks.png")

# ============ 图4 论坛日期 × 场馆分布 ============
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4.2))
dates = ["7月17日", "7月18日", "7月19日", "7月20日"]
dvals = [27, 65, 64, 19]
bars = ax1.bar(dates, dvals, color=[BLUE, RED, RED, BLUE], width=0.55)
for b, v in zip(bars, dvals):
    ax1.text(b.get_x() + b.get_width() / 2, v + 1.2, f"{v} 场",
             ha="center", fontsize=10)
ax1.set_ylim(0, 75)
ax1.set_title("论坛日期分布")
ax1.spines[["top", "right"]].set_visible(False)
venues = [("世博中心", 91), ("世博展览馆", 28), ("西岸国际会展中心", 25),
          ("世博桐森酒店", 13), ("其他场馆", 8), ("世博滨江酒店", 5),
          ("张江科学会堂", 5)]
vn = [v[0] for v in venues][::-1]
vv = [v[1] for v in venues][::-1]
bars = ax2.barh(vn, vv, color=BLUE, height=0.6)
bars[-1].set_color(RED)
for b, v in zip(bars, vv):
    ax2.text(b.get_width() + 1, b.get_y() + b.get_height() / 2,
             f"{v}", va="center", fontsize=10)
ax2.set_xlim(0, 105)
ax2.set_title("论坛场馆分布")
ax2.spines[["top", "right"]].set_visible(False)
fig.suptitle("图 4  论坛的时间与空间分布：7 月 18—19 日承载 73.7%，世博中心承载 52%",
             fontsize=15, fontweight="bold", y=1.02)
note(fig, "数据来源：WAIC2026 全量资源整合总表·论坛日程（175 场）")
save(fig, "chart04_forum_dates_venues.png")

# ============ 图5 参展商注册地分布 ============
geo = [("上海", 271), ("北京", 189), ("广东（含深圳）", 134), ("浙江（含杭州）", 81),
       ("江苏", 45), ("香港", 20), ("天津", 11), ("山东", 9),
       ("四川", 8), ("湖北", 7)]
fig, ax = plt.subplots(figsize=(9, 4.8))
names = [g[0] for g in geo][::-1]
vals = [g[1] for g in geo][::-1]
colors = [RED if n in ("上海", "北京") else BLUE for n in names]
bars = ax.barh(names, vals, color=colors, height=0.62)
for b, v in zip(bars, vals):
    ax.text(b.get_width() + 3, b.get_y() + b.get_height() / 2,
            f"{v} 家（{v/963*100:.1f}%）", va="center", fontsize=10)
ax.set_xlim(0, 320)
ax.set_title("图 5  参展商注册地分布 Top 10（沪京双核占 47.8%）")
ax.spines[["top", "right"]].set_visible(False)
note(fig, "数据来源：WAIC2026 全量资源整合总表·参展商库（963 家，按注册地归并到省级；另有 48 家信息待补充）")
save(fig, "chart05_exhibitor_geo.png")

# ============ 图6 参展商融资阶段结构 ============
fin = [("已上市", 188), ("天使/种子轮", 85), ("A轮系列", 72), ("B轮系列", 41),
       ("战略融资", 19), ("D轮及以后", 15), ("C轮系列", 12),
       ("其他/政府资助等", 83), ("未披露", 448)]
fig, ax = plt.subplots(figsize=(9, 4.8))
names = [f[0] for f in fin][::-1]
vals = [f[1] for f in fin][::-1]
colors = []
for n in names:
    if n == "已上市":
        colors.append(RED)
    elif n in ("天使/种子轮", "A轮系列"):
        colors.append("#E8734A")
    elif n == "未披露":
        colors.append("#C9CFD8")
    else:
        colors.append(BLUE)
bars = ax.barh(names, vals, color=colors, height=0.62)
for b, v in zip(bars, vals):
    ax.text(b.get_width() + 4, b.get_y() + b.get_height() / 2,
            f"{v} 家（{v/963*100:.1f}%）", va="center", fontsize=10)
ax.set_xlim(0, 530)
ax.set_title("图 6  参展商融资阶段结构：\u201c已上市 + 早期\u201d哑铃型分布")
ax.spines[["top", "right"]].set_visible(False)
note(fig, "数据来源：WAIC2026 全量资源整合总表·参展商库\u201c最新融资轮次\u201d字段归类统计（红/橙为哑铃两端）")
save(fig, "chart06_financing.png")

# ============ 图7 大模型综合能力 Top 20 ============
models = json.load(open(os.path.join(DATA, "xsct_models.json")))
DOMESTIC_KEYS = ["doubao", "qwen", "kimi", "glm", "deepseek", "minimax",
                 "mimo", "step", "hy3", "hunyuan", "longcat", "ernie",
                 "spark", "yi-", "baichuan", "sense", "千问", "混元"]


def is_domestic(m):
    s = (m["slug"] + " " + m["name"]).lower()
    if any(k in s for k in ["claude", "gpt", "gemini", "gemma", "grok",
                            "nemotron", "llama", "mistral", "elephant",
                            "openai", "o3", "o4"]):
        return False
    return any(k in s for k in DOMESTIC_KEYS)


top20 = models[:20]
fig, ax = plt.subplots(figsize=(9.5, 7))
names = [f"{m['rank']}. {m['name']}" for m in top20][::-1]
vals = [m["overall"] for m in top20][::-1]
colors = [RED if is_domestic(m) else GRAY for m in top20][::-1]
bars = ax.barh(names, vals, color=colors, height=0.62)
for b, v in zip(bars, vals):
    ax.text(b.get_width() + 0.15, b.get_y() + b.get_height() / 2,
            f"{v:.1f}", va="center", fontsize=9.5)
ax.set_xlim(80, 93)
ax.set_title("图 7  大模型综合能力榜 Top 20（XSCT Bench，满分 100）")
ax.spines[["top", "right"]].set_visible(False)
from matplotlib.patches import Patch
ax.legend(handles=[Patch(color=RED, label="国产模型（13/20）"),
                   Patch(color=GRAY, label="海外模型（7/20）")],
          loc="lower right")
note(fig, "数据来源：xsct.ai（XSCT Bench）综合榜，2026-08 抓取；综合分 = 基础×30% + 进阶×40% + 困难×30%")
save(fig, "chart07_llm_top20.png")

# ============ 图8 得分 × 价格散点（性价比象限） ============
def out_price(m):
    """返回可用的输出定价；剔除未评测（0 分）与占位式异常定价。"""
    try:
        p = float(m["price_out"].replace("$", "").replace(",", ""))
    except ValueError:
        return None
    if p <= 0 or p > 100 or m["overall"] < 30:
        return None
    return p


fig, ax = plt.subplots(figsize=(9.5, 6.2))
xs, ys, cs, labels = [], [], [], []
for m in models:
    p = out_price(m)
    if p is None:
        continue
    xs.append(p)
    ys.append(m["overall"])
    cs.append(RED if is_domestic(m) else GRAY)
    labels.append(m["name"])
ax.scatter(xs, ys, c=cs, s=46, alpha=0.85, edgecolors="white", linewidths=0.6)
ax.set_xscale("log")
label_set = {"Anthropic: Claude Sonnet 4.6", "Claude Opus 4.6",
             "qwen3.6-plus-preview", "GLM-5v-turbo", "deepseek-v4-pro",
             "OpenAI: GPT-5.4", "deepseek-v4-flash", "kimi-k2.6",
             "Google: Gemini 3.1 Pro Preview", "qwen3.5-flash",
             "Google: Gemma 4 26B A4B"}
for x, y, l in zip(xs, ys, labels):
    if l in label_set:
        ha = "left" if "Claude" not in l and "GPT" not in l else "right"
        dx = 5 if ha == "left" else -5
        ax.annotate(l, (x, y), fontsize=8, xytext=(dx, 5), ha=ha,
                    textcoords="offset points", color="#444")
import numpy as np
ax.axhline(np.median(ys), color="#C9CFD8", lw=1, ls="--")
ax.axvline(np.median(xs), color="#C9CFD8", lw=1, ls="--")
ax.text(0.035, 89.3, "高分低价\n（最具性价比象限）", fontsize=10, color=RED,
        fontweight="bold")
ax.set_xlabel("输出价格（美元 / 百万 token，对数刻度）")
ax.set_ylabel("XSCT Bench 综合得分")
ax.set_title("图 8  大模型\u201c能力—价格\u201d分布：国产模型占据性价比象限")
ax.spines[["top", "right"]].set_visible(False)
ax.legend(handles=[Patch(color=RED, label="国产模型"),
                   Patch(color=GRAY, label="海外模型")], loc="lower right")
note(fig, "数据来源：xsct.ai 榜单公开定价（有公开付费定价且完成评测的模型，n=%d）；虚线为中位数" % len(xs))
save(fig, "chart08_llm_price_perf.png")

# ============ 图9 厂商模型梯队 ============
from collections import Counter
prov_cn = {"阿里云百炼": "阿里（通义千问）", "火山引擎": "字节（豆包）",
           "智谱开放平台": "智谱（GLM）", "月之暗面": "月之暗面（Kimi）",
           "Xiaomi MiMo": "小米（MiMo）", "MiniMax": "MiniMax",
           "深度求索": "深度求索（DeepSeek）", "腾讯混元": "腾讯（混元）",
           "阶跃星辰": "阶跃星辰（Step）"}
cnt = Counter()
for m in models:
    p = m["provider"]
    if p in prov_cn:
        cnt[prov_cn[p]] += 1
    elif p in ("OpenRouter", "PipeLLM"):
        s = (m["slug"] + m["name"]).lower()
        if "claude" in s:
            cnt["Anthropic（Claude）"] += 1
        elif "gpt" in s or "openai" in s:
            cnt["OpenAI（GPT）"] += 1
        elif "gemini" in s or "gemma" in s or "google" in s:
            cnt["Google（Gemini/Gemma）"] += 1
        elif "grok" in s:
            cnt["xAI（Grok）"] += 1
        elif "qwen" in s:
            cnt["阿里（通义千问）"] += 1
        elif "kimi" in s:
            cnt["月之暗面（Kimi）"] += 1
        elif "deepseek" in s:
            cnt["深度求索（DeepSeek）"] += 1
        elif "hy3" in s or "tencent" in s:
            cnt["腾讯（混元）"] += 1
        elif "step" in s:
            cnt["阶跃星辰（Step）"] += 1
        else:
            cnt["其他"] += 1
    else:
        cnt["其他"] += 1
items = [(k, v) for k, v in cnt.most_common() if k != "其他"][:12]
fig, ax = plt.subplots(figsize=(9, 5.4))
names = [i[0] for i in items][::-1]
vals = [i[1] for i in items][::-1]
DOM = ("阿里", "字节", "智谱", "月之暗面", "小米", "MiniMax", "深度求索",
       "腾讯", "阶跃")
colors = [RED if any(n.startswith(d) or d in n for d in DOM) else GRAY
          for n in names]
bars = ax.barh(names, vals, color=colors, height=0.62)
for b, v in zip(bars, vals):
    ax.text(b.get_width() + 0.2, b.get_y() + b.get_height() / 2,
            f"{v} 款", va="center", fontsize=10)
ax.set_xlim(0, 28)
ax.set_title("图 9  XSCT Bench 在评模型的厂商梯队（按上榜模型数）")
ax.spines[["top", "right"]].set_visible(False)
ax.legend(handles=[Patch(color=RED, label="国产厂商"),
                   Patch(color=GRAY, label="海外厂商")], loc="lower right")
note(fig, "数据来源：xsct.ai 综合榜 95 款在评模型，按厂商归并统计（聚合平台上的模型按原厂归属）")
save(fig, "chart09_llm_providers.png")

# ============ 图10 国产 vs 海外：能力与价格对比 ============
dom = [m for m in models if is_domestic(m)]
ovs = [m for m in models if not is_domestic(m)]


def top_avg(ms, k=10):
    ss = sorted(ms, key=lambda m: -m["overall"])[:k]
    return sum(m["overall"] for m in ss) / len(ss)


def avg_price(ms):
    ps = [out_price(m) for m in ms]
    ps = [p for p in ps if p is not None]
    return sum(ps) / len(ps)


fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4.4))
x = [0, 1]
v1 = [top_avg(dom), top_avg(ovs)]
bars = ax1.bar(["国产 Top10 均分", "海外 Top10 均分"], v1,
               color=[RED, GRAY], width=0.45)
for b, v in zip(bars, v1):
    ax1.text(b.get_x() + b.get_width() / 2, v + 0.15, f"{v:.1f}",
             ha="center", fontsize=11, fontweight="bold")
ax1.set_ylim(80, 92)
ax1.set_title("综合能力：Top10 平均分")
ax1.spines[["top", "right"]].set_visible(False)
v2 = [avg_price(dom), avg_price(ovs)]
bars = ax2.bar(["国产模型均价", "海外模型均价"], v2,
               color=[RED, GRAY], width=0.45)
for b, v in zip(bars, v2):
    ax2.text(b.get_x() + b.get_width() / 2, v + 0.2, f"${v:.2f}/M",
             ha="center", fontsize=11, fontweight="bold")
ax2.set_title("输出定价：全体在评模型平均")
ax2.spines[["top", "right"]].set_visible(False)
fig.suptitle("图 10  国产与海外大模型对比：能力差距收敛至 1 分以内，价格差距逾 4 倍",
             fontsize=15, fontweight="bold", y=1.03)
note(fig, "数据来源：xsct.ai 综合榜（2026-08）；价格为有公开付费定价模型的输出价均值（美元/百万 token）")
save(fig, "chart10_llm_cn_vs_global.png")

# ============ 图11 具身智能场景观察 ============
scenes = [("工业制造/仓储物流", 9, "搬运、分拣、上下料最拥挤；柔性装配更具挑战"),
          ("零售服务", 7, "便利店值守、药房取药、洗衣烘干、具身社区"),
          ("核心零部件", 2, "灵巧手 vs 抓夹：性能—成本—可靠性三角")]
fig, ax = plt.subplots(figsize=(9, 3.6))
names = [s[0] for s in scenes][::-1]
vals = [s[1] for s in scenes][::-1]
bars = ax.barh(names, vals, color=[GRAY, BLUE, RED], height=0.5)
for b, v, s in zip(bars, vals, [s[2] for s in scenes][::-1]):
    ax.text(b.get_width() + 0.15, b.get_y() + b.get_height() / 2,
            f"{v} 家 · {s}", va="center", fontsize=9.5)
ax.set_xlim(0, 16)
ax.set_title("图 11  H3 具身智能展厅：结构化观察样本的场景分布")
ax.spines[["top", "right"]].set_visible(False)
note(fig, "数据来源：WAIC2026 全量资源整合总表·具身智能场景观察（H3 展厅超 150 家企业中的 18 个结构化样本）")
save(fig, "chart11_embodied_scenes.png")

# ============ 图12 观猹热门产品类别分布 ============
wcat = [("效率工具", 18), ("通用助手", 15), ("其他类型", 14), ("编程开发", 6),
        ("图像生成", 6), ("Agent 构建", 4), ("知识管理", 4), ("科研辅助", 3),
        ("虚拟陪伴", 3), ("智能搜索", 3), ("写作辅助", 2), ("视频创作", 1)]
fig, ax = plt.subplots(figsize=(9, 5.2))
names = [w[0] for w in wcat][::-1]
vals = [w[1] for w in wcat][::-1]
colors = [RED if n in ("效率工具", "通用助手") else BLUE for n in names]
bars = ax.barh(names, vals, color=colors, height=0.62)
for b, v in zip(bars, vals):
    ax.text(b.get_width() + 0.15, b.get_y() + b.get_height() / 2,
            f"{v}", va="center", fontsize=10)
ax.set_xlim(0, 21)
ax.set_title("图 12  观猹（watcha.cn）热门 Top 50 AI 产品类别分布")
ax.spines[["top", "right"]].set_visible(False)
note(fig, "数据来源：watcha.cn 观猹社区热门产品接口（2026-08 抓取，一款产品可属多个类别）")
save(fig, "chart12_watcha_categories.png")

# ============ 图13 观猹社区口碑 Top 产品 ============
hot = json.load(open(os.path.join(DATA, "watcha_hot.json")))["data"]["items"]
ranked = sorted(hot, key=lambda p: -(p["stats"].get("stars", 0)))[:10]
fig, ax = plt.subplots(figsize=(9.5, 5.2))
names = [p["name"] for p in ranked][::-1]
stars = [p["stats"].get("stars", 0) for p in ranked][::-1]
up = [p["stats"].get("upvotes", 0) for p in ranked][::-1]
y = np.arange(len(names))
ax.barh(y + 0.2, stars, color=BLUE, height=0.38, label="收藏（stars）")
ax.barh(y - 0.2, up, color=RED, height=0.38, label="点赞（upvotes）")
ax.set_yticks(y, names)
xmax = max(stars + up)
for yy, v in zip(y + 0.2, stars):
    ax.text(v + xmax * 0.01, yy, str(v), va="center", fontsize=9)
for yy, v in zip(y - 0.2, up):
    ax.text(v + xmax * 0.01, yy, str(v), va="center", fontsize=9)
ax.set_xlim(0, xmax * 1.12)
ax.set_title("图 13  观猹社区互动热度 Top 10 AI 产品")
ax.spines[["top", "right"]].set_visible(False)
ax.legend(loc="lower right")
note(fig, "数据来源：watcha.cn 观猹热门产品接口 stats 字段（2026-08 抓取）")
save(fig, "chart13_watcha_top10.png")

# ============ 图14 AI 原生园区能力雷达（概念图） ============
dims = ["算力密度", "场景开放度", "数据要素供给", "人才与社群",
        "资本可得性", "空间柔性"]
trad = [2.0, 2.5, 1.5, 3.0, 2.5, 2.0]
ainative = [4.5, 4.5, 4.0, 4.5, 4.0, 4.5]
angles = np.linspace(0, 2 * np.pi, len(dims), endpoint=False).tolist()
angles += angles[:1]
fig, ax = plt.subplots(figsize=(7.2, 6.4), subplot_kw=dict(polar=True))
for vals_, color, label in [(trad, GRAY, "传统产业园区"),
                            (ainative, RED, "AI 原生园区（目标形态）")]:
    v = vals_ + vals_[:1]
    ax.plot(angles, v, color=color, linewidth=2, label=label)
    ax.fill(angles, v, color=color, alpha=0.15)
ax.set_xticks(angles[:-1])
ax.set_xticklabels(dims, fontsize=11)
ax.set_ylim(0, 5)
ax.set_yticks([1, 2, 3, 4, 5])
ax.set_yticklabels(["1", "2", "3", "4", "5"], fontsize=8, color=GRAY)
ax.set_title("图 14  传统园区与 AI 原生园区能力体系对比（示意）",
             pad=30, fontsize=14, fontweight="bold")
ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.06), ncol=2)
note(fig, "说明：本图为课题组基于 WAIC2026 调研形成的概念评估框架（5 分制示意），非实测数据")
save(fig, "chart14_park_radar.png")

# ============ 图15 展馆×行业：空间分区即产业地图 ============
halls = ["世博展览馆", "张江科学会堂", "西岸国际会展中心", "世博中心"]
inds = ["机器人与智能硬件", "算力·芯片·基础设施", "大模型与生成式AI",
        "行业AI应用", "AI技术与算法", "企业服务与营销", "机构与平台"]
hall_data = {
    "世博展览馆": [216, 135, 108, 90, 63, 66, 45],
    "张江科学会堂": [3, 80, 3, 5, 10, 10, 3],
    "西岸国际会展中心": [12, 3, 22, 12, 6, 4, 4],
    "世博中心": [2, 8, 4, 6, 10, 6, 15],
}
fig, ax = plt.subplots(figsize=(10.2, 5.2))
x = np.arange(len(halls))
bottom = np.zeros(len(halls))
colors15 = ["#C8102E", "#0E4E9B", "#2E7CC3", "#3AA17E", "#E8734A",
            "#8B6BB8", "#8A93A3"]
for i, ind in enumerate(inds):
    vals = [hall_data[h][i] for h in halls]
    ax.bar(x, vals, bottom=bottom, color=colors15[i], width=0.58, label=ind)
    bottom += np.array(vals)
ax.set_xticks(x, ["世博展览馆\n733 家", "张江科学会堂\n114 家",
                  "西岸会展中心\n63 家", "世博中心\n52 家"])
ax.set_ylabel("参展商家数")
ax.set_title("图 15  展馆 × 行业：产业内容决定空间分区")
ax.legend(loc="upper right", fontsize=8.5)
ax.spines[["top", "right"]].set_visible(False)
note(fig, "数据来源：WAIC2026 全量资源整合总表·参展商库（963 家按展馆归并；智能网联汽车因样本量小并入未单列）")
save(fig, "chart15_hall_industry.png")

# ============ 图16 沪京粤浙产业分工 ============
regions = ["上海", "北京", "广东", "浙江"]
# 机器人 / 算力 / 大模型
r_robot = [58, 30, 50, 40]
r_compute = [54, 47, 46, 18]
r_model = [57, 31, 9, 9]
fig, ax = plt.subplots(figsize=(9.2, 4.8))
x = np.arange(len(regions))
w = 0.25
ax.bar(x - w, r_robot, w, color=RED, label="机器人与智能硬件")
ax.bar(x, r_compute, w, color=BLUE, label="算力·芯片·基础设施")
ax.bar(x + w, r_model, w, color="#3AA17E", label="大模型与生成式AI")
for xs, vs in [(x - w, r_robot), (x, r_compute), (x + w, r_model)]:
    for xx, v in zip(xs, vs):
        ax.text(xx, v + 0.8, str(v), ha="center", fontsize=9)
ax.set_xticks(x, ["上海（275 家）", "北京（193 家）", "广东（141 家）", "浙江（97 家）"])
ax.set_ylabel("参展商家数")
ax.set_title("图 16  国内四极产业分工：上海均衡、北京算力、广东硬件、浙江机器人")
ax.legend(loc="upper right")
ax.spines[["top", "right"]].set_visible(False)
note(fig, "数据来源：WAIC2026 全量资源整合总表·参展商库（按注册地归并；上海含上海市，广东含深圳/广州，浙江含杭州）")
save(fig, "chart16_regional_mix.png")

# ============ 图17 行业AI应用垂直分布 ============
verts = [("医疗健康/生物医药", 26), ("能源/电力/工业", 18),
         ("金融科技/支付", 15), ("智慧城市/政务", 15),
         ("文旅/文娱/游戏", 12), ("教育科技", 11),
         ("零售/电商/消费", 8), ("农业/食品科技", 5),
         ("物流/供应链", 4)]
fig, ax = plt.subplots(figsize=(9, 4.8))
names = [v[0] for v in verts][::-1]
vals = [v[1] for v in verts][::-1]
bars = ax.barh(names, vals, color=BLUE, height=0.62)
bars[-1].set_color(RED)
for b, v in zip(bars, vals):
    ax.text(b.get_width() + 0.3, b.get_y() + b.get_height() / 2,
            f"{v} 家", va="center", fontsize=10)
ax.set_xlim(0, 32)
ax.set_title("图 17  行业 AI 应用垂直分布：医疗领先，物流与农业仍是空白带")
ax.spines[["top", "right"]].set_visible(False)
note(fig, "数据来源：WAIC2026 全量资源整合总表·行业统计（行业AI应用大类共 114 家）")
save(fig, "chart17_vertical_apps.png")

# ============ 图18 跨展会供应链外延 ============
xex = [("2026 无人机参展商", 1138), ("无人机名片合集", 534),
       ("华南国际工业博览会", 431), ("中国物流装备展", 380),
       ("上海国际传感器会展", 340), ("小电机磁性材料名单", 211),
       ("物联网智能传感器展", 192), ("深圳人工智能会展", 153),
       ("2026 人工智能眼镜展", 69)]
fig, ax = plt.subplots(figsize=(9, 4.8))
names = [x[0] for x in xex][::-1]
vals = [x[1] for x in xex][::-1]
bars = ax.barh(names, vals, color=BLUE, height=0.62)
for b, v in zip(bars, vals):
    ax.text(b.get_width() + 8, b.get_y() + b.get_height() / 2,
            str(v), va="center", fontsize=10)
ax.set_xlim(0, 1350)
ax.set_title("图 18  跨展会名录：AI 产业的硬件供应链外延（去重前）")
ax.spines[["top", "right"]].set_visible(False)
note(fig, "数据来源：WAIC2026 全量资源整合总表·跨展会名录统计（联系人已去重并入品牌主库）")
save(fig, "chart18_cross_expo.png")

# ============ 图19 杨浦园区产业构成 ============
yp_ind = [("在线新经济", 65, 825), ("AI/大模型", 58, 1238),
          ("设计创意", 54, 861), ("专精特新制造服务", 50, 856),
          ("检验检测/科技服务", 47, 1222), ("科技金融", 46, 734)]
fig, ax = plt.subplots(figsize=(9.2, 4.6))
names = [p[0] for p in yp_ind][::-1]
cnts = [p[1] for p in yp_ind][::-1]
areas = [p[2] for p in yp_ind][::-1]
colors = [RED if n == "AI/大模型" else BLUE for n in names]
bars = ax.barh(names, cnts, color=colors, height=0.58)
for b, v, a in zip(bars, cnts, areas):
    ax.text(b.get_width() + 0.6, b.get_y() + b.get_height() / 2,
            f"{v} 家 · 户均 {a:.0f}㎡", va="center", fontsize=10)
ax.set_xlim(0, 88)
ax.set_title("图 19  杨浦产业园区入驻结构：AI/大模型户均面积最大（1238㎡）")
ax.spines[["top", "right"]].set_visible(False)
note(fig, "数据来源：本中心《上海商务楼宇与产业园区市场报告·杨浦区》样本口径（入驻企业样本 320 家）")
save(fig, "chart19_yangpu_industry.png")

# ============ 图20 杨浦板块租金与空置 ============
plates = [("大创智", 4.88, 17.9, 24), ("五角场商圈", 4.80, 16.5, 7),
          ("新江湾", 4.48, 17.7, 5), ("同济周边", 4.09, 18.0, 0),
          ("杨浦滨江", 4.01, 15.4, 8), ("内环商务", 3.51, 18.0, 7),
          ("国定路高校带", 3.21, 16.3, 7), ("外环", 2.94, 11.4, 0),
          ("滨江东延伸", 2.87, 12.9, 0)]
fig, ax = plt.subplots(figsize=(9.4, 5.4))
xs = [p[1] for p in plates]
ys = [p[2] for p in plates]
ss = [max(80, p[3] * 18) for p in plates]
ax.scatter(xs, ys, s=ss, c=[RED if p[3] >= 8 else BLUE for p in plates],
           alpha=0.85, edgecolors="white", linewidths=0.7)
offsets = {
    "大创智": (8, 8), "五角场商圈": (-72, -18), "新江湾": (8, 6),
    "同济周边": (8, -14), "杨浦滨江": (8, 6), "内环商务": (8, 6),
    "国定路高校带": (8, 6), "外环": (8, -16), "滨江东延伸": (8, 6),
}
for p in plates:
    dx, dy = offsets.get(p[0], (5, 5))
    ax.annotate(f"{p[0]}（AI {p[3]}家）", (p[1], p[2]), fontsize=8,
                xytext=(dx, dy), textcoords="offset points", color="#333")
ax.set_xlabel("成交租金（元/㎡·天）")
ax.set_ylabel("空置率（%）")
ax.set_title("图 20  杨浦板块：租金梯度与 AI 企业集聚（气泡大小≈AI 企业数）")
ax.spines[["top", "right"]].set_visible(False)
note(fig, "数据来源：本中心杨浦区办公园区市场报告；大创智以 24 家 AI/大模型企业成为区内最显著的 AI 集聚板块")
save(fig, "chart20_yangpu_plates.png")

# ============ 图21 世博展馆展位结构 ============
booths = [("H4 创新创投", 301), ("H3 具身智能", 162),
          ("H1 大模型/Agent", 150), ("H2 算力基建", 98),
          ("其他展位", 252)]
fig, ax = plt.subplots(figsize=(8.8, 4.2))
names = [b[0] for b in booths]
vals = [b[1] for b in booths]
colors = [RED, "#C8102E", BLUE, "#2E7CC3", GRAY]
# H3 and H4 red-ish - actually H4 is largest so maybe blue with H3 red
colors = [BLUE, RED, "#2E7CC3", "#5FA8DC", GRAY]
bars = ax.bar(names, vals, color=colors, width=0.58)
for b, v in zip(bars, vals):
    ax.text(b.get_x() + b.get_width() / 2, v + 6, f"{v} 家",
            ha="center", fontsize=10)
ax.set_ylim(0, 360)
ax.set_title("图 21  世博展览馆展位结构：创投与具身智能占据最大物理面积")
ax.spines[["top", "right"]].set_visible(False)
note(fig, "数据来源：WAIC2026 全量资源整合总表·参展商库展位号前缀归类（H1–H4 为世博展览馆主展区）")
save(fig, "chart21_booth_halls.png")

# ============ 图22 中美欧产业路径对照（示意评分） ============
dims2 = ["前沿模型", "性价比", "硬件/具身", "场景落地", "治理标准", "产业空间供给"]
cn = [4.2, 4.8, 4.7, 4.5, 3.6, 4.6]
us = [4.8, 2.8, 3.2, 3.5, 3.4, 3.3]
eu = [3.6, 3.0, 2.8, 3.2, 4.8, 3.4]
angles = np.linspace(0, 2 * np.pi, len(dims2), endpoint=False).tolist()
angles += angles[:1]
fig, ax = plt.subplots(figsize=(7.4, 6.6), subplot_kw=dict(polar=True))
for vals_, color, label in [(cn, RED, "中国路径"),
                            (us, BLUE, "美国路径"),
                            (eu, GRAY, "欧洲路径")]:
    v = vals_ + vals_[:1]
    ax.plot(angles, v, color=color, linewidth=2, label=label)
    ax.fill(angles, v, color=color, alpha=0.12)
ax.set_xticks(angles[:-1])
ax.set_xticklabels(dims2, fontsize=11)
ax.set_ylim(0, 5)
ax.set_yticks([1, 2, 3, 4, 5])
ax.set_yticklabels(["1", "2", "3", "4", "5"], fontsize=8, color=GRAY)
ax.set_title("图 22  中美欧人工智能产业路径对照（示意）",
             pad=30, fontsize=14, fontweight="bold")
ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.06), ncol=3)
note(fig, "说明：5 分制示意框架，综合 WAIC 产业结构、XSCT 定价与公开治理进展；非单一指标实测")
save(fig, "chart22_cn_us_eu.png")

print("all charts done")
