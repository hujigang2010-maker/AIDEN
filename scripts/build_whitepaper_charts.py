# -*- coding: utf-8 -*-
"""《面向2030：新能源、储能与新型电力系统白皮书》图表。

数据口径见白皮书附录 A。2030 年为规划坐标，2025 年为官方统计或权威转述。
"""
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from matplotlib import font_manager
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

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
    "legend.fontsize": 9.5,
})
BLUE = "#0E4E9B"
RED = "#C8102E"
TEAL = "#3AA17E"
ORANGE = "#E8734A"
GOLD = "#F2A65A"
GRAY = "#8A93A3"
LIGHT = "#DCE6F1"
PALETTE = [BLUE, "#2E7CC3", TEAL, ORANGE, RED, GOLD, "#8B6BB8", "#6B7B8C"]
OUT = "/workspace/whitepaper/assets/charts"
os.makedirs(OUT, exist_ok=True)


def note(fig, text):
    fig.text(0.01, 0.01, text, fontsize=8, color=GRAY)


def save(fig, name):
    fig.savefig(os.path.join(OUT, name), bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("saved", name)


# ---------- 图1 装机 2025 vs 2030 ----------
labels = ["风光合计", "常规水电", "核电", "抽水蓄能", "新型储能", "西电东送"]
y2025 = [18.4, 3.8, 0.62, 0.66, 1.36, 3.4]
y2030 = [28.0, 4.1, 1.1, 1.6, 3.0, 4.2]
x = np.arange(len(labels))
w = 0.36
fig, ax = plt.subplots(figsize=(10.2, 5.2))
b1 = ax.bar(x - w / 2, y2025, w, label="2025 年实绩", color=BLUE)
b2 = ax.bar(x + w / 2, y2030, w, label="2030 年规划", color=RED)
for bars in (b1, b2):
    for b in bars:
        ax.text(b.get_x() + b.get_width() / 2, b.get_height() + 0.25,
                f"{b.get_height():g}", ha="center", va="bottom", fontsize=8.5)
ax.set_xticks(x)
ax.set_xticklabels(labels)
ax.set_ylabel("亿千瓦")
ax.set_title("图1  中国关键电力指标：2025年实绩与2030年规划坐标")
ax.set_ylim(0, 32)
ax.legend(frameon=False)
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
note(fig, "数据：国家能源局、十五五规划公开要点。风光2030年为28亿千瓦以上下限。单位：亿千瓦。")
save(fig, "chart01_capacity_2025_2030.png")


# ---------- 图2 调节资源 ----------
names = ["新型储能", "抽水蓄能", "虚拟电厂\n调节能力", "车网互动\n可调充电", "需求响应\n（峰值负荷5%）"]
vals = [3.0, 1.6, 0.5, 0.5, 0.8]
# 需求响应：以2030年最大负荷粗略按16亿千瓦×5%=0.8亿千瓦示意
colors = [BLUE, "#2E7CC3", TEAL, ORANGE, GOLD]
fig, ax = plt.subplots(figsize=(9.6, 5.0))
bars = ax.barh(names[::-1], vals[::-1], color=colors[::-1], height=0.58)
for b, v in zip(bars, vals[::-1]):
    ax.text(b.get_width() + 0.06, b.get_y() + b.get_height() / 2,
            f"{v:g} 亿千瓦", va="center", fontsize=10)
ax.set_xlim(0, 3.7)
ax.set_xlabel("亿千瓦")
ax.set_title("图2  2030年关键调节资源规划规模")
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
note(fig, "虚拟电厂、车网互动为规划目标；需求响应按最大用电负荷5%以上、以约16亿千瓦峰值示意，非官方点目标。")
save(fig, "chart02_flexibility_2030.png")


# ---------- 图3 风光内部结构 ----------
fig, axes = plt.subplots(1, 2, figsize=(10.2, 4.8))
wind = [5.9, 0.47]
wlab = ["陆上风电 5.9", "海上风电 0.47"]
axes[0].pie(wind, labels=wlab, colors=[BLUE, "#93C6E8"], startangle=90,
            wedgeprops=dict(width=0.45, edgecolor="white"),
            textprops={"fontsize": 9})
axes[0].set_title("风电累计 6.4 亿千瓦", fontsize=12)
pv = [6.7, 5.3]
plab = ["集中式 6.7", "分布式 5.3"]
axes[1].pie(pv, labels=plab, colors=[TEAL, GOLD], startangle=90,
            wedgeprops=dict(width=0.45, edgecolor="white"),
            textprops={"fontsize": 9})
axes[1].set_title("光伏累计 12.0 亿千瓦", fontsize=12)
fig.suptitle("图3  2025年底风光装机内部结构（亿千瓦）", fontsize=14, fontweight="bold")
note(fig, "数据：国家能源局《2025年可再生能源并网运行情况》。风光合计18.4亿千瓦，历史性超过火电装机。")
save(fig, "chart03_wind_solar_mix.png")


# ---------- 图4 储能增长 ----------
years = ["2020", "2025", "2026.6", "2030"]
gw = [0.03, 1.36, 1.53, 3.00]
fig, ax = plt.subplots(figsize=(9.2, 5.0))
ax.plot(years, gw, color=BLUE, marker="o", linewidth=2.4, markersize=8)
ax.fill_between(range(len(years)), gw, color=BLUE, alpha=0.12)
for i, v in enumerate(gw):
    ax.text(i, v + 0.12, f"{v} 亿千瓦", ha="center", fontsize=10, color=BLUE)
ax.set_ylabel("亿千瓦")
ax.set_title("图4  中国新型储能装机：从“十四五”跃升到2030")
ax.set_ylim(0, 3.6)
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.annotate("十四五：300万→1.36亿千瓦", xy=(1, 1.36), xytext=(0.15, 2.15),
            fontsize=9, color=GRAY,
            arrowprops=dict(arrowstyle="->", color=GRAY))
note(fig, "2020年为约数；2025、2026.6来自官方报告；2030年为规划目标3亿千瓦。不含抽水蓄能。")
save(fig, "chart04_storage_growth.png")


# ---------- 图5 时长分工 ----------
fig, ax = plt.subplots(figsize=(10.4, 5.1))
rows = [
    (0.2, 0.9, "秒级—15分钟\n飞轮 / 超容 / 高功率锂电", BLUE),
    (1.0, 3.5, "1—4小时\n锂电（液态/半固态）· 部分钠电", "#2E7CC3"),
    (4.0, 8.5, "4—12小时\n液流 · 压缩空气 · 长时锂电/钠电", TEAL),
    (8.0, 24, "日以上 / 季节性\n抽蓄 · 储热储冷 · 氢氨醇", ORANGE),
]
for y, (x0, x1, lab, c) in enumerate(rows):
    ax.barh(y, x1 - x0, left=x0, height=0.55, color=c, alpha=0.9)
    ax.text((x0 + x1) / 2, y, lab, ha="center", va="center",
            color="white", fontsize=9, fontweight="bold")
ax.set_yticks([])
ax.set_xlabel("连续放电时长（小时，示意）")
ax.set_xlim(0, 26)
ax.set_title("图5  储能技术按调节时长的分工图谱")
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.spines["left"].set_visible(False)
note(fig, "示意图，非精确技术边界。政策付费应按“提供了多长、多可靠的调节”，而不是按技术路线站队。")
save(fig, "chart05_storage_duration.png")


# ---------- 图6 西电东送 ----------
fig, ax = plt.subplots(figsize=(8.6, 4.8))
labs = ["2025 年\n约 3.4 亿千瓦", "2030 年\n超过 4.2 亿千瓦"]
vs = [3.4, 4.2]
bars = ax.bar(labs, vs, color=[BLUE, RED], width=0.45)
for b, v in zip(bars, vs):
    ax.text(b.get_x() + b.get_width() / 2, v + 0.08, f"{v} 亿千瓦",
            ha="center", fontsize=11, fontweight="bold")
ax.set_ylim(0, 5.1)
ax.set_ylabel("亿千瓦")
ax.set_title("图6  西电东送能力：大范围互济仍须加力")
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.annotate("+0.8 亿千瓦以上\n（规划新增外送能力）", xy=(1, 4.2), xytext=(0.35, 4.55),
            fontsize=9, color=RED,
            arrowprops=dict(arrowstyle="->", color=RED))
note(fig, "2025年约3.4亿千瓦见国家统计局解读；2030年超过4.2亿千瓦见十五五规划公开要点。")
save(fig, "chart06_west_to_east.png")


# ---------- 图7 住房能源节点四层结构 ----------
fig, ax = plt.subplots(figsize=(10.4, 5.6))
ax.set_xlim(0, 10)
ax.set_ylim(0, 6)
ax.axis("off")
ax.set_title("图7  住房能源节点：可确权、可接入、可分享、可托底", pad=12)
layers = [
    (0.4, 4.35, 9.2, 1.15, BLUE, "城市与电网层  ·  配电网承载力 · 虚拟电厂市场入口 · 极端天气应急"),
    (0.7, 3.05, 8.6, 1.15, "#2E7CC3", "社区与物业层  ·  屋顶/车位公共决策 · 统建统服 · 共享储能与收益账户"),
    (1.0, 1.75, 8.0, 1.15, TEAL, "建筑与设施层  ·  光储直柔 · 有序充电/V2G · 柔性空调与直流配电"),
    (1.3, 0.45, 7.4, 1.15, ORANGE, "住户与公平层  ·  账单可读可负担 · 租户分享 · 能源贫困对冲"),
]
for x, y, w, h, c, t in layers:
    box = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.15",
                         facecolor=c, edgecolor="none", alpha=0.92)
    ax.add_patch(box)
    ax.text(x + w / 2, y + h / 2, t, ha="center", va="center",
            color="white", fontsize=10.5)
note(fig, "复旦大学住房政策研究会提出：住房即能源节点，四层缺一则分布式能源停在最后一公里。")
save(fig, "chart07_housing_energy_node.png")


# ---------- 图8 五类住房 ----------
fig, ax = plt.subplots(figsize=(10.2, 5.2))
cats = ["新建商品房", "存量商品房", "老旧小区/\n城中村/单电源高层", "保障房/\n公租房/人才公寓", "农村住房/\n县域社区"]
priority = [8.5, 6.5, 9.2, 8.8, 7.6]
potential = [7.0, 6.0, 4.5, 7.5, 9.5]
x = np.arange(len(cats))
w = 0.35
ax.bar(x - w / 2, priority, w, label="近期政策优先级", color=RED)
ax.bar(x + w / 2, potential, w, label="资源/制度潜力", color=BLUE)
ax.set_xticks(x)
ax.set_xticklabels(cats, fontsize=9)
ax.set_ylabel("指数（研究会评估，1—10）")
ax.set_ylim(0, 12)
ax.set_title("图8  五类住房空间：近期优先级与中长期潜力")
ax.legend(frameon=False, loc="upper right")
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
note(fig, "指数为研究会基于产权集中度、配网约束、安全欠账与资源潜力的定性评估，便于比较，不是统计调查得分。")
save(fig, "chart08_housing_types.png")


# ---------- 图9 新业态跃迁 ----------
fig, ax = plt.subplots(figsize=(9.8, 5.1))
items = ["新型储能", "虚拟电厂\n调节能力", "车网互动\n可调充电", "充电网络\n支撑车辆"]
now = [1.36, 0.18, 0.10, 0.35]   # 车辆用亿辆示意：约0.35亿以上保有/渗透，这里用“支撑能力”相对值不合适
# Better: use explicit units via grouped bars with annotations
# 储能亿千瓦、虚拟电厂亿千瓦、V2G亿千瓦、车辆亿辆
now = [1.36, 0.17, 0.10, None]
fut = [3.00, 0.50, 0.50, 1.10]
# For vehicles we only have 2030 target 1.1亿辆; 2025 EV fleet ~0.3+亿. Use 0.32.
now = [1.36, 0.17, 0.10, 0.32]
fut = [3.00, 0.50, 0.50, 1.10]
x = np.arange(len(items))
w = 0.34
b1 = ax.bar(x - w / 2, now, w, label="2025 年约数", color=BLUE)
b2 = ax.bar(x + w / 2, fut, w, label="2030 年规划", color=RED)
for i, (n, f) in enumerate(zip(now, fut)):
    unit = "亿千瓦" if i < 3 else "亿辆"
    ax.text(i - w / 2, n + 0.04, f"{n:g}", ha="center", fontsize=8, color=BLUE)
    ax.text(i + w / 2, f + 0.04, f"{f:g}", ha="center", fontsize=8, color=RED)
ax.set_xticks(x)
ax.set_xticklabels(items)
ax.set_ylabel("储能/调节：亿千瓦；车辆：亿辆")
ax.set_title("图9  用电侧新业态：2025—2030 目标跃迁")
ax.set_ylim(0, 3.6)
ax.legend(frameon=False)
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
note(fig, "虚拟电厂2025年按规划“增长近2倍”反推约1700万千瓦，为示意；车辆为充电网络支撑能力目标，不是保有量预测。")
save(fig, "chart09_new_business.png")


# ---------- 图10 路线图 ----------
fig, ax = plt.subplots(figsize=(10.6, 5.4))
ax.set_xlim(0, 10)
ax.set_ylim(0, 6.2)
ax.axis("off")
ax.set_title("图10  面向2030：建机制 → 上规模 → 验收初步建成", pad=8)
phases = [
    (0.35, 1.2, 2.9, 4.2, BLUE, "2026—2027",
     "建机制 · 补欠账 · 设试点\n\n容量电价落地\n配电与单电源摸排\n保障房光储直柔\n有序充电默认开通"),
    (3.55, 1.2, 2.9, 4.2, "#2E7CC3", "2028—2029",
     "上规模 · 通市场 · 进社区\n\n通道与配网投产高峰\n长时储能商业化\n能源就绪交付标配\n统一市场关键规则"),
    (6.75, 1.2, 2.9, 4.2, RED, "2030 验收",
     "初步建成新型电力系统\n\n非化石电量约50%\n储能3亿+抽蓄1.6亿\n分布式接入9亿千瓦\n账单可读、极端天气可托底"),
]
for x, y, w, h, c, title, body in phases:
    box = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.03,rounding_size=0.12",
                         facecolor=c, edgecolor="none", alpha=0.93)
    ax.add_patch(box)
    ax.text(x + w / 2, y + h - 0.45, title, ha="center", va="center",
            color="white", fontsize=13, fontweight="bold")
    ax.text(x + w / 2, y + 1.85, body, ha="center", va="center",
            color="white", fontsize=9.2, linespacing=1.45)
for x in (3.25, 6.45):
    ax.annotate("", xy=(x + 0.25, 3.3), xytext=(x, 3.3),
                arrowprops=dict(arrowstyle="->", color=GRAY, lw=1.8))
note(fig, "路线图为研究会归纳，用于政策讨论。正式验收以国家规划文本为准。")
save(fig, "chart10_roadmap.png")

print("all charts done")
