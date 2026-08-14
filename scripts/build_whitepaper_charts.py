# -*- coding: utf-8 -*-
"""医疗、养老与提效服务白皮书配套图。数据为公开口径示意，详见文末来源。"""
import os
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import FancyBboxPatch

FONT = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
font_manager.fontManager.addfont(FONT)
plt.rcParams["font.family"] = "WenQuanYi Micro Hei"
plt.rcParams["axes.unicode_minus"] = False

OUT = "/workspace/whitepaper/assets/charts"
os.makedirs(OUT, exist_ok=True)
BLUE, RED, TEAL, AMBER, GRAY = "#0E4E9B", "#C8102E", "#1A7A6D", "#C47B17", "#5B616B"
LIGHT = "#DCE6F1"


def save(fig, name):
    fig.savefig(os.path.join(OUT, name), dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("saved", name)


# 1 老龄化底数（指标卡，避免不同量纲同轴）
fig, ax = plt.subplots(figsize=(9.6, 4.6))
ax.set_xlim(0, 12)
ax.set_ylim(0, 5.2)
ax.axis("off")
ax.set_title("2025 年末老龄底数（老龄事业发展公报）", fontsize=14, color=BLUE, pad=8)
kpis = [
    (0.25, BLUE, "3.23 亿", "60 岁及以上人口", "占总人口 23.0%"),
    (3.2, TEAL, "2.24 亿", "65 岁及以上人口", "占总人口 15.9%"),
    (6.15, AMBER, "23.1%", "65 岁抚养比", "每 100 名劳动年龄人口\n负担约 23 名老年人"),
    (9.1, RED, "79.25 岁", "人均预期寿命", "2030 年规划目标 80 岁"),
]
for x, c, big, mid, small in kpis:
    ax.add_patch(FancyBboxPatch((x, 0.55), 2.7, 4.0, boxstyle="round,pad=0.04,rounding_size=0.12",
                                facecolor=c, edgecolor="none"))
    ax.text(x + 1.35, 3.55, big, ha="center", color="white", fontsize=20)
    ax.text(x + 1.35, 2.55, mid, ha="center", color="white", fontsize=12)
    ax.text(x + 1.35, 1.45, small, ha="center", color="white", fontsize=10)
save(fig, "chart01_aging.png")

# 2 养老供给结构
fig, ax = plt.subplots(figsize=(9.2, 5.2))
cats = ["机构床位", "其中护理型\n（约）", "社区床位", "床位合计"]
nums = [498.4, 498.4 * 0.699, 269.5, 767.9]
cols = [BLUE, TEAL, AMBER, RED]
bars = ax.barh(cats[::-1], nums[::-1], color=cols[::-1], height=0.55, zorder=3)
for b, n in zip(bars, nums[::-1]):
    ax.text(n + 8, b.get_y() + b.get_height() / 2, f"{n:.1f} 万张", va="center", fontsize=10)
ax.set_xlim(0, 920)
ax.set_title("2025 年末养老床位结构：总量优化，护理型上升", fontsize=13, color=BLUE)
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.text(0, -0.18, "机构 498.4 万张、护理型占比 69.9%；社区 269.5 万张。十五五目标：2030 年护理型占比 73%。",
        transform=ax.transAxes, fontsize=8, color=GRAY)
save(fig, "chart02_beds.png")

# 3 医疗资源（指标卡）
fig, ax = plt.subplots(figsize=(9.6, 4.6))
ax.set_xlim(0, 12)
ax.set_ylim(0, 5.2)
ax.axis("off")
ax.set_title("2025 年医疗卫生资源：机构增加、床位下降、诊疗上升", fontsize=13, color=BLUE, pad=8)
kpis = [
    (0.25, BLUE, "110.7 万个", "医疗卫生机构", "基层机构 105.5 万个"),
    (3.2, TEAL, "1340 万人", "卫生技术人员", "医师 529 万 / 护士 603 万"),
    (6.15, AMBER, "1009 万张", "医疗卫生床位", "较上年减少约 28 万张"),
    (9.1, RED, "105.8 亿", "全年诊疗人次", "基层占 52.6%"),
]
for x, c, big, mid, small in kpis:
    ax.add_patch(FancyBboxPatch((x, 0.55), 2.7, 4.0, boxstyle="round,pad=0.04,rounding_size=0.12",
                                facecolor=c, edgecolor="none"))
    ax.text(x + 1.35, 3.55, big, ha="center", color="white", fontsize=18)
    ax.text(x + 1.35, 2.55, mid, ha="center", color="white", fontsize=12)
    ax.text(x + 1.35, 1.45, small, ha="center", color="white", fontsize=10)
save(fig, "chart03_health.png")

# 4 十五五民生锚点
fig, ax = plt.subplots(figsize=(9.6, 5.4))
labels = ["预期寿命\n（岁）", "执业医师\n（人/千人）", "注册护士\n（人/千人）", "护理型床位\n占比（%）"]
y2025 = [79.25, 3.1, 4.3, 68]
y2030 = [80, 3.7, 5.1, 73]
import numpy as np
idx = np.arange(len(labels))
w = 0.36
ax.bar(idx - w / 2, y2025, w, color=LIGHT, label="2025（规划基期）", zorder=3, edgecolor=BLUE)
ax.bar(idx + w / 2, y2030, w, color=BLUE, label="2030 目标", zorder=3)
ax.set_xticks(idx)
ax.set_xticklabels(labels)
ax.set_title("“十五五”规划民生锚点：健康、养老同步加码", fontsize=13, color=BLUE)
ax.legend(frameon=False)
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.grid(axis="y", color="#E6E8EC", zorder=0)
ax.text(0, -0.2, "规划基期护理型床位占比 68%；公报 2025 年末实际为 69.9%。入托率另要求提高 6 个百分点。",
        transform=ax.transAxes, fontsize=8, color=GRAY)
save(fig, "chart04_15th_plan.png")

# 5 服务三层
fig, ax = plt.subplots(figsize=(10.6, 4.8))
ax.set_xlim(0, 12)
ax.set_ylim(0, 5)
ax.axis("off")
ax.set_title("三类服务是同一道题：把家庭隐形税改写成可及的社会化服务", fontsize=13, color=BLUE)
blocks = [
    (0.3, BLUE, "医疗", "从住院中心到连续照护\n家庭医生 / 医共体\n康复护理 / 家庭病床"),
    (4.15, TEAL, "养老", "从床位竞赛到服务网络\n家庭养老床位\n一刻钟圈 / 长护险"),
    (8.0, AMBER, "提效服务", "把照护时间还给劳动\n普惠托育 / 家政物业\n延迟退休的配套条件"),
]
for x, c, title, body in blocks:
    ax.add_patch(FancyBboxPatch((x, 0.45), 3.55, 3.9, boxstyle="round,pad=0.04,rounding_size=0.12",
                                facecolor=c, edgecolor="none"))
    ax.text(x + 1.77, 3.7, title, ha="center", color="white", fontsize=16)
    ax.text(x + 1.77, 1.9, body, ha="center", va="center", color="white", fontsize=11)
save(fig, "chart05_three_services.png")

# 6 空间四层
fig, ax = plt.subplots(figsize=(10.4, 5.0))
ax.set_xlim(0, 12)
ax.set_ylim(0, 5)
ax.axis("off")
ax.set_title("住房是总接口：住宅、社区、设施、城市", fontsize=14, color=BLUE)
items = [
    (0.25, BLUE, "住宅", "适老化三档\n家庭养老床位\n加装电梯"),
    (3.2, TEAL, "社区", "嵌入式设施\n一刻钟生活圈\n完整社区"),
    (6.15, AMBER, "设施网", "县—街道—社区\n与医共体叠合\n一老一小综合体"),
    (9.1, RED, "城市", "照护密度入规\n护理人才住房\n存量盘活优先"),
]
for x, c, t, b in items:
    ax.add_patch(FancyBboxPatch((x, 0.4), 2.7, 4.0, boxstyle="round,pad=0.04,rounding_size=0.12",
                                facecolor=c, edgecolor="none"))
    ax.text(x + 1.35, 3.7, t, ha="center", color="white", fontsize=15)
    ax.text(x + 1.35, 1.9, b, ha="center", va="center", color="white", fontsize=11)
save(fig, "chart06_space.png")

# 7 长护险时间轴
fig, ax = plt.subplots(figsize=(10.6, 4.8))
ax.set_xlim(2024.6, 2030.8)
ax.set_ylim(0, 6.2)
ax.axis("off")
ax.set_title("2025—2030：支付、网络与规划目标对齐", fontsize=14, color=BLUE)
ax.plot([2025, 2030], [0.7, 0.7], color=BLUE, lw=3)
for yr in range(2025, 2031):
    ax.plot([yr, yr], [0.55, 0.85], color=BLUE, lw=2)
    ax.text(yr, 0.28, str(yr), ha="center", fontsize=9, color=GRAY)
ev = [
    (2025.05, 5.3, "延迟退休启动 / 中央养老意见", RED),
    (2025.9, 4.2, "老龄底数：3.23 亿；托位 4.73／千人", TEAL),
    (2026.25, 3.1, "长护险实施方案：三年建成", BLUE),
    (2028.9, 2.15, "长护险计划基本全面覆盖", AMBER),
    (2029.2, 5.3, "养老服务网络基本建成（中央目标）", GRAY),
    (2030.0, 4.15, "十五五收官：预期寿命 80 岁", RED),
]
for x, y, text, c in ev:
    ax.plot([x, x], [0.7, y - 0.2], color=c, lw=1)
    ax.scatter([x], [y], color=c, s=38, zorder=5)
    ax.text(x + 0.06, y, text, va="center", fontsize=9)
save(fig, "chart07_timeline.png")

# 8 2030 情景（定性评分）
fig, ax = plt.subplots(figsize=(9.4, 5.4))
dims = ["服务进家", "支付可及", "社区覆盖", "托育普惠", "劳动释放"]
A = [75, 80, 70, 68, 62]
B = [88, 90, 82, 80, 78]
C = [45, 70, 50, 48, 35]
x = np.arange(len(dims))
ax.plot(x, C, "o-", color=RED, label="承压 C", lw=2)
ax.plot(x, A, "s-", color=BLUE, label="基准 A", lw=2)
ax.plot(x, B, "D-", color=TEAL, label="加速 B", lw=2)
ax.set_xticks(x)
ax.set_xticklabels(dims)
ax.set_ylim(0, 100)
ax.set_ylabel("中心定性评分（满分 100）")
ax.set_title("展望 2030：按基准配空间与人才，按承压守“服务进家”", fontsize=13, color=BLUE)
ax.legend(frameon=False, loc="lower right")
ax.grid(axis="y", color="#E6E8EC")
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.text(0, -0.2, "评分为方向性研判，用于比较情景，不是官方考核分数。",
        transform=ax.transAxes, fontsize=8, color=GRAY)
save(fig, "chart08_scenarios.png")

# 9 政策建议
fig, ax = plt.subplots(figsize=(10.4, 4.8))
ax.set_xlim(0, 12)
ax.set_ylim(0, 5)
ax.axis("off")
ax.set_title("政策建议四端：空间、体系、劳动、制度", fontsize=14, color=BLUE)
cols = [
    (0.2, BLUE, "城市与住房", "照护密度入规\n配建限期运营\n改造打包推进"),
    (3.2, TEAL, "医疗与养老", "考核转介成功率\n长护优先上门\n评估一次互认"),
    (6.2, AMBER, "提效与产业", "护理员有房有证\n银发做辅具住宅\n先托育再延退"),
    (9.2, RED, "国家制度", "2028 险成网\n2030 服务进家\n资金支持可运营设施"),
]
for x, c, t, b in cols:
    ax.add_patch(FancyBboxPatch((x, 0.4), 2.7, 4.0, boxstyle="round,pad=0.04,rounding_size=0.12",
                                facecolor=c, edgecolor="none"))
    ax.text(x + 1.35, 3.7, t, ha="center", color="white", fontsize=13)
    ax.text(x + 1.35, 1.9, b, ha="center", va="center", color="white", fontsize=11)
save(fig, "chart09_policy.png")

print("all charts done")
