# -*- coding: utf-8 -*-
"""白皮书配套统计图。数据均为公开口径的示意可视化，详见白皮书文末来源。"""
import os

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib import font_manager
from matplotlib.patches import FancyBboxPatch

FONT = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
font_manager.fontManager.addfont(FONT)
plt.rcParams["font.family"] = "WenQuanYi Micro Hei"
plt.rcParams["axes.unicode_minus"] = False

OUT = "/workspace/whitepaper/assets/charts"
os.makedirs(OUT, exist_ok=True)

BLUE = "#0E4E9B"
RED = "#C8102E"
TEAL = "#1A7A6D"
AMBER = "#C47B17"
GRAY = "#5B616B"
LIGHT = "#DCE6F1"
BG = "#F7F9FC"


def save(fig, name):
    fig.savefig(os.path.join(OUT, name), dpi=180, bbox_inches="tight",
                facecolor="white")
    plt.close(fig)
    print("saved", name)


# 1 跨境电商规模
fig, ax = plt.subplots(figsize=(9.2, 5.2))
years = ["2020", "2025", "2026Q1"]
vals = [1.69, 2.75, 0.618]
colors = [BLUE, BLUE, TEAL]
bars = ax.bar(years, vals, color=colors, width=0.55, zorder=3)
ax.set_ylabel("万亿元人民币")
ax.set_title("中国跨境电商进出口规模（海关口径）", fontsize=15, color=BLUE, pad=12)
for b, v, note in zip(bars, vals, ["约 1.69", "2.75\n较2020 +69.7%", "0.62\n出口 0.47"]):
    ax.text(b.get_x() + b.get_width() / 2, v + 0.08, note, ha="center",
            va="bottom", fontsize=10, color=GRAY)
ax.set_ylim(0, 3.4)
ax.axhline(2.75, color=RED, ls="--", lw=0.8, alpha=0.5)
ax.grid(axis="y", color="#E6E8EC", zorder=0)
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.text(0, -0.18, "数据：海关总署 2026-01-14、2026-04-14 国新办发布会；2020 年约数为公开报道口径。",
        transform=ax.transAxes, fontsize=8, color=GRAY)
save(fig, "chart01_cbec_scale.png")

# 2 出海模式演进
fig, ax = plt.subplots(figsize=(10.4, 4.6))
ax.set_xlim(0, 10)
ax.set_ylim(0, 4.2)
ax.axis("off")
ax.set_title("中国企业出海：四层叠加，而不再按旧顺序展开", fontsize=14, color=BLUE, pad=8)
stages = [
    (0.4, "贸易型\n货物过海", "跨境电商 / 一般贸易\n快、轻、可试错", BLUE),
    (2.9, "品牌型\n心智过海", "商标、内容、售后\n告别铺货", TEAL),
    (5.4, "产能型\n工厂过海", "中国 + N 属地制造\n本地含量 / 原产地", AMBER),
    (7.9, "体系型\n规则过海", "标准、数据、金融\n仓厂房城闭环", RED),
]
for x, title, desc, c in stages:
    box = FancyBboxPatch((x, 1.15), 2.15, 2.35, boxstyle="round,pad=0.04,rounding_size=0.15",
                         facecolor=c, edgecolor="none")
    ax.add_patch(box)
    ax.text(x + 1.075, 2.85, title, ha="center", va="center", color="white",
            fontsize=12, fontweight="bold")
    ax.text(x + 1.075, 1.85, desc, ha="center", va="center", color="white", fontsize=9)
    if x < 7:
        ax.annotate("", xy=(x + 2.35, 2.3), xytext=(x + 2.15, 2.3),
                    arrowprops=dict(arrowstyle="->", color=GRAY, lw=1.6))
ax.text(5, 0.45, "2026 年特征：三层同时发生在同一家企业、同一个区域", ha="center",
        fontsize=10, color=GRAY)
save(fig, "chart02_outbound_layers.png")

# 3 全球监管时间轴
fig, ax = plt.subplots(figsize=(10.6, 5.4))
ax.set_xlim(2024.6, 2030.6)
ax.set_ylim(0, 7)
ax.axis("off")
ax.set_title("2025—2030 全球规则时间轴（现行公开安排）", fontsize=14, color=BLUE)
events = [
    (2025.65, 6.1, "2025.8 美国暂停各国低值免税", RED),
    (2026.0, 5.1, "2026.1 CBAM 确定期 / 印尼 EV 本地含量", AMBER),
    (2026.25, 4.1, "2026.3—4 国内：政府工作报告、跨关区退货", BLUE),
    (2026.5, 3.1, "2026.7 欧盟取消 150 欧元关税豁免", RED),
    (2027.15, 2.1, "2027.2 欧盟电池护照（现行时间表）", TEAL),
    (2028.0, 1.1, "2028 欧盟低值包裹转入全额征税", RED),
    (2030.0, 6.1, "2030 十五五收官 / 印尼 TKDN 规划 80%", GRAY),
]
ax.plot([2025, 2030], [0.55, 0.55], color=BLUE, lw=3)
for yr in range(2025, 2031):
    ax.plot([yr, yr], [0.42, 0.68], color=BLUE, lw=2)
    ax.text(yr, 0.18, str(yr), ha="center", fontsize=9, color=GRAY)
for x, y, text, c in events:
    ax.plot([x, x], [0.55, y - 0.15], color=c, lw=1)
    ax.scatter([x], [y], color=c, s=40, zorder=5)
    ax.text(x + 0.08, y, text, va="center", fontsize=9, color="#1A1A1A")
save(fig, "chart03_rule_timeline.png")

# 4 海外仓样本增长
fig, ax = plt.subplots(figsize=(9.2, 5.2))
yrs = [2021, 2022, 2023, 2024, 2025, 2026]
n = [1260, 1800, 2600, 4436, 5300, 6200]
ax.fill_between(yrs, n, color=LIGHT, alpha=0.9)
ax.plot(yrs, n, color=BLUE, lw=2.6, marker="o")
for x, y in zip(yrs, n):
    ax.text(x, y + 180, str(y), ha="center", fontsize=9, color=BLUE)
ax.set_title("行业调研口径：海外仓样本数量（跨境眼）", fontsize=14, color=BLUE)
ax.set_ylabel("个")
ax.grid(axis="y", color="#E6E8EC")
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.text(0, -0.18,
        "注：行业样本口径，含各类海外仓；与“专注跨境电商海外仓超 1800 个、面积超 2200 万㎡”的媒体口径不可直接加总。",
        transform=ax.transAxes, fontsize=8, color=GRAY)
save(fig, "chart04_overseas_warehouse.png")

# 5 市场分层
fig, ax = plt.subplots(figsize=(9.6, 5.6))
ax.set_xlim(0, 10)
ax.set_ylim(0, 8)
ax.axis("off")
ax.set_title("2026—2030 市场分层：利润、增长、期权", fontsize=14, color=BLUE)
blocks = [
    (0.4, 4.3, 4.4, 3.1, BLUE, "欧美 · 利润与规则高地",
     "品牌精耕 / 高合规成本\n低值直邮窗口关闭\n适合：高附加值与本地服务"),
    (5.2, 4.3, 4.4, 3.1, TEAL, "东盟 + 中东 · 增长引擎",
     "产能与消费双栖\n本地含量与园区配套\n适合：中国 + N 与仓网加密"),
    (0.4, 0.5, 4.4, 3.1, AMBER, "拉美 / 非洲 / 中亚 · 布局期权",
     "电商渗透率仍在跃升\n制度方差大，必须一国一策\n适合：丝路电商与人民币结算"),
    (5.2, 0.5, 4.4, 3.1, RED, "中国国内 · 总部与产业带",
     "研发、品牌、合规中台\n综试区 / 人才住房 / 检测认证\n把决策留在国内"),
]
for x, y, w, h, c, title, body in blocks:
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.03,rounding_size=0.12",
                                facecolor=c, edgecolor="none", alpha=0.92))
    ax.text(x + 0.25, y + h - 0.55, title, color="white", fontsize=12, fontweight="bold")
    ax.text(x + 0.25, y + 0.55, body, color="white", fontsize=10, va="bottom")
save(fig, "chart05_market_layers.png")

# 6 四维合规
fig, ax = plt.subplots(figsize=(9.2, 5.4))
labels = ["产品安全\n与准入", "税务\n与关税", "环保\n与循环", "数据\n与平台治理"]
vals = [88, 95, 80, 78]
bars = ax.barh(labels[::-1], vals[::-1], color=[TEAL, RED, AMBER, BLUE][::-1], height=0.55)
ax.set_xlim(0, 110)
ax.set_xlabel("作为 2030 年市场准入门槛的重要性（中心定性评分，满分 100）")
ax.set_title("跨境经营四维合规：缺一即可能失去市场", fontsize=14, color=BLUE)
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
for b, t in zip(bars, ["GDPR / 数据出境 / 平台 KYC",
                       "EPR / CBAM / 电池护照",
                       "美低值免税暂停 / 欧 IOSS 与全额税",
                       "GPSR / CE / CPSC / FDA"][::-1]):
    ax.text(b.get_width() + 1.5, b.get_y() + b.get_height() / 2, t,
            va="center", fontsize=8, color=GRAY)
save(fig, "chart06_compliance.png")

# 7 空间四层
fig, ax = plt.subplots(figsize=(9.8, 5.8))
ax.set_xlim(0, 10)
ax.set_ylim(0, 10)
ax.axis("off")
ax.set_title("出海的空间框架：货、产、人、城", fontsize=15, color=BLUE)
rings = [
    (5, 5, 4.3, LIGHT, "城  综试区 / 总部 / 产业带城市"),
    (5, 5, 3.3, "#B8CDE4", "人  外派住房 / 属地职住 / 国内人才公寓"),
    (5, 5, 2.3, "#7FA3CC", "产  海外工厂 / 合作园区 / 本地含量"),
    (5, 5, 1.3, BLUE, "货  海外仓 / 退货翻新 / 智能分拨"),
]
for x, y, r, c, label in rings:
    circ = plt.Circle((x, y), r, color=c, ec="white", lw=3)
    ax.add_patch(circ)
ax.set_aspect("equal")
ax.text(5, 5, "货", ha="center", va="center", color="white", fontsize=16, fontweight="bold")
ax.text(5, 6.55, "产", ha="center", color="white", fontsize=13, fontweight="bold")
ax.text(5, 7.7, "人", ha="center", color=BLUE, fontsize=13, fontweight="bold")
ax.text(5, 8.85, "城", ha="center", color=BLUE, fontsize=13, fontweight="bold")
ax.text(5, 0.45, "住房政策的分析工具——选址、密度、配套、可负担性、职住——可直接用于观察出海",
        ha="center", fontsize=9, color=GRAY)
save(fig, "chart07_space_four.png")

# 8 2030 情景
fig, ax = plt.subplots(figsize=(9.4, 5.4))
names = ["承压情景 C", "基准情景 A", "乐观情景 B"]
low = [3.3, 4.2, 5.0]
high = [3.6, 4.8, 5.5]
mid = [(a + b) / 2 for a, b in zip(low, high)]
err_lo = [m - a for m, a in zip(mid, low)]
err_hi = [b - m for m, b in zip(mid, high)]
colors = [RED, BLUE, TEAL]
ax.bar(names, mid, color=colors, width=0.5, zorder=3)
ax.errorbar(names, mid, yerr=[err_lo, err_hi], fmt="none", ecolor="#1A1A1A",
            capsize=8, lw=1.4, zorder=4)
ax.axhline(2.75, color=GRAY, ls="--", lw=1)
ax.text(2.35, 2.82, "2025 年 2.75 万亿", fontsize=8, color=GRAY)
ax.set_ylabel("2030 年跨境电商进出口（万亿元，中心情景区间）")
ax.set_title("展望 2030：按基准配置能力，按承压做底线，按乐观做期权", fontsize=13, color=BLUE)
ax.set_ylim(0, 6.4)
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.grid(axis="y", color="#E6E8EC", zorder=0)
ax.text(0, -0.18, "区间为方向性研判，不是官方目标或投资建议。件数增长受免税取消抑制，增量更多来自品牌、仓网与新兴市场。",
        transform=ax.transAxes, fontsize=8, color=GRAY)
save(fig, "chart08_2030_scenarios.png")

# 9 政策建议三端
fig, ax = plt.subplots(figsize=(10.2, 4.8))
ax.set_xlim(0, 12)
ax.set_ylim(0, 5)
ax.axis("off")
ax.set_title("政策建议：企业系统、城市空间、国家规则三端发力", fontsize=14, color=BLUE)
cols = [
    (0.3, BLUE, "对企业", "四年交出全球经营系统\n· 合规中台\n· 仓网与库存\n· 品牌与产业\n· 人与空间"),
    (4.2, TEAL, "对城市与园区", "把出海写入空间规划\n· 保障退货与监管空间\n· 园区配人才公寓\n· 总部复合功能\n· 海外资产台账"),
    (8.1, AMBER, "对国家政策", "制度型开放对冲规则封闭\n· 对接 DEPA / CPTPP\n· 中小卖家公共合规\n· 海外仓补功能\n· 安居纳入走出去服务"),
]
for x, c, title, body in cols:
    ax.add_patch(FancyBboxPatch((x, 0.4), 3.6, 4.0, boxstyle="round,pad=0.04,rounding_size=0.12",
                                facecolor=c, edgecolor="none"))
    ax.text(x + 1.8, 3.85, title, ha="center", color="white", fontsize=14, fontweight="bold")
    ax.text(x + 1.8, 2.0, body, ha="center", va="center", color="white", fontsize=11)
save(fig, "chart09_policy.png")

print("all charts done")
