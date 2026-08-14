# -*- coding: utf-8 -*-
"""生成白皮书配图。运行：python3 scripts/build_charts.py"""

from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import FancyBboxPatch

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "whitepaper" / "assets" / "charts"
OUT.mkdir(parents=True, exist_ok=True)

FONT = "WenQuanYi Micro Hei"
for p in font_manager.findSystemFonts():
    if "wqy" in p.lower() or "microhei" in p.lower() or "droid" in p.lower():
        font_manager.fontManager.addfont(p)

plt.rcParams["font.sans-serif"] = [FONT, "WenQuanYi Micro Hei", "Droid Sans Fallback"]
plt.rcParams["axes.unicode_minus"] = False

NAVY = "#1A2A4A"
RED = "#8B1A1A"
OCHRE = "#C45C26"
SAGE = "#3D6B5A"
GOLD = "#B8860B"
BG = "#F7F4EF"
CARD = "#FFFFFF"
GRAY = "#5A646E"


def _save(fig, name):
    path = OUT / name
    fig.savefig(path, dpi=180, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    print(f"wrote {path}")


def chart_architecture():
    fig, ax = plt.subplots(figsize=(11.2, 6.4), facecolor=BG)
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 7)
    ax.axis("off")

    # base housing layer
    ax.add_patch(FancyBboxPatch((0.4, 0.35), 11.2, 1.15, boxstyle="round,pad=0.04,rounding_size=0.12",
                                facecolor=NAVY, edgecolor="none"))
    ax.text(6.0, 0.92, "住房 / 社区 / 园区空间底座", ha="center", va="center", color="white", fontsize=12, fontweight="bold")
    ax.text(6.0, 0.58, "好房子 · 完整社区 · 一刻钟生活圈 · 产业园区职住服务 · 数字家庭与边缘算力",
            ha="center", va="center", color="#D7DEE8", fontsize=8.5)

    services = [
        (0.5, 2.0, 3.5, 2.4, RED, "医疗服务", "家庭健康监测\n社区卫生站 / 家庭病床\n县域医共体智能辅助\n急救“最后一米”"),
        (4.25, 2.0, 3.5, 2.4, OCHRE, "养老服务", "家庭养老床位\n嵌入式日间照料\n机构专业照护\n长护险支付闭环"),
        (8.0, 2.0, 3.5, 2.4, SAGE, "提效服务", "工业 / 仓储具身作业\n社区物流与物业\n家政与康养辅助\n人机协同排程"),
    ]
    for x, y, w, h, c, title, body in services:
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.04,rounding_size=0.12",
                                    facecolor=CARD, edgecolor=c, linewidth=1.8))
        ax.add_patch(FancyBboxPatch((x, y + h - 0.55), w, 0.55, boxstyle="round,pad=0.02,rounding_size=0.08",
                                    facecolor=c, edgecolor="none"))
        ax.text(x + w / 2, y + h - 0.28, title, ha="center", va="center", color="white", fontsize=12, fontweight="bold")
        ax.text(x + w / 2, y + 1.05, body, ha="center", va="center", color=NAVY, fontsize=9, linespacing=1.45)

    ax.add_patch(FancyBboxPatch((0.5, 4.85), 11.0, 1.55, boxstyle="round,pad=0.04,rounding_size=0.12",
                                facecolor="#EFE8DC", edgecolor=GOLD, linewidth=1.4))
    ax.text(6.0, 5.95, "科技与制度赋能层（2026—2030）", ha="center", va="center", color=NAVY, fontsize=11, fontweight="bold")
    ax.text(6.0, 5.35, "人工智能+  ·  医疗/养老智能体  ·  具身机器人  ·  长护险  ·  医养结合标准  ·  城市更新与存量盘活",
            ha="center", va="center", color=GRAY, fontsize=9)
    _save(fig, "chart01_architecture.png")


def chart_timeline():
    fig, ax = plt.subplots(figsize=(11.2, 6.6), facecolor=BG)
    ax.set_xlim(2024.6, 2031.2)
    ax.set_ylim(0, 8)
    ax.axis("off")

    ax.plot([2025, 2030.8], [4.0, 4.0], color=NAVY, linewidth=2.2, zorder=1)
    years = [2025, 2026, 2027, 2028, 2029, 2030]
    for y in years:
        ax.scatter([y], [4.0], s=46, color=RED, zorder=3)
        ax.text(y, 3.55, str(y), ha="center", va="top", color=NAVY, fontsize=10, fontweight="bold")

    items = [
        (2025.15, 6.7, "2025.08\n国务院《人工智能+》行动\n2027 普及率超70%，2030 超90%", RED, 1),
        (2025.85, 5.55, "2025.10 医养结合促进行动\n2025.11 AI+医疗卫生实施意见", OCHRE, 1),
        (2026.35, 6.85, "2026.03 长护险全国建制\n2026.05 智能体实施意见\n2026.05 城市更新“十五五”规划", NAVY, 1),
        (2026.55, 1.15, "2026.07 WAIC：康养具身工程化\n郑州众擎T800量产下线\n2026.09 医养结合国标施行", SAGE, 0),
        (2027.4, 5.7, "2027 医养结合县域全覆盖\n智能终端/智能体普及率超70%", GOLD, 1),
        (2028.35, 1.35, "2028 年底\n长护险全国基本全覆盖", OCHRE, 0),
        (2030.0, 6.55, "2030\n基层诊疗智能辅助基本全覆盖\n二级以上医院普遍开展影像与临床辅助\n城市更新取得重要进展", RED, 1),
    ]
    for x, y, text, color, above in items:
        ax.annotate(
            "",
            xy=(x if x < 2030 else 2030, 4.0),
            xytext=(x, y + (0.35 if above else -0.15)),
            arrowprops=dict(arrowstyle="-", color=color, lw=1.1),
        )
        ax.text(x, y, text, ha="center", va="center", fontsize=8, color=NAVY,
                bbox=dict(boxstyle="round,pad=0.35", facecolor=CARD, edgecolor=color, linewidth=1.2))
    _save(fig, "chart02_timeline.png")


def chart_penetration():
    fig, ax = plt.subplots(figsize=(11.0, 5.8), facecolor=BG)
    scenes = ["工业与仓储", "商业与物业", "康养机构", "社区驿站", "家庭居家"]
    y2026 = [28, 16, 12, 8, 4]
    y2030_base = [72, 48, 45, 38, 22]
    y2030_acc = [88, 70, 68, 58, 40]
    x = range(len(scenes))
    w = 0.24
    ax.bar([i - w for i in x], y2026, width=w, color=NAVY, label="2026 现状（示意）")
    ax.bar(list(x), y2030_base, width=w, color=OCHRE, label="2030 基准情景")
    ax.bar([i + w for i in x], y2030_acc, width=w, color=RED, label="2030 加速情景")
    ax.set_xticks(list(x))
    ax.set_xticklabels(scenes, fontsize=10)
    ax.set_ylabel("人机协同服务渗透率（%）", fontsize=10, color=NAVY)
    ax.set_ylim(0, 100)
    ax.set_facecolor(BG)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="y", linestyle="--", alpha=0.35)
    ax.legend(frameon=False, loc="upper right")
    fig.text(0.5, 0.02, "说明：渗透率指该场景中稳定可用的人机协同服务包覆盖比例，用于政策讨论，不替代官方统计。",
             ha="center", fontsize=8, color=GRAY)
    fig.subplots_adjust(bottom=0.16)
    _save(fig, "chart03_penetration.png")


def chart_scenarios():
    fig, ax = plt.subplots(figsize=(11.2, 6.2), facecolor=BG)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.axis("off")

    ax.annotate("", xy=(9.4, 1.6), xytext=(0.8, 1.6), arrowprops=dict(arrowstyle="->", color=NAVY, lw=1.4))
    ax.annotate("", xy=(0.8, 9.1), xytext=(0.8, 1.6), arrowprops=dict(arrowstyle="->", color=NAVY, lw=1.4))
    ax.text(5.1, 1.15, "支付与制度完备程度 →", ha="center", color=GRAY, fontsize=9)
    ax.text(0.35, 5.4, "工程落地与安全可靠性 →", ha="center", va="center", rotation=90, color=GRAY, fontsize=9)

    boxes = [
        (1.2, 5.5, 3.6, 3.2, OCHRE, "基准情景（中心主情景）", "工业与机构先行，社区跟进；\n家庭以监测、订阅、远程兜底为主；\n县域与乡村节奏明显滞后。"),
        (5.4, 5.5, 3.6, 3.2, SAGE, "加速情景", "长护险、医保与机构采购同步；\n标准与质检跟上量产；\n社区驿站成为标配接口。"),
        (1.2, 2.0, 3.6, 3.0, RED, "迟滞情景", "安全事件、支付缺口或数据治理\n引发停摆；技术停留在展厅，\n住房改造投入沉淀为沉没成本。"),
        (5.4, 2.0, 3.6, 3.0, GOLD, "分化情景", "高能级城区形成样板，\n老旧小区与县城出现“服务断崖”；\n住房公平问题被放大。"),
    ]
    for x, y, w, h, c, title, body in boxes:
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.04,rounding_size=0.12",
                                    facecolor=CARD, edgecolor=c, linewidth=1.7))
        ax.text(x + 0.22, y + h - 0.45, title, ha="left", va="center", color=c, fontsize=11, fontweight="bold")
        ax.text(x + 0.22, y + h / 2 - 0.15, body, ha="left", va="center", color=NAVY, fontsize=9, linespacing=1.45)
    _save(fig, "chart04_scenarios.png")


def chart_henan():
    fig, ax = plt.subplots(figsize=(11.0, 5.6), facecolor=BG)
    labels = ["常住人口\n规模压力", "老龄化\n照护压力", "劳动力\n供给压力", "住房存量\n更新压力", "制造体系\n提效压力"]
    values = [88, 82, 76, 70, 84]
    colors = [RED, OCHRE, GOLD, NAVY, SAGE]
    bars = ax.barh(labels[::-1], values[::-1], color=colors[::-1], height=0.62)
    ax.set_xlim(0, 100)
    ax.set_xlabel("对住房—服务体系的结构性压力（示意指数 0–100）", fontsize=10, color=NAVY)
    ax.set_facecolor(BG)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="x", linestyle="--", alpha=0.3)
    for b, v in zip(bars, values[::-1]):
        ax.text(v + 1.2, b.get_y() + b.get_height() / 2, str(v), va="center", fontsize=9, color=NAVY)
    fig.text(0.5, 0.02, "指数用于比较压力相对强度，便于讨论政策优先级，不构成官方统计或预测。",
             ha="center", fontsize=8, color=GRAY)
    fig.subplots_adjust(bottom=0.16)
    _save(fig, "chart05_henan_pressure.png")


def main():
    chart_architecture()
    chart_timeline()
    chart_penetration()
    chart_scenarios()
    chart_henan()


if __name__ == "__main__":
    main()
