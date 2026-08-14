# -*- coding: utf-8 -*-
"""绘制 2026 Google 开发者大会观察白皮书配套图表。"""
import json
import os
from collections import Counter

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib import font_manager
from matplotlib.patches import FancyBboxPatch

FONT_PATH = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
font_manager.fontManager.addfont(FONT_PATH)
plt.rcParams["font.sans-serif"] = ["WenQuanYi Micro Hei"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["figure.facecolor"] = "white"
plt.rcParams["axes.facecolor"] = "white"
plt.rcParams["axes.edgecolor"] = "#D0D5DD"
plt.rcParams["axes.grid"] = False
plt.rcParams["savefig.dpi"] = 180
plt.rcParams["savefig.bbox"] = "tight"

BLUE = "#0E4E9B"
RED = "#C8102E"
GOLD = "#C9A227"
TEAL = "#1A7A6D"
ORANGE = "#D9762C"
PURPLE = "#5B4B8A"
GRAY = "#5B616B"
LIGHT = "#E8EEF6"
COLORS = [BLUE, TEAL, ORANGE, PURPLE, GOLD, RED, "#3D7EA6", "#7A9D3C"]

OUT = "/workspace/whitepaper/assets/charts"
os.makedirs(OUT, exist_ok=True)
DATA = "/workspace/whitepaper/data"


def save(fig, name):
    path = os.path.join(OUT, name)
    fig.savefig(path, facecolor="white")
    plt.close(fig)
    print("saved", path)


def caption_style(ax, title):
    ax.set_title(title, fontsize=14, color=BLUE, pad=12, fontweight="bold")
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)


# 图1 2026 Google 开发者日历
def chart01():
    fig, ax = plt.subplots(figsize=(11.2, 4.8))
    items = [
        (0.5, "4月18日", "Gemma 4\n黑客松启动"),
        (2.0, "5月", "Google I/O 2026\n智能体 Gemini 时代"),
        (3.5, "6月13日", "京沪深三城\n黑客松半决赛"),
        (5.0, "7月14日", "上海站报名\n通道开启"),
        (6.5, "7月17–20日", "WAIC 2026\n同址世博片区"),
        (8.0, "8月12–13日", "I/O Connect China\n上海世博中心"),
    ]
    ax.plot([0.3, 8.4], [0.55, 0.55], color=BLUE, lw=3, zorder=1)
    for i, (x, date, label) in enumerate(items):
        color = RED if i == 5 else BLUE
        ax.scatter([x], [0.55], s=160, color=color, zorder=3)
        ax.text(x, 0.78, date, ha="center", va="bottom", fontsize=10,
                color=color, fontweight="bold")
        ax.text(x, 0.28, label, ha="center", va="top", fontsize=9.5, color=GRAY)
    ax.set_xlim(0, 8.8)
    ax.set_ylim(0, 1.15)
    ax.axis("off")
    caption_style(ax, "图1  2026 年 Google 开发者网络关键节点")
    save(fig, "chart01_calendar.png")


# 图2 四大技术板块
def chart02():
    fig, ax = plt.subplots(figsize=(9.5, 5.2))
    labels = ["AI\n智能体与模型", "Android\n跨端应用", "Chrome / Web\n智能体用户", "Cloud\n算力与运行时"]
    values = [32, 26, 22, 20]
    bars = ax.barh(labels[::-1], values[::-1], color=[PURPLE, TEAL, ORANGE, BLUE][::-1], height=0.62)
    ax.set_xlabel("公开议程相对权重（课题组根据官方报道与博客主题编码，%）")
    ax.set_xlim(0, 42)
    for bar, v in zip(bars, values[::-1]):
        ax.text(v + 0.6, bar.get_y() + bar.get_height() / 2, f"{v}%",
                va="center", fontsize=11, color=BLUE, fontweight="bold")
    caption_style(ax, "图2  大会四大技术板块的公开叙事权重")
    save(fig, "chart02_tracks.png")


# 图3 全栈 AI 层级
def chart03():
    fig, ax = plt.subplots(figsize=(10.6, 5.6))
    layers = [
        (0.08, 0.78, 0.84, 0.14, "产品与平台层", "Play / Search / Chrome / Android / 出海加速器", GOLD),
        (0.08, 0.60, 0.84, 0.14, "开发与编排层", "AI Studio · Antigravity 2.0 · 托管式智能体 · Agent Studio", ORANGE),
        (0.08, 0.42, 0.84, 0.14, "模型层", "Gemini 3.6 Flash / 3.5 Flash-Lite · Gemma 4 / 12B · DiffusionGemma", TEAL),
        (0.08, 0.24, 0.84, 0.14, "基础设施层", "TPU 8t（训练）· TPU 8i（推理）· Boardfly 拓扑 · Agent Runtime", BLUE),
        (0.08, 0.06, 0.84, 0.14, "安全与治理层", "IntentSanitizer · SynthID / Content Credentials · 远程沙盒", PURPLE),
    ]
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    for x, y, w, h, title, sub, c in layers:
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.012,rounding_size=0.02",
                                    facecolor=c, edgecolor="none", alpha=0.92))
        ax.text(x + 0.025, y + h * 0.62, title, color="white", fontsize=13, fontweight="bold")
        ax.text(x + 0.025, y + h * 0.22, sub, color="white", fontsize=10)
    caption_style(ax, "图3  Google 全栈式 AI 创新路线（据 2026 年 8 月 12 日官方博客整理）")
    save(fig, "chart03_fullstack.png")


# 图4 Android 三项更新
def chart04():
    fig, ax = plt.subplots(figsize=(9.8, 4.8))
    names = ["Android CLI\n+ 智能体", "Android Studio\n智能体模式", "Android Bench\n评估工具"]
    metric = ["词元消耗约降 70%\n任务完成最多提速 3 倍", "可选远程或本地模型\n支持离线深度开发", "衡量成本、延迟\n与真实开发挑战"]
    xs = [0.18, 0.50, 0.82]
    colors = [BLUE, TEAL, ORANGE]
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    for x, name, m, c in zip(xs, names, metric, colors):
        circ = plt.Circle((x, 0.62), 0.12, color=c)
        ax.add_patch(circ)
        ax.text(x, 0.62, "AI", ha="center", va="center", color="white", fontsize=12, fontweight="bold")
        ax.text(x, 0.38, name, ha="center", va="top", fontsize=12, color=c, fontweight="bold")
        ax.text(x, 0.16, m, ha="center", va="top", fontsize=9.5, color=GRAY)
    caption_style(ax, "图4  Android 开发工具三项核心更新")
    save(fig, "chart04_android.png")


# 图5 黑客松四赛道
def chart05():
    fig, ax = plt.subplots(figsize=(8.6, 8.6))
    sizes = [25, 25, 25, 25]
    labels = [
        "Track A\nAI Agent\n多步规划 / 工具调用",
        "Track B\nMultimodal\n视听语一体化",
        "Track C\nEdge AI\n端侧离线部署",
        "Track D\nAI for Social Good\n无障碍 / 气候 / 公卫",
    ]
    wedges, _ = ax.pie(sizes, colors=[BLUE, TEAL, ORANGE, PURPLE], startangle=90,
                       wedgeprops=dict(width=0.42, edgecolor="white", linewidth=3))
    ax.text(0, 0, "Gemma 4\nHackathon", ha="center", va="center", fontsize=13,
            color=BLUE, fontweight="bold")
    ax.legend(wedges, labels, loc="center left", bbox_to_anchor=(0.95, 0.5),
              frameon=False, fontsize=10)
    ax.set_title("图5  Gemma 4 黑客松四条核心赛道", fontsize=14, color=BLUE,
                 pad=16, fontweight="bold")
    save(fig, "chart05_hackathon_tracks.png")


# 图6 入围项目类型
def chart06():
    with open(os.path.join(DATA, "hackathon_finalists.json"), encoding="utf-8") as f:
        data = json.load(f)
    counts = Counter(x["category"] for x in data["finalists"])
    order = ["适老化", "医疗健康", "教育", "端侧隐私", "社会公益", "空间智能", "通用应用"]
    labels = [k for k in order if k in counts]
    vals = [counts[k] for k in labels]
    fig, ax = plt.subplots(figsize=(9.6, 5.0))
    bars = ax.bar(labels, vals, color=COLORS[:len(labels)], width=0.62)
    ax.set_ylabel("入围项目数（件）")
    ax.set_ylim(0, max(vals) + 1.4)
    for b, v in zip(bars, vals):
        ax.text(b.get_x() + b.get_width() / 2, v + 0.12, str(v),
                ha="center", fontsize=11, color=BLUE, fontweight="bold")
    caption_style(ax, "图6  Gemma 4 黑客松公开入围项目的场景类型（n=15）")
    save(fig, "chart06_finalist_types.png")


# 图7 TPU 对照
def chart07():
    fig, ax = plt.subplots(figsize=(10.2, 4.6))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.axis("off")
    boxes = [
        (0.4, 1.2, 4.3, 3.6, BLUE, "TPU 8t · 训练引擎",
         ["专为大规模预训练", "单 Pod 算力接近上一代近 3 倍", "块级乘法内置 MXU", "原生量化降低 VPU 开销"]),
        (5.3, 1.2, 4.3, 3.6, TEAL, "TPU 8i · 推理引擎",
         ["专为推理与强化学习", "Boardfly 拓扑网络", "支撑百万级智能体近零延迟并发", "强调时延与能效"]),
    ]
    for x, y, w, h, c, title, bullets in boxes:
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.08,rounding_size=0.15",
                                    facecolor=c, edgecolor="none", alpha=0.92))
        ax.text(x + 0.25, y + h - 0.55, title, color="white", fontsize=13, fontweight="bold")
        for i, b in enumerate(bullets):
            ax.text(x + 0.25, y + h - 1.15 - i * 0.55, "•  " + b, color="white", fontsize=10.5)
    caption_style(ax, "图7  第八代 TPU 的训练 / 推理双芯片分工")
    save(fig, "chart07_tpu.png")


# 图8 开发者角色变迁
def chart08():
    fig, ax = plt.subplots(figsize=(10.4, 4.4))
    stages = ["写每一行代码", "调用模型 API", "编排单个智能体", "定义规范并\n管理智能体队列"]
    years = ["2023 前", "2024–2025", "2025–2026", "I/O Connect\n2026 叙事"]
    xs = [1, 2, 3, 4]
    ax.plot(xs, [1, 1, 1, 1], color=BLUE, lw=3)
    for x, s, y in zip(xs, stages, years):
        ax.scatter([x], [1], s=220, color=BLUE if x < 4 else RED, zorder=3)
        ax.text(x, 1.08, s, ha="center", va="bottom", fontsize=11, color=BLUE, fontweight="bold")
        ax.text(x, 0.90, y, ha="center", va="top", fontsize=9.5, color=GRAY)
    ax.set_xlim(0.4, 4.6)
    ax.set_ylim(0.7, 1.28)
    ax.axis("off")
    caption_style(ax, "图8  官方叙事中的开发者角色变迁：从编码者到智能体管理者")
    save(fig, "chart08_role_shift.png")


# 图9 世博片区连续大会对照
def chart09():
    fig, ax = plt.subplots(figsize=(10.8, 5.4))
    cols = ["对照维度", "WAIC 2026", "Google I/O Connect China 2026"]
    rows = [
        ["时间", "7 月 17–20 日", "8 月 12–13 日"],
        ["主场馆", "世博中心 / 世博展览馆等", "上海世博中心"],
        ["公开规模", "963 家参展商、175 场论坛", "近 2,000 名开发者"],
        ["连接对象", "政府、产业、治理多边", "全球开发者与出海团队"],
        ["技术主线", "具身智能 + 产业落地", "智能体全栈 + 跨海出圈"],
        ["空间含义", "园区 / 产线 / 场景", "人才 / 工具链 / 全球接口"],
    ]
    ax.axis("off")
    table = ax.table(cellText=rows, colLabels=cols, loc="center", cellLoc="center")
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.15, 1.55)
    for (r, c), cell in table.get_celld().items():
        cell.set_edgecolor("#D0D5DD")
        if r == 0:
            cell.set_facecolor(BLUE)
            cell.set_text_props(color="white", fontweight="bold")
        elif c == 0:
            cell.set_facecolor(LIGHT)
            cell.set_text_props(fontweight="bold", color=BLUE)
        else:
            cell.set_facecolor("white")
    ax.set_title("图9  同一世博片区、一个月内两场全球科技大会的功能分工",
                 fontsize=13, color=BLUE, pad=8, fontweight="bold")
    save(fig, "chart09_expo_compare.png")


# 图10 四大支柱
def chart10():
    fig, ax = plt.subplots(figsize=(10.2, 5.0))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.axis("off")
    items = [
        (0.4, 3.2, "本土社区连接", "GDG / GDE\n京沪深杭工作坊"),
        (2.7, 3.2, "技术导师指导", "现场工作坊\n黑客松 Office Hours"),
        (5.0, 3.2, "出海创业加速", "初创出海计划\n跨境电商加速中心"),
        (7.3, 3.2, "中文技术文档", "本地化文档\nDeveloper Program"),
    ]
    for x, y, title, sub in items:
        ax.add_patch(FancyBboxPatch((x, y - 1.6), 2.1, 2.6,
                                    boxstyle="round,pad=0.06,rounding_size=0.12",
                                    facecolor=LIGHT, edgecolor=BLUE, linewidth=1.6))
        ax.text(x + 1.05, y + 0.55, title, ha="center", fontsize=12, color=BLUE, fontweight="bold")
        ax.text(x + 1.05, y - 0.55, sub, ha="center", fontsize=10, color=GRAY)
    ax.annotate("", xy=(9.4, 4.6), xytext=(0.5, 4.6),
                arrowprops=dict(arrowstyle="->", color=RED, lw=2))
    ax.text(5, 4.85, "把本土创新接到全球规模", ha="center", color=RED, fontsize=12, fontweight="bold")
    caption_style(ax, "图10  Google 连接中国开发者的四大支柱")
    save(fig, "chart10_four_pillars.png")


# 图11 空间含义雷达
def chart11():
    import numpy as np
    labels = ["会展承载力", "短期住宿弹性", "人才住房适配", "灵活办公供给",
              "端侧算力入户", "全球接口功能"]
    values = [86, 72, 58, 64, 47, 81]
    values += values[:1]
    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    angles += angles[:1]
    fig, ax = plt.subplots(figsize=(7.2, 7.2), subplot_kw=dict(polar=True))
    ax.plot(angles, values, color=BLUE, lw=2)
    ax.fill(angles, values, color=BLUE, alpha=0.18)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels, fontsize=11)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=8, color=GRAY)
    ax.set_ylim(0, 100)
    ax.set_title("图11  智能体大会对城市空间系统的冲击强度（课题组观察评分）",
                 fontsize=12, color=BLUE, pad=18, fontweight="bold")
    save(fig, "chart11_space_radar.png")


# 图12 三层部署
def chart12():
    fig, ax = plt.subplots(figsize=(10.4, 4.8))
    layers = [
        (0.5, 1.1, 2.8, 2.8, BLUE, "云端智能体", "Gemini Spark\n24/7 虚拟机\n长程后台任务"),
        (3.7, 1.1, 2.8, 2.8, TEAL, "托管式智能体", "一次 API 调用\n预配智能体\n+ 远程沙盒"),
        (6.9, 1.1, 2.8, 2.8, ORANGE, "端侧智能体", "Gemma 4 本地运行\n笔记本 / 手机\n离线与隐私优先"),
    ]
    ax.set_xlim(0, 10.2)
    ax.set_ylim(0, 5)
    ax.axis("off")
    for x, y, w, h, c, title, sub in layers:
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.08,rounding_size=0.16",
                                    facecolor=c, edgecolor="none"))
        ax.text(x + w / 2, y + h - 0.55, title, ha="center", color="white",
                fontsize=14, fontweight="bold")
        ax.text(x + w / 2, y + 1.05, sub, ha="center", color="white", fontsize=11)
    caption_style(ax, "图12  智能体部署的三层结构：云端 / 托管 / 端侧")
    save(fig, "chart12_deploy_layers.png")


# 图13 I/O Connect 巡回
def chart13():
    fig, ax = plt.subplots(figsize=(10.2, 4.6))
    cities = ["柏林\nBerlin", "班加罗尔\nBengaluru", "上海\nShanghai"]
    notes = ["欧洲开发者枢纽\n工业软件与治理语境", "南亚开发者枢纽\n工程人才密度高", "中国与亚太接口\n出海与智能体落地"]
    xs = [1.5, 5.1, 8.7]
    ax.plot([1.5, 8.7], [2.4, 2.4], color=BLUE, lw=3)
    for x, city, note, c in zip(xs, cities, notes, [PURPLE, TEAL, RED]):
        ax.scatter([x], [2.4], s=280, color=c, zorder=3)
        ax.text(x, 3.15, city, ha="center", fontsize=13, color=c, fontweight="bold")
        ax.text(x, 1.35, note, ha="center", fontsize=10, color=GRAY)
    ax.set_xlim(0, 10.2)
    ax.set_ylim(0.4, 4.2)
    ax.axis("off")
    caption_style(ax, "图13  2026 Google I/O Connect 全球巡回的三城结构")
    save(fig, "chart13_tour.png")


# 图14 人才住房政策含义
def chart14():
    fig, ax = plt.subplots(figsize=(10.6, 5.2))
    labels = ["会展高峰\n短期住宿", "青年开发者\n租赁住房", "一人公司/\n灵活办公", "出海团队\n多时区居住", "端侧入户\n家庭算力", "高校周边\n社区配套"]
    now = [78, 55, 48, 42, 28, 50]
    future = [82, 70, 74, 66, 58, 68]
    x = range(len(labels))
    w = 0.36
    ax.bar([i - w / 2 for i in x], now, width=w, color=BLUE, label="当前冲击（2026）")
    ax.bar([i + w / 2 for i in x], future, width=w, color=ORANGE, label="2027–2028 潜在上升")
    ax.set_xticks(list(x))
    ax.set_xticklabels(labels, fontsize=9.5)
    ax.set_ylabel("政策关注优先级（0–100）")
    ax.set_ylim(0, 100)
    ax.legend(frameon=False, loc="upper right")
    caption_style(ax, "图14  智能体开发者生态对住房与空间政策的关注优先级")
    save(fig, "chart14_housing_priority.png")


if __name__ == "__main__":
    chart01()
    chart02()
    chart03()
    chart04()
    chart05()
    chart06()
    chart07()
    chart08()
    chart09()
    chart10()
    chart11()
    chart12()
    chart13()
    chart14()
    print("all charts done")
