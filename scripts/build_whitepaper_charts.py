# -*- coding: utf-8 -*-
"""《住房产业三链融合白皮书》图表生成脚本。

图表均为课题组根据公开政策目标与行业公开数据绘制的研究性归纳，
不构成官方统计公报。
"""
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
from matplotlib import font_manager

FONT_CANDIDATES = [
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
]
FONT_FAMILY = "Noto Sans CJK SC"
for f in FONT_CANDIDATES:
    if os.path.exists(f):
        font_manager.fontManager.addfont(f)
plt.rcParams.update({
    "font.family": FONT_FAMILY,
    "axes.unicode_minus": False,
    "figure.dpi": 150,
    "savefig.dpi": 160,
    "axes.edgecolor": "#C9CFD8",
    "axes.linewidth": 0.8,
    "axes.titlesize": 15,
    "axes.titleweight": "bold",
    "axes.labelsize": 11,
    "xtick.labelsize": 10,
    "ytick.labelsize": 10,
    "legend.fontsize": 10,
})
BLUE = "#0E4E9B"
RED = "#C8102E"
GRAY = "#8A93A3"
TEAL = "#3AA17E"
ORANGE = "#E8734A"
GOLD = "#F2A65A"
PALETTE = [BLUE, "#2E7CC3", "#5FA8DC", TEAL, RED, ORANGE, GOLD, "#8B6BB8"]
OUT = "/workspace/whitepaper/assets/charts"
os.makedirs(OUT, exist_ok=True)


def note(fig, text):
    fig.text(0.01, 0.01, text, fontsize=8, color=GRAY)


def save(fig, name):
    fig.savefig(os.path.join(OUT, name), bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("saved", name)


def rounded(ax, x, y, w, h, text, fc, ec=None, ts=10, tc="white", lw=0):
    box = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.08",
                         facecolor=fc, edgecolor=ec or fc, linewidth=lw,
                         transform=ax.transAxes, clip_on=False)
    ax.add_patch(box)
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center",
            fontsize=ts, color=tc, transform=ax.transAxes, linespacing=1.35)


# ============ 图1 三链融合架构 ============
fig, ax = plt.subplots(figsize=(11.2, 6.4))
ax.set_xlim(0, 1)
ax.set_ylim(0, 1)
ax.axis("off")
ax.set_title("住房产业三链融合架构", pad=12, color=BLUE, fontsize=16)

layers = [
    (0.04, 0.72, 0.92, 0.22, BLUE, "软件层  ·  住房操作系统",
     "BIM / CIM / 数字孪生 / AI 审图与设计 / 家庭 OS 与智能体\n房屋健康档案 · 物业与生活服务平台 · 产业监管与数据身份"),
    (0.04, 0.42, 0.92, 0.22, TEAL, "硬件层  ·  感知与执行",
     "建筑机器人与智能施工装备 / 模块吊装与产线装备\n结构与管线传感器 · 房屋体检设备 · 智能家居终端 · 边缘算力"),
    (0.04, 0.12, 0.92, 0.22, ORANGE, "供应链  ·  产品化总装",
     "绿色建材与部品部件 / 模块化工厂（MiC · CMC）\n集中采购与材料护照 · 物流节拍 · 供应链金融 · 集成商总装"),
]
for x, y, w, h, c, title, body in layers:
    rounded(ax, x, y, w, h, "", c, ts=11)
    ax.text(x + 0.02, y + h - 0.055, title, color="white", fontsize=13,
            fontweight="bold", transform=ax.transAxes)
    ax.text(x + 0.02, y + 0.055, body, color="white", fontsize=10,
            transform=ax.transAxes, linespacing=1.45)
ax.text(0.5, 0.03, "融合结果：可制造、可感知、可运维、可迭代的长期住房产品",
        ha="center", fontsize=11, color=BLUE, fontweight="bold",
        transform=ax.transAxes)
note(fig, "资料来源：课题组绘制。")
save(fig, "chart01_three_chain.png")


# ============ 图2 政策时间轴 ============
fig, ax = plt.subplots(figsize=(11.2, 5.8))
ax.set_xlim(2019.5, 2026.8)
ax.set_ylim(0, 7)
ax.axis("off")
ax.set_title("智能建造与数字住建政策演进（2020—2026）", pad=8, color=BLUE)

events = [
    (2020.6, 1.2, "2020\n十三部门指导意见\n新型建筑工业化", BLUE),
    (2022.8, 3.1, "2022\n24 城智能建造试点", "#2E7CC3"),
    (2024.2, 5.0, "2024\n数字住建布局\n中办国办韧性城市", TEAL),
    (2025.2, 2.2, "2025\n智能建造技术导则\n好房子品质意见", ORANGE),
    (2026.3, 4.4, "2026\n政府工作报告：\n培育现代化建筑产业链", RED),
]
ax.plot([2020, 2026.5], [0.55, 0.55], color="#C9CFD8", lw=4, solid_capstyle="round")
for yr, y, text, c in events:
    ax.plot([yr, yr], [0.55, y], color=c, lw=1.4)
    ax.scatter([yr], [0.55], s=70, color=c, zorder=5)
    ax.text(yr, y + 0.15, text, ha="center", va="bottom", fontsize=9.5,
            color=c, fontweight="bold", linespacing=1.35)
ax.set_xticks([2020, 2022, 2024, 2025, 2026])
ax.tick_params(axis="x", length=0, labelsize=11, colors=GRAY)
ax.spines[:].set_visible(False)
note(fig, "资料来源：公开政策文件，课题组整理。")
save(fig, "chart02_policy_timeline.png")


# ============ 图3 试点成绩单 ============
fig, ax = plt.subplots(figsize=(10.6, 5.4))
items = [
    ("协调机制城市", 24, "座"),
    ("骨干培育企业", 506, "家"),
    ("国家级高新企业", 214, "家"),
    ("示范工程项目", 758, "个"),
    ("相关标准导则", 47, "项"),
    ("开设专业高校", 99, "所"),
]
names = [i[0] for i in items]
vals = [i[1] for i in items]
units = [i[2] for i in items]
colors = [BLUE, BLUE, "#2E7CC3", TEAL, ORANGE, RED]
bars = ax.barh(names[::-1], vals[::-1], color=colors[::-1], height=0.62)
for b, v, u in zip(bars, vals[::-1], units[::-1]):
    ax.text(b.get_width() + 8, b.get_y() + b.get_height() / 2,
            f"{v} {u}", va="center", fontsize=11, color="#333")
ax.set_xlim(0, 900)
ax.set_title("智能建造试点城市公开成绩单（2023 年度口径）")
ax.tick_params(axis="x", colors=GRAY)
note(fig, "资料来源：住建部关于智能建造试点城市 2023 年度工作情况的通报。另：2023 年智能建造专业招生 5539 人，同比 +55%。")
save(fig, "chart03_pilot_scorecard.png")


# ============ 图4 三项制度 ============
fig, ax = plt.subplots(figsize=(11.0, 5.6))
ax.set_xlim(0, 1)
ax.set_ylim(0, 1)
ax.axis("off")
ax.set_title("房屋全生命周期安全“三项制度”与数字档案", pad=10, color=BLUE)
blocks = [
    (0.04, 0.42, 0.28, 0.42, BLUE, "房屋体检",
     "定期检测结构、围护\n与设备管线\n优先房龄 30 年以上\n需要传感与检测装备"),
    (0.36, 0.42, 0.28, 0.42, TEAL, "房屋养老金",
     "个人账户：维修资金\n公共账户：政府筹集\n用于体检与保险\n取之于房、用之于房"),
    (0.68, 0.42, 0.28, 0.42, ORANGE, "房屋保险",
     "质量与安全风险分担\n依赖可核保数据\n与体检结果挂钩\n市场化机制仍在发育"),
]
for x, y, w, h, c, t, b in blocks:
    rounded(ax, x, y, w, h, "", c)
    ax.text(x + w / 2, y + h - 0.07, t, ha="center", color="white",
            fontsize=13, fontweight="bold", transform=ax.transAxes)
    ax.text(x + w / 2, y + 0.16, b, ha="center", color="white",
            fontsize=10, transform=ax.transAxes, linespacing=1.45)
rounded(ax, 0.18, 0.08, 0.64, 0.24, "", "#0E4E9B")
ax.text(0.5, 0.20, "共同底座：一房一档的数字身份\n设计模型 · 材料护照 · 竣工孪生 · 体检记录 · 维修与理赔",
        ha="center", va="center", color="white", fontsize=11,
        transform=ax.transAxes, linespacing=1.5)
note(fig, "资料来源：《十五五》规划纲要及住建部公开口径，课题组绘制。2025 年起 42 城开展全生命周期安全管理试点。")
save(fig, "chart04_three_systems.png")


# ============ 图5 模块化效率 ============
fig, ax = plt.subplots(figsize=(10.4, 5.6))
labels = ["工期", "现场用工", "建筑垃圾", "单模块吊装时间"]
trad = [100, 100, 100, 100]
mic = [50, 30, 20, 12.5]  # 40min -> 5min = 12.5%
x = range(len(labels))
w = 0.36
b1 = ax.bar([i - w / 2 for i in x], trad, w, color="#C9CFD8", label="传统现浇（基准=100）")
b2 = ax.bar([i + w / 2 for i in x], mic, w, color=BLUE, label="模块化智能产线（公开典型值）")
for bar, v in zip(b2, mic):
    ax.text(bar.get_x() + bar.get_width() / 2, v + 2, f"{v:.0f}",
            ha="center", fontsize=10, color=BLUE, fontweight="bold")
ax.set_xticks(list(x))
ax.set_xticklabels(labels)
ax.set_ylim(0, 125)
ax.set_ylabel("相对传统现浇的指数")
ax.set_title("模块化建造相对传统现浇的效率对照（公开典型区间）")
ax.legend(frameon=False, loc="upper right")
note(fig, "资料来源：2026 年 CMC/MiC 产线与项目公开报道的典型口径（工期约 -50%、用工约 -70%、垃圾约 -80%、吊装 40 分钟→5 分钟）。用于结构对照，非精确普查。")
save(fig, "chart05_mic_efficiency.png")


# ============ 图6 家庭智能演进 ============
fig, ax = plt.subplots(figsize=(11.2, 4.8))
ax.set_xlim(0, 1)
ax.set_ylim(0, 1)
ax.axis("off")
ax.set_title("家庭智能：从单品联网到空间智能体", pad=8, color=BLUE)
stages = [
    (0.03, "2014—2018", "单品智能", "App 遥控\n品牌孤岛", "#93C6E8"),
    (0.27, "2019—2023", "全屋场景", "中控与场景包\n配网仍脆弱", "#5FA8DC"),
    (0.51, "2024—2026", "互联重构", "Matter 1.6\n智家统一互联", BLUE),
    (0.75, "2027—2030", "空间智能体", "主动服务\n可退出、可适老", RED),
]
for x, year, title, body, c in stages:
    rounded(ax, x, 0.18, 0.22, 0.62, "", c)
    ax.text(x + 0.11, 0.68, year, ha="center", color="white", fontsize=9,
            transform=ax.transAxes)
    ax.text(x + 0.11, 0.54, title, ha="center", color="white", fontsize=13,
            fontweight="bold", transform=ax.transAxes)
    ax.text(x + 0.11, 0.34, body, ha="center", color="white", fontsize=10,
            transform=ax.transAxes, linespacing=1.4)
    if x < 0.7:
        ax.annotate("", xy=(x + 0.235, 0.49), xytext=(x + 0.215, 0.49),
                    xycoords=ax.transAxes, textcoords=ax.transAxes,
                    arrowprops=dict(arrowstyle="->", color=GRAY, lw=1.6))
note(fig, "资料来源：Matter 1.6（2026 年 6 月）、AWE 2026《智家统一互联标准》及企业公开产品，课题组归纳。")
save(fig, "chart06_smarthome_evolution.png")


# ============ 图7 软件四层 ============
fig, ax = plt.subplots(figsize=(11.0, 5.8))
ax.set_xlim(0, 1)
ax.set_ylim(0, 1)
ax.axis("off")
ax.set_title("住房软件的四层操作系统", pad=10, color=BLUE)
rows = [
    (0.78, RED, "城市级 CIM", "住房、小区、管网、电梯进入同一空间底座，支撑体检、更新与治理"),
    (0.56, ORANGE, "产业级监管与数据", "资金监管、材料护照、数据资产、保险核保模型"),
    (0.34, TEAL, "企业级 ERP / 产业互联网", "工厂节拍、采购、进度、质量、供应链协同"),
    (0.12, BLUE, "项目级 BIM / 数字孪生", "正向设计、碰撞检测、竣工模型、入住后 IoT 回流"),
]
for y, c, t, b in rows:
    rounded(ax, 0.06, y, 0.88, 0.18, "", c)
    ax.text(0.10, y + 0.11, t, color="white", fontsize=13, fontweight="bold",
            transform=ax.transAxes)
    ax.text(0.10, y + 0.045, b, color="white", fontsize=10.5,
            transform=ax.transAxes)
note(fig, "资料来源：2026 工程数智大会公开论述及上海“数字住建 4321”框架，课题组绘制。")
save(fig, "chart07_software_stack.png")


# ============ 图8 生命周期 ============
fig, ax = plt.subplots(figsize=(11.2, 5.2))
ax.set_xlim(0, 1)
ax.set_ylim(0, 1)
ax.axis("off")
ax.set_title("住房全生命周期中的三链嵌入", pad=8, color=BLUE)
phases = ["策划设计", "工厂制造", "现场总装", "交付入住", "运维体检", "更新拆除"]
xs = [0.04, 0.20, 0.36, 0.52, 0.68, 0.84]
for x, p in zip(xs, phases):
    rounded(ax, x, 0.72, 0.14, 0.16, p, BLUE, ts=10)
    if x < 0.8:
        ax.annotate("", xy=(x + 0.155, 0.80), xytext=(x + 0.135, 0.80),
                    xycoords=ax.transAxes, textcoords=ax.transAxes,
                    arrowprops=dict(arrowstyle="->", color=GRAY, lw=1.4))
embeds = [
    ("供应链", ORANGE, ["户型模块与\n材料护照", "产线节拍与\n部品入库", "物流与吊装\n窗口", "竣工清册\n移交", "维修部品\n可替换", "材料回收\n再制造"]),
    ("硬件", TEAL, ["传感点位\n预埋设计", "焊接检测\n机器人", "智能吊装\n工地机器人", "家庭终端\n调试验收", "结构监测\n巡检装备", "低扰动\n拆解装备"]),
    ("软件", BLUE, ["BIM 正向\nAI 审图", "二维码身份\nMES", "数字孪生\n工地", "一房一档\n家庭 OS", "工单·保险\nCIM", "拆除模型\n碳账本"]),
]
for i, (name, c, cells) in enumerate(embeds):
    y = 0.50 - i * 0.20
    ax.text(0.015, y + 0.06, name, fontsize=10, color=c, fontweight="bold",
            rotation=90, va="center", transform=ax.transAxes)
    for x, cell in zip(xs, cells):
        rounded(ax, x, y, 0.14, 0.16, cell, "#F4F7FB", ec=c, ts=8.5, tc="#222", lw=1.2)
note(fig, "资料来源：课题组绘制。")
save(fig, "chart08_lifecycle.png")


# ============ 图9 上海角色 ============
fig, ax = plt.subplots(figsize=(10.8, 5.6))
ax.set_xlim(0, 1)
ax.set_ylim(0, 1)
ax.axis("off")
ax.set_title("上海在全国住房三链融合中的角色", pad=8, color=BLUE)
items = [
    (0.06, 0.55, "存量更新需求", "十五五约 3000 万㎡\n老旧小区改造"),
    (0.38, 0.55, "现场样本", "田林路 1044 户\n混凝土 MiC 交付"),
    (0.70, 0.55, "地方标准", "混凝土模块化\n建筑技术导则"),
    (0.06, 0.12, "数字底座", "数字住建 4321\n开放 CIM"),
    (0.38, 0.12, "产业配套", "长三角模块工厂\n软件与机器人"),
    (0.70, 0.12, "政策实验室", "跟踪入住后体验\n档案与物业接驳"),
]
for x, y, t, b in items:
    rounded(ax, x, y, 0.26, 0.32, "", BLUE)
    ax.text(x + 0.13, y + 0.23, t, ha="center", color="white", fontsize=12,
            fontweight="bold", transform=ax.transAxes)
    ax.text(x + 0.13, y + 0.10, b, ha="center", color="white", fontsize=10,
            transform=ax.transAxes, linespacing=1.4)
note(fig, "资料来源：上海市规划与住建公开文件、田林路项目 2026 年 6 月公开报道，课题组绘制。")
save(fig, "chart09_shanghai_role.png")


# ============ 图10 2030 展望 ============
fig, ax = plt.subplots(figsize=(10.8, 5.8))
indicators = [
    "保障房工业化总装渗透",
    "政府投资项目 BIM 全过程",
    "30 年以上住房体检覆盖",
    "新房数字家庭互操作达标",
    "房屋健康档案完整率",
    "危繁脏重工序机器人化",
]
now = [18, 35, 12, 22, 8, 15]
future = [70, 85, 75, 80, 70, 55]
y = range(len(indicators))
ax.barh([i + 0.18 for i in y], future[::-1], 0.34, color="#93C6E8",
        label="2030 情景（课题组）")
ax.barh([i - 0.18 for i in y], now[::-1], 0.34, color=BLUE, label="2026 年约数（研究判断）")
ax.set_yticks(list(y))
ax.set_yticklabels(indicators[::-1])
ax.set_xlim(0, 100)
ax.set_xlabel("%")
ax.set_title("面向 2030 年的技术与制度渗透展望（情景判断）")
ax.legend(frameon=False, loc="lower right")
note(fig, "说明：2026 年为课题组对当前渗透的粗估，2030 年为基于现行政策目标的情景，而非官方预测或承诺。")
save(fig, "chart10_2030_outlook.png")


# ============ 图11 好房子对应 ============
fig, ax = plt.subplots(figsize=(10.8, 5.6))
ax.set_xlim(0, 1)
ax.set_ylim(0, 1)
ax.axis("off")
ax.set_title("好房子四维目标与三链能力对应", pad=8, color=BLUE)
quads = [
    (0.06, 0.52, "安全", RED, "结构监测 · 房屋体检\n燃气消防传感 · 保险核保\n隐蔽工程可追溯"),
    (0.52, 0.52, "舒适", BLUE, "隔声防水材料护照\n户型模块与可变隔断\n适老控件与无障碍"),
    (0.06, 0.08, "绿色", TEAL, "绿色建材认证\n能效等级与光储直柔\n垃圾减量与材料回收"),
    (0.52, 0.08, "智慧", ORANGE, "BIM/CIM 一模到底\n家庭 OS 与空间智能体\n可退出、可互操作"),
]
for x, y, t, c, b in quads:
    rounded(ax, x, y, 0.42, 0.38, "", "#F4F7FB", ec=c, lw=2, tc="#222")
    ax.text(x + 0.04, y + 0.28, t, color=c, fontsize=16, fontweight="bold",
            transform=ax.transAxes)
    ax.text(x + 0.04, y + 0.10, b, color="#333", fontsize=11,
            transform=ax.transAxes, linespacing=1.45)
note(fig, "资料来源：住建部《关于提升住房品质的意见》四维目标，课题组对应三链能力绘制。")
save(fig, "chart11_good_house.png")


# ============ 图12 政策矩阵 ============
fig, ax = plt.subplots(figsize=(10.6, 6.0))
ax.set_xlim(0, 10)
ax.set_ylim(0, 10)
ax.set_xlabel("实施紧迫性  →", fontsize=11)
ax.set_ylabel("公共价值（民生 · 安全 · 可核查）  →", fontsize=11)
ax.set_title("政策建议优先级矩阵")
points = [
    (8.6, 9.0, "1 身份接口\n一房一档", RED),
    (8.2, 7.6, "2 互联接口", ORANGE),
    (7.4, 8.6, "3 安全接口", RED),
    (7.8, 6.4, "4 保障房头雁", BLUE),
    (6.2, 6.8, "5 更新工具箱", BLUE),
    (6.8, 8.0, "6 体验门槛\n6633", TEAL),
    (5.4, 5.6, "7 承认集成商", "#2E7CC3"),
    (5.0, 6.6, "8 供应链金融", GOLD),
    (6.0, 7.6, "9 养老金需求侧", TEAL),
    (4.2, 5.2, "10 上海标准化", "#8B6BB8"),
    (3.6, 7.2, "11 CIM 开放", "#8B6BB8"),
    (3.2, 4.4, "12 三链人才培养", GRAY),
]
ax.axhline(5, color="#E6E9EE", lw=1)
ax.axvline(5, color="#E6E9EE", lw=1)
ax.fill_between([5, 10], 5, 10, color="#F4F7FB", zorder=0)
for x, y, t, c in points:
    ax.scatter([x], [y], s=90, color=c, zorder=3)
    ax.text(x + 0.15, y + 0.15, t, fontsize=8.5, color="#222", linespacing=1.2)
ax.set_xticks([])
ax.set_yticks([])
note(fig, "说明：位置为课题组对第八章 12 条建议的相对排序，供讨论使用。")
save(fig, "chart12_policy_matrix.png")

print("all charts done")
