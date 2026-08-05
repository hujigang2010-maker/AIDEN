# -*- coding: utf-8 -*-
"""生成《赛普客户：房企十五五竞争的破局关键——客户服务体系化建设》报告配图。

统一视觉规范：
- 主色：深藏青 #1f2a44（标题、主结构）
- 强调色：赛普红 #e94d4f（关键节点、高亮）
- 辅助色：钢蓝 #3d5a80、雾灰蓝 #8d99ae、浅底 #f4f6f8
"""
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Rectangle
from matplotlib import font_manager

# ---------------------------------------------------------------- 字体与全局
FONT_PATH = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
font_manager.fontManager.addfont(FONT_PATH)
CN = font_manager.FontProperties(fname=FONT_PATH).get_name()
plt.rcParams["font.family"] = CN
plt.rcParams["axes.unicode_minus"] = False

NAVY = "#1f2a44"
RED = "#e94d4f"
BLUE = "#3d5a80"
GRAY = "#8d99ae"
LIGHT = "#f4f6f8"
GOLD = "#d9a441"
WHITE = "#ffffff"

OUT = "/workspace/assets"
DPI = 200


def newfig(w, h):
    fig = plt.figure(figsize=(w, h))
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")
    return fig, ax


def box(ax, x, y, w, h, fc=NAVY, ec="none", lw=0, r=0.6, alpha=1.0, z=2):
    p = FancyBboxPatch(
        (x, y), w, h,
        boxstyle=f"round,pad=0,rounding_size={r}",
        fc=fc, ec=ec, lw=lw, alpha=alpha, zorder=z,
    )
    ax.add_patch(p)
    return p


def txt(ax, x, y, s, size=10, color=WHITE, weight="normal", ha="center", va="center", z=5, spacing=1.4):
    ax.text(x, y, s, fontsize=size, color=color, fontweight=weight,
            ha=ha, va=va, zorder=z, linespacing=spacing)


def arrow(ax, x1, y1, x2, y2, color=GRAY, lw=2.2, style="-|>", z=3, ms=14):
    a = FancyArrowPatch((x1, y1), (x2, y2), arrowstyle=style,
                        mutation_scale=ms, color=color, lw=lw, zorder=z)
    ax.add_patch(a)


def save(fig, name):
    fig.savefig(f"{OUT}/{name}", dpi=DPI, facecolor="white", bbox_inches=None)
    plt.close(fig)
    print("saved", name)


# ==================================================== 图1 政策脉络时间轴
def fig_timeline():
    fig, ax = newfig(11.4, 5.0)
    events = [
        ("2023.07", "中央政治局会议", "房地产市场供求关系\n发生重大变化", GRAY),
        ("2024.12", "全国住建工作会议", "增加改善性住房\n特别是好房子供给", GRAY),
        ("2025.03", "政府工作报告", "“好房子”首次写入\n安全·舒适·绿色·智慧", BLUE),
        ("2025.05", "《住宅项目规范》实施", "层高≥3米等强制标准\n“好房子”有了国标", BLUE),
        ("2025.07", "中央城市工作会议", "时隔十年再召开\n城市发展转向存量提质", BLUE),
        ("2025.08", "城市高质量发展意见", "加快构建房地产\n发展新模式", BLUE),
        ("2025.10", "二十届四中全会", "“十五五”规划建议：\n推动房地产高质量发展\n实施物业服务质量提升行动", RED),
    ]
    n = len(events)
    y_axis = 52
    ax.plot([3, 97], [y_axis, y_axis], color=NAVY, lw=3, zorder=2)
    arrow(ax, 97, y_axis, 99.5, y_axis, color=NAVY, lw=3, ms=18)
    for i, (date, title, desc, color) in enumerate(events):
        x = 5.5 + i * (86.5 / (n - 1))
        up = i % 2 == 0
        ax.add_patch(plt.Circle((x, y_axis), 1.05, color=color, zorder=4))
        ax.add_patch(plt.Circle((x, y_axis), 0.45, color=WHITE, zorder=5))
        y_card = 66 if up else 38
        ax.plot([x, x], [y_axis + (1.5 if up else -1.5), y_card - (6.5 if up else -6.5)],
                color=color, lw=1.4, zorder=3)
        card_h = 13
        wcard = 12.4
        box(ax, x - wcard / 2, y_card - card_h / 2, wcard, card_h,
            fc=WHITE, ec=color, lw=1.6, r=1.2, z=4)
        txt(ax, x, y_card + 3.6, title, size=9.2, color=NAVY, weight="bold")
        txt(ax, x, y_card - 2.2, desc, size=7.4, color="#4a5568", spacing=1.35)
        txt(ax, x, y_axis - 6.8 if up else y_axis + 6.8, date, size=10,
            color=color, weight="bold", va="center")
    txt(ax, 3, 90, "从“有没有”到“好不好”：房地产政策脉络（2023—2025）", size=14,
        color=NAVY, weight="bold", ha="left")
    txt(ax, 3, 82.5, "“十五五”规划《建议》首次将“实施物业服务质量提升行动”纳入顶层设计，客户服务上升为国家战略议程",
        size=9.5, color=RED, ha="left")
    save(fig, "fig_timeline.png")


# ==================================================== 图2 市场之变：规模回落
def fig_market():
    fig = plt.figure(figsize=(11.4, 5.2))
    ax = fig.add_axes([0.08, 0.16, 0.66, 0.72])
    years = ["2021", "2022", "2023", "2024", "2025E"]
    area = [17.94, 13.58, 11.17, 9.74, 9.2]
    colors = [GRAY, GRAY, GRAY, BLUE, RED]
    bars = ax.bar(years, area, width=0.56, color=colors, zorder=3)
    for b, v in zip(bars, area):
        ax.text(b.get_x() + b.get_width() / 2, v + 0.35, f"{v:.2f}",
                ha="center", fontsize=11, color=NAVY, fontweight="bold")
    ax.set_ylim(0, 20.5)
    ax.set_ylabel("商品房销售面积（亿平方米）", fontsize=11, color=NAVY)
    ax.spines[["top", "right"]].set_visible(False)
    ax.spines[["left", "bottom"]].set_color(GRAY)
    ax.tick_params(colors=NAVY, labelsize=11)
    ax.yaxis.grid(True, color="#e2e8f0", lw=0.8, zorder=0)
    ax.set_axisbelow(True)
    ax.set_title("全国商品房销售面积较高峰回落近五成", fontsize=14, color=NAVY,
                 fontweight="bold", loc="left", pad=14)
    ax2 = fig.add_axes([0.78, 0.16, 0.20, 0.72])
    ax2.axis("off")
    notes = [
        ("17.94→9.74", "亿平方米\n三年累计降幅约46%", RED),
        ("供求关系", "发生重大变化\n买方市场全面确立", NAVY),
        ("存量时代", "重点城市二手房成交\n占比持续超过新房", BLUE),
    ]
    y = 0.98
    for t, d, c in notes:
        ax2.text(0, y, t, fontsize=13, color=c, fontweight="bold", va="top")
        ax2.text(0, y - 0.09, d, fontsize=9.5, color="#4a5568", va="top", linespacing=1.5)
        y -= 0.36
    fig.text(0.08, 0.035, "资料来源：国家统计局，赛普研究院整理；2025E为前三季度趋势外推的约数", fontsize=8, color=GRAY)
    save(fig, "fig_market.png")


# ==================================================== 图3 14554 蓝图全景
def fig_14554():
    fig, ax = newfig(11.4, 6.6)
    # 顶部：体系目标
    box(ax, 20, 86, 60, 10.5, fc=NAVY, r=1.0)
    txt(ax, 50, 93.4, "1 个体系目标", size=13, weight="bold")
    txt(ax, 50, 88.9, "全周期覆盖 · 全维度满意——构建以客户为中心的服务生态", size=9.5)
    arrow(ax, 50, 85.5, 50, 80.5, color=GRAY, lw=2)
    # 第二层：四维穿透
    box(ax, 4, 68, 41, 12, fc=RED, r=1.0)
    txt(ax, 24.5, 75.6, "4 维穿透", size=12.5, weight="bold")
    txt(ax, 24.5, 71.2, "场景穿透 · 流程穿透 · 岗位穿透 · 标准穿透", size=8.8)
    box(ax, 55, 68, 41, 12, fc=BLUE, r=1.0)
    txt(ax, 75.5, 75.6, "5 链协同", size=12.5, weight="bold")
    txt(ax, 75.5, 71.2, "客关 · 营销 · 物业 · 签约 · 维修（第三方）", size=8.8)
    # 第三层：五大关键举措
    box(ax, 4, 46, 92, 15, fc=LIGHT, ec=NAVY, lw=1.4, r=1.0)
    txt(ax, 50, 57.6, "5 大关键举措", size=12, color=NAVY, weight="bold")
    acts = ["营销前置管控\n6大对客会签节点", "客关定期巡检\n5场景16个动作", "客诉机制优化\n三级处理+自主关单",
            "维修服务重塑\n简路径·优机制·精奖罚", "物业服务提升\n协议化梯度激励"]
    for i, a in enumerate(acts):
        x = 8.5 + i * 18.2
        box(ax, x - 7.9, 46.8, 15.8, 8.2, fc=WHITE, ec=GRAY, lw=1.0, r=0.8)
        txt(ax, x, 50.9, a, size=7.6, color="#2d3748", spacing=1.45)
    # 第四层：四大落地抓手
    box(ax, 4, 26, 92, 13, fc=LIGHT, ec=BLUE, lw=1.4, r=1.0)
    txt(ax, 50, 35.6, "4 项落地抓手", size=12, color=NAVY, weight="bold")
    grips = ["考核“导向标”\n绩效强关联", "培训“加油站”\n能力内化", "计划“导航仪”\n过程可视", "激励“强心剂”\n动力续航"]
    for i, g in enumerate(grips):
        x = 14 + i * 24
        box(ax, x - 10.5, 27, 21, 7.4, fc=WHITE, ec=GRAY, lw=1.0, r=0.8)
        txt(ax, x, 30.7, g, size=7.8, color="#2d3748", spacing=1.45)
    # 底部数字带
    box(ax, 4, 8, 92, 12, fc=NAVY, r=1.0)
    nums = [("5", "大客户接触场景"), ("25", "个客户关注点"), ("104", "个内部服务流程"),
            ("20", "个第三方关键岗位"), ("7", "大开发阶段"), ("16", "个巡检服务动作")]
    for i, (num, lab) in enumerate(nums):
        x = 11.5 + i * 15.4
        txt(ax, x, 16.2, num, size=15, color=RED, weight="bold")
        txt(ax, x, 11.2, lab, size=7.8, color=WHITE)
    txt(ax, 50, 99, "“14554”客户服务体系蓝图", size=15, color=NAVY, weight="bold")
    save(fig, "fig_14554.png")


# ==================================================== 图4 体系目标矩阵
def fig_goal():
    fig, ax = newfig(11.4, 5.2)
    txt(ax, 50, 96, "体系目标：全周期覆盖 × 全维度满意", size=14, color=NAVY, weight="bold")
    # 左：全周期覆盖
    box(ax, 4, 8, 44, 78, fc=LIGHT, ec=NAVY, lw=1.4, r=1.0)
    txt(ax, 26, 80.5, "全周期服务覆盖", size=12, color=NAVY, weight="bold")
    stages = [("前期策划", "服务标准从项目立项开始\n深度植入，服务与产品同步规划"),
              ("中期营造", "建设过程持续维护客户关系\n增强信任感与期待值"),
              ("后期交付", "完善交付体验及售后服务\n建立长期关系维护机制")]
    y = 68
    for idx, (t, d) in enumerate(stages):
        box(ax, 8, y - 12, 36, 14, fc=WHITE, ec=BLUE, lw=1.2, r=0.8)
        txt(ax, 26, y - 2.5, t, size=10.5, color=BLUE, weight="bold")
        txt(ax, 26, y - 8, d, size=8, color="#4a5568", spacing=1.4)
        if idx < len(stages) - 1:
            arrow(ax, 26, y - 12.6, 26, y - 17.4, color=GRAY, lw=1.6, ms=11)
        y -= 20
    # 右：全维度满意
    box(ax, 52, 8, 44, 78, fc=LIGHT, ec=RED, lw=1.4, r=1.0)
    txt(ax, 74, 80.5, "全维度客户满意", size=12, color=RED, weight="bold")
    dims = [("准业主满意", "首次到访起打造极致第一印象与购房体验"),
            ("磨合期业主满意", "交付后关键磨合期提供及时、专业的服务支持"),
            ("稳定期业主满意", "长期陪伴机制，持续提升居住体验与归属感"),
            ("总体·服务·品质满意", "三维度综合评价，确保服务品质全面提升")]
    y = 70
    for t, d in dims:
        box(ax, 56, y - 8.5, 36, 11.5, fc=WHITE, ec=GRAY, lw=1.0, r=0.8)
        txt(ax, 74, y - 0.5, t, size=9.8, color=NAVY, weight="bold")
        txt(ax, 74, y - 5.4, d, size=7.8, color="#4a5568")
        y -= 15.5
    save(fig, "fig_goal.png")


# ==================================================== 图5 客户旅程五大场景
def fig_scenarios():
    fig, ax = newfig(11.4, 5.6)
    txt(ax, 50, 96.5, "场景穿透：客户接触旅程的五大核心场景", size=14, color=NAVY, weight="bold")
    scenes = [
        ("01", "案场到访", "第一印象塑造", "案场环境与接待服务细节\n提升客户到访体验", BLUE),
        ("02", "认购签约", "流程优化与信任建立", "高效、透明的服务流程\n回应客户关注重点", BLUE),
        ("03", "等待期关怀", "粘性维系与情感连接", "多样化互动活动\n增强期待、提升忠诚度", BLUE),
        ("04", "满意交付", "期望兑现的关键一跃", "交付细节管理\n产品质量与服务承诺一致", RED),
        ("05", "持续服务", "长期关系经营", "完善售后服务体系\n持续关注需求并提供支持", BLUE),
    ]
    y0 = 55
    ax.plot([8, 92], [y0, y0], color=NAVY, lw=3, zorder=2)
    for i, (no, t, sub, d, c) in enumerate(scenes):
        x = 10 + i * 20
        ax.add_patch(plt.Circle((x, y0), 2.6, color=c, zorder=4))
        txt(ax, x, y0, no, size=10, weight="bold")
        box(ax, x - 8.6, y0 + 8, 17.2, 26, fc=WHITE, ec=c, lw=1.6, r=1.0)
        txt(ax, x, y0 + 29.5, t, size=11.5, color=NAVY, weight="bold")
        txt(ax, x, y0 + 24.5, sub, size=8.2, color=c, weight="bold")
        txt(ax, x, y0 + 15.5, d, size=7.8, color="#4a5568", spacing=1.5)
        if i < 4:
            arrow(ax, x + 2.9, y0, x + 17.1, y0, color=GRAY, lw=2, ms=12)
    box(ax, 8, 8, 84, 14, fc=LIGHT, ec=GRAY, lw=1.0, r=1.0)
    txt(ax, 50, 17.6, "五大场景 × 客户典型行为、核心需求与痛点问题", size=10.5, color=NAVY, weight="bold")
    txt(ax, 50, 12.2, "梳理形成 25 个客户关注点，逐一转化为内部管理要点，实现“客户想要什么”到“我们该做什么”的精准转译",
        size=9, color="#4a5568")
    save(fig, "fig_scenarios.png")


# ==================================================== 图6 流程穿透矩阵
def fig_matrix():
    fig, ax = newfig(11.4, 5.4)
    txt(ax, 50, 96, "流程穿透：7 大阶段 × 5 条专业线的服务管控网络（104 个内部控制点）",
        size=13, color=NAVY, weight="bold")
    stages = ["立项", "定位", "首开备货", "开盘", "现场施工", "竣备交付", "持续服务"]
    lines = [("客关线", RED), ("营销线", BLUE), ("物业线", BLUE), ("签约线", BLUE), ("维修线（第三方）", BLUE)]
    x0, y0 = 14, 12
    cw, ch = 12.2, 12.6
    for j, s in enumerate(stages):
        box(ax, x0 + j * cw, y0 + 5 * ch + 2, cw - 1.2, 8, fc=NAVY, r=0.5)
        txt(ax, x0 + j * cw + (cw - 1.2) / 2, y0 + 5 * ch + 6, s, size=8.6, weight="bold")
    for i, (ln, c) in enumerate(lines):
        yy = y0 + (4 - i) * ch
        box(ax, 1.5, yy, 11.5, ch - 1.6, fc=c, r=0.5)
        txt(ax, 7.2, yy + (ch - 1.6) / 2, ln, size=8.6, weight="bold")
        for j in range(7):
            hot = (i, j) in [(0, 0), (0, 3), (0, 5), (0, 6), (1, 2), (1, 3), (2, 5), (2, 6), (3, 1), (3, 3), (4, 4), (4, 5)]
            fc = "#fde8e8" if hot else WHITE
            ec = RED if hot else "#cbd5e0"
            box(ax, x0 + j * cw, yy, cw - 1.2, ch - 1.6, fc=fc, ec=ec, lw=1.3 if hot else 0.8, r=0.5)
            if hot:
                ax.add_patch(plt.Circle((x0 + j * cw + (cw - 1.2) / 2, yy + (ch - 1.6) / 2), 2.4,
                                        color=RED, zorder=5))
    box(ax, 14, 1.2, 85, 7, fc=LIGHT, ec="none", r=0.6)
    txt(ax, 56.5, 4.7, "● 关键控制点示例：每个客户触点都有对应管理动作支撑，确保服务标准在项目一线有效落地",
        size=8.8, color="#2d3748", ha="center")
    save(fig, "fig_matrix.png")


# ==================================================== 图7 四维穿透转化链
def fig_penetration():
    fig, ax = newfig(11.4, 4.6)
    txt(ax, 50, 95, "四维穿透：从客户视角到企业执行的完整转译链", size=14, color=NAVY, weight="bold")
    steps = [
        ("客户视角", "客户想要什么", "5大场景\n25个关注点", GRAY, "外"),
        ("场景穿透", "需求转译为场景", "客户旅程\n行为路径与痛点", BLUE, "从外到内"),
        ("流程穿透", "场景落实为流程", "7阶段×5条线\n104个控制点", BLUE, "从粗到细"),
        ("岗位穿透", "流程分解到岗位", "第三方20个\n关键服务岗位", BLUE, "从虚到实"),
        ("标准穿透", "岗位固化为标准", "内容·成果·标准·奖罚\n“温度值”考核", RED, "可量化"),
    ]
    x = 4
    for i, (t, sub, d, c, tag) in enumerate(steps):
        box(ax, x, 30, 16.4, 40, fc=WHITE, ec=c, lw=1.8, r=1.2)
        box(ax, x, 58, 16.4, 12, fc=c, r=1.2)
        txt(ax, x + 8.2, 64, t, size=10.5, weight="bold")
        txt(ax, x + 8.2, 52, sub, size=8, color=NAVY, weight="bold")
        txt(ax, x + 8.2, 41, d, size=7.8, color="#4a5568", spacing=1.45)
        if i < 4:
            arrow(ax, x + 16.7, 50, x + 19.3, 50, color=NAVY, lw=2.4, ms=13)
        x += 19.6
    txt(ax, 50, 18, "视角转换方法论：从外到内，将客户需求转化为管理要求；从粗到细，将宏观理念转化为执行标准；从虚到实，将抽象概念转化为可量化指标",
        size=9, color="#4a5568")
    save(fig, "fig_penetration.png")


# ==================================================== 图8 服务经营价值
def fig_value():
    fig = plt.figure(figsize=(11.4, 5.0))
    ax = fig.add_axes([0.06, 0.14, 0.9, 0.74])
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")
    cards = [
        ("30%+", "老业主复购与推荐率", "良好客户服务可将老业主\n复购推荐率提升至30%以上\n显著降低营销成本、稳定现金流", RED),
        ("5%—10%", "住宅项目品牌溢价", "卓越物业服务带来\n住宅项目品牌溢价\n直接增厚项目收益", NAVY),
        ("8%—15%", "二手房价格优势", "优质物业管理的社区\n二手房价格普遍高出\n同地段物业8%—15%", BLUE),
        ("+20%", "专项服务满意度提升", "体系化动作实施一年内\n投诉、维修等核心痛点\n专项满意度提升20%", GOLD),
    ]
    for i, (num, t, d, c) in enumerate(cards):
        x = 3 + i * 24.5
        box(ax, x, 18, 22.5, 70, fc=LIGHT, ec=c, lw=1.6, r=1.2)
        txt(ax, x + 11.25, 74, num, size=22, color=c, weight="bold")
        txt(ax, x + 11.25, 62, t, size=10.5, color=NAVY, weight="bold")
        ax.plot([x + 4, x + 18.5], [56, 56], color=c, lw=1.2)
        txt(ax, x + 11.25, 40, d, size=8.6, color="#4a5568", spacing=1.6)
    ax.text(3, 96, "客户服务的经营账：服务力 = 现金流 + 溢价 + 资产保值", fontsize=14,
            color=NAVY, fontweight="bold", va="top")
    ax.text(3, 6, "资料来源：赛普研究院行业研究与项目实践数据", fontsize=8, color=GRAY)
    save(fig, "fig_value.png")


# ==================================================== 图9 体系建设五步法
def fig_steps():
    fig, ax = newfig(11.4, 4.6)
    txt(ax, 50, 95, "客户服务体系化建设五步法", size=14, color=NAVY, weight="bold")
    steps = [
        ("STEP 1", "体系规划", "聚焦问题深度诊断\n6E体检·VCM测评\n战略输入定方向"),
        ("STEP 2", "蓝图设计", "顶层设计一张蓝图\n一把手工程\n群策群力全维解构"),
        ("STEP 3", "体系深化", "流程·标准·机制\n分层分类细化\n工具化表单化"),
        ("STEP 4", "落地跟踪", "计划咬合·过程可视\n体系审核\n执行符合性审查"),
        ("STEP 5", "持续提升", "管理评审·定期复盘\n触发-诊断-建设-迭代\n体系动态进化"),
    ]
    x = 3
    for i, (no, t, d) in enumerate(steps):
        c = RED if i == 1 else (NAVY if i in (0, 4) else BLUE)
        box(ax, x, 34, 17.2, 42, fc=WHITE, ec=c, lw=1.8, r=1.2)
        box(ax, x, 66, 17.2, 10, fc=c, r=1.2)
        txt(ax, x + 8.6, 71, no, size=9.5, weight="bold")
        txt(ax, x + 8.6, 59, t, size=11.5, color=c, weight="bold")
        txt(ax, x + 8.6, 44, d, size=7.8, color="#4a5568", spacing=1.5)
        if i < 4:
            arrow(ax, x + 17.5, 55, x + 20.3, 55, color=GRAY, lw=2.2, ms=12)
        x += 19.6
    # 循环箭头
    arrow(ax, 88.6, 32, 12.6, 32, color=GRAY, lw=1.6, style="-|>", ms=11)
    txt(ax, 50, 26.5, "闭环迭代：落地跟踪发现的问题与外部环境变化持续反馈至体系建设，形成常态化升级机制",
        size=8.8, color="#4a5568")
    save(fig, "fig_steps.png")


# ==================================================== 图10 十五五路线图
def fig_roadmap():
    fig, ax = newfig(11.4, 5.4)
    txt(ax, 50, 96, "“十五五”客户服务体系建设路线图（2026—2030）", size=14, color=NAVY, weight="bold")
    phases = [
        ("2026—2027", "夯基垒台", "完成体系诊断与蓝图设计\n服务标准与流程全面建档\n考核·培训·计划·激励四抓手就位", BLUE),
        ("2028—2029", "立柱架梁", "四维穿透全项目覆盖\n第三方服务标准全面落地\n服务数据反哺产品迭代形成闭环", NAVY),
        ("2030", "积厚成势", "客户满意度进入行业优良区间\n复购推荐率成为稳定增长极\n服务品牌兑现资产溢价", RED),
    ]
    y0 = 42
    ax.plot([8, 92], [y0, y0], color=NAVY, lw=3, zorder=2)
    arrow(ax, 92, y0, 95, y0, color=NAVY, lw=3, ms=16)
    for i, (yr, t, d, c) in enumerate(phases):
        x = 16 + i * 30
        ax.add_patch(plt.Circle((x, y0), 1.1, color=c, zorder=4))
        box(ax, x - 13.5, y0 + 8, 27, 34, fc=WHITE, ec=c, lw=1.8, r=1.2)
        box(ax, x - 13.5, y0 + 34, 27, 8, fc=c, r=1.2)
        txt(ax, x, y0 + 38, yr, size=11, weight="bold")
        txt(ax, x, y0 + 29.5, t, size=11.5, color=c, weight="bold")
        txt(ax, x, y0 + 18.5, d, size=8.2, color="#4a5568", spacing=1.55)
    txt(ax, 50, 30, "与国家“十五五”规划节奏同频：2025年完成顶层设计，2026年开局起步，2030年形成体系化竞争力",
        size=9, color="#4a5568")
    box(ax, 8, 8, 84, 14, fc=LIGHT, ec=RED, lw=1.2, r=1.0)
    txt(ax, 50, 17.4, "闭环管理节点：年度体系审核 + 半年度管理评审 + 季度服务数据复盘", size=10, color=NAVY, weight="bold")
    txt(ax, 50, 11.8, "每年结合客户满意度第三方测评结果滚动修订服务标准，确保体系与“十五五”规划同周期进化",
        size=8.8, color="#4a5568")
    save(fig, "fig_roadmap.png")


if __name__ == "__main__":
    fig_timeline()
    fig_market()
    fig_14554()
    fig_goal()
    fig_scenarios()
    fig_matrix()
    fig_penetration()
    fig_value()
    fig_steps()
    fig_roadmap()
    print("all charts done")
