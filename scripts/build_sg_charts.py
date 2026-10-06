"""生成《以新为鉴》书稿所需的数据图表。

所有数据均来自书稿附录 C 列出的公开来源（新加坡统计局、人力部、建屋局、
公积金局、财政部、国家统计局等），检索日期为 2026 年 10 月。
"""

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "book" / "assets" / "charts"
OUT.mkdir(parents=True, exist_ok=True)

for path in (
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
):
    if Path(path).exists():
        font_manager.fontManager.addfont(path)
plt.rcParams["font.sans-serif"] = ["WenQuanYi Micro Hei", "Droid Sans Fallback"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["figure.dpi"] = 150

RED = "#C8102E"
NAVY = "#1F3A5F"
GREY = "#8A8F98"
GOLD = "#C9A227"
TEAL = "#2A9D8F"


def _finish(fig, name, source):
    fig.text(0.01, 0.01, f"数据来源：{source}", fontsize=7.5, color=GREY, ha="left")
    fig.tight_layout(rect=(0, 0.035, 1, 1))
    fig.savefig(OUT / name, bbox_inches="tight")
    plt.close(fig)


def chart_gdp_kondratiev():
    growth = {
        1976: 7.40, 1977: 7.46, 1978: 8.68, 1979: 9.40, 1980: 10.05, 1981: 10.73,
        1982: 7.19, 1983: 8.57, 1984: 8.82, 1985: -0.65, 1986: 1.29, 1987: 10.77,
        1988: 11.07, 1989: 10.23, 1990: 10.04, 1991: 6.69, 1992: 7.09, 1993: 11.54,
        1994: 11.10, 1995: 7.20, 1996: 7.47, 1997: 8.32, 1998: -2.19, 1999: 5.72,
        2000: 9.04, 2001: -1.07, 2002: 3.92, 2003: 4.55, 2004: 9.94, 2005: 7.37,
        2006: 9.01, 2007: 9.02, 2008: 1.86, 2009: 0.13, 2010: 14.52, 2011: 6.21,
        2012: 4.44, 2013: 4.82, 2014: 3.94, 2015: 2.98, 2016: 3.91, 2017: 4.33,
        2018: 3.19, 2019: 1.16, 2020: -3.61, 2021: 10.14, 2022: 4.01, 2023: 1.45,
        2024: 5.34, 2025: 5.03,
    }
    years = list(growth)
    vals = [growth[y] for y in years]
    fig, ax = plt.subplots(figsize=(10, 4.8))
    phases = [
        (1975.5, 1985.5, "#F3E9D2", "第四波康波尾声\n（石化·电子装配）"),
        (1985.5, 2000.5, "#E3EEF7", "第五波回升—繁荣\n（ICT·晶圆·硬盘）"),
        (2000.5, 2008.5, "#EAF4EA", "第五波繁荣后段\n（金融·生物医药）"),
        (2008.5, 2024.5, "#F2F2F2", "第五波衰退—萧条\n（低增长·存量博弈）"),
        (2024.5, 2025.7, "#FBE3E6", "六"),
    ]
    for x0, x1, color, label in phases:
        ax.axvspan(x0, x1, color=color, zorder=0)
        ax.text((x0 + x1) / 2, 15.3, label, ha="center", va="top", fontsize=7.5, color=NAVY)
    colors = [RED if v < 0.5 else NAVY for v in vals]
    ax.bar(years, vals, color=colors, width=0.75, zorder=2)
    for y, (note, ty, tx) in {1985: ("1985 首次衰退", -6.2, 1985), 1998: ("1998 亚洲金融危机", -6.2, 1996),
                              2001: ("2001 科网泡沫", -4.6, 2003), 2009: ("2009 全球金融危机", -6.2, 2009),
                              2020: ("2020 新冠", -6.2, 2020)}.items():
        ax.annotate(note, (y, growth[y]), xytext=(tx, ty), ha="center", fontsize=7.5,
                    color=RED, arrowprops=dict(arrowstyle="-", color=RED, lw=0.6))
    ax.axhline(0, color="black", lw=0.6)
    ax.set_ylim(-7.5, 16)
    ax.set_xlim(1975.3, 2025.7)
    ax.set_ylabel("实际 GDP 增长率（%）")
    ax.set_title("图1  新加坡五十年增长曲线与康波阶段（1976—2025）", fontsize=12, color=NAVY)
    ax.spines[["top", "right"]].set_visible(False)
    _finish(fig, "fig01_gdp_kondratiev.png",
            "IMF WEO 年度实际 GDP 增长（ycharts 整理）；康波阶段划分为作者整理，仅作分析框架")


def chart_tfr():
    tfr_desc = [0.87, 0.97, 0.97, 1.04, 1.12, 1.10, 1.14, 1.14, 1.16, 1.20, 1.24, 1.25, 1.19,
                1.29, 1.20, 1.15, 1.22, 1.28, 1.29, 1.28, 1.26, 1.26, 1.27, 1.37, 1.41, 1.60,
                1.47, 1.48, 1.61, 1.66, 1.67, 1.71, 1.74, 1.72, 1.73, 1.83, 1.75, 1.96, 1.62,
                1.43, 1.61, 1.62, 1.61, 1.74, 1.78, 1.82, 1.79, 1.79, 1.82, 2.11, 2.07, 2.35,
                2.79, 3.04, 3.02, 3.07, 3.22, 3.53, 3.91, 4.46, 4.66, 4.97, 5.16, 5.21, 5.41, 5.76]
    years = list(range(2025, 1959, -1))
    fig, ax = plt.subplots(figsize=(10, 4.6))
    ax.plot(years, tfr_desc, color=NAVY, lw=2)
    ax.fill_between(years, tfr_desc, color=NAVY, alpha=0.08)
    ax.axhline(2.1, color=GREY, ls="--", lw=1)
    ax.text(1961, 2.18, "更替水平 2.1", fontsize=8, color=GREY)
    notes = [
        (1972, 3.04, "1972 “两个就够”\n节育运动", 4.3),
        (1977, 1.82, "1977 跌破更替水平", 2.9),
        (1987, 1.62, "1987 “有能力就生三个”", 0.75),
        (2001, 1.41, "2001 婴儿花红", 2.35),
        (2023, 0.97, "2023 首次跌破 1", 1.75),
        (2025, 0.87, "2025 年 0.87", 0.45),
    ]
    for x, y, txt, ty in notes:
        ax.annotate(txt, (x, y), xytext=(x, ty), fontsize=7.8, ha="center", color=RED,
                    arrowprops=dict(arrowstyle="->", color=RED, lw=0.7))
    ax.set_ylim(0, 6.2)
    ax.set_xlim(1959, 2026.5)
    ax.set_ylabel("居民总和生育率")
    ax.set_title("图2  新加坡居民总和生育率六十五年（1960—2025）", fontsize=12, color=NAVY)
    ax.spines[["top", "right"]].set_visible(False)
    _finish(fig, "fig02_tfr.png", "新加坡统计局 Births and Fertility Rates（data.gov.sg）；2025 年为副总理颜金勇 2026 年 2 月公布值")


def chart_hdb_rpi():
    data = {
        2011: [126.4, 130.4, 135.4, 137.7], 2012: [138.5, 140.3, 143.1, 146.7],
        2013: [148.6, 149.4, 148.1, 145.8], 2014: [143.5, 141.5, 139.1, 137.0],
        2015: [135.6, 135.0, 134.6, 134.8], 2016: [134.7, 134.7, 134.7, 134.6],
        2017: [133.9, 133.7, 132.8, 132.6], 2018: [131.6, 131.7, 131.6, 131.4],
        2019: [131.0, 130.8, 130.9, 131.5], 2020: [131.5, 131.9, 133.9, 138.1],
        2021: [142.2, 146.4, 150.6, 155.7], 2022: [159.5, 163.9, 168.1, 171.9],
        2023: [173.6, 176.2, 178.5, 180.4], 2024: [183.7, 187.9, 192.9, 197.9],
        2025: [201.0, 202.9, 203.7, 203.6], 2026: [203.4, 202.8],
    }
    xs, ys = [], []
    for y, qs in data.items():
        for i, v in enumerate(qs):
            xs.append(y + i * 0.25)
            ys.append(v)
    fig, ax = plt.subplots(figsize=(10, 4.6))
    ax.plot(xs, ys, color=NAVY, lw=2)
    ax.axvspan(2013.25, 2019.5, color="#F2F2F2")
    ax.axvspan(2020.25, 2025.5, color="#FBE3E6")
    ax.text(2016.4, 192, "2013—2019 连续降温\n累计约 -12.5%", ha="center", fontsize=8.5, color=NAVY)
    ax.text(2022.9, 140, "2020—2025 疫后上行\n累计约 +55%", ha="center", fontsize=8.5, color=RED)
    ax.annotate("2026Q2 202.8\n连续两季回落", (2026.25, 202.8), xytext=(2024.3, 212),
                fontsize=8, color=RED, arrowprops=dict(arrowstyle="->", color=RED, lw=0.7))
    ax.set_ylim(115, 220)
    ax.set_ylabel("组屋转售价格指数（2009Q1=100）")
    ax.set_title("图3  组屋转售价格指数：一轮完整的“降温—反弹—再降温”", fontsize=12, color=NAVY)
    ax.spines[["top", "right"]].set_visible(False)
    _finish(fig, "fig03_hdb_rpi.png", "新加坡建屋发展局 Resale Price Index 1Q1990—2Q2026")


def chart_housing_compare():
    fig, axes = plt.subplots(1, 2, figsize=(10, 4.4))
    labels = ["组屋 1-2 房", "组屋 3 房", "组屋 4 房", "组屋 5 房及以上", "共管公寓等", "有地住宅"]
    vals = [7.3, 16.6, 31.2, 22.1, 17.9, 4.7]
    colors = ["#9DB4CF", "#6E92BA", NAVY, "#14273F", GOLD, GREY]
    axes[0].pie(vals, labels=labels, colors=colors, autopct="%1.1f%%", startangle=90,
                textprops={"fontsize": 8}, pctdistance=0.75,
                wedgeprops=dict(width=0.45, edgecolor="white"))
    axes[0].set_title("新加坡居民家庭住宅类型（2025）\n组屋合计 77.2%，自有率 91.2%", fontsize=10, color=NAVY)
    years = ["2021", "2025"]
    sales = [181930, 83937]
    bars = axes[1].bar(years, sales, color=[GREY, RED], width=0.5)
    for b, v in zip(bars, sales):
        axes[1].text(b.get_x() + b.get_width() / 2, v + 3000, f"{v/10000:.2f} 万亿元",
                     ha="center", fontsize=9)
    axes[1].set_ylim(0, 210000)
    axes[1].set_yticks([])
    axes[1].set_title("中国新建商品房销售额\n2025 年较 2021 年峰值下降约 54%", fontsize=10, color=NAVY)
    axes[1].spines[["top", "right", "left"]].set_visible(False)
    fig.suptitle("图4  两种住房体系：以保障为主的存量体系 vs 以增量开发为主的市场体系",
                 fontsize=12, color=NAVY)
    _finish(fig, "fig04_housing_compare.png", "新加坡统计局 Resident Households 2025；国家统计局 2021、2025 年全国房地产市场基本情况")


def chart_labour():
    years = list(range(2016, 2026))
    six = [64.6, 64.6, 62.9, 64.4, 61.6, 65.8, 68.9, 63.7, 58.4, 57.3]
    twelve = [77.9, 75.5, 75.4, 74.5, 74.1, 76.1, 77.1, 75.3, 72.8, 72.1]
    fig, axes = plt.subplots(1, 2, figsize=(10, 4.4), gridspec_kw={"width_ratios": [1.4, 1]})
    axes[0].plot(years, six, marker="o", color=RED, label="裁员后 6 个月再就业率")
    axes[0].plot(years, twelve, marker="s", color=NAVY, label="裁员后 12 个月再就业率")
    for x, y in zip(years, six):
        axes[0].text(x, y - 2.2, f"{y}", fontsize=7, ha="center", color=RED)
    axes[0].set_ylim(50, 82)
    axes[0].set_ylabel("%")
    axes[0].legend(fontsize=8, loc="upper right")
    axes[0].set_title("新加坡被裁员居民的再就业率", fontsize=10, color=NAVY)
    axes[0].spines[["top", "right"]].set_visible(False)

    cats = ["新加坡\n居民失业率\n(2025)", "中国 30-59 岁\n失业率\n(2026.8)", "中国\n城镇调查\n失业率\n(2026.8)",
            "中国 16-24 岁\n不含在校生\n失业率\n(2026.8)"]
    vals = [2.8, 3.9, 5.3, 18.9]
    colors = [NAVY, TEAL, GREY, RED]
    bars = axes[1].bar(range(4), vals, color=colors, width=0.6)
    for b, v in zip(bars, vals):
        axes[1].text(b.get_x() + b.get_width() / 2, v + 0.4, f"{v}%", ha="center", fontsize=9)
    axes[1].set_xticks(range(4))
    axes[1].set_xticklabels(cats, fontsize=7.2)
    axes[1].set_yticks([])
    axes[1].set_ylim(0, 22)
    axes[1].set_title("就业关键指标（口径不同，仅作量级参考）", fontsize=10, color=NAVY)
    axes[1].spines[["top", "right", "left"]].set_visible(False)
    fig.suptitle("图5  就业：低失业率之下的“再就业速度”才是真正的安全感", fontsize=12, color=NAVY)
    _finish(fig, "fig05_labour.png", "新加坡人力部 Re-entry into Employment、Labour Force 2025；新加坡教育部 GES 2025；国家统计局 2026 年 8 月数据")


def chart_cpf():
    ages = ["55岁及以下", "55—60岁", "60—65岁", "65—70岁", "70岁以上"]
    employer = [17, 16, 12.5, 9, 7.5]
    employee = [20, 18, 12.5, 7.5, 5]
    fig, ax = plt.subplots(figsize=(10, 4.4))
    x = range(len(ages))
    ax.bar(x, employer, color=NAVY, label="雇主缴纳")
    ax.bar(x, employee, bottom=employer, color=GOLD, label="雇员缴纳")
    for i in x:
        ax.text(i, employer[i] + employee[i] + 0.8, f"合计 {employer[i] + employee[i]:g}%",
                ha="center", fontsize=9)
    ax.set_xticks(list(x))
    ax.set_xticklabels(ages)
    ax.set_ylim(0, 44)
    ax.set_ylabel("占月工资比例（%）")
    ax.legend(fontsize=9)
    ax.set_title("图6  新加坡中央公积金缴费率（2026 年 1 月起，月薪上限 8,000 新元）", fontsize=12, color=NAVY)
    ax.spines[["top", "right"]].set_visible(False)
    _finish(fig, "fig06_cpf.png", "新加坡中央公积金局 CPF-related changes taking place in 2026；55—60 岁雇主/雇员分拆按公积金局费率表")


def chart_ageing():
    fig, axes = plt.subplots(1, 2, figsize=(10, 4.4))
    years = ["2015", "2024", "2025", "2030（预测）"]
    vals = [13.1, 19.9, 20.7, 23.9]
    bars = axes[0].bar(years, vals, color=[GREY, NAVY, NAVY, RED], width=0.55)
    for b, v in zip(bars, vals):
        axes[0].text(b.get_x() + b.get_width() / 2, v + 0.5, f"{v}%", ha="center", fontsize=9)
    axes[0].set_ylim(0, 28)
    axes[0].set_yticks([])
    axes[0].set_title("新加坡公民 65 岁及以上占比", fontsize=10, color=NAVY)
    axes[0].spines[["top", "right", "left"]].set_visible(False)

    cats = ["60 岁及以上", "65 岁及以上"]
    vals2 = [23.0, 15.9]
    bars = axes[1].bar(cats, vals2, color=[GOLD, RED], width=0.5)
    for b, v in zip(bars, vals2):
        axes[1].text(b.get_x() + b.get_width() / 2, v + 0.5, f"{v}%", ha="center", fontsize=9)
    axes[1].text(0.5, 26.5, "60 岁及以上 3.23 亿人", ha="center", fontsize=9, color=NAVY)
    axes[1].set_ylim(0, 28)
    axes[1].set_yticks([])
    axes[1].set_title("中国老年人口占比（2025 年末）", fontsize=10, color=NAVY)
    axes[1].spines[["top", "right", "left"]].set_visible(False)
    fig.suptitle("图7  两个正在变老的社会：新加坡已进入超老龄，中国紧随其后", fontsize=12, color=NAVY)
    _finish(fig, "fig07_ageing.png", "新加坡 Population in Brief 2025；国家统计局 2025 年统计公报")


def chart_inequality():
    fig, axes = plt.subplots(1, 2, figsize=(10, 4.4))
    labels = ["市场收入\n基尼系数", "转移支付与\n税收后"]
    vals = [0.452, 0.379]
    bars = axes[0].bar(labels, vals, color=[GREY, TEAL], width=0.5)
    for b, v in zip(bars, vals):
        axes[0].text(b.get_x() + b.get_width() / 2, v + 0.01, f"{v:.3f}", ha="center", fontsize=10)
    axes[0].set_ylim(0, 0.55)
    axes[0].set_yticks([])
    axes[0].set_title("新加坡 2025：再分配使基尼系数下降 0.073", fontsize=10, color=NAVY)
    axes[0].spines[["top", "right", "left"]].set_visible(False)

    yrs = ["2015", "2020", "2025"]
    ratio = [0.51, 0.52, 0.55]
    axes[1].plot(yrs, ratio, marker="o", color=RED, lw=2)
    for x, y in zip(yrs, ratio):
        axes[1].text(x, y + 0.004, f"{y:.2f}", ha="center", fontsize=10)
    axes[1].set_ylim(0.48, 0.58)
    axes[1].set_title("P20/P50 收入比：低薪者追赶中位数", fontsize=10, color=NAVY)
    axes[1].spines[["top", "right"]].set_visible(False)
    fig.suptitle("图8  “先做大蛋糕，再精准分蛋糕”：新加坡的分配账本", fontsize=12, color=NAVY)
    _finish(fig, "fig08_inequality.png", "新加坡统计局 Key Household Income Trends 2025；人力部 Labour Force in Singapore 2025")


def chart_fiscal():
    fig, ax = plt.subplots(figsize=(10, 4.4))
    items = ["FY2025 总支出", "FY2026 总支出（预算）", "FY2026 储备金投资收益贡献 NIRC"]
    vals = [124.46, 137.32, 28.48]
    bars = ax.barh(items, vals, color=[GREY, NAVY, GOLD], height=0.5)
    for b, v in zip(bars, vals):
        ax.text(v + 1.5, b.get_y() + b.get_height() / 2, f"{v * 10:,.1f} 亿新元", va="center", fontsize=9)
    ax.set_xlim(0, 165)
    ax.set_xlabel("十亿新元")
    ax.invert_yaxis()
    ax.set_title("图9  储备金收益：新加坡财政的“第三根支柱”（FY2026 预算）", fontsize=12, color=NAVY)
    ax.spines[["top", "right"]].set_visible(False)
    _finish(fig, "fig09_fiscal.png", "新加坡财政部 Budget 2026 Revenue and Expenditure Estimates")


def chart_kondratiev_map():
    fig, ax = plt.subplots(figsize=(10, 4.8))
    import numpy as np

    knots = [(1940, -0.5), (1966, 0.9), (1983, -0.9), (2005, 0.9), (2025, -0.9), (2050, 0.9),
             (2062, 0.3)]

    def wave(xv):
        for (x0, y0), (x1, y1) in zip(knots, knots[1:]):
            if x0 <= xv <= x1:
                t = (xv - x0) / (x1 - x0)
                return y0 + (y1 - y0) * (1 - np.cos(np.pi * t)) / 2
        return knots[-1][1]

    x = np.linspace(1940, 2060, 600)
    y = [wave(v) for v in x]
    ax.plot(x, y, color=NAVY, lw=2)
    for x0, x1, label in [(1940, 1983, "第四波（汽车·石化·电子）"), (1983, 2025, "第五波（信息技术）"),
                          (2025, 2060, "第六波（AI·新能源·生物科技）")]:
        ax.text((x0 + x1) / 2, 1.42, label, ha="center", fontsize=9, color=NAVY)
    ax.axvline(2026, color=RED, ls="--", lw=1)
    ax.text(2027, -1.45, "2026 年：萧条尾声 / 回升起点", color=RED, fontsize=8)
    marks = [
        (1965, "1965 独立\n裕廊工业区·炼油", -0.75),
        (1985, "1985 首次衰退\n公积金降费", 0.5),
        (1997, "1990s 硬盘·晶圆\n“东方硅谷”", -0.55),
        (2008, "2008 金融危机\n第五波见顶回落", -0.65),
        (2018, "2015—2025 萧条期\n存量博弈·AI 酝酿", -0.75),
        (2040, "第六波回升\nAI 资本开支·能源转型", -0.6),
    ]
    for xm, txt, dy in marks:
        ym = wave(xm)
        ax.scatter([xm], [ym], color=GOLD, zorder=3)
        ax.annotate(txt, (xm, ym), xytext=(xm, ym + dy),
                    fontsize=7.8, ha="center", color=NAVY,
                    arrowprops=dict(arrowstyle="-", color=GREY, lw=0.6))
    ax.set_yticks([])
    ax.set_xlim(1940, 2060)
    ax.set_ylim(-1.6, 1.6)
    ax.set_title("图10  康波示意：新加坡的四次产业押注与第六波（示意图，非精确测算）",
                 fontsize=12, color=NAVY)
    ax.spines[["top", "right", "left"]].set_visible(False)
    _finish(fig, "fig10_kondratiev_map.png", "作者根据康德拉季耶夫长波理论与新加坡产业史整理，曲线为示意")


if __name__ == "__main__":
    chart_gdp_kondratiev()
    chart_tfr()
    chart_hdb_rpi()
    chart_housing_compare()
    chart_labour()
    chart_cpf()
    chart_ageing()
    chart_inequality()
    chart_fiscal()
    chart_kondratiev_map()
    print("charts ->", OUT)
