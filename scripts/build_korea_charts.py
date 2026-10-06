#!/usr/bin/env python3
"""生成《以韩为鉴》书稿配图，输出到 manuscript/figures/。

数据均来自公开统计（韩国国家数据处/统计厅、韩国银行、OECD、韩国人事革新处、
韩国国税厅、KB 不动产、中国国家统计局、国家公务员局等），出处见附录与 Excel 数据表。
"""
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

ROOT = Path(__file__).resolve().parent.parent
FIG_DIR = ROOT / "manuscript" / "figures"
FIG_DIR.mkdir(parents=True, exist_ok=True)

for path in (
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
):
    if Path(path).exists():
        font_manager.fontManager.addfont(path)
plt.rcParams["font.sans-serif"] = ["WenQuanYi Micro Hei", "Droid Sans Fallback", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["figure.dpi"] = 150

NAVY = "#1f3a5f"
RED = "#c0392b"
GOLD = "#d4a017"
GREY = "#7f8c8d"
TEAL = "#16a085"


def save(fig, name, source):
    fig.text(0.01, 0.01, f"数据来源：{source}", fontsize=7.5, color=GREY, ha="left")
    fig.tight_layout(rect=(0, 0.04, 1, 1))
    fig.savefig(FIG_DIR / name, bbox_inches="tight")
    plt.close(fig)
    print("已生成", name)


def fig_fertility():
    years = [1970, 1980, 1990, 2000, 2005, 2010, 2015, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025]
    tfr = [4.53, 2.82, 1.57, 1.48, 1.09, 1.23, 1.24, 1.05, 0.98, 0.92, 0.84, 0.81, 0.78, 0.72, 0.75, 0.80]
    births = [100.7, 86.3, 65.0, 64.0, 43.5, 47.0, 43.8, 35.8, 32.7, 30.3, 27.2, 26.1, 24.9, 23.0, 23.8, 25.4]
    fig, ax1 = plt.subplots(figsize=(9, 4.8))
    ax1.bar(years, births, width=1.6, color=NAVY, alpha=0.75, label="出生人口（万人）")
    ax1.set_ylabel("出生人口（万人）")
    ax2 = ax1.twinx()
    ax2.plot(years, tfr, color=RED, marker="o", lw=2, label="总和生育率")
    ax2.axhline(2.1, color=GREY, ls="--", lw=1)
    ax2.text(1971, 2.18, "人口更替水平 2.1", color=GREY, fontsize=8)
    ax2.set_ylabel("总和生育率")
    for x, y in [(2023, 0.72), (2025, 0.80)]:
        ax2.annotate(f"{y:.2f}", (x, y), textcoords="offset points", xytext=(0, 8), ha="center", color=RED, fontsize=9)
    ax1.set_title("图1  韩国出生人口与总和生育率（1970—2025）")
    h1, l1 = ax1.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax1.legend(h1 + h2, l1 + l2, loc="upper right", fontsize=9)
    save(fig, "fig01_korea_fertility.png", "韩国国家数据处（原统计厅）出生统计，2025年为2026年8月发布的确定值")


def fig_gdp_kondratieff():
    years = list(range(1990, 2026))
    growth = [9.9, 10.8, 6.2, 6.9, 9.3, 9.6, 7.9, 6.2, -4.9, 11.6, 9.1, 4.9, 7.7, 3.1, 5.2, 4.3, 5.3, 5.8,
              3.0, 0.8, 7.0, 3.7, 2.4, 3.2, 3.2, 2.8, 2.9, 3.2, 2.9, 2.3, -0.7, 4.3, 2.7, 1.4, 2.0, 1.0]
    fig, ax = plt.subplots(figsize=(10, 4.8))
    phases = [
        (1990, 2007.5, "#e8f1fb", "第五次康波·繁荣期\n（信息技术扩散）"),
        (2007.5, 2015.5, "#fdf2e3", "衰退期"),
        (2015.5, 2025.5, "#f6e6e6", "萧条期"),
    ]
    for start, end, color, label in phases:
        ax.axvspan(start, end, color=color, zorder=0)
        ax.text((start + end) / 2, 15.6, label, ha="center", va="top", fontsize=8.5, color=NAVY)
    colors = [RED if g < 0 else NAVY for g in growth]
    ax.bar(years, growth, color=colors, width=0.7, zorder=2)
    ax.axhline(0, color="black", lw=0.8)
    events = {1998: "亚洲金融危机", 2003: "信用卡危机", 2009: "全球金融危机", 2020: "新冠疫情", 2025: "建设投资萎缩"}
    for x, text in events.items():
        y = growth[years.index(x)]
        ax.annotate(text, (x, y), textcoords="offset points", xytext=(0, -14 if y < 0 else 6),
                    ha="center", fontsize=7.5, color=RED)
    ax.set_ylim(-7, 16)
    ax.set_ylabel("实际GDP增长率（%）")
    ax.set_title("图2  韩国实际GDP增长率与康波阶段（1990—2025）")
    save(fig, "fig02_korea_gdp_kondratieff.png",
         "韩国银行国民账户（部分年份经基期修订，可能有±0.3个百分点差异）；康波分期参照周金涛框架")


def fig_imf_unemployment():
    labels = ["1997.10", "1998.02", "1998.03", "1998.04", "1998.05", "1998.06", "1998.07", "1998.08", "1999.02"]
    rate = [2.1, 5.9, 6.5, 6.7, 6.9, 7.0, 7.6, 7.4, 8.7]
    fig, ax = plt.subplots(figsize=(9, 4.4))
    ax.plot(labels, rate, color=RED, marker="o", lw=2.2)
    ax.fill_between(range(len(labels)), rate, color=RED, alpha=0.12)
    for i, v in enumerate(rate):
        ax.text(i, v + 0.25, f"{v}%", ha="center", fontsize=8.5)
    ax.set_ylim(0, 10)
    ax.set_ylabel("失业率（%）")
    ax.set_title("图3  亚洲金融危机期间韩国失业率：16个月从2.1%升至8.7%")
    ax.text(0.2, 8.8, "1998年全年失业人口146.3万，\n比1997年增加90.7万；\n1998年6月仅约7%的失业者领到失业金", fontsize=8.5, color=NAVY)
    save(fig, "fig03_imf_unemployment.png", "韩国统计厅月度雇佣动向；KDI《雇佣创出研究》；KDI School 工作论文 w99-01")


def fig_civil_service():
    kr_years = [2011, 2016, 2021, 2022, 2023, 2024, 2025]
    kr_ratio = [93.3, 53.8, 35.0, 29.2, 22.8, 21.8, 24.3]
    cn_years = [2020, 2021, 2022, 2023, 2024, 2025, 2026]
    cn_ratio = [60, 61, 68, 70, 77, 86, 98]
    fig, ax = plt.subplots(figsize=(9, 4.6))
    ax.plot(kr_years, kr_ratio, color=NAVY, marker="o", lw=2, label="韩国国家职9级公务员公开招聘 平均竞争率")
    ax.plot(cn_years, cn_ratio, color=RED, marker="s", lw=2, label="中国国考 过审人数与录用计划之比")
    for x, y in zip(kr_years, kr_ratio):
        ax.text(x, y + 2.5, f"{y}", ha="center", fontsize=8, color=NAVY)
    for x, y in zip(cn_years, cn_ratio):
        ax.text(x, y + 2.5, f"{y}", ha="center", fontsize=8, color=RED)
    ax.set_ylabel("竞争比（X:1）")
    ax.set_ylim(0, 110)
    ax.legend(fontsize=9, loc="lower left")
    ax.set_title("图4  考公热的两条曲线：韩国退潮，中国升温")
    save(fig, "fig04_civil_service_ratio.png", "韩国人事革新处历年公告；国家公务员局及新京报、京报网报道（年份为考试年度）")


def fig_teacher():
    years = [2014, 2016, 2018, 2020, 2022, 2023, 2024, 2025, 2026]
    hires = [7386, 6591, 4089, 3916, 3758, 3561, 3157, 4272, 3113]
    pass_years = [2018, 2020, 2021, 2022, 2023, 2024]
    pass_rate = [68.9, 53.9, 50.8, 48.6, 47.7, 43.6]
    fig, ax1 = plt.subplots(figsize=(9, 4.6))
    ax1.bar(years, hires, color=TEAL, width=0.8, label="公立小学教师新录用人数（人）")
    for x, y in zip(years, hires):
        ax1.text(x, y + 120, str(y), ha="center", fontsize=8)
    ax1.set_ylabel("录用人数")
    ax2 = ax1.twinx()
    ax2.plot(pass_years, pass_rate, color=RED, marker="o", lw=2, label="小学教师录用考试合格率（%）")
    ax2.set_ylim(30, 80)
    ax2.set_ylabel("合格率（%）")
    h1, l1 = ax1.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax1.legend(h1 + h2, l1 + l2, fontsize=8.5, loc="upper right")
    ax1.set_title("图5  “教师铁饭碗”的收缩：韩国小学教师录用十年减少近六成")
    save(fig, "fig05_korea_teacher.png",
         "韩国教育部；Newspim 2024-04-11；东亚日报 2025-09-10（2025年因转岗“放学后照护室长”临时扩招）")


def fig_wage_gap():
    fig, ax = plt.subplots(figsize=(8, 4.4))
    cats = ["大企业", "非营利企业", "中小企业", "全体平均", "全体中位数"]
    v2024 = [613, 357, 307, 375, 288]
    bars = ax.bar(cats, v2024, color=[NAVY, GREY, RED, GOLD, GOLD])
    for b, v in zip(bars, v2024):
        ax.text(b.get_x() + b.get_width() / 2, v + 8, f"{v}万韩元", ha="center", fontsize=9)
    ax.set_ylabel("月平均收入（万韩元）")
    ax.set_ylim(0, 700)
    ax.set_title("图6  2024年韩国工资岗位月平均收入：大企业是中小企业的2倍")
    save(fig, "fig06_wage_gap.png", "韩国国家数据处《2024年工资劳动岗位收入（报酬）结果》，2026-02-23")


def fig_closures():
    years = [2019, 2020, 2021, 2022, 2023, 2024, 2025]
    closed = [92.2, 89.5, 88.5, 86.7, 98.6, 100.8, 97.6]
    fig, ax = plt.subplots(figsize=(8.5, 4.4))
    bars = ax.bar(years, closed, color=[GREY] * 4 + [RED, RED, NAVY])
    for b, v in zip(bars, closed):
        ax.text(b.get_x() + b.get_width() / 2, v + 0.8, f"{v}", ha="center", fontsize=9)
    ax.axhline(100, color=RED, ls="--", lw=1)
    ax.text(2018.6, 101.2, "100万", color=RED, fontsize=8)
    ax.set_ylim(70, 108)
    ax.set_ylabel("关店（停业）经营者数（万）")
    ax.set_title("图7  韩国年度停业经营者数量：2024年首次突破100万")
    save(fig, "fig07_closures.png", "韩国国税厅国税统计；韩国中小风险企业部（2025年为97.6万家，口径为经营者）")


def fig_elderly_poverty():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4.4), gridspec_kw={"width_ratios": [1.1, 1]})
    cats = ["韩国\n66岁+", "韩国\n76岁+", "日本\n66岁+", "OECD平均\n66岁+"]
    vals = [39.7, 54.0, 20.0, 14.8]
    bars = ax1.bar(cats, vals, color=[RED, RED, NAVY, GREY])
    for b, v in zip(bars, vals):
        ax1.text(b.get_x() + b.get_width() / 2, v + 1, f"{v}%", ha="center", fontsize=9)
    ax1.set_ylim(0, 62)
    ax1.set_title("老年相对收入贫困率（OECD 2025）", fontsize=10)
    years = [2015, 2017, 2019, 2021, 2023, 2025]
    trend = [49.6, 45.7, 43.8, 43.4, 40.4, 39.7]
    ax2.plot(years, trend, color=RED, marker="o", lw=2)
    for x, y in zip(years, trend):
        ax2.text(x, y + 0.6, f"{y}", ha="center", fontsize=8.5)
    ax2.set_ylim(35, 53)
    ax2.set_title("韩国老年贫困率：历版《Pensions at a Glance》", fontsize=10)
    fig.suptitle("图8  OECD 最贫困的老人：在改善，但仍是OECD平均的2.7倍", fontsize=11.5)
    save(fig, "fig08_elderly_poverty.png", "OECD Pensions at a Glance 2025；韩国国民年金研究院；横轴年份为报告版次")


def fig_pir():
    fig, ax = plt.subplots(figsize=(8.5, 4.4))
    cats = ["收入第1分位\n（最低20%）", "收入第3分位\n（中间20%）", "收入第5分位\n（最高20%）"]
    vals = [29.36, 10.49, 4.44]
    bars = ax.barh(cats, vals, color=[RED, GOLD, NAVY])
    for b, v in zip(bars, vals):
        ax.text(v + 0.4, b.get_y() + b.get_height() / 2, f"{v} 年", va="center", fontsize=10)
    ax.invert_yaxis()
    ax.set_xlim(0, 34)
    ax.set_xlabel("购买首尔中间价位住房所需年收入倍数（PIR）")
    ax.set_title("图9  首尔“不吃不喝”买房年限：低收入家庭需29年")
    save(fig, "fig09_seoul_pir.png", "KB不动产数据中心，2026年3月；首尔公寓均价7月为13.62亿韩元（韩国不动产院）")


def fig_debt():
    fig, ax = plt.subplots(figsize=(8.5, 4.4))
    labels = ["韩国\n2019Q3", "韩国\n2021Q3(峰值)", "韩国\n2024末", "韩国\n2025末", "中国\n2025末"]
    vals = [88.3, 99.1, 89.6, 88.6, 59.4]
    bars = ax.bar(labels, vals, color=[NAVY, RED, NAVY, NAVY, GOLD])
    for b, v in zip(bars, vals):
        ax.text(b.get_x() + b.get_width() / 2, v + 1.2, f"{v}%", ha="center", fontsize=9)
    ax.axhline(80, color=GREY, ls="--", lw=1)
    ax.text(4, 70, "韩国银行：超过80%—85%后\n债务对消费和增长的拖累加剧", fontsize=8, color=GREY, ha="center")
    ax.set_ylim(0, 110)
    ax.set_ylabel("居民（家庭）债务/GDP（%）")
    ax.set_title("图10  居民杠杆率：韩国近年杠杆率下降主要靠名义GDP膨胀")
    save(fig, "fig10_household_debt.png",
         "BIS（经首尔经济日报、KBS报道）；国家金融与发展实验室《2025年四季度宏观杠杆率报告》。韩国口径未含约165.7万亿韩元传贳贷款")


def fig_kondratieff_map():
    fig, ax = plt.subplots(figsize=(11, 4.6))
    waves = [
        (1948, 1966, "#dce9f5", "第四次康波 繁荣"),
        (1966, 1982, "#f3e3d3", "第四次康波 衰退/萧条"),
        (1982, 1991, "#e2f0e3", "第五次康波 回升"),
        (1991, 2007, "#dce9f5", "第五次康波 繁荣"),
        (2007, 2015, "#f3e3d3", "衰退"),
        (2015, 2025, "#f1d9d9", "萧条"),
        (2025, 2040, "#e2f0e3", "第六次康波 回升（推测）"),
    ]
    for s, e, c, t in waves:
        ax.axvspan(s, e, ymin=0.62, ymax=0.95, color=c)
        ax.text((s + e) / 2, 0.86, t, ha="center", va="center", fontsize=8.5, color=NAVY)
    korea = [
        (1962, "第一个五年计划\n出口导向"), (1973, "重化工业宣言"), (1988, "汉城奥运\n民主化"),
        (1997, "IMF危机"), (2004, "半导体/手机\n全盛"), (2008, "全球金融危机"),
        (2018, "生育率跌破1"), (2024, "超老龄社会\nAI存储超级周期"),
    ]
    china = [
        (1978, "改革开放"), (2001, "加入WTO"), (2008, "四万亿"), (2015, "股灾/去库存"),
        (2021, "房地产拐点"), (2025, "人口连降四年\nAI与新能源"),
    ]
    for i, (x, t) in enumerate(korea):
        y = 0.52 if i % 2 == 0 else 0.38
        ax.plot([x, x], [y, 0.6], color=RED, lw=0.8)
        ax.text(x, y, t, ha="center", va="center", fontsize=7.5, color=RED,
                bbox=dict(boxstyle="round,pad=0.2", fc="white", ec=RED, lw=0.6))
    for i, (x, t) in enumerate(china):
        y = 0.2 if i % 2 == 0 else 0.07
        ax.text(x, y, t, ha="center", va="center", fontsize=7.5, color=GOLD,
                bbox=dict(boxstyle="round,pad=0.2", fc="white", ec=GOLD, lw=0.6))
    ax.text(1945, 0.45, "韩国", fontsize=10, color=RED, ha="right", va="center")
    ax.text(1945, 0.13, "中国", fontsize=10, color=GOLD, ha="right", va="center")
    ax.set_xlim(1938, 2041)
    ax.set_ylim(0, 1)
    ax.set_yticks([])
    ax.set_title("图11  康波坐标中的韩国与中国（示意）")
    save(fig, "fig11_kondratieff_map.png", "康波分期参照周金涛等研究；事件为作者整理。康波是分析框架，不是确定性规律")


def main():
    fig_fertility()
    fig_gdp_kondratieff()
    fig_imf_unemployment()
    fig_civil_service()
    fig_teacher()
    fig_wage_gap()
    fig_closures()
    fig_elderly_poverty()
    fig_pir()
    fig_debt()
    fig_kondratieff_map()


if __name__ == "__main__":
    main()
