"""基于聚合结果生成图表（PNG，中文字体），供报告与 PPT 引用。"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
import pandas as pd

import config

# 注册中文字体
font_manager.fontManager.addfont(config.CJK_FONT)
CJK = font_manager.FontProperties(fname=config.CJK_FONT).get_name()
plt.rcParams["font.sans-serif"] = [CJK]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["font.size"] = 11

BLUE = "#1f4e79"
ORANGE = "#e07b39"
GREY = "#8a8a8a"
PALETTE = ["#1f4e79", "#2e75b6", "#5b9bd5", "#9dc3e6", "#e07b39", "#a5a5a5"]
D = config.DATA_DIR
C = config.CHART_DIR


def _save(fig, name):
    fig.tight_layout()
    fig.savefig(C / name, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print("saved", name)


def chart_supply():
    df = pd.read_csv(D / "agg_supply_by_subdistrict.csv")
    df = df.sort_values("可租面积(万㎡)")
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.barh(df["subdistrict"], df["可租面积(万㎡)"], color=BLUE)
    ax.set_xlabel("可租赁面积（万㎡）")
    ax.set_title("图1 · 杨浦区分街道可租赁面积分布", fontweight="bold")
    _save(fig, "01_supply_by_subdistrict.png")


def chart_grade_pie():
    df = pd.read_csv(D / "assets.csv")
    cnt = df.groupby("grade")["rentable_area_wan"].sum()
    fig, ax = plt.subplots(figsize=(6, 6))
    ax.pie(cnt.values, labels=cnt.index, autopct="%1.0f%%",
           colors=PALETTE, startangle=90, wedgeprops={"edgecolor": "white"})
    ax.set_title("图2 · 物业等级面积结构（办公空间金字塔）", fontweight="bold")
    _save(fig, "02_grade_structure.png")


def chart_rent_vacancy():
    df = pd.read_csv(D / "agg_market_by_plate.csv")
    fig, ax1 = plt.subplots(figsize=(9, 5))
    x = range(len(df))
    ax1.bar(x, df["成交租金"], color=BLUE, label="成交租金(元/㎡/天)")
    ax1.set_ylabel("成交租金（元/㎡/天）", color=BLUE)
    ax1.set_xticks(list(x))
    ax1.set_xticklabels(df["plate"], rotation=20, ha="right")
    ax2 = ax1.twinx()
    ax2.plot(x, df["空置率%"], color=ORANGE, marker="o", linewidth=2, label="空置率%")
    ax2.set_ylabel("空置率（%）", color=ORANGE)
    ax1.set_title("图3 · 分板块成交租金与空置率分化", fontweight="bold")
    _save(fig, "03_rent_vacancy_by_plate.png")


def chart_industry():
    df = pd.read_csv(D / "agg_industry_total.csv").sort_values("企业数")
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.barh(df["industry_tag"], df["企业数"], color=PALETTE)
    for i, (v, a) in enumerate(zip(df["企业数"], df["平均面积(㎡)"])):
        ax.text(v + 1, i, f"均{int(a)}㎡", va="center", fontsize=9, color=GREY)
    ax.set_xlabel("入驻企业数（样本）")
    ax.set_title("图4 · 主导产业企业数与平均办公面积", fontweight="bold")
    _save(fig, "04_industry_mix.png")


def chart_inout():
    df = pd.read_csv(D / "agg_inout_ratio.csv").sort_values("InOut比")
    colors = [ORANGE if v < 1 else BLUE for v in df["InOut比"]]
    fig, ax = plt.subplots(figsize=(8, 4.5))
    ax.barh(df["板块"], df["InOut比"], color=colors)
    ax.axvline(1.0, color=GREY, linestyle="--", linewidth=1)
    ax.set_xlabel("企业流入/流出比（>1 净吸引）")
    ax.set_title("图5 · 杨浦内部板块招商吸引力 In/Out Ratio", fontweight="bold")
    _save(fig, "05_inout_ratio.png")


def chart_migration():
    df = pd.read_csv(D / "migrations.csv")
    to_yp = df[df.to_area.str.startswith("杨浦")]
    top = to_yp["from_area"].value_counts().head(6).sort_values()
    fig, ax = plt.subplots(figsize=(8, 4.5))
    ax.barh(top.index, top.values, color=BLUE)
    ax.set_xlabel("迁入杨浦企业数（样本）")
    ax.set_title("图6 · 迁入杨浦的主要来源地", fontweight="bold")
    _save(fig, "06_migration_source.png")


if __name__ == "__main__":
    chart_supply()
    chart_grade_pie()
    chart_rent_vacancy()
    chart_industry()
    chart_inout()
    chart_migration()
    print("All charts done ->", C)
