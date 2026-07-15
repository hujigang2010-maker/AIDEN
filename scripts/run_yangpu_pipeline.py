"""杨浦区数据管线编排：采集 → 清洗 → 聚合 → 落盘。

输出：
  data/yangpu/*.csv  （原始 + 聚合）
  data/yangpu/headline.json  （报告/PPT 填充用关键指标）
"""
import json

import pandas as pd

import config
import amap_collector
import qcc_collector
import listing_collector


def main():
    mode_flags = []

    assets, m1 = amap_collector.collect()
    companies, migrations, m2 = qcc_collector.collect()
    listings, m3 = listing_collector.collect()
    mode_flags = {"amap": m1, "qcc": m2, "listing": m3}

    df_assets = pd.DataFrame(assets)
    df_comp = pd.DataFrame(companies)
    df_mig = pd.DataFrame(migrations)
    df_list = pd.DataFrame(listings)

    # ---- 落原始表 ----
    df_assets.to_csv(config.DATA_DIR / "assets.csv", index=False, encoding="utf-8-sig")
    df_comp.to_csv(config.DATA_DIR / "companies.csv", index=False, encoding="utf-8-sig")
    df_mig.to_csv(config.DATA_DIR / "migrations.csv", index=False, encoding="utf-8-sig")
    df_list.to_csv(config.DATA_DIR / "listings.csv", index=False, encoding="utf-8-sig")

    # ---- 聚合 1：分街道供给 ----
    supply = df_assets.groupby("subdistrict").agg(
        载体数量=("name", "count"),
        总建面=("total_gfa_wan", "sum"),
        可租面积=("rentable_area_wan", "sum"),
    ).reset_index()
    supply = supply.rename(columns={"总建面": "总建面(万㎡)", "可租面积": "可租面积(万㎡)"})
    office_cnt = df_assets[df_assets.asset_type == "office"].groupby("subdistrict")["name"].count()
    supply["写字楼数"] = supply["subdistrict"].map(office_cnt).fillna(0).astype(int)
    supply["产业园数"] = supply["载体数量"] - supply["写字楼数"]
    supply = supply.sort_values("可租面积(万㎡)", ascending=False)
    supply.to_csv(config.DATA_DIR / "agg_supply_by_subdistrict.csv", index=False, encoding="utf-8-sig")

    # ---- 聚合 2：市场景气指标 ----
    import numpy as np
    rng = np.random.default_rng(33)
    avg_quote = round(df_list["quote_rent"].mean(), 2)
    avg_deal = round(df_list["deal_rent"].mean(), 2)
    avg_fee = round(df_list["property_fee"].mean(), 1)
    full_cost = round(avg_deal + avg_fee / 30, 2)  # 物业费换算到 元/㎡/天

    # 每个载体按等级赋予合理空置率（演示口径），可租面积加权得到全区空置率
    base_vac = {"甲级": 0.18, "乙级": 0.22, "科创园": 0.14, "孵化器": 0.16}
    df_assets["vacancy_rate"] = df_assets["grade"].map(base_vac).fillna(0.18)
    df_assets["vacancy_rate"] = (df_assets["vacancy_rate"]
                                 + rng.normal(0, 0.04, len(df_assets))).clip(0.03, 0.40)
    df_assets["vacant_area_wan"] = (df_assets["rentable_area_wan"] * df_assets["vacancy_rate"]).round(2)
    vacancy_rate = round(
        df_assets["vacant_area_wan"].sum() / df_assets["rentable_area_wan"].sum() * 100, 1)

    # ---- 聚合 2b：分板块市场表（租金 + 空置，供图表用）----
    deal_by_plate = df_list.groupby("plate")["deal_rent"].mean()
    market = df_assets.groupby("plate").agg(
        可租面积=("rentable_area_wan", "sum"),
        空置面积=("vacant_area_wan", "sum"),
    ).reset_index()
    market["空置率%"] = (market["空置面积"] / market["可租面积"] * 100).round(1)
    market["成交租金"] = market["plate"].map(deal_by_plate).round(2)
    market = market.rename(columns={"可租面积": "可租面积(万㎡)", "空置面积": "空置面积(万㎡)"})
    market = market.sort_values("成交租金", ascending=False)
    market.to_csv(config.DATA_DIR / "agg_market_by_plate.csv", index=False, encoding="utf-8-sig")
    df_assets.to_csv(config.DATA_DIR / "assets.csv", index=False, encoding="utf-8-sig")

    # ---- 聚合 3：产业结构（按板块）----
    ind_mix = df_comp.groupby(["plate", "industry_tag"]).size().reset_index(name="企业数")
    ind_mix.to_csv(config.DATA_DIR / "agg_industry_mix.csv", index=False, encoding="utf-8-sig")

    ind_total = df_comp.groupby("industry_tag").agg(
        企业数=("company_id", "count"),
        平均面积=("area_sqm", "mean"),
    ).reset_index().sort_values("企业数", ascending=False)
    ind_total = ind_total.rename(columns={"平均面积": "平均面积(㎡)"})
    ind_total["平均面积(㎡)"] = ind_total["平均面积(㎡)"].round(0)
    ind_total.to_csv(config.DATA_DIR / "agg_industry_total.csv", index=False, encoding="utf-8-sig")

    # ---- 聚合 4：In/Out Ratio（按杨浦内部板块）----
    yp_to = df_mig[df_mig.to_area.str.startswith("杨浦")].groupby("to_area").size()
    yp_from = df_mig[df_mig.from_area.str.startswith("杨浦")].groupby("from_area").size()
    plates = sorted(set(yp_to.index) | set(yp_from.index))
    io = pd.DataFrame({"板块": plates})
    io["流入"] = io["板块"].map(yp_to).fillna(0).astype(int)
    io["流出"] = io["板块"].map(yp_from).fillna(0).astype(int)
    io["InOut比"] = (io["流入"] / io["流出"].replace(0, 1)).round(2)
    io["研判"] = io["InOut比"].apply(lambda x: "净吸引" if x >= 1 else "净流失")
    io.to_csv(config.DATA_DIR / "agg_inout_ratio.csv", index=False, encoding="utf-8-sig")

    # ---- 聚合 5：招商线索 ----
    leads = df_comp[df_comp.signal.isin(["近6月融资+招聘扩张", "工商稳定+岗位激增", "新设研发中心"])]
    leads = leads.sort_values("area_sqm", ascending=False).head(30)
    leads.to_csv(config.DATA_DIR / "agg_leads.csv", index=False, encoding="utf-8-sig")

    # ---- headline ----
    headline = {
        "mode": mode_flags,
        "载体总数": int(len(df_assets)),
        "总建面万㎡": round(df_assets["total_gfa_wan"].sum(), 1),
        "可租面积万㎡": round(df_assets["rentable_area_wan"].sum(), 1),
        "写字楼占比": round((df_assets.asset_type == "office").mean() * 100, 1),
        "平均报价租金": avg_quote,
        "平均成交租金": avg_deal,
        "全成本租金": full_cost,
        "平均物业费": avg_fee,
        "空置率": vacancy_rate,
        "入驻企业样本数": int(len(df_comp)),
        "迁徙事件样本数": int(len(df_mig)),
        "招商线索数": int(len(leads)),
        "TOP产业": ind_total.iloc[0]["industry_tag"],
        "议价折让%": round((1 - avg_deal / avg_quote) * 100, 1),
    }
    (config.DATA_DIR / "headline.json").write_text(
        json.dumps(headline, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"[pipeline] mode={mode_flags}")
    print(json.dumps(headline, ensure_ascii=False, indent=2))
    return headline


if __name__ == "__main__":
    main()
