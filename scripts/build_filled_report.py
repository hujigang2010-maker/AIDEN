"""读取聚合结果，生成"杨浦区数据填充版"Markdown 报告（内嵌图表与数据表）。"""
import json
import pandas as pd

import config

D = config.DATA_DIR
OUT = config.ROOT / "上海商办楼宇与产业园区市场报告-杨浦区数据填充版.md"
CHART_REL = "charts"  # 相对路径供 pandoc 引用


def md_table(df, cols=None):
    if cols:
        df = df[cols]
    header = "| " + " | ".join(map(str, df.columns)) + " |"
    sep = "| " + " | ".join(["---"] * len(df.columns)) + " |"
    rows = ["| " + " | ".join(map(lambda v: str(v), r)) + " |"
            for r in df.itertuples(index=False)]
    return "\n".join([header, sep] + rows)


def main():
    h = json.loads((D / "headline.json").read_text(encoding="utf-8"))
    supply = pd.read_csv(D / "agg_supply_by_subdistrict.csv")
    market = pd.read_csv(D / "agg_market_by_plate.csv")
    ind = pd.read_csv(D / "agg_industry_total.csv")
    io = pd.read_csv(D / "agg_inout_ratio.csv")
    leads = pd.read_csv(D / "agg_leads.csv")

    mode_note = ("> ⚠️ **本版数据为「演示数据模式」自动生成**（云端未注入高德/企查查 API Key）。"
                 "脚本已具备真实采集能力，在 `scripts/.env` 填入真实 Key 后重跑即切换为实采数据。\n")

    leads_view = leads[["name", "industry_tag", "scale", "area_sqm", "plate", "signal"]].head(15)
    leads_view = leads_view.rename(columns={
        "name": "企业(脱敏)", "industry_tag": "产业标签", "scale": "规模",
        "area_sqm": "当前面积(㎡)", "plate": "所在板块", "signal": "信号"})

    parts = []
    parts.append("# 上海市杨浦区街道级商办市场月度动态（数据填充版）\n")
    parts.append("> 易居房地产研究院 × AI 数据获取方 ｜ 由采集管线自动生成 ｜ 版本：填充版 V1.0\n")
    parts.append(mode_note)

    parts.append("## 一、关键指标速览\n")
    kpi = pd.DataFrame([
        ["载体总数（样本）", h["载体总数"], "总建面（万㎡）", h["总建面万㎡"]],
        ["可租面积（万㎡）", h["可租面积万㎡"], "写字楼占比（%）", h["写字楼占比"]],
        ["平均报价租金（元/㎡/天）", h["平均报价租金"], "平均成交租金（元/㎡/天）", h["平均成交租金"]],
        ["全成本租金（元/㎡/天）", h["全成本租金"], "议价折让（%）", h["议价折让%"]],
        ["空置率（招商口径,%）", h["空置率"], "TOP 产业", h["TOP产业"]],
        ["入驻企业样本数", h["入驻企业样本数"], "迁徙事件样本数", h["迁徙事件样本数"]],
    ], columns=["指标", "数值", "指标 ", "数值 "])
    parts.append(md_table(kpi) + "\n")

    parts.append("## 二、全域供给：分街道分布\n")
    parts.append(f"![分街道可租赁面积]({CHART_REL}/01_supply_by_subdistrict.png)\n")
    parts.append(md_table(supply, ["subdistrict", "载体数量", "总建面(万㎡)",
                                   "可租面积(万㎡)", "写字楼数", "产业园数"]) + "\n")
    parts.append(f"\n![物业等级结构]({CHART_REL}/02_grade_structure.png)\n")

    parts.append("## 三、市场行情：分板块租金与空置分化\n")
    parts.append(f"![分板块租金与空置]({CHART_REL}/03_rent_vacancy_by_plate.png)\n")
    parts.append(md_table(market, ["plate", "可租面积(万㎡)", "空置率%", "成交租金"]) + "\n")

    parts.append("## 四、需求结构：主导产业画像\n")
    parts.append(f"![主导产业]({CHART_REL}/04_industry_mix.png)\n")
    parts.append(md_table(ind, ["industry_tag", "企业数", "平均面积(㎡)"]) + "\n")

    parts.append("## 五、企业迁徙与招商吸引力\n")
    parts.append(f"![In/Out Ratio]({CHART_REL}/05_inout_ratio.png)\n")
    parts.append(md_table(io, ["板块", "流入", "流出", "InOut比", "研判"]) + "\n")
    parts.append(f"\n![迁入来源地]({CHART_REL}/06_migration_source.png)\n")

    parts.append("## 六、精准招商线索（TOP15）\n")
    parts.append(md_table(leads_view) + "\n")
    parts.append("\n> 筛选逻辑：近 6 个月融资 + 招聘扩张 + 工商稳定（未异常）交叉。企业名称已脱敏。\n")

    parts.append("\n---\n")
    parts.append("> 数据来源：高德 POI/AOI（载体）、企查查/天眼查（企业与工商变更）、"
                 "好租/办办网/点点租（挂牌行情）；演示模式下为合成样本。"
                 "正式版填入真实 Key 后由相同管线产出。遵循《上海市数据条例》及各平台条款。\n")

    OUT.write_text("\n".join(parts), encoding="utf-8")
    print("written", OUT)


if __name__ == "__main__":
    main()
