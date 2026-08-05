#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""上海板块 × 学区融合分析：按差异化选房逻辑评分，输出 Excel 与 HTML。"""

from __future__ import annotations

import json
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# 允许从 scripts 目录直接运行
sys.path.insert(0, str(Path(__file__).resolve().parent))
from shanghai_plates_data import (  # noqa: E402
    CENTRAL_DISTRICTS,
    EXCLUDED_DISTRICTS,
    EXPLICIT_EXCLUDE,
    EXPLICIT_RECOMMEND,
    MOSTLY_EXCLUDED_DISTRICTS,
    PLATES,
)

OUTPUT_DIR = Path(__file__).resolve().parent.parent / "output"

# 评分权重（差异化选房逻辑）
WEIGHTS = {
    "population_hollow": 0.30,  # 人口空心化：越高越好
    "low_new_resident": 0.25,   # 新上海人占比低：越高越好
    "low_competition": 0.25,    # 升学竞争低：越高越好
    "primary_quality": 0.12,    # 小学质量
    "price_fit": 0.08,          # 200–300 万区间匹配度
}

TIER_COLORS = {
    "⭐⭐⭐ 强烈推荐": "#1a7f37",
    "⭐⭐ 可考虑": "#0969da",
    "⭐ 谨慎": "#bf8700",
    "❌ 排除": "#cf222e",
}

TIER_FILLS = {
    "⭐⭐⭐ 强烈推荐": "C6EFCE",
    "⭐⭐ 可考虑": "D0E8FF",
    "⭐ 谨慎": "FFF2CC",
    "❌ 排除": "FADBD8",
}


def _norm(value: float, low: float = 1, high: float = 5) -> float:
    return max(0.0, min(1.0, (value - low) / (high - low)))


def price_fit_score(price_low: int, price_high: int) -> float:
    """200–300 万为最优区间（纪要推荐潍坊板块价位）。"""
    target_low, target_high = 200, 300
    if price_low <= target_high and price_high >= target_low:
        overlap = min(price_high, target_high) - max(price_low, target_low)
        span = max(price_high - price_low, 1)
        return 1.0 + 4.0 * (overlap / span)
    if price_low > target_high:
        # 越贵越低分
        return max(1.0, 5.0 - (price_low - target_high) / 100)
    # 太便宜可能学校资源弱，略降
    return max(2.0, 4.0 - (target_low - price_high) / 50)


def compute_score(plate: dict) -> float:
    hollow = plate["population_hollow"]
    low_new = 6 - plate["new_resident_ratio"]  # 反向
    low_comp = 6 - plate["school_competition"]
    primary = plate["primary_quality"]
    price = price_fit_score(plate["price_low"], plate["price_high"])

    score = (
        WEIGHTS["population_hollow"] * hollow
        + WEIGHTS["low_new_resident"] * low_new
        + WEIGHTS["low_competition"] * low_comp
        + WEIGHTS["primary_quality"] * primary
        + WEIGHTS["price_fit"] * price
    )
    return round(score, 2)


def assign_tier(plate: dict, score: float) -> str:
    key = (plate["district"], plate["plate"])
    district = plate["district"]

    if key in EXPLICIT_EXCLUDE:
        return "❌ 排除"
    if key in EXPLICIT_RECOMMEND:
        return "⭐⭐⭐ 强烈推荐"
    if district in EXCLUDED_DISTRICTS:
        return "❌ 排除"
    if district in MOSTLY_EXCLUDED_DISTRICTS and key not in EXPLICIT_RECOMMEND:
        if score >= 3.8 and plate["population_hollow"] >= 4:
            return "⭐⭐ 可考虑"
        return "❌ 排除"
    # 自动强烈推荐：中心城区 + 人口空心化 + 低导入 + 小学质量达标
    if (
        score >= 3.6
        and plate["population_hollow"] >= 4
        and plate["new_resident_ratio"] <= 3
        and plate["primary_quality"] >= 3
        and district in CENTRAL_DISTRICTS
    ):
        return "⭐⭐⭐ 强烈推荐"
    # 空心化明显但学校一般的中心板块 → 可考虑
    if (
        score >= 3.5
        and plate["population_hollow"] >= 4
        and plate["new_resident_ratio"] <= 3
        and district in CENTRAL_DISTRICTS
    ):
        return "⭐⭐ 可考虑"
    if score >= 3.0:
        return "⭐⭐ 可考虑"
    if score >= 2.2:
        return "⭐ 谨慎"
    return "❌ 排除"


def enrich_plates() -> list[dict]:
    rows = []
    for p in PLATES:
        row = dict(p)
        row["score"] = compute_score(p)
        row["tier"] = assign_tier(p, row["score"])
        row["price_range"] = f"{p['price_low']}–{p['price_high']}万"
        row["marked"] = row["tier"] == "⭐⭐⭐ 强烈推荐"
        rows.append(row)
    rows.sort(key=lambda r: (-r["score"], r["district"], r["plate"]))
    return rows


def write_excel(rows: list[dict], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "板块学区融合"

    headers = [
        "推荐等级", "综合分", "行政区", "板块", "对应街道",
        "对口小学（代表）", "对口初中（代表）", "入学方式", "五年一户",
        "人口空心化", "新上海人占比", "升学竞争", "小学质量", "初中质量",
        "总价区间", "备注",
    ]
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 标题行
    ws.merge_cells("A1:P1")
    ws["A1"] = f"上海板块 × 学区融合分析（差异化选房逻辑）— 生成日期 {date.today()}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center")

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2F5496")
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border

    for i, r in enumerate(rows, 3):
        values = [
            r["tier"], r["score"], r["district"], r["plate"], r["street"],
            r["primary_schools"], r["middle_schools"], r["enrollment"], r["five_year"],
            r["population_hollow"], r["new_resident_ratio"], r["school_competition"],
            r["primary_quality"], r["middle_quality"], r["price_range"], r["notes"],
        ]
        fill_color = TIER_FILLS.get(r["tier"], "FFFFFF")
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if col == 1:
                cell.fill = PatternFill("solid", fgColor=fill_color)
                cell.font = Font(bold=True)

    # 说明 sheet
    ws2 = wb.create_sheet("评分说明")
    notes = [
        ["维度", "权重", "说明"],
        ["人口空心化", "30%", "居住人口相对学校容量偏少，升学竞争压力小（纪要核心指标）"],
        ["新上海人占比（反向）", "25%", "导入人口少、老上海聚集的板块优先"],
        ["升学竞争（反向）", "25%", "同板块同龄生源密度与热门程度"],
        ["小学质量", "12%", "公办小学口碑（非官方排名，仅供参考）"],
        ["价格匹配 200–300万", "8%", "纪要推荐潍坊板块价位区间"],
        ["", "", ""],
        ["推荐等级", "", "说明"],
        ["⭐⭐⭐ 强烈推荐", "", "符合人口空心化+低竞争+纪要点名推荐"],
        ["⭐⭐ 可考虑", "", "部分指标良好，需结合具体对口与入户年限核实"],
        ["⭐ 谨慎", "", "学区或竞争存在明显短板"],
        ["❌ 排除", "", "纪要建议排除区域（嘉定/松江/大部分闵行浦东等）或高竞争新区"],
        ["", "", ""],
        ["重要提示", "", "本表为决策辅助工具，非官方学区划片；请以各区教育局当年招生简章为准"],
    ]
    for ri, row in enumerate(notes, 1):
        for ci, val in enumerate(row, 1):
            ws2.cell(row=ri, column=ci, value=val)

    col_widths = [14, 8, 8, 16, 14, 28, 24, 12, 10, 10, 12, 10, 8, 8, 12, 36]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A3"
    wb.save(path)


def write_html(rows: list[dict], path: Path) -> None:
    recommended = [r for r in rows if r["tier"] == "⭐⭐⭐ 强烈推荐"]
    data_json = json.dumps(rows, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>上海板块 × 学区融合分析</title>
<style>
:root {{
  --green: #1a7f37; --blue: #0969da; --yellow: #bf8700; --red: #cf222e;
  --bg: #f6f8fa; --card: #fff; --border: #d0d7de;
}}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
  background: var(--bg); color: #1f2328; line-height: 1.6; padding: 24px; }}
h1 {{ font-size: 1.5rem; margin-bottom: 8px; }}
.subtitle {{ color: #656d76; margin-bottom: 20px; font-size: 0.95rem; }}
.summary {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 24px; }}
.stat {{ background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 16px; }}
.stat .num {{ font-size: 1.8rem; font-weight: 700; }}
.stat .label {{ font-size: 0.85rem; color: #656d76; }}
.filters {{ display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 16px; align-items: center; }}
.filters input, .filters select {{ padding: 8px 12px; border: 1px solid var(--border); border-radius: 6px; font-size: 14px; }}
.filters label {{ font-size: 14px; color: #656d76; }}
.card {{ background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 20px; margin-bottom: 20px; }}
.card h2 {{ font-size: 1.1rem; margin-bottom: 12px; }}
.highlight-list {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 10px; }}
.highlight-item {{ border-left: 4px solid var(--green); padding: 10px 14px; background: #dafbe1; border-radius: 0 6px 6px 0; }}
.highlight-item strong {{ display: block; margin-bottom: 4px; }}
table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
th, td {{ border: 1px solid var(--border); padding: 8px 10px; text-align: left; vertical-align: top; }}
th {{ background: #2f5496; color: #fff; position: sticky; top: 0; z-index: 1; }}
tr:nth-child(even) {{ background: #f6f8fa; }}
tr.row-strong {{ background: #dafbe1 !important; }}
.tier {{ font-weight: 700; white-space: nowrap; }}
.tier-strong {{ color: var(--green); }}
.tier-ok {{ color: var(--blue); }}
.tier-caution {{ color: var(--yellow); }}
.tier-exclude {{ color: var(--red); }}
.table-wrap {{ overflow-x: auto; max-height: 70vh; overflow-y: auto; border: 1px solid var(--border); border-radius: 8px; }}
.logic {{ font-size: 0.9rem; color: #424a53; }}
.logic li {{ margin-left: 20px; margin-bottom: 6px; }}
.footer {{ margin-top: 24px; font-size: 0.8rem; color: #656d76; }}
</style>
</head>
<body>
<h1>上海板块 × 学区融合分析</h1>
<p class="subtitle">基于差异化选房逻辑 · 生成日期 {date.today()} · 共 {len(rows)} 个板块</p>

<div class="summary">
  <div class="stat"><div class="num" style="color:var(--green)">{len(recommended)}</div><div class="label">⭐⭐⭐ 强烈推荐</div></div>
  <div class="stat"><div class="num">{len([r for r in rows if r['tier']=='⭐⭐ 可考虑'])}</div><div class="label">⭐⭐ 可考虑</div></div>
  <div class="stat"><div class="num">{len([r for r in rows if r['tier']=='⭐ 谨慎'])}</div><div class="label">⭐ 谨慎</div></div>
  <div class="stat"><div class="num">{len([r for r in rows if r['tier']=='❌ 排除'])}</div><div class="label">❌ 排除</div></div>
</div>

<div class="card">
<h2>🎯 强烈推荐板块（已标记）</h2>
<div class="highlight-list" id="recommended-list"></div>
</div>

<div class="card">
<h2>📋 选房逻辑摘要</h2>
<ul class="logic">
<li><strong>核心目标</strong>：获取上学资格，非投资增值；优先小学/初中学区，高中靠统考</li>
<li><strong>差异化策略</strong>：选人口空心化、老上海聚集、本地人口流出的老板块，避开新上海人集中导入的新区</li>
<li><strong>纪要推荐</strong>：黄浦全区、杨浦滨江、浦东潍坊新村（明珠小学，200–300万）</li>
<li><strong>纪要排除</strong>：嘉定、松江、闵行大部分、浦东大部分、九亭等高竞争板块</li>
<li><strong>学段策略</strong>：可先解决小学学区，初中阶段再搬家换学区，不必一步到位</li>
</ul>
</div>

<div class="card">
<div class="filters">
  <input type="search" id="search" placeholder="搜索板块/学校/备注…" style="min-width:220px">
  <select id="district-filter"><option value="">全部行政区</option></select>
  <select id="tier-filter">
    <option value="">全部等级</option>
    <option value="⭐⭐⭐ 强烈推荐">⭐⭐⭐ 强烈推荐</option>
    <option value="⭐⭐ 可考虑">⭐⭐ 可考虑</option>
    <option value="⭐ 谨慎">⭐ 谨慎</option>
    <option value="❌ 排除">❌ 排除</option>
  </select>
  <label><input type="checkbox" id="only-marked"> 仅看强烈推荐</label>
</div>
<div class="table-wrap">
<table id="data-table">
<thead><tr>
  <th>等级</th><th>综合分</th><th>区</th><th>板块</th><th>小学</th><th>初中</th>
  <th>入学</th><th>空心化</th><th>新上海人</th><th>竞争</th><th>总价</th><th>备注</th>
</tr></thead>
<tbody></tbody>
</table>
</div>
</div>

<p class="footer">⚠️ 数据为公开信息整理 + 纪要逻辑评分，非官方学区划片。报名前务必查阅所在区 2025/2026 年招生简章，核实对口地段与入户年限。</p>

<script>
const DATA = {data_json};

const tierClass = {{
  "⭐⭐⭐ 强烈推荐": "tier-strong",
  "⭐⭐ 可考虑": "tier-ok",
  "⭐ 谨慎": "tier-caution",
  "❌ 排除": "tier-exclude"
}};

// 推荐卡片
const recList = document.getElementById("recommended-list");
DATA.filter(r => r.tier === "⭐⭐⭐ 强烈推荐").forEach(r => {{
  const div = document.createElement("div");
  div.className = "highlight-item";
  div.innerHTML = `<strong>${{r.district}} · ${{r.plate}}</strong>
    综合分 ${{r.score}} · ${{r.price_range}}<br>
    小学：${{r.primary_schools}}<br>
    ${{r.notes}}`;
  recList.appendChild(div);
}});

// 区筛选
const districts = [...new Set(DATA.map(r => r.district))].sort();
const dSel = document.getElementById("district-filter");
districts.forEach(d => {{ const o = document.createElement("option"); o.value = d; o.textContent = d; dSel.appendChild(o); }});

function render() {{
  const q = document.getElementById("search").value.toLowerCase();
  const dist = document.getElementById("district-filter").value;
  const tier = document.getElementById("tier-filter").value;
  const onlyMarked = document.getElementById("only-marked").checked;
  const tbody = document.querySelector("#data-table tbody");
  tbody.innerHTML = "";

  DATA.filter(r => {{
    if (onlyMarked && r.tier !== "⭐⭐⭐ 强烈推荐") return false;
    if (dist && r.district !== dist) return false;
    if (tier && r.tier !== tier) return false;
    if (q) {{
      const hay = [r.plate, r.district, r.primary_schools, r.middle_schools, r.notes, r.street].join(" ").toLowerCase();
      if (!hay.includes(q)) return false;
    }}
    return true;
  }}).forEach(r => {{
    const tr = document.createElement("tr");
    if (r.tier === "⭐⭐⭐ 强烈推荐") tr.className = "row-strong";
    tr.innerHTML = `
      <td class="tier ${{tierClass[r.tier] || ''}}">${{r.tier}}</td>
      <td>${{r.score}}</td>
      <td>${{r.district}}</td>
      <td><strong>${{r.plate}}</strong><br><small>${{r.street}}</small></td>
      <td>${{r.primary_schools}}</td>
      <td>${{r.middle_schools}}</td>
      <td>${{r.enrollment}}</td>
      <td>${{r.population_hollow}}/5</td>
      <td>${{r.new_resident_ratio}}/5</td>
      <td>${{r.school_competition}}/5</td>
      <td>${{r.price_range}}</td>
      <td>${{r.notes}}</td>`;
    tbody.appendChild(tr);
  }});
}}

["search","district-filter","tier-filter","only-marked"].forEach(id =>
  document.getElementById(id).addEventListener("input", render));
render();
</script>
</body>
</html>"""

    path.write_text(html, encoding="utf-8")


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    rows = enrich_plates()

    xlsx_path = OUTPUT_DIR / "上海板块学区融合分析.xlsx"
    html_path = OUTPUT_DIR / "上海板块学区融合分析.html"

    write_excel(rows, xlsx_path)
    write_html(rows, html_path)

    recommended = [r for r in rows if r["tier"] == "⭐⭐⭐ 强烈推荐"]
    print(f"已生成 {len(rows)} 个板块数据")
    print(f"Excel: {xlsx_path}")
    print(f"HTML:  {html_path}")
    print(f"\n⭐⭐⭐ 强烈推荐（{len(recommended)} 个）：")
    for r in recommended:
        print(f"  · {r['district']} {r['plate']}（{r['score']}分）— {r['primary_schools']}")


if __name__ == "__main__":
    main()
