# -*- coding: utf-8 -*-
"""生成『群邦·领事会客厅』策划案 Markdown 版（便于阅读与二次编辑）。

用法：
    python3 scripts/generate_md.py [输出路径.md]
默认输出 output/群邦-领事会客厅-策划案.md
"""
import sys
import os
import content as C


def md_table(t):
    out = [f"### {t['title']}", ""]
    out.append("| " + " | ".join(t["headers"]) + " |")
    out.append("| " + " | ".join(["---"] * len(t["headers"])) + " |")
    for row in t["rows"]:
        out.append("| " + " | ".join(str(x) for x in row) + " |")
    if t.get("note"):
        out += ["", f"> 说明：{t['note']}"]
    out.append("")
    return "\n".join(out)


def build(path):
    L = []
    L.append(f"# {C.PROJECT_NAME}")
    L.append(f"## {C.PROJECT_SUBTITLE}")
    L.append(f"**{C.PROJECT_TAG} · {C.VERSION}**")
    L.append("")
    L.append("> 合作三方：复旦大学政策研究中心 × 上海市杨浦区科技企业联合会 × "
             "郡邦（上海）文化交流发展有限公司（总领事俱乐部 CGC 运营方，金茂大厦86层）")
    L.append("")

    L.append("## 一、合作对象：总领事俱乐部（CGC）")
    for b in C.CGC_PROFILE:
        L.append(f"- {b}")
    L.append("")
    L.append("**CGC 既有品牌资产（升级变现）**")
    for b in C.EXISTING_ASSETS:
        L.append(f"- {b}")
    L.append("")

    L.append("## 二、背景与机遇")
    L.append("**资源现状**")
    for b in C.BACKGROUND:
        L.append(f"- {b}")
    L.append("")
    L.append("**机遇与思路**")
    for b in C.OPPORTUNITY:
        L.append(f"- {b}")
    L.append("")

    L.append("## 三、合作三方与定位")
    for p in C.PARTIES:
        L.append(f"### {p['name'].replace(chr(10), ' ')}　—　{p['role']}")
        for d in p["duties"]:
            L.append(f"- {d}")
        L.append("")

    L.append("## 四、长三角领区：把领事资源变成招商抓手")
    L.append(C.REGION["intro"])
    L.append("")
    L.append("覆盖：上海 · 江苏 · 浙江 · 安徽")
    L.append("")
    for x in C.REGION["apps"]:
        L.append(f"- {x}")
    L.append("")

    L.append("## 五、总体战略：以领事为纽带的引领模式")
    L.append("价值闭环：**活动获客 → 会籍沉淀 → 出海·招商 → 政企复购**，"
             "形成可持续现金流。‘小道大道’：先以小型高频活动建立链接，"
             "再放大为城市级开放平台与变现体系。")
    L.append("")
    for k, v in C.STRATEGY_PILLARS:
        L.append(f"- **{k}**：{v}")
    L.append("")

    L.append("## 六、产品矩阵：六大产品线")
    for pl in C.PRODUCT_LINES:
        L.append(f"### {pl['no']} {pl['name']} —— {pl['tagline']}")
        for pt in pl["points"]:
            L.append(f"- {pt}")
        L.append("")

    L.append("## 七、旗舰产品：国家会客厅")
    L.append(f"**定位**：{C.SALON_MODEL['concept']}")
    L.append("")
    L.append("**四大功能**")
    for x in C.SALON_MODEL["functions"]:
        L.append(f"- {x}")
    L.append("")
    L.append("**收入来源**")
    for x in C.SALON_MODEL["revenue"]:
        L.append(f"- {x}")
    L.append("")
    L.append("**政府 / 园区价值**")
    for x in C.SALON_MODEL["gov_value"]:
        L.append(f"- {x}")
    L.append("")

    L.append("## 八、高端商务出海 + 人文特色")
    L.append(f"**节奏**：{C.OUTBOUND['rhythm']}")
    L.append("")
    for x in C.OUTBOUND["highlights"]:
        L.append(f"- {x}")
    L.append("")

    L.append("## 九、收费体系与测算（表格）")
    for t in C.ALL_TABLES:
        L.append(md_table(t))

    L.append("## 十、合规与风险管理")
    for x in C.COMPLIANCE:
        L.append(f"- {x}")
    L.append("")

    L.append("## 十一、下一步行动建议")
    for i, x in enumerate(C.NEXT_STEPS, 1):
        L.append(f"{i}. {x}")
    L.append("")

    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    print(f"已生成 Markdown：{path}")


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "output/群邦-领事会客厅-策划案.md"
    build(out)
