# -*- coding: utf-8 -*-
"""生成《首场·出海东南亚 总领事商务论坛》方案 Markdown 版。

用法：
    python3 scripts/generate_event_md.py [输出路径.md]
"""
import sys
import os
import event_content as C


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
    L.append("> 背景来源：总领事俱乐部首场活动筹备会议（AI 纪要 + 完整逐字稿）。")
    L.append("")

    L.append("## 一、三大目标")
    for x in C.OBJECTIVES:
        L.append(f"- {x}")
    L.append("")
    L.append("## 择期：黄道吉日")
    L.append(f"- **活动日期**：{C.HUANGLI['date']}（下午 13:00–18:30）")
    L.append(f"- **【宜】**：{C.HUANGLI['yi']}")
    L.append(f"- **【忌】**：{C.HUANGLI['ji']}")
    L.append(f"- **【冲】**：{C.HUANGLI['chong']}")
    L.append(f"- **择此日之由**：{C.HUANGLI['why']}")
    for x in C.DATE_BACKUP:
        L.append(f"- {x}")
    L.append(f"- {C.HOLIDAY_NOTE}")
    L.append("")
    L.append("## 二、首场定位")
    for x in C.POSITIONING:
        L.append(f"- {x}")
    L.append("")
    L.append("## 三、盈利逻辑")
    for x in C.KEY_MODEL:
        L.append(f"- {x}")
    L.append("")

    L.append("## 四、活动方案、收费与赞助（表格）")
    for t in C.ALL_TABLES:
        L.append(md_table(t))

    L.append("## 五、复旦品牌价值保护")
    for x in C.BRAND_PROTECT:
        L.append(f"- {x}")
    L.append("")
    L.append("## 六、双平台价值最大化")
    for name, desc in C.PLATFORM_VALUE:
        L.append(f"- **{name}**：{desc}")
    L.append("")
    L.append("## 七、可行性评估")
    for x in C.FEASIBILITY:
        L.append(f"- {x}")
    L.append("")
    L.append("## 八、下一步行动")
    for i, x in enumerate(C.NEXT_STEPS, 1):
        L.append(f"{i}. {x}")
    L.append("")

    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    print(f"已生成 Markdown：{path}")


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "output/首场出海论坛-方案.md"
    build(out)
