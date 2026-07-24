# -*- coding: utf-8 -*-
"""生成策划案 Markdown 版，便于阅读与二次编辑。"""
import os

import content as C

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")


def build():
    L = []
    L.append(f"# {C.PROJECT_TITLE}")
    L.append(f"## {C.PROJECT_SUBTITLE}")
    L.append(f"**{C.PROJECT_TAG} · {C.VERSION} · {C.DOC_DATE}**")
    L.append("")
    L.append(f"> {C.ORGANIZER_LINE}  ")
    L.append(f"> 提交对象：**{C.RECIPIENT}**  ")
    L.append(f"> {C.PARTICIPANT_LINE}")
    L.append("")

    L.append("## 一、项目背景与目的")
    L.append("### 1.1 背景")
    for b in C.BACKGROUND:
        L.append(f"- {b}")
    L.append("")
    L.append("### 1.2 考察目的")
    for p in C.PURPOSE:
        L.append(f"- {p}")
    L.append("")

    L.append("## 二、组织方与接待单位")
    for party in C.PARTIES:
        L.append(f"### {party['role']} — {party['name']}")
        for d in party["duties"]:
            L.append(f"- {d}")
        L.append("")

    L.append("## 三、考察人员构成")
    L.append(f"**{C.SCALE}**")
    L.append("")
    L.append(C.TRAVEL_NOTE)
    L.append("")
    L.append("| 类别 | 说明 |")
    L.append("| --- | --- |")
    for a, b in C.PARTICIPANT_PROFILE:
        L.append(f"| {a} | {b} |")
    L.append("")

    L.append("## 四、考察主题与长期合作方向")
    for theme in C.COOP_THEMES:
        L.append(f"### {theme['name']}")
        for pt in theme["points"]:
            L.append(f"- {pt}")
        L.append("")

    L.append("## 五、时间安排建议")
    L.append(f"**{C.TIME_WINDOW}**")
    L.append("")
    L.append(C.DATE_CONFIRM)
    L.append("")
    L.append("| 方案 | 时长 | 适用情形 |")
    L.append("| --- | --- | --- |")
    for row in C.TIME_OPTIONS:
        L.append("| " + " | ".join(row) + " |")
    L.append("")
    L.append("**建议窗口**")
    for m in C.PREFERRED_MONTHS:
        L.append(f"- {m}")
    L.append("")

    L.append("## 六、建议行程方案")
    L.append("### 6.1 第一日（精华行程）")
    L.append("| 时间 | 事项 | 说明 |")
    L.append("| --- | --- | --- |")
    for row in C.DAY1_ITINERARY:
        L.append("| " + " | ".join(row) + " |")
    L.append("")
    L.append("### 6.2 第二日（两日深度行 · 推荐）")
    L.append("| 时间 | 事项 | 说明 |")
    L.append("| --- | --- | --- |")
    for row in C.DAY2_ITINERARY:
        L.append("| " + " | ".join(row) + " |")
    L.append("")
    L.append("### 6.3 建议参访与对接点")
    L.append("| 建议点位 | 交流重点 |")
    L.append("| --- | --- |")
    for a, b in C.SUGGESTED_SITES:
        L.append(f"| {a} | {b} |")
    L.append("")

    L.append("## 七、预期成果与长效合作")
    L.append("### 7.1 本次预期成果")
    for o in C.OUTCOMES:
        L.append(f"- {o}")
    L.append("")
    L.append("### 7.2 长效合作设想")
    for name, desc in C.LONG_TERM:
        L.append(f"- **{name}**：{desc}")
    L.append("")

    L.append("## 八、恳请接待单位支持事项")
    L.append("| 支持事项 | 具体说明 |")
    L.append("| --- | --- |")
    for a, b in C.SUPPORT_REQUESTS:
        L.append(f"| {a} | {b} |")
    L.append("")

    L.append("## 九、费用与分工原则")
    for c in C.COST_PRINCIPLES:
        L.append(f"- {c}")
    L.append("")

    L.append("## 十、下一步工作安排")
    L.append("| 步骤 | 工作内容 |")
    L.append("| --- | --- |")
    for a, b in C.NEXT_STEPS:
        L.append(f"| {a} | {b} |")
    L.append("")
    L.append(C.CONTACT_NOTE)
    L.append("")
    L.append("## 结语")
    L.append(C.CLOSING)
    L.append("")
    L.append("---")
    L.append("")
    L.append("复旦大学住房政策研究中心  ")
    L.append("上海市杨浦区科技企业联合会  ")
    L.append(C.DOC_DATE)
    L.append("")

    os.makedirs(OUT_DIR, exist_ok=True)
    out = os.path.join(OUT_DIR, "赴宁波港及宁波经济技术开发区考察交流策划案.md")
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    print("saved:", out)


if __name__ == "__main__":
    build()
