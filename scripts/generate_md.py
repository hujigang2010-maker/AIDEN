# -*- coding: utf-8 -*-
"""生成思路一页纸 Markdown 与致潘嘉琰沟通稿。"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import content as C

OUT_MD = Path(__file__).resolve().parent.parent / "output" / "港大经管上海中心_联合策划思路一页纸.md"
OUT_WX = Path(__file__).resolve().parent.parent / "output" / "致潘嘉琰_沟通稿.txt"


def build_md(path: Path | None = None) -> Path:
    path = path or OUT_MD
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        f"# {C.PROJECT_NAME}",
        "",
        f"**{C.PROJECT_SUBTITLE}**  ",
        f"{C.VERSION} · {C.DATE_CN}  ",
        f"致：{C.THEIR_CONTACT}（{C.THEIR_TITLE}） / {C.THEIR_UNIT}",
        "",
        "## 一句话",
        "",
        C.ONE_LINER,
        "",
        "## 为什么找上海中心",
        "",
    ]
    for x in C.THEIR_WINDOW:
        lines.append(f"- {x}")
    lines += [
        "",
        "## 思路",
        "",
        "港大经管上海中心不缺品牌、场地和院方师资。缺的是把「产业转化平台」变成高频、高质量、可跟进的到场人群。",
        "",
        "潘老师负责华东 EMBA，客群是 8 年以上管理经验的董事长、总裁、创始人。杨浦科企联与产业圈层里，正是这一批人。",
        "",
        "所以不做宣讲会外包，做 **联合招生场景**：",
        "",
    ]
    for x in C.PRODUCT_LOGIC:
        lines.append(f"- {x}")
    lines += [
        "",
        "### 三件套",
        "",
    ]
    for p in C.PRODUCTS:
        lines.append(f"- **{p['name']}**（{p['tag']}）：{p['desc']}")
    lines += [
        "",
        "## 如何形成合作：只收一笔前期费用",
        "",
        "不碰学费分成，不承诺录取人数。签一份短协议，合作就算成立。",
        "",
        f"- **费用**：{C.FEE_NAME} **{C.FEE_AMOUNT_CN}（¥{C.FEE_AMOUNT:,}）**",
        f"- **支付**：协议生效后 {C.FEE_DAYS} 个工作日内一次性支付",
        f"- **覆盖**：{C.PLAN_DAYS} 天策划 + 首场闭门课（{C.FIRST_EVENT_SIZE}）+ 会后纪要与共管名单",
        "",
        "覆盖明细：",
        "",
    ]
    for x in C.FEE_COVERS:
        lines.append(f"- {x}")
    lines += [
        "",
        "不包含：",
        "",
    ]
    for x in C.FEE_NOT_COVERED:
        lines.append(f"- {x}")
    lines += [
        "",
        "## 90 天",
        "",
        "| 阶段 | 节点 | 做什么 |",
        "| --- | --- | --- |",
    ]
    for a, b, c in C.NINETY_DAY:
        lines.append(f"| {a} | {b} | {c} |")
    lines += [
        "",
        "## 下一步",
        "",
    ]
    for i, x in enumerate(C.NEXT_STEPS, 1):
        lines.append(f"{i}. {x}")
    lines += [
        "",
        "---",
        "",
        f"接口：{C.THEIR_CONTACT}  {C.THEIR_TEL}  {C.THEIR_EMAIL}  {C.THEIR_ADDR}",
        "",
        "配套：联合策划方案 PPT、合作协议建议稿、前期费用与交付 Excel。",
        "",
    ]
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


def build_wechat(path: Path | None = None) -> Path:
    path = path or OUT_WX
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(C.WECHAT.strip() + "\n", encoding="utf-8")
    return path


def build() -> tuple[Path, Path]:
    return build_md(), build_wechat()


if __name__ == "__main__":
    md, wx = build()
    print(f"已生成 {md}")
    print(f"已生成 {wx}")
