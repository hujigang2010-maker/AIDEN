# -*- coding: utf-8 -*-
"""生成思路一页纸 Markdown 与致潘嘉琰沟通稿。"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import content as C

OUT_MD = Path(__file__).resolve().parent.parent / "output" / "港大经管上海中心_联合策划思路一页纸.md"
OUT_WX = Path(__file__).resolve().parent.parent / "output" / "致潘嘉琰_沟通稿.txt"
OUT_NEEDS = Path(__file__).resolve().parent.parent / "output" / "港大经管上海中心_对方需求整理_20260903.md"


def build_needs(path: Path | None = None) -> Path:
    path = path or OUT_NEEDS
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        f"# 对方需求整理 · {C.MEETING_DATE}",
        "",
        f"来源：港大经管上海中心与校友活动平台合作洽谈交流记录。整理口径用于方案 V1.1，不替代原始录音。",
        "",
        "## 对方是谁、管什么",
        "",
        f"- 接口人：{C.THEIR_CONTACT}。架构上属于 EMBA 条线，借上海中心平台做 **华东、华中** 招生与市场。",
        "- 自己的定位：EMBA 粘性，不是中心主任。经管学院层合作要主任点头，招生官不能越权承诺。",
        "- 中心四个点位：外滩 22 号总部接待处（装修中、无办学功能）、经管中心约 5000㎡（已用）、张江计教学院约 3 万㎡（已用）、漕河泾智能研究所（未启用）。",
        "",
        "## 听懂的四句话",
        "",
    ]
    for t, d in C.HEARD:
        lines.append(f"- **{t}**：{d}")
    lines += [
        "",
        "## 需求四件事",
        "",
    ]
    for t, d in C.THEIR_NEEDS:
        lines.append(f"- **{t}**：{d}")
    lines += [
        "",
        "## 中心职能（对方口述）",
        "",
    ]
    for x in C.CENTRE_FUNCTIONS:
        lines.append(f"- {x}")
    lines += [
        "",
        "## 方案怎么接住",
        "",
        "| 对方要的 | 本期怎么做 | 是否验收 |",
        "| --- | --- | --- |",
        "| 高质量到场，做华东华中高端教育 | 外滩出海闭门课 + 名单共管 + 会后一对一 | 验收 |",
        "| 先让主任有直观感受 | 借我方总领事/企业家出海场，书面邀请主任 | 验收邀约，不强制出席 |",
        "| 从经管学院做，不限于 EMBA | 协议分层：A 粘性 / B 主任 / C 延伸 | 层 C 只出备忘 |",
        "| 出海课、原创谷、海外模块 | 90 天内书面备忘，另签再生效 | 不验收 |",
        "",
        "## 故意不用的口径",
        "",
        "洽谈闲聊涉及稳定币清结算、自贸区转口套利等非正式跨境玩法。对外方案、协议、现场一律不采用。"
        "出海只讲合法的企业国际化、港股路径、跨境贸易合规与金融基础设施。",
        "",
        f"商业结构不变：{C.FEE_NAME} {C.FEE_AMOUNT_CN}，一次性支付。",
        "",
    ]
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


def build_md(path: Path | None = None) -> Path:
    path = path or OUT_MD
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        f"# {C.PROJECT_NAME}",
        "",
        f"**{C.PROJECT_SUBTITLE}**  ",
        f"{C.VERSION} · {C.DATE_CN}  ",
        f"依据：{C.MEETING_DATE} 外滩中心交流",
        "",
        "## 一句话",
        "",
        C.ONE_LINER,
        "",
        "## 对方要什么",
        "",
    ]
    for t, d in C.HEARD:
        lines.append(f"- **{t}**：{d}")
    lines += [
        "",
        "## 三层合作",
        "",
    ]
    for p in C.LAYERS:
        lines.append(f"- **{p['name']}**（{p['tag']}）：{p['desc']}")
    lines += [
        "",
        "所以不做宣讲会外包，主赛道按对方口径定为 **出海**：",
        "",
    ]
    for x in C.PRODUCT_LOGIC:
        lines.append(f"- {x}")
    lines += [
        "",
        "## 如何形成合作：只收一笔前期费用",
        "",
        "不碰学费分成，不承诺录取人数。签一份短协议，合作就算成立。",
        "",
        f"- **费用**：{C.FEE_NAME} **{C.FEE_AMOUNT_CN}（¥{C.FEE_AMOUNT:,}）**",
        f"- **支付**：协议生效后 {C.FEE_DAYS} 个工作日内一次性支付",
        f"- **覆盖**：{C.PLAN_DAYS} 天策划 + 外滩出海首场（{C.FIRST_EVENT_SIZE}）+ 一次主任体验邀约 + 会后纪要",
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


def build() -> tuple[Path, Path, Path]:
    return build_md(), build_wechat(), build_needs()


if __name__ == "__main__":
    md, wx, needs = build()
    print(f"已生成 {md}")
    print(f"已生成 {wx}")
    print(f"已生成 {needs}")
