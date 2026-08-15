#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 8 月 14 日沟通后的对外回复方案。"""

import sys
from pathlib import Path

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
import generate_wangdefeng_proposal as g  # noqa: E402
from generate_external_outputs import set_cell_margins  # noqa: E402

OUT = Path(__file__).resolve().parents[1] / "output"


def build_reply():
    doc = g.new_doc()
    g.setup_section(doc, "10 月 31 日王德峰老师活动 · 合作事项回复（对外）")

    g.add_para(doc, "", size=8, space_after=8)
    g.add_para(
        doc,
        "关于 10 月 31 日活动合作事项的回复",
        size=20,
        bold=True,
        color=g.NAVY,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=6,
    )
    g.add_para(
        doc,
        "供合作方沟通使用  ·  2026 年 8 月 15 日",
        size=11,
        color=g.GOLD,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=10,
    )
    g.add_quote_box(
        doc,
        "回复原则",
        "我方认同双方继续推进联合运营。票价、分成、社群权限应一并确认，不单边调整其中一项。抽成若下调，必须对等让利；官方活动群由我方任群主。",
    )

    g.add_h1(doc, "一、票价：维持 499 元标准票")
    g.add_para(
        doc,
        "感谢贵方提出将对外价格调整为 599 / 999 元。我方理解这是希望体现老师价值，但 10 月 31 日若为约 2.5–3 小时公开讲座，对外传播的主力价格仍建议维持 499 元，不改为 599 元标准票。",
        first_line=True,
    )
    g.add_para(
        doc,
        "599 元可以作为优选 / 前区票，999 元保留为 VIP 票，并配套前排、专属问答、签名或合影等可见权益。299 元引流票、399 元早鸟票建议保留，用于形成开售节奏。不建议把 599 元做成全场基础价，也不建议取消低价票档后只保留 599 / 999。",
        first_line=True,
    )
    g.add_table(
        doc,
        ["票档", "我方确认", "说明"],
        [
            ["限量引流票", "299 元，保留", "后区、限量，制造开售速度"],
            ["早鸟普通票", "399 元，保留", "开售前 7–10 天限量"],
            ["标准票（对外主价格）", "499 元，不改为 599", "核心销售票档"],
            ["优选 / 前区票", "可用 599 或 699 元", "前区座位、赠书或资料"],
            ["VIP 票", "999 元，保留", "须配套前排、问答、签名或合影"],
        ],
        [4.6, 4.8, 7.0],
        emphasize_col=1,
    )

    g.add_h1(doc, "二、分成：可以微调，但必须对等")
    g.add_para(
        doc,
        "我方独立开发的报名，此前书面建议为 60% : 40%。贵方希望下调至 55%。我方可以讨论微调，但目前只调整了我方客户的比例，贵方客户仍按 70% : 30% 结算，属于单边调整，我方不能接受。",
        first_line=True,
    )
    g.add_para(
        doc,
        "原则是：要么两边都不调，要么两边一起调。请贵方在以下两个方案中选择其一。",
        first_line=True,
        bold=True,
    )
    g.add_h2(doc, "方案 A：维持原结构（首选）")
    g.add_table(
        doc,
        ["收入来源", "我方", "贵方"],
        [
            ["我方独立带来的报名 / 企业客户", "60%", "40%"],
            ["贵方独立带来的报名", "30%", "70%"],
            ["双方共同渠道", "50%", "50%"],
            ["我方引入赞助", "70%", "30%"],
            ["贵方引入赞助", "30%", "70%"],
        ],
        [8.4, 4.0, 4.0],
        emphasize_col=1,
    )
    g.add_h2(doc, "方案 B：对等微调")
    g.add_para(
        doc,
        "若我方客户从 60% 下调至 55%，则贵方客户须同步从 70% : 30% 调整为 60% : 40%。这才是互相让利。",
        first_line=True,
    )
    g.add_table(
        doc,
        ["收入来源", "我方", "贵方"],
        [
            ["我方独立带来的报名 / 企业客户", "55%", "45%"],
            ["贵方独立带来的报名", "40%", "60%"],
            ["双方共同渠道", "50%", "50%"],
            ["我方引入赞助", "70%", "30%"],
            ["贵方引入赞助", "30%", "70%"],
        ],
        [8.4, 4.0, 4.0],
        emphasize_col=1,
        header_fill="3D5A40",
    )
    g.add_quote_box(
        doc,
        "不能接受的结构",
        "我方客户 55% : 45%，同时贵方客户仍为 70% : 30%。该结构只改我方贡献、不改贵方贡献，不是共担。",
        fill=g.ROSE_HEX,
        title_color=g.RED,
    )
    g.add_para(
        doc,
        "分成微调的前提是：标准票维持 499 元，官方活动群由我方任群主。没有这两条，分成仍按方案 A 执行，不单独下调。",
        first_line=True,
    )

    g.add_h1(doc, "三、官方社群：我方任群主，贵方任管理员")
    g.add_para(
        doc,
        "本次社群是 10 月 31 日活动交流群，不是贵方其他课程的招生群。报名客户、现场运营和客户服务主要由我方承担，因此官方群须由我方创建并担任群主。贵方担任管理员，负责老师侧内容答疑与必要互动。",
        first_line=True,
    )
    g.add_table(
        doc,
        ["事项", "我方确认"],
        [
            ["群名称", "「王德峰老师 10 月 31 日活动交流群」，不使用「课程群」作为官方群名"],
            ["建群主体", "我方企业微信创建；不使用个人微信或单方账户作为唯一建群主体"],
            ["群主", "我方担任群主，不可转让"],
            ["管理员", "贵方可设管理员，负责内容答疑"],
            ["对外表述", "可写「双方联合运营社群」，但群主权限在我方"],
            ["禁止事项", "不得擅自解散、转让群主、导出名单另建群，或在群内单独销售其他课程"],
            ["原有粉丝群", "贵方原有社群仍由贵方管理；本次新报名客户不迁入他群"],
            ["建群时间", "分成、群规则写入书面确认后再建正式群；开售前如需咨询群，同样由我方任群主"],
        ],
        [3.8, 12.6],
    )

    g.add_h1(doc, "四、文字、头像与 logo：可以改，须保留联合主办")
    g.add_para(
        doc,
        "老师形象、授权范围内的头像、活动文案和视觉规范，我方愿意配合调整。调整后须双方书面确认再对外使用。",
        first_line=True,
    )
    g.add_bullets(
        doc,
        [
            "主视觉、售票页、海报、社群头像须保留联合主办 / 联合运营双方露出。",
            "王德峰老师姓名、肖像、logo 按书面授权使用，双方均有终审权。",
            "活动名称和社群头像不得改成贵方其他课程的附属品牌。",
            "我方负责的场地、票务与运营身份，不在对外物料中消失。",
        ],
    )

    g.add_h1(doc, "五、请贵方确认的四件事")
    g.add_table(
        doc,
        ["序号", "事项", "请确认"],
        [
            ["1", "票价", "标准票 499 元；599 如采用，作为优选票，不替代标准票"],
            ["2", "分成", "选择方案 A（60/40 与 30/70）或方案 B（55/45 与 40/60）"],
            ["3", "社群", "我方企业微信建群并任群主，贵方任管理员"],
            ["4", "视觉", "文案 / 头像 / logo 双方书面确认后对外，保留联合主办"],
        ],
        [2.0, 3.2, 11.2],
    )
    g.add_para(
        doc,
        "以上四点请一并回复。我方希望尽快进入场地、票务和开售准备，避免单项口头承诺、其余事项继续悬空。",
        first_line=True,
    )
    g.add_para(
        doc,
        "（本回复为沟通确认函，不构成合同。最终以双方签署的协议为准。）",
        size=9.5,
        color=g.MUTED,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_before=14,
    )

    doc.add_page_break()
    g.add_h1(doc, "附件：微信沟通稿（可直接转发）")
    g.add_para(
        doc,
        "以下文字可复制到微信，发给对接人或转达大海老师。请整段发送，不要先单独答应抽成下调。",
        size=10.5,
        color=g.MUTED,
        space_after=10,
    )
    wechat = (
        "关于 10 月 31 日活动，我方确认如下：\n\n"
        "1. 价格：对外标准票维持 499 元，不改为 599 主力。599 如需要，作为优选 / 前区票；VIP 仍为 999 元，并配套前排和互动权益。299、399 引流和早鸟票保留。\n\n"
        "2. 分成：可以微调，但必须对等。只把我方客户从 60% 降到 55%、贵方客户仍是 70% : 30%，这个结构我方不能接受。两条路请选一条：\n"
        "方案 A：维持原比例，我方客户 60% : 40%，贵方客户 30% : 70%。\n"
        "方案 B：对等让利，我方客户 55% : 45%，贵方客户同步调到 40% : 60%。\n"
        "没有对等调整，我方客户仍按 60% 执行。\n\n"
        "3. 社群：本次是活动交流群，不是课程招生群。由我方企业微信建群并担任群主，贵方任管理员，负责内容答疑。群主权限不能转让。正式群在书面确认后再建。\n\n"
        "4. 视觉：文字、头像、logo 可以按老师授权调整，但必须保留联合主办身份，双方书面确认后对外。\n\n"
        "以上四点请一并回复。价格不涨、群主在我方，才讨论分成微调。"
    )
    box = doc.add_table(rows=1, cols=1)
    cell = box.cell(0, 0)
    g.shade_cell(cell, g.SOFT_HEX)
    g.set_cell_border(
        cell,
        top={"val": "single", "sz": "8", "color": g.NAVY_HEX, "space": "0"},
        left={"val": "single", "sz": "16", "color": g.NAVY_HEX, "space": "0"},
        bottom={"val": "single", "sz": "8", "color": g.NAVY_HEX, "space": "0"},
        right={"val": "single", "sz": "8", "color": g.NAVY_HEX, "space": "0"},
    )
    cell.text = ""
    set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.35
    run = p.add_run(wechat)
    g.set_run_font(run, size=11, color=g.INK)
    g.set_table_full_width(box, [16.4])

    path = OUT / "王德峰老师10月31日活动_合作事项回复_对外.docx"
    doc.save(path)
    return path


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    path = build_reply()
    print(f"已生成回复方案：{path}")


if __name__ == "__main__":
    main()
