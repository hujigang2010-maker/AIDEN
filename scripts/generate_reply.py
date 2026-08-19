#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成可直接发给合作方的确认函。"""

import sys
from pathlib import Path

from docx.enum.text import WD_ALIGN_PARAGRAPH

sys.path.insert(0, str(Path(__file__).resolve().parent))
import generate_wangdefeng_proposal as g  # noqa: E402

OUT = Path(__file__).resolve().parents[1] / "output"


def build_letter():
    doc = g.new_doc()
    g.setup_section(doc, "10 月 31 日王德峰老师活动 · 合作事项确认函")

    g.add_para(doc, "", size=6, space_after=6)
    g.add_para(
        doc,
        "10 月 31 日王德峰老师活动",
        size=14,
        bold=True,
        color=g.GOLD,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=4,
    )
    g.add_para(
        doc,
        "合作事项确认函",
        size=22,
        bold=True,
        color=g.NAVY,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=8,
    )
    g.add_para(
        doc,
        "致：合作方　　日期：2026 年 8 月 19 日",
        size=11,
        color=g.MUTED,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=12,
    )
    g.add_para(
        doc,
        "感谢贵方就 10 月 31 日活动持续沟通。现就票价、分成、官方社群及视觉物料四项，书面确认我方意见。请贵方对下列事项一并回复后，双方再进入场地、票务与开售准备。",
        first_line=True,
    )
    g.add_quote_box(
        doc,
        "确认原则",
        "双方继续按联合运营推进。票价、分成、社群权限应一并确认，不单边调整其中一项。分成若下调，须对等让利；官方活动群由我方任群主。",
    )

    g.add_h1(doc, "一、票价")
    g.add_para(
        doc,
        "感谢贵方提出将对外价格调整为 599 / 999 元。我方理解这是希望体现老师价值。若本次为约 2.5–3 小时公开讲座，对外传播的主力价格仍维持 499 元，不改为 599 元标准票。",
        first_line=True,
    )
    g.add_para(
        doc,
        "599 元可作为优选 / 前区票；999 元保留为 VIP 票，并配套前排、专属问答、签名或合影等可见权益。299 元引流票、399 元早鸟票建议保留，以形成开售节奏。不建议取消低价票档后只保留 599 / 999。",
        first_line=True,
    )
    g.add_table(
        doc,
        ["票档", "我方确认", "说明"],
        [
            ["限量引流票", "299 元，保留", "后区、限量，用于开售节奏"],
            ["早鸟普通票", "399 元，保留", "开售前 7–10 天限量"],
            ["标准票（对外主价格）", "499 元", "不改为 599 元主力"],
            ["优选 / 前区票", "599 元或 699 元", "前区座位、赠书或资料"],
            ["VIP 票", "999 元，保留", "须配套前排、问答、签名或合影"],
        ],
        [4.6, 4.8, 7.0],
        emphasize_col=1,
    )

    g.add_h1(doc, "二、分成")
    g.add_para(
        doc,
        "我方独立开发的报名，此前书面建议为 60% : 40%。贵方希望下调至 55%。我方可以讨论微调，但目前仅调整了我方客户比例，贵方客户仍按 70% : 30% 结算，属于单边调整，我方不能接受。",
        first_line=True,
    )
    g.add_para(
        doc,
        "原则：要么两边都不调，要么两边一起调。请贵方在下列两个方案中选择其一。",
        first_line=True,
        bold=True,
    )
    g.add_h2(doc, "方案 A：维持原结构")
    g.add_table(
        doc,
        ["收入来源", "我方", "贵方"],
        [
            ["我方独立带来的报名 / 企业客户", "60%", "40%"],
            ["贵方独立带来的报名", "30%", "70%"],
            ["双方共同渠道", "50%", "50%"],
            ["我方引入的赞助", "70%", "30%"],
            ["贵方引入的赞助", "30%", "70%"],
        ],
        [8.4, 4.0, 4.0],
        emphasize_col=1,
    )
    g.add_h2(doc, "方案 B：对等微调")
    g.add_para(
        doc,
        "若我方客户从 60% 下调至 55%，则贵方客户须同步从 70% : 30% 调整为 60% : 40%。",
        first_line=True,
    )
    g.add_table(
        doc,
        ["收入来源", "我方", "贵方"],
        [
            ["我方独立带来的报名 / 企业客户", "55%", "45%"],
            ["贵方独立带来的报名", "40%", "60%"],
            ["双方共同渠道", "50%", "50%"],
            ["我方引入的赞助", "70%", "30%"],
            ["贵方引入的赞助", "30%", "70%"],
        ],
        [8.4, 4.0, 4.0],
        emphasize_col=1,
        header_fill="3D5A40",
    )
    g.add_quote_box(
        doc,
        "我方不能接受的结构",
        "我方客户 55% : 45%，同时贵方客户仍为 70% : 30%。该结构只调整我方贡献，未相应调整贵方贡献。",
        fill=g.ROSE_HEX,
        title_color=g.RED,
    )
    g.add_para(
        doc,
        "分成微调的前提是：标准票维持 499 元，官方活动群由我方任群主。若上述两项不能同时确认，分成按方案 A 执行。",
        first_line=True,
    )

    g.add_h1(doc, "三、官方社群")
    g.add_para(
        doc,
        "本次社群为 10 月 31 日活动交流群，不是其他课程的招生群。报名客户、现场运营和客户服务主要由我方承担，因此官方群须由我方创建并担任群主。贵方担任管理员，负责老师侧内容答疑与必要互动。",
        first_line=True,
    )
    g.add_table(
        doc,
        ["事项", "我方确认"],
        [
            ["群名称", "「王德峰老师 10 月 31 日活动交流群」"],
            ["建群主体", "我方企业微信创建"],
            ["群主", "我方担任，权限不转让"],
            ["管理员", "贵方担任，负责内容答疑"],
            ["对外表述", "可写「双方联合运营社群」，群主权限在我方"],
            ["使用边界", "不得擅自解散、转让群主、导出名单另建群，或在群内单独销售其他课程"],
            ["原有社群", "贵方原有粉丝群仍由贵方管理；本次新报名客户不迁入他群"],
            ["建群时间", "本函事项书面确认后再建正式群；如需开售咨询群，同样由我方任群主"],
        ],
        [3.8, 12.6],
    )

    g.add_h1(doc, "四、文字、头像与 logo")
    g.add_para(
        doc,
        "老师形象、授权范围内的头像、活动文案和视觉规范，我方愿意配合调整。调整后须双方书面确认再对外使用。",
        first_line=True,
    )
    g.add_bullets(
        doc,
        [
            "主视觉、售票页、海报、社群头像须保留联合主办 / 联合运营双方露出。",
            "王德峰老师姓名、肖像及相关标识按书面授权使用，双方均有终审权。",
            "活动名称和社群头像不改为其他课程的附属品牌。",
        ],
    )

    g.add_h1(doc, "五、请一并确认")
    g.add_table(
        doc,
        ["序号", "事项", "请贵方确认"],
        [
            ["1", "票价", "标准票 499 元；599 元如采用，作为优选票，不替代标准票"],
            ["2", "分成", "选择方案 A，或选择方案 B（两边同步调整）"],
            ["3", "社群", "我方企业微信建群并任群主，贵方任管理员"],
            ["4", "视觉", "文案、头像、logo 双方书面确认后对外，保留联合主办"],
        ],
        [2.0, 3.0, 11.4],
    )
    g.add_para(
        doc,
        "请贵方就以上四点一并书面回复。我方愿尽快与贵方共同推进本次活动。",
        first_line=True,
        space_before=8,
    )
    g.add_para(
        doc,
        "此致",
        first_line=True,
        space_before=10,
        space_after=2,
    )
    g.add_para(
        doc,
        "敬礼",
        first_line=True,
        space_after=16,
    )
    g.add_para(
        doc,
        "我方（联合运营方）",
        size=11,
        bold=True,
        color=g.NAVY,
        align=WD_ALIGN_PARAGRAPH.RIGHT,
        space_after=2,
    )
    g.add_para(
        doc,
        "2026 年 8 月 19 日",
        size=11,
        color=g.MUTED,
        align=WD_ALIGN_PARAGRAPH.RIGHT,
        space_after=12,
    )
    g.add_para(
        doc,
        "本函为沟通确认文件，不构成合同。最终合作内容以双方签署的协议为准。",
        size=9.5,
        color=g.MUTED,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_before=8,
    )

    path = OUT / "10月31日王德峰老师活动_合作事项确认函.docx"
    doc.save(path)
    return path


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    path = build_letter()
    print(f"已生成确认函：{path}")


if __name__ == "__main__":
    main()
