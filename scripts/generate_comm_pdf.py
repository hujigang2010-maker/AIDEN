#!/usr/bin/env python3
"""生成肇事方沟通方案与影像分析 PDF。"""

from pathlib import Path

from reportlab.lib.colors import HexColor, white
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from comm_plan import (
    BANS,
    CN_FONT_PATH,
    COMM_GOAL,
    CT_PLAIN,
    DOC_SUB,
    DOC_TITLE,
    FLOW,
    LAWYER_NOW,
    MED_FOR_TALK,
    MED_SUMMARY,
    OPENING_SCRIPT,
    RIDER_REPLIES,
    SCHEMES,
    SHOW_LIST,
)
from content import (
    ACCIDENT,
    ASR_CORRECTIONS,
    CLOUD_FILM,
    CT_REPORTS,
    DISABILITY,
    HOSPITAL_CHOICE,
    HOSPITAL_TRACK,
    INJURY_NOT_THIS_ACCIDENT,
    INJURY_THIS_ACCIDENT,
    MONEY,
    PARTIES,
    TALKING_POINTS,
)

NAVY = HexColor("#0B2F5B")
RED = HexColor("#A63D2F")
GOLD = HexColor("#C4A35A")
MUTED = HexColor("#5C6B7A")
LIGHT = HexColor("#F3F6FA")
CREAM = HexColor("#F8F1E9")
GREEN = HexColor("#2F6B4F")

MUST_GET = [
    "微信群：骑手 + 美团站点 + 平台保险 + 家属（胡继刚）",
    "保险对接人姓名、电话、保单/工号；是否必须先有认定书才能垫付",
    "市北区人民医院（抚顺路院区）后续费用认不认；齐鲁已发生费用认不认",
    "已发生费用处理路径：谁先垫、转到哪、要哪些发票",
]


def _register_font() -> str:
    kwargs = {"subfontIndex": 0} if CN_FONT_PATH.endswith(".ttc") else {}
    pdfmetrics.registerFont(TTFont("CN", CN_FONT_PATH, **kwargs))
    return "CN"


def _styles(font: str) -> dict[str, ParagraphStyle]:
    return {
        "title": ParagraphStyle(
            "title",
            fontName=font,
            fontSize=16,
            leading=24,
            textColor=NAVY,
            alignment=TA_CENTER,
            spaceAfter=6,
        ),
        "sub": ParagraphStyle(
            "sub",
            fontName=font,
            fontSize=10,
            leading=16,
            textColor=MUTED,
            alignment=TA_CENTER,
            spaceAfter=4,
        ),
        "h1": ParagraphStyle(
            "h1",
            fontName=font,
            fontSize=13,
            leading=20,
            textColor=NAVY,
            spaceBefore=12,
            spaceAfter=6,
        ),
        "h2": ParagraphStyle(
            "h2",
            fontName=font,
            fontSize=11,
            leading=16,
            textColor=HexColor("#1F4E79"),
            spaceBefore=8,
            spaceAfter=4,
        ),
        "body": ParagraphStyle(
            "body",
            fontName=font,
            fontSize=10,
            leading=16,
            textColor=HexColor("#1F2A37"),
            alignment=TA_JUSTIFY,
            firstLineIndent=22,
            spaceAfter=6,
        ),
        "bullet": ParagraphStyle(
            "bullet",
            fontName=font,
            fontSize=10,
            leading=15,
            textColor=HexColor("#1F2A37"),
            leftIndent=14,
            spaceAfter=3,
        ),
        "ban": ParagraphStyle(
            "ban",
            fontName=font,
            fontSize=10,
            leading=15,
            textColor=RED,
            leftIndent=14,
            spaceAfter=3,
        ),
        "cell": ParagraphStyle(
            "cell",
            fontName=font,
            fontSize=8,
            leading=12,
            textColor=HexColor("#1F2A37"),
        ),
        "cell_h": ParagraphStyle(
            "cell_h",
            fontName=font,
            fontSize=8,
            leading=12,
            textColor=white,
            alignment=TA_CENTER,
        ),
        "quote": ParagraphStyle(
            "quote",
            fontName=font,
            fontSize=10,
            leading=16,
            textColor=HexColor("#1F2A37"),
            leftIndent=10,
            rightIndent=10,
            borderPadding=8,
            spaceAfter=8,
            spaceBefore=4,
        ),
        "footer": ParagraphStyle(
            "footer",
            fontName=font,
            fontSize=8,
            leading=12,
            textColor=MUTED,
            alignment=TA_CENTER,
        ),
    }


def P(text: str, style: ParagraphStyle) -> Paragraph:
    return Paragraph(str(text).replace("\n", "<br/>"), style)


def make_table(headers: list[str], rows: list[list[str]], col_widths: list[float], styles: dict) -> Table:
    head = [P(h, styles["cell_h"]) for h in headers]
    data = [head]
    for row in rows:
        data.append([P(c, styles["cell"]) for c in row])
    t = Table(data, colWidths=col_widths, repeatRows=1)
    cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), white),
        ("FONTNAME", (0, 0), (-1, -1), "CN"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("GRID", (0, 0), (-1, -1), 0.4, HexColor("#D0D7DE")),
        ("BACKGROUND", (0, 1), (-1, 1), CREAM),
    ]
    for i in range(2, len(data)):
        if i % 2 == 0:
            cmds.append(("BACKGROUND", (0, i), (-1, i), LIGHT))
    t.setStyle(TableStyle(cmds))
    return t


def _header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(NAVY)
    canvas.rect(0, A4[1] - 12 * mm, A4[0], 12 * mm, fill=1, stroke=0)
    canvas.setFillColor(white)
    canvas.setFont("CN", 8)
    canvas.drawString(18 * mm, A4[1] - 8 * mm, "青岛抚顺路和哈尔滨路路口交通事故 · 肇事方沟通方案 · 内部使用")
    canvas.setFillColor(GOLD)
    canvas.rect(0, 0, A4[0], 10 * mm, fill=1, stroke=0)
    canvas.setFillColor(NAVY)
    canvas.setFont("CN", 8)
    canvas.drawString(18 * mm, 4 * mm, "非律师函 · 非伤残鉴定 · 以齐鲁已审核 CT 和发票为准")
    canvas.drawRightString(A4[0] - 18 * mm, 4 * mm, f"第 {doc.page} 页")
    canvas.restoreState()


def build_comm_pdf(output_path: Path) -> None:
    font = _register_font()
    S = _styles(font)
    story = []

    story.append(Spacer(1, 6))
    story.append(P(DOC_TITLE, S["title"]))
    story.append(P(DOC_SUB, S["sub"]))
    story.append(P("内部使用 · 非律师函 · 非伤残鉴定书 · 依据为齐鲁医院已审核 CT", S["sub"]))
    story.append(Spacer(1, 8))

    story.append(P("一、先把「体检报告」说清楚", S["h1"]))
    for line in MED_SUMMARY:
        story.append(P("• " + line, S["bullet"]))
    story.append(
        P(
            f"医院：{CLOUD_FILM['hospital']}。云影像患者：{CLOUD_FILM['patient_display']}，"
            f"{CLOUD_FILM['sex']}，{CLOUD_FILM['age']} 岁。{CLOUD_FILM['disclaimer']}"
            "报告只能证明伤了什么、做了什么手术，不能当残级，也不能当入职体检结论。",
            S["body"],
        )
    )
    story.append(P(PARTIES["injured"] + " " + PARTIES["family"], S["body"]))
    story.append(P(PARTIES["other"] + " " + PARTIES["police"], S["body"]))
    story.append(P(f"地点：{ACCIDENT['place']}。{ACCIDENT['process']} {ACCIDENT['police_range']}", S["body"]))

    story.append(P("二、三份 CT 对照分析（对外口径）", S["h1"]))
    usable = A4[0] - 36 * mm
    story.append(
        make_table(
            ["检查", "白话结论", "对骑手怎么说", "索赔"],
            [[c["title"], c["plain"], c["to_rider"], c["claim"]] for c in CT_PLAIN],
            [usable * 0.22, usable * 0.32, usable * 0.32, usable * 0.14],
            S,
        )
    )
    story.append(Spacer(1, 6))
    story.append(
        make_table(
            ["日期", "检查号", "项目 / 科室", "影像诊断（报告原文）", "本次"],
            [
                [
                    r["date"],
                    r["no"],
                    f"{r['item']} / {r['place']}",
                    "；".join(r["diagnosis"]),
                    "计入" if r["accident_related"] else "不计入",
                ]
                for r in CT_REPORTS
            ],
            [usable * 0.12, usable * 0.14, usable * 0.28, usable * 0.36, usable * 0.10],
            S,
        )
    )
    story.append(P("本次事故计入的伤", S["h2"]))
    for x in INJURY_THIS_ACCIDENT:
        story.append(P("• " + x, S["bullet"]))
    story.append(P("不计入本次事故", S["h2"]))
    for x in INJURY_NOT_THIS_ACCIDENT:
        story.append(P("• " + x, S["bullet"]))
    story.append(P("通话里必须改口", S["h2"]))
    story.append(
        make_table(
            ["不要再说", "报告实际写的"],
            [[a, b] for a, b in ASR_CORRECTIONS],
            [usable * 0.40, usable * 0.60],
            S,
        )
    )
    story.append(P(DISABILITY["now"] + " " + DISABILITY["cannot_grade"], S["body"]))
    story.append(P(DISABILITY["when"], S["body"]))
    story.append(P(HOSPITAL_TRACK["now"], S["body"]))
    story.append(P(HOSPITAL_TRACK["not_rehab"], S["body"]))
    story.append(P(HOSPITAL_TRACK["two_tracks"], S["body"]))
    story.append(P(MONEY["paid_family_816"] + " 对外只说以发票为准、还在发生。", S["body"]))
    story.append(P(MONEY["insurance"], S["body"]))
    story.append(P(HOSPITAL_CHOICE["headline"], S["body"]))

    story.append(PageBreak())
    story.append(P("三、跟肇事者沟通：这一场到底要什么", S["h1"]))
    for line in COMM_GOAL:
        story.append(P("• " + line, S["bullet"]))
    story.append(P("这一场必须拿到", S["h2"]))
    for x in MUST_GET:
        story.append(P("• " + x, S["bullet"]))
    story.append(P("对骑手的四句态度", S["h2"]))
    for x in TALKING_POINTS["to_rider"]:
        story.append(P("• " + x, S["bullet"]))
    story.append(P("伤情对外只准这样说", S["h2"]))
    for x in MED_FOR_TALK:
        story.append(P("• " + x, S["bullet"]))

    story.append(P("四、建议开场（可照着念）", S["h1"]))
    story.append(P(OPENING_SCRIPT, S["quote"]))
    story.append(
        P(
            "念完停住，等她接。不要自己往下加残级、全责、现金数字。第一次由胡继刚一人出面，岳父场外看稿，律师不到场。",
            S["body"],
        )
    )

    story.append(P("五、对方常见接话，怎么接", S["h1"]))
    story.append(
        make_table(
            ["对方可能说", "建议你怎么接"],
            [[a, b] for a, b in RIDER_REPLIES],
            [usable * 0.32, usable * 0.68],
            S,
        )
    )

    story.append(P("六、这一场绝对不能说", S["h1"]))
    for x in BANS:
        story.append(P("× " + x, S["ban"]))
    for x in TALKING_POINTS["do_not_say"]:
        story.append(P("× " + x, S["ban"]))

    story.append(PageBreak())
    story.append(P("七、要不要请律师", S["h1"]))
    story.append(P(LAWYER_NOW["answer"], S["body"]))
    story.append(P("现在不请诉讼律师的原因", S["h2"]))
    for x in LAWYER_NOW["why_not"]:
        story.append(P("• " + x, S["bullet"]))
    story.append(P(LAWYER_NOW["why_yes_internal"], S["body"]))
    story.append(P("出现下面任一情形，再对外用律师", S["h2"]))
    for x in LAWYER_NOW["triggers"]:
        story.append(P("• " + x, S["bullet"]))
    story.append(
        P(
            "对外用律师时：先发律师函催保险进场和指定医院，再等认定书后起诉。"
            "执行盯美团平台险，不指望骑手个人财产。"
            "不要用「走一般程序、拘留」威胁对方——己方三轮无正规号牌，一般程序可能把无牌上路、甚至无证驾驶的处罚砸回来。",
            S["body"],
        )
    )

    story.append(P("八、三套完整方案", S["h1"]))
    for sch in SCHEMES:
        block = [
            P(sch["name"], S["h2"]),
            P("何时用：" + sch["use"], S["bullet"]),
            P("怎么走：" + sch["steps"], S["bullet"]),
            P("律师：" + sch["lawyer"], S["bullet"]),
            P("风险：" + sch["risk"], S["bullet"]),
            P("不做：" + sch["dont"], S["bullet"]),
        ]
        story.append(KeepTogether(block))

    story.append(P("九、接下来沟通顺序（12 步）", S["h1"]))
    story.append(
        P(
            "按天推进，不要跳步。第 1–7 步现在就走方案甲；卡住再进方案乙或丙。每一步「拿到什么才算过」没拿到，就不要假装过关。",
            S["body"],
        )
    )
    story.append(
        make_table(
            ["步", "何时", "谁 → 找谁", "做什么", "过关", "律师"],
            [
                [
                    str(s["n"]),
                    s["when"],
                    f"{s['who']} → {s['to']}",
                    s["do"],
                    s["get"],
                    s["lawyer"],
                ]
                for s in FLOW
            ],
            [usable * 0.06, usable * 0.12, usable * 0.18, usable * 0.32, usable * 0.16, usable * 0.16],
            S,
        )
    )

    story.append(P("十、现场决策树", S["h1"]))
    for title, text in HOSPITAL_CHOICE["tree"]:
        story.append(P(f"• {title}：{text}", S["bullet"]))
    story.append(
        P(
            "认定书出具后，交警不再参与赔多少钱。已发生医疗费按责任比例报保险；伤残等治疗终结再鉴定。"
            "协商不成只能起诉。评残走司法鉴定、用途写道路交通事故，不要走人社局工伤通道。",
            S["body"],
        )
    )

    story.append(P("十一、这一场可以给对方看的材料", S["h1"]))
    for x in SHOW_LIST:
        story.append(P("• " + x, S["bullet"]))

    story.append(P("十二、使用边界", S["h1"]))
    story.append(
        P(
            "本文是家属内部沟通脚本和影像对照，不是医学鉴定、不是司法鉴定、不是律师函、不是向法院提交的诉状。"
            "对外出示时，以医院已审核报告原文、发票原件、交警认定书为准。"
            "费用两套口述数字冲突，对外只说以发票为准、还在发生。",
            S["body"],
        )
    )
    story.append(Spacer(1, 10))
    story.append(
        P(
            "配套：同名 Word；《肇事方沟通流程表》Excel（12 步可勾选、对方接话、律师触发、出示清单）。",
            S["footer"],
        )
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=16 * mm,
        title=DOC_TITLE,
        author="内部整理",
    )
    doc.build(story, onFirstPage=_header_footer, onLaterPages=_header_footer)
    print("Wrote", output_path)


if __name__ == "__main__":
    out = Path(__file__).resolve().parent.parent / "deliverables" / "青岛抚顺路和哈尔滨路路口交通事故_肇事方沟通方案与体检分析_20260817.pdf"
    build_comm_pdf(out)
