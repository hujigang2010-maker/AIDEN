# -*- coding: utf-8 -*-
"""生成 Word：会议纪要 + 文字版可执行落地方案。"""
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

import content as C

OUT = Path(__file__).resolve().parent.parent / "deliverables"
OUT.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT / "上海AI博物馆_会议纪要与可执行方案.docx"

NAVY = (0x0B, 0x3D, 0x5C)
TEAL = (0x0F, 0x7A, 0x6E)
INK = (0x1A, 0x2B, 0x2E)
GREY = (0x5C, 0x6B, 0x70)


def set_run_font(run, size=10.5, bold=False, color=None, name="微软雅黑"):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)


def set_cell_bg(cell, color_hex: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)


def add_p(doc, text, bold=False, size=10.5, color=None, align=None, space_after=6):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    return p


def add_h(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        set_run_font(
            run,
            size={0: 22, 1: 16, 2: 13, 3: 11}.get(level, 11),
            bold=True,
            color=NAVY,
        )
    return h


def add_bullets(doc, items):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        run = p.add_run(it)
        set_run_font(run, size=10.5, color=INK)
        p.paragraph_format.space_after = Pt(3)


def add_table(doc, headers, rows, header_fill="0B3D5C"):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        r = p.add_run(h)
        set_run_font(r, size=9.5, bold=True, color=(255, 255, 255))
        set_cell_bg(cell, header_fill)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = table.rows[ri + 1].cells[ci]
            cell.text = ""
            p = cell.paragraphs[0]
            r = p.add_run(str(val))
            set_run_font(r, size=9, color=INK)
            if ri % 2 == 1:
                set_cell_bg(cell, "D7EBE6")
    doc.add_paragraph()
    return table


def build():
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "微软雅黑"
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.3)
        section.right_margin = Cm(2.3)

    # ===== 封面 =====
    for _ in range(2):
        add_p(doc, "")
    add_p(doc, "工作会议材料", bold=True, size=12, color=TEAL,
          align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p(doc, C.PROJECT, bold=True, size=22, color=NAVY,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    add_p(doc, "会议纪要与可执行落地方案", bold=True, size=16, color=NAVY,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    add_p(
        doc,
        "核心聚焦：投资方怎么投 · 项目怎么赚 · 政策支持怎么落地",
        bold=True, size=12, color=TEAL,
        align=WD_ALIGN_PARAGRAPH.CENTER, space_after=20,
    )
    add_p(doc, f"版本：{C.VERSION}　　日期：{C.DATE_STR}", size=11,
          color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p(doc, "资料来源：得到大脑分享笔记（两场）", size=10, color=GREY,
          align=WD_ALIGN_PARAGRAPH.CENTER)
    for u in C.SOURCE_NOTES:
        add_p(doc, u, size=9, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    doc.add_page_break()

    # ===== 目录 =====
    add_h(doc, "目录", level=1)
    toc = [
        "第一部分　双场会议纪要",
        "　　一、会议概要",
        "　　二、分议题结论",
        "　　三、待办事项",
        "第二部分　可执行落地方案（文字版）",
        "　　一、项目定位与差异化",
        "　　二、投资方如何投资（重点）",
        "　　三、盈利模式如何赚钱（重点）",
        "　　四、政策性支持与扶持基金（重点）",
        "　　五、场地与空间经济",
        "　　六、内容结构与资源",
        "　　七、组织分工与推进节奏",
        "　　八、风险与下一步",
    ]
    for t in toc:
        add_p(doc, t, size=11, space_after=4)
    doc.add_page_break()

    # ===== 第一部分 纪要 =====
    add_h(doc, "第一部分　双场会议纪要", level=1)
    add_h(doc, "一、会议概要", level=2)
    add_p(doc, C.MEETING_SUMMARY)
    add_p(doc, "会议元信息：", bold=True)
    add_table(
        doc,
        ["场次", "时间", "时长", "人数", "类型"],
        [[m["场次"], m["时间"], m["时长"], m["人数"], m["类型"]] for m in C.MEETING_META],
    )
    add_p(doc, "本方案提炼时特别关注的三点：", bold=True)
    add_bullets(doc, C.CORE_CONCERNS)

    add_h(doc, "二、分议题结论", level=2)
    for kp in C.MEETING_KEY_POINTS:
        add_p(doc, kp["议题"], bold=True, color=TEAL, space_after=2)
        add_p(doc, kp["结论"], space_after=8)

    add_h(doc, "三、待办事项", level=2)
    add_table(
        doc,
        ["责任人", "事项", "优先级", "时间要求"],
        C.TODOS,
    )
    add_p(
        doc,
        "说明：场次一后半段为地铁出行与私人事务，与项目无关内容未写入本纪要。",
        size=9, color=GREY,
    )
    doc.add_page_break()

    # ===== 第二部分 方案 =====
    add_h(doc, "第二部分　可执行落地方案（文字版）", level=1)
    add_p(
        doc,
        "本部分把会议共识整理为可路演、可谈判、可排期的执行文本。"
        "阅读顺序建议：先看投资与盈利、再看政策匹配，最后落到场地谈判与90天节奏。",
    )

    add_h(doc, "一、项目定位与差异化", level=2)
    add_p(
        doc,
        "定位：建设人工智能产业相关的常设展示与配套服务空间，以可签约企业展位、"
        "可核验应用案例和研学服务为主，并探索与城市重大展会的会后延伸合作。"
        "一期控制可运营面积与预算，验证后再评估扩展或对外输出。",
    )
    add_p(doc, "差异化要点：", bold=True)
    add_bullets(doc, C.DIFFERENTIATION)

    add_h(doc, "二、投资方如何投资（重点）", level=2)
    add_p(doc, C.INVEST_PRINCIPLE)
    add_p(
        doc,
        "实务上，把「投资」拆成可组合路径：物业租金与装修共担、主赞助、企业展位合作、"
        "政府专项、社团合作、运营方启动金与后期跟投、授权类内容合作。"
        "P0优先闭环：物业条款、主赞助、政府专项申报与启动金——否则不宜开工。",
    )
    add_table(
        doc,
        ["路径", "出资方式", "出资方画像", "回报机制", "优先级"],
        [
            [it["路径"], it["出资方式"], it["出资方画像"], it["回报机制"], it["优先级"]]
            for it in C.INVEST_PATHS
        ],
    )
    add_p(doc, "赞助层级（便于企业立项）：", bold=True)
    add_table(doc, C.SPONSOR_TIERS[0], C.SPONSOR_TIERS[1:])
    add_p(doc, "关键落地动作：", bold=True)
    add_bullets(doc, [
        "出具《物业合作需求书》（租期、租金阶梯/装修期减免、装修分摊、退出条款）。",
        "主赞助：一页纸权益包 + 对齐对方科创合作KPI的立项叙述。",
        "企业展位：建设费+年度更新合同，约定露出、撤展与合作期限。",
        "科协等社团：合作备忘录/正式文件，明确挂牌与可申报事项。",
        "政府：争取纳入区相关项目库，匹配租金支持与专项申报窗口；专项按保守到账入账。",
    ])

    add_h(doc, "三、盈利模式如何赚钱（重点）", level=2)
    add_p(doc, C.PROFIT_LOGIC)
    add_p(
        doc,
        "开办期必须单列启动金与运营储备，不可假设项目公司零出资。"
        "开业后利润不依赖门票；启动顺序建议："
        "配套空间经营 → 研学服务 → 展位与赞助 → 入驻服务；咨询输出与跟投放后期。",
    )
    add_table(
        doc,
        ["收入线", "描述", "稳态占比", "里程碑"],
        [
            [r["收入线"], r["描述"], r["目标占比(稳态)"], r["里程碑"]]
            for r in C.REVENUE_STREAMS
        ],
    )
    add_p(doc, "空间与经济假设（讨论稿，待测算锁定）：", bold=True)
    add_table(doc, C.UNIT_ECONOMICS[0], C.UNIT_ECONOMICS[1:])

    add_h(doc, "四、政策性支持与扶持基金（重点）", level=2)
    add_p(
        doc,
        "政策需转化成合同与批复，而不是口号。优先级："
        "①研学/科普基地与课程采购；"
        "②物业租金优惠或装修支持（写入租约）；"
        "③科创/科普/文旅专项按窗口申报，预算只计保守到账比例。",
    )
    add_table(
        doc,
        ["政策/抓手", "要点", "对项目价值", "落地步骤"],
        [
            [x["政策/抓手"], x["要点"], x["对项目价值"], x["落地步骤"]]
            for x in C.POLICY_SUPPORT
        ],
    )
    add_p(doc, "扶持资金类型与用途匹配：", bold=True)
    add_table(doc, C.POLICY_FUND_MATCH[0], C.POLICY_FUND_MATCH[1:])
    add_p(
        doc,
        "执行口诀：先谈清租约与装修分摊 → 锁定启动金与主赞助 → 再按窗口申报专项；"
        "未批复资金不计入必达开办条件。",
        bold=True, color=TEAL,
    )

    add_h(doc, "五、场地与空间经济", level=2)
    add_p(
        doc,
        "一期建议控制在可运营规模（如创智汇约3300㎡顶楼试点，或1500–3000㎡展示区）。"
        "租金、装修与设备采用共担与分期，不以超长期零租金或场地方大额无偿出资作为前提。",
    )
    add_table(doc, C.SITE_CANDIDATES[0], C.SITE_CANDIDATES[1:])

    add_h(doc, "六、内容结构与资源", level=2)
    add_table(doc, C.CONTENT_LAYERS[0], C.CONTENT_LAYERS[1:])
    add_bullets(doc, [
        "内容以可签约展位与可授权案例为主，不堆砌概念演示，不预设不可控传播事项。",
        "陈胜补充海外授权线索；国内头部企业多数在沪，适合集中拜访谈展位。",
        "涉外与品牌使用先法务预审，再制作物料。",
    ])

    add_h(doc, "七、组织分工与推进节奏", level=2)
    add_table(doc, C.ORG_ROLES[0], C.ORG_ROLES[1:])
    add_p(doc, "阶段节奏：", bold=True)
    add_table(doc, C.ROADMAP[0], C.ROADMAP[1:])
    add_p(
        doc,
        "方法论约束：先完成预算框架与合作条款清单，再进入设计深化；"
        "开办资金未闭环前，不启动大额装修合同。",
    )

    add_h(doc, "八、风险与下一步", level=2)
    add_table(doc, C.RISKS[0], C.RISKS[1:])
    add_p(doc, "建议立即执行的下一步：", bold=True)
    add_bullets(doc, [
        "陈胜发送极简文字材料，并补充海外授权可行性清单。",
        "胡继刚以本套 Word / PPT / Excel 形成可外发路演包。",
        "陈红苗并行推进：创智汇与复兴岛租约比选、潜在赞助沟通、教育渠道对接。",
        "指定专人梳理扶持基金窗口表，专项按保守到账写入预算。",
    ])

    add_p(doc, "")
    add_p(
        doc,
        "— 文档结束 —",
        align=WD_ALIGN_PARAGRAPH.CENTER, color=GREY, size=10,
    )
    add_p(
        doc,
        "配套文件：上海AI博物馆_可执行落地方案.pptx　｜　上海AI博物馆_投资盈利政策落地表.xlsx",
        align=WD_ALIGN_PARAGRAPH.CENTER, color=GREY, size=9,
    )

    doc.save(OUT_FILE)
    print(f"已生成: {OUT_FILE}")


if __name__ == "__main__":
    build()
