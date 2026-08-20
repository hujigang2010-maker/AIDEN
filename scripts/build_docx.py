# -*- coding: utf-8 -*-
"""生成 Word：AIDEN × 西电产学研课题介绍（完整版）。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

import content as C

OUT = Path(__file__).resolve().parents[1] / "exports"
OUT.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT / "AIDEN_西电产学研课题介绍.docx"

NAVY = (0x0B, 0x3D, 0x5C)
TEAL = (0x0F, 0x7A, 0x6E)
INK = (0x1A, 0x2B, 0x2E)
GREY = (0x5C, 0x6B, 0x70)


def set_run_font(run, size=11, bold=False, color=None, name="微软雅黑"):
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


def add_p(doc, text, bold=False, size=11, color=None, align=None, space_after=8, first_line=False):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    if first_line:
        pf.first_line_indent = Cm(0.74)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    return p


def add_h(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        set_run_font(run, size={1: 16, 2: 13, 3: 12}.get(level, 12), bold=True, color=NAVY)
    return h


def add_bullets(doc, items):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        run = p.add_run(it)
        set_run_font(run, size=11, color=INK)
        p.paragraph_format.space_after = Pt(3)


def add_table(doc, headers, rows, header_fill="0B3D5C"):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        r = p.add_run(h)
        set_run_font(r, size=10, bold=True, color=(255, 255, 255))
        set_cell_bg(cell, header_fill)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = table.rows[ri + 1].cells[ci]
            cell.text = ""
            p = cell.paragraphs[0]
            r = p.add_run(str(val))
            set_run_font(r, size=9.5, color=INK)
            if ri % 2 == 1:
                set_cell_bg(cell, "D7EBE6")
    doc.add_paragraph()
    return table


def build():
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "微软雅黑"
    style.font.size = Pt(11)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    for section in doc.sections:
        section.top_margin = Cm(2.2)
        section.bottom_margin = Cm(2.2)
        section.left_margin = Cm(2.4)
        section.right_margin = Cm(2.4)

    add_p(doc, f"{C.SCHOOL} 课题发布配套材料", bold=True, size=12, color=TEAL,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    add_p(doc, C.TITLE, bold=True, size=22, color=NAVY,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    add_p(doc, C.SUBTITLE, size=13, color=TEAL,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    add_p(doc, C.ORG_LINE, size=11, color=GREY,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_p(doc, f"对接导师：{C.MENTOR}  ·  {C.MENTOR_TITLE}", size=11, color=GREY,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_p(doc, f"{C.VERSION}  ·  {C.DATE_STR}  ·  计划 {C.RELEASE_DATE} 正式发布",
          size=11, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)

    add_h(doc, "一、写在前面", 1)
    for para in C.LETTER:
        add_p(doc, para, first_line=True)
    add_p(doc, "本文是完整课题说明。若只需课堂上讲 15 分钟，请用配套 PPT；若要收志愿，请用配套 Excel；若要发学生群，请用发群一页稿（Markdown）。")

    add_h(doc, "二、我们是谁，课题从哪来", 1)
    for para in C.ABOUT:
        add_p(doc, para, first_line=True)
    add_p(doc, "这些题目已经在 AIDEN 仓库里留下可打开的产物（网页、小程序演示、规格、白皮书、数据看板）。同学进组后是接着做，不是从零虚构一个「实习项目」。")

    add_h(doc, "三、原则", 1)
    add_bullets(doc, C.PRINCIPLES)

    add_h(doc, "四、三条赛道", 1)
    for tr in C.TRACKS:
        add_h(doc, tr["name"], 2)
        add_p(doc, tr["blurb"])
        names = [f"{t['id']}  {t['name']}" for t in C.TOPICS if t["id"] in tr["ids"]]
        add_bullets(doc, names)

    add_h(doc, "五、十条课题总览", 1)
    add_table(
        doc,
        ["编号", "课题", "赛道", "状态", "人数", "周期"],
        [[t["id"], t["name"], t["track"], t["status"], t["quota"], t["weeks"]] for t in C.TOPICS],
    )

    add_h(doc, "六、分题说明（老师可直接剪贴发群）", 1)
    add_p(doc, "每一题都按同一结构写：一句话、已经做到哪、本组目标、任务、角色、交付物。技能写在总览匹配表里，避免每题重复。")

    for t in C.TOPICS:
        add_h(doc, f"{t['id']}  {t['name']}", 2)
        add_table(
            doc,
            ["项", "内容"],
            [
                ["赛道 / 状态", f"{t['track']}  ·  {t['status']}"],
                ["人数 / 周期", f"{t['quota']}  ·  {t['weeks']}"],
                ["适合专业", t["major"]],
                ["关键技能", t["skills"]],
                ["角色建议", t["roles"]],
            ],
            header_fill="1A7A6D",
        )
        add_h(doc, "一句话", 3)
        add_p(doc, t["one_liner"])
        add_h(doc, "已经开工的部分", 3)
        add_p(doc, t["started"])
        add_h(doc, "本组目标", 3)
        add_p(doc, t["goal"])
        add_h(doc, "建议任务", 3)
        add_bullets(doc, t["tasks"])
        add_h(doc, "结题交付", 3)
        add_p(doc, t["deliverables"])

    add_h(doc, "七、选题与组队", 1)
    add_bullets(doc, C.HOW_TO_JOIN)
    add_p(doc, "热门题超额：先看备选，再看技能匹配，最后抽签。无人认领的题目由老师决定合并到相邻赛道或本学期缓开。")

    add_h(doc, "八、时间表", 1)
    add_table(doc, ["节点", "做什么"], C.RHYTHM)

    add_h(doc, "九、收获与约定", 1)
    add_h(doc, "同学侧收获", 2)
    add_bullets(doc, C.WHAT_YOU_GET)
    add_h(doc, "项目组约定", 2)
    add_bullets(doc, C.WHAT_WE_EXPECT)

    add_h(doc, "十、结题包", 1)
    add_table(
        doc,
        ["组别", "除共同必交外，还要交"],
        [
            ["产品组（T01–T03）", "可运行地址或录屏、模块说明、未完成清单"],
            ["协议/工程组（T04–T06、T09）", "原型或脚本、验收记录、失败用例"],
            ["研究组（T07、T08、T10）", "修订对照、证据袋、一页摘要"],
            ["共同必交", "周报合订、复盘一页、成员分工表"],
        ],
    )

    add_h(doc, "十一、发布时怎么用这套材料", 1)
    add_bullets(doc, [
        "学生群：粘贴发群一页稿（Markdown / 纯文本），附 Excel 选题表。",
        "宣讲会：用 PPT，控制在 15 分钟，细节让同学回去看 Word。",
        "老师留底：Word 全文 + Excel 编组。",
        "学生群：粘贴发群一页稿（Markdown / 纯文本）。",
        C.WECHAT_CONTACT_NOTE,
        "AIDEN 侧不在学生群单独招生。编组由西电老师确认后，再拉各项目组工作群。",
    ])

    add_h(doc, "十二、填报字段（与 Excel 一致）", 1)
    add_table(
        doc,
        ["字段", "怎么填"],
        [
            ["姓名 / 学院专业 / 年级", "与学籍一致"],
            ["第一志愿编号", "T01–T10 选一"],
            ["备选编号", "必须与第一志愿不同"],
            ["我能贡献什么（≤80 字）", "写技能或已有作品"],
            ["每周可投入小时", "建议不少于 8，写真实数字"],
            ["能否短期赴沪", "能 / 不能 / 寒暑假可以"],
            ["角色意向", "组长 / 开发 / 主笔 / 数据 / 演示"],
        ],
    )

    add_p(doc, f"（完）  {C.ORG_LINE}  |  {C.DATE_STR}", size=10, color=GREY,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0)

    doc.save(OUT_FILE)
    print(f"已生成 {OUT_FILE}")
    return OUT_FILE


if __name__ == "__main__":
    build()
