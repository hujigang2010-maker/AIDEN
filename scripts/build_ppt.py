# -*- coding: utf-8 -*-
"""生成 PPT：AIDEN × 西电产学研课题介绍。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

import content as C
from ppt_utils import (
    ACCENT,
    BG_DEEP,
    DARK,
    GOLD,
    GREY,
    LIGHT,
    ORANGE,
    PRIMARY,
    SOFT,
    WHITE,
    add_rect,
    add_rounded,
    add_text,
    add_bullet_list,
    make_card,
    make_table,
    slide_header,
)

OUT = Path(__file__).resolve().parents[1] / "exports"
OUT.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT / "AIDEN_西电产学研课题介绍.pptx"
TOTAL = 16


def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    page = [0]

    def new_slide():
        page[0] += 1
        return prs.slides.add_slide(blank), page[0]

    # 1 封面
    s, _ = new_slide()
    add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), fill=BG_DEEP)
    add_rect(s, Inches(0), Inches(0), Inches(0.18), Inches(7.5), fill=ACCENT)
    add_rect(s, Inches(0), Inches(5.15), Inches(13.333), Inches(0.04), fill=GOLD)
    add_text(s, Inches(0.9), Inches(1.15), Inches(11.5), Inches(0.4),
             f"{C.SCHOOL}  ·  课题发布配套材料  ·  {C.VERSION}", size=16, color=SOFT)
    add_text(s, Inches(0.9), Inches(1.7), Inches(11.8), Inches(1.1),
             C.TITLE, size=36, bold=True, color=WHITE)
    add_text(s, Inches(0.9), Inches(2.9), Inches(11.5), Inches(0.5),
             C.SUBTITLE, size=20, color=SOFT)
    add_text(s, Inches(0.9), Inches(3.6), Inches(11.5), Inches(0.7),
             [
                 "发给经过初筛的同学：按兴趣选题，按组跟进。",
                 f"计划 {C.RELEASE_DATE} 正式发出课题需求。",
             ],
             size=16, color=WHITE)
    add_text(s, Inches(0.9), Inches(5.45), Inches(11.5), Inches(0.35),
             C.ORG_LINE, size=14, color=GOLD)
    add_text(s, Inches(0.9), Inches(6.05), Inches(11.5), Inches(0.55),
             [f"对接导师：{C.MENTOR}  ·  {C.MENTOR_TITLE}", f"讨论稿  |  {C.DATE_STR}"],
             size=13, color=WHITE)

    # 2 这封材料做什么
    s, p = new_slide()
    slide_header(s, "这份材料给谁用", "老师发群、同学选题、组长认领，同一套口径", p, TOTAL)
    add_bullet_list(s, Inches(0.55), Inches(1.45), Inches(12.2), Inches(1.5), C.LETTER, size=15, color=DARK, bullet="▸")
    make_card(s, Inches(0.5), Inches(3.15), Inches(4.0), Inches(3.7),
              "老师侧", ["下周一把需求发出去", "用选题表收志愿、编项目组", "PPT 可用于宣讲，Word 可细读"],
              accent=PRIMARY, body_size=13)
    make_card(s, Inches(4.65), Inches(3.15), Inches(4.0), Inches(3.7),
              "同学侧", ["十条都是已开工题目", "看赛道、人数、技能再报名", "进组后接现有产物，不从零空想"],
              accent=ACCENT, body_size=13)
    make_card(s, Inches(8.8), Inches(3.15), Inches(4.0), Inches(3.7),
              "我们侧", ["只提供真实在研课题", "周会验收，不代写", "结业材料随成果走，不开空证明"],
              accent=ORANGE, body_size=13)

    # 3 我们是谁
    s, p = new_slide()
    slide_header(s, "课题从哪来", "上海产业场景 × 已经在转的 AIDEN 工作区", p, TOTAL)
    add_bullet_list(s, Inches(0.55), Inches(1.45), Inches(12.2), Inches(1.6), C.ABOUT, size=15, color=DARK, bullet="▸")
    headers = ["类型", "例子（均已有产物）", "学生能接什么"]
    rows = [
        ["可运行产品", "MP Typer 排版器、杨浦楼宇看板、复旦链接演示", "功能迭代、试用、文档"],
        ["协议与工程", "活动报名 MCP 规格、TicNote API、文档生成校验", "最小原型、评测、规范"],
        ["研究报告", "具身智能白皮书、产业空间、住房/能源/出海", "证据补强、图表、更新摘要"],
        ["工作方法", "AIDEN 分支交付、90+ 真实任务在跑", "手册、复盘、周报与验收口"],
    ]
    make_table(s, Inches(0.5), Inches(3.2), Inches(12.3), Inches(3.5), headers, rows, font_size=12)

    # 4 原则
    s, p = new_slide()
    slide_header(s, "六条原则", "先把规矩讲清，再让同学挑题目", p, TOTAL)
    for i, t in enumerate(C.PRINCIPLES):
        col = i % 2
        row = i // 2
        x = 0.5 + col * 6.4
        y = 1.5 + row * 1.75
        add_rounded(s, Inches(x), Inches(y), Inches(6.15), Inches(1.55), fill=LIGHT)
        add_rect(s, Inches(x), Inches(y), Inches(0.1), Inches(1.55), fill=ACCENT if i % 2 == 0 else PRIMARY)
        add_text(s, Inches(x + 0.28), Inches(y + 0.18), Inches(0.7), Inches(0.4),
                 f"0{i+1}", size=18, bold=True, color=ACCENT)
        add_text(s, Inches(x + 0.28), Inches(y + 0.62), Inches(5.6), Inches(0.75),
                 t, size=14, color=DARK)

    # 5 总览表
    s, p = new_slide()
    slide_header(s, "十条课题总览", "每题都已开工；人数按组，不按「来者不拒」", p, TOTAL)
    headers = ["编号", "课题", "赛道", "状态", "人数"]
    rows = [[t["id"], t["name"], t["track"], t["status"], t["quota"]] for t in C.TOPICS]
    make_table(s, Inches(0.35), Inches(1.4), Inches(12.6), Inches(5.55), headers, rows, font_size=11)

    # 6 三条赛道
    s, p = new_slide()
    slide_header(s, "先选赛道，再选题", "兴趣对不上编号时，按赛道找相邻题目", p, TOTAL)
    colors = [PRIMARY, ACCENT, ORANGE]
    for i, tr in enumerate(C.TRACKS):
        x = 0.45 + i * 4.25
        add_rounded(s, Inches(x), Inches(1.5), Inches(4.05), Inches(5.35), fill=LIGHT)
        add_rect(s, Inches(x), Inches(1.5), Inches(4.05), Inches(0.12), fill=colors[i])
        add_text(s, Inches(x + 0.2), Inches(1.8), Inches(3.65), Inches(0.7),
                 tr["name"], size=16, bold=True, color=colors[i])
        add_text(s, Inches(x + 0.2), Inches(2.5), Inches(3.65), Inches(1.1),
                 tr["blurb"], size=13, color=DARK)
        names = [t["name"].split("：")[0] for t in C.TOPICS if t["id"] in tr["ids"]]
        add_bullet_list(s, Inches(x + 0.15), Inches(3.7), Inches(3.7), Inches(2.8),
                        names, size=13, color=DARK)

    # 7 课题 1-5
    s, p = new_slide()
    slide_header(s, "课题卡片（一）T01–T05", "一句话 + 已开工状态，细节见 Word", p, TOTAL)
    for i, t in enumerate(C.TOPICS[:5]):
        y = 1.38 + i * 1.12
        add_rounded(s, Inches(0.4), Inches(y), Inches(12.5), Inches(1.04), fill=LIGHT)
        add_rect(s, Inches(0.4), Inches(y), Inches(1.15), Inches(1.04), fill=PRIMARY)
        add_text(s, Inches(0.45), Inches(y), Inches(1.05), Inches(1.04),
                 t["id"], size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(1.7), Inches(y + 0.06), Inches(10.9), Inches(0.36),
                 f"{t['name']}    ·    {t['quota']}    ·    {t['status']}",
                 size=13, bold=True, color=PRIMARY)
        add_text(s, Inches(1.7), Inches(y + 0.46), Inches(10.9), Inches(0.5),
                 t["one_liner"], size=12, color=DARK)

    # 8 课题 6-10
    s, p = new_slide()
    slide_header(s, "课题卡片（二）T06–T10", "T09 即当前主线工作流，欢迎直接认领", p, TOTAL)
    for i, t in enumerate(C.TOPICS[5:]):
        y = 1.38 + i * 1.12
        add_rounded(s, Inches(0.4), Inches(y), Inches(12.5), Inches(1.04), fill=LIGHT)
        add_rect(s, Inches(0.4), Inches(y), Inches(1.15), Inches(1.04), fill=ACCENT if t["id"] != "T09" else ORANGE)
        add_text(s, Inches(0.45), Inches(y), Inches(1.05), Inches(1.04),
                 t["id"], size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(1.7), Inches(y + 0.06), Inches(10.9), Inches(0.36),
                 f"{t['name']}    ·    {t['quota']}    ·    {t['status']}",
                 size=13, bold=True, color=PRIMARY)
        add_text(s, Inches(1.7), Inches(y + 0.46), Inches(10.9), Inches(0.5),
                 t["one_liner"], size=12, color=DARK)

    # 9 怎么组队
    s, p = new_slide()
    slide_header(s, "怎么选题、怎么组队", "第一志愿 + 备选；老师编组，避免热门题堆人", p, TOTAL)
    make_card(s, Inches(0.5), Inches(1.45), Inches(6.15), Inches(5.4),
              "报名动作", C.HOW_TO_JOIN, accent=PRIMARY, title_size=16, body_size=14)
    make_card(s, Inches(6.85), Inches(1.45), Inches(5.95), Inches(5.4),
              "组内角色（参考）", [
                  "组长：节奏、周报、对外接口",
                  "主开发 / 主笔：扛模块或章节",
                  "数据或资料：口径、引用、台账",
                  "体验或演示：让外人 5 分钟能看懂",
                  "每组至少交出「能打开的东西」",
              ], accent=ACCENT, title_size=16, body_size=14)

    # 10 节奏
    s, p = new_slide()
    slide_header(s, "时间表与周节奏", f"发布日：{C.RELEASE_DATE}", p, TOTAL)
    headers = ["节点", "做什么"]
    make_table(s, Inches(0.7), Inches(1.5), Inches(11.9), Inches(5.3), headers, C.RHYTHM, font_size=14)

    # 11 技能匹配
    s, p = new_slide()
    slide_header(s, "技能与专业匹配", "对不上也没关系，备选可以跨赛道", p, TOTAL)
    headers = ["编号", "适合专业", "关键技能"]
    rows = [[t["id"], t["major"], t["skills"]] for t in C.TOPICS]
    make_table(s, Inches(0.35), Inches(1.4), Inches(12.6), Inches(5.55), headers, rows, font_size=11)

    # 12 你得到 / 我们要求
    s, p = new_slide()
    slide_header(s, "你能得到什么，我们要求什么", "对等：真实题目换稳定投入", p, TOTAL)
    make_card(s, Inches(0.5), Inches(1.45), Inches(6.15), Inches(5.4),
              "同学侧收获", C.WHAT_YOU_GET, accent=ACCENT, title_size=16, body_size=14)
    make_card(s, Inches(6.85), Inches(1.45), Inches(5.95), Inches(5.4),
              "项目组约定", C.WHAT_WE_EXPECT, accent=ORANGE, title_size=16, body_size=14)

    # 13 交付物
    s, p = new_slide()
    slide_header(s, "结题时交什么", "能打开、能讲清、能指出自己做了哪一块", p, TOTAL)
    items = [
        ["产品组", "可运行地址或录屏 + 模块说明 + 未完成清单"],
        ["协议/工程组", "原型或脚本 + 验收记录 + 失败用例"],
        ["研究组", "修订对照 + 证据袋 + 一页摘要"],
        ["共同必交", "周报合订 + 复盘一页 + 成员分工表"],
    ]
    for i, (title, body) in enumerate(items):
        y = 1.5 + i * 1.3
        add_rounded(s, Inches(0.55), Inches(y), Inches(12.2), Inches(1.15), fill=LIGHT)
        add_rect(s, Inches(0.55), Inches(y), Inches(2.3), Inches(1.15), fill=PRIMARY)
        add_text(s, Inches(0.65), Inches(y), Inches(2.1), Inches(1.15),
                 title, size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(3.05), Inches(y), Inches(9.4), Inches(1.15),
                 body, size=16, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

    # 14 下周一怎么发
    s, p = new_slide()
    slide_header(s, "下周一建议这样发出去", "群里用短稿，宣讲用 PPT，细读用 Word，收表用 Excel", p, TOTAL)
    files = [
        ["PPT", "课堂或线上宣讲 15 分钟"],
        ["Word", "老师留底、同学细读每题任务"],
        ["Excel", "志愿、编组、技能、时间表"],
        ["群发短稿", "直接粘贴到学生群"],
        ["HTML", "浏览器打开的一页总览"],
    ]
    for i, (name, desc) in enumerate(files):
        x = 0.45 + i * 2.55
        add_rounded(s, Inches(x), Inches(1.6), Inches(2.4), Inches(3.3), fill=LIGHT)
        add_rect(s, Inches(x), Inches(1.6), Inches(2.4), Inches(0.12), fill=GOLD)
        add_text(s, Inches(x + 0.1), Inches(2.0), Inches(2.2), Inches(0.9),
                 name, size=18, bold=True, color=PRIMARY, align=PP_ALIGN.CENTER)
        add_text(s, Inches(x + 0.12), Inches(2.95), Inches(2.16), Inches(1.6),
                 desc, size=13, color=DARK, align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.55), Inches(5.15), Inches(12.2), Inches(1.6),
             [
                 C.WECHAT_CONTACT_NOTE,
                 "AIDEN 侧不在学生群里单独招生；编组结果由西电老师确认后，再拉项目组工作群。",
                 "热门题超额时：先看备选，再看技能匹配，最后抽签。空题由老师决定是否合并或缓开。",
             ],
             size=14, color=DARK)

    # 15 报名填写
    s, p = new_slide()
    slide_header(s, "同学填报时请写清这些", "Excel「选题意向」表可直接发；以下是字段说明", p, TOTAL)
    headers = ["字段", "怎么填"]
    rows = [
        ["姓名 / 学院专业 / 年级", "与学籍一致，方便老师编组"],
        ["第一志愿编号", "T01–T10 选一"],
        ["备选编号", "必须与第一志愿不同"],
        ["我能贡献什么（≤80 字）", "写技能或已有作品，不写空话"],
        ["每周可投入小时", "建议不少于 8；写真实数字"],
        ["能否短期赴沪", "能 / 不能 / 寒暑假可以"],
        ["角色意向", "组长 / 开发 / 主笔 / 数据 / 演示"],
    ]
    make_table(s, Inches(0.7), Inches(1.45), Inches(11.9), Inches(5.4), headers, rows, font_size=14)

    # 16 封底
    s, _ = new_slide()
    add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), fill=BG_DEEP)
    add_rect(s, Inches(0), Inches(0), Inches(0.18), Inches(7.5), fill=GOLD)
    add_text(s, Inches(0.9), Inches(1.8), Inches(11.5), Inches(0.5),
             "先选题，再组队，再开工", size=18, color=GOLD)
    add_text(s, Inches(0.9), Inches(2.4), Inches(11.8), Inches(1.0),
             "十条课题都在路上。挑你愿意跟完一学期的那一条。",
             size=28, bold=True, color=WHITE)
    add_text(s, Inches(0.9), Inches(3.7), Inches(11.5), Inches(1.2),
             [
                 f"正式发布：{C.RELEASE_DATE}",
                 "配套：PPT · Word · Excel 选题表 · 群发短稿 · HTML 一页总览",
             ],
             size=16, color=SOFT)
    add_text(s, Inches(0.9), Inches(5.4), Inches(11.5), Inches(0.8),
             [C.ORG_LINE, f"{C.MENTOR}  |  {C.DATE_STR}  |  {C.VERSION}"],
             size=14, color=WHITE)

    prs.save(OUT_FILE)
    print(f"已生成 {OUT_FILE}")
    return OUT_FILE


if __name__ == "__main__":
    build()
