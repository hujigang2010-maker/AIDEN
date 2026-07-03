# -*- coding: utf-8 -*-
"""合集版《全球自贸365街区项目群 · 策划服务建议书(合并版)》Word。

合并:广州知识城 + 广州黄埔九佛TOD 两个项目的服务建议书(含提资清单与报价),
并附市场数据(出口TOP20、消费类出口TOP20、楼层功能布局)。商务大气版式。
"""
import os

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor

import content_zsc as ZSC
import content_hp as HP
import content_cgc as CGC
from style_docx import (BLUE, GOLD, GRAY, NAVY, NAVY_HEX, WD_ALIGN_PARAGRAPH,
                        banner, bullet, cn, h1, h2, para, styled_table)

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "deliverables")
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHTGOLD = RGBColor(0xE8, 0xC9, 0x77)

DOC_TITLE = "全球自贸365街区项目群"
DOC_SUB = "策划服务建议书 · 合并版"
PROVIDER_LINE = ZSC.PROVIDER_LINE
DOC_DATE = "2026年7月"


def cover(doc):
    for _ in range(2):
        doc.add_paragraph()
    banner(doc, [
        (DOC_TITLE, 30, WHITE, True),
        ("GLOBAL FREE TRADE 365 BLOCK", 11, LIGHTGOLD, True),
    ])
    p = para(doc, "", space_after=2)
    banner(doc, [
        (DOC_SUB, 22, NAVY, True),
        ("（含提资清单与服务报价 · 附市场数据）", 12, GRAY, False),
    ], fill_hex="F6EEDA")
    for _ in range(2):
        doc.add_paragraph()
    para(doc, "本合集包含以下两个项目:", size=12, bold=True,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    para(doc, "① 广州知识城“全球自贸365街区”项目", size=13, bold=True,
         color=BLUE, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    para(doc, "② 广州黄埔区九佛TOD“全球自贸365街区”项目", size=13, bold=True,
         color=BLUE, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    for _ in range(5):
        doc.add_paragraph()
    para(doc, f"提供方:{PROVIDER_LINE}", size=13, bold=True,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    para(doc, DOC_DATE, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()


def toc(doc):
    h1(doc, "目录", no="·")
    items = [
        "前言", "第一篇 广州知识城“全球自贸365街区”项目",
        "　　一、项目概要　二、服务范围　三、成果交付　四、提资清单　五、服务报价",
        "第二篇 广州黄埔区九佛TOD“全球自贸365街区”项目",
        "　　一、项目概要　二、服务范围　三、成果交付　四、提资清单　五、服务报价",
        "第三篇 附件·市场数据",
        "　　附件一 广州出口TOP20　附件二 消费类出口TOP20　附件三 楼层功能布局",
        "服务团队与联系方式",
    ]
    for it in items:
        para(doc, it, size=12, bold=not it.startswith("　"),
             color=NAVY if not it.startswith("　") else GRAY,
             space_after=6, line=1.3)
    doc.add_page_break()


def preface(doc):
    h1(doc, "前言", no="序")
    para(doc,
         "“全球自贸365街区”以进博会、广交会的常态化外延为核心理念,将短期展会延展为"
         "365天常态化展示与交易场景,叠加综合保税区、自贸区等制度型开放红利,打造品牌出海"
         "商贸3.0基地与永不落幕的展销合作平台。", size=11.5, indent=24)
    para(doc,
         "本合并版建议书由复旦大学住房政策研究中心、上海市杨浦区科技企业联合会联合编制,"
         "涵盖广州知识城与广州黄埔九佛TOD两个“全球自贸365街区”项目的策划服务方案,统一"
         "呈现项目概要、服务范围、成果交付、提资清单与服务报价,并附广州出口市场数据供参考。",
         size=11.5, indent=24)
    doc.add_page_break()


def render_project(doc, C, part_label, part_title):
    h1(doc, part_title, no=part_label)
    # 一、项目概要
    h2(doc, "一、项目概要")
    para(doc, f"{C.CONCEPT_TITLE}:", size=12, bold=True, color=NAVY)
    for pt in C.CONCEPT_POINTS:
        bullet(doc, pt)
    h2(doc, "产业方向")
    para(doc, C.INDUSTRY_INTRO, size=11.5)
    for name, detail in C.INDUSTRY_GROUPS:
        bullet(doc, f"{name}:{detail}")
    h2(doc, "下阶段重点")
    para(doc, C.NEXT_STAGE_TEXT, size=11.5, indent=24)
    # 二、服务范围
    h2(doc, "二、服务范围(策划服务八大模块)")
    for i, (name, detail) in enumerate(C.SERVICE_MODULES, 1):
        para(doc, f"{i}. {name}", size=11.5, bold=True, color=BLUE, space_after=2)
        para(doc, detail, size=11, indent=24, space_after=6)
    # 三、成果交付
    h2(doc, "三、成果交付")
    for d in C.DELIVERABLES:
        bullet(doc, d)
    # 四、提资清单
    h2(doc, "四、提资清单(请委托方提供的资料)")
    rows = [(str(i), cat, name, desc, prio) for i, (cat, name, desc, prio)
            in enumerate(C.INFO_REQUEST_ITEMS, 1)]
    styled_table(doc, ["序号", "类别", "资料名称", "说明", "优先级"], rows,
                 col_widths_cm=[1.1, 2.6, 4.0, 5.6, 1.5],
                 center_cols={0, 4}, body_size=9.5, header_size=10)
    para(doc, "", size=4)
    for n in C.INFO_REQUEST_NOTES:
        bullet(doc, n, size=10.5)
    # 五、服务报价
    h2(doc, "五、服务报价")
    qrows = [(s, w, o, str(f)) for (s, w, o, f) in C.QUOTATION_ITEMS]
    qrows.append(("合计", "", "", str(C.QUOTATION_TOTAL)))
    styled_table(doc, ["工作阶段", "主要工作内容", "主要成果", "报价(万元)"], qrows,
                 col_widths_cm=[3.3, 5.6, 3.6, 2.3], center_cols={3},
                 body_size=9.5, header_size=10, last_row_bold=True)
    para(doc, "", size=4)
    for n in C.QUOTATION_NOTES:
        bullet(doc, n, size=10.5)
    doc.add_page_break()


def appendix(doc):
    h1(doc, "第三篇 附件 · 市场数据", no="附")
    h2(doc, "附件一  2025年广州出口TOP20品类")
    rows = [(str(r[0]), r[1], r[2], r[3], r[4]) for r in CGC.EXPORT_TOP20]
    styled_table(doc, ["排名", "品类", "代表品牌", "广州口岸交易额", "核心说明"], rows,
                 col_widths_cm=[1.1, 3.2, 3.2, 2.8, 4.5], center_cols={0, 3},
                 body_size=8.8, header_size=9.5)
    para(doc, "说明:" + CGC.EXPORT_TOP20_NOTE, size=9.5, color=GRAY, space_before=4)
    doc.add_page_break()
    h2(doc, "附件二  2025年前10月广州消费类出口20强(单位:亿元)")
    rows = [(str(r[0]), r[1], r[2], r[3], r[4], r[5]) for r in CGC.CONSUMER_TOP20]
    styled_table(doc, ["排名", "品类", "出口额", "同比", "核心出口品牌", "核心市场"], rows,
                 col_widths_cm=[1.1, 3.0, 1.8, 1.6, 4.0, 3.0],
                 center_cols={0, 2, 3}, body_size=8.8, header_size=9.5)
    para(doc, "说明:" + CGC.CONSUMER_TOP20_NOTE, size=9.5, color=GRAY, space_before=4)
    doc.add_page_break()
    h2(doc, "附件三  黄埔九佛TOD“全球自贸365街区”楼层功能布局")
    rows = [(r[0], r[1], r[2], r[3]) for r in CGC.FLOOR_LAYOUT]
    styled_table(doc, ["区位楼层", "功能定位", "核心业态 / 服务", "数据支撑与参考案例"],
                 rows, col_widths_cm=[2.6, 3.2, 5.2, 4.2], center_cols={0},
                 body_size=8.8, header_size=9.5)
    doc.add_page_break()


def team(doc):
    h1(doc, "服务团队与联系方式", no="队")
    para(doc,
         "本项目由复旦大学住房政策研究中心与上海市杨浦区科技企业联合会联合组建服务团队:"
         "研究中心负责政策研究、项目定位与策划方案编制;科技企业联合会负责产业资源组织、"
         "企业对接与招商建议。两家机构协同为两个“全球自贸365街区”项目提供一体化策划服务。",
         size=11.5, indent=24)
    para(doc, "联系方式:另行提供(以正式合同联络人为准)。", size=11.5, indent=24)


def main():
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(2.4); sec.bottom_margin = Cm(2.4)
        sec.left_margin = Cm(2.6); sec.right_margin = Cm(2.6)
    cover(doc)
    toc(doc)
    preface(doc)
    render_project(doc, ZSC, "壹", "第一篇  广州知识城“全球自贸365街区”项目")
    render_project(doc, HP, "贰", "第二篇  广州黄埔区九佛TOD“全球自贸365街区”项目")
    appendix(doc)
    team(doc)
    os.makedirs(OUT_DIR, exist_ok=True)
    out = os.path.join(OUT_DIR, "全球自贸365街区项目群_策划服务建议书_合并版.docx")
    doc.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main()
