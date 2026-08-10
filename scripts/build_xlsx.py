# -*- coding: utf-8 -*-
"""生成 Excel：上海人工智能产业展馆 · 对外汇报配套表。"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import content as C

OUT = Path(__file__).resolve().parent.parent / "deliverables"
OUT.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT / "上海人工智能产业展馆_投资盈利政策落地表.xlsx"

PRIMARY = "0B3D5C"
ACCENT = "0F7A6E"
LIGHT = "D7EBE6"
SAND = "F4F7F6"
WHITE = "FFFFFF"
DARK = "1A2B2E"
GREY = "5C6B70"


def thin():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)


def header_row(ws, row, headers, fill=PRIMARY):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(name="微软雅黑", size=11, bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = thin()
    ws.row_dimensions[row].height = 28


def write_row(ws, row, data, fill=None, bold=False, size=10):
    for i, v in enumerate(data, 1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name="微软雅黑", size=size, bold=bold, color=DARK)
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c.border = thin()
        if fill:
            c.fill = PatternFill("solid", fgColor=fill)
        elif row % 2 == 0:
            c.fill = PatternFill("solid", fgColor=SAND)


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def banner(ws, row, text, span, fill=ACCENT):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name="微软雅黑", size=12, bold=True, color=WHITE)
    c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 24


def big_title(ws, span_col, text, sub):
    ws.merge_cells(f"A1:{span_col}1")
    c = ws["A1"]
    c.value = text
    c.font = Font(name="微软雅黑", size=16, bold=True, color=WHITE)
    c.fill = PatternFill("solid", fgColor=PRIMARY)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 34
    ws.merge_cells(f"A2:{span_col}2")
    c2 = ws["A2"]
    c2.value = sub
    c2.font = Font(name="微软雅黑", size=10, italic=True, color=GREY)
    c2.fill = PatternFill("solid", fgColor=LIGHT)
    c2.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 20


def add_status_dv(ws, col, start, end):
    dv = DataValidation(
        type="list",
        formula1='"未开始,进行中,已完成,阻塞,取消"',
        allow_blank=True,
    )
    ws.add_data_validation(dv)
    dv.add(f"{col}{start}:{col}{end}")


def sheet_dashboard(wb):
    ws = wb.active
    ws.title = "00_总览"
    big_title(
        ws, "H",
        f"{C.PROJECT} · 对外汇报总览",
        f"{C.VERSION} | {C.DATE_STR} | 对标 CHM | 租金装修重点 | 多元盈利 | 支持单位",
    )
    banner(ws, 4, "一、一句话", 8, PRIMARY)
    ws.merge_cells("A5:H5")
    ws["A5"] = C.EXEC_SUMMARY
    ws["A5"].alignment = Alignment(wrap_text=True, vertical="center")
    ws["A5"].font = Font(name="微软雅黑", size=11, color=DARK)
    ws.row_dimensions[5].height = 60

    banner(ws, 7, "二、核心议题", 8)
    header_row(ws, 8, ["序号", "议题", "对应Sheet", "状态"])
    focus = [
        (1, C.CORE_CONCERNS[0], "03_租金生态/04_装修分摊", "进行中"),
        (2, C.CORE_CONCERNS[1], "02_CHM对标", "进行中"),
        (3, C.CORE_CONCERNS[2], "05_盈利模式/06_支持单位", "进行中"),
    ]
    for i, row in enumerate(focus, 9):
        write_row(ws, i, list(row))
    add_status_dv(ws, "D", 9, 11)

    banner(ws, 13, "三、阶段里程碑", 8, PRIMARY)
    header_row(ws, 14, ["阶段", "时间窗", "关键交付", "退出标准", "状态", "负责人", "备注"])
    for i, r in enumerate(C.ROADMAP[1:], 15):
        write_row(ws, i, list(r) + ["未开始", "", ""])
    add_status_dv(ws, "E", 15, 14 + len(C.ROADMAP[1:]))
    set_widths(ws, [18, 36, 36, 28, 12, 14, 18, 18])


def sheet_chm(wb):
    ws = wb.create_sheet("02_CHM对标")
    big_title(ws, "C", "对标加州计算机历史博物馆（一对一）", C.BENCHMARK)
    banner(ws, 4, "样板馆画像", 3, PRIMARY)
    header_row(ws, 5, ["维度", "要点"])
    for i, (k, v) in enumerate(C.CHM_PROFILE.items(), 6):
        write_row(ws, i, [k, v])
        ws.row_dimensions[i].height = 36
    r0 = 6 + len(C.CHM_PROFILE) + 1
    banner(ws, r0, "一对一对照", 3)
    header_row(ws, r0 + 1, C.CHM_COMPARE[0])
    for i, row in enumerate(C.CHM_COMPARE[1:], r0 + 2):
        write_row(ws, i, row)
        ws.row_dimensions[i].height = 48
    set_widths(ws, [22, 45, 45])


def sheet_rent(wb):
    ws = wb.create_sheet("03_租金生态")
    big_title(ws, "E", "租金产业生态（四级结构）", C.RENT_ECOLOGY)
    header_row(ws, 4, ["层级", "逻辑", "谁承担", "谈判点", "状态"])
    for i, x in enumerate(C.RENT_LAYERS, 5):
        write_row(ws, i, [x["层级"], x["逻辑"], x["谁承担"], x["谈判点"], "未开始"])
        ws.row_dimensions[i].height = 55
    add_status_dv(ws, "E", 5, 4 + len(C.RENT_LAYERS))
    set_widths(ws, [28, 42, 36, 42, 12])


def sheet_fitout(wb):
    ws = wb.create_sheet("04_装修分摊")
    big_title(ws, "E", "装修费分层共担", C.FITOUT_ECOLOGY)
    header_row(ws, 4, ["科目", "典型内容", "建议承担", "回收方式", "状态"])
    for i, x in enumerate(C.FITOUT_LAYERS, 5):
        write_row(ws, i, [x["科目"], x["典型内容"], x["建议承担"], x["回收方式"], "未开始"])
        ws.row_dimensions[i].height = 50
    add_status_dv(ws, "E", 5, 4 + len(C.FITOUT_LAYERS))
    r0 = 5 + len(C.FITOUT_LAYERS) + 1
    banner(ws, r0, "落地步骤", 5, ACCENT)
    header_row(ws, r0 + 1, C.RENT_FITOUT_PLAYBOOK[0] + ["状态", "负责人"])
    for i, row in enumerate(C.RENT_FITOUT_PLAYBOOK[1:], r0 + 2):
        write_row(ws, i, row + ["未开始", ""])
    add_status_dv(ws, "D", r0 + 2, r0 + 1 + len(C.RENT_FITOUT_PLAYBOOK[1:]))
    set_widths(ws, [18, 40, 36, 32, 12])


def sheet_profit(wb):
    ws = wb.create_sheet("05_盈利模式")
    big_title(ws, "F", "多元盈利模式", C.PROFIT_LOGIC)
    header_row(ws, 4, ["收入线", "描述", "延伸玩法", "占比", "里程碑", "状态"])
    for i, r in enumerate(C.REVENUE_STREAMS, 5):
        write_row(ws, i, [
            r["收入线"], r["描述"], r["想象力延伸"], r["占比"], r["里程碑"], "未开始",
        ])
        ws.row_dimensions[i].height = 48
    add_status_dv(ws, "F", 5, 4 + len(C.REVENUE_STREAMS))
    set_widths(ws, [26, 40, 40, 14, 28, 12])


def sheet_support(wb):
    ws = wb.create_sheet("06_支持单位")
    big_title(ws, "D", "支持单位协同", "备忘录 + 年度清单")
    header_row(ws, 4, C.SUPPORT_ORG_ROLES[0] + ["对接状态"])
    for i, row in enumerate(C.SUPPORT_ORG_ROLES[1:], 5):
        write_row(ws, i, row + ["未开始"])
        ws.row_dimensions[i].height = 50
    add_status_dv(ws, "D", 5, 4 + len(C.SUPPORT_ORG_ROLES[1:]))
    set_widths(ws, [32, 22, 55, 12])


def sheet_invest(wb):
    ws = wb.create_sheet("07_投资路径")
    big_title(ws, "G", "投资与赞助", C.INVEST_PRINCIPLE)
    header_row(ws, 4, ["路径", "出资方式", "出资方画像", "回报", "优先级", "状态", "跟进人"])
    for i, it in enumerate(C.INVEST_PATHS, 5):
        write_row(ws, i, [
            it["路径"], it["出资方式"], it["出资方画像"], it["回报机制"],
            it["优先级"], "未开始", "",
        ])
        ws.row_dimensions[i].height = 48
    add_status_dv(ws, "F", 5, 4 + len(C.INVEST_PATHS))
    set_widths(ws, [22, 36, 36, 32, 10, 12, 12])


def sheet_policy(wb):
    ws = wb.create_sheet("08_政策资金")
    big_title(ws, "F", "政策与资金匹配", "保守到账入预算")
    header_row(ws, 4, ["政策/抓手", "要点", "价值", "步骤", "责任方", "状态"])
    for i, x in enumerate(C.POLICY_SUPPORT, 5):
        write_row(ws, i, [
            x["政策/抓手"], x["要点"], x["对项目价值"], x["落地步骤"], x["责任方"], "未开始",
        ])
        ws.row_dimensions[i].height = 40
    add_status_dv(ws, "F", 5, 4 + len(C.POLICY_SUPPORT))
    set_widths(ws, [24, 28, 24, 32, 16, 12])


def sheet_site(wb):
    ws = wb.create_sheet("09_选址内容")
    big_title(ws, "G", "选址与内容", "一期试点")
    header_row(ws, 4, C.SITE_CANDIDATES[0] + ["状态"])
    for i, row in enumerate(C.SITE_CANDIDATES[1:], 5):
        write_row(ws, i, row + ["未开始"])
        ws.row_dimensions[i].height = 48
    r0 = 5 + len(C.SITE_CANDIDATES[1:]) + 1
    banner(ws, r0, "内容结构（对标 CHM）", 7)
    header_row(ws, r0 + 1, C.CONTENT_LAYERS[0] + ["状态", "负责人"])
    for i, row in enumerate(C.CONTENT_LAYERS[1:], r0 + 2):
        write_row(ws, i, row + ["未开始", ""])
    set_widths(ws, [22, 18, 36, 28, 28, 22, 12])


def sheet_tracker(wb):
    ws = wb.create_sheet("10_对接跟踪")
    big_title(ws, "J", "对外对接跟踪", "路演后回填")
    header_row(ws, 4, [
        "对象", "类型", "路径", "对接人", "日期", "意向", "卡点", "下一步", "状态", "备注",
    ])
    seeds = [
        ["杨浦科创/中建四局", "物业", "A", "陈红苗", "", "", "租金阶梯+壳装", "创智汇路演", "未开始", ""],
        ["复兴岛主管单位", "物业", "A", "陈红苗", "", "", "改造估算", "踏勘", "未开始", ""],
        ["主赞助候选", "赞助", "B", "胡继刚", "", "", "权益包", "立项材料", "未开始", ""],
        ["复旦大学住房政策研究中心", "支持单位", "E", "胡继刚", "", "", "备忘录", "联合课题", "未开始", ""],
        ["上海市科技企业联合会", "支持单位", "E", "胡继刚", "", "", "会员对接", "备忘录", "未开始", ""],
        ["杨浦区科技企业联合会", "支持单位", "E", "陈红苗", "", "", "属地企业", "活动联办", "未开始", ""],
        ["虹口区科技企业联合会", "支持单位", "E", "陈红苗", "", "", "跨区联动", "权益互通", "未开始", ""],
        ["在沪科技企业", "展位", "C", "", "", "", "年更合同", "集中拜访", "未开始", ""],
    ]
    for i, row in enumerate(seeds, 5):
        write_row(ws, i, row)
    add_status_dv(ws, "I", 5, 4 + len(seeds))
    set_widths(ws, [26, 12, 8, 12, 12, 10, 18, 16, 12, 14])


def build():
    wb = Workbook()
    sheet_dashboard(wb)
    sheet_chm(wb)
    sheet_rent(wb)
    sheet_fitout(wb)
    sheet_profit(wb)
    sheet_support(wb)
    sheet_invest(wb)
    sheet_policy(wb)
    sheet_site(wb)
    sheet_tracker(wb)
    wb.save(OUT_FILE)
    print(f"已生成: {OUT_FILE}")


if __name__ == "__main__":
    build()
