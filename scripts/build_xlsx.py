# -*- coding: utf-8 -*-
"""生成 Excel：上海AI博物馆投资 / 盈利 / 政策落地表。"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import content as C

OUT = Path(__file__).resolve().parent.parent / "deliverables"
OUT.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT / "上海AI博物馆_投资盈利政策落地表.xlsx"

PRIMARY = "0B3D5C"
ACCENT = "0F7A6E"
LIGHT = "D7EBE6"
SAND = "F4F7F6"
WHITE = "FFFFFF"
DARK = "1A2B2E"
GREY = "5C6B70"
AMBER = "C47B2D"


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
        f"{C.PROJECT} · 落地总览",
        f"{C.VERSION} | {C.DATE_STR} | 核心：投资路径 / 盈利模式 / 政策扶持",
    )
    banner(ws, 4, "一、项目一句话", 8, PRIMARY)
    ws.merge_cells("A5:H5")
    ws["A5"] = (
        "科协牵头、多主体出资建设「AI全产业链博物馆」，项目方轻资产运营："
        "用场地免租+设备配套+赞助捐赠完成建设，用租金/研学/孵化/模板输出/海外内容实现盈利。"
    )
    ws["A5"].alignment = Alignment(wrap_text=True, vertical="center")
    ws["A5"].font = Font(name="微软雅黑", size=11, color=DARK)
    ws.row_dimensions[5].height = 48

    banner(ws, 7, "二、核心关注（本表重点）", 8)
    header_row(ws, 8, ["序号", "关注点", "对应Sheet", "当前状态"])
    focus = [
        (1, C.CORE_CONCERNS[0], "02_投资路径", "进行中"),
        (2, C.CORE_CONCERNS[1], "03_盈利模式", "进行中"),
        (3, C.CORE_CONCERNS[2], "04_政策扶持", "进行中"),
    ]
    for i, row in enumerate(focus, 9):
        write_row(ws, i, list(row))
    add_status_dv(ws, "D", 9, 11)

    banner(ws, 13, "三、阶段里程碑", 8, PRIMARY)
    header_row(ws, 14, ["阶段", "时间窗", "关键交付", "退出标准", "状态", "负责人", "备注"])
    for i, r in enumerate(C.ROADMAP[1:], 15):
        write_row(ws, i, list(r) + ["未开始", "", ""])
    add_status_dv(ws, "E", 15, 15 + len(C.ROADMAP[1:]) - 1)

    banner(ws, 23, "四、来源笔记", 8)
    write_row(ws, 24, ["场次一", C.SOURCE_NOTES[1], "2026-08-08", "出行讨论"])
    write_row(ws, 25, ["场次二", C.SOURCE_NOTES[0], "2026-08-10", "筹建工作会议"])
    set_widths(ws, [18, 36, 28, 28, 12, 14, 18, 18])


def sheet_minutes(wb):
    ws = wb.create_sheet("01_会议纪要")
    big_title(ws, "F", "双场会议纪要提炼", "仅保留与立项/投资/盈利/政策相关的结论")
    banner(ws, 4, "会议元信息", 6, PRIMARY)
    header_row(ws, 5, ["场次", "时间", "时长", "人数", "类型", "来源"])
    for i, m in enumerate(C.MEETING_META, 6):
        write_row(ws, i, [m["场次"], m["时间"], m["时长"], m["人数"], m["类型"], m["来源"]])

    banner(ws, 9, "纪要总括", 6)
    ws.merge_cells("A10:F10")
    ws["A10"] = C.MEETING_SUMMARY
    ws["A10"].alignment = Alignment(wrap_text=True)
    ws["A10"].font = Font(name="微软雅黑", size=10)
    ws.row_dimensions[10].height = 55

    banner(ws, 12, "分议题结论", 6, PRIMARY)
    header_row(ws, 13, ["议题", "结论"])
    for i, kp in enumerate(C.MEETING_KEY_POINTS, 14):
        write_row(ws, i, [kp["议题"], kp["结论"]])
        ws.row_dimensions[i].height = 48

    start = 14 + len(C.MEETING_KEY_POINTS) + 1
    banner(ws, start, "待办事项", 6)
    header_row(ws, start + 1, ["责任人", "事项", "优先级", "时间", "状态"])
    for i, t in enumerate(C.TODOS, start + 2):
        write_row(ws, i, t + ["未开始"])
    add_status_dv(ws, "E", start + 2, start + 1 + len(C.TODOS))
    set_widths(ws, [22, 70, 14, 14, 12, 40])


def sheet_invest(wb):
    ws = wb.create_sheet("02_投资路径")
    big_title(ws, "H", "投资方怎么投资（七条可组合路径）", C.INVEST_PRINCIPLE)
    banner(ws, 4, "出资路径明细", 8, PRIMARY)
    header_row(ws, 5, ["路径", "出资方式", "出资方画像", "回报机制", "落地动作", "优先级", "状态", "跟进人"])
    for i, it in enumerate(C.INVEST_PATHS, 6):
        write_row(ws, i, [
            it["路径"], it["出资方式"], it["出资方画像"], it["回报机制"],
            it["落地动作"], it["优先级"], "未开始", "",
        ])
        ws.row_dimensions[i].height = 55
    add_status_dv(ws, "G", 6, 5 + len(C.INVEST_PATHS))

    r0 = 6 + len(C.INVEST_PATHS) + 1
    banner(ws, r0, "赞助层级产品化（便于企业立项）", 8)
    header_row(ws, r0 + 1, C.SPONSOR_TIERS[0] + ["对接状态", "意向方", "备注"])
    for i, row in enumerate(C.SPONSOR_TIERS[1:], r0 + 2):
        write_row(ws, i, row + ["未开始", "", ""])
    add_status_dv(ws, "E", r0 + 2, r0 + 1 + len(C.SPONSOR_TIERS[1:]))
    set_widths(ws, [22, 40, 36, 40, 36, 10, 12, 14])


def sheet_profit(wb):
    ws = wb.create_sheet("03_盈利模式")
    big_title(ws, "G", "项目如何盈利（六条收入线）", C.PROFIT_LOGIC)
    banner(ws, 4, "收入线设计", 7, PRIMARY)
    header_row(ws, 5, ["收入线", "描述", "启动条件", "稳态占比", "里程碑", "状态", "负责人"])
    for i, r in enumerate(C.REVENUE_STREAMS, 6):
        write_row(ws, i, [
            r["收入线"], r["描述"], r["启动条件"], r["目标占比(稳态)"],
            r["里程碑"], "未开始", "",
        ])
        ws.row_dimensions[i].height = 50
    add_status_dv(ws, "F", 6, 5 + len(C.REVENUE_STREAMS))

    r0 = 6 + len(C.REVENUE_STREAMS) + 1
    banner(ws, r0, "空间与经济假设（讨论稿）", 7)
    header_row(ws, r0 + 1, C.UNIT_ECONOMICS[0] + ["确认状态", "备注"])
    for i, row in enumerate(C.UNIT_ECONOMICS[1:], r0 + 2):
        write_row(ws, i, row + ["待确认", ""])
    set_widths(ws, [28, 55, 36, 16, 32, 12, 14])


def sheet_policy(wb):
    ws = wb.create_sheet("04_政策扶持")
    big_title(ws, "G", "政策性支持与扶持基金落地", "把政策变成可申报、可签约的动作清单")
    banner(ws, 4, "政策抓手清单", 7, PRIMARY)
    header_row(ws, 5, ["政策/抓手", "要点", "对项目价值", "落地步骤", "责任方", "状态", "材料准备"])
    for i, x in enumerate(C.POLICY_SUPPORT, 6):
        write_row(ws, i, [
            x["政策/抓手"], x["要点"], x["对项目价值"], x["落地步骤"],
            x["责任方"], "未开始", "",
        ])
        ws.row_dimensions[i].height = 48
    add_status_dv(ws, "F", 6, 5 + len(C.POLICY_SUPPORT))

    r0 = 6 + len(C.POLICY_SUPPORT) + 1
    banner(ws, r0, "扶持资金类型匹配", 7)
    header_row(ws, r0 + 1, C.POLICY_FUND_MATCH[0] + ["窗口日期", "申报状态", "金额目标"])
    for i, row in enumerate(C.POLICY_FUND_MATCH[1:], r0 + 2):
        write_row(ws, i, row + ["", "未开始", ""])
    add_status_dv(ws, "F", r0 + 2, r0 + 1 + len(C.POLICY_FUND_MATCH[1:]))
    set_widths(ws, [28, 40, 28, 40, 18, 12, 16])


def sheet_site(wb):
    ws = wb.create_sheet("05_场地候选")
    big_title(ws, "G", "场地候选与谈判要点", "谈判核心：免租10–20年 + 约5000万设备装修")
    header_row(ws, 4, C.SITE_CANDIDATES[0] + ["谈判状态"])
    for i, row in enumerate(C.SITE_CANDIDATES[1:], 5):
        write_row(ws, i, row + ["未开始"])
        ws.row_dimensions[i].height = 55
    add_status_dv(ws, "G", 5, 4 + len(C.SITE_CANDIDATES[1:]))
    banner(ws, 8, "谈判话术要点（内部）", 7, ACCENT)
    tips = [
        "对场地方：盘活闲置 + 高层考察场景 + 剩余面积招商优先/租金分成",
        "对区政府：WAIC会后365天政绩 + 研学指标 + 科创名片",
        "对资方：科创考核故事 + AI历史席位（比广告好立项）",
        "红线：建设期项目方尽量零出资本金；先封闭场地与设备资金再画施工图",
    ]
    for i, t in enumerate(tips, 9):
        write_row(ws, i, [t])
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=7)
    set_widths(ws, [20, 16, 36, 28, 32, 28, 12])


def sheet_exhibit(wb):
    ws = wb.create_sheet("06_展陈资源")
    big_title(ws, "F", "展陈分层与资源对接", "美方理论 / 中方应用")
    header_row(ws, 4, C.EXHIBIT_LAYERS[0] + ["对接负责人", "状态"])
    for i, row in enumerate(C.EXHIBIT_LAYERS[1:], 5):
        write_row(ws, i, row + ["", "未开始"])
        ws.row_dimensions[i].height = 40
    add_status_dv(ws, "F", 5, 4 + len(C.EXHIBIT_LAYERS[1:]))
    set_widths(ws, [20, 40, 36, 28, 14, 12])


def sheet_org(wb):
    ws = wb.create_sheet("07_组织分工")
    big_title(ws, "E", "组织角色", "中美双团队推进")
    header_row(ws, 4, C.ORG_ROLES[0] + ["联系人", "状态"])
    for i, row in enumerate(C.ORG_ROLES[1:], 5):
        write_row(ws, i, row + ["", "未开始"])
    add_status_dv(ws, "E", 5, 4 + len(C.ORG_ROLES[1:]))
    set_widths(ws, [16, 28, 50, 14, 12])


def sheet_risks(wb):
    ws = wb.create_sheet("08_风险台账")
    big_title(ws, "F", "风险与缓释", "谈判与预算为当前最高优先级风险")
    header_row(ws, 4, C.RISKS[0] + ["责任人", "状态"])
    for i, row in enumerate(C.RISKS[1:], 5):
        write_row(ws, i, row + ["", "未开始"])
        ws.row_dimensions[i].height = 40
    add_status_dv(ws, "F", 5, 4 + len(C.RISKS[1:]))
    set_widths(ws, [28, 8, 36, 40, 12, 12])


def sheet_tracker(wb):
    ws = wb.create_sheet("09_资方对接跟踪")
    big_title(ws, "J", "资方 / 场地方 / 政策方对接跟踪表", "路演后务必回填")
    header_row(ws, 4, [
        "对象", "类型", "对应路径", "对接人", "最近接触日",
        "意向度(高/中/低)", "卡点", "下一步", "状态", "备注",
    ])
    seeds = [
        ["华润相关主体", "主赞助/场地", "B / A", "", "", "", "同学通道材料", "一页纸权益包", "未开始", ""],
        ["杨浦科创集团/中建四局", "场地方", "A", "", "", "", "免租+5000万", "创智汇路演", "未开始", "3300㎡一期"],
        ["复兴岛主管单位", "场地方", "A", "", "", "", "产权与改造", "踏勘", "未开始", ""],
        ["中国科协", "主管/专项", "E", "", "", "", "正式文件", "汇报材料", "未开始", "已初步沟通"],
        ["区政府/科创委", "政策资金", "D", "", "", "", "入库与补贴", "重点项目申报", "未开始", ""],
        ["头部大模型企业（沪）", "金牌赞助", "C", "", "", "", "展席权益", "集中拜访", "未开始", ""],
        ["曾达", "生态合作", "C/G", "", "", "", "角色定义", "邀约", "未开始", ""],
        ["海外科技企业", "内容授权", "G", "", "", "", "接受度/合规", "陈院长清单", "未开始", ""],
    ]
    for i, row in enumerate(seeds, 5):
        write_row(ws, i, row)
    add_status_dv(ws, "I", 5, 4 + len(seeds))
    set_widths(ws, [22, 14, 12, 12, 14, 14, 18, 18, 12, 20])


def build():
    wb = Workbook()
    sheet_dashboard(wb)
    sheet_minutes(wb)
    sheet_invest(wb)
    sheet_profit(wb)
    sheet_policy(wb)
    sheet_site(wb)
    sheet_exhibit(wb)
    sheet_org(wb)
    sheet_risks(wb)
    sheet_tracker(wb)
    wb.save(OUT_FILE)
    print(f"已生成: {OUT_FILE}")


if __name__ == "__main__":
    build()
