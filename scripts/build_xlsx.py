# -*- coding: utf-8 -*-
"""生成《复兴岛全球创客岛收官答卷大会》执行计划 Excel。"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import content as C

OUT = Path(__file__).resolve().parent.parent / "deliverables"
OUT.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT / "复兴岛_全球创客岛_科创出海与具身智能国际大会_执行计划表.xlsx"

FONT = "Microsoft YaHei"
NAVY = "0B1F3A"
TEAL = "0A6E6A"
GOLD = "C48A2A"
LIGHT = "E6F0EE"
ROW2 = "F3F7F6"
WHITE = "FFFFFF"
INK = "1B2A44"

thin = Side(style="thin", color="D5DEEB")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def hfill(color):
    return PatternFill("solid", fgColor=color)


def style_title(ws, text, ncol, color=TEAL):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    c = ws.cell(1, 1, text)
    c.font = Font(name=FONT, size=14, bold=True, color=WHITE)
    c.fill = hfill(color)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 28


def style_sub(ws, text, ncol, row=2):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncol)
    c = ws.cell(row, 1, text)
    c.font = Font(name=FONT, size=10, italic=True, color="5A6B86")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 20


def header_row(ws, row, headers, widths, color=NAVY):
    for j, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row, j, h)
        c.font = Font(name=FONT, size=10.5, bold=True, color=WHITE)
        c.fill = hfill(color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[row].height = 26


def data_rows(ws, start, rows, aligns=None, base_h=24):
    for i, r in enumerate(rows):
        ri = start + i
        for j, val in enumerate(r, start=1):
            c = ws.cell(ri, j, val)
            c.font = Font(name=FONT, size=10, color=INK)
            c.fill = hfill(WHITE if i % 2 == 0 else ROW2)
            al = aligns[j - 1] if aligns else "left"
            c.alignment = Alignment(
                horizontal=al,
                vertical="center",
                wrap_text=True,
                indent=(1 if al == "left" else 0),
            )
            c.border = BORDER
        ws.row_dimensions[ri].height = base_h
    return start + len(rows)


wb = Workbook()

# 1 活动总览
ws = wb.active
ws.title = "1.活动总览"
style_title(ws, f"{C.PROJECT_NAME}｜执行计划表", 4)
style_sub(ws, f"{C.PROJECT_FULL}｜{C.VERSION}｜建议日期 {C.EVENT_DATE_SHORT}", 4)
hr = 4
header_row(ws, hr, ["项目", "内容", "项目", "内容"], [18, 42, 18, 42])
pairs = []
ov = list(C.OVERVIEW)
for i in range(0, len(ov), 2):
    a = ov[i]
    b = ov[i + 1] if i + 1 < len(ov) else ("", "")
    pairs.append((a[0], a[1], b[0], b[1]))
data_rows(ws, hr + 1, pairs, aligns=["center", "left", "center", "left"], base_h=32)

# 2 择日
ws = wb.create_sheet("2.择日专章")
style_title(ws, "择日专章｜9 月 15 日前最靠近收官节点的黄道吉日", 2, GOLD)
style_sub(ws, "主推 9/12；领导周六不便则改 9/9；规避 9/10 杨公忌日", 2)
hr = 4
header_row(ws, hr, ["项目", "内容"], [22, 90], TEAL)
rows = [
    ["首选日期", C.HUANGLI["date"]],
    ["干支 / 值日", f"{C.HUANGLI['ganzhi']} · {C.HUANGLI['zhiri']}"],
    ["冲煞", C.HUANGLI["chong"]],
    ["宜", C.HUANGLI["yi"]],
    ["忌", C.HUANGLI["ji"]],
    ["择日理由", C.HUANGLI["why"]],
]
for b in C.DATE_BACKUP:
    rows.append(["备选", f"{b['date']}｜{b['note']}"])
for a in C.DATE_AVOID:
    rows.append(["规避", a])
data_rows(ws, hr + 1, rows, base_h=36)

# 3 议程
ws = wb.create_sheet("3.详细议程")
style_title(ws, f"详细议程｜{C.EVENT_DATE}", 4)
style_sub(ws, "上午：规格 + 揭牌签约；下午：分论坛 + 国际对接", 4)
hr = 4
header_row(ws, hr, ["时间", "环节", "内容要点", "责任方"], [16, 28, 48, 16])
data_rows(ws, hr + 1, C.AGENDA, aligns=["center", "left", "left", "center"], base_h=28)

# 4 嘉宾
ws = wb.create_sheet("4.嘉宾邀请")
style_title(ws, "嘉宾分层与邀约目标", 3)
style_sub(ws, "政治高位与国际高位必须同时在场", 3)
hr = 4
header_row(ws, hr, ["层级", "邀约对象", "目标"], [22, 70, 36])
rows = [[g["tier"], "；".join(g["targets"]), g["goal"]] for g in C.GUEST_TIERS]
data_rows(ws, hr + 1, rows, base_h=48)

# 5 国别池
ws = wb.create_sheet("5.一带一路国别池")
style_title(ws, "拟邀「一带一路」国家参考池（揭牌优先）", 3, TEAL)
style_sub(ws, "建议确认 6–10 国总领事出席；揭牌国须本人到场", 3)
hr = 4
header_row(ws, hr, ["序号", "国家", "合作侧重"], [10, 22, 60])
rows = [[i + 1, n, f] for i, (n, f) in enumerate(C.BRI_COUNTRY_POOL)]
data_rows(ws, hr + 1, rows, aligns=["center", "center", "left"], base_h=26)

# 6 揭牌
ws = wb.create_sheet("6.揭牌落位")
style_title(ws, "国际会议厅 / 会客厅揭牌落位计划", 4, GOLD)
style_sub(ws, "先锁意向、再上仪式；一国一厅、一厅一责", 4)
hr = 4
header_row(ws, hr, ["落位类型", "数量建议", "形式", "价值"], [32, 18, 40, 40])
rows = [[u["name"], u["count"], u["form"], u["value"]] for u in C.UNVEILING]
data_rows(ws, hr + 1, rows, base_h=40)
r = hr + 1 + len(rows) + 1
ws.cell(r, 1, "落位原则").font = Font(name=FONT, size=11, bold=True, color=TEAL)
for i, p in enumerate(C.UNVEILING_PRINCIPLES):
    ws.cell(r + 1 + i, 1, f"· {p}").font = Font(name=FONT, size=10, color=INK)
    ws.merge_cells(start_row=r + 1 + i, start_column=1, end_row=r + 1 + i, end_column=4)

# 7 席位
ws = wb.create_sheet("7.席位结构")
style_title(ws, "席位结构（200–300 人）", 3)
hr = 4
header_row(ws, hr, ["席别", "人数", "组成"], [28, 16, 70])
data_rows(ws, hr + 1, C.SEAT_PLAN, aligns=["left", "center", "left"], base_h=28)

# 8 组织
ws = wb.create_sheet("8.组织架构")
style_title(ws, "组织架构与职责", 2)
hr = 4
header_row(ws, hr, ["组别", "职责"], [22, 90])
data_rows(ws, hr + 1, C.ORG_STRUCTURE, base_h=30)

# 9 倒排期
ws = wb.create_sheet("9.倒排期")
style_title(ws, "执行倒排期（自即日起）", 3, TEAL)
style_sub(ws, "外事报批建议按 ≥40 天窗口提前启动", 3)
hr = 4
header_row(ws, hr, ["节点", "关键任务", "状态"], [28, 80, 14])
rows = [[a, b, "待启动"] for a, b in C.TIMELINE]
data_rows(ws, hr + 1, rows, aligns=["left", "left", "center"], base_h=32)

# 10 预算
ws = wb.create_sheet("10.预算测算")
style_title(ws, "预算测算（示意，单位：万元）", 3, GOLD)
style_sub(ws, C.BUDGET_NOTE, 3)
hr = 4
header_row(ws, hr, ["成本项", "预算(万元)", "说明"], [36, 16, 55])
data_rows(ws, hr + 1, C.BUDGET, aligns=["left", "center", "left"], base_h=28)

# 11 KPI
ws = wb.create_sheet("11.KPI成效")
style_title(ws, "预期成效与 KPI", 2)
hr = 4
header_row(ws, hr, ["维度", "量化目标"], [22, 90])
data_rows(ws, hr + 1, C.KPIS, base_h=32)

# 12 风险
ws = wb.create_sheet("12.风险预案")
style_title(ws, "风险预案", 3, "B33A3A")
hr = 4
header_row(ws, hr, ["风险", "对策", "等级"], [32, 70, 10])
data_rows(ws, hr + 1, C.RISKS, aligns=["left", "left", "center"], base_h=36)

# 13 下一步
ws = wb.create_sheet("13.下一步行动")
style_title(ws, "下一步行动清单", 2, TEAL)
hr = 4
header_row(ws, hr, ["序号", "行动项"], [10, 100])
rows = [[i + 1, t] for i, t in enumerate(C.NEXT_STEPS)]
data_rows(ws, hr + 1, rows, aligns=["center", "left"], base_h=30)

wb.save(OUT_FILE)
print(f"已生成：{OUT_FILE}")
