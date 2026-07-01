# -*- coding: utf-8 -*-
"""生成《上海 / 长三角产业研学考察计划》Excel 排期表 + 合作框架。"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import content as C

# ---------------------------------------------------------------- 样式
NAVY = "142B45"
BLUE = "1F4E79"
STEEL = "2E6DA4"
AMBER = "E88A1A"
LIGHT = "EEF3F8"
BAND1 = "E3ECF4"
BAND2 = "EDF3EA"
BAND3 = "F6EEE1"
WHITE = "FFFFFF"

F = "微软雅黑"
thin = Side(style="thin", color="C9D6E3")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="center")
WRAP_L = Alignment(wrap_text=True, vertical="center", horizontal="left")
CENTER = Alignment(wrap_text=True, vertical="center", horizontal="center")


def title_font(sz=11, color=WHITE, bold=True):
    return Font(name=F, size=sz, bold=bold, color=color)


def body_font(sz=10, color="222A33", bold=False):
    return Font(name=F, size=sz, bold=bold, color=color)


def fill(hexc):
    return PatternFill("solid", fgColor=hexc)


def style_header_row(ws, row, ncols, bg=NAVY):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill(bg)
        cell.font = title_font()
        cell.alignment = CENTER
        cell.border = BORDER


wb = Workbook()

# ================================================================ Sheet 1 说明
ws = wb.active
ws.title = "说明"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 22
ws.column_dimensions["C"].width = 90

ws["B2"] = C.PROJECT_TITLE
ws["B2"].font = Font(name=F, size=18, bold=True, color=NAVY)
ws["B3"] = C.PROJECT_SUBTITLE
ws["B3"].font = Font(name=F, size=13, bold=True, color=AMBER)

rows = [
    ("主办三方", C.HOSTS),
    ("形式", "四期 · 每期两天半（2.5 天） · 上海（长三角延伸）"),
    ("规模", C.SCALE),
    ("保障", C.LOGISTICS),
    ("考察内容", "标杆项目（房地产标杆项目，非房企） · 标杆园区 · 智能制造 · 科技企业 · 供需闭门撮合"),
    ("战略目标", C.STRATEGIC_GOAL),
    ("本表结构", "四期总览 → 各期排期（时间节点） → 合作框架"),
]
r = 5
for k, v in rows:
    ws.cell(row=r, column=2, value=k).font = body_font(11, NAVY, True)
    ws.cell(row=r, column=2).alignment = WRAP
    ws.cell(row=r, column=2).fill = fill(LIGHT)
    ws.cell(row=r, column=2).border = BORDER
    c = ws.cell(row=r, column=3, value=v)
    c.font = body_font(11)
    c.alignment = WRAP_L
    c.border = BORDER
    ws.row_dimensions[r].height = 30
    r += 1

# ================================================================ Sheet 2 四期总览
ws = wb.create_sheet("四期总览")
ws.sheet_view.showGridLines = False
widths = [10, 14, 26, 40, 34, 30]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.merge_cells("A1:F1")
ws["A1"] = "四期考察总览"
ws["A1"].font = Font(name=F, size=15, bold=True, color=NAVY)
ws["A1"].alignment = WRAP_L
ws.row_dimensions[1].height = 26

heads = ["期次", "时间窗口", "主题", "聚焦方向", "用钢新场景", "本期主题课"]
for c, h in enumerate(heads, start=1):
    ws.cell(row=2, column=c, value=h)
style_header_row(ws, 2, len(heads))
ws.row_dimensions[2].height = 24
bands = [BAND1, BAND2, BAND3, LIGHT]
for i, t in enumerate(C.TOURS):
    row = 3 + i
    vals = [t["code"], t["window"], t["theme"], t["focus"], t["steel"], t["lecture"]]
    for c, v in enumerate(vals, start=1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.fill = fill(bands[i % len(bands)])
        cell.font = body_font(10, NAVY, c in (1, 3))
        cell.alignment = WRAP_L
        cell.border = BORDER
    ws.row_dimensions[row].height = 62

# 参访点清单（追加区块）
start = 8
ws.cell(row=start, column=1, value="各期考察点清单").font = Font(name=F, size=13, bold=True, color=NAVY)
start += 1
heads2 = ["期次", "类别", "考察点（项目 / 园区 / 工厂）", "看点 · 用钢关联"]
for c, h in enumerate(heads2, start=1):
    ws.cell(row=start, column=c, value=h)
style_header_row(ws, start, 4)
ws.row_dimensions[start].height = 22
# widen note columns via merge C:F for readability isn't ideal; keep 4 cols, extend width
ws.column_dimensions["D"].width = 40
r = start + 1
for i, t in enumerate(C.TOURS):
    for j, (cat, name, note) in enumerate(t["sites"]):
        code = t["code"] if j == 0 else ""
        vals = [code, cat, name, note]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.fill = fill(bands[i % len(bands)] if code else WHITE)
            cell.font = body_font(10, STEEL if c == 2 else "222A33", c in (1, 2, 3))
            cell.alignment = WRAP_L
            cell.border = BORDER
        ws.row_dimensions[r].height = 26
        r += 1

# ================================================================ 各期排期表
def schedule_rows(tour):
    """将 2.5 天模板与本期具体考察点合并。"""
    site_idx = 0
    out = []
    for d, tm, act, note in C.DAY_TEMPLATE:
        act2, note2 = act, note
        if act.startswith("参访点"):
            cat, name, sn = tour["sites"][site_idx]
            marker = act[3]  # ①②③④⑤ 符号
            act2 = "参访 " + marker + "：" + name
            note2 = "[" + cat + "] " + sn
            site_idx += 1
        elif "主题报告" in act:
            note2 = tour["lecture"]
        out.append((d, tm, act2, note2))
    return out


def make_schedule_sheet(tour):
    ws = wb.create_sheet(tour["code"] + "排期")
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCD", [12, 16, 40, 52]):
        ws.column_dimensions[col].width = w
    ws.merge_cells("A1:D1")
    ws["A1"] = tour["code"] + " · " + tour["window"] + " · " + tour["theme"] + "（两天半）"
    ws["A1"].font = Font(name=F, size=14, bold=True, color=NAVY)
    ws["A1"].alignment = WRAP_L
    ws.row_dimensions[1].height = 26
    ws.merge_cells("A2:D2")
    ws["A2"] = "聚焦：" + tour["focus"] + "　|　用钢新场景：" + tour["steel"]
    ws["A2"].font = body_font(10, STEEL, True)
    ws["A2"].alignment = WRAP_L
    ws.row_dimensions[2].height = 30

    heads = ["日期", "时间", "环节", "说明 / 考察点看点"]
    for c, h in enumerate(heads, start=1):
        ws.cell(row=3, column=c, value=h)
    style_header_row(ws, 3, 4)
    ws.row_dimensions[3].height = 22
    day_bg = {"第 1 天": BAND1, "第 2 天": BAND2, "第 3 天": BAND3}
    r = 4
    for d, tm, act, note in schedule_rows(tour):
        vals = [d, tm, act, note]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.fill = fill(day_bg.get(d, WHITE))
            cell.font = body_font(10, NAVY if c == 1 else "222A33", c in (1, 3))
            cell.alignment = CENTER if c in (1, 2) else WRAP_L
            cell.border = BORDER
        ws.row_dimensions[r].height = 30
        r += 1


for t in C.TOURS:
    make_schedule_sheet(t)

# ================================================================ 合作框架 Sheet
ws = wb.create_sheet("合作框架")
ws.sheet_view.showGridLines = False
for col, w in zip("ABC", [4, 30, 92]):
    ws.column_dimensions[col].width = w

r = 2
ws.cell(row=r, column=2, value="统筹合作框架").font = Font(name=F, size=16, bold=True, color=NAVY)
r += 1
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
ws.cell(row=r, column=2, value=C.FRAMEWORK_INTRO).font = body_font(11)
ws.cell(row=r, column=2).alignment = WRAP_L
ws.row_dimensions[r].height = 44
r += 2


def section(title):
    global r
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    c = ws.cell(row=r, column=2, value=title)
    c.font = title_font(12, WHITE, True)
    c.fill = fill(BLUE)
    c.alignment = WRAP_L
    ws.row_dimensions[r].height = 24
    r += 1


def kv(k, v, kbg=LIGHT):
    global r
    kc = ws.cell(row=r, column=2, value=k)
    kc.font = body_font(11, NAVY, True); kc.fill = fill(kbg)
    kc.alignment = WRAP_L; kc.border = BORDER
    vc = ws.cell(row=r, column=3, value=v)
    vc.font = body_font(11); vc.alignment = WRAP_L; vc.border = BORDER
    txt = str(v)
    est_lines = sum(max(1, (len(seg) // 44) + 1) for seg in txt.split("\n"))
    ws.row_dimensions[r].height = max(28, 17 * est_lines + 6)
    r += 1


section("一、两阶段合作设计")
for ph in C.PHASES:
    kv(ph["name"], "\n".join("• " + p for p in ph["points"]))
r += 1

section("二、盈利底线与分配原则（红线）")
for title, body in C.PROFIT_PRINCIPLES:
    kv(title, body)
r += 1

section("三、针对对方顾虑（痛点）的应对")
for k, v in C.PAIN_RESPONSES:
    kv(k, v)
r += 1

section("四、收费结构（初期让利，承接原报价逻辑）")
for k, v in C.PRICING:
    kv(k, v)
r += 1

section("五、目标里程碑")
for k, v in C.MILESTONES:
    kv(k, v)
r += 1

section("六、战略目标")
kv("总战略合作伙伴", C.STRATEGIC_GOAL, kbg=BAND3)

# ---------------------------------------------------------------- 打印布局：横向 + 适应页宽
from openpyxl.worksheet.properties import PageSetupProperties
for sh in wb.worksheets:
    sh.page_setup.orientation = "landscape"
    sh.page_setup.fitToWidth = 1
    sh.page_setup.fitToHeight = 0
    sh.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    sh.page_margins.left = sh.page_margins.right = 0.3
    sh.page_margins.top = sh.page_margins.bottom = 0.4
    sh.freeze_panes = None

OUT = "output/中钢协_上海长三角产业研学考察计划_排期表.xlsx"
wb.save(OUT)
print("saved:", OUT, "sheets:", wb.sheetnames)
