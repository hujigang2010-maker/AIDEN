# -*- coding: utf-8 -*-
"""生成 Excel：90 天任务跟踪、选址比选、对接清单。"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import content as C

OUT = Path(__file__).resolve().parent.parent / "deliverables"
OUT.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT / "东昇聚变_政府事务与上海产业落地_90天工作台账.xlsx"

NAVY = "071A2B"
CYAN = "1FA8B4"
EMBER = "E36B2C"
SAND = "F4F1EA"
MIST = "EEF3F6"
WHITE = "FFFFFF"
INK = "1A2433"
GREY = "5B6775"

thin = Border(
    left=Side(style="thin", color="D5DDE5"),
    right=Side(style="thin", color="D5DDE5"),
    top=Side(style="thin", color="D5DDE5"),
    bottom=Side(style="thin", color="D5DDE5"),
)


def font(size=11, bold=False, color=INK, name="微软雅黑"):
    return Font(name=name, size=size, bold=bold, color=color)


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def align(h="left", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def style_header(ws, row, cols, fill_color=NAVY):
    for col in range(1, cols + 1):
        cell = ws.cell(row, col)
        cell.font = font(10, bold=True, color=WHITE)
        cell.fill = fill(fill_color)
        cell.alignment = align("center")
        cell.border = thin


def style_body(ws, start, end, cols):
    for r in range(start, end + 1):
        for c in range(1, cols + 1):
            cell = ws.cell(r, c)
            cell.font = font(10)
            cell.alignment = align("left" if c > 1 else "center")
            cell.border = thin
            cell.fill = fill(WHITE if (r - start) % 2 == 0 else MIST)


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def freeze_title(ws, title, subtitle, cols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=cols)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=cols)
    a = ws.cell(1, 1, title)
    a.font = font(16, bold=True, color=WHITE)
    a.fill = fill(NAVY)
    a.alignment = align("left")
    b = ws.cell(2, 1, subtitle)
    b.font = font(10, color=WHITE)
    b.fill = fill(CYAN)
    b.alignment = align("left")
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 20
    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.print_title_rows = "1:3"
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def write_matrix(ws, headers, rows, start_row=3):
    for i, h in enumerate(headers, 1):
        ws.cell(start_row, i, h)
    style_header(ws, start_row, len(headers))
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row, 1):
            ws.cell(start_row + 1 + ri, ci, val)
        ws.row_dimensions[start_row + 1 + ri].height = 36
    style_body(ws, start_row + 1, start_row + len(rows), len(headers))


def build():
    wb = Workbook()

    # 总览
    ws = wb.active
    ws.title = "90天总览"
    freeze_title(ws, f"{C.COMPANY} · {C.DOC_TITLE}", f"{C.AUTHOR}  |  {C.DATE_STR}  |  {C.VERSION}  |  {C.TAGLINE}", 4)
    write_matrix(
        ws,
        ["阶段", "时间", "一句话", "阶段成果"],
        [[p["name"], p["span"], p["one"], "；".join(p["outcomes"])] for p in C.PHASES],
    )
    set_widths(ws, [16, 18, 36, 70])
    ws.cell(8, 1, "三件套")
    ws.cell(8, 1).font = font(12, bold=True, color=NAVY)
    write_matrix(
        ws,
        ["交付物", "内容", "审定"],
        [[d["name"], d["what"], d["owner"]] for d in C.DELIVERABLES],
        start_row=9,
    )

    # 分周
    ws = wb.create_sheet("分周计划")
    freeze_title(ws, "分周计划", "每周五提交一页纸周报：见了谁、进展、风险、需决策事项", 5)
    rows = []
    status_options = '"未开始,进行中,已完成,暂缓"'
    for i, row in enumerate(C.WEEKLY_PLAN[1:]):
        rows.append(list(row) + ["未开始"])
    write_matrix(ws, C.WEEKLY_PLAN[0] + ["状态"], rows)
    set_widths(ws, [12, 12, 36, 42, 12])
    dv = DataValidation(type="list", formula1=status_options, allow_blank=False)
    dv.error = "请选择状态"
    dv.errorTitle = "无效状态"
    dv.prompt = "选择当前状态"
    dv.promptTitle = "状态"
    ws.add_data_validation(dv)
    dv.add("E4:E20")

    # 产出物
    ws = wb.create_sheet("产出物")
    freeze_title(ws, "产出物时间表", "每个阶段都有纸面结果，避免一直在跑、没有交卷", 4)
    write_matrix(ws, C.OUTPUT_LIST[0] + ["状态"], [list(r) + ["未开始"] for r in C.OUTPUT_LIST[1:]])
    set_widths(ws, [14, 42, 22, 12])
    dv2 = DataValidation(type="list", formula1=status_options, allow_blank=False)
    ws.add_data_validation(dv2)
    dv2.add("D4:D20")

    # 对接
    ws = wb.create_sheet("政府关系作战图")
    freeze_title(ws, "政府关系作战图（90天口径）", "补位不抢位。拜访清单须分管高管过目。", 5)
    write_matrix(
        ws,
        C.GOV_CONTACTS[0] + ["进展备注"],
        [list(r) + [""] for r in C.GOV_CONTACTS[1:]],
    )
    set_widths(ws, [28, 36, 36, 22, 24])

    # 选址
    ws = wb.create_sheet("选址比选")
    freeze_title(ws, "上海产业落地选址比选", "重资产看财政、场域、承重；轻资产看高校与窗口期", 5)
    write_matrix(ws, C.SITE_ROWS[0], C.SITE_ROWS[1:])
    set_widths(ws, [16, 36, 36, 36, 28])
    for r in range(4, 4 + len(C.SITE_ROWS[1:])):
        name = str(ws.cell(r, 1).value or "")
        if "金桥" in name or "张江" in name:
            ws.cell(r, 5).fill = fill("D6F3F5")
        elif "马桥" in name:
            ws.cell(r, 5).fill = fill("F8D7C8")
        elif "徐汇" in name or "枢纽" in name:
            ws.cell(r, 5).fill = fill(SAND)

    # 现场六指标
    ws = wb.create_sheet("现场六指标")
    freeze_title(ws, "现场核验记录", "容积率、建筑密度、单位荷载、货梯、建筑净深、室内净高", 8)
    headers = ["板块/楼栋", "容积率", "建筑密度", "单位荷载 kg/㎡", "货梯", "净深", "净高", "结论"]
    write_matrix(
        ws,
        headers,
        [
            ["浦东金桥（待填）", "", "", "", "", "", "", ""],
            ["浦东张江（待填）", "", "", "", "", "", "", ""],
            ["杨浦滨江（如需）", "", "", "", "", "", "", ""],
            ["复兴岛（如需）", "", "", "", "", "", "", ""],
        ],
    )
    set_widths(ws, [22, 12, 12, 18, 14, 12, 12, 28])
    note_row = 9
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=8)
    ws.cell(note_row, 1, "备忘：" + "；".join(C.PHASE2_SITE))
    ws.cell(note_row, 1).font = font(9, color=GREY)
    ws.cell(note_row, 1).alignment = align("left")
    ws.row_dimensions[note_row].height = 80

    # 风险
    ws = wb.create_sheet("风险与支持")
    freeze_title(ws, "风险与需要公司支持的事项", "岗位价值取决于补位是否干净、口径是否经得起问", 3)
    write_matrix(ws, C.RISKS[0], C.RISKS[1:])
    set_widths(ws, [22, 36, 50])
    start = 4 + len(C.RISKS[1:]) + 2
    ws.cell(start, 1, "需要公司打开的门")
    ws.cell(start, 1).font = font(12, bold=True, color=NAVY)
    for i, item in enumerate(C.SUPPORT_NEEDED):
        ws.cell(start + 1 + i, 1, i + 1)
        ws.merge_cells(start_row=start + 1 + i, start_column=2, end_row=start + 1 + i, end_column=3)
        ws.cell(start + 1 + i, 2, item)
        for c in range(1, 4):
            ws.cell(start + 1 + i, c).font = font(10)
            ws.cell(start + 1 + i, c).border = thin
            ws.cell(start + 1 + i, c).alignment = align()
            ws.cell(start + 1 + i, c).fill = fill(WHITE if i % 2 == 0 else MIST)
        ws.row_dimensions[start + 1 + i].height = 28

    wb.save(OUT_FILE)
    print(f"已生成 {OUT_FILE}  工作表 {wb.sheetnames}")
    return OUT_FILE


if __name__ == "__main__":
    build()
