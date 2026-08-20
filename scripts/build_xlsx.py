# -*- coding: utf-8 -*-
"""生成 Excel：课题总览、技能匹配、选题意向、编组、时间表、群发短稿。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import content as C

OUT = Path(__file__).resolve().parents[1] / "exports"
OUT.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT / "AIDEN_西电课题选题与编组表.xlsx"

NAVY = "0B3D5C"
TEAL = "1A7A6D"
GOLD = "C4A35A"
LIGHT = "E8F3F1"
WHITE = "FFFFFF"
INK = "1A2A33"
THIN = Border(
    left=Side(style="thin", color="C5D5D1"),
    right=Side(style="thin", color="C5D5D1"),
    top=Side(style="thin", color="C5D5D1"),
    bottom=Side(style="thin", color="C5D5D1"),
)


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def font(bold=False, color=INK, size=11, name="微软雅黑"):
    return Font(name=name, bold=bold, color=color, size=size)


def align(wrap=True, h="left", v="center"):
    return Alignment(wrap_text=wrap, horizontal=h, vertical=v, indent=0)


def style_header(ws, row, cols, bg=NAVY):
    for col in range(1, cols + 1):
        cell = ws.cell(row, col)
        cell.fill = fill(bg)
        cell.font = font(bold=True, color=WHITE, size=11)
        cell.alignment = align(h="center")
        cell.border = THIN


def style_body(ws, start, end, cols):
    for r in range(start, end + 1):
        for c in range(1, cols + 1):
            cell = ws.cell(r, c)
            cell.font = font(size=10)
            cell.alignment = align(v="center")
            cell.border = THIN
            cell.fill = fill(WHITE if (r - start) % 2 == 0 else LIGHT)


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def freeze_header(ws):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.print_title_rows = "1:1"
    ws.page_setup.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def add_title_sheet_banner(ws, title, subtitle):
    ws.merge_cells("A1:F1")
    ws["A1"] = title
    ws["A1"].font = font(bold=True, color=WHITE, size=16)
    ws["A1"].fill = fill(NAVY)
    ws["A1"].alignment = align(h="left")
    ws.merge_cells("A2:F2")
    ws["A2"] = subtitle
    ws["A2"].font = font(color=INK, size=11)
    ws["A2"].fill = fill(LIGHT)
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 22


def build():
    wb = Workbook()

    # 0 封面说明
    ws = wb.active
    ws.title = "使用说明"
    add_title_sheet_banner(
        ws,
        f"{C.TITLE}  ·  {C.VERSION}",
        f"计划 {C.RELEASE_DATE} 发给西电初筛同学  |  {C.ORG_LINE}",
    )
    notes = [
        ["表", "用途", "谁来填"],
        ["课题总览", "十条题目、状态、人数、一句话", "只读"],
        ["技能匹配", "专业与技能，方便老师编组", "只读"],
        ["选题意向", "学生填志愿，可复印多行", "学生"],
        ["项目组编组", "老师确认名单与组长", "老师"],
        ["时间表", "发布到结题节点", "只读"],
        ["群发短稿", "可直接复制到微信群", "老师粘贴"],
    ]
    start = 4
    for i, row in enumerate(notes):
        for c, val in enumerate(row, 1):
            ws.cell(start + i, c, val)
    style_header(ws, start, 3)
    style_body(ws, start + 1, start + len(notes) - 1, 3)
    ws.merge_cells("A12:F14")
    ws["A12"] = (
        "编组规则：每人 1 个第一志愿 + 1 个备选。超额先看备选，再看技能，最后抽签。"
        "空题可合并或缓开。AIDEN 侧不在学生群单独招生。"
    )
    ws["A12"].alignment = align()
    ws["A12"].font = font(size=11)
    set_widths(ws, [16, 42, 18, 18, 18, 18])
    ws.row_dimensions[12].height = 48
    ws.sheet_view.showGridLines = False

    # 1 总览
    ws = wb.create_sheet("课题总览")
    headers = ["编号", "课题", "赛道", "状态", "人数", "周期", "一句话", "结题交付"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    for t in C.TOPICS:
        ws.append([t["id"], t["name"], t["track"], t["status"], t["quota"], t["weeks"], t["one_liner"], t["deliverables"]])
    style_body(ws, 2, 1 + len(C.TOPICS), len(headers))
    for r in range(2, 2 + len(C.TOPICS)):
        ws.row_dimensions[r].height = 48
    set_widths(ws, [8, 42, 14, 22, 12, 12, 48, 36])
    freeze_header(ws)

    # 2 技能
    ws = wb.create_sheet("技能匹配")
    headers = ["编号", "适合专业", "关键技能", "角色建议", "已经开工的部分"]
    ws.append(headers)
    style_header(ws, 1, len(headers), bg=TEAL)
    for t in C.TOPICS:
        ws.append([t["id"], t["major"], t["skills"], t["roles"], t["started"]])
    style_body(ws, 2, 1 + len(C.TOPICS), len(headers))
    for r in range(2, 2 + len(C.TOPICS)):
        ws.row_dimensions[r].height = 56
    set_widths(ws, [8, 28, 40, 32, 55])
    freeze_header(ws)

    # 3 选题意向
    ws = wb.create_sheet("选题意向")
    headers = [
        "姓名", "学院专业", "年级", "手机或微信",
        "第一志愿", "备选", "角色意向", "每周小时",
        "能否短期赴沪", "我能贡献什么（≤80字）", "备注",
    ]
    ws.append(headers)
    style_header(ws, 1, len(headers), bg=NAVY)
    for i in range(40):
        ws.append([""] * len(headers))
        ws.row_dimensions[2 + i].height = 22
    style_body(ws, 2, 41, len(headers))
    ids = ",".join(t["id"] for t in C.TOPICS)
    dv1 = DataValidation(type="list", formula1=f'"{ids}"', allow_blank=True)
    dv2 = DataValidation(type="list", formula1=f'"{ids}"', allow_blank=True)
    dv3 = DataValidation(type="list", formula1='"组长,开发,主笔,数据,演示,均可"', allow_blank=True)
    dv4 = DataValidation(type="list", formula1='"能,不能,寒暑假可以"', allow_blank=True)
    for dv, col in ((dv1, "E"), (dv2, "F"), (dv3, "G"), (dv4, "I")):
        dv.add(f"{col}2:{col}41")
        ws.add_data_validation(dv)
    set_widths(ws, [10, 18, 10, 16, 12, 12, 12, 12, 16, 40, 16])
    freeze_header(ws)
    ws.freeze_panes = "A2"

    # 4 编组
    ws = wb.create_sheet("项目组编组")
    headers = ["编号", "课题", "计划人数", "组长", "组员", "状态", "老师确认", "备注"]
    ws.append(headers)
    style_header(ws, 1, len(headers), bg=TEAL)
    for t in C.TOPICS:
        ws.append([t["id"], t["name"], t["quota"], "", "", "待编组", "", ""])
    style_body(ws, 2, 1 + len(C.TOPICS), len(headers))
    for r in range(2, 2 + len(C.TOPICS)):
        ws.row_dimensions[r].height = 28
    dv = DataValidation(type="list", formula1='"待编组,已满员,缓开,合并"', allow_blank=True)
    dv.add("F2:F11")
    ws.add_data_validation(dv)
    set_widths(ws, [8, 42, 12, 12, 36, 12, 12, 20])
    freeze_header(ws)

    # 5 时间表
    ws = wb.create_sheet("时间表")
    headers = ["节点", "做什么"]
    ws.append(headers)
    style_header(ws, 1, 2, bg=NAVY)
    for row in C.RHYTHM:
        ws.append(row)
    style_body(ws, 2, 1 + len(C.RHYTHM), 2)
    for r in range(2, 2 + len(C.RHYTHM)):
        ws.row_dimensions[r].height = 28
    set_widths(ws, [18, 70])
    freeze_header(ws)

    # 6 群发短稿（纯文本，便于复制）
    ws = wb.create_sheet("群发短稿")
    ws["A1"] = "微信群发布稿（全选 A 列复制）"
    ws["A1"].font = font(bold=True, color=WHITE, size=14)
    ws["A1"].fill = fill(NAVY)
    ws.merge_cells("A1:B1")
    lines = wechat_lines()
    for i, line in enumerate(lines, 3):
        ws.cell(i, 1, line)
        ws.cell(i, 1).alignment = align(wrap=True)
        ws.cell(i, 1).font = font(size=11)
        ws.row_dimensions[i].height = 18 if line else 8
    set_widths(ws, [90, 20])
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 26

    wb.save(OUT_FILE)
    print(f"已生成 {OUT_FILE}")
    return OUT_FILE


def wechat_lines():
    from build_wechat import render_lines
    return render_lines()


if __name__ == "__main__":
    build()
