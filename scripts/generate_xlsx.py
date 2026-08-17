#!/usr/bin/env python3
"""生成伤情、费用、待办与口径工作簿。"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from content import (
    ACCIDENT,
    ACTIONS_72H,
    ASR_CORRECTIONS,
    CALLS,
    CLOUD_FILM,
    CT_REPORTS,
    DISABILITY,
    INJURY_NOT_THIS_ACCIDENT,
    INJURY_THIS_ACCIDENT,
    QINGXIAN,
    TALKING_POINTS,
)

NAVY = "0B2F5B"
GOLD = "C4A35A"
LIGHT = "F3F6FA"
RED = "A63D2F"
GREEN = "2F6B4F"
WHITE = "FFFFFF"
THIN = Border(
    left=Side(style="thin", color="D0D7DE"),
    right=Side(style="thin", color="D0D7DE"),
    top=Side(style="thin", color="D0D7DE"),
    bottom=Side(style="thin", color="D0D7DE"),
)


def header_font():
    return Font(name="微软雅黑", size=11, bold=True, color=WHITE)


def title_font():
    return Font(name="微软雅黑", size=14, bold=True, color=NAVY)


def body_font(bold=False, color="1F2A37"):
    return Font(name="微软雅黑", size=10, bold=bold, color=color)


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def wrap():
    return Alignment(wrap_text=True, vertical="center")


def style_header(ws, row, cols):
    for col in range(1, cols + 1):
        cell = ws.cell(row, col)
        cell.font = header_font()
        cell.fill = fill(NAVY)
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = THIN


def style_rows(ws, start, end, cols):
    for r in range(start, end + 1):
        for c in range(1, cols + 1):
            cell = ws.cell(r, c)
            cell.font = body_font()
            cell.alignment = wrap()
            cell.border = THIN
            if r % 2 == 0:
                cell.fill = fill(LIGHT)


def widths(ws, values):
    for i, w in enumerate(values, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def freeze(ws):
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5
    ws.print_title_rows = "1:2"


def write_title(ws, text, cols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=cols)
    cell = ws.cell(1, 1, text)
    cell.font = title_font()
    cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 22


def build_workbook(output_path: Path) -> None:
    wb = Workbook()

    # 1 伤情鉴定
    ws = wb.active
    ws.title = "伤情鉴定"
    write_title(ws, "齐鲁医院云影像 · 谁受伤、伤了什么（2026-08-17 核对）", 6)
    headers = ["日期", "检查号", "项目 / 科室", "影像所见（报告原文）", "影像诊断（报告原文）", "本次事故？"]
    for i, h in enumerate(headers, 1):
        ws.cell(2, i, h)
    style_header(ws, 2, 6)
    for i, r in enumerate(CT_REPORTS, 3):
        ws.cell(i, 1, r["date"])
        ws.cell(i, 2, r["no"])
        ws.cell(i, 3, f"{r['item']}\n{r['place']} · {r['status']}")
        ws.cell(i, 4, r["findings"])
        ws.cell(i, 5, "\n".join(r["diagnosis"]))
        ws.cell(i, 6, "是，计入本次" if r["accident_related"] else "否，退变/陈旧，不计入评残")
        if not r["accident_related"]:
            ws.cell(i, 6).font = body_font(True, GREEN)
        else:
            ws.cell(i, 6).font = body_font(True, RED)
    style_rows(ws, 3, 5, 6)
    ws.cell(7, 1, f"患者：{CLOUD_FILM['patient_display']} / {CLOUD_FILM['sex']} / {CLOUD_FILM['age']}岁　　医院：{CLOUD_FILM['hospital']}")
    ws.merge_cells("A7:F7")
    ws.cell(7, 1).font = body_font(True)
    ws.cell(8, 1, CLOUD_FILM["disclaimer"] + " 这是影像诊断，不是伤残鉴定。")
    ws.merge_cells("A8:F8")
    widths(ws, [14, 16, 28, 42, 36, 22])
    ws.row_dimensions[3].height = 72
    ws.row_dimensions[4].height = 56
    ws.row_dimensions[5].height = 48
    freeze(ws)

    # 2 伤残情况
    ws = wb.create_sheet("伤残情况")
    write_title(ws, "伤残：现在不能定级。只准备材料。", 3)
    headers = ["问题", "结论", "家属立刻怎么做"]
    for i, h in enumerate(headers, 1):
        ws.cell(2, i, h)
    style_header(ws, 2, 3)
    rows = [
        ["谁可能评残", DISABILITY["who"], "只为伤者本人建病历档案，不要给别人做伤残材料"],
        ["现在有没有残级", DISABILITY["now"], "对外禁止说「已经几级」"],
        ["什么时候评", DISABILITY["when"], "先治病，3 个月窗口前不要催鉴定"],
        ["用哪套标准", DISABILITY["standard"], "交通事故用人体损伤致残程度分级，不要套工伤"],
        ["评哪些伤", DISABILITY["what_counts"], "踝关节功能+行走；右足骨刺删除"],
        ["现在能否报价", DISABILITY["cannot_grade"], "面谈不谈残级、不签总包死价"],
    ]
    for i, row in enumerate(rows, 3):
        for c, val in enumerate(row, 1):
            ws.cell(i, c, val)
    style_rows(ws, 3, 8, 3)
    ws.cell(10, 1, "评残前材料清单")
    ws.merge_cells("A10:C10")
    ws.cell(10, 1).font = body_font(True, NAVY)
    for i, item in enumerate(DISABILITY["prep"], 11):
        ws.cell(i, 1, i - 10)
        ws.cell(i, 2, item)
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=3)
    style_rows(ws, 11, 14, 3)
    widths(ws, [18, 70, 36])
    for r in range(3, 9):
        ws.row_dimensions[r].height = 48
    freeze(ws)

    # 3 外伤 vs 退变
    ws = wb.create_sheet("外伤与退变对照")
    write_title(ws, "通话口述 vs 医院 CT：必须改口的清单", 3)
    headers = ["来源", "内容", "处理"]
    for i, h in enumerate(headers, 1):
        ws.cell(2, i, h)
    style_header(ws, 2, 3)
    r = 3
    ws.cell(r, 1, "CT：本次外伤")
    ws.cell(r, 2, "\n".join(f"• {x}" for x in INJURY_THIS_ACCIDENT))
    ws.cell(r, 3, "对外、对保险、对鉴定都用这套")
    ws.row_dimensions[r].height = 110
    r = 4
    ws.cell(r, 1, "CT：非本次外伤")
    ws.cell(r, 2, "\n".join(f"• {x}" for x in INJURY_NOT_THIS_ACCIDENT))
    ws.cell(r, 3, "不要写成跟骨粉碎骨折")
    ws.row_dimensions[r].height = 70
    r = 5
    for i, (wrong, right) in enumerate(ASR_CORRECTIONS, 5):
        ws.cell(i, 1, "通话须改口")
        ws.cell(i, 2, wrong)
        ws.cell(i, 3, right)
        ws.row_dimensions[i].height = 40
    style_rows(ws, 3, 8, 3)
    widths(ws, [18, 55, 55])
    freeze(ws)

    # 4 费用台账
    ws = wb.create_sheet("费用台账")
    write_title(ws, "用发票重做台账。8月15日「约7万」和8月16日「不到3.5万」冲突，对外前必须核死。", 8)
    headers = ["日期", "项目", "金额（元）", "已付/未付", "付款人", "有无发票", "票据号", "备注"]
    for i, h in enumerate(headers, 1):
        ws.cell(2, i, h)
    style_header(ws, 2, 8)
    presets = [
        ["2026-08-14", "急诊右小腿CT（10008056847）", "", "已付待核", "家属", "待贴", "", "齐鲁医院"],
        ["2026-08-14", "急诊右足CT（10008056848）", "", "已付待核", "家属", "待贴", "", "齐鲁医院"],
        ["2026-08-15", "住院右踝CT术后复查（10008059257）", "", "已付待核", "家属", "待贴", "", "手足与显微重建外科"],
        ["2026-08-15", "右踝骨折内外固定手术", "", "已付待核", "家属", "待贴", "", "以手术记录为准"],
        ["", "住院押金/预缴", "", "已付待核", "家属", "待贴", "", "8/16口径累计不到35000，须对账单"],
        ["", "药品（院内）", "", "已付待核", "家属", "待贴", "", ""],
        ["", "护工费", "", "已付待核", "家属", "待贴", "", "不确定项，必须合同+发票"],
        ["", "护理耗材（垫、便盆等）", "", "已付待核", "家属", "待贴", "", ""],
        ["", "专业转运/救护车", "", "未发生或待核", "", "待贴", "", "责任方承担，不自掏不留票"],
        ["", "雇主垫付（通话口径）", 20000, "待核是否属实", "货主?", "待核", "", "8/15口述，未与3.5万口径对上"],
        ["", "未结工钱（通话口径）", 5000, "未付", "货主", "无", "", "不要直接冲抵事故赔偿"],
        ["", "责任方已付", 0, "未付", "骑手/美团", "无", "", "截至8/16通话称一分未付"],
    ]
    for i, row in enumerate(presets, 3):
        for c, val in enumerate(row, 1):
            ws.cell(i, c, val)
        ws.cell(i, 3).number_format = "#,##0.00"
    style_rows(ws, 3, 14, 8)
    ws.cell(16, 1, "合计（仅填写了数字的行）")
    ws.cell(16, 3, "=SUM(C3:C14)")
    ws.cell(16, 3).number_format = "#,##0.00"
    ws.cell(16, 1).font = body_font(True)
    ws.cell(16, 3).font = body_font(True, RED)
    dv = DataValidation(type="list", formula1='"已付,未付,已付待核,未发生或待核"', allow_blank=True)
    dv.add("D3:D30")
    ws.add_data_validation(dv)
    dv2 = DataValidation(type="list", formula1='"有,无,待贴,待核"', allow_blank=True)
    dv2.add("F3:F30")
    ws.add_data_validation(dv2)
    widths(ws, [14, 34, 14, 16, 14, 12, 14, 36])
    freeze(ws)
    ws.auto_filter.ref = "A2:H30"

    # 5 待办
    ws = wb.create_sheet("72小时待办")
    write_title(ws, "本周只办治疗不断档、证据固定、保险进群。不谈残级。", 5)
    headers = ["序号", "谁", "做什么", "为什么", "状态"]
    for i, h in enumerate(headers, 1):
        ws.cell(2, i, h)
    style_header(ws, 2, 5)
    for i, a in enumerate(ACTIONS_72H, 3):
        ws.cell(i, 1, i - 2)
        ws.cell(i, 2, a["who"])
        ws.cell(i, 3, a["what"])
        ws.cell(i, 4, a["why"])
        ws.cell(i, 5, "未完成")
        ws.row_dimensions[i].height = 36
    style_rows(ws, 3, 8, 5)
    dv = DataValidation(type="list", formula1='"未完成,进行中,已完成,律师跟进"', allow_blank=False)
    dv.add("E3:E20")
    ws.add_data_validation(dv)
    widths(ws, [8, 18, 55, 36, 14])
    freeze(ws)

    # 6 口径
    ws = wb.create_sheet("沟通口径卡")
    write_title(ws, "面谈只说这些。全责、残级、跟骨粉碎骨折，全部停用。", 2)
    headers = ["场合", "原话/要点"]
    for i, h in enumerate(headers, 1):
        ws.cell(2, i, h)
    style_header(ws, 2, 2)
    r = 3
    mapping = [
        ("对交警", TALKING_POINTS["to_police"]),
        ("对骑手/平台", TALKING_POINTS["to_rider"]),
        ("对医院", TALKING_POINTS["to_hospital"]),
        ("绝对不要说", TALKING_POINTS["do_not_say"]),
    ]
    for title, items in mapping:
        for item in items:
            ws.cell(r, 1, title)
            ws.cell(r, 2, item)
            if title == "绝对不要说":
                ws.cell(r, 2).font = body_font(True, RED)
            r += 1
    style_rows(ws, 3, r - 1, 2)
    widths(ws, [18, 90])
    freeze(ws)

    # 7 责任路径
    ws = wb.create_sheet("责任与程序")
    write_title(ws, "交警视频已排除全责。程序选择不要赌全责。", 3)
    headers = ["事项", "交警口径", "家属对策"]
    for i, h in enumerate(headers, 1):
        ws.cell(2, i, h)
    style_header(ws, 2, 3)
    rows = [
        ["地点", ACCIDENT["place"], "材料里写抚顺路市政道路，不要写成市场内部"],
        ["过程", ACCIDENT["process"], "以监控为准，不补充猜测"],
        ["三轮过错", "；".join(ACCIDENT["tri_faults"]), "承认监控里看得到的问题，不争口头全责"],
        ["骑手过错", "；".join(ACCIDENT["rider_faults"]), "要求主责，保底同等"],
        ["责任范围", ACCIDENT["police_range"], "停止使用「对方全责」"],
        ["简易程序", "不审资质，快出认定书", "可谈，但先确认平台保险能赔"],
        ["一般程序", "鉴定约一个多月；无牌/可能无证要处罚", "除非保险拒赔或对方不谈，否则不主动升级"],
        ["交警赔钱", "只出认定书，不参与金额", "金额找保险和法院，别逼交警"],
    ]
    for i, row in enumerate(rows, 3):
        for c, val in enumerate(row, 1):
            ws.cell(i, c, val)
        ws.row_dimensions[i].height = 48
    style_rows(ws, 3, 10, 3)
    widths(ws, [16, 55, 40])
    freeze(ws)

    # 8 录音索引
    ws = wb.create_sheet("录音来源")
    write_title(ws, "得到大脑 10 份录音。与 CT 冲突的以 CT 为准。", 4)
    headers = ["日期", "标题", "时长", "本表如何使用"]
    for i, h in enumerate(headers, 1):
        ws.cell(2, i, h)
    style_header(ws, 2, 4)
    for i, c in enumerate(CALLS, 3):
        ws.cell(i, 1, c["date"])
        ws.cell(i, 2, c["title"])
        ws.cell(i, 3, c["mins"])
        ws.cell(i, 4, c["use"])
    style_rows(ws, 3, 12, 4)
    widths(ws, [14, 36, 10, 50])
    freeze(ws)

    # 9 轻弦附录
    ws = wb.create_sheet("附录轻弦洽谈")
    write_title(ws, QINGXIAN["title"], 2)
    headers = ["序号", "要点"]
    for i, h in enumerate(headers, 1):
        ws.cell(2, i, h)
    style_header(ws, 2, 2)
    for i, p in enumerate(QINGXIAN["points"], 3):
        ws.cell(i, 1, i - 2)
        ws.cell(i, 2, p)
    style_rows(ws, 3, 6, 2)
    widths(ws, [8, 90])
    freeze(ws)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


if __name__ == "__main__":
    root = Path(__file__).resolve().parents[1]
    build_workbook(root / "deliverables" / "青岛红枫路交通事故_伤情伤残与行动表_20260817.xlsx")
