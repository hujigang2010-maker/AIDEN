# -*- coding: utf-8 -*-
"""生成《泰隆银行·科创企业成长伙伴计划（2027）》资源排期与预算表。"""

import os
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import partner2027_data as D

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "deliverables", "泰隆银行科创企业成长伙伴计划2027_资源排期与预算表.xlsx")

NAVY = "0C1E33"
GOLD = "A6844A"
CREAM = "F6F3EE"
PAPER = "FFFCF8"
INK = "1B2836"
MUTED = "5C6874"
RED = "8E2F2F"
RED_SOFT = "F8EEEC"
GOLD_SOFT = "F4ECDC"
NAVY_SOFT = "E7EEF4"
YELLOW = "FFF4D6"
GREEN_SOFT = "E7F2EC"
LINE = "E4DBCE"
WHITE = "FFFFFF"

thin = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
font_name = "微软雅黑"


def font(size=10, bold=False, color=INK):
    return Font(name=font_name, size=size, bold=bold, color=color)


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def align(h="left", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def apply_header_footer(ws, title):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.45
    ws.page_margins.right = 0.45
    ws.page_margins.top = 0.7
    ws.page_margins.bottom = 0.55
    ws.page_margins.header = 0.28
    ws.page_margins.footer = 0.25
    ws.oddHeader.left.text = f"&8&K5C6874{title}"
    ws.oddHeader.right.text = "&8&K5C6874洽谈稿 · 待正式确认"
    ws.oddFooter.left.text = "&8&K5C6874泰隆银行·科创企业成长伙伴计划（2027）"
    ws.oddFooter.right.text = "&8&K5C6874第 &P 页"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.sheet_properties.tabColor = NAVY
    ws.page_setup.pageOrder = "downThenOver"
    ws.print_options.horizontalCentered = True


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write(ws, row, col, value, size=10, bold=False, color=INK, bg=None, h="left", v="center", fmt=None, border=True):
    cell = ws.cell(row, col, value)
    cell.font = font(size, bold, color)
    cell.alignment = align(h, v)
    if bg:
        cell.fill = fill(bg)
    if border:
        cell.border = thin
    if fmt:
        cell.number_format = fmt
    return cell


def merge(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


def title_block(ws, text, subtitle, cols):
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 32
    write(ws, 1, 1, "泰隆银行 · 科创企业成长伙伴计划（2027）", 11, True, GOLD, NAVY, "left")
    for c in range(2, cols + 1):
        write(ws, 1, c, "", 11, False, WHITE, NAVY)
    merge(ws, 1, 1, 1, cols)
    write(ws, 2, 1, text, 16, True, NAVY, CREAM, "left")
    for c in range(2, cols + 1):
        write(ws, 2, c, "", 16, True, NAVY, CREAM)
    merge(ws, 2, 1, 2, cols)
    ws.row_dimensions[3].height = 36
    write(ws, 3, 1, subtitle, 9, False, MUTED, PAPER, "left")
    for c in range(2, cols + 1):
        write(ws, 3, c, "", 9, False, MUTED, PAPER)
    merge(ws, 3, 1, 3, cols)
    ws.auto_filter.ref = None
    ws.freeze_panes = "A5"
    ws.print_title_rows = "1:3"
    ws.page_setup.fitToHeight = 1
    ws.sheet_view.zoomScale = 110


def header_row(ws, row, labels, bg=NAVY):
    ws.row_dimensions[row].height = 22
    for i, label in enumerate(labels, 1):
        write(ws, row, i, label, 10, True, WHITE, bg, "center")


def paint_row(ws, row, ncol, bg):
    for c in range(1, ncol + 1):
        cell = ws.cell(row, c)
        if cell.fill is None or cell.fill.fgColor is None or cell.fill.fgColor.rgb in (None, "00000000"):
            cell.fill = fill(bg)
        cell.border = thin
        cell.alignment = align("left" if c > 1 else "center", "center")


def sheet_note(ws):
    ws.sheet_view.view = "pageBreakPreview"
    ws.sheet_view.view = "normal"


def build_cover(wb):
    ws = wb.active
    ws.title = "00说明"
    apply_header_footer(ws, "使用说明")
    cols = 6
    set_widths(ws, [24, 28, 28, 26, 24, 32])
    title_block(ws, "资源排期与预算表", "与商务沟通稿、简报配套。黄色单元格可以改，金额会随公式重算。编制日期 " + D.DATE, cols)
    ws.row_dimensions[4].height = 20
    write(ws, 4, 1, "本表怎么用", 12, True, WHITE, NAVY)
    for c in range(2, cols + 1):
        write(ws, 4, c, "", 12, True, WHITE, NAVY)
    merge(ws, 4, 1, 4, cols)

    notes = [
        ("给谁", "可随商务沟通稿一并提交给泰隆银行。资源状态、预算和排期都是洽谈口径。"),
        ("黄格", "09预算报价中的场次、单价和年度建议可以改。不要手改金额列，金额由公式生成。"),
        ("资源", "04参访资源列出A/B/C三类。是否已确认接待默认为“否”，约到之后再改。"),
        ("预算", "参访15万—20万元与路演14万元分开。页面下方的并列合计只供内部对照，不直接报批。"),
        ("绩效", "11绩效跟踪预填了计划场次。会后填写有效企业、需求、对接、满意度和授权。"),
        ("边界", "不把口头经费写成批复；不把名单写成企业已同意；未授权不使用高校品牌；不承诺入学；不对个人按结果付佣。"),
    ]
    for i, (k, v) in enumerate(notes):
        r = 5 + i
        ws.row_dimensions[r].height = 32
        write(ws, r, 1, k, 11, True, NAVY, GOLD_SOFT, "center")
        write(ws, r, 2, v, 10, False, INK, PAPER)
        for c in range(3, cols + 1):
            write(ws, r, c, "", 10, False, INK, PAPER)
        merge(ws, r, 2, r, cols)

    ws.row_dimensions[12].height = 22
    write(ws, 12, 1, "工作表", 12, True, WHITE, NAVY)
    for c in range(2, cols + 1):
        write(ws, 12, c, "", 12, True, WHITE, NAVY)
    merge(ws, 12, 1, 12, cols)
    sheets = [
        ("01一页纸", "可打印的合作共识草案，附沟通意见栏"),
        ("02六方向", "六个合作方向、优先级和阶段"),
        ("03待确认", "提交前需要银行或相关单位确认的八件事"),
        ("04参访资源", "A/B类企业清单，可筛选；接待确认状态可改"),
        ("05参访年历", "2027年12场储备排期，可按月调整优先资源"),
        ("06参访流程", "半天标准流程、三条业务线和分工"),
        ("07路演年历", "4场重点 + 4场小型，含规模和内容"),
        ("08路演专题", "四个专题、筛选标准、三类项目库和议程"),
        ("09预算报价", "路演14万元公式，以及参访第一轮建议"),
        ("10执行计划", "2026年11月至2027年第一季度，及路演四阶段"),
        ("11绩效跟踪", "会后填写的跟踪表，计划场次已预填"),
    ]
    header_row(ws, 13, ["工作表", "内容", "", "", "", ""])
    merge(ws, 13, 2, 13, cols)
    for i, (name, desc) in enumerate(sheets):
        r = 14 + i
        ws.row_dimensions[r].height = 20
        cell = write(ws, r, 1, name, 10, True, NAVY, NAVY_SOFT if i % 2 == 0 else PAPER, "center")
        cell.hyperlink = f"#'{name}'!A1"
        cell.font = Font(name=font_name, size=10, bold=True, color="1A3A5C", underline="single")
        write(ws, r, 2, desc, 10, False, INK, NAVY_SOFT if i % 2 == 0 else PAPER)
        for c in range(3, cols + 1):
            write(ws, r, c, "", 10, False, INK, NAVY_SOFT if i % 2 == 0 else PAPER)
        merge(ws, r, 2, r, cols)

    ws.row_dimensions[26].height = 22
    write(ws, 26, 1, "资源结构（自动统计）", 12, True, WHITE, NAVY)
    for c in range(2, 4):
        write(ws, 26, c, "", 12, True, WHITE, NAVY)
    merge(ws, 26, 1, 26, 3)
    header_row(ws, 27, ["类别", "家数", "含义"])
    stats = [
        ("A类", '=COUNTIF(\'04参访资源\'!A5:A22,"A")', "重点推进，线索或可协助约访"),
        ("B类", '=COUNTIF(\'04参访资源\'!A5:A22,"B")', "历史参访资源，需重新协调"),
        ("合计", "=COUNTA('04参访资源'!B5:B22)", "A类与B类，不含C类方向"),
    ]
    for i, (k, formula, note) in enumerate(stats):
        r = 28 + i
        ws.row_dimensions[r].height = 20
        write(ws, r, 1, k, 10, True, NAVY, GOLD_SOFT if i < 2 else CREAM, "center")
        write(ws, r, 2, formula, 12, True, NAVY, PAPER, "center", fmt="0")
        write(ws, r, 3, note, 10, False, INK, PAPER)
    write(ws, 31, 1, "已确认接待", 10, True, RED, RED_SOFT, "center")
    write(ws, 31, 2, '=COUNTIF(\'04参访资源\'!I5:I22,"是")', 12, True, RED, PAPER, "center", fmt="0")
    write(ws, 31, 3, "约访成功后，把04表“是否已确认接待”改为“是”。", 10, False, INK, PAPER)
    for c in range(4, 7):
        write(ws, 31, c, "", 10, False, INK, PAPER)
    merge(ws, 31, 3, 31, 6)
    ws.row_dimensions[31].height = 24
    ws.auto_filter.ref = "A13:F24"
    ws.freeze_panes = "A5"
    ws.page_setup.fitToHeight = 1
    ws.print_area = "A1:F31"
    ws.oddHeader.left.text = "&8&K5C6874使用说明"


def build_onepager(wb):
    ws = wb.create_sheet("01一页纸")
    apply_header_footer(ws, "一页纸共识草案")
    set_widths(ws, [28, 72])
    title_block(ws, "合作共识草案", "本页供沟通确认。签署不构成合同、采购承诺或预算批准。", 2)
    rows = [
        ("项目名称", "泰隆银行·科创企业成长伙伴计划"),
        ("建议合作单位", D.PARTNERS),
        ("拟联动单位", D.LINKED + "。联合主办、品牌、课程和证书须另行授权。"),
        ("合作架构", "1个联合服务平台 + 3类资源（企业 / 教育 / 金融） + 6项业务"),
        ("服务闭环", D.LOOP),
        ("近期动作", "核实11月18日约100人科创路演是否举行，以及茶歇、会务、场地支持是否获批；2026年12月完成1场20—30人参访试点。"),
        ("2027年参访", "原则上每月1场，全年12场，单场20—30名企业高管，年度240—360人次。"),
        ("2027年路演", "4场重点产业路演 + 4场小型专项对接，与参访经费分别列支。"),
        ("第一轮经费", "参访年度建议15万—20万元，对应口头讨论的十几万至二十几万元，尚未批准。路演年度建议14万元，作为第二项独立方案。"),
        ("分工", "银行负责客户、金融服务、采购审批；联合会及项目执行方负责资源、策划、组织和总结。"),
        ("请确认", "预算量级、首场参访、承办主体与结算、年度活动计划、客户服务与资源权益。"),
        ("明确不做", "未授权不冠名；不承诺招生录取或特殊入学权益；不按贷款、开户或转介结果向个人支付佣金。"),
    ]
    header_row(ws, 4, ["事项", "草案表述"])
    for i, (k, v) in enumerate(rows):
        r = 5 + i
        ws.row_dimensions[r].height = 32
        write(ws, r, 1, k, 10, True, NAVY, GOLD_SOFT, "center")
        write(ws, r, 2, v, 10, False, INK, PAPER if i % 2 else CREAM)
    r = 18
    ws.row_dimensions[r].height = 22
    write(ws, r, 1, "沟通意见（不代表预算批准）", 11, True, WHITE, NAVY)
    write(ws, r, 2, "", 11, True, WHITE, NAVY)
    merge(ws, r, 1, r, 2)
    header_row(ws, 19, ["", "泰隆银行", ])
    # redo as 3 columns conceptually inside col2 by using a second small table starting row 19 with more columns
    # Rebuild signature on columns A-B is awkward. Use A as label and B as blank line text.
    sig = [
        ("银行确认人、职务、日期", ""),
        ("银行沟通意见", ""),
        ("执行方确认人、职务、日期", ""),
        ("执行方沟通意见", ""),
    ]
    for i, (k, v) in enumerate(sig):
        rr = 19 + i
        ws.row_dimensions[rr].height = 28
        write(ws, rr, 1, k, 10, True, NAVY, NAVY_SOFT, "center")
        write(ws, rr, 2, v, 10, False, MUTED, PAPER)
    ws.row_dimensions[23].height = 36
    write(ws, 23, 2, "意见栏留空，由当面沟通填写。", 9, False, MUTED, PAPER)
    ws.freeze_panes = "A5"
    ws.print_area = "A1:B23"
    ws.page_setup.fitToHeight = 1
    ws.page_setup.orientation = "landscape"


def build_directions(wb):
    ws = wb.create_sheet("02六方向")
    apply_header_footer(ws, "六个合作方向")
    set_widths(ws, [8, 28, 62, 12, 36])
    title_block(ws, "六个合作方向", "第一项、第二项为2026年年底至2027年第一季度重点。大型赞助不作为主要收入来源。", 5)
    header_row(ws, 4, ["序号", "方向", "内容", "优先级", "阶段"])
    for i, d in enumerate(D.DIRECTIONS):
        r = 5 + i
        ws.row_dimensions[r].height = 36
        bg = RED_SOFT if d["priority"] == "最高" else (GOLD_SOFT if d["priority"] in ("高", "中高") else PAPER)
        write(ws, r, 1, d["no"], 12, True, NAVY, bg, "center")
        write(ws, r, 2, d["name"], 11, True, NAVY, bg)
        write(ws, r, 3, d["detail"], 10, False, INK, bg)
        write(ws, r, 4, d["priority"], 11, True, RED if d["priority"] == "最高" else NAVY, bg, "center")
        write(ws, r, 5, d["window"], 10, False, INK, bg)
    ws.row_dimensions[12].height = 22
    write(ws, 12, 1, "六项业务", 11, True, WHITE, NAVY)
    for c in range(2, 6):
        write(ws, 12, c, "", 11, True, WHITE, NAVY)
    merge(ws, 12, 1, 12, 5)
    header_row(ws, 13, ["序号", "业务", "2027年位置", "", ""])
    positions = ["基础产品，第一轮主方案", "独立扩展，第二项方案", "月度小规模沙龙", "两条业务线，授权后开展", "结合国际业务，中高优先级", "视预算择机，不作为主要收入"]
    for i, name in enumerate(D.BUSINESSES):
        r = 14 + i
        ws.row_dimensions[r].height = 22
        bg = CREAM if i % 2 == 0 else PAPER
        write(ws, r, 1, i + 1, 10, True, NAVY, bg, "center")
        write(ws, r, 2, name, 10, True, INK, bg)
        write(ws, r, 3, positions[i], 10, False, INK, bg)
        write(ws, r, 4, "", 10, False, INK, bg)
        write(ws, r, 5, "", 10, False, INK, bg)
        merge(ws, r, 3, r, 5)
    ws.auto_filter.ref = "A4:E10"
    ws.print_area = "A1:E19"
    ws.page_setup.fitToHeight = 1


def build_confirm(wb):
    ws = wb.create_sheet("03待确认")
    apply_header_footer(ws, "待确认事项")
    set_widths(ws, [8, 42, 48, 28, 22, 16])
    title_block(ws, "提交前需要确认的事项", "状态截至2026年11月17日。核实完成后在“最新进展”中更新，不要覆盖“当前状态”。", 6)
    header_row(ws, 4, ["序号", "事项", "当前状态", "建议确认方", "时间", "最新进展"])
    for i, (item, status, who, when) in enumerate(D.CONFIRM):
        r = 5 + i
        ws.row_dimensions[r].height = 42
        bg = PAPER if i % 2 else CREAM
        write(ws, r, 1, i + 1, 11, True, NAVY, bg, "center")
        write(ws, r, 2, item, 10, True, NAVY, bg)
        write(ws, r, 3, status, 10, False, INK, bg)
        write(ws, r, 4, who, 10, False, INK, bg)
        write(ws, r, 5, when, 10, False, INK, bg, "center")
        write(ws, r, 6, "", 10, False, INK, YELLOW)
    dv = DataValidation(type="list", formula1='"未核实,核实中,已确认,暂不推进"', allow_blank=True)
    dv.error = "请从列表选择"
    dv.errorTitle = "进展"
    dv.add("F5:F12")
    ws.add_data_validation(dv)
    ws.auto_filter.ref = "A4:F12"
    ws.auto_filter.ref = "A4:F12"
    note_row = 14
    ws.row_dimensions[note_row].height = 36
    write(ws, note_row, 1, "使用", 10, True, NAVY, GOLD_SOFT, "center")
    write(ws, note_row, 2, "黄色“最新进展”供会后填写。11月18日路演和年度预算两项，建议在11月20日前有结论。", 10, False, INK, PAPER)
    for c in range(3, 7):
        write(ws, note_row, c, "", 10, False, INK, PAPER)
    merge(ws, note_row, 2, note_row, 6)
    ws.print_area = "A1:F14"
    ws.page_setup.fitToHeight = 1


def build_resources(wb):
    ws = wb.create_sheet("04参访资源")
    apply_header_footer(ws, "参访资源清单")
    headers = ["类别", "序号", "拟参访单位", "核心主题", "产业方向", "资源状态", "系列", "建议月份", "是否已确认接待", "备注"]
    widths = [10, 8, 38, 42, 18, 28, 24, 14, 18, 42]
    set_widths(ws, widths)
    title_block(ws, "2027年度科技企业标杆参访资源清单", "A类为重点推进，B类为历史资源。是否已确认接待默认为否。约访结果只改第9列。", len(headers))
    header_row(ws, 4, headers)
    for i, v in enumerate(D.VISITS):
        r = 5 + i
        ws.row_dimensions[r].height = 32
        bg = RED_SOFT if v["cat"] == "A" else NAVY_SOFT
        vals = [v["cat"], v["no"], v["org"], v["theme"], v["track"], v["status"], v["series"], v["month"], "否", v["note"]]
        for c, val in enumerate(vals, 1):
            bold = c in (1, 2, 3)
            write(ws, r, c, val, 10, bold, NAVY if c <= 3 else INK, YELLOW if c == 9 else bg,
                  "center" if c in (1, 2, 8, 9) else "left")
    last = 4 + len(D.VISITS)
    dv = DataValidation(type="list", formula1='"否,是"', allow_blank=False)
    dv.add(f"I5:I{last}")
    ws.add_data_validation(dv)
    ws.auto_filter.ref = f"A4:J{last}"
    # C class below
    start = last + 2
    ws.row_dimensions[start].height = 22
    write(ws, start, 1, "C类 · 拓展储备（尚未落实具体接待单位）", 11, True, WHITE, NAVY)
    for c in range(2, 11):
        write(ws, start, c, "", 11, True, WHITE, NAVY)
    merge(ws, start, 1, start, 10)
    header_row(ws, start + 1, ["类别", "序号", "产业领域", "拟开发资源", "参访价值", "年历位置", "", "", "", ""])
    for i, (name, kind, value, month) in enumerate(D.VISIT_DIRS):
        r = start + 2 + i
        ws.row_dimensions[r].height = 28
        bg = GOLD_SOFT
        vals = ["C", i + 1, name, kind, value, month, "", "", "否", "方向储备，不计入已承诺场次"]
        for c, val in enumerate(vals, 1):
            write(ws, r, c, val, 10, c <= 3, INK, bg, "center" if c in (1, 2, 6, 9) else "left")
    c_last = start + 1 + len(D.VISIT_DIRS)
    ws.auto_filter.ref = f"A4:J{last}"
    ws.auto_filter.ref = f"A4:J{c_last}"
    # Filtering including C header row in the middle is messy. Keep filter on A/B only.
    ws.auto_filter.ref = f"A4:J{last}"
    ws.print_area = f"A1:J{c_last}"
    ws.print_title_rows = "1:4"
    ws.page_setup.fitToHeight = 2
    ws.auto_filter.ref = f"A4:J{last}"
    ws.freeze_panes = "C5"


def build_calendar(wb):
    ws = wb.create_sheet("05参访年历")
    apply_header_footer(ws, "2027参访年历")
    set_widths(ws, [12, 36, 42, 28, 22])
    title_block(ws, "2027年参访排期建议", "储备方案。两家并列的月份，以先确认接待的单位为准。黄色列可按报名和预算修改。", 5)
    header_row(ws, 4, ["月份", "参访主题", "优先资源", "调整后资源（会后填写）", "状态"])
    for i, (month, theme, org) in enumerate(D.VISIT_CALENDAR):
        r = 5 + i
        ws.row_dimensions[r].height = 24
        bg = PAPER if i % 2 else CREAM
        write(ws, r, 1, month, 11, True, NAVY, bg, "center")
        write(ws, r, 2, theme, 10, False, INK, bg)
        write(ws, r, 3, org, 10, False, INK, bg)
        write(ws, r, 4, "", 10, False, INK, YELLOW)
        write(ws, r, 5, "待确认接待", 10, False, RED, bg, "center")
    dv = DataValidation(type="list", formula1='"待确认接待,已锁定,已完成,改期,取消"', allow_blank=False)
    dv.add("E5:E16")
    ws.add_data_validation(dv)
    ws.row_dimensions[18].height = 48
    write(ws, 18, 1, "说明", 10, True, NAVY, GOLD_SOFT, "center")
    write(ws, 18, 2, "魔力社区、魔丸熊、腾讯、霸王茶姬在资源清单中，未单列进本表，可作为替换场次写入“调整后资源”。1月、6月、8月为二选一或可并场。", 10, False, INK, PAPER)
    for c in range(3, 6):
        write(ws, 18, c, "", 10, False, INK, PAPER)
    merge(ws, 18, 2, 18, 5)
    ws.print_area = "A1:E18"
    ws.page_setup.fitToHeight = 1
    ws.freeze_panes = "A5"


def build_flow(wb):
    ws = wb.create_sheet("06参访流程")
    apply_header_footer(ws, "参访流程与分工")
    set_widths(ws, [22, 55, 40])
    title_block(ws, "标准参访流程与分工", "规模20—30人，约半天。晚餐、园区走访或融资对接另做方案。", 3)
    header_row(ws, 4, ["时间", "内容", "备注"])
    for i, (tm, act) in enumerate(D.VISIT_FLOW):
        r = 5 + i
        ws.row_dimensions[r].height = 22
        bg = PAPER if i % 2 else CREAM
        write(ws, r, 1, tm, 10, True, NAVY, bg, "center")
        write(ws, r, 2, act, 10, False, INK, bg)
        note = "需求收集，转化由银行评价" if "泰隆" in act else ""
        write(ws, r, 3, note, 10, False, MUTED, bg)
    ws.row_dimensions[12].height = 22
    write(ws, 12, 1, "三个系列", 11, True, WHITE, NAVY)
    write(ws, 12, 2, "", 11, True, WHITE, NAVY)
    write(ws, 12, 3, "", 11, True, WHITE, NAVY)
    merge(ws, 12, 1, 12, 3)
    header_row(ws, 13, ["系列", "优先资源", ""])
    for i, (name, body) in enumerate(D.SERIES):
        r = 14 + i
        ws.row_dimensions[r].height = 22
        write(ws, r, 1, name, 10, True, NAVY, GOLD_SOFT)
        write(ws, r, 2, body, 10, False, INK, PAPER)
        write(ws, r, 3, "", 10, False, INK, PAPER)
        merge(ws, r, 2, r, 3)
    ws.row_dimensions[18].height = 22
    write(ws, 18, 1, "分工", 11, True, WHITE, NAVY)
    for c in range(2, 4):
        write(ws, 18, c, "", 11, True, WHITE, NAVY)
    merge(ws, 18, 1, 18, 3)
    header_row(ws, 19, ["主体", "负责事项", ""])
    for i, (who, what) in enumerate(D.ROLES):
        r = 20 + i
        ws.row_dimensions[r].height = 36
        write(ws, r, 1, who, 10, True, NAVY, CREAM if i % 2 == 0 else PAPER)
        write(ws, r, 2, what, 10, False, INK, CREAM if i % 2 == 0 else PAPER)
        write(ws, r, 3, "", 10, False, INK, CREAM if i % 2 == 0 else PAPER)
        merge(ws, r, 2, r, 3)
    ws.row_dimensions[24].height = 22
    write(ws, 24, 1, "研修两条线", 11, True, WHITE, NAVY)
    for c in range(2, 4):
        write(ws, 24, c, "", 11, True, WHITE, NAVY)
    merge(ws, 24, 1, 24, 3)
    header_row(ws, 25, ["对象", "内容", ""])
    for i, (who, what) in enumerate(D.EDU_LINES):
        r = 26 + i
        ws.row_dimensions[r].height = 32
        write(ws, r, 1, who, 10, True, NAVY, GOLD_SOFT)
        write(ws, r, 2, what, 10, False, INK, PAPER)
        write(ws, r, 3, "", 10, False, INK, PAPER)
        merge(ws, r, 2, r, 3)
    ws.print_area = "A1:C27"
    ws.page_setup.fitToHeight = 1
    ws.freeze_panes = "A5"


def build_roadshow(wb):
    ws = wb.create_sheet("07路演年历")
    apply_header_footer(ws, "2027路演年历")
    set_widths(ws, [12, 46, 14, 16, 48, 18])
    title_block(ws, "2027年路演与对接安排", "6月、9月、12月用于复访、推进和复盘，不在本表强制加场。规模和项目数是组织目标。", 6)
    header_row(ws, 4, ["时间", "活动名称", "类型", "规模", "重点内容", "状态"])
    for i, (month, name, kind, scale, focus) in enumerate(D.ROADSHOWS):
        r = 5 + i
        ws.row_dimensions[r].height = 30
        bg = GOLD_SOFT if kind == "重点" else PAPER
        write(ws, r, 1, month, 11, True, NAVY, bg, "center")
        write(ws, r, 2, name, 10, True, NAVY, bg)
        write(ws, r, 3, "重点路演" if kind == "重点" else "小型对接", 10, True, RED if kind == "重点" else NAVY, bg, "center")
        write(ws, r, 4, scale, 10, False, INK, bg, "center")
        write(ws, r, 5, focus, 10, False, INK, bg)
        write(ws, r, 6, "计划", 10, False, MUTED, YELLOW, "center")
    dv = DataValidation(type="list", formula1='"计划,筹备中,已举办,改期,取消"', allow_blank=False)
    dv.add("F5:F12")
    ws.add_data_validation(dv)
    ws.row_dimensions[14].height = 22
    write(ws, 14, 1, "年度目标（组织工作目标，不是融资承诺）", 11, True, WHITE, NAVY)
    for c in range(2, 7):
        write(ws, 14, c, "", 11, True, WHITE, NAVY)
    merge(ws, 14, 1, 14, 6)
    header_row(ws, 15, ["指标", "年度计划", "", "", "", ""])
    for i, (k, v) in enumerate(D.ROADSHOW_TARGETS):
        r = 16 + i
        ws.row_dimensions[r].height = 20
        bg = CREAM if i % 2 == 0 else PAPER
        write(ws, r, 1, k, 10, True, NAVY, bg)
        write(ws, r, 2, v, 10, False, INK, bg, "center")
        for c in range(3, 7):
            write(ws, r, c, "", 10, False, INK, bg)
        merge(ws, r, 2, r, 6)
    ws.print_area = "A1:F21"
    ws.page_setup.fitToHeight = 1
    ws.freeze_panes = "A5"


def build_themes(wb):
    ws = wb.create_sheet("08路演专题")
    apply_header_footer(ws, "路演专题与流程")
    set_widths(ws, [24, 70])
    title_block(ws, "四个重点专题、筛选与议程", "重点场一般安排5—6家企业路演。授信和投资与否由银行、机构独立决定。", 2)
    r = 4
    for i, t in enumerate(D.THEMES):
        ws.row_dimensions[r].height = 22
        write(ws, r, 1, f"专题{i+1}  {t['name']}", 11, True, WHITE, NAVY)
        write(ws, r, 2, "", 11, True, WHITE, NAVY)
        merge(ws, r, 1, r, 2)
        fields = [("拟定名称", t["title"]), ("重点赛道", t["tracks"]), ("拟邀对象", t["guests"]), ("合作价值", t["value"])]
        for j, (k, v) in enumerate(fields):
            rr = r + 1 + j
            ws.row_dimensions[rr].height = 28
            write(ws, rr, 1, k, 10, True, NAVY, GOLD_SOFT, "center")
            write(ws, rr, 2, v, 10, False, INK, PAPER)
        r = r + 6
    header_row(ws, r, ["时间", "重点路演议程（60—80人）"])
    for i, (tm, act) in enumerate(D.ROADSHOW_FLOW):
        rr = r + 1 + i
        ws.row_dimensions[rr].height = 20
        bg = CREAM if i % 2 == 0 else PAPER
        write(ws, rr, 1, tm, 10, True, NAVY, bg, "center")
        write(ws, rr, 2, act, 10, False, INK, bg)
    r2 = r + 1 + len(D.ROADSHOW_FLOW) + 1
    ws.row_dimensions[r2].height = 22
    write(ws, r2, 1, "筛选与项目库", 11, True, WHITE, NAVY)
    write(ws, r2, 2, "", 11, True, WHITE, NAVY)
    merge(ws, r2, 1, r2, 2)
    for i, text in enumerate(D.SCREEN):
        rr = r2 + 1 + i
        ws.row_dimensions[rr].height = 20
        write(ws, rr, 1, f"条件{i+1}", 10, True, NAVY, NAVY_SOFT, "center")
        write(ws, rr, 2, text, 10, False, INK, PAPER)
    base = r2 + 1 + len(D.SCREEN)
    for i, (name, desc) in enumerate(D.POOLS):
        rr = base + i
        ws.row_dimensions[rr].height = 22
        write(ws, rr, 1, name, 10, True, NAVY, GOLD_SOFT, "center")
        write(ws, rr, 2, desc, 10, False, INK, PAPER)
    ws.print_area = f"A1:B{base + len(D.POOLS) - 1}"
    ws.page_setup.fitToHeight = 2
    ws.freeze_panes = "A4"


def build_budget(wb):
    ws = wb.create_sheet("09预算报价")
    ws.sheet_properties.tabColor = GOLD
    apply_header_footer(ws, "预算报价")
    set_widths(ws, [42, 16, 18, 18, 48])
    title_block(
        ws,
        "预算报价（谈判用，不是已确认采购价）",
        "黄色单元格为可调整输入。金额列为公式。参访与路演分别测算；下方并列合计不要直接作为报批总额。",
        5,
    )
    write(ws, 4, 1, "一、路演标准版（第二项独立方案）", 11, True, WHITE, NAVY)
    for c in range(2, 6):
        write(ws, 4, c, "", 11, True, WHITE, NAVY)
    merge(ws, 4, 1, 4, 5)
    header_row(ws, 5, ["项目", "数量", "单价（元）", "金额（元）", "说明"])
    for i, (name, qty, price, note) in enumerate(D.ROADSHOW_BUDGET):
        r = 6 + i
        ws.row_dimensions[r].height = 24
        write(ws, r, 1, name, 10, True, NAVY, PAPER)
        write(ws, r, 2, qty, 11, True, NAVY, YELLOW, "center", fmt="0")
        write(ws, r, 3, price, 11, False, INK, YELLOW, "right", fmt='¥#,##0')
        write(ws, r, 4, f"=B{r}*C{r}", 11, True, NAVY, CREAM, "right", fmt='¥#,##0')
        write(ws, r, 5, note, 9, False, MUTED, PAPER)
    ws.row_dimensions[9].height = 24
    write(ws, 9, 1, "路演年度合计", 11, True, WHITE, NAVY)
    write(ws, 9, 2, "—", 11, True, WHITE, NAVY, "center")
    write(ws, 9, 3, "", 11, True, WHITE, NAVY)
    write(ws, 9, 4, "=SUM(D6:D8)", 12, True, WHITE, NAVY, "right", fmt='¥#,##0')
    write(ws, 9, 5, "目标值 140,000 元。改黄格后以本格为准。", 9, False, WHITE, NAVY)

    write(ws, 11, 1, "二、参访年度（第一轮主方案）", 11, True, WHITE, NAVY)
    for c in range(2, 6):
        write(ws, 11, c, "", 11, True, WHITE, NAVY)
    merge(ws, 11, 1, 11, 5)
    inputs = [
        (12, "建议场次", 12, "0", "原则上每月1场"),
        (13, "单场服务费下限（元）", 5000, '¥#,##0', "只含策划、协调、组织执行"),
        (14, "单场服务费上限（元）", 10000, '¥#,##0', "不含企业接待、交通、专家、场地"),
        (15, "服务费合计下限", "=B12*B13", '¥#,##0', "公式"),
        (16, "服务费合计上限", "=B12*B14", '¥#,##0', "公式"),
        (17, "第一轮年度建议下限", 150000, '¥#,##0', "含资源维护及按实成本的建议盘子"),
        (18, "第一轮年度建议上限", 200000, '¥#,##0', "对应口头讨论的十几万至二十几万元"),
    ]
    write(ws, 12, 3, "数值", 10, True, WHITE, NAVY, "center")
    # row 12 is also first input. I overlapped. Let me not add a subheader and just use the inputs starting row 12.
    # I already wrote title at row 11 and then inputs include row 12. The header write at row 12 col 3 conflicts.
    # Remove that header write by not doing it - I'll overwrite properly in the loop only.
    for r, label, val, fmt, note in inputs:
        ws.row_dimensions[r].height = 22
        write(ws, r, 1, label, 10, True, NAVY, GOLD_SOFT)
        is_formula = isinstance(val, str) and str(val).startswith("=")
        write(ws, r, 2, val, 12, True, NAVY, CREAM if is_formula else YELLOW, "right", fmt=fmt)
        write(ws, r, 3, "", 10, False, INK, PAPER)
        write(ws, r, 4, "", 10, False, INK, PAPER)
        write(ws, r, 5, note, 9, False, MUTED, PAPER)
        merge(ws, r, 3, r, 4)
    ws.row_dimensions[19].height = 32
    write(ws, 19, 1, "口头讨论口径", 10, True, RED, RED_SOFT, "center")
    write(ws, 19, 2, "十几万至二十几万元", 11, True, RED, RED_SOFT, "center")
    write(ws, 19, 3, "11月6日提出，尚未审批，不能视为已批准预算。", 10, False, INK, RED_SOFT)
    write(ws, 19, 4, "", 10, False, INK, RED_SOFT)
    write(ws, 19, 5, "", 10, False, INK, RED_SOFT)
    merge(ws, 19, 3, 19, 5)

    write(ws, 21, 1, "三、若两项并列时的对照（不要直接报批）", 11, True, WHITE, RED)
    for c in range(2, 6):
        write(ws, 21, c, "", 11, True, WHITE, RED)
    merge(ws, 21, 1, 21, 5)
    write(ws, 22, 1, "对照下限 = 参访建议下限 + 路演合计", 10, True, NAVY, PAPER)
    write(ws, 22, 2, "=B17+D9", 12, True, RED, PAPER, "right", fmt='¥#,##0')
    write(ws, 22, 3, "", 10, False, INK, PAPER)
    write(ws, 22, 4, "", 10, False, INK, PAPER)
    write(ws, 22, 5, "仅供内部看盘子大小", 9, False, MUTED, PAPER)
    merge(ws, 22, 3, 22, 4)
    write(ws, 23, 1, "对照上限 = 参访建议上限 + 路演合计", 10, True, NAVY, PAPER)
    write(ws, 23, 2, "=B18+D9", 12, True, RED, PAPER, "right", fmt='¥#,##0')
    write(ws, 23, 3, "", 10, False, INK, PAPER)
    write(ws, 23, 4, "", 10, False, INK, PAPER)
    write(ws, 23, 5, "银行若只有一个年度预算，应重新分配", 9, False, MUTED, PAPER)
    merge(ws, 23, 3, 23, 4)
    ws.row_dimensions[22].height = 22
    ws.row_dimensions[23].height = 22

    write(ws, 25, 1, "报价三档（谈判区间）", 11, True, WHITE, NAVY)
    for c in range(2, 6):
        write(ws, 25, c, "", 11, True, WHITE, NAVY)
    merge(ws, 25, 1, 25, 5)
    header_row(ws, 26, ["合作模式", "建议费用", "", "", "收益方式"])
    for i, (mode, fee, how) in enumerate(D.PRICING):
        r = 27 + i
        ws.row_dimensions[r].height = 22
        bg = CREAM if i % 2 == 0 else PAPER
        write(ws, r, 1, mode, 10, True, NAVY, bg)
        write(ws, r, 2, fee, 11, True, NAVY, bg, "center")
        write(ws, r, 3, "", 10, False, INK, bg)
        write(ws, r, 4, "", 10, False, INK, bg)
        merge(ws, r, 2, r, 4)
        write(ws, r, 5, how, 10, False, INK, bg)
    ws.row_dimensions[31].height = 48
    write(ws, 31, 1, "结算", 10, True, NAVY, GOLD_SOFT, "center")
    write(ws, 31, 2, "单场验收后结算，再进入年度框架。一页纸可以先共识；承办主体、发票、审批和结算走银行采购与财务流程。不按贷款审批、开户或转介结果向个人支付佣金。开户、授信和资金转化由银行自行评价。", 10, False, INK, PAPER)
    for c in range(3, 6):
        write(ws, 31, c, "", 10, False, INK, PAPER)
    merge(ws, 31, 2, 31, 5)

    chart = BarChart()
    chart.type = "col"
    chart.title = "路演标准版构成"
    chart.y_axis.title = None
    chart.x_axis.title = None
    data = Reference(ws, min_col=4, min_row=5, max_row=8)
    cats = Reference(ws, min_col=1, min_row=6, max_row=8)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.legend = None
    chart.style = 10
    chart.y_axis.numFmt = '¥#,##0'
    chart.width = 15
    chart.height = 6
    ws.add_chart(chart, "A33")
    ws.row_dimensions[33].height = 20
    ws.print_area = "A1:E31"
    ws.page_setup.fitToHeight = 1
    ws.freeze_panes = "A5"


def build_plan(wb):
    ws = wb.create_sheet("10执行计划")
    apply_header_footer(ws, "执行计划")
    set_widths(ws, [28, 36, 52, 18])
    title_block(ws, "2026年11月至2027年第一季度", "近端计划用于沟通。路演四阶段写在下方，供年度框架讨论。", 4)
    header_row(ws, 4, ["时间", "核心任务", "具体成果", "状态"])
    for i, (when, task, result) in enumerate(D.TIMELINE):
        r = 5 + i
        ws.row_dimensions[r].height = 36
        bg = GOLD_SOFT if i < 3 else PAPER
        write(ws, r, 1, when, 10, True, NAVY, bg, "center")
        write(ws, r, 2, task, 10, True, INK, bg)
        write(ws, r, 3, result, 10, False, INK, bg)
        write(ws, r, 4, "进行中" if i == 0 else "未开始", 10, False, RED if i == 0 else MUTED, YELLOW, "center")
    dv = DataValidation(type="list", formula1='"未开始,进行中,已完成,阻塞"', allow_blank=False)
    dv.add("D5:D9")
    ws.add_data_validation(dv)
    write(ws, 11, 1, "路演实施四阶段", 11, True, WHITE, NAVY)
    for c in range(2, 5):
        write(ws, 11, c, "", 11, True, WHITE, NAVY)
    merge(ws, 11, 1, 11, 4)
    header_row(ws, 12, ["阶段", "时间", "工作", ""])
    for i, (phase, when, work) in enumerate(D.PHASES):
        r = 13 + i
        ws.row_dimensions[r].height = 32
        bg = CREAM if i % 2 == 0 else PAPER
        write(ws, r, 1, phase, 10, True, NAVY, bg, "center")
        write(ws, r, 2, when, 10, False, INK, bg, "center")
        write(ws, r, 3, work, 10, False, INK, bg)
        write(ws, r, 4, "", 10, False, INK, bg)
        merge(ws, r, 3, r, 4)
    ws.print_area = "A1:D16"
    ws.page_setup.fitToHeight = 1
    ws.freeze_panes = "A5"


def build_kpi(wb):
    ws = wb.create_sheet("11绩效跟踪")
    apply_header_footer(ws, "绩效跟踪")
    headers = ["类型", "计划时间", "活动名称", "实际日期", "有效企业数", "需求条数", "银行对接数", "满意度（1-5）", "授权情况", "备注"]
    widths = [14, 14, 42, 16, 14, 12, 14, 16, 16, 28]
    set_widths(ws, widths)
    title_block(ws, "活动绩效跟踪表", "会后填写黄色列。开户、授信和资金结果不在本表统计，由银行自行管理。7日 / 30日 / 90日见表末。", len(headers))
    header_row(ws, 4, headers)
    rows = []
    for month, theme, org in D.VISIT_CALENDAR:
        rows.append(("参访", "2027年" + month, theme + "｜" + org))
    for month, name, kind, scale, focus in D.ROADSHOWS:
        rows.append(("路演" if kind == "重点" else "对接", "2027年" + month, name))
    for i, (kind, when, name) in enumerate(rows):
        r = 5 + i
        ws.row_dimensions[r].height = 22
        bg = NAVY_SOFT if kind == "参访" else (GOLD_SOFT if kind == "路演" else PAPER)
        write(ws, r, 1, kind, 10, True, NAVY, bg, "center")
        write(ws, r, 2, when, 10, False, INK, bg, "center")
        write(ws, r, 3, name, 10, False, INK, bg)
        for c in range(4, 11):
            write(ws, r, c, None, 10, False, INK, YELLOW, "center")
    last = 4 + len(rows)
    dv = DataValidation(type="list", formula1='"未收集,已授权,部分授权,不适用"', allow_blank=True)
    dv.add(f"I5:I{last}")
    ws.add_data_validation(dv)
    dv2 = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=True)
    dv2.add(f"H5:H{last}")
    ws.add_data_validation(dv2)
    sum_row = last + 1
    ws.row_dimensions[sum_row].height = 22
    write(ws, sum_row, 1, "合计", 11, True, WHITE, NAVY, "center")
    write(ws, sum_row, 2, f'=COUNTA(C5:C{last})', 11, True, WHITE, NAVY, "center")
    write(ws, sum_row, 3, "场（计划）", 10, True, WHITE, NAVY)
    for col in (5, 6, 7):
        letter = get_column_letter(col)
        write(ws, sum_row, col, f"=SUM({letter}5:{letter}{last})", 11, True, WHITE, NAVY, "center", fmt="0")
    write(ws, sum_row, 4, "", 10, True, WHITE, NAVY)
    write(ws, sum_row, 8, f"=IF(COUNT(H5:H{last})=0,\"\",ROUND(AVERAGE(H5:H{last}),1))", 11, True, WHITE, NAVY, "center")
    write(ws, sum_row, 9, "", 10, True, WHITE, NAVY)
    write(ws, sum_row, 10, "满意度为已填场次的平均值", 9, False, WHITE, NAVY)
    base = sum_row + 2
    write(ws, base, 1, "跟踪节点", 11, True, WHITE, NAVY)
    for c in range(2, 11):
        write(ws, base, c, "", 11, True, WHITE, NAVY)
    merge(ws, base, 1, base, 4)
    header_row(ws, base + 1, ["节点", "完成内容", "", ""])
    for i, (when, what) in enumerate(D.TRACK):
        r = base + 2 + i
        ws.row_dimensions[r].height = 22
        write(ws, r, 1, when, 10, True, NAVY, GOLD_SOFT, "center")
        write(ws, r, 2, what, 10, False, INK, PAPER)
        write(ws, r, 3, "", 10, False, INK, PAPER)
        write(ws, r, 4, "", 10, False, INK, PAPER)
        merge(ws, r, 2, r, 4)
    ws.freeze_panes = "D5"
    ws.print_title_rows = "1:4"
    ws.auto_filter.ref = f"A4:J{last}"
    ws.print_area = f"A1:J{base + 4}"
    ws.page_setup.fitToHeight = 2
    ws.auto_filter.ref = f"A4:J{last}"


def build():
    wb = Workbook()
    wb.properties.title = D.TITLE
    wb.properties.creator = "上海市杨浦区科技企业联合会"
    wb.properties.subject = "资源排期与预算表"
    wb.properties.description = "洽谈稿。预算与接待待正式确认。"
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    build_cover(wb)
    build_onepager(wb)
    build_directions(wb)
    build_confirm(wb)
    build_resources(wb)
    build_calendar(wb)
    build_flow(wb)
    build_roadshow(wb)
    build_themes(wb)
    build_budget(wb)
    build_plan(wb)
    build_kpi(wb)
    # drop unused imports side effects: none
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
