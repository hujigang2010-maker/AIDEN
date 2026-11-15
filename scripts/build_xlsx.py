# -*- coding: utf-8 -*-
"""港交所18C分行业估值模板：可改输入、公式计算的工作簿。

蓝字奶油底是输入。黑色是公式。绿色是结果。
演示数字只复现手册虚构算例，不是行业基准或真实发行人数据。
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.legend import Legend
from openpyxl.chart.series import SeriesLabel
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "output" / "港交所18C分行业估值模型.xlsx"

FONT = "微软雅黑"
NAVY = "0E2A24"
GOLD = "C6A15B"
TEAL = "0F6E62"
INK = "1C2421"
MUTED = "5E6A64"
WHITE = "FFFFFF"
CREAM = "F7F4EE"
INPUT = "FFF6E4"
BLUE = "1F4E79"
GREEN = "E7F4EF"
GREEN_DK = "0F6E62"
ORANGE = "FBF0E4"
ORANGE_DK = "8A5A32"
RED = "F8EBE8"
RED_DK = "A33B32"
LINE = "E4DCCF"
EXAMPLE = "E7EEF6"
PALE = "FBF9F4"

thin = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
no_border = Border()


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def fnt(size=10, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def align(horizontal="left", wrap=True):
    return Alignment(horizontal=horizontal, vertical="center", wrap_text=wrap)


def put(ws, addr, value, kind="label", fmt=None, horizontal="left", comment=None, size=10):
    cell = ws[addr]
    cell.value = value
    cell.alignment = align(horizontal, True)
    cell.border = thin if kind not in ("title", "plain", "section") else no_border
    locked = True
    if kind == "title":
        cell.font = fnt(size or 18, True, WHITE)
        cell.fill = fill(NAVY)
        cell.alignment = align("left", True)
    elif kind == "section":
        cell.font = fnt(size or 12, True, WHITE)
        cell.fill = fill(NAVY)
    elif kind == "header":
        cell.font = fnt(size, True, WHITE)
        cell.fill = fill(NAVY)
        cell.alignment = align(horizontal or "center", True)
    elif kind == "input":
        cell.font = fnt(size, False, BLUE)
        cell.fill = fill(INPUT)
        locked = False
    elif kind == "example":
        cell.font = fnt(size, False, BLUE)
        cell.fill = fill(EXAMPLE)
        locked = False
    elif kind == "formula":
        cell.font = fnt(size, False, INK)
        cell.fill = fill(WHITE)
    elif kind == "output":
        cell.font = fnt(size, True, GREEN_DK)
        cell.fill = fill(GREEN)
    elif kind == "warn":
        cell.font = fnt(size, False, ORANGE_DK)
        cell.fill = fill(ORANGE)
    elif kind == "note":
        cell.font = fnt(size, False, MUTED, True)
        cell.fill = fill(PALE)
    elif kind == "label":
        cell.font = fnt(size, False, INK)
        cell.fill = fill(CREAM)
    elif kind == "plain":
        cell.font = fnt(size, False, INK)
        cell.fill = fill(WHITE)
    else:
        cell.font = fnt(size, False, INK)
    if fmt:
        cell.number_format = fmt
    cell.protection = Protection(locked=locked)
    if comment:
        cell.comment = Comment(comment, "模板", width=260, height=70)
    return cell


def merge(ws, range_str, value, kind="title", size=18):
    ws.merge_cells(range_str)
    addr = range_str.split(":")[0]
    put(ws, addr, value, kind=kind, size=size)
    return ws[addr]


def widths(ws, spec):
    for col, width in spec.items():
        ws.column_dimensions[col].width = width


def heights(ws, spec):
    for row, height in spec.items():
        ws.row_dimensions[row].height = height


def finish(ws, tab, freeze="A5", landscape=True, fit_height=1):
    ws.sheet_properties.tabColor = tab
    ws.freeze_panes = freeze
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_height
    ws.page_setup.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.45, right=0.45, top=0.6, bottom=0.5, header=0.25, footer=0.25)
    ws.oddHeader.left.text = "港交所18C分行业估值模板"
    ws.oddHeader.right.text = "虚构演示 · 非投资建议"
    ws.oddFooter.left.text = "蓝字可改 · 黑字为公式 · 灯号不是法律意见"
    ws.oddFooter.right.text = "第 &P 页"
    ws.oddHeader.left.font = FONT
    ws.oddFooter.left.font = FONT
    ws.sheet_view.view = "normal"
    # openpyxl 里这些开关为 True 时，表示该操作被禁止。
    # 未锁定的蓝字必须可以点选，否则不先撤销整表保护就无法改假设。
    ws.protection.sheet = True
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False
    ws.protection.autoFilter = False
    ws.protection.sort = False
    ws.page_setup.scale = 70


def yes_no(ws, cells):
    dv = DataValidation(type="list", formula1='"是,否"', allow_blank=True)
    dv.error = "请选择是或否"
    dv.errorTitle = "输入"
    dv.add_list(cells) if False else None
    for addr in cells:
        dv.add(ws[addr])
    ws.add_data_validation(dv)


def choice(ws, addr, options):
    formula = '"' + ",".join(options) + '"'
    dv = DataValidation(type="list", formula1=formula, allow_blank=False)
    dv.add(ws[addr])
    ws.add_data_validation(dv)


def banner(ws, last_col, title, subtitle):
    merge(ws, f"A1:{last_col}1", title, "title", 18)
    merge(ws, f"A2:{last_col}2", subtitle, "note", 9)
    ws["A2"].alignment = align("left", True)
    heights(ws, {1: 28, 2: 32})


def build_cover(wb):
    ws = wb.active
    ws.title = "00_投委会摘要"
    widths(ws, {c: 16 for c in "ABCDEFGHIJKL"})
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["F"].width = 20
    banner(
        ws,
        "L",
        "港交所18C  ·  分行业估值模板",
        "研究基准日 2026-11-06  ·  V1.0  ·  下面四块是四条分开的计算链，不能把经营EV、演示倍数、回调桥和上市门槛加在一起。",
    )
    merge(ws, "A3:L3", "性质：虚构演示与可复用公式。不是独立评估、投资建议、会计或法律意见，也不是行业均值。", "warn", 9)
    heights(ws, {3: 22})

    merge(ws, "A5:D5", "计算链 A　经营DCF（手册第4.5节虚构AI企业）", "section", 11)
    labels = [
        (6, "VAL_EV  经营企业价值", "='05_DCF'!A18", "0.00"),
        (7, "预测期FCFF现值", "='05_DCF'!A14", "0.00"),
        (8, "终值现值", "='05_DCF'!A17", "0.00"),
        (9, "终值现值 / 经营EV", "='05_DCF'!A19", "0.0%"),
        (10, "诊断", "='05_DCF'!B20", None),
    ]
    for row, label, formula, fmt in labels:
        put(ws, f"A{row}", label, "label")
        kind = "warn" if row == 10 else "output"
        put(ws, f"B{row}", formula, kind, fmt, "center" if row < 10 else "left")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        heights(ws, {row: 22 if row < 10 else 48})

    merge(ws, "E5:H5", "计算链 B　市场法对照（演示倍数，不是行情）", "section", 11)
    market = [
        (6, "NTM收入（2027）", "='07_市场法'!C4", "0.000"),
        (7, "采用倍数", "='07_市场法'!C8", "0.00"),
        (8, "调整后指示EV", "='07_市场法'!C10", "0.00"),
        (9, "与DCF的差额", "='07_市场法'!C12", "0.00"),
        (10, "倍数来源", "='07_市场法'!C7", None),
    ]
    for row, label, formula, fmt in market:
        put(ws, f"E{row}", label, "label")
        put(ws, f"F{row}", formula, "output", fmt, "center" if fmt else "left")
        ws.merge_cells(start_row=row, start_column=6, end_row=row, end_column=8)

    merge(ws, "A12:D12", "计算链 C　回调结构演示（不是上面的DCF）", "section", 11)
    for row, label, formula, fmt in (
        (13, "上期股权价值", "='08_股权桥'!C17", "0.00"),
        (14, "计算得到的本期", "='08_股权桥'!C22", "0.00"),
        (15, "桥的残差", "='08_股权桥'!C24", "0.00"),
        (16, "勾稽", "='08_股权桥'!C25", None),
    ):
        put(ws, f"A{row}", label, "label")
        put(ws, f"B{row}", formula, "output", fmt, "center")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)

    merge(ws, "E12:H12", "计算链 D　普通股稀释演示（无优先权）", "section", 11)
    for row, label, formula, fmt in (
        (13, "情景一新投资者", "='09_稀释与IPO'!C8", "0.0%"),
        (14, "情景二新投资者", "='09_稀释与IPO'!C14", "0.0%"),
        (15, "持股上升", "='09_稀释与IPO'!C16", "0.0%"),
        (16, "rNPV单项演示", "='12_生物医药'!C8", "0.00"),
    ):
        put(ws, f"E{row}", label, "label")
        put(ws, f"F{row}", formula, "output", fmt, "center")
        ws.merge_cells(start_row=row, start_column=6, end_row=row, end_column=8)

    merge(ws, "A18:H18", "18C数值灯号　PASS只代表表内门槛，不代表可以申报", "section", 11)
    for col, text in (("A", "对象"), ("B", "阶段"), ("C", "总灯"), ("E", "对象"), ("F", "阶段"), ("G", "总灯")):
        put(ws, f"{col}19", text, "header", horizontal="center")
    put(ws, "A20", "本项目（默认空白）", "label")
    put(ws, "B20", "='02_18C合规'!C25", "output", horizontal="center")
    put(ws, "C20", "='02_18C合规'!C40", "output", horizontal="center")
    put(ws, "E20", "示例甲·已商业化", "label")
    put(ws, "F20", "='02_18C合规'!D25", "output", horizontal="center")
    put(ws, "G20", "='02_18C合规'!D40", "output", horizontal="center")
    put(ws, "A21", "示例乙·未商业化", "label")
    put(ws, "B21", "='02_18C合规'!E25", "output", horizontal="center")
    put(ws, "C21", "='02_18C合规'!E40", "output", horizontal="center")
    put(ws, "E21", "采用汇率状态", "label")
    put(ws, "F21", "='01_使用说明'!C15", "warn", horizontal="left")
    ws.merge_cells("F21:H21")

    merge(ws, "A23:H23", "模型校验", "section", 11)
    put(ws, "A24", "身份恒等式异常数", "label")
    put(ws, "B24", "='17_校验'!C4", "output", "0", "center")
    put(ws, "C24", "提示数（不是计算错误）", "label")
    put(ws, "D24", "='17_校验'!C5", "warn", "0", "center")
    put(ws, "E24", "打开模型后若异常数不是0，先不要用结果。", "note")
    ws.merge_cells("E24:H24")

    merge(ws, "A26:H26", "怎么用", "section", 11)
    notes = [
        "1  只改奶油底蓝字。要插入行，先在「审阅 → 撤销工作表保护」，本文件没有密码。",
        "2  AI演示从「03_AI输入」改增长、毛利率、WACC和g，「04_AI预测」「05_DCF」「06_敏感性」会跟着变。",
        "3  18C请填「02_18C合规」的「本项目」列。示例甲、示例乙是规则演算，不是这家虚构AI公司的申报结论。",
        "4  市场法默认3.0倍和-20%调整来自手册的纯演示。正式项目改「07_市场法」，或改为引用可比中位数。",
        "5  汇率在「01_使用说明」。1.09只为把公式跑通，不是2026-11-06的市场汇率。",
    ]
    for idx, text in enumerate(notes):
        row = 27 + idx
        merge(ws, f"A{row}:H{row}", text, "plain", 10)
        heights(ws, {row: 20})

    # 图表数据放在远处，打印区域不包含它
    put(ws, "AA6", "DCF经营EV", "plain")
    put(ws, "AB6", "='05_DCF'!A18", "formula", "0.00")
    put(ws, "AA7", "市场法原始", "plain")
    put(ws, "AB7", "='07_市场法'!C9", "formula", "0.00")
    put(ws, "AA8", "市场法调整后", "plain")
    put(ws, "AB8", "='07_市场法'!C10", "formula", "0.00")

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "经营价值（亿元）"
    chart.y_axis.title = "亿元人民币"
    data = Reference(ws, min_col=28, min_row=6, max_row=8)
    cats = Reference(ws, min_col=27, min_row=6, max_row=8)
    chart.add_data(data, from_rows=False, titles_from_data=False)
    chart.set_categories(cats)
    chart.shape = 4
    chart.legend = None
    chart.style = 10
    chart.y_axis.majorGridlines = None
    chart.series[0].graphicalProperties.solidFill = NAVY
    chart.series[0].dLbls = DataLabelList()
    chart.series[0].dLbls.showVal = True
    chart.series[0].dLbls.showCatName = False
    chart.series[0].dLbls.showSerName = False
    chart.width = 12
    chart.height = 6
    ws.add_chart(chart, "I5")

    for addr in ("C20", "G20", "C21"):
        ws.conditional_formatting.add(
            addr,
            CellIsRule(operator="equal", formula=['"PASS"'], fill=fill("CDE8DF"), font=fnt(10, True, GREEN_DK)),
        )
        ws.conditional_formatting.add(
            addr,
            CellIsRule(operator="equal", formula=['"FAIL"'], fill=fill(RED), font=fnt(10, True, RED_DK)),
        )
        ws.conditional_formatting.add(
            addr,
            CellIsRule(operator="equal", formula=['"CONDITIONAL"'], fill=fill(ORANGE), font=fnt(10, True, ORANGE_DK)),
        )
        ws.conditional_formatting.add(
            addr,
            CellIsRule(operator="equal", formula=['"NOT-ASSESSABLE"'], fill=fill("EEEAE2"), font=fnt(10, True, MUTED)),
        )
    ws.auto_filter.ref = None
    ws.print_title_rows = "1:2"
    ws.page_setup.fitToHeight = 1
    ws.sheet_view.showGridLines = False
    finish(ws, GOLD, freeze="A5", fit_height=1)
    ws.column_dimensions["AA"].hidden = True
    ws.column_dimensions["AB"].hidden = True
    return ws


def build_readme(wb):
    ws = wb.create_sheet("01_使用说明")
    widths(ws, {"A": 28, "B": 42, "C": 24, "D": 55, "E": 22})
    banner(ws, "E", "使用说明、颜色和汇率", "先统一估值日、主体、币种和股权类别。本表默认人民币经营预测，港币只用于18C规则测试。")
    rows = [
        ("颜色", "含义", "能不能改"),
        ("奶油底、蓝字", "输入：假设、合同事实、行情、汇率", "能"),
        ("白底、黑字", "公式", "不能。要改就回到蓝字"),
        ("绿底", "结果：EV、灯号、勾稽", "不能"),
        ("橙底", "提示：终值过高、演示汇率、税损未建", "阅读，不必改公式"),
        ("浅蓝底、蓝字", "规则示例，仍可改，用来看灯号怎么变", "能"),
    ]
    for col, text in enumerate(("项目", "含义", "操作"), 1):
        put(ws, f"{get_column_letter(col)}4", text, "header", horizontal="center")
    body = rows[1:]
    for idx, (a, b, c) in enumerate(body):
        kind = "input" if idx == 0 else "formula" if idx == 1 else "output" if idx == 2 else "warn" if idx == 3 else "example"
        put(ws, f"A{5+idx}", a, kind if idx < 5 else "label")
        put(ws, f"B{5+idx}", b, "plain")
        put(ws, f"C{5+idx}", c, "plain")

    merge(ws, "A11:C11", "汇率（港币 / 1人民币）", "section", 11)
    put(ws, "A12", "正式汇率", "label")
    put(ws, "C12", None, "input", "0.0000", "center", "填估值日指定汇率。空着时下面的演示汇率才会被采用。")
    put(ws, "A13", "演示汇率", "label")
    put(ws, "C13", 1.09, "example", "0.0000", "center", "只为把港币公式跑通。不是2026-11-06市场汇率，不能写进正式意见。")
    put(ws, "A14", "当前采用", "label")
    put(ws, "C14", '=IF(C12="",C13,C12)', "output", "0.0000", "center")
    put(ws, "A15", "状态", "label")
    put(ws, "C15", '=IF(C12="","正在使用演示汇率，港币结果只供公式演练","正在使用正式汇率")', "warn")
    ws.merge_cells("C15:E15")

    merge(ws, "A17:E17", "六项输出，名称不要混用", "section", 11)
    outputs = [
        ("VAL_EV", "经营企业价值", "经营性资产。本演示等于DCF经营EV，尚未扣净债务。"),
        ("VAL_EQ_PRE", "融资前全部股权价值", "经营EV加减现金、债务、租赁、少数股东和或有项目之后。"),
        ("VAL_EQ_POST", "融资后全部股权价值", "只有新现金确实进入主体时才加上。不能和Pre-money各加一次。"),
        ("VAL_SHARE", "完全摊薄每股价值", "优先股转换、期权和权证处理完再除。"),
        ("VAL_RANGE", "区间", "悲观、基准、乐观都要留下权重和条件。"),
        ("HK18C_GATE", "18C适用性诊断", "PASS / CONDITIONAL / FAIL / NOT-ASSESSABLE。不替代估值。"),
    ]
    for col, text in enumerate(("代码", "输出", "口径"), 1):
        put(ws, f"{get_column_letter(col)}18", text, "header")
    for i, item in enumerate(outputs):
        for col, text in enumerate(item, 1):
            put(ws, f"{get_column_letter(col)}{19+i}", text, "plain")

    merge(ws, "A26:E26", "工作表地图", "section", 11)
    sheets = [
        "00 摘要看四条计算链是否被混用",
        "02 填本项目的18C事实；示例列用来理解灯号",
        "03–06 改AI驱动、看五年预测、DCF和敏感性",
        "07 市场法。默认不读取空的可比表",
        "08–09 股权桥、回调桥、稀释和IPO股数",
        "10–14 五个行业的公式骨架，自检数可以整格替换",
        "15–16 方法库和回调顺序，给投委会翻",
        "17 恒等式。异常数应为0",
        "18–20 来源、报告底稿、参数字典",
    ]
    for i, text in enumerate(sheets):
        merge(ws, f"A{27+i}:E{27+i}", text, "plain", 10)
    finish(ws, "1A433A", fit_height=1)
    return ws


def build_compliance(wb):
    ws = wb.create_sheet("02_18C合规")
    widths(ws, {"A": 38, "B": 16, "C": 24, "D": 26, "E": 26, "F": 36})
    banner(
        ws,
        "F",
        "18C数值门槛",
        "收入只填特专科技业务，不填合同额、在手订单、ARR、GMV或未确认订阅。临时市值门槛：2024-09-01至2027-08-31为40亿/80亿港元；规则正文为60亿/100亿。",
    )
    headers = ["项目", "单位", "本项目", "示例甲·已商业化", "示例乙·未商业化", "填写提示"]
    for col, text in enumerate(headers, 1):
        put(ws, f"{get_column_letter(col)}4", text, "header", horizontal="center")

    input_rows = [
        (6, "最近审计年特专科技收入", "港元", None, 280_000_000, 120_000_000, "只计合资格分部，剔除无关业务和内部交易"),
        (7, "预计上市日", "日期", None, date(2026, 11, 6), date(2026, 11, 6), "用来判断临时门槛是否适用"),
        (8, "第1年合资格研发支出", "港元", None, 40_000_000, 80_000_000, "按指引口径，不是利润表研发费用原数"),
        (9, "第1年总经营支出", "港元", None, 200_000_000, 150_000_000, "总费用－销售成本－融资费用＋未费用化的合资格研发"),
        (10, "第2年合资格研发支出", "港元", None, 45_000_000, 90_000_000, ""),
        (11, "第2年总经营支出", "港元", None, 210_000_000, 160_000_000, ""),
        (12, "第3年合资格研发支出", "港元", None, 50_000_000, 100_000_000, ""),
        (13, "第3年总经营支出", "港元", None, 220_000_000, 170_000_000, ""),
        (14, "可用营运资金", "港元", None, None, 200_000_000, "未商业化才做125%测试"),
        (15, "上市文件刊发日起12个月相关成本", "港元", None, None, 140_000_000, "按交易所定义，含应计入的资本支出"),
        (16, "预计发行市值", "港元", None, 5_000_000_000, 9_000_000_000, "发行价×发行后等价股数，不是DCF"),
        (17, "Pathfinder资深投资者人数", "人", None, 3, 3, "指示性基准通常2–5名"),
        (18, "均至少于申请前12个月投资", "是/否", None, "是", "是", ""),
        (19, "申请日合计持股", "比例", None, 0.12, 0.12, "与下一行是“或”的关系，不能只看一个"),
        (20, "合计投资额", "港元", None, 1_600_000_000, 1_600_000_000, "指示性基准15亿港元"),
        (21, "至少两名各达3%或各4.5亿港元", "是/否", None, "是", "是", "指示性筛查，不是完整资格意见"),
    ]
    yes_no(ws, ["C18", "D18", "E18", "C21", "D21", "E21"])
    for row, label, unit, project, ex_a, ex_b, tip in input_rows:
        put(ws, f"A{row}", label, "label")
        put(ws, f"B{row}", unit, "note", horizontal="center")
        fmt = "YYYY-MM-DD" if row == 7 else "0.0%" if row == 19 else "#,##0" if unit == "港元" else "0"
        put(ws, f"C{row}", project, "input", fmt, "center")
        put(ws, f"D{row}", ex_a, "example", fmt, "center")
        put(ws, f"E{row}", ex_b, "example", fmt, "center")
        put(ws, f"F{row}", tip, "note", size=8)

    merge(ws, "A23:F23", "规则演算　示例列的数字是虚构的门槛样本，不是第4章那家AI公司", "section", 11)

    def gate_formulas(col):
        src = col
        return {
            25: f'=IF({src}6="","NOT-ASSESSABLE",IF({src}6>=250000000,"COMMERCIAL","PRE-COMMERCIAL"))',
            26: (
                f'=IF(OR({src}6="",{src}7=""),"",'
                f'IF({src}6>=250000000,'
                f'IF(AND({src}7>=DATE(2024,9,1),{src}7<=DATE(2027,8,31)),4000000000,6000000000),'
                f'IF(AND({src}7>=DATE(2024,9,1),{src}7<=DATE(2027,8,31)),8000000000,10000000000)))'
            ),
            27: (
                f'=IF({src}7="","日期空缺",'
                f'IF(AND({src}7>=DATE(2024,9,1),{src}7<=DATE(2027,8,31)),"临时门槛","规则正文门槛"))'
            ),
            28: f'=IF({src}6="","",IF({src}6>=250000000,0.15,IF({src}6>=150000000,0.3,0.5)))',
            29: f'=IF(OR({src}8="",{src}9="",{src}9=0),"",{src}8/{src}9)',
            30: f'=IF(OR({src}10="",{src}11="",{src}11=0),"",{src}10/{src}11)',
            31: f'=IF(OR({src}12="",{src}13="",{src}13=0),"",{src}12/{src}13)',
            32: (
                f'=IF(AND(ISNUMBER({src}29),{src}29>={src}28),1,0)'
                f'+IF(AND(ISNUMBER({src}30),{src}30>={src}28),1,0)'
                f'+IF(AND(ISNUMBER({src}31),{src}31>={src}28),1,0)'
            ),
            33: f'=IF(({src}9+{src}11+{src}13)=0,"",({src}8+{src}10+{src}12)/({src}9+{src}11+{src}13))',
            34: f'=IF(OR({src}28="",{src}33=""),"",IF(AND({src}32>=2,{src}33>={src}28),"RND-PASS","RND-CHECK"))',
            35: f'=IF(OR({src}16="",{src}26=""),"NOT-ASSESSABLE",IF({src}16>={src}26,"CAP-PASS","CAP-FAIL"))',
            36: (
                f'=IF({src}25="COMMERCIAL","商业化公司走一般营运资金要求，此格不自动给通过",'
                f'IF(OR({src}14="",{src}15=""),"NOT-ASSESSABLE",IF({src}14>={src}15*1.25,"WC-PASS","WC-CHECK")))'
            ),
            37: (
                f'=IF(OR({src}17="",{src}18="",{src}19="",{src}20="",{src}21=""),"NOT-ASSESSABLE",'
                f'IF(AND({src}17>=2,{src}17<=5,{src}18="是",OR({src}19>=0.1,{src}20>=1500000000),{src}21="是"),'
                f'"SII-SCREEN-PASS","SII-SCREEN-CHECK"))'
            ),
            40: (
                f'=IF(OR({src}6="",{src}7="",{src}8="",{src}9="",{src}10="",{src}11="",{src}12="",{src}13=""),"NOT-ASSESSABLE",'
                f'IF(OR({src}34<>"RND-PASS",{src}35<>"CAP-PASS"),"FAIL",'
                f'IF({src}37<>"SII-SCREEN-PASS","CONDITIONAL",'
                f'IF(AND({src}25="PRE-COMMERCIAL",{src}36<>"WC-PASS"),"CONDITIONAL","PASS"))))'
            ),
        }

    labels = {
        25: "阶段",
        26: "最低市值",
        27: "门槛口径",
        28: "要求的研发比率",
        29: "第1年研发比率",
        30: "第2年研发比率",
        31: "第3年研发比率",
        32: "单年达标个数",
        33: "三年合计研发比率",
        34: "研发测试",
        35: "市值测试",
        36: "营运资金测试",
        37: "资深投资者指示性筛查",
        40: "表内总灯",
    }
    for row, label in labels.items():
        put(ws, f"A{row}", label, "label")
        put(ws, f"B{row}", "港元" if row == 26 else "%" if row in (28, 29, 30, 31, 33) else "—", "note", horizontal="center")
    for col in ("C", "D", "E"):
        formulas = gate_formulas(col)
        for row, formula in formulas.items():
            fmt = '"HK$"#,##0' if row == 26 else "0.0%" if row in (28, 29, 30, 31, 33) else "0" if row == 32 else None
            put(ws, f"{col}{row}", formula, "output", fmt, "center")
    put(ws, "F25", "收入≥2.5亿港元为已商业化，否则为未商业化。", "note", size=8)
    put(ws, "F34", "三年中至少两年各自达标，且三年合计达标。", "note", size=8)
    put(ws, "F37", "这是指引里的指示性基准筛查，不能代替完整规则判断。", "note", size=8)
    put(ws, "F40", "PASS不等于法律意见。配售50%、披露和豁免都不在此灯里。", "note", size=8)
    merge(
        ws,
        "A42:F42",
        "若把示例乙的发行市值改成 7,000,000,000，市值测试会从 CAP-PASS 变成 CAP-FAIL，总灯变为 FAIL。可以用这一格确认公式听得懂门槛。",
        "note",
        9,
    )
    heights(ws, {42: 32})
    finish(ws, TEAL, freeze="A6", fit_height=1)
    return ws


def build_ai_inputs(wb):
    ws = wb.create_sheet("03_AI输入")
    widths(ws, {"A": 32, "B": 16, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 18, "I": 42})
    banner(ws, "I", "AI经营驱动", "全部是手册第4章的虚构演示。改蓝字后，预测、DCF和敏感性会重算。不要把这些比例写成行业正常值。")
    merge(ws, "A4:I4", "A. 2026年订阅与Token收入桥", "section", 11)
    for col, text in enumerate(("参数", "代码", "数值", "单位", "说明"), 1):
        put(ws, f"{get_column_letter(col)}5", text, "header")
    ws.merge_cells("E5:I5")
    items = [
        (6, "期初订阅ARR", "A01", 1.20, "亿元", "input", "MRR合同与开票。演示1.20"),
        (7, "净收入留存NRR", "A02", 1.15, "倍", "input", "已含增购和流失。演示115%，不要再扣一次流失"),
        (8, "新增ARR", "A03", 0.80, "亿元", "input", "只计已生效的付费合同"),
        (9, "期末ARR", "A04", "=C6*C7+C8", "亿元", "formula", "期初×NRR＋新增"),
        (10, "平均ARR", "A05", "=(C6+C9)/2", "亿元", "formula", "手册用年初年末简单平均，正式模型要按月确认"),
        (11, "Token已确认收入", "A06", 0.45, "亿元", "input", "免费调用不算收入"),
        (12, "桥表合计 = 2026收入", "A07", "=C10+C11", "亿元", "output", "演示应等于2.14。预测以此为基数"),
    ]
    for row, name, code, value, unit, kind, note in items:
        put(ws, f"A{row}", name, "label")
        put(ws, f"B{row}", code, "note", horizontal="center")
        fmt = "0.00%" if row == 7 else "0.00"
        # NRR stored as 1.15 and shown as 115.0% if format is percent? 1.15 with 0.00% becomes 115.00%. Use 0.00 for multiple to avoid confusion, label says 倍.
        fmt = "0.00"
        put(ws, f"C{row}", value, kind if kind != "formula" else "formula", fmt, "center", note if kind == "input" else None)
        if kind == "output":
            put(ws, f"C{row}", value, "output", fmt, "center")
        put(ws, f"D{row}", unit, "note", horizontal="center")
        put(ws, f"E{row}", note, "note", size=8)
        ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=9)

    merge(ws, "A14:I14", "B. 2027–2031经营比例", "section", 11)
    for col, text in enumerate(("项目", "代码", 2027, 2028, 2029, 2030, 2031, "单位", "说明"), 1):
        put(ws, f"{get_column_letter(col)}15", text, "header", horizontal="center")
    drivers = [
        (17, "收入增长", "B01", [0.60, 0.45, 0.30, 0.20, 0.12], "相对上年"),
        (18, "毛利率", "B02", [0.50, 0.56, 0.61, 0.65, 0.68], "已确认收入口径"),
        (19, "费用化研发/收入", "B03", [0.42, 0.35, 0.28, 0.23, 0.20], "费用化部分。资本化现金放在CAPEX"),
        (20, "销管/收入", "B04", [0.32, 0.26, 0.23, 0.20, 0.18], ""),
        (21, "资本开支/收入", "B05", [0.08, 0.08, 0.07, 0.06, 0.05], "含应资本化的现金"),
        (22, "D&A/收入", "B06", [0.03, 0.03, 0.03, 0.03, 0.03], "手册简化为收入的3%"),
        (23, "营运资金/增量收入", "B07", [0.04, 0.04, 0.04, 0.04, 0.04], "ΔNWC = 当年收入增量 × 该比例"),
        (24, "所得税率", "B08", [0.25, 0.25, 0.25, 0.25, 0.25], "EBIT为负时税额取0。没有做税损结转"),
    ]
    for row, name, code, values, note in drivers:
        put(ws, f"A{row}", name, "label")
        put(ws, f"B{row}", code, "note", horizontal="center")
        for offset, value in enumerate(values):
            put(ws, f"{get_column_letter(3+offset)}{row}", value, "input", "0.0%", "center")
        put(ws, f"H{row}", "%", "note", horizontal="center")
        put(ws, f"I{row}", note, "note", size=8)

    merge(ws, "A26:I26", "C. 折现。直接输入用于复现手册；右边的搭建栏默认不参与计算", "section", 11)
    put(ws, "A27", "WACC", "label")
    put(ws, "C27", 0.16, "input", "0.0%", "center", "演示16%，敏感性看14%到18%。不是AI行业资本成本。")
    put(ws, "D27", "终值增长率 g", "label")
    put(ws, "E27", 0.03, "input", "0.0%", "center", "演示3%。必须小于WACC。只在稳态使用。")
    put(ws, "A28", "折现约定", "label")
    put(ws, "C28", "2026年末为第0点，2027年末为第1期，使用年末折现。", "note")
    ws.merge_cells("C28:I28")
    finish(ws, "1F4E79", freeze="A5", fit_height=1)
    return ws


def build_forecast(wb):
    ws = wb.create_sheet("04_AI预测")
    widths(ws, {"A": 28, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 36})
    banner(ws, "H", "AI五年预测", "金额单位：亿元人民币。全部公式。税额在EBIT为负时为0，税损结转没有建模。")
    for col, year in enumerate((2026, 2027, 2028, 2029, 2030, 2031), 2):
        put(ws, f"{get_column_letter(col)}4", year, "header", "0", "center")
    put(ws, "A4", "项目", "header")
    labels = [
        (5, "收入增长"),
        (6, "收入"),
        (7, "毛利率"),
        (8, "毛利"),
        (9, "费用化研发"),
        (10, "销售及管理费用"),
        (11, "EBITDA"),
        (12, "D&A"),
        (13, "EBIT"),
        (14, "经营现金税"),
        (15, "NOPAT"),
        (16, "资本开支"),
        (17, "营运资金增加"),
        (18, "FCFF"),
    ]
    for row, label in labels:
        put(ws, f"A{row}", label, "label")
    put(ws, "B5", "—", "note", horizontal="center")
    put(ws, "B6", "='03_AI输入'!C12", "output", "0.000", "center")
    for row in (7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18):
        put(ws, f"B{row}", "—", "note", horizontal="center")
    prev = "B"
    for col in "CDEFG":
        put(ws, f"{col}5", f"='03_AI输入'!{col}17", "formula", "0.0%", "center")
        put(ws, f"{col}6", f"={prev}6*(1+{col}5)", "formula", "0.000", "center")
        put(ws, f"{col}7", f"='03_AI输入'!{col}18", "formula", "0.0%", "center")
        put(ws, f"{col}8", f"={col}6*{col}7", "formula", "0.000", "center")
        put(ws, f"{col}9", f"={col}6*'03_AI输入'!{col}19", "formula", "0.000", "center")
        put(ws, f"{col}10", f"={col}6*'03_AI输入'!{col}20", "formula", "0.000", "center")
        put(ws, f"{col}11", f"={col}8-{col}9-{col}10", "output", "0.000", "center")
        put(ws, f"{col}12", f"={col}6*'03_AI输入'!{col}22", "formula", "0.000", "center")
        put(ws, f"{col}13", f"={col}11-{col}12", "formula", "0.000", "center")
        put(ws, f"{col}14", f"=MAX(0,{col}13)*'03_AI输入'!{col}24", "formula", "0.000", "center")
        put(ws, f"{col}15", f"={col}13-{col}14", "formula", "0.000", "center")
        put(ws, f"{col}16", f"={col}6*'03_AI输入'!{col}21", "formula", "0.000", "center")
        put(ws, f"{col}17", f"=({col}6-{prev}6)*'03_AI输入'!{col}23", "formula", "0.000", "center")
        put(ws, f"{col}18", f"={col}15+{col}12-{col}16-{col}17", "output", "0.000", "center")
        prev = col
    put(ws, "H6", "来自收入桥，不是另填的一个收入。", "note", size=8)
    put(ws, "H11", "毛利－研发－销管。与“收入×（毛利率－研发率－销管率）”是同一件事。", "note", size=8)
    put(ws, "H18", "NOPAT＋D&A－CAPEX－ΔNWC。D&A加回后，不再把折旧重复扣一次。", "note", size=8)
    put(ws, "A20", "2027 EBITDA交叉", "label")
    put(ws, "C20", "=C6*(C7-'03_AI输入'!C19-'03_AI输入'!C20)", "formula", "0.000", "center")
    put(ws, "D20", '=IF(ABS(C20-C11)<0.0000001,"与EBITDA一致","异常")', "output", horizontal="center")
    finish(ws, "3E4A45", freeze="B5", fit_height=1)
    return ws


def build_dcf(wb):
    ws = wb.create_sheet("05_DCF")
    widths(ws, {"A": 32, "B": 14, "C": 16, "D": 16, "E": 16, "F": 18, "G": 28, "H": 18, "I": 22})
    banner(ws, "I", "FCFF折现", "年末折现。终值只用戈登增长，且仅当WACC大于g。退出倍数留空，填了才出现第二条终值，不替代主结果。")
    put(ws, "A3", "采用的WACC", "label")
    put(ws, "B3", "='03_AI输入'!C27", "output", "0.0%", "center")
    put(ws, "A4", "采用的g", "label")
    put(ws, "B4", "='03_AI输入'!E27", "output", "0.0%", "center")
    put(ws, "A5", "折现率检验", "label")
    put(ws, "B5", '=IF(B3>B4,"WACC > g","ERROR：WACC必须大于g")', "output", horizontal="center")
    ws.merge_cells("B5:C5")

    for col, text in enumerate(("年度", "期数 t", "FCFF", "折现因子", "现值"), 1):
        put(ws, f"{get_column_letter(col)}7", text, "header", horizontal="center")
    for i, col in enumerate("CDEFG"):
        row = 8 + i
        put(ws, f"A{row}", f"='04_AI预测'!{col}4", "formula", "0", "center")
        put(ws, f"B{row}", i + 1, "formula", "0", "center")
        put(ws, f"C{row}", f"='04_AI预测'!{col}18", "formula", "0.000", "center")
        put(ws, f"D{row}", f"=1/(1+$B$3)^B{row}", "formula", "0.000", "center")
        put(ws, f"E{row}", f"=C{row}*D{row}", "output", "0.000", "center")

    put(ws, "A14", "FCFF现值合计", "label")
    put(ws, "E14", "=SUM(E8:E12)", "output", "0.000", "center")
    put(ws, "A15", "最后一年FCFF", "label")
    put(ws, "E15", "=C12", "formula", "0.000", "center")
    put(ws, "A16", "戈登终值", "label")
    put(ws, "E16", '=IF(B3<=B4,"",E15*(1+B4)/(B3-B4))', "formula", "0.000", "center")
    put(ws, "A17", "终值现值", "label")
    put(ws, "E17", '=IF(E16="","",E16/(1+B3)^5)', "output", "0.000", "center")
    put(ws, "A18", "经营EV", "label")
    put(ws, "E18", '=IF(E16="","",E14+E17)', "output", "0.000", "center")
    put(ws, "A19", "终值现值 / 经营EV", "label")
    put(ws, "E19", '=IF(OR(E18="",E18=0),"",E17/E18)', "output", "0.0%", "center")
    put(ws, "A20", "诊断", "label")
    put(
        ws,
        "B20",
        '=IF(E18="","",IF(E17>E18,"终值现值高于经营EV：短期现金消耗大，高度依赖远期，不能直接当作可投定价","终值现值没有超过经营EV"))',
        "warn",
    )
    ws.merge_cells("B20:E20")
    heights(ws, {20: 32})

    merge(ws, "A22:E22", "退出倍数只作交叉，不写进经营EV", "section", 11)
    put(ws, "A23", "退出EV/EBITDA倍数", "label")
    put(ws, "C23", None, "input", "0.00", "center", "留空。这不是手册演示的主结果。")
    put(ws, "A24", "退出法终值", "label")
    put(ws, "C24", '=IF(C23="","",\'04_AI预测\'!G11*C23)', "formula", "0.000", "center")
    put(ws, "A25", "退出法经营EV", "label")
    put(ws, "C25", '=IF(C24="","",E14+C24/(1+B3)^5)', "output", "0.000", "center")
    put(ws, "D25", "FCFF现值＋退出终值的现值。不替代E18。", "note", size=8)
    ws.merge_cells("D25:I25")

    merge(ws, "G3:I3", "资本成本搭建（默认不接入）", "section", 11)
    put(ws, "G4", "此处故意留空。成熟公司的WACC统计不能直接抄到亏损的前商业化企业。要启用，先填完，再在报告里说明，并把手册演示WACC替换掉。", "note", size=8)
    ws.merge_cells("G4:I6")
    build_rows = [
        (7, "无风险利率", None),
        (8, "杠杆Beta", None),
        (9, "股权风险溢价", None),
        (10, "国家风险溢价", None),
        (11, "特定风险溢价", None),
        (12, "税前债务成本", None),
        (13, "债务权重", None),
    ]
    for row, label, _ in build_rows:
        put(ws, f"G{row}", label, "label")
        fmt = "0.00" if row == 8 else "0.0%"
        put(ws, f"H{row}", None, "input", fmt, "center")
    put(ws, "G14", "股权成本Ke", "label")
    put(ws, "H14", '=IF(COUNT(H7,H8,H9)<3,"",H7+H8*H9+IF(H10="",0,H10)+IF(H11="",0,H11))', "formula", "0.0%", "center")
    put(ws, "G15", "税后债务成本", "label")
    put(ws, "H15", '=IF(H12="","",H12*(1-\'03_AI输入\'!C24))', "formula", "0.0%", "center")
    put(ws, "G16", "搭建WACC", "label")
    put(ws, "H16", '=IF(OR(H14="",H12="",H13=""),"",(1-H13)*H14+H13*H15)', "output", "0.0%", "center")
    put(ws, "I16", "本模板的经营EV仍引用C27的直接WACC，避免半成品折现率悄悄进入结论。", "note", size=8)
    finish(ws, GREEN_DK, freeze="A7", fit_height=1)
    # 摘要和校验引用的是 A18 这一套易读地址，这里做一列镜像，避免使用者找E列。
    put(ws, "A18", "经营EV", "label")
    return ws


def build_sensitivity(wb):
    ws = wb.create_sheet("06_敏感性")
    widths(ws, {"A": 22, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 42})
    banner(ws, "H", "WACC与终值增长率", "FCFF保持当前预测不变，只改折现率和g。黄轴可改。基准格应等于05页经营EV。")
    put(ws, "B6", "终值增长率 g", "header")
    ws.merge_cells("B6:G6")
    put(ws, "B7", "WACC \\ g", "header", horizontal="center")
    for col, value in zip("CDEFG", (0.01, 0.02, 0.03, 0.04, 0.05)):
        put(ws, f"{col}7", value, "input", "0.0%", "center")
    for row, value in zip(range(8, 13), (0.12, 0.14, 0.16, 0.18, 0.20)):
        put(ws, f"B{row}", value, "input", "0.0%", "center")
        for col in "CDEFG":
            formula = (
                f'=IF({col}$7>=$B{row},"n.m.",'
                f"'04_AI预测'!C18/(1+$B{row})"
                f"+'04_AI预测'!D18/(1+$B{row})^2"
                f"+'04_AI预测'!E18/(1+$B{row})^3"
                f"+'04_AI预测'!F18/(1+$B{row})^4"
                f"+'04_AI预测'!G18/(1+$B{row})^5"
                f"+('04_AI预测'!G18*(1+{col}$7)/($B{row}-{col}$7))/(1+$B{row})^5)"
            )
            put(ws, f"{col}{row}", formula, "output", "0.00", "center")
    ws.conditional_formatting.add(
        "C8:G12",
        ColorScaleRule(
            start_type="min", start_color="F4C7C3",
            mid_type="percentile", mid_value=50, mid_color="F7F4EE",
            end_type="max", end_color="8FCBB8",
        ),
    )
    put(ws, "A14", "当前WACC / g 对应的格子", "label")
    put(
        ws,
        "C14",
        '=IFERROR(INDEX(C8:G12,MATCH(\'05_DCF\'!B3,B8:B12,0),MATCH(\'05_DCF\'!B4,C7:G7,0)),"轴上没有当前参数")',
        "output",
        "0.00",
        "center",
    )
    put(ws, "E14", '=IF(OR(C14="",NOT(ISNUMBER(C14))),"",IF(ABS(C14-\'05_DCF\'!E18)<0.0001,"与经营EV一致","与经营EV不一致"))', "output")
    ws.merge_cells("E14:H14")
    put(ws, "A16", "怎么读", "section")
    ws.merge_cells("A16:H16")
    put(
        ws,
        "A17",
        "手册印出的九格是14%/16%/18%对2%/3%/4%。业务还没进入稳态时，这张表本身就是投委会材料，不要只报中间那一格。若2031年仍在高增长，应先拉长预测期，而不是把g调高。",
        "note",
    )
    ws.merge_cells("A17:H18")
    heights(ws, {17: 22, 18: 22})
    finish(ws, ORANGE_DK, freeze="C8", fit_height=1)
    return ws


def build_market(wb):
    ws = wb.create_sheet("07_市场法")
    widths(ws, {get_column_letter(i): w for i, w in enumerate([28, 18, 16, 12, 14, 14, 14, 14, 14, 14, 14, 14, 16, 14], 1)})
    banner(ws, "N", "市场法", "3.0倍和-20%是手册里的纯演示，用来检查差额能不能被解释。空的可比表不会自动覆盖这个演示。")
    put(ws, "A4", "NTM收入", "label")
    put(ws, "C4", "='04_AI预测'!C6", "output", "0.000", "center")
    put(ws, "D4", "亿元，取2027年预测收入", "note", size=8)
    put(ws, "A5", "手工演示倍数", "label")
    put(ws, "C5", 3.0, "input", "0.00", "center", "纯演示。不是AI行业实际倍数。")
    put(ws, "A6", "经营质量调整", "label")
    put(ws, "C6", -0.20, "input", "0.0%", "center", "必须写得清这20%来自增长、利润、风险还是流动性。不能再叠一个IPO折价。")
    put(ws, "A7", "采用哪一个倍数", "label")
    put(ws, "C7", "手工演示倍数", "input", horizontal="center")
    choice(ws, "C7", ["手工演示倍数", "可比中位数"])
    put(ws, "A8", "采用倍数", "label")
    put(ws, "C8", '=IF(C7="可比中位数",IF(C30="","",C30),C5)', "output", "0.00", "center")
    put(ws, "A9", "原始指示EV", "label")
    put(ws, "C9", '=IF(OR(C4="",C8=""),"",C4*C8)', "output", "0.000", "center")
    put(ws, "A10", "调整后指示EV", "label")
    put(ws, "C10", '=IF(C9="","",C9*(1+C6))', "output", "0.000", "center")
    put(ws, "A11", "DCF经营EV", "label")
    put(ws, "C11", "='05_DCF'!E18", "output", "0.000", "center")
    put(ws, "A12", "差额（市场法－DCF）", "label")
    put(ws, "C12", '=IF(C10="","",C10-C11)', "output", "0.000", "center")
    put(ws, "D12", "差额要解释，不能靠加权把它藏起来。两种方法都无效时，不要为了得出一个数而赋权。", "note", size=8)
    ws.merge_cells("D12:G12")

    merge(ws, "A14:N14", "可比公司（正式使用时填写。货币和会计口径要先对齐）", "section", 11)
    headers = ["代码", "公司", "纳入", "货币", "市值", "净债务", "EV", "NTM收入", "NTM毛利", "NTM EBITDA", "EV/Sales", "EV/GP", "EV/EBITDA", "纳入后的EV/Sales"]
    for col, text in enumerate(headers, 1):
        put(ws, f"{get_column_letter(col)}15", text, "header", horizontal="center", size=8)
    choice_cells = []
    for row in range(16, 26):
        for col in range(1, 7):
            put(ws, f"{get_column_letter(col)}{row}", None, "input", "#,##0.00" if col in (5, 6) else None, "center")
        put(ws, f"C{row}", None, "input", horizontal="center")
        choice_cells.append(f"C{row}")
        put(ws, f"G{row}", f'=IF(OR(E{row}="",F{row}=""),"",E{row}+F{row})', "formula", "#,##0.00", "center")
        for col in ("H", "I", "J"):
            put(ws, f"{col}{row}", None, "input", "#,##0.00", "center")
        put(ws, f"K{row}", f'=IF(OR(G{row}="",H{row}="",H{row}=0),"",G{row}/H{row})', "formula", "0.00", "center")
        put(ws, f"L{row}", f'=IF(OR(G{row}="",I{row}="",I{row}=0),"",G{row}/I{row})', "formula", "0.00", "center")
        put(ws, f"M{row}", f'=IF(OR(G{row}="",J{row}="",J{row}=0),"",G{row}/J{row})', "formula", "0.00", "center")
        put(ws, f"N{row}", f'=IF(AND(C{row}="是",ISNUMBER(K{row})),K{row},"")', "formula", "0.00", "center")
    dv = DataValidation(type="list", formula1='"是,否"', allow_blank=True)
    for addr in choice_cells:
        dv.add(ws[addr])
    ws.add_data_validation(dv)
    put(ws, "A30", "纳入样本的EV/Sales中位数", "label")
    put(ws, "C30", '=IF(COUNT(N16:N25)=0,"",MEDIAN(N16:N25))', "output", "0.00", "center")
    put(ws, "D30", "只有C7选成“可比中位数”且这里有数，演示倍数才会被替换。", "note", size=8)
    ws.merge_cells("D30:G30")
    finish(ws, "8A5A32", freeze="A16", fit_height=1)
    return ws


def build_equity(wb):
    ws = wb.create_sheet("08_股权桥")
    widths(ws, {"A": 34, "B": 18, "C": 18, "D": 52})
    banner(ws, "D", "从经营EV到股权，以及另一条回调桥", "上半部分接DCF。下半部分是手册第10.5节的结构演示，两套数字不要加总。")
    merge(ws, "A4:D4", "一、经营EV到全部股权价值", "section", 11)
    rows = [
        (5, "经营EV", "='05_DCF'!E18", "output"),
        (6, "超额现金", 0, "input"),
        (7, "非经营资产", 0, "input"),
        (8, "有息负债", 0, "input"),
        (9, "租赁负债", 0, "input"),
        (10, "少数股东及其他索取权", 0, "input"),
        (11, "未纳入预测的或有项目", 0, "input"),
    ]
    for row, label, value, kind in rows:
        put(ws, f"A{row}", label, "label")
        fmt = "0.000"
        comment = "演示先填0，表示手册没有给出这项。正式项目必须换成估值日数字，0不是审计结论。" if kind == "input" else None
        put(ws, f"C{row}", value, kind, fmt, "center", comment)
    put(ws, "A12", "全部股权价值", "label")
    put(ws, "C12", "=C5+C6+C7-C8-C9-C10-C11", "output", "0.000", "center")
    put(ws, "A13", "桥项状态", "label")
    put(ws, "C13", '=IF(COUNT(C6:C11)<6,"缺少桥项",IF(SUM(C6:C11)=0,"六项仍为演示0，不能把C12写成正式股权价值","已替换演示0"))', "warn")
    ws.merge_cells("C13:D13")

    merge(ws, "A16:D16", "二、回调结构演示：冻结其他因素，一次只换一个因素", "section", 11)
    bridge = [
        (17, "上期股权价值", 8.00),
        (18, "经营增长下修", -1.10),
        (19, "算力毛利恶化", -0.45),
        (20, "可比倍数下降", -0.80),
        (21, "融资净现金及其他", 0.25),
    ]
    for row, label, value in bridge:
        put(ws, f"A{row}", label, "label")
        put(ws, f"C{row}", value, "example", "0.00", "center", "结构演示，不是任何真实公司的减项。")
    put(ws, "A22", "计算得到的本期", "label")
    put(ws, "C22", "=SUM(C17:C21)", "output", "0.00", "center")
    put(ws, "A23", "若另报一个新估值", "label")
    put(ws, "C23", 5.90, "example", "0.00", "center")
    put(ws, "A24", "残差", "label")
    put(ws, "C24", "=C23-C22", "output", "0.00", "center")
    put(ws, "A25", "勾稽", "label")
    put(ws, "C25", '=IF(ABS(C24)<0.0001,"桥已轧平","残差未解释")', "output", horizontal="center")
    put(ws, "D17", "顺序变了，交互项的归属要单独写明。", "note", size=8)
    finish(ws, "0E2A24", fit_height=1)
    return ws


def build_dilution(wb):
    ws = wb.create_sheet("09_稀释与IPO")
    widths(ws, {"A": 36, "B": 16, "C": 20, "D": 48})
    banner(ws, "D", "普通股稀释与IPO股数", "这里没有优先权、期权池和交易费用。有这些条款时，不能用新现金除以融资后价值当作普通股比例。")
    merge(ws, "A4:D4", "一、两轮普通股定价", "section", 11)
    put(ws, "A5", "情景一·融资前股权价值", "label")
    put(ws, "C5", 8, "input", "0.00", "center")
    put(ws, "A6", "新投入现金", "label")
    put(ws, "C6", 2, "input", "0.00", "center")
    put(ws, "A7", "融资后", "label")
    put(ws, "C7", "=C5+C6", "output", "0.00", "center")
    put(ws, "A8", "新投资者持股", "label")
    put(ws, "C8", '=IF(C7=0,"",C6/C7)', "output", "0.0%", "center")
    put(ws, "A9", "原股东持股", "label")
    put(ws, "C9", '=IF(C7=0,"",C5/C7)', "output", "0.0%", "center")
    put(ws, "A11", "情景二·融资前股权价值", "label")
    put(ws, "C11", 6, "input", "0.00", "center")
    put(ws, "A12", "新投入现金", "label")
    put(ws, "C12", "=C6", "formula", "0.00", "center")
    put(ws, "A13", "融资后", "label")
    put(ws, "C13", "=C11+C12", "output", "0.00", "center")
    put(ws, "A14", "新投资者持股", "label")
    put(ws, "C14", '=IF(C13=0,"",C12/C13)', "output", "0.0%", "center")
    put(ws, "A15", "原股东持股", "label")
    put(ws, "C15", '=IF(C13=0,"",C11/C13)', "output", "0.0%", "center")
    put(ws, "A16", "新投资者持股上升", "label")
    put(ws, "C16", '=IF(OR(C8="",C14=""),"",C14-C8)', "output", "0.0%", "center")
    put(ws, "D16", "手册默认输入下为5个百分点。优先权和ESOP会让实际分配不同。", "note", size=8)

    merge(ws, "A18:D18", "二、IPO股数与市值（空白就保持不可评估）", "section", 11)
    ipo = [
        (19, "IPO前等价股数", None, "0.0000"),
        (20, "发行价（港元/股）", None, "0.00"),
        (21, "募资总额（港元）", None, "#,##0"),
        (22, "发行费用（港元）", None, "#,##0"),
    ]
    for row, label, value, fmt in ipo:
        put(ws, f"A{row}", label, "label")
        put(ws, f"C{row}", value, "input", fmt, "center")
    put(ws, "A23", "新发股份", "label")
    put(ws, "C23", '=IF(OR(C20="",C20=0,C21=""),"",C21/C20)', "formula", "#,##0.00", "center")
    put(ws, "A24", "IPO后等价股数", "label")
    put(ws, "C24", '=IF(OR(C19="",C23=""),"",C19+C23)', "formula", "#,##0.00", "center")
    put(ws, "A25", "发行市值（港元）", "label")
    put(ws, "C25", '=IF(C24="","",C20*C24)', "output", '"HK$"#,##0', "center")
    put(ws, "A26", "净募集（港元）", "label")
    put(ws, "C26", '=IF(C21="","",C21-IF(C22="",0,C22))', "output", '"HK$"#,##0', "center")
    put(ws, "A27", "本项目最低市值门槛", "label")
    put(ws, "C27", "='02_18C合规'!C26", "formula", '"HK$"#,##0', "center")
    put(ws, "A28", "门槛对照", "label")
    put(ws, "C28", '=IF(OR(C25="",C27=""),"NOT-ASSESSABLE",IF(C25>=C27,"CAP-PASS","CAP-FAIL"))', "output", horizontal="center")
    put(ws, "A30", "情景一融资前价值折合亿港元", "label")
    put(ws, "C30", "=C5*'01_使用说明'!C14", "formula", "0.00", "center")
    put(ws, "D30", "只是币种换算。它不是发行市值，不能拿去过40亿或80亿的测试。", "note", size=8)
    finish(ws, "1A433A", fit_height=1)
    return ws


def industry_banner_rows(ws, last, title, subtitle, rows):
    widths(ws, {"A": 38, "B": 16, "C": 18, "D": 62})
    banner(ws, last, title, subtitle)
    for row, label, value, kind, fmt, note in rows:
        put(ws, f"A{row}", label, "label")
        put(ws, f"C{row}", value, kind, fmt, "center")
        put(ws, f"D{row}", note, "note", size=8)


def build_hardware(wb):
    ws = wb.create_sheet("10_硬科技")
    industry_banner_rows(
        ws,
        "D",
        "硬科技公式骨架",
        "自检数只证明“投片×裸片×良率×交付率×ASP”这条链。它不是芯片公司估值。",
        [
            (4, "晶圆投片量（片）", 1000, "example", "#,##0", "公式自检"),
            (5, "每片理论裸片", 500, "example", "#,##0", "公式自检"),
            (6, "可售良率", 0.80, "example", "0.0%", "公式自检"),
            (7, "交付率", 0.90, "example", "0.0%", "公式自检"),
            (8, "ASP（元/颗）", 10, "example", "0.00", "不是市场价格"),
            (9, "可售芯片量", "=C4*C5*C6*C7", "output", "#,##0.00", "片数×裸片×良率×交付率"),
            (10, "芯片收入（元）", "=C9*C8", "output", "#,##0.00", "再乘ASP"),
            (12, "量产延期（月）", None, "input", "0", "只记录。不会自动把收入后移"),
            (13, "ASP变动", None, "input", "0.0%", "例如-15%就填-15%"),
            (14, "良率变动（加法）", None, "input", "0.0%", "下降10个百分点填-10%"),
            (15, "订单实现率", None, "input", "0.0%", "留空则下面按100%看压力收入"),
            (16, "压力后收入（元）", '=C4*C5*(C6+IF(C14="",0,C14))*C7*(C8*(1+IF(C13="",0,C13)))*IF(C15="",1,C15)', "output", "#,##0.00", "延期月数不在这个乘法里"),
        ],
    )
    merge(ws, "A18:D18", "回调顺序：先改客户认证和流片成功的时间，再改单位成本和现金消耗，最后才重算倍数。不要把每个风险溢价同时上调。", "note", 9)
    heights(ws, {18: 32})
    finish(ws, "3E4A45", fit_height=1)


def build_energy(wb):
    ws = wb.create_sheet("11_新能源")
    industry_banner_rows(
        ws,
        "D",
        "新能源单位经济骨架",
        "产能、利用率和单价是公式自检，不是电芯或系统的市场价格。",
        [
            (4, "产能（MWh）", 1000, "example", "#,##0", "自检"),
            (5, "利用率", 0.80, "example", "0.0%", "自检"),
            (6, "合格品率", 0.95, "example", "0.0%", "自检"),
            (7, "销售实现率", 0.90, "example", "0.0%", "自检"),
            (8, "单价（元/MWh）", 1000, "example", "#,##0.00", "任意自检单价"),
            (9, "出货量（MWh）", "=C4*C5*C6*C7", "output", "#,##0.00", ""),
            (10, "收入（元）", "=C9*C8", "output", "#,##0.00", ""),
            (12, "ASP变动", None, "input", "0.0%", "与原料分开调"),
            (13, "利用率变动", None, "input", "0.0%", "加减的是百分点"),
            (14, "原料价格变动", None, "input", "0.0%", "不要把它写成收入增速"),
            (15, "压力后出货", "=C4*(C5+IF(C13=\"\",0,C13))*C6*C7", "output", "#,##0.00", ""),
            (16, "压力后收入", "=C15*C8*(1+IF(C12=\"\",0,C12))", "output", "#,##0.00", "原料变动要另改成本，不在收入公式里"),
        ],
    )
    merge(ws, "A18:D18", "持有长期能源项目的公司按项目期限做现金流，不要因为集团叙事使用永续高增长。价格和原料分开回调。", "note", 9)
    heights(ws, {18: 32})
    finish(ws, "0F6E62", fit_height=1)


def build_bio(wb):
    ws = wb.create_sheet("12_生物医药")
    widths(ws, {"A": 28, "B": 22, "C": 18, "D": 18, "E": 18, "F": 18, "G": 36})
    banner(ws, "G", "rNPV", "已经用阶段成功概率调整过的现金流，不要再把同一临床失败加进折现率。0.30亿元是算术演示，不是临床成功率。")
    put(ws, "A4", "手册单项目演示", "section")
    ws.merge_cells("A4:C4")
    put(ws, "A5", "全部成功时的现值", "label")
    put(ws, "C5", 12, "example", "0.00", "center", "亿元。虚构。")
    put(ws, "A6", "演示用全流程成功概率", "label")
    put(ws, "C6", 0.15, "example", "0.0%", "center", "不是任何临床阶段的真实成功率。")
    put(ws, "A7", "按阶段加权的研发支出现值", "label")
    put(ws, "C7", 1.5, "example", "0.00", "center")
    put(ws, "A8", "rNPV", "label")
    put(ws, "C8", "=C5*C6-C7", "output", "0.00", "center")
    put(ws, "D8", "亿元人民币", "note")

    merge(ws, "A10:G10", "管线表：第一行预填演示，其余留给项目。平台服务收入不要再加进这些格子。", "section", 11)
    for col, text in enumerate(("资产", "成功情形现值", "成功概率", "研发支出现值", "rNPV", "备注"), 1):
        put(ws, f"{get_column_letter(col)}11", text, "header", horizontal="center")
    put(ws, "A12", "手册演示管线", "example")
    put(ws, "B12", "=C5", "formula", "0.00", "center")
    put(ws, "C12", "=C6", "formula", "0.0%", "center")
    put(ws, "D12", "=C7", "formula", "0.00", "center")
    put(ws, "E12", "=B12*C12-D12", "output", "0.00", "center")
    put(ws, "F12", "与C8同一条算式", "note", size=8)
    for row in range(13, 17):
        put(ws, f"A{row}", None, "input")
        put(ws, f"B{row}", None, "input", "0.00", "center")
        put(ws, f"C{row}", None, "input", "0.0%", "center")
        put(ws, f"D{row}", None, "input", "0.00", "center")
        put(ws, f"E{row}", f'=IF(OR(B{row}="",C{row}="",D{row}=""),"",B{row}*C{row}-D{row})', "output", "0.00", "center")
        put(ws, f"F{row}", None, "input")
    put(ws, "A18", "管线合计", "label")
    put(ws, "E18", "=SUM(E12:E16)", "output", "0.00", "center")
    merge(ws, "A20:G20", "18A最低市值15亿港元，不能套到18C。核心产品若是需临床试验批准的药物，先走18A分流，不要因为使用了AI就改章节。", "warn", 9)
    heights(ws, {20: 32})
    finish(ws, RED_DK, fit_height=1)


def build_equipment(wb):
    ws = wb.create_sheet("13_高端装备")
    industry_banner_rows(
        ws,
        "D",
        "高端装备订单骨架",
        "设备、软件和服务要分开。签单额不是当期收入。",
        [
            (4, "出货台数", 100, "example", "#,##0", "自检"),
            (5, "实现售价（元）", 100000, "example", "#,##0", "自检"),
            (6, "验收确认系数", 0.80, "example", "0.0%", "未验收的不确认"),
            (7, "新设备收入（元）", "=C4*C5*C6", "output", "#,##0", ""),
            (9, "期初在手订单", 1000, "example", "#,##0.00", "自检，单位自定"),
            (10, "新签有效订单", 400, "example", "#,##0.00", ""),
            (11, "已确认交付", 300, "example", "#,##0.00", ""),
            (12, "取消", 50, "example", "#,##0.00", ""),
            (13, "期末在手订单", "=C9+C10-C11-C12", "output", "#,##0.00", ""),
            (15, "装机保有量", None, "input", "#,##0", "正式项目再填"),
            (16, "服务渗透率", None, "input", "0.0%", ""),
            (17, "年度ARPU", None, "input", "#,##0.00", ""),
            (18, "维护及耗材收入", '=IF(OR(C15="",C16="",C17=""),"",C15*C16*C17)', "output", "#,##0.00", ""),
        ],
    )
    merge(ws, "A20:D20", "验收延期要同时改收入确认、回款、质保和营运资金。只把销售额下调，会漏掉现金。", "note", 9)
    finish(ws, "5E6A64", fit_height=1)


def build_sme(wb):
    ws = wb.create_sheet("14_专精特新")
    industry_banner_rows(
        ws,
        "D",
        "专精特新不是18C行业",
        "先回到半导体、装备、材料或新能源里的具体技术，再决定用不用18C。",
        [
            (4, "客户产量或采购预算", 100, "example", "#,##0.00", "自检"),
            (5, "公司份额", 0.20, "example", "0.0%", "自检"),
            (6, "有效售价系数", 1, "example", "0.00", "自检"),
            (7, "品类收入", "=C4*C5*C6", "output", "#,##0.00", ""),
            (9, "报告EBITDA", None, "input", "#,##0.00", ""),
            (10, "有证据的非经常调整", None, "input", "#,##0.00", "加项为正，减项为负"),
            (11, "正常化EBITDA", '=IF(C9="","",C9+IF(C10="",0,C10))', "output", "#,##0.00", ""),
        ],
    )
    merge(ws, "A13:D13", "四个常见陷阱，填“是/否”。空白表示还没检查。", "section", 11)
    questions = [
        (14, "是否把政府资质直接当成估值溢价"),
        (15, "是否用专利件数代替现金流质量"),
        (16, "是否已单独建客户集中和应收账款"),
        (17, "是否把净利润、净现金和资本化研发重复相加"),
    ]
    cells = []
    for row, text in questions:
        put(ws, f"A{row}", text, "label")
        put(ws, f"C{row}", None, "input", horizontal="center")
        cells.append(f"C{row}")
    yes_no(ws, cells)
    put(ws, "D14", "期望答案：否", "note", size=8)
    put(ws, "D15", "期望答案：否", "note", size=8)
    put(ws, "D16", "期望答案：是", "note", size=8)
    put(ws, "D17", "期望答案：否", "note", size=8)
    finish(ws, "8A5A32", fit_height=1)


def build_methods(wb):
    ws = wb.create_sheet("15_方法库")
    widths(ws, {"A": 28, "B": 42, "C": 55, "D": 55})
    banner(ws, "D", "方法选择", "只有口径一致、已经校准、而且确实有信息的方法才能加权。权重不是会计准则。")
    for col, text in enumerate(("方法", "主要用途", "核心关系", "不要这样做"), 1):
        put(ws, f"{get_column_letter(col)}4", text, "header")
    rows = [
        ("FCFF-DCF", "收入、成本、资本开支和营运资金能建模型", "EV = Σ FCFF/(1+WACC)^t + TV/(1+WACC)^N", "未商业化不要让终值主导还不自知"),
        ("FCFE-DCF", "股权现金流和债务政策都稳定", "股权价值用Ke折现FCFE", "FCFE不能配WACC"),
        ("可比公司", "18C市场定价、IPO比较", "同一期间、同一口径的指标 × 调整后倍数", "不要拿某一天的高点"),
        ("可比交易", "控制权收购", "交易倍数，处理控制和协同", "协同溢价不要自动用于少数股权"),
        ("最近融资校准", "创投和Pre-IPO", "用当时的价格校准技术，再用今天的输入重估", "融资价格不自动等于今天的公允价值"),
        ("rNPV", "创新药和未成熟项目", "成功概率后的现金流现值，减去阶段研发", "临床风险不要在概率和折现率里各罚一次"),
        ("PWERM", "前商业化的多条路径", "路径概率 × 该路径证券回款的现值", "概率要和优先权放在一起算"),
        ("OPM", "多轮优先股", "按股权价值、波动率、期限和分配断点", "不要用公司EV直接代替每一类股份"),
        ("调整净资产", "重资产或清算下限", "资产公允价值减负债公允价值", "无形资产和未来利润会被看低"),
        ("VC退出法", "极早期交叉检验", "退出股权价值 / (1+目标回报)^n", "后续稀释和生存概率不要扣两次"),
        ("分部加总", "软件、硬件和项目混在一起", "各分部EV，减去中央费用和净债务", "内部交易和共享成本不要算两次"),
    ]
    for i, row in enumerate(rows):
        for col, text in enumerate(row, 1):
            put(ws, f"{get_column_letter(col)}{5+i}", text, "plain", size=9)
        heights(ws, {5 + i: 32})
    finish(ws, NAVY, freeze="A5", fit_height=1)


def build_callback(wb):
    ws = wb.create_sheet("16_回调手册")
    widths(ws, {"A": 28, "B": 36, "C": 36, "D": 36})
    banner(ws, "D", "回调：四件事不要合成一个折价", "经济估值再校准、融资价格重谈、合同反稀释、IPO簿记，是四条不同的工作。")
    steps = [
        ("0", "锁定上期版本", "上期估值日和上期价值", "已批准的模型和报告"),
        ("1", "核对新的财务实际", "收入、毛利、研发、资本开支、现金、负债", "新报表、总账、银行流水"),
        ("2", "经营差异归因", "ARR、Token、订单、ASP、流失、临床概率", "合同、分群、工程或临床报告"),
        ("3", "时间和资金", "商业化时间、下一轮、12个月跑道", "月度现金、薪酬和算力预算"),
        ("4", "重抽同期市场", "可比池、LTM/NTM、倍数", "截止日行情和净债务"),
        ("5", "更新资本成本和终值", "无风险利率、ERP、Beta、g、稳态利润", "市场观测和委员会批准"),
        ("6", "重算并检验", "DCF、市场法、PWERM或rNPV", "校验页没有未解释异常"),
        ("7", "拆变化", "经营、倍数、汇率、时间、净债务、优先权", "变动桥残差为零"),
        ("8", "出结论", "区间、每股、触发条件", "签核人和数据日"),
    ]
    for col, text in enumerate(("步骤", "做什么", "更新什么", "留下什么"), 1):
        put(ws, f"{get_column_letter(col)}4", text, "header")
    for i, row in enumerate(steps):
        for col, text in enumerate(row, 1):
            put(ws, f"{get_column_letter(col)}{5+i}", text, "plain", size=9)
    merge(ws, "A15:D15", "事实出现以后，先动经营，再动倍数", "section", 11)
    triggers = [
        ("调用增加但收费率下降", "净收费Token、实现单价、免费量", "单位成本和毛利", "直接提高PS"),
        ("客户流失", "NRR、获客成本、合同额", "长期利润率和可比池", "只上调WACC"),
        ("新的市场化融资", "条款、价格和摊薄", "回看DCF隐含了什么", "把优先股价格当成普通股价格"),
        ("流片失败或延迟", "量产概率和时间", "追加研发和现金跑道", "继续用一个不变的永续增长"),
        ("产品降价", "ASP和原料分开", "库存、应收、质保", "利润按降价幅度等比例下降"),
        ("临床失利", "该阶段概率和路径", "研发、授权和定价的时间", "同一失败再加一次风险溢价"),
        ("验收延期", "确认收入的时间", "回款、质保、营运资金", "把签单额当成当期收入"),
        ("利率或市场倍数下移", "同期倍数、无风险利率、ERP", "融资窗口和现金情景", "没有证据就打七折"),
        ("离18C门槛还很远", "结论、阶段和路径", "延期、其他上市可能", "用门槛倒推终值"),
    ]
    for col, text in enumerate(("看到的事实", "先调", "接着调", "通常不要先做"), 1):
        put(ws, f"{get_column_letter(col)}16", text, "header")
    for i, row in enumerate(triggers):
        for col, text in enumerate(row, 1):
            put(ws, f"{get_column_letter(col)}{17+i}", text, "plain", size=9)
        heights(ws, {17 + i: 28})
    finish(ws, ORANGE_DK, freeze="A5", fit_height=1)


def build_qa(wb):
    ws = wb.create_sheet("17_校验")
    widths(ws, {"A": 8, "B": 42, "C": 16, "D": 70})
    banner(ws, "D", "恒等式校验", "异常应为0。提示是手册自己写明的局限，不是公式坏了。")
    put(ws, "A4", "异常", "header", horizontal="center")
    put(ws, "C4", '=COUNTIF(C7:C40,"异常")', "output", "0", "center")
    put(ws, "B4", "个身份或规则示例没有通过", "label")
    put(ws, "B5", "条需要阅读的提示", "label")
    put(ws, "C5", '=COUNTIF(C7:C40,"提示")', "warn", "0", "center")
    for col, text in enumerate(("编号", "检查", "结果", "怎么读"), 1):
        put(ws, f"{get_column_letter(col)}6", text, "header")

    checks = [
        ("1", "收入桥可以独立复算", '=IF(ABS(\'03_AI输入\'!C12-((\'03_AI输入\'!C6+\'03_AI输入\'!C6*\'03_AI输入\'!C7+\'03_AI输入\'!C8)/2+\'03_AI输入\'!C11))<0.0000001,"通过","异常")', "期末ARR、平均ARR和Token应加回2026收入"),
        ("2", "WACC大于g", '=IF(\'05_DCF\'!B3>\'05_DCF\'!B4,"通过","异常")', "否则终值没有定义"),
        ("3", "2027收入等于上年乘增长", '=IF(ABS(\'04_AI预测\'!C6-\'04_AI预测\'!B6*(1+\'04_AI预测\'!C5))<0.0000001,"通过","异常")', ""),
        ("4", "2031收入链式相乘", '=IF(ABS(\'04_AI预测\'!G6-\'04_AI预测\'!F6*(1+\'04_AI预测\'!G5))<0.0000001,"通过","异常")', ""),
        ("5", "2027 EBITDA恒等式", '=IF(ABS(\'04_AI预测\'!C11-\'04_AI预测\'!C20)<0.0000001,"通过","异常")', "交叉格在预测表"),
        ("6", "2027 FCFF恒等式", '=IF(ABS(\'04_AI预测\'!C18-(\'04_AI预测\'!C15+\'04_AI预测\'!C12-\'04_AI预测\'!C16-\'04_AI预测\'!C17))<0.0000001,"通过","异常")', ""),
        ("7", "亏损年不征经营现金税", '=IF(\'04_AI预测\'!C13>=0,"通过",IF(ABS(\'04_AI预测\'!C14)<0.0000001,"通过","异常"))', "税损结转仍未建模"),
        ("8", "经营EV等于两段现值之和", '=IF(ABS(\'05_DCF\'!E18-(\'05_DCF\'!E14+\'05_DCF\'!E17))<0.0000001,"通过","异常")', ""),
        ("9", "戈登终值公式", '=IF(ABS(\'05_DCF\'!E16-\'05_DCF\'!E15*(1+\'05_DCF\'!B4)/(\'05_DCF\'!B3-\'05_DCF\'!B4))<0.0000001,"通过","异常")', ""),
        ("10", "敏感性基准格等于经营EV", '=IF(AND(ISNUMBER(\'06_敏感性\'!C14),ABS(\'06_敏感性\'!C14-\'05_DCF\'!E18)<0.0001),"通过","异常")', "轴被改到没有当前WACC时会异常，这是提醒"),
        ("11", "市场法差额勾稽", '=IF(ABS(\'07_市场法\'!C12-(\'07_市场法\'!C10-\'07_市场法\'!C11))<0.0000001,"通过","异常")', ""),
        ("12", "回调桥残差", '=IF(ABS(\'08_股权桥\'!C24)<0.0001,"通过","异常")', "另报新估值若改掉，残差必须解释"),
        ("13", "情景一持股合计为1", '=IF(ABS(\'09_稀释与IPO\'!C8+\'09_稀释与IPO\'!C9-1)<0.0000001,"通过","异常")', ""),
        ("14", "示例甲为已商业化且表内通过", '=IF(AND(\'02_18C合规\'!D25="COMMERCIAL",\'02_18C合规\'!D40="PASS"),"通过","异常")', "用来确认规则公式没有被改坏"),
        ("15", "示例乙为未商业化且表内通过", '=IF(AND(\'02_18C合规\'!E25="PRE-COMMERCIAL",\'02_18C合规\'!E40="PASS"),"通过","异常")', "含50%研发比率和125%营运资金"),
        ("16", "本项目空白时不可评估", '=IF(\'02_18C合规\'!C6="",IF(\'02_18C合规\'!C40="NOT-ASSESSABLE","通过","异常"),"通过")', "开始填写以后，这条不再强制空白状态"),
        ("17", "rNPV算式", '=IF(ABS(\'12_生物医药\'!C8-(\'12_生物医药\'!C5*\'12_生物医药\'!C6-\'12_生物医药\'!C7))<0.0000001,"通过","异常")', ""),
        ("18", "芯片数量算式", '=IF(ABS(\'10_硬科技\'!C9-\'10_硬科技\'!C4*\'10_硬科技\'!C5*\'10_硬科技\'!C6*\'10_硬科技\'!C7)<0.000001,"通过","异常")', ""),
        ("19", "储能出货算式", '=IF(ABS(\'11_新能源\'!C9-\'11_新能源\'!C4*\'11_新能源\'!C5*\'11_新能源\'!C6*\'11_新能源\'!C7)<0.000001,"通过","异常")', ""),
        ("20", "终值是否压过经营EV", '=IF(\'05_DCF\'!E17>\'05_DCF\'!E18,"提示","通过")', "演示案例预期就是提示"),
        ("21", "汇率是否仍为演示", '=IF(\'01_使用说明\'!C12="","提示","通过")', "填了正式汇率后变为通过"),
        ("22", "股权桥是否仍是演示零", '=IF(\'08_股权桥\'!C13="六项仍为演示0，不能把C12写成正式股权价值","提示","通过")', ""),
        ("23", "税损结转", '"提示"', "手册明确本演示没有结转亏损"),
        ("24", "市场倍数是否仍是手工演示", '=IF(AND(\'07_市场法\'!C7="手工演示倍数",ABS(\'07_市场法\'!C5-3)<0.0001),"提示","通过")', "换成证据倍数后，这条提示会消失"),
        ("25", "专精特新陷阱有没有填反",
         '=IF(COUNTIF(\'14_专精特新\'!C14:C17,"")=4,"提示",IF(OR(\'14_专精特新\'!C14="是",\'14_专精特新\'!C15="是",\'14_专精特新\'!C16="否",\'14_专精特新\'!C17="是"),"提示","通过"))',
         "空白或填反都只提示，不假装已经检查过"),
    ]
    for i, (num, name, formula, note) in enumerate(checks):
        row = 7 + i
        put(ws, f"A{row}", num, "note", horizontal="center")
        put(ws, f"B{row}", name, "label")
        put(ws, f"C{row}", formula, "output", horizontal="center")
        put(ws, f"D{row}", note, "note", size=8)
        ws.conditional_formatting.add(f"C{row}", CellIsRule(operator="equal", formula=['"异常"'], fill=fill(RED), font=fnt(10, True, RED_DK)))
        ws.conditional_formatting.add(f"C{row}", CellIsRule(operator="equal", formula=['"提示"'], fill=fill(ORANGE), font=fnt(10, True, ORANGE_DK)))
        ws.conditional_formatting.add(f"C{row}", CellIsRule(operator="equal", formula=['"通过"'], fill=fill(GREEN), font=fnt(10, True, GREEN_DK)))
    finish(ws, RED_DK, freeze="A7", fit_height=1)


def build_sources(wb):
    ws = wb.create_sheet("18_来源")
    widths(ws, {"A": 14, "B": 78, "C": 18, "D": 78})
    banner(ws, "D", "公开来源", "链接是入口，不是转载授权。行情和倍数要在估值日重新下载。")
    for col, text in enumerate(("编号", "材料", "入口", "地址"), 1):
        put(ws, f"{get_column_letter(col)}4", text, "header")
    sources = [
        ("HKEX-1", "香港交易所《18C, Explained》（2026-05-04）", "https://www.hkexgroup.com/Media-Centre/Insight/Insight/2026/HKEX-Insight/18C-Explained?sc_lang=en"),
        ("HKEX-2", "港交所《18C.03》", "https://en-rules.hkex.com.hk/rulebook/18c03"),
        ("HKEX-3", "18C临时市值门槛公告（2024-08-23）", "https://www.hkex.com.hk/News/Regulatory-Announcements/2024/240823news?sc_lang=en"),
        ("HKEX-4", "港交所《18C.04》研发", "https://en-rules.hkex.com.hk/rulebook/18c04"),
        ("HKEX-5", "上市申请人指引第2.5章", "https://www.hkex.com.hk/listing/rules-and-resources/guidance/ipo/listing-of-specialist-technology-companies?sc_lang=en"),
        ("HKEX-6", "港交所《18C.06》商业化路径", "https://en-rules.hkex.com.hk/rulebook/18c06"),
        ("HKEX-7", "港交所《18C.07》营运资金", "https://en-rules.hkex.com.hk/rulebook/18c07"),
        ("HKEX-8", "港交所《18C.08》独立定价投资者", "https://en-rules.hkex.com.hk/rulebook/18c08"),
        ("HKEX-9", "港交所《18A》生物科技", "https://en-rules.hkex.com.hk/entiresection/5193"),
        ("IPEV-1", "IPEV估值指引，2025年12月版", "https://www.privateequityvaluation.com/Valuation-Guidelines"),
        ("PWC-1", "PwC Alternatives Valuation Insights 2025", "https://www.pwc.com/us/en/industries/financial-services/library/pdf/pwc-awm-benchmarking-insights-responding-to-market-uncertainty-alternatives-valuation-snapshot-2025.pdf"),
        ("PWC-2", "PwC AI in deals 2026", "https://www.pwc.com/us/en/services/consulting/deals/library/deals-age-ai-acceleration-vs-erosion-value.html"),
        ("DELOITTE-1", "德勤瑞典组合估值调查 2025", "https://www.deloitte.com/content/dam/assets-zone2/se/sv/docs/services/consulting/2024/PVS%20Survey%20Sweden%202025_vF.pdf"),
        ("DELOITTE-2", "德勤 2025 Fair Valuation Pricing Survey", "https://www.deloitte.com/us/en/industries/financial-services/articles/annual-fair-valuation-survey.html"),
        ("EY-1", "安永对IPEV的2023年解读（旧版）", "https://www.ey.com/content/dam/ey-unified-site/ey-com/en-in/services/strategy-transactions/images/valuation-modeling-economics/ey-nternational-private-equity-and-venture-capital-valuation-guidelines.pdf"),
        ("EY-2", "安永：创业公司DCF六步", "https://www.ey.com/en_nl/services/finance-navigator/startup-valuation-applying-the-discounted-cash-flow-method-in-six-easy-step"),
        ("EY-3", "安永 The AI valuation shift 2026", "https://www.ey.com/en_us/insights/strategy/the-ai-valuation-shift-a-point-of-view-for-industrial-ceos-and-their-deal-teams"),
        ("KPMG-1", "毕马威 Valuations Services 2019", "https://assets.kpmg.com/content/dam/kpmg/xx/pdf/2019/07/global-valuations-brochure.pdf"),
        ("KPMG-2", "毕马威对IPEV 2025的说明", "https://kpmg.com/dp/en/insights/2026/01/ipev-2025.html"),
        ("KPMG-3", "毕马威资本成本研究 2026", "https://kpmg.com/ch/en/insights/deals/cost-capital-study.html"),
        ("CASE-1", "晶泰科技2024年招股书", "https://www1.hkexnews.hk/listedco/listconews/sehk/2024/0604/2024060400065.pdf"),
        ("CASE-2", "越疆科技2024年招股书", "https://www1.hkexnews.hk/listedco/listconews/sehk/2024/1213/2024121300029.pdf"),
        ("CASE-3", "MiniMax 2026-03-13转商业化公告", "https://www.hkexnews.hk/listedco/listconews/sehk/2026/0313/2026031301613.pdf"),
        ("CASE-4", "壁仞科技2025年报", "https://www1.hkexnews.hk/listedco/listconews/sehk/2026/0422/2026042200045.pdf"),
        ("CASE-5", "Skadden 2026年18C上市公司报告", "https://www.skadden.com/insights/publications/2026/10/2026-report-on-hong-kong-listed-specialist-technology-companies"),
    ]
    for i, (code, title, url) in enumerate(sources):
        row = 5 + i
        put(ws, f"A{row}", code, "label", horizontal="center")
        put(ws, f"B{row}", title, "plain", size=9)
        cell = put(ws, f"C{row}", "打开", "plain", horizontal="center")
        cell.hyperlink = url
        cell.font = fnt(10, False, TEAL)
        put(ws, f"D{row}", url, "note", size=8)
        heights(ws, {row: 22})
    ws.auto_filter.ref = f"A4:D{4+len(sources)}"
    finish(ws, "5E6A64", freeze="A5", fit_height=2)
    ws.auto_filter.ref = f"A4:D{4+len(sources)}"


def build_report(wb):
    ws = wb.create_sheet("19_报告底稿")
    widths(ws, {"A": 32, "B": 70})
    banner(ws, "B", "项目估值报告底稿", "方括号留给项目。已经能从模型取出的数字用公式接上，叙事仍要人写。")
    fields = [
        (4, "企业名称", None),
        (5, "估值日期", date(2026, 11, 6)),
        (6, "估值对象", None),
        (7, "估值目的", None),
        (8, "币种", "人民币经营预测；港币用于规则测试"),
        (9, "财务资料", None),
        (10, "执行与复核", None),
        (12, "DCF经营EV（亿元）", "='05_DCF'!E18"),
        (13, "市场法调整后EV（亿元）", "='07_市场法'!C10"),
        (14, "方法差额（亿元）", "='07_市场法'!C12"),
        (15, "全部股权价值（亿元）", "='08_股权桥'!C12"),
        (16, "股权桥状态", "='08_股权桥'!C13"),
        (17, "18C本项目总灯", "='02_18C合规'!C40"),
        (18, "最大的三个变量", None),
        (20, "不采用某方法的理由", None),
        (21, "特专科技领域及证据", None),
        (22, "上次到本次的变动桥", "='08_股权桥'!C25"),
        (23, "重新估值的触发事件", None),
        (24, "尚未确认的事项", None),
    ]
    for row, label, value in fields:
        put(ws, f"A{row}", label, "label")
        if isinstance(value, str) and value.startswith("="):
            put(ws, f"B{row}", value, "output", "0.00" if "亿元" in label else None)
        elif value is None:
            put(ws, f"B{row}", None, "input")
        elif isinstance(value, date):
            put(ws, f"B{row}", value, "input", "YYYY-MM-DD")
        else:
            put(ws, f"B{row}", value, "input")
        heights(ws, {row: 24})
    finish(ws, GOLD, fit_height=1)


def build_dictionary(wb):
    ws = wb.create_sheet("20_参数字典")
    widths(ws, {"A": 16, "B": 32, "C": 16, "D": 18, "E": 55})
    banner(ws, "E", "参数字典", "数值连到输入格。改字典不会改模型，改输入格才会。")
    for col, text in enumerate(("代码", "含义", "单位", "当前值", "核验"), 1):
        put(ws, f"{get_column_letter(col)}4", text, "header")
    rows = [
        ("A01", "期初订阅ARR", "亿元", "='03_AI输入'!C6", "分月合同、开票、递延收入"),
        ("A02", "净收入留存", "倍", "='03_AI输入'!C7", "队列。不要重复扣流失"),
        ("A03", "新增ARR", "亿元", "='03_AI输入'!C8", "生效日和实施进度"),
        ("A06", "Token已确认收入", "亿元", "='03_AI输入'!C11", "计费日志，剔除赠送"),
        ("A07", "2026收入", "亿元", "='03_AI输入'!C12", "平均ARR＋Token"),
        ("B01", "2027收入增长", "%", "='03_AI输入'!C17", "与客户分群预测一致"),
        ("A11", "WACC", "%", "='03_AI输入'!C27", "演示值。正式项目重做资本成本"),
        ("A12", "终值增长率", "%", "='03_AI输入'!E27", "币种和稳态一致，且小于WACC"),
        ("A13", "所得税率", "%", "='03_AI输入'!C24", "按纳税主体。本演示无税损结转"),
        ("M01", "演示EV/NTM收入", "倍", "='07_市场法'!C5", "不是市场统计"),
        ("M02", "质量调整", "%", "='07_市场法'!C6", "写明调整归因"),
        ("FX", "采用汇率", "HKD/RMB", "='01_使用说明'!C14", "正式汇率优先"),
        ("R01", "rNPV演示", "亿元", "='12_生物医药'!C8", "不是临床基准"),
    ]
    for i, row in enumerate(rows):
        line = 5 + i
        for col, value in enumerate(row, 1):
            if col == 4:
                fmt = "0.00%" if row[2] == "%" else "0.00"
                put(ws, f"D{line}", value, "output", fmt, "center")
            else:
                put(ws, f"{get_column_letter(col)}{line}", value, "plain" if col > 2 else "label", size=9)
    finish(ws, "1F4E79", freeze="A5", fit_height=1)


def build():
    wb = Workbook()
    wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
    wb.properties.title = "港交所18C分行业估值模型"
    wb.properties.subject = "虚构演示与可复用公式，非投资建议"
    wb.properties.creator = "AIDEN"
    wb.properties.category = "估值模型"
    build_cover(wb)
    build_readme(wb)
    build_compliance(wb)
    build_ai_inputs(wb)
    build_forecast(wb)
    build_dcf(wb)
    build_sensitivity(wb)
    build_market(wb)
    build_equity(wb)
    build_dilution(wb)
    build_hardware(wb)
    build_energy(wb)
    build_bio(wb)
    build_equipment(wb)
    build_sme(wb)
    build_methods(wb)
    build_callback(wb)
    build_qa(wb)
    build_sources(wb)
    build_report(wb)
    build_dictionary(wb)
    # 摘要引用了 DCF 的 A18，但经营EV实际放在 E18。上面 cover 已直接引用 E 列的镜像前，我在 cover 写的是 A18。
    # 在 DCF 页把 A18 设成标签后，值在 E18。修正摘要公式。
    cover = wb["00_投委会摘要"]
    cover["B6"] = "='05_DCF'!E18"
    cover["B7"] = "='05_DCF'!E14"
    cover["B8"] = "='05_DCF'!E17"
    cover["B9"] = "='05_DCF'!E19"
    cover["B10"] = "='05_DCF'!B20"
    cover["AB6"] = "='05_DCF'!E18"
    wb["19_报告底稿"]["B12"] = "='05_DCF'!E18"
    wb["08_股权桥"]["C5"] = "='05_DCF'!E18"
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT)
    return OUTPUT


if __name__ == "__main__":
    print(build())
