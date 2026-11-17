# -*- coding: utf-8 -*-
"""生成合作工作表。输入格为浅金色，金额由公式计算。"""
from __future__ import annotations

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation

from coop_data import (
    AI_PLAN,
    ANNUAL_GOALS,
    BUSINESSES,
    CHANNELS,
    CLIENT_ACTIONS,
    CLIENTS,
    COMPETITORS,
    DECISION_FACTORS,
    EXISTING_AIGC,
    FUNNEL,
    INTERVIEW_FIELDS,
    KPI,
    META,
    MILESTONES,
    MODULES,
    MOU_ITEMS,
    NEAR_TERM_A,
    NEAR_TERM_B,
    PRICE_LADDER,
    PRIORITY_VISITS,
    PRODUCTS,
    PROJECTS_2027,
    QUARTERS,
    RISKS,
    ROLES,
    ROADMAP,
    ROUTES,
    SCALE_EXAMPLE,
    SCENARIOS,
    SEGMENTS,
    SOURCES,
    VISIT_STATUS,
    VISIT_TIERS,
    VISITS,
)

NAVY = "13294B"
CRIMSON = "8C1D2C"
GOLD = "8A6A32"
WHITE = "FFFFFF"
INK = "1C2430"
MUTED = "5E6A78"
PAPER = "F7F4EF"
ZEBRA = "F8F6F2"
SOFT_RED = "F8E8EA"
SOFT_GOLD = "F8F3E8"
SOFT_NAVY = "E8EEF5"
SOFT_GREEN = "E7F3EC"
INPUT = "FFF4D6"
LINE = "D9D3C8"

thin = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
fill_navy = PatternFill("solid", fgColor=NAVY)
fill_crimson = PatternFill("solid", fgColor=CRIMSON)
fill_paper = PatternFill("solid", fgColor=PAPER)
fill_input = PatternFill("solid", fgColor=INPUT)
fill_zebra = PatternFill("solid", fgColor=ZEBRA)
fill_gold = PatternFill("solid", fgColor=SOFT_GOLD)
fill_red = PatternFill("solid", fgColor=SOFT_RED)
fill_green = PatternFill("solid", fgColor=SOFT_GREEN)
fill_navy_soft = PatternFill("solid", fgColor=SOFT_NAVY)
fill_white = PatternFill("solid", fgColor=WHITE)

font_white = Font(name="微软雅黑", size=10, color=WHITE, bold=True)
font_title = Font(name="微软雅黑", size=18, color=NAVY, bold=True)
font_h = Font(name="微软雅黑", size=13, color=CRIMSON, bold=True)
font_body = Font(name="微软雅黑", size=10, color=INK)
font_small = Font(name="微软雅黑", size=9, color=MUTED)
font_input = Font(name="微软雅黑", size=11, color=CRIMSON, bold=True)
font_link = Font(name="微软雅黑", size=11, color=CRIMSON, underline="single")

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
left_top = Alignment(horizontal="left", vertical="top", wrap_text=True)


def apply_header(ws, row, start, end, fill=fill_navy):
    for col in range(start, end + 1):
        cell = ws.cell(row, col)
        cell.fill = fill
        cell.font = font_white
        cell.alignment = center
        cell.border = thin


def paint_body(cell, zebra=False, align=None):
    cell.font = font_body
    cell.alignment = align or left_align
    cell.border = thin
    cell.fill = fill_zebra if zebra else fill_white


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def banner(ws, row, cols, text, fill, font, height=32):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    cell = ws.cell(row, 1, text)
    cell.fill = fill
    cell.font = font
    cell.alignment = left_align
    ws.row_dimensions[row].height = height
    return cell


def finish(ws, freeze="A3", landscape=True, fit_width=1):
    ws.freeze_panes = freeze
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = fit_width
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.horizontalCentered = True
    ws.oddHeader.left.text = f"{META['party_a']} × {META['party_b']}"
    ws.oddHeader.right.text = META["version"]
    ws.oddFooter.left.text = "洽谈工作底稿 · 建议数与假设数均可改，但须重新经学院确认"
    ws.oddFooter.right.text = "第 &P 页"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.7
    ws.page_margins.bottom = 0.6
    ws.sheet_properties.tabColor = NAVY


def priority_fill(value):
    if value in ("P0", "最高", "核心", "第一批", "A"):
        return fill_red
    if value in ("P1", "高", "重点", "第二批", "B"):
        return fill_gold
    if value in ("P2", "中高", "中期培育", "储备", "C", "待核实"):
        return fill_navy_soft
    return None


def build(path):
    wb = Workbook()
    wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True)
    wb.properties.title = "复旦终身教育学院与杨浦科联合作工作表"
    wb.properties.creator = META["compiler"]
    wb.properties.subject = META["nature"]
    wb.properties.description = "2026—2027年合作洽谈工作底稿。金色单元格为假设输入。"

    cover(wb.active)
    sheet_framework(wb.create_sheet("01_框架总览"))
    sheet_business(wb.create_sheet("02_六大业务"))
    sheet_visits(wb.create_sheet("03_参访资源库"))
    sheet_priority_visits(wb.create_sheet("04_重点参访"))
    sheet_routes(wb.create_sheet("05_精品线路"))
    sheet_tiers(wb.create_sheet("06_服务层级"))
    sheet_clients(wb.create_sheet("07_客户线索"))
    sheet_internal(wb.create_sheet("08_内部跟进"))
    sheet_products(wb.create_sheet("09_课程产品"))
    sheet_actions(wb.create_sheet("10_客户推进"))
    sheet_interview(wb.create_sheet("11_访谈台账"))
    sheet_segments(wb.create_sheet("12_客群与决策"))
    sheet_channels(wb.create_sheet("13_招生渠道"))
    sheet_modules(wb.create_sheet("14_课程模块"))
    sheet_pricing(wb.create_sheet("15_定价情景"))
    sheet_market(wb.create_sheet("16_竞争与风险"))
    sheet_projects(wb.create_sheet("17_2027项目"))
    sheet_quarters(wb.create_sheet("18_季度与目标"))
    sheet_roles(wb.create_sheet("19_分工与节点"))
    sheet_mou(wb.create_sheet("20_备忘录与评价"))

    wb.save(path)


def cover(ws):
    ws.title = "00_封面"
    ws.sheet_properties.tabColor = CRIMSON
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    set_widths(ws, [28, 78, 28])
    banner(ws, 1, 3, "  2026—2027年合作工作表", fill_navy, Font(name="微软雅黑", size=22, color=WHITE, bold=True), height=42)
    banner(
        ws, 2, 3,
        f"  {META['party_a']}  ×  {META['party_b']}    {META['version']}    {META['compiled']}",
        fill_crimson,
        Font(name="微软雅黑", size=12, color=WHITE),
        height=26,
    )
    notes = [
        (4, "文件性质", META["nature"]),
        (5, "怎么用", "外发学院使用 03、04、05、07、09、17 等业务表。08_内部跟进只供联合会自己填写联系人和金额，不要连同外发。"),
        (6, "数字约定", "金色单元格是假设，可以改。学费、成本和年度规模都由公式重算。改完后的数字仍然只是测算，不是学院核准收费。"),
        (7, "状态约定", "参访 A/B/C 与客户 P0/P1/P2 表示推进顺序，不表示对方已经同意接待或已经有培训预算。"),
        (8, "合规约定", "正式结算采用真实服务、真实合同、依法开票和校方审批。内部管理费比例、票据路径和个人账户安排不进入本表。"),
    ]
    for row, k, v in notes:
        ws.cell(row, 1, k).font = Font(name="微软雅黑", size=11, color=WHITE, bold=True)
        ws.cell(row, 1).fill = fill_navy
        ws.cell(row, 1).alignment = center
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        ws.cell(row, 2, v).font = font_body
        ws.cell(row, 2).alignment = left_align
        ws.cell(row, 2).fill = fill_paper
        ws.row_dimensions[row].height = 32

    ws.cell(10, 1, "工作表").font = font_white
    ws.cell(10, 1).fill = fill_navy
    ws.merge_cells("B10:C10")
    ws.cell(10, 2, "内容").font = font_white
    ws.cell(10, 2).fill = fill_navy
    ws.cell(10, 2).alignment = left_align
    links = [
        ("01_框架总览", "定位、目标、原则和双方职责"),
        ("02_六大业务", "六项业务的优先级和产品形态"),
        ("03_参访资源库", "37条候选资源，可按产业和状态筛选"),
        ("04_重点参访", "建议先核实的12个项目"),
        ("05_精品线路", "五条现场教学组合"),
        ("06_服务层级", "基础、专题、深度三档服务"),
        ("07_客户线索", "22条外发安全版线索，不含电话和预算"),
        ("08_内部跟进", "内部台账。外发前不要附上"),
        ("09_课程产品", "六门待设计课程"),
        ("10_客户推进", "第一轮对接动作"),
        ("11_访谈台账", "空白访谈表，带进度下拉"),
        ("12_客群与决策", "研修班客群和购买决策问题"),
        ("13_招生渠道", "25人配置目标，不是转化率预测"),
        ("14_课程模块", "8天模块草案"),
        ("15_定价情景", "可改单价和成本，自动重算平衡点"),
        ("16_竞争与风险", "公开课程参照和七项风险"),
        ("17_2027项目", "四项核心、三项重点、三项储备"),
        ("18_季度与目标", "年度目标、季度安排和近期路线"),
        ("19_分工与节点", "职责分工和时间节点"),
        ("20_备忘录与评价", "备忘录十条、评价指标和资料依据"),
    ]
    for i, (name, desc) in enumerate(links):
        row = 11 + i
        cell = ws.cell(row, 1, name)
        cell.font = font_link
        cell.hyperlink = f"#{name}!A1"
        cell.alignment = left_align
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        ws.cell(row, 2, desc).font = font_body
        ws.cell(row, 2).alignment = left_align
        if i % 2:
            cell.fill = fill_paper
            ws.cell(row, 2).fill = fill_paper
        ws.row_dimensions[row].height = 20
    ws.row_dimensions[32].height = 36
    ws.merge_cells("A32:C32")
    ws["A32"] = "资料依据：" + " ".join(f"{i}. {s}" for i, s in enumerate(SOURCES, 1))
    ws["A32"].font = font_small
    ws["A32"].alignment = left_top
    ws.oddFooter.right.text = "第 &P 页"
    ws.print_title_rows = "1:2"
    ws.page_setup.horizontalCentered = True
    ws.sheet_view.zoomScale = 120


def sheet_framework(ws):
    set_widths(ws, [22, 78])
    banner(ws, 1, 2, "  合作框架总览", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    rows = [
        ("一个定位", "高校教育资源 + 产业组织平台 + 市场化运营。学院负责教学与标准，联合会负责客户、渠道和产业资源。"),
        ("近期目标", "2026年底：对接机制、合作备忘录、企业参访资源库及首批课程方案。"),
        ("中期目标", "2027年：企业委托培训、人工智能创新人才研修班、产业参访正式落地。"),
        ("长期目标", "2027年以后：形成持续招生和产业特色，拓展长三角及全国企业、园区和政府委托项目。"),
        ("合作原则", "优势互补、专业分工、项目化运作、依法合规、市场导向、成果共享。具体合作以学校审批、正式协议和项目合同为准。"),
        ("学院职责", "立项审批、课程与质量、师资、场地与学员服务、经批准的考核证书、政府与国央企项目交付、教育品牌和教学标准。"),
        ("联合会职责", "客户拓展、需求收集、社群与论坛、参访资源、产业专家对接、宣传推广和经批准的配套服务。"),
        ("基本分工", "学院主导教学，联合会侧重市场与产业资源。运营事项按项目协议确定。"),
        ("候选人员", "已有课程产品和运营经验的人员可作为课程开发或项目运营的候选合作人员，不当然取得校方岗位或项目授权。"),
        ("优先模式", "模式A：企业或政府直接委托学院。模式B须具备真实签约和采购关系。模式C按实际服务与风险定价。"),
        ("试点原则", "首期一年、非排他。备忘录只载明意向。费用、品牌、招生和交付另签专项合同。"),
        ("框架主线", "以委托培训形成业务基础，以产业参访形成差异化，以科创人才短训形成增长空间。"),
    ]
    apply_header(ws, 3, 1, 2)
    ws.cell(3, 1, "项目")
    ws.cell(3, 2, "内容")
    for i, (k, v) in enumerate(rows):
        paint_body(ws.cell(4 + i, 1, k), i % 2 == 1, center)
        ws.cell(4 + i, 1).font = Font(name="微软雅黑", size=10, color=NAVY, bold=True)
        paint_body(ws.cell(4 + i, 2, v), i % 2 == 1)
        ws.row_dimensions[4 + i].height = 32
    finish(ws, "A4")
    ws.auto_filter.ref = "A3:B15"


def sheet_business(ws):
    set_widths(ws, [10, 16, 42, 24, 36])
    banner(ws, 1, 5, "  六大合作业务", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    headers = ["序号", "业务", "对象与主题", "优先级", "产品形态"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    apply_header(ws, 3, 1, 5)
    for r, row in enumerate(BUSINESSES):
        for c, val in enumerate(row, 1):
            paint_body(ws.cell(4 + r, c, val), r % 2 == 1, center if c in (1, 4) else left_align)
        pf = priority_fill(row[3].split("，")[0])
        if pf:
            ws.cell(4 + r, 4).fill = pf
        ws.row_dimensions[4 + r].height = 28
    finish(ws, "A4")
    ws.auto_filter.ref = "A3:E9"


def sheet_visits(ws):
    set_widths(ws, [8, 24, 40, 52, 14, 42])
    banner(ws, 1, 6, "  科技企业标杆参访资源库（候选目录，不是接待承诺）", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    headers = ["序号", "产业方向", "拟参访企业或平台", "建议参访主题", "资源状态", "状态含义"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    apply_header(ws, 3, 1, 6)
    for r, row in enumerate(VISITS):
        status = row[4]
        values = list(row) + [VISIT_STATUS[status]]
        for c, val in enumerate(values, 1):
            paint_body(ws.cell(4 + r, c, val), r % 2 == 1, center if c in (1, 5) else left_align)
        ws.cell(4 + r, 5).fill = priority_fill(status) or fill_white
        ws.row_dimensions[4 + r].height = 30
    note_row = 4 + len(VISITS)
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=6)
    ws.cell(note_row, 1, "厂区、生产线、研发实验室和商业秘密区域，只在企业批准范围内开放。医疗研发环境不列为默认参观内容。")
    ws.cell(note_row, 1).font = font_small
    ws.cell(note_row, 1).alignment = left_align
    ws.auto_filter.ref = f"A3:F{3 + len(VISITS)}"
    ws.print_title_rows = "1:3"
    finish(ws, "A4")


def sheet_priority_visits(ws):
    set_widths(ws, [14, 42, 36, 28])
    banner(ws, 1, 4, "  首批12个重点参访项目（优先核实，尚未获得接待确认）", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    headers = ["批次", "参访资源", "首选教学方向", "核实结果"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    apply_header(ws, 3, 1, 4)
    for r, row in enumerate(PRIORITY_VISITS):
        for c, val in enumerate(row, 1):
            paint_body(ws.cell(4 + r, c, val), r % 2 == 1, center if c == 1 else left_align)
        ws.cell(4 + r, 1).fill = priority_fill(row[0]) or fill_white
        cell = ws.cell(4 + r, 4, "待核实")
        paint_body(cell, False, center)
        cell.fill = fill_input
        cell.font = font_input
        cell.comment = Comment("核实接待窗口、人数、时长、费用和开放范围后，把“待核实”改成结论。", "联合会")
        ws.row_dimensions[4 + r].height = 24
    dv = DataValidation(type="list", formula1='"待核实,可接待,有条件接待,暂不开放,未联系上"', allow_blank=False)
    dv.error = "请从列表中选择"
    dv.errorTitle = "核实结果"
    ws.add_data_validation(dv)
    dv.add(f"D4:D{3 + len(PRIORITY_VISITS)}")
    finish(ws, "A4")
    ws.auto_filter.ref = f"A3:D{3 + len(PRIORITY_VISITS)}"


def sheet_routes(ws):
    set_widths(ws, [12, 28, 62, 40])
    banner(ws, 1, 4, "  五条精品现场教学线路（组合建议，顺序以实际预约为准）", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    headers = ["线路", "主题", "组合", "适合对象"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    apply_header(ws, 3, 1, 4)
    for r, row in enumerate(ROUTES):
        for c, val in enumerate(row, 1):
            paint_body(ws.cell(4 + r, c, val), r % 2 == 1, center if c == 1 else left_align)
        ws.row_dimensions[4 + r].height = 32
    finish(ws, "A4")


def sheet_tiers(ws):
    set_widths(ws, [20, 22, 70])
    banner(ws, 1, 3, "  参访服务层级（实际内容由接待机构确认）", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    headers = ["服务类型", "建议时长", "建议配置"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    apply_header(ws, 3, 1, 3)
    for r, row in enumerate(VISIT_TIERS):
        for c, val in enumerate(row, 1):
            paint_body(ws.cell(4 + r, c, val), r % 2 == 1, center if c < 3 else left_align)
        ws.row_dimensions[4 + r].height = 28
    ws.merge_cells("A8:C8")
    ws["A8"] = "报价分开列示：企业实际接待费用、交通会务支出、资源协调服务费用。已包含在培训合同中的参访，不重复计入独立收入。"
    ws["A8"].font = font_small
    ws["A8"].alignment = left_align
    finish(ws, "A4", landscape=False)


def sheet_clients(ws):
    set_widths(ws, [8, 32, 18, 46, 46, 12])
    banner(ws, 1, 6, "  首批客户及需求线索（外发版：不填未确认的姓名、电话和预算）", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    headers = ["序号", "潜在客户", "类型", "现有关系基础", "拟匹配需求（待确认）", "优先级"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    apply_header(ws, 3, 1, 6)
    for r, row in enumerate(CLIENTS):
        for c, val in enumerate(row, 1):
            paint_body(ws.cell(4 + r, c, val), r % 2 == 1, center if c in (1, 6) else left_align)
        ws.cell(4 + r, 6).fill = priority_fill(row[5]) or fill_white
        ws.cell(4 + r, 6).font = Font(name="微软雅黑", size=10, color=CRIMSON if row[5] == "P0" else INK, bold=True)
        ws.row_dimensions[4 + r].height = 32
    finish(ws, "A4")
    ws.auto_filter.ref = f"A3:F{3 + len(CLIENTS)}"
    ws.print_title_rows = "1:3"


def sheet_internal(ws):
    set_widths(ws, [8, 32, 10, 16, 18, 22, 14, 16, 28, 16, 12, 14, 22, 12, 20])
    banner(
        ws, 1, 15,
        "  内部跟进台账 · 含待填写的联系人和金额 · 提交学院时不要附上本表",
        fill_crimson,
        Font(name="微软雅黑", size=14, color=WHITE, bold=True),
        height=28,
    )
    headers = ["序号", "客户", "优先级", "联系人", "部门岗位", "电话或邮箱", "介绍人", "首次沟通", "客户反馈", "预计金额（元）", "成交概率", "当前进度", "下一步", "责任人", "备注"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    apply_header(ws, 3, 1, 15, fill_crimson)
    start = 4
    for i, row in enumerate(CLIENTS):
        r = start + i
        ws.cell(r, 1, f"='07_客户线索'!A{4 + i}")
        ws.cell(r, 2, f"='07_客户线索'!B{4 + i}")
        ws.cell(r, 3, f"='07_客户线索'!F{4 + i}")
        for c in range(1, 16):
            paint_body(ws.cell(r, c), i % 2 == 1, center if c in (1, 3, 11, 12) else left_align)
            if c >= 4:
                ws.cell(r, c).fill = fill_input
                ws.cell(r, c).protection = Protection(locked=False)
        ws.row_dimensions[r].height = 22
    last = start + len(CLIENTS) - 1
    dv_p = DataValidation(type="list", formula1='"待联系,已联系,需求访谈,方案提交,洽谈中,已签约,暂缓"', allow_blank=True)
    dv_prob = DataValidation(type="list", formula1='"空白,低,中,高,已签约"', allow_blank=True)
    ws.add_data_validation(dv_p)
    ws.add_data_validation(dv_prob)
    dv_p.add(f"L{start}:L{last}")
    dv_prob.add(f"K{start}:K{last}")
    ws.auto_filter.ref = f"A3:O{last}"
    finish(ws, "A4")
    ws.sheet_properties.tabColor = CRIMSON
    ws.auto_filter.ref = f"A3:O{last}"


def sheet_products(ws):
    set_widths(ws, [38, 32, 16, 28])
    banner(ws, 1, 4, "  首批拟开发课程（周期和形式为初步参数，由学院确定）", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    headers = ["产品名称", "目标客户", "建议周期", "建议组织形式"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    apply_header(ws, 3, 1, 4)
    for r, row in enumerate(PRODUCTS):
        for c, val in enumerate(row, 1):
            paint_body(ws.cell(4 + r, c, val), r % 2 == 1, center if c == 3 else left_align)
        ws.row_dimensions[4 + r].height = 26
    finish(ws, "A4")


def sheet_actions(ws):
    set_widths(ws, [24, 42, 32, 28])
    banner(ws, 1, 4, "  首批重点客户推进（2026年11—12月完成第一轮）", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    headers = ["客户或渠道", "首轮对接事项", "建议责任分工", "目标成果"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    apply_header(ws, 3, 1, 4)
    for r, row in enumerate(CLIENT_ACTIONS):
        for c, val in enumerate(row, 1):
            paint_body(ws.cell(4 + r, c, val), r % 2 == 1)
        ws.row_dimensions[4 + r].height = 30
    ws.merge_cells("A12:D12")
    ws["A12"] = "先把国泰海通、国药控股、中远相关企业的培训负责人引荐给学院对接人。三家中有一家形成真实订单，备忘录就有了落脚点。"
    ws["A12"].font = Font(name="微软雅黑", size=10, color=CRIMSON, bold=True)
    ws["A12"].alignment = left_align
    finish(ws, "A4")


def sheet_interview(ws):
    headers = ["序号"] + [f[0] for f in INTERVIEW_FIELDS]
    set_widths(ws, [8] + [18] * (len(headers) - 1))
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["G"].width = 28
    ws.column_dimensions["N"].width = 28
    banner(ws, 1, len(headers), "  客户需求访谈台账 · 未获得的信息填“待确认”，不要把行业通用需求写成客户已经确认的预算", fill_navy, Font(name="微软雅黑", size=14, color=WHITE, bold=True))
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    apply_header(ws, 3, 1, len(headers))
    for i in range(12):
        r = 4 + i
        paint_body(ws.cell(r, 1, i + 1), i % 2 == 1, center)
        for c in range(2, len(headers) + 1):
            paint_body(ws.cell(r, c, None), i % 2 == 1)
            ws.cell(r, c).fill = fill_input
        ws.row_dimensions[r].height = 22
    dv = DataValidation(type="list", formula1='"待联系,已联系,需求访谈,方案提交,洽谈中,已签约,暂缓"', allow_blank=True)
    dv2 = DataValidation(type="list", formula1='"直接委托,公开采购,招标,尚未明确"', allow_blank=True)
    ws.add_data_validation(dv)
    ws.add_data_validation(dv2)
    dv.add("M4:M15")
    dv2.add("K4:K15")
    # 字段说明
    base = 17
    ws.cell(base, 1, "字段").font = font_white
    ws.cell(base, 1).fill = fill_navy
    ws.merge_cells(start_row=base, start_column=2, end_row=base, end_column=4)
    ws.cell(base, 2, "填写要求").font = font_white
    ws.cell(base, 2).fill = fill_navy
    for i, (name, rule) in enumerate(INTERVIEW_FIELDS):
        paint_body(ws.cell(base + 1 + i, 1, name), i % 2 == 1, center)
        ws.merge_cells(start_row=base + 1 + i, start_column=2, end_row=base + 1 + i, end_column=4)
        paint_body(ws.cell(base + 1 + i, 2, rule), i % 2 == 1)
    finish(ws, "A4")
    ws.auto_filter.ref = "A3:N15"


def sheet_segments(ws):
    set_widths(ws, [34, 36, 24, 16])
    banner(ws, 1, 4, "  研修班目标客户与购买决策问题", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    headers = ["目标客户", "主要关注问题", "建议课程方向", "招生价值"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    apply_header(ws, 3, 1, 4)
    for r, row in enumerate(SEGMENTS):
        for c, val in enumerate(row, 1):
            paint_body(ws.cell(4 + r, c, val), r % 2 == 1)
        ws.row_dimensions[4 + r].height = 26
    ws.cell(11, 1, "决策因素").font = font_white
    ws.cell(11, 1).fill = fill_navy
    ws.merge_cells("B11:D11")
    ws.cell(11, 2, "访谈时需要确认的问题").font = font_white
    ws.cell(11, 2).fill = fill_navy
    for i, (k, v) in enumerate(DECISION_FACTORS):
        paint_body(ws.cell(12 + i, 1, k), i % 2 == 1, center)
        ws.merge_cells(start_row=12 + i, start_column=2, end_row=12 + i, end_column=4)
        paint_body(ws.cell(12 + i, 2, v), i % 2 == 1)
    finish(ws, "A4")


def sheet_channels(ws):
    set_widths(ws, [36, 14, 12, 36])
    banner(ws, 1, 4, "  首期招生渠道配置目标（不是已验证的转化率）", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    headers = ["渠道", "拟招生人数", "占比", "说明"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    apply_header(ws, 3, 1, 4)
    notes = [
        "一对一访谈和企业团体报名优先",
        "采购周期须单独确认",
        "先办小规模说明会，再看转化",
        "按身份重分类，不按历史总人数测算",
        "须经学院授权后使用",
    ]
    for i, (name, n, _ratio) in enumerate(CHANNELS):
        paint_body(ws.cell(4 + i, 1, name), i % 2 == 1)
        cell = ws.cell(4 + i, 2, n)
        paint_body(cell, i % 2 == 1, center)
        cell.fill = fill_input
        cell.font = font_input
        paint_body(ws.cell(4 + i, 3, f"=B{4 + i}/$B$9"), i % 2 == 1, center)
        ws.cell(4 + i, 3).number_format = "0%"
        paint_body(ws.cell(4 + i, 4, notes[i]), i % 2 == 1)
        ws.row_dimensions[4 + i].height = 24
    paint_body(ws.cell(9, 1, "合计"), False, left_align)
    ws.cell(9, 1).font = Font(name="微软雅黑", size=10, color=WHITE, bold=True)
    ws.cell(9, 1).fill = fill_navy
    paint_body(ws.cell(9, 2, "=SUM(B4:B8)"), False, center)
    ws.cell(9, 2).font = Font(name="微软雅黑", size=12, color=WHITE, bold=True)
    ws.cell(9, 2).fill = fill_navy
    ws.cell(9, 2).number_format = "0"
    paint_body(ws.cell(9, 3, "=SUM(C4:C8)"), False, center)
    ws.cell(9, 3).number_format = "0%"
    ws.cell(9, 3).font = Font(name="微软雅黑", size=10, color=WHITE, bold=True)
    ws.cell(9, 3).fill = fill_navy
    paint_body(ws.cell(9, 4, "配置目标合计。改金色人数后，占比自动重算。"), False)
    ws.cell(9, 4).fill = fill_navy
    ws.cell(9, 4).font = Font(name="微软雅黑", size=10, color=WHITE)

    ws.cell(11, 1, "环节").font = font_white
    ws.cell(11, 1).fill = fill_navy
    ws.cell(11, 2, "试点目标人数").font = font_white
    ws.cell(11, 2).fill = fill_navy
    for i, (name, n) in enumerate(FUNNEL):
        paint_body(ws.cell(12 + i, 1, name), i % 2 == 1)
        cell = ws.cell(12 + i, 2, n)
        paint_body(cell, False, center)
        cell.fill = fill_input
        cell.font = font_input
    ws.merge_cells("A18:D18")
    ws["A18"] = "漏斗是目标情景，不代表已经掌握相应数量的客户。正式立项前先完成不少于30位企业管理者访谈。"
    ws["A18"].font = font_small
    ws["A18"].alignment = left_align
    finish(ws, "A4")


def sheet_modules(ws):
    set_widths(ws, [28, 48, 14])
    banner(ws, 1, 3, f"  {AI_PLAN['name']} · 模块草案（学时由学院核定）", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    headers = ["模块", "主要内容", "建议天数"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    apply_header(ws, 3, 1, 3)
    for r, row in enumerate(MODULES):
        for c, val in enumerate(row, 1):
            paint_body(ws.cell(4 + r, c, val), r % 2 == 1, center if c == 3 else left_align)
            if c == 3:
                ws.cell(4 + r, c).fill = fill_input
                ws.cell(4 + r, c).font = font_input
        ws.row_dimensions[4 + r].height = 24
    paint_body(ws.cell(11, 1, "合计"), False)
    ws.cell(11, 1).font = Font(name="微软雅黑", size=10, color=WHITE, bold=True)
    ws.cell(11, 1).fill = fill_navy
    paint_body(ws.cell(11, 2, "课堂教学与现场教学"), False)
    ws.cell(11, 2).fill = fill_navy
    ws.cell(11, 2).font = Font(name="微软雅黑", size=10, color=WHITE)
    paint_body(ws.cell(11, 3, "=SUM(C4:C10)"), False, center)
    ws.cell(11, 3).fill = fill_navy
    ws.cell(11, 3).font = Font(name="微软雅黑", size=12, color=WHITE, bold=True)
    ws.cell(11, 3).number_format = "0"
    finish(ws, "A4", landscape=False)


def sheet_pricing(ws):
    set_widths(ws, [28, 20, 22, 24, 36])
    banner(ws, 1, 5, "  定价与可行性测算 · 金色为假设输入，结果不是可分配利润", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    ws["A3"] = "假设单价（元/人）"
    ws["B3"] = AI_PLAN["price"]
    ws["C3"] = "洽谈区间为数千元至2万元。16,800用于测试，不是核准学费。"
    ws["A4"] = "假设班均直接成本（元）"
    ws["B4"] = AI_PLAN["cost_assumption"]
    ws["C4"] = "含教学组织、讲师、场地、参访、市场与运营的合并假设，不分项。"
    ws["A5"] = "假设可支配比例"
    ws["B5"] = AI_PLAN["disposable_ratio"]
    ws["C5"] = "仅为覆盖直接支出的资金比例假设，不是学院管理费标准。"
    for row in (3, 4, 5):
        ws.cell(row, 1).font = Font(name="微软雅黑", size=10, color=NAVY, bold=True)
        ws.cell(row, 1).alignment = left_align
        ws.cell(row, 1).fill = fill_paper
        ws.cell(row, 2).font = font_input
        ws.cell(row, 2).fill = fill_input
        ws.cell(row, 2).alignment = center
        ws.cell(row, 2).border = thin
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=5)
        ws.cell(row, 3).font = font_small
        ws.cell(row, 3).alignment = left_align
        ws.row_dimensions[row].height = 24
    ws["B3"].number_format = '#,##0'
    ws["B4"].number_format = '#,##0'
    ws["B5"].number_format = "0%"
    ws["B3"].comment = Comment("修改这里会重算下方全部情景。", "测算")

    headers = ["招生人数", "理论学费", "假设可支配资金", "扣除假设成本后的余额", "说明"]
    for i, h in enumerate(headers, 1):
        ws.cell(7, i, h)
    apply_header(ws, 7, 1, 5)
    reads = [
        "接近假设平衡点，余额很薄",
        "基准情景",
        "积极情景，仍须先验证付费意愿",
    ]
    for i, n in enumerate(SCENARIOS):
        r = 8 + i
        cell_n = ws.cell(r, 1, n)
        paint_body(cell_n, False, center)
        cell_n.fill = fill_input
        cell_n.font = font_input
        paint_body(ws.cell(r, 2, f"=A{r}*$B$3"), False, center)
        paint_body(ws.cell(r, 3, f"=B{r}*$B$5"), False, center)
        paint_body(ws.cell(r, 4, f"=C{r}-$B$4"), False, center)
        paint_body(ws.cell(r, 5, reads[i]), False)
        for c in range(2, 5):
            ws.cell(r, c).number_format = '#,##0'
        ws.row_dimensions[r].height = 24
    red = FormulaRule(formula=["D8<0"], fill=fill_red, font=Font(name="微软雅黑", color=CRIMSON, bold=True))
    green = FormulaRule(formula=["D8>=0"], fill=fill_green, font=Font(name="微软雅黑", color="1B7A45", bold=True))
    ws.conditional_formatting.add("D8:D10", red)
    ws.conditional_formatting.add("D8:D10", green)

    ws["A12"] = "假设盈亏平衡人数"
    ws["B12"] = '=IF(B3*B5=0,"—",ROUNDUP(B4/(B3*B5),0))'
    ws["C12"] = "按当前三个假设计算。人数变化引起的成本变化、税费和分摊都没有计入。"
    ws["A12"].font = Font(name="微软雅黑", size=10, color=WHITE, bold=True)
    ws["A12"].fill = fill_crimson
    ws["A12"].alignment = left_align
    ws["B12"].font = Font(name="微软雅黑", size=16, color=CRIMSON, bold=True)
    ws["B12"].fill = fill_red
    ws["B12"].alignment = center
    ws["B12"].border = thin
    ws["B12"].number_format = "0"
    ws.merge_cells("C12:E12")
    ws["C12"].font = font_small
    ws["C12"].alignment = left_align
    ws.row_dimensions[12].height = 28

    ws["A14"] = "产品层级"
    ws["B14"] = "建议价格"
    ws["C14"] = "定位"
    apply_header(ws, 14, 1, 3)
    for i, row in enumerate(PRICE_LADDER):
        for c, val in enumerate(row, 1):
            paint_body(ws.cell(15 + i, c, val), i % 2 == 1, center if c == 2 else left_align)
        ws.merge_cells(start_row=15 + i, start_column=3, end_row=15 + i, end_column=5)

    ws["A19"] = "年度规模示例"
    ws["A19"].font = font_h
    headers2 = ["收费项目", "数量", "单价（元）", "理论收入（元）", "说明"]
    for i, h in enumerate(headers2, 1):
        ws.cell(20, i, h)
    apply_header(ws, 20, 1, 5)
    scale_notes = [
        "示例单价，待客户预算确认",
        "与上方单价分开。若要联动，把 C22 改成 =B3",
        "已打包进培训合同的参访不要再计",
        "示例单价，待采购方式确认",
    ]
    for i, (name, qty, price, _unit) in enumerate(SCALE_EXAMPLE):
        r = 21 + i
        paint_body(ws.cell(r, 1, name), i % 2 == 1)
        for col, val in ((2, qty), (3, price)):
            cell = ws.cell(r, col, val)
            paint_body(cell, False, center)
            cell.fill = fill_input
            cell.font = font_input
            cell.number_format = '#,##0'
        paint_body(ws.cell(r, 4, f"=B{r}*C{r}"), False, center)
        ws.cell(r, 4).number_format = '#,##0'
        paint_body(ws.cell(r, 5, scale_notes[i]), False)
        ws.row_dimensions[r].height = 22
    paint_body(ws.cell(25, 1, "合计"), False)
    ws.cell(25, 1).fill = fill_navy
    ws.cell(25, 1).font = Font(name="微软雅黑", size=10, color=WHITE, bold=True)
    for col in (2, 3):
        ws.cell(25, col).fill = fill_navy
        ws.cell(25, col).border = thin
    paint_body(ws.cell(25, 4, "=SUM(D21:D24)"), False, center)
    ws.cell(25, 4).number_format = '#,##0'
    ws.cell(25, 4).font = Font(name="微软雅黑", size=14, color=WHITE, bold=True)
    ws.cell(25, 4).fill = fill_crimson
    paint_body(ws.cell(25, 5, "理论收入示例。不是营业收入，也不是联合会收益。"), False)
    ws.cell(25, 5).fill = fill_crimson
    ws.cell(25, 5).font = Font(name="微软雅黑", size=10, color=WHITE, bold=True)
    ws.row_dimensions[25].height = 26
    finish(ws, "A7")
    ws.freeze_panes = "A7"


def sheet_market(ws):
    set_widths(ws, [62, 28, 18, 36])
    banner(ws, 1, 4, "  公开课程参照与主要风险", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    headers = ["项目及主体", "学制", "公开参考学费", "使用注意"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    apply_header(ws, 3, 1, 4)
    for r, row in enumerate(COMPETITORS):
        for c, val in enumerate(row, 1):
            paint_body(ws.cell(4 + r, c, val), r % 2 == 1, center if c in (2, 3) else left_align)
        ws.row_dimensions[4 + r].height = 28
    ws.merge_cells("A9:D9")
    ws["A9"] = (
        f"学院既有项目已核对：{EXISTING_AIGC['name']}，{EXISTING_AIGC['term']}，"
        f"项目编号 {EXISTING_AIGC['codes']}，{EXISTING_AIGC['price_note']}，每期{EXISTING_AIGC['seats']}席。"
        f"{EXISTING_AIGC['days']}。证书表述以该项目已批准的规则为准，新项目不得自行沿用。"
    )
    ws["A9"].font = font_small
    ws["A9"].alignment = left_align
    ws.row_dimensions[9].height = 48

    ws.cell(11, 1, "风险").font = font_white
    ws.cell(11, 1).fill = fill_navy
    ws.merge_cells("B11:D11")
    ws.cell(11, 2, "应对").font = font_white
    ws.cell(11, 2).fill = fill_navy
    for i, (k, v) in enumerate(RISKS):
        paint_body(ws.cell(12 + i, 1, k), i % 2 == 1, center)
        ws.merge_cells(start_row=12 + i, start_column=2, end_row=12 + i, end_column=4)
        paint_body(ws.cell(12 + i, 2, v), i % 2 == 1)
        ws.row_dimensions[12 + i].height = 30
    finish(ws, "A4")


def sheet_projects(ws):
    set_widths(ws, [8, 42, 40, 14, 26, 12])
    banner(ws, 1, 6, "  2027年拟合作项目（建议目标，不是已有订单）", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    headers = ["序号", "项目", "主要服务对象", "时间", "年度目标", "层级"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    apply_header(ws, 3, 1, 6)
    for r, row in enumerate(PROJECTS_2027):
        for c, val in enumerate(row, 1):
            paint_body(ws.cell(4 + r, c, val), r % 2 == 1, center if c in (1, 4, 6) else left_align)
        ws.cell(4 + r, 6).fill = priority_fill(row[5]) or fill_white
        ws.row_dimensions[4 + r].height = 28
    finish(ws, "A4")
    ws.auto_filter.ref = "A3:F13"


def sheet_quarters(ws):
    set_widths(ws, [18, 62, 48])
    banner(ws, 1, 3, "  年度目标、季度安排与 2026 年底前路线", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    ws.cell(3, 1, "指标").font = font_white
    ws.cell(3, 1).fill = fill_navy
    ws.merge_cells("B3:C3")
    ws.cell(3, 2, "2027年建议目标").font = font_white
    ws.cell(3, 2).fill = fill_navy
    for i, (k, v) in enumerate(ANNUAL_GOALS):
        paint_body(ws.cell(4 + i, 1, k), i % 2 == 1, center)
        ws.merge_cells(start_row=4 + i, start_column=2, end_row=4 + i, end_column=3)
        paint_body(ws.cell(4 + i, 2, v), i % 2 == 1)
        ws.row_dimensions[4 + i].height = 22
    base = 13
    headers = ["季度", "核心工作", "重点成果"]
    for i, h in enumerate(headers, 1):
        ws.cell(base, i, h)
    apply_header(ws, base, 1, 3)
    for i, row in enumerate(QUARTERS):
        for c, val in enumerate(row, 1):
            paint_body(ws.cell(base + 1 + i, c, val), i % 2 == 1, center if c == 1 else left_align)
        ws.row_dimensions[base + 1 + i].height = 32
    base2 = 19
    headers = ["阶段", "重点任务", "预期成果"]
    for i, h in enumerate(headers, 1):
        ws.cell(base2, i, h)
    apply_header(ws, base2, 1, 3)
    for i, row in enumerate(ROADMAP):
        for c, val in enumerate(row, 1):
            paint_body(ws.cell(base2 + 1 + i, c, val), i % 2 == 1, center if c == 1 else left_align)
        ws.row_dimensions[base2 + 1 + i].height = 30
    finish(ws, "A4")


def sheet_roles(ws):
    set_widths(ws, [28, 36, 36])
    banner(ws, 1, 3, "  职责分工与时间节点", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    headers = ["工作事项", "复旦大学终身教育学院（拟）", "杨浦区科技企业联合会"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    apply_header(ws, 3, 1, 3)
    for r, row in enumerate(ROLES):
        for c, val in enumerate(row, 1):
            paint_body(ws.cell(4 + r, c, val), r % 2 == 1, center if c > 1 else left_align)
        ws.row_dimensions[4 + r].height = 24
    base = 14
    for i, h in enumerate(["工作事项", "建议牵头方", "拟完成节点"], 1):
        ws.cell(base, i, h)
    apply_header(ws, base, 1, 3)
    for i, row in enumerate(MILESTONES):
        for c, val in enumerate(row, 1):
            paint_body(ws.cell(base + 1 + i, c, val), i % 2 == 1, center if c > 1 else left_align)
        ws.row_dimensions[base + 1 + i].height = 22
    ws.merge_cells("A24:C24")
    ws["A24"] = "节点以正式合作和学校立项、审批结果为前提。联合会负责市场开发和产业资源，不自动取得培训业务代理权。"
    ws["A24"].font = font_small
    ws["A24"].alignment = left_align
    finish(ws, "A4")


def sheet_mou(ws):
    set_widths(ws, [10, 70, 40])
    banner(ws, 1, 3, "  合作备忘录建议条款、评价指标与近期材料", fill_navy, Font(name="微软雅黑", size=16, color=WHITE, bold=True))
    headers = ["序号", "备忘录建议明确的事项", "备注"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    apply_header(ws, 3, 1, 3)
    for i, item in enumerate(MOU_ITEMS):
        paint_body(ws.cell(4 + i, 1, i + 1), i % 2 == 1, center)
        paint_body(ws.cell(4 + i, 2, item), i % 2 == 1)
        note = "写入意向，专项合同再落实费用与品牌" if i in (4, 5, 6, 9) else "意向层面"
        paint_body(ws.cell(4 + i, 3, note), i % 2 == 1)
        ws.row_dimensions[4 + i].height = 22
    base = 15
    ws.cell(base, 1, "评价").font = font_white
    ws.cell(base, 1).fill = fill_navy
    ws.merge_cells(start_row=base, start_column=2, end_row=base, end_column=3)
    ws.cell(base, 2, "试点期看这些，不只看签约金额").font = font_white
    ws.cell(base, 2).fill = fill_navy
    for i, (k, v) in enumerate(KPI):
        paint_body(ws.cell(base + 1 + i, 1, k), i % 2 == 1, center)
        ws.merge_cells(start_row=base + 1 + i, start_column=2, end_row=base + 1 + i, end_column=3)
        paint_body(ws.cell(base + 1 + i, 2, v), i % 2 == 1)
    base2 = 23
    ws.cell(base2, 1, "联合会近期提交").font = font_white
    ws.cell(base2, 1).fill = fill_crimson
    ws.merge_cells(start_row=base2, start_column=2, end_row=base2, end_column=3)
    ws.cell(base2, 2, "本套 Word、PPT、工作表即这四份材料的合订呈现").font = font_white
    ws.cell(base2, 2).fill = fill_crimson
    for i, name in enumerate(NEAR_TERM_B):
        paint_body(ws.cell(base2 + 1 + i, 1, i + 1), False, center)
        ws.merge_cells(start_row=base2 + 1 + i, start_column=2, end_row=base2 + 1 + i, end_column=3)
        paint_body(ws.cell(base2 + 1 + i, 2, name), False)
    base3 = 29
    ws.cell(base3, 1, "请学院协调").font = font_white
    ws.cell(base3, 1).fill = fill_navy
    ws.merge_cells(start_row=base3, start_column=2, end_row=base3, end_column=3)
    ws.cell(base3, 2, "资料交换后再确定备忘录文本和第一批试点").font = font_white
    ws.cell(base3, 2).fill = fill_navy
    for i, name in enumerate(NEAR_TERM_A):
        paint_body(ws.cell(base3 + 1 + i, 1, i + 1), i % 2 == 1, center)
        ws.merge_cells(start_row=base3 + 1 + i, start_column=2, end_row=base3 + 1 + i, end_column=3)
        paint_body(ws.cell(base3 + 1 + i, 2, name), i % 2 == 1)
    finish(ws, "A4")


if __name__ == "__main__":
    from pathlib import Path
    out = Path(__file__).resolve().parents[1] / "output"
    out.mkdir(exist_ok=True)
    target = out / "复旦终身教育学院与杨浦科联合作工作表.xlsx"
    build(target)
    print(target)
