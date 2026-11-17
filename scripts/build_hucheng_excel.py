#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成《沪蓉科技产业协同发展合作建议书》讨论稿配套工作底表。

与 PPT 使用同一口径，可随讨论稿发送。
标为「待填写」的单元格留给双方确认，不预填未经核对的数字。
"""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins

NAVY = "10243A"
TEAL = "1B6B7A"
GOLD = "8C6A2F"
CREAM = "F6F3EC"
WHITE = "FFFFFF"
ZEBRA = "F8F5EE"
INK = "1C2833"
MUTED = "5C6B76"
INPUT = "FFF6D8"
SOFT_TEAL = "E6F3F1"
SOFT_GOLD = "F8F1DE"
SOFT_RED = "F8EEEE"
LINE = "E4DACC"

FONT = "微软雅黑"
thin = Side(style="thin", color="D9CBB8")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
TITLE_FONT = Font(name=FONT, size=16, bold=True, color=NAVY)
NOTE_FONT = Font(name=FONT, size=10, italic=True, color=MUTED)
HEAD_FONT = Font(name=FONT, size=11, bold=True, color=WHITE)
BODY = Font(name=FONT, size=10.5, color=INK)
BOLD = Font(name=FONT, size=10.5, bold=True, color=NAVY)
WRAP = Alignment(wrap_text=True, vertical="center", horizontal="left")
CENTER = Alignment(wrap_text=True, vertical="center", horizontal="center")
LEFT_CENTER = Alignment(wrap_text=True, vertical="center", horizontal="left")

NAVY_FILL = PatternFill("solid", fgColor=NAVY)
TEAL_FILL = PatternFill("solid", fgColor=TEAL)
GOLD_FILL = PatternFill("solid", fgColor=GOLD)
CREAM_FILL = PatternFill("solid", fgColor=CREAM)
ZEBRA_FILL = PatternFill("solid", fgColor=ZEBRA)
WHITE_FILL = PatternFill("solid", fgColor=WHITE)
INPUT_FILL = PatternFill("solid", fgColor=INPUT)
SOFT_TEAL_FILL = PatternFill("solid", fgColor=SOFT_TEAL)
SOFT_GOLD_FILL = PatternFill("solid", fgColor=SOFT_GOLD)
SOFT_RED_FILL = PatternFill("solid", fgColor=SOFT_RED)


def apply_page(ws, title):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.7, bottom=0.6, header=0.3, footer=0.3)
    ws.oddHeader.left.text = "上海市杨浦区科技企业联合会"
    ws.oddHeader.right.text = "战略合作建议讨论稿"
    ws.oddFooter.left.text = title
    ws.oddFooter.right.text = "不构成法律约束力协议  ·  第 &P 页"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.sheet_properties.tabColor = NAVY
    ws.auto_filter.ref = None
    ws.oddHeader.left.font = FONT
    ws.print_title_rows = "1:3"


def write_title(ws, columns, title, note):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=columns)
    cell = ws.cell(1, 1, title)
    cell.font = TITLE_FONT
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    cell.fill = CREAM_FILL
    for col in range(1, columns + 1):
        ws.cell(1, col).fill = CREAM_FILL
    ws.row_dimensions[1].height = 32

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=columns)
    note_cell = ws.cell(2, 1, note)
    note_cell.font = NOTE_FONT
    note_cell.alignment = Alignment(vertical="center", wrap_text=True)
    for col in range(1, columns + 1):
        ws.cell(2, col).fill = CREAM_FILL
    ws.row_dimensions[2].height = 28


def write_header(ws, headers):
    for col, header in enumerate(headers, 1):
        cell = ws.cell(3, col, header)
        cell.font = HEAD_FONT
        cell.fill = NAVY_FILL
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[3].height = 24
    ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}3"
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}3"


def paint_row(ws, row, columns, fill):
    for col in range(1, columns + 1):
        cell = ws.cell(row, col)
        cell.fill = fill
        cell.border = BORDER
        cell.font = BOLD if col == 1 else BODY
        cell.alignment = CENTER if col == 1 else WRAP


def set_widths(ws, widths):
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(index)].width = width


def wrapped_lines(value, width):
    text = "" if value is None else str(value)
    # 按偏宽的中文字宽估算，避免微软雅黑与替代字体下换行被裁切。
    capacity = max(4, int((width - 2) / 2.6))
    lines = 0
    for part in text.split("\n"):
        lines += max(1, (len(part) + capacity - 1) // capacity)
    return max(1, lines)


def fitted_height(row, widths):
    lines = 1
    for value, width in zip(row, widths):
        lines = max(lines, wrapped_lines(value, width))
    return min(128, max(36, 16 + lines * 20))


def add_rows(ws, rows, fills=None, heights=None, input_cols=None, widths=None):
    columns = ws.max_column
    for offset, row in enumerate(rows):
        excel_row = 4 + offset
        fill = (fills[offset] if fills else None) or (ZEBRA_FILL if offset % 2 else WHITE_FILL)
        for col, value in enumerate(row, 1):
            cell = ws.cell(excel_row, col, value)
            cell.font = BOLD if col == 1 else BODY
            cell.alignment = CENTER if col == 1 and not isinstance(value, str) else WRAP
            if col == 1 and isinstance(value, (int, float)):
                cell.alignment = CENTER
            elif col == 1 and isinstance(value, str) and len(value) <= 8:
                cell.alignment = CENTER
            else:
                cell.alignment = LEFT_CENTER
            cell.border = BORDER
            cell.fill = INPUT_FILL if input_cols and col in input_cols else fill
        if heights:
            height = heights[offset]
        elif widths:
            height = fitted_height(row, widths)
        else:
            height = 42
        ws.row_dimensions[excel_row].height = height
    last = 3 + len(rows)
    ws.auto_filter.ref = f"A3:{get_column_letter(columns)}{last}"
    return last


def status_validation(ws, cell_range):
    dv = DataValidation(type="list", formula1='"待启动,进行中,已完成,暂缓"', allow_blank=True)
    dv.error = "请从列表中选择状态"
    dv.errorTitle = "状态"
    dv.prompt = "选择当前状态"
    dv.promptTitle = "状态"
    dv.add(cell_range)
    ws.add_data_validation(dv)


def confirm_validation(ws, cell_range):
    dv = DataValidation(type="list", formula1='"待确认,已确认,不适用"', allow_blank=True)
    dv.error = "请从列表中选择"
    dv.errorTitle = "确认状态"
    dv.prompt = "确认后选择状态，并把结论写在右侧"
    dv.promptTitle = "确认"
    dv.add(cell_range)
    ws.add_data_validation(dv)


def build():
    wb = Workbook()
    wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True)
    wb.properties.title = "沪蓉科技产业协同发展合作建议书讨论稿工作底表"
    wb.properties.creator = "上海市杨浦区科技企业联合会"
    wb.properties.subject = "2027—2029年战略合作讨论稿"
    wb.properties.description = "讨论文件，不构成法律约束力协议。面积与会费为交流参考口径。"

    overview(wb.active)
    directions(wb.create_sheet("01-五个方向"))
    content(wb.create_sheet("02-合作内容"))
    roles(wb.create_sheet("03-职责分工"))
    plan(wb.create_sheet("04-推进计划"))
    outcomes(wb.create_sheet("05-首年成果"))
    events(wb.create_sheet("06-活动与路演"))
    commercial(wb.create_sheet("07-商业机制"))
    priorities(wb.create_sheet("08-优先抓手"))
    confirm(wb.create_sheet("09-发送前确认"))

    colors = {
        "00-总览": NAVY,
        "01-五个方向": TEAL,
        "02-合作内容": TEAL,
        "03-职责分工": TEAL,
        "04-推进计划": GOLD,
        "05-首年成果": GOLD,
        "06-活动与路演": GOLD,
        "07-商业机制": "8E3B34",
        "08-优先抓手": "8E3B34",
        "09-发送前确认": NAVY,
    }
    for name, color in colors.items():
        wb[name].sheet_properties.tabColor = color

    path = "deliverables/杨浦科技企业联合会_成都民营产业园集团_战略合作建议书_工作底表.xlsx"
    wb.save(path)
    print(f"已生成 {path} ，共 {len(wb.sheetnames)} 张工作表")


def overview(ws):
    ws.title = "00-总览"
    apply_page(ws, "00 总览")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A8"
    ws.print_title_rows = "1:2"
    set_widths(ws, [18, 28, 42, 28, 24, 36])

    ws.merge_cells("A1:F1")
    ws["A1"] = "上海市杨浦区科技企业联合会  ×  成都民营产业园集团"
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = Alignment(vertical="center")
    ws["A1"].fill = CREAM_FILL
    for col in range(1, 7):
        ws.cell(1, col).fill = CREAM_FILL
    ws.row_dimensions[1].height = 34

    ws.merge_cells("A2:F2")
    ws["A2"] = "战略合作建议书（讨论稿）  ·  共建沪蓉科技产业协同发展与创新企业服务平台  ·  2026年11月"
    ws["A2"].font = Font(name=FONT, size=12, color=TEAL)
    ws["A2"].fill = CREAM_FILL
    for col in range(1, 7):
        ws.cell(2, col).fill = CREAM_FILL
    ws.row_dimensions[2].height = 22

    info = [
        ("建议发起单位", "上海市杨浦区科技企业联合会"),
        ("拟合作单位", "成都民营产业园集团（签约主体待确认）"),
        ("拟合作周期", "2027—2029年"),
        ("交流依据", "2026年9月24日双方洽谈"),
        ("文件性质", "讨论稿，不构成法律约束力协议"),
        ("合作原则", "资源共享、产业协同、市场化运作、优势互补、长期共赢"),
    ]
    for index, (label, value) in enumerate(info):
        row = 4 + index // 2
        col = 1 if index % 2 == 0 else 4
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col)
        label_cell = ws.cell(row, col, label)
        label_cell.font = HEAD_FONT
        label_cell.fill = NAVY_FILL
        label_cell.alignment = CENTER
        label_cell.border = BORDER
        ws.merge_cells(start_row=row, start_column=col + 1, end_row=row, end_column=col + 2)
        value_cell = ws.cell(row, col + 1, value)
        value_cell.font = BOLD
        value_cell.alignment = LEFT_CENTER
        value_cell.border = BORDER
        value_cell.fill = WHITE_FILL
        ws.cell(row, col + 2).border = BORDER
        ws.cell(row, col + 2).fill = WHITE_FILL
        ws.row_dimensions[row].height = 26

    ws.merge_cells("A8:F8")
    ws["A8"] = "五个合作方向"
    ws["A8"].font = Font(name=FONT, size=13, bold=True, color=WHITE)
    ws["A8"].fill = TEAL_FILL
    ws["A8"].alignment = LEFT_CENTER
    for col in range(1, 7):
        ws.cell(8, col).fill = TEAL_FILL
    ws.row_dimensions[8].height = 24

    headers = ["序号", "方向", "首年要看到的结果", "联合会主要贡献", "集团主要贡献", "放大条件"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(9, col, header)
        cell.font = HEAD_FONT
        cell.fill = NAVY_FILL
        cell.alignment = CENTER
        cell.border = BORDER
    directions_rows = [
        ("01", "上海产业招商", "上海园区形成可核对的特色招商进展", "企业筛选、推荐、招商活动", "空间、租赁商务、落地服务", "先核对可招商面积与条件"),
        ("02", "人工智能产业集聚", "跑通项目筛选、对接和活动机制", "产业资源、高校与项目", "场景、孵化与载体", "平台名称另行确认"),
        ("03", "科技金融赋能", "完成首批企业与投资机构匹配", "机构、路演、专业服务", "企业需求与项目组织", "投资由持牌主体独立开展"),
        ("04", "沪蓉双城联动", "完成需求摸排和至少一次考察", "上海企业与科研创新资源", "成都园区、产业链、在地服务", "明确成都总部授权"),
        ("05", "国际化出海", "建立新加坡及出海交流机制", "国际渠道与专业机构", "海外平台与新加坡业务资源", "先做资源库，再扩平台"),
    ]
    for offset, row in enumerate(directions_rows):
        excel_row = 10 + offset
        for col, value in enumerate(row, 1):
            cell = ws.cell(excel_row, col, value)
            cell.font = BOLD if col <= 2 else BODY
            cell.alignment = CENTER if col == 1 else WRAP
            cell.border = BORDER
            cell.fill = WHITE_FILL if offset % 2 == 0 else ZEBRA_FILL
        ws.row_dimensions[excel_row].height = 36

    ws.merge_cells("A16:F16")
    ws["A16"] = "建议优先推进"
    ws["A16"].font = Font(name=FONT, size=13, bold=True, color=WHITE)
    ws["A16"].fill = PatternFill("solid", fgColor="8E3B34")
    for col in range(1, 7):
        ws.cell(16, col).fill = PatternFill("solid", fgColor="8E3B34")
        ws.cell(16, col).font = Font(name=FONT, size=13, bold=True, color=WHITE)
    ws.row_dimensions[16].height = 24

    pri_headers = ["优先级", "抓手", "商务价值", "首步动作", "费用归属", "依赖确认"]
    for col, header in enumerate(pri_headers, 1):
        cell = ws.cell(17, col, header)
        cell.font = HEAD_FONT
        cell.fill = NAVY_FILL
        cell.alignment = CENTER
        cell.border = BORDER
    priorities_rows = [
        ("第一优先", "副会长单位入会 + 园区联合招商", "建立正式关系，并形成可核对成果", "确认申请主体，梳理可招商空间", "会费与招商服务费分开", "工商主体、签约权限、会费标准"),
        ("第二优先", "人工智能活动 + 科技金融路演", "为企业和对接机构提供持续到场场景", "选定首场主题、场地与嘉宾", "活动费另行约定", "场地场次、容量和设施"),
        ("第三优先", "成都约80万㎡载体的沪蓉招商", "由上海试点延伸为集团层面的长期协作", "对接成都总部产业合作负责人", "跨区域项目另签专项协议", "载体范围、总部授权"),
    ]
    fills = [SOFT_TEAL_FILL, SOFT_GOLD_FILL, SOFT_RED_FILL]
    for offset, row in enumerate(priorities_rows):
        excel_row = 18 + offset
        for col, value in enumerate(row, 1):
            cell = ws.cell(excel_row, col, value)
            cell.font = BOLD if col <= 2 else BODY
            cell.alignment = WRAP
            cell.border = BORDER
            cell.fill = fills[offset]
        ws.row_dimensions[excel_row].height = 40

    ws.merge_cells("A22:F22")
    ws["A22"] = (
        "费用原则：会费、招商服务费、活动合作费分别约定。副会长单位参考会费为每年人民币10,000元，"
        "以正式会员管理规定为准；该会费不包含后续招商及专业服务。载体面积为2026年9月24日交流口径，发送正式文本前请核对。"
    )
    ws["A22"].font = Font(name=FONT, size=10.5, color=INK)
    ws["A22"].alignment = Alignment(wrap_text=True, vertical="center")
    ws["A22"].fill = SOFT_GOLD_FILL
    for col in range(1, 7):
        ws.cell(22, col).fill = SOFT_GOLD_FILL
        ws.cell(22, col).border = BORDER
    ws.row_dimensions[22].height = 42

    ws.merge_cells("A23:F23")
    ws["A23"] = "使用顺序：00总览 → 09发送前确认 → 07商业机制 → 04推进计划。黄色单元格为待填写项。工作表可筛选、可打印。"
    ws["A23"].font = NOTE_FONT
    ws["A23"].alignment = LEFT_CENTER
    ws.row_dimensions[23].height = 22
    ws.page_setup.fitToHeight = 1
    ws.print_options.horizontalCentered = True
    ws.oddFooter.left.text = "00 总览"


def directions(ws):
    apply_page(ws, "01 五个方向")
    headers = ["序号", "合作方向", "对应目标", "重点对象", "联合会动作", "集团动作", "首年可见进展", "暂不承诺"]
    write_title(
        ws,
        len(headers),
        "五个合作方向与总体目标",
        "方向来自讨论稿。首年可见进展用于内部对照，正式量化指标在合作启动后共同确定。",
    )
    write_header(ws, headers)
    rows = [
        [
            "01",
            "上海产业招商",
            "把上海园区建成连接复旦周边、杨浦企业和成都载体的服务节点。",
            "有扩租、搬迁、研发办公或区域布局需求的科技企业。",
            "筛选、推荐、组织考察和招商活动。",
            "提供空间、租赁条件和落地服务。",
            "可核对的推荐、带看和商务洽谈记录。",
            "不预设出租率或成交面积。",
        ],
        [
            "02",
            "人工智能产业集聚",
            "提升园区在人工智能、机器人、集成电路、先进制造和科技服务上的集聚度。",
            "人工智能、具身智能、硬科技、集成电路、可控核聚变等相关项目。",
            "产业资源、高校团队、活动组织。",
            "在园企业、应用场景、孵化与载体。",
            "项目清单、专题交流和产业链协作线索。",
            "平台名称和挂牌另行协商。",
        ],
        [
            "03",
            "科技金融赋能",
            "建立融资路演、机构对接、投后服务和成长辅导机制。",
            "有明确融资需求、技术基础或商业化进展的企业。",
            "引入机构和专业服务，组织路演。",
            "汇集企业需求，推荐在园及储备项目。",
            "首批企业与机构的匹配记录。",
            "不代为投资，不承诺融资结果。",
        ],
        [
            "04",
            "沪蓉双城联动",
            "促进上海成果、企业和资本与成都空间、产业链需求匹配。",
            "拟赴成都布局的上海企业，以及拟来上海拓展的成都企业。",
            "组织上海侧资源和企业。",
            "提供成都空间、产业链和考察支持。",
            "需求摸排，以及至少一次双城交流或考察。",
            "首年不承诺跨城落地数量。",
        ],
        [
            "05",
            "国际化出海",
            "面向新加坡、东南亚及其他重点市场，探索出海服务和跨境合作。",
            "有国际市场拓展需求的科技企业。",
            "国际渠道和专业机构对接。",
            "海外孵化经验、新加坡业务及相关平台资源。",
            "专项交流机制和出海服务资源名录。",
            "新的海外平台待资源成熟后再议。",
        ],
    ]
    widths = [8, 18, 36, 32, 28, 28, 30, 26]
    add_rows(ws, rows, widths=widths)
    set_widths(ws, widths)


def content(ws):
    apply_page(ws, "02 合作内容")
    headers = ["板块", "事项", "合作内容", "联合会动作", "集团动作", "交付物", "建议阶段", "费用归属"]
    write_title(
        ws,
        len(headers),
        "重点合作内容分解",
        "本表把讨论稿第三章拆成可执行事项。费用归属与07-商业机制一致：会费不覆盖招商服务和活动专业服务。",
    )
    write_header(ws, headers)
    rows = [
        ["产业招商", "招商信息共享", "交换可招商面积、租赁条件、空间产品、适配行业和配套。", "据此更新企业匹配依据。", "提供并维护信息。", "招商条件一览", "首月", "框架内，不单列"],
        ["产业招商", "企业筛选", "筛选有扩租、搬迁、研发办公、区域布局需求的企业。", "结合《杨浦一百》和资源库筛选。", "反馈空间适配意见。", "首批企业名单", "首月", "框架内，不单列"],
        ["产业招商", "考察与带看", "定期组织考察、路演、推介和现场带看。", "组织企业。", "安排场地与接待。", "考察记录", "第2–3个月起", "活动费另计"],
        ["产业招商", "重点跟踪", "跟踪被投企业、高校成果转化企业和成长型科技企业。", "提供线索。", "跟进落地条件。", "跟踪台账", "持续", "招商服务费另计"],
        ["产业招商", "三个特色方向", "人工智能与具身智能；硬科技、集成电路与先进制造；科技服务与产业国际化。", "按方向组织企业。", "明确空间和政策适配。", "定位说明", "首月", "框架内，不单列"],
        ["产业招商", "租赁与服务组合", "探索灵活租赁、成长型企业优惠和产业服务组合。", "反馈企业阶段与需求。", "设计商务方案。", "方案要点", "按项目", "专项协议"],
        ["产业招商", "合作方式", "资源推荐、招商代理、专题活动招商，按项目选择。", "按约定提供服务。", "确认成交与结算。", "专项服务协议", "按项目", "招商服务费另计"],
        ["人工智能", "项目筛选", "人工智能、机器人、集成电路、可控核聚变等领域。", "提供项目线索。", "提供场景和载体意见。", "项目清单", "第2–3个月", "框架内，不单列"],
        ["人工智能", "专题交流", "研发合作、产业应用、成果转化、企业融资。", "组织议题与嘉宾。", "组织园区企业。", "交流纪要", "按活动", "活动费另计"],
        ["人工智能", "三方对接", "高校科研团队、科技企业、投资机构。", "邀请与匹配。", "提供场景。", "对接记录", "持续", "框架内，不单列"],
        ["人工智能", "产业链协作", "园区企业与联合会服务企业协作。", "撮合需求。", "组织在园企业。", "合作线索", "持续", "按项目另计"],
        ["人工智能", "活动形态", "开发者交流、技术展示、产业沙龙、项目路演。", "内容策划。", "场地支持。", "活动方案", "按活动", "活动费另计"],
        ["人工智能", "成都孵化延伸", "具备条件的项目，协调成都总部孵化平台。", "推荐项目。", "成都平台评估承接。", "承接意见", "第4–12个月", "专项协议"],
        ["人工智能", "平台设立", "名称、设立形式、对外挂牌另行协商。", "共同确认。", "共同确认。", "确认纪要", "待协商", "另行约定"],
        ["科技金融", "机构引入", "创投、产业基金、银行、证券和专业服务机构。", "协助引入。", "组织企业需求。", "机构与需求清单", "第2–3个月", "框架内，不单列"],
        ["科技金融", "季度路演", "四类主题，优先有融资需求或商业化进展的项目。", "组织机构和议程。", "组织项目企业。", "路演安排", "季度", "活动费另计"],
        ["科技金融", "综合培育", "空间、辅导、资本对接、产业资源。", "资本与辅导资源。", "空间与在地服务。", "培育方案", "按企业", "专项协议"],
        ["科技金融", "合规边界", "基金投资和投资管理由具备资质的主体依法独立开展。", "只组织对接。", "只组织对接。", "合规说明", "持续", "独立约定"],
        ["沪蓉协同", "通道分工", "上海组织创新资源，成都提供空间、产业链和考察。", "上海资源组织。", "成都在地服务。", "对接机制", "首月启动", "框架内，不单列"],
        ["沪蓉协同", "上海企业西进", "支持设立研发、生产或区域业务机构。", "筛选企业。", "成都选址与落地支持。", "意向清单", "第4–12个月", "招商服务费另计"],
        ["沪蓉协同", "成都企业东进", "支持来上海开展研发合作、市场拓展和资本对接。", "安排上海对接。", "推荐企业。", "意向清单", "第4–12个月", "按项目另计"],
        ["沪蓉协同", "企业间合作", "产品应用、供应链合作和项目落地。", "撮合。", "撮合。", "合作线索", "持续", "按项目另计"],
        ["沪蓉协同", "双城考察", "每年组织一次上海—成都产业考察。", "组织上海侧。", "组织成都参访与接待。", "考察方案", "首年至少1次", "活动费另计"],
        ["国际化", "新加坡合作", "市场拓展、商务对接、项目孵化、跨境商业合作。", "提供国际渠道。", "提供海外平台和新加坡资源。", "交流机制", "第4–12个月", "专项协议"],
        ["国际化", "出海资源库", "市场信息、商务渠道、法律税务、知识产权、跨境业务。", "对接专业机构。", "对接海外资源。", "资源名录", "第4–12个月", "服务费另计"],
        ["国际化", "出海活动", "出海沙龙、东南亚专题、国际项目对接。", "内容与企业组织。", "海外资源对接。", "活动方案", "按活动", "活动费另计"],
        ["国际化", "其他平台", "按实际资源逐步探索其他海外创新平台。", "评估企业需求。", "提供平台线索。", "机会清单", "条件成熟后", "另行论证"],
        ["活动品牌", "场地共享", "为适宜规模的科技企业活动提供场地。", "提出需求。", "确认场次、容量、设施。", "使用安排", "首月明确", "活动费另计"],
        ["活动品牌", "五类品牌活动", "人工智能、双城、投融资、成果转化、出海五类主题。", "策划、邀请、组织。", "场地及必要配套。", "年度活动安排", "首年4–6场", "活动费另计"],
        ["活动品牌", "《杨浦一百》", "走访、案例交流、产业研究；宣传须授权并经企业同意。", "研究方法与案例组织。", "组织园区企业。", "案例与对接记录", "首年", "框架内，不单列"],
    ]
    widths = [14, 18, 42, 28, 26, 20, 16, 18]
    add_rows(ws, rows, widths=widths)
    set_widths(ws, widths)
    ws.auto_filter.ref = f"A3:H{3 + len(rows)}"


def roles(ws):
    apply_page(ws, "03 职责分工")
    headers = ["合作事项", "杨浦区科技企业联合会", "成都民营产业园集团", "共同事项", "不宜由单方承诺的内容"]
    write_title(ws, len(headers), "双方资源与职责分工", "与讨论稿第四章一致。共同事项需要双方负责人口头或书面确认后执行。")
    write_header(ws, headers)
    rows = [
        ["园区招商", "企业筛选、项目推荐、招商活动", "空间提供、租赁商务、落地服务", "确定重点行业和首批企业", "出租率、免租期、成交价格"],
        ["人工智能合作", "产业资源、创新项目、活动组织", "企业孵化、产业场景、载体支持", "确认平台名称是否对外使用", "未经协商的挂牌和品牌授权"],
        ["科技金融", "投资机构、融资路演、专业服务对接", "企业需求组织、项目推荐", "确定路演主题和项目标准", "融资结果、投资决策、收益承诺"],
        ["沪蓉产业协同", "上海企业和科研创新资源", "成都园区、产业链及在地服务", "每年考察的主题、规模和时间", "未经授权的跨区域招商承诺"],
        ["企业国际化", "国际合作渠道及专业机构对接", "海外平台及相关业务资源对接", "先核定新加坡及现有海外资源", "新增海外主体和联合设点"],
        ["品牌活动", "内容策划、嘉宾邀请、企业组织", "场地及必要配套支持", "场次、容量、议程和费用", "默认全部活动免费或无限场次"],
    ]
    widths = [18, 36, 36, 32, 34]
    add_rows(ws, rows, widths=widths)
    set_widths(ws, widths)


def plan(ws):
    apply_page(ws, "04 推进计划")
    headers = ["阶段", "时间", "序号", "工作事项", "主要责任", "交付物", "完成标志", "状态"]
    write_title(
        ws,
        len(headers),
        "十二个月推进计划",
        "状态列请双方更新。黄色单元格可填写。完成标志用于判断阶段是否可以进入下一步，不是对外考核合同。",
    )
    write_header(ws, headers)
    rows = [
        ["第一阶段", "首月", 1, "双方确定业务负责人", "双方", "负责人及联系方式", "对接人可以独立推进日常事项", "待启动"],
        ["第一阶段", "首月", 2, "建立常态化对接机制", "双方", "会议节奏和纪要模板", "至少约定双周或月度对接", "待启动"],
        ["第一阶段", "首月", 3, "梳理上海可招商空间及租赁政策", "集团", "面积、条件、产品、行业、配套", "联合会可以据此筛选企业", "待启动"],
        ["第一阶段", "首月", 4, "明确首批重点服务企业名单", "联合会", "候选企业及需求类型", "名单经双方确认优先顺序", "待启动"],
        ["第一阶段", "首月", 5, "筛选首场产业活动", "双方", "主题、规模、场地需求", "确定一场可执行活动", "待启动"],
        ["第一阶段", "首月", 6, "对接成都总部产业合作负责人", "集团牵头", "负责人及授权范围", "双城事项有明确接口人", "待启动"],
        ["第一阶段", "首月", 7, "确认重点项目与优先级", "双方", "优先级确认记录", "三个抓手的先后顺序得到确认", "待启动"],
        ["第二阶段", "第2—3个月", 8, "举办首场人工智能或科技金融活动", "双方", "活动方案、签到和纪要", "活动按方案完成", "待启动"],
        ["第二阶段", "第2—3个月", 9, "建立招商线索与项目台账", "双方", "台账字段和更新人", "每条线索有状态和下一步", "待启动"],
        ["第二阶段", "第2—3个月", 10, "组织实地看场与商务洽谈", "集团主办、联合会组织企业", "看场记录和跟进事项", "至少形成一批有效洽谈", "待启动"],
        ["第二阶段", "第2—3个月", 11, "完成首批融资企业与机构匹配", "联合会牵头", "匹配清单", "企业需求与机构方向相互可见", "待启动"],
        ["第二阶段", "第2—3个月", 12, "启动沪蓉需求摸排", "双方", "东进、西进需求表", "知道哪些企业具备跨城意愿", "待启动"],
        ["第三阶段", "第4—12个月", 13, "将招商、孵化、金融转入常态机制", "双方", "季度工作安排", "不再依赖单次活动推进", "待启动"],
        ["第三阶段", "第4—12个月", 14, "完成至少一次沪蓉考察或交流", "双方", "考察方案和纪要", "有参访、交流和后续项目", "待启动"],
        ["第三阶段", "第4—12个月", 15, "形成一项可持续协同或出海项目", "双方", "项目说明和责任人", "项目具备继续运营的条件", "待启动"],
        ["第三阶段", "第4—12个月", 16, "评估是否延伸到成都总部及其他载体", "集团牵头", "扩展建议", "有明确的可招商范围和授权", "待启动"],
        ["第三阶段", "第4—12个月", 17, "复盘首年并确定下一年度量化指标", "双方", "复盘纪要", "指标由双方共同确认", "待启动"],
    ]
    widths = [14, 14, 8, 38, 24, 32, 32, 12]
    last = add_rows(ws, rows, input_cols={8}, widths=widths)
    status_validation(ws, f"H4:H{last}")
    set_widths(ws, widths)
    ws.conditional_formatting.add(
        f"H4:H{last}",
        FormulaRule(formula=['$H4="已完成"'], fill=SOFT_TEAL_FILL),
    )
    ws.conditional_formatting.add(
        f"H4:H{last}",
        FormulaRule(formula=['$H4="进行中"'], fill=SOFT_GOLD_FILL),
    )
    ws.conditional_formatting.add(
        f"H4:H{last}",
        FormulaRule(formula=['$H4="暂缓"'], fill=SOFT_RED_FILL),
    )


def outcomes(ws):
    apply_page(ws, "05 首年成果")
    headers = ["序号", "建议成果", "为什么把它列为成果", "当前表述", "启动后补充的量化指标", "主要责任", "复核时点"]
    write_title(
        ws,
        len(headers),
        "首年预期成果",
        "讨论稿将下列事项列为建议性目标。黄色单元格留待启动后填写，避免在未核对资源前写入不可执行的数字。",
    )
    write_header(ws, headers)
    rows = [
        [1, "稳定的联合工作机制", "没有固定负责人和节奏，后续事项无法持续。", "负责人、对接节奏、项目台账均在运行。", "", "双方", "第3个月"],
        [2, "4—6场产业活动", "活动是企业、机构和园区见面的载体。", "围绕人工智能、硬科技、科技金融。", "", "联合会牵头", "每季度"],
        [3, "有需求的项目储备", "招商和融资都需要可跟踪的项目池。", "项目具备空间、融资或跨区域布局中的至少一种明确需求。", "", "双方", "第6个月"],
        [4, "上海园区招商成果", "试点需要看见企业落地进展。", "在特色方向上形成可核对的洽谈或入驻进展。", "", "集团为主", "第12个月"],
        [5, "一次沪蓉交流或考察", "把双城合作从意向变成一次真实组织。", "包含参访、企业交流和项目洽谈。", "", "双方", "第12个月"],
        [6, "一项可持续项目", "避免合作停留在单次活动。", "双城协同或国际化服务中，有一项可继续运营。", "", "双方", "第12个月"],
    ]
    widths = [8, 24, 36, 40, 28, 16, 14]
    last = add_rows(ws, rows, input_cols={5}, widths=widths)
    set_widths(ws, widths)
    ws.cell(last + 2, 1, "说明：量化指标由双方在资源核对后共同填写，本表不预置成交面积、融资金额或到场人数。")
    ws.merge_cells(start_row=last + 2, start_column=1, end_row=last + 2, end_column=7)
    ws.cell(last + 2, 1).font = NOTE_FONT
    ws.row_dimensions[last + 2].height = 24


def events(ws):
    apply_page(ws, "06 活动与路演")
    headers = ["类型", "名称", "目的", "建议节奏", "联合会", "集团", "费用归属", "状态"]
    write_title(
        ws,
        len(headers),
        "活动品牌与季度路演清单",
        "五类活动是品牌菜单，首年总量建议 4—6 场，不必每类都举办。路演按季度专题组织。",
    )
    write_header(ws, headers)
    rows = [
        ["品牌活动", "杨浦人工智能与硬科技创新沙龙", "集聚人工智能、机器人和硬科技企业。", "首年择期", "策划、邀请、组织", "场地及配套", "活动费另计", "待启动"],
        ["品牌活动", "沪蓉科技企业产业协同交流会", "收集东进、西进和产业链合作需求。", "配合双城考察", "上海企业组织", "成都资源对接", "活动费另计", "待启动"],
        ["品牌活动", "科技企业投融资与产业资本对接会", "让企业和投资机构直接见面。", "可与季度路演合并", "机构邀请", "项目企业组织", "活动费另计", "待启动"],
        ["品牌活动", "科技成果转化与企业成长专题沙龙", "连接高校成果、园区场景和成长辅导。", "首年择期", "高校与内容", "在园企业", "活动费另计", "待启动"],
        ["品牌活动", "科技企业出海与国际市场拓展交流会", "讨论新加坡、东南亚及跨境服务需求。", "资源核定后", "国际渠道", "海外平台资源", "活动费另计", "待启动"],
        ["季度路演", "人工智能与具身智能", "展示有融资或合作需求的项目。", "季度专题之一", "机构与议程", "项目组织", "活动费另计", "待启动"],
        ["季度路演", "硬科技与集成电路", "对接硬科技项目与产业资本。", "季度专题之一", "机构与议程", "项目组织", "活动费另计", "待启动"],
        ["季度路演", "新能源及未来产业", "覆盖新能源和双方认可的未来产业项目。", "季度专题之一", "机构与议程", "项目组织", "活动费另计", "待启动"],
        ["季度路演", "专精特新与先进制造", "服务制造类成长企业的融资与场景需求。", "季度专题之一", "机构与议程", "项目组织", "活动费另计", "待启动"],
        ["双城活动", "上海—成都科技企业产业考察", "园区参访、企业交流、项目洽谈。", "每年一次，首年至少一次", "上海团组", "成都接待与参访", "活动费另计", "待启动"],
        ["研究合作", "《杨浦一百》园区企业走访与案例", "形成可授权使用的企业案例和资源对接。", "首年择机", "研究与案例方法", "组织受访企业", "框架内；宣传另经同意", "待启动"],
    ]
    widths = [14, 36, 36, 22, 20, 20, 18, 12]
    last = add_rows(ws, rows, input_cols={8}, widths=widths)
    status_validation(ws, f"H4:H{last}")
    set_widths(ws, widths)


def commercial(ws):
    apply_page(ws, "07 商业机制")
    headers = ["机制", "解决什么问题", "建议文件", "费用如何处理", "是否包含在会员会费中", "边界"]
    write_title(
        ws,
        len(headers),
        "合作模式与商业机制",
        "核心原则：会费、招商服务费、活动合作费分别约定。参考会费是建立关系的入口，不包干后续专业服务。",
    )
    write_header(ws, headers)
    rows = [
        [
            "战略合作机制",
            "先确定资源共享、产业协作和联合活动的方向。",
            "战略合作框架协议",
            "框架协议本身不替代专项收费。",
            "不适用",
            "框架协议不自动构成每一项业务的成交承诺。",
        ],
        [
            "会员合作机制",
            "让集团上海项目进入联合会的正式会员体系。",
            "入会申请及审批文件",
            "参考会费每年人民币10,000元，以正式会员管理规定为准。",
            "这就是会费本身",
            "按现行制度和审批程序办理，不先行承诺审批结果。",
        ],
        [
            "专项招商机制",
            "企业租赁、产业项目引进和其他商业落地。",
            "专项服务协议",
            "按投入和成交另行约定范围、认定、收费和结算。",
            "不包含",
            "没有专项协议，不视为已委托有偿招商。",
        ],
        [
            "产业活动机制",
            "联合举办沙龙、路演、考察和对接会。",
            "活动方案；涉及收费时另签协议",
            "优先使用现有资源。赞助和服务费按方案结算。",
            "不包含",
            "场地支持的场次、容量和设施须事先确认。",
        ],
        [
            "重大项目机制",
            "园区改造、载体运营、跨区域投资、基金合作。",
            "专项论证后的商务协议",
            "单独洽谈，不从会费或单场活动费中列支。",
            "不包含",
            "基金投资由具备资质的市场主体依法独立开展。",
        ],
    ]
    fills = [WHITE_FILL, SOFT_GOLD_FILL, SOFT_TEAL_FILL, WHITE_FILL, ZEBRA_FILL]
    widths = [18, 36, 28, 40, 24, 40]
    add_rows(ws, rows, fills=fills, widths=widths)
    set_widths(ws, widths)

    ws.cell(10, 1, "三类费用对照")
    ws.merge_cells("A10:F10")
    ws["A10"].font = Font(name=FONT, size=13, bold=True, color=WHITE)
    for col in range(1, 7):
        ws.cell(10, col).fill = TEAL_FILL
        ws.cell(10, col).font = Font(name=FONT, size=13, bold=True, color=WHITE)
    ws.row_dimensions[10].height = 24
    fee_headers = ["费用类型", "参考口径", "包含什么", "不包含什么", "约定文件", "当前状态"]
    for col, header in enumerate(fee_headers, 1):
        cell = ws.cell(11, col, header)
        cell.font = HEAD_FONT
        cell.fill = NAVY_FILL
        cell.alignment = CENTER
        cell.border = BORDER
    fees = [
        ["副会长单位年费", "每年人民币10,000元，最终以正式规定为准", "按会员制度享有的会员权益", "招商代理、成交服务、活动赞助、专业顾问和基金服务", "入会审批文件", "待确认"],
        ["招商服务费", "按实际投入和成交结果另行约定", "约定范围内的推荐、代理、促成和结算", "未写入专项协议的服务", "专项服务协议", "待约定"],
        ["活动合作费", "按单场方案约定；能用现有资源则不新增收费", "约定的策划、组织、场地、嘉宾和物料", "默认无限场次或全部免费", "活动方案或独立协议", "待约定"],
    ]
    for offset, row in enumerate(fees):
        excel_row = 12 + offset
        for col, value in enumerate(row, 1):
            cell = ws.cell(excel_row, col, value)
            cell.font = BOLD if col == 1 else BODY
            cell.alignment = WRAP
            cell.border = BORDER
            cell.fill = SOFT_GOLD_FILL if offset == 0 else WHITE_FILL
        ws.row_dimensions[excel_row].height = 48
    ws.row_dimensions[12].height = 52


def priorities(ws):
    apply_page(ws, "08 优先抓手")
    headers = ["优先级", "合作抓手", "商务价值", "首月动作", "需要的对方条件", "成功时看什么", "主要风险"]
    write_title(
        ws,
        len(headers),
        "三个优先合作抓手",
        "排序按启动难度和关系建立顺序。第三优先决定合作能否从上海试点延伸到集团载体。",
    )
    write_header(ws, headers)
    rows = [
        [
            "第一优先",
            "副会长单位入会 + 园区联合招商",
            "尽快建立正式合作关系，并形成可核对的业务成果。",
            "确认申请主体；梳理上海可招商空间；确定首批企业。",
            "工商主体、签约权限、会费标准、可招商条件。",
            "入会程序启动，并有企业推荐与带看记录。",
            "把后续招商和专业服务都写进会员权益。",
        ],
        [
            "第二优先",
            "人工智能产业活动 + 科技金融路演",
            "集聚企业和投资机构，为园区带来持续的项目来源。",
            "选定首场主题、规模、场地和双方负责人。",
            "场地场次、容量、设施，以及可邀请的在园企业。",
            "首场活动完成，并留下可跟踪的企业或机构名单。",
            "活动办完没有台账，无法转化成招商或融资线索。",
        ],
        [
            "第三优先",
            "成都约80万平方米载体的沪蓉招商合作",
            "由上海约6000平方米试点，延伸为面向长三角科技企业的集团层面协作。",
            "对接成都总部产业合作负责人，核对载体范围。",
            "总部授权、可招商区域、产业链配套和考察安排。",
            "形成东进、西进需求表，并排定一次双城考察。",
            "只围绕上海物业，合作空间停留在单个项目。",
        ],
    ]
    fills = [SOFT_TEAL_FILL, SOFT_GOLD_FILL, SOFT_RED_FILL]
    widths = [12, 32, 36, 36, 32, 34, 32]
    add_rows(ws, rows, fills=fills, widths=widths)
    set_widths(ws, widths)


def confirm(ws):
    apply_page(ws, "09 发送前确认")
    headers = ["序号", "确认事项", "建议确认方", "为什么要先确认", "当前状态", "确认结果", "影响的合作内容"]
    write_title(
        ws,
        len(headers),
        "正式文本发出前的确认清单",
        "黄色单元格请填写。讨论稿可以先用于交流；替换成正式签约文本前，建议完成本表。",
    )
    write_header(ws, headers)
    rows = [
        [1, "集团准确工商主体名称", "集团", "当前使用“成都民营产业园集团”作为暂定表述。", "待确认", "", "封面、协议主体、开票与用印"],
        [2, "上海项目签约主体及权限", "集团", "上海载体可能由项目公司运营，不一定由集团总部签约。", "待确认", "", "园区招商、场地、会员申请主体"],
        [3, "成都总部签约或授权权限", "集团", "双城和集团层面合作需要总部确认接口与授权。", "待确认", "", "沪蓉招商、成都孵化、年度考察"],
        [4, "成都约80万平方米的统计范围", "集团", "需明确区域、业态、自有或可调度、可对外招商面积。", "待确认", "", "第三优先抓手、对外表述"],
        [5, "上海约6000平方米的统计范围", "集团", "需明确自有面积、可招商面积、产品形态和租赁条件。", "待确认", "", "首批试点和招商筛选"],
        [6, "副会长单位申请主体", "双方", "讨论稿建议为集团上海项目，最终按审批程序确定。", "待确认", "", "会员合作机制"],
        [7, "副会长年费标准", "联合会", "交流参考口径为每年人民币10,000元。", "待确认", "", "会员费用，不影响专项收费"],
        [8, "会费与专项费用分列", "双方", "避免会费被理解为包干招商、活动和专业服务。", "待确认", "", "全部商业机制"],
        [9, "活动场地支持条件", "集团", "场次、容量、设施和使用安排尚未写成确定承诺。", "待确认", "", "产业活动和路演"],
        [10, "双方业务负责人", "双方", "首月机制依赖有权推进的日常接口人。", "待确认", "", "推进计划全部事项"],
        [11, "平台名称是否对外使用", "双方", "“沪蓉人工智能与硬科技产业协同创新中心”只是建议名称。", "待确认", "", "对外宣传和挂牌"],
        [12, "企业案例与信息授权", "双方及企业", "《杨浦一百》延伸宣传须双方授权并经企业同意。", "待确认", "", "案例、参访和公开传播"],
        [13, "文件法律性质", "双方", "本套文件是讨论稿，正式条款以后续确认文件为准。", "已明确", "讨论稿，不构成法律约束力协议", "全部文本"],
    ]
    widths = [8, 30, 16, 42, 14, 36, 28]
    last = add_rows(ws, rows, input_cols={5, 6}, widths=widths)
    confirm_validation(ws, f"E4:E{last}")
    ws.cell(last, 5).fill = SOFT_TEAL_FILL
    set_widths(ws, widths)


if __name__ == "__main__":
    build()
