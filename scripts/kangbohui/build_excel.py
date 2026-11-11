"""可勾选、可回款、可自动判断的执行台账。"""

from __future__ import annotations

import os
from datetime import date

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT = os.path.normpath(
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "..",
        "deliverables",
        "2026康博会专题分论坛",
        "03-内部-执行管理台账.xlsx",
    )
)

INK = "102433"
TEAL = "0F6E62"
GOLD = "B8893A"
CREAM = "F6F3EC"
MIST = "F3F6F5"
YELLOW = "F8F1E3"
ROSE = "F8EEEE"
GREEN = "E5F4EC"
WHITE = "FFFFFF"
SLATE = "1E2A32"
MUTED = "5C6B73"
RED = "8C3E3E"

thin = Border(
    left=Side(style="thin", color="E3DCD0"),
    right=Side(style="thin", color="E3DCD0"),
    top=Side(style="thin", color="E3DCD0"),
    bottom=Side(style="thin", color="E3DCD0"),
)
fill_ink = PatternFill("solid", fgColor=INK)
fill_teal = PatternFill("solid", fgColor=TEAL)
fill_cream = PatternFill("solid", fgColor=CREAM)
fill_mist = PatternFill("solid", fgColor=MIST)
fill_yellow = PatternFill("solid", fgColor=YELLOW)
fill_rose = PatternFill("solid", fgColor=ROSE)
fill_green = PatternFill("solid", fgColor=GREEN)
fill_white = PatternFill("solid", fgColor=WHITE)
fill_gold = PatternFill("solid", fgColor="E8D7B0")

font_title = Font(name="微软雅黑", size=18, bold=True, color=INK)
font_header = Font(name="微软雅黑", size=10, bold=True, color=WHITE)
font_base = Font(name="微软雅黑", size=10, color=SLATE)
font_small = Font(name="微软雅黑", size=9, color=MUTED)
font_bold = Font(name="微软雅黑", size=10, bold=True, color=INK)
font_kpi = Font(name="微软雅黑", size=16, bold=True, color=TEAL)
font_white = Font(name="微软雅黑", size=12, bold=True, color=WHITE)

wrap = Alignment(wrap_text=True, vertical="center", horizontal="left")
center = Alignment(wrap_text=True, vertical="center", horizontal="center")


def apply_header_row(ws, row, cols):
    for col in range(1, cols + 1):
        cell = ws.cell(row, col)
        cell.font = font_header
        cell.fill = fill_ink
        cell.alignment = center
        cell.border = thin


def paint_input(cell):
    cell.fill = fill_yellow
    cell.font = font_base
    cell.alignment = wrap
    cell.border = thin


def paint_text(cell, bold=False):
    cell.font = font_bold if bold else font_base
    cell.alignment = wrap
    cell.border = thin
    cell.fill = fill_white


def list_validation(ws, cell_range, options, name):
    formula = '"' + ",".join(options) + '"'
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    dv.error = "请从列表中选择"
    dv.errorTitle = name
    dv.prompt = "选择" + name
    dv.promptTitle = name
    dv.add(cell_range)
    ws.add_data_validation(dv)
    return dv


def status_colors(ws, cell_range, good, bad, mid=None):
    ws.conditional_formatting.add(
        cell_range,
        CellIsRule(operator="equal", formula=[f'"{good}"'], fill=fill_green),
    )
    ws.conditional_formatting.add(
        cell_range,
        CellIsRule(operator="equal", formula=[f'"{bad}"'], fill=fill_rose),
    )
    if mid:
        ws.conditional_formatting.add(
            cell_range,
            CellIsRule(operator="equal", formula=[f'"{mid}"'], fill=fill_gold),
        )


def setup_page(ws, landscape=True, fit_height=0):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_height
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.horizontalCentered = True
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.6
    ws.page_margins.bottom = 0.5
    ws.oddHeader.left.text = "2026康博会专题分论坛 · 内部执行台账"
    ws.oddFooter.left.text = "黄底单元格需要填写。结论和结余由公式生成。"
    ws.oddFooter.right.text = "第 &P 页"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.sheet_properties.tabColor = TEAL
    ws.page_setup.scale = 100


def title(ws, text, subtitle, last_col):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    c1 = ws.cell(1, 1, text)
    c1.font = font_title
    c1.alignment = Alignment(vertical="center", horizontal="left")
    c2 = ws.cell(2, 1, subtitle)
    c2.font = font_small
    c2.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 32
    for col in range(1, last_col + 1):
        ws.cell(1, col).fill = fill_cream
        ws.cell(2, col).fill = fill_cream


def widths(ws, pairs):
    for col, width in pairs.items():
        ws.column_dimensions[col].width = width


def build_overview(wb):
    ws = wb.active
    ws.title = "总览"
    title(
        ws,
        "执行管理台账",
        "内部使用。黄底格子填写，其余数字自动计算。当前结论直接读取“决策检查”。默认全部未满足，避免在授权完成前误判为可以发布。",
        6,
    )
    ws["A4"] = "当前结论"
    ws["A4"].font = font_bold
    ws.merge_cells("A5:F5")
    ws["A5"] = "='决策检查'!A18"
    ws["A5"].font = Font(name="微软雅黑", size=14, bold=True, color=INK)
    ws["A5"].alignment = Alignment(wrap_text=True, vertical="center")
    ws["A5"].fill = fill_gold
    for col in range(1, 7):
        ws.cell(5, col).fill = fill_gold
        ws.cell(5, col).border = thin
    ws.row_dimensions[5].height = 48

    labels = [
        (7, 1, "嘉宾已确认（含视频）"),
        (7, 3, "邀请已发出或待回复"),
        (7, 5, "已婉拒或改邀"),
        (9, 1, "候选目标合计（元）"),
        (9, 3, "签约金额（元）"),
        (9, 5, "已回款（元）"),
        (11, 1, "预算上限（元）"),
        (11, 3, "实际支出（元）"),
        (11, 5, "预算结余（元）"),
        (13, 1, "任务已完成"),
        (13, 3, "任务总数"),
        (13, 5, "任务完成率"),
    ]
    value_map = {
        (8, 1): '=COUNTIF(嘉宾邀约!I5:I40,"已确认")+COUNTIF(嘉宾邀约!I5:I40,"视频替代")',
        (8, 3): '=COUNTIF(嘉宾邀约!I5:I40,"已发邀请")+COUNTIF(嘉宾邀约!I5:I40,"待回复")',
        (8, 5): '=COUNTIF(嘉宾邀约!I5:I40,"已婉拒")+COUNTIF(嘉宾邀约!I5:I40,"改邀替补")',
        (10, 1): "=SUM(招商台账!F5:F40)",
        (10, 3): "=SUM(招商台账!K5:K40)",
        (10, 5): "=SUM(招商台账!L5:L40)",
        (12, 1): "=SUM(预算控制!C5:C12)",
        (12, 3): "=SUM(预算控制!E5:E12)",
        (12, 5): "=A12-C12",
        (14, 1): '=COUNTIF(任务倒排!H5:H40,"已完成")',
        (14, 3): '=COUNTA(任务倒排!E5:E40)',
        (14, 5): '=IF(C14=0,"—",A14/C14)',
    }
    for row, col, text in labels:
        cell = ws.cell(row, col, text)
        cell.font = font_small
        cell.alignment = wrap
    for (row, col), formula in value_map.items():
        cell = ws.cell(row, col, formula)
        cell.font = font_kpi
        cell.alignment = Alignment(vertical="center", horizontal="left")
        cell.fill = fill_mist
        cell.border = thin
    for row in (8, 10, 12, 14):
        ws.row_dimensions[row].height = 28
    for coord in ("A10", "C10", "E10", "A12", "C12", "E12"):
        ws[coord].number_format = '#,##0'
    ws["E14"].number_format = "0%"

    ws["A16"] = "怎么用这本台账"
    ws["A16"].font = font_bold
    lines = [
        "1. 先改“决策检查”。七项底线里只要还有“未满足”，结论就是不以独立总承办身份接手。",
        "2. 11月14日前，主旨渠道、傅利叶或替代主线、场地技术条件三项没有通过，就不要大规模发布。",
        "3. 11月20日前，现金或等额资源不足3万元，就把非必要支出停掉。预算表里的茶歇和设备运输按这个开关执行。",
        "4. 嘉宾、招商、任务的黄底单元格由责任人每天更新。总览数字不要手改。",
        "5. “商务条件”和“对外口径”只在内部使用。赞助商只接收对外赞助方案。",
        "6. 胡继刚的圆桌角色已纳入方案，海报仍要等联合承办授权完成后再使用校名和会标。",
    ]
    for i, line in enumerate(lines):
        ws.merge_cells(start_row=17 + i, start_column=1, end_row=17 + i, end_column=6)
        cell = ws.cell(17 + i, 1, line)
        cell.font = font_base
        cell.alignment = wrap
        ws.row_dimensions[17 + i].height = 22
    widths(ws, {"A": 28, "B": 22, "C": 28, "D": 18, "E": 24, "F": 18})
    setup_page(ws, landscape=True, fit_height=1)
    ws.sheet_properties.tabColor = INK
    ws.auto_filter.ref = None
    ws.freeze_panes = "A4"
    ws.oddFooter.left.text = "内部执行台账 · 不对外发送"
    return ws


def build_decision(wb):
    ws = wb.create_sheet("决策检查")
    title(ws, "承接条件检查", "请只修改 D 列。结论写在第18行，并显示在总览。判定日期：11月14日看第12—14行，11月20日看第15行。", 6)
    headers = ["编号", "闸门", "条件", "判定", "证据或待办", "未满足时的处理"]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    apply_header_row(ws, 4, 6)
    rows = [
        (1, "承接底线", "策划权已在书面协议中明确", "协议写明主题、议程、嘉宾、合作单位、主持和赞助方案由项目团队主导", "不独立承办"),
        (2, "承接底线", "执行权已经明确", "现场执行、嘉宾联络和对接服务有指定统筹人", "不独立承办"),
        (3, "承接底线", "可销售权益已与大会已售权益核对", "冠名、展示、发言和采购对接清单确认没有冲突", "不独立承办"),
        (4, "承接底线", "收款主体、开票和分配已书面约定", "账户、结算时点、成本扣除和分配方式", "不独立承办"),
        (5, "承接底线", "场地、基础设备、现场人员和官方宣传已落实", "会议室、音视频、人员和宣传渠道有书面清单", "不独立承办，或先缩小规模"),
        (6, "承接底线", "不要求个人无上限垫资，也不承担大会招商缺口", "风险划分已写进协议", "不垫付，不独立承办"),
        (7, "承接底线", "未授权的媒体、领导和主办身份承诺可以删除", "旧方案中的央媒、电视直播和名单共享已删除", "不对外发布"),
        (8, "11月14日", "主旨嘉宾的邀约渠道已经明确", "复旦老龄研究专家的联系人和发出记录", "暂停对外发布"),
        (9, "11月14日", "傅利叶合作意向或替代技术主线已经明确", "企业回复，或已经确定的备用企业", "暂停对外发布"),
        (10, "11月14日", "场地技术条件已经核验", "电源、通道、基础音视频和演示限制", "暂停对外发布"),
        (11, "11月20日", "现金或等额可执行资源不少于3万元", "到账凭证，或正式合同载明的等额资源", "压缩为低成本对接，停止非必要支出"),
    ]
    for i, (num, gate, cond, evidence, action) in enumerate(rows):
        r = 5 + i
        values = [num, gate, cond, "未满足", evidence, action]
        for c, value in enumerate(values, 1):
            cell = ws.cell(r, c, value)
            paint_text(cell, bold=(c == 1))
            if c == 4:
                paint_input(cell)
        ws.row_dimensions[r].height = 32
    list_validation(ws, "D5:D15", ["已满足", "部分满足", "未满足", "不适用"], "判定")
    status_colors(ws, "D5:D15", "已满足", "未满足", "部分满足")
    ws.merge_cells("A17:F17")
    ws["A17"] = "当前结论（自动生成，不要手改）"
    ws["A17"].font = font_bold
    ws.merge_cells("A18:F18")
    ws["A18"] = (
        '=IF(COUNTIF(D5:D15,"")>0,"还有条件未判定，暂不对外发布",'
        'IF(COUNTIF(D5:D11,"未满足")>0,"不建议以独立总承办身份接手",'
        'IF(COUNTIF(D5:D11,"部分满足")>0,"底线尚未全部满足，可先改为有偿单项服务，不要大规模发布",'
        'IF(COUNTIF(D12:D14,"未满足")+COUNTIF(D12:D14,"部分满足")>0,"底线已齐，但11月14日闸门未过，暂停对外发布",'
        'IF(OR(D15="未满足",D15="部分满足"),"11月20日资金或资源不足：压缩为低成本对接，停止非必要支出",'
        '"条件满足，按现行方案推进")))))'
    )
    ws["A18"].font = Font(name="微软雅黑", size=14, bold=True, color=RED)
    ws["A18"].alignment = Alignment(wrap_text=True, vertical="center")
    ws["A18"].fill = fill_gold
    for col in range(1, 7):
        ws.cell(18, col).fill = fill_gold
        ws.cell(18, col).border = thin
    ws.row_dimensions[18].height = 42
    ws.conditional_formatting.add(
        "A18:F18",
        FormulaRule(formula=['ISNUMBER(SEARCH("条件满足",A18))'], fill=fill_green),
    )
    widths(ws, {"A": 8, "B": 16, "C": 48, "D": 14, "E": 52, "F": 36})
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:F15"
    ws.auto_filter.ref = "A4:F15"
    setup_page(ws)
    ws.print_title_rows = "1:4"
    ws.sheet_properties.tabColor = GOLD
    return ws


def build_tasks(wb):
    ws = wb.create_sheet("任务倒排")
    title(ws, "倒排任务", "今天是2026年11月11日。前五项标为进行中。状态只从列表选择。验收标准没有达到时，不要改成已完成。", 8)
    headers = ["编号", "开始", "结束", "阶段", "任务", "验收标准", "责任方", "状态"]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    apply_header_row(ws, 4, 8)
    tasks = [
        (date(2026, 11, 11), date(2026, 11, 13), "授权", "取得星智耀与组委会对承办范围的书面说明", "书面回复包含授权范围和执行主体", "星智耀 + 胡继刚", "进行中"),
        (date(2026, 11, 11), date(2026, 11, 13), "授权", "确认杨浦科企联联合承办授权", "单位同意署名，并指定对接人", "杨浦科企联", "进行中"),
        (date(2026, 11, 11), date(2026, 11, 13), "授权", "确认复旦住房政策研究中心联合承办授权", "书面同意使用机构名称；校名和标识另行批准", "住策中心", "进行中"),
        (date(2026, 11, 11), date(2026, 11, 13), "财务", "确定收款主体、开票和结算方式", "签约主体、账户和分配原则写成文字", "商务与财务", "进行中"),
        (date(2026, 11, 11), date(2026, 11, 13), "场地", "核验会议室档期、电源、通道和基础音视频", "组委会给出可用条件清单", "星智耀 + 组委会", "进行中"),
        (date(2026, 11, 12), date(2026, 11, 14), "嘉宾", "发出复旦老龄研究院专家邀请", "彭希哲、吴玉韶或代表，至少形成一条有效邀请记录", "学术组", "未开始"),
        (date(2026, 11, 12), date(2026, 11, 14), "嘉宾", "联系傅利叶业务负责人", "确认人选、演示形式，以及能否使用企业名称", "产业组", "未开始"),
        (date(2026, 11, 12), date(2026, 11, 14), "场景", "通过养老协会锁定场景机构", "至少两家机构有参会或试点讨论意愿", "协会渠道", "未开始"),
        (date(2026, 11, 13), date(2026, 11, 16), "招商", "完成与大会已售赞助的冲突核对", "书面结论：可售、不可售或需修改", "招商组", "未开始"),
        (date(2026, 11, 13), date(2026, 11, 16), "招商", "向10—15家匹配企业发送已批准的方案", "发送记录包含企业、档位和日期", "招商组", "未开始"),
        (date(2026, 11, 16), date(2026, 11, 20), "财务", "完成合同并落实至少3万元", "到账凭证，或合同载明的等额资源", "商务与财务", "未开始"),
        (date(2026, 11, 18), date(2026, 11, 23), "内容", "锁定主旨题目、圆桌名单和主持稿", "主持人确认档期；圆桌不超过6人", "总策划 + 主持人", "未开始"),
        (date(2026, 11, 18), date(2026, 11, 23), "邀约", "按席位结构完成定向邀约", "买方、科技企业和投资机构分别有名单", "会务组 + 协会", "未开始"),
        (date(2026, 11, 21), date(2026, 11, 25), "物料", "海报和现场物料", "仅在授权完成且预算允许后发布", "执行组", "未开始"),
        (date(2026, 11, 23), date(2026, 11, 26), "现场", "确认演示方案、座席、安全和备份", "企业与场馆书面确认，或改为视频", "执行组 + 技术组", "未开始"),
        (date(2026, 11, 26), date(2026, 11, 27), "现场", "联调和嘉宾接待确认", "流程走一遍，应急联系人到人", "现场总控", "未开始"),
        (date(2026, 11, 28), date(2026, 11, 28), "举办", "论坛执行并记录每组对接", "对接记录包含双方、需求和下一步", "现场总控", "未开始"),
        (date(2026, 11, 29), date(2026, 12, 5), "收口", "影像、权益核销、结算和合作台账", "台账在7个工作日内可交接", "项目负责人 + 财务", "未开始"),
    ]
    for i, row in enumerate(tasks, 1):
        r = 4 + i
        ws.cell(r, 1, i)
        ws.cell(r, 2, row[0])
        ws.cell(r, 3, row[1])
        for c in range(1, 9):
            paint_text(ws.cell(r, c))
        for c, value in enumerate(row[2:], 4):
            ws.cell(r, c, value)
            paint_text(ws.cell(r, c))
        paint_input(ws.cell(r, 8))
        ws.cell(r, 2).number_format = "YYYY-MM-DD"
        ws.cell(r, 3).number_format = "YYYY-MM-DD"
        ws.row_dimensions[r].height = 30
    list_validation(ws, "H5:H40", ["未开始", "进行中", "已完成", "阻塞", "取消"], "任务状态")
    status_colors(ws, "H5:H40", "已完成", "阻塞", "进行中")
    widths(ws, {"A": 8, "B": 14, "C": 14, "D": 12, "E": 46, "F": 40, "G": 24, "H": 12})
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:H22"
    setup_page(ws)
    ws.print_title_rows = "1:4"
    return ws


def build_guests(wb):
    ws = wb.create_sheet("嘉宾邀约")
    title(ws, "嘉宾邀约台账", "姓名在确认前都是拟邀。I列为状态。胡继刚已纳入方案，但校名和会标仍要等授权。新增嘉宾从第20行往下填。", 12)
    headers = ["编号", "梯队", "优先级", "姓名", "单位及身份", "建议角色", "建议主题或讨论点", "邀约责任", "状态", "计划联系日", "是否上台", "备注"]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    apply_header_row(ws, 4, 12)
    people = [
        ("第一梯队", "P0", "彭希哲", "复旦大学老龄研究院院长（发出前核对）", "学术主旨", "“十五五”老龄社会与银发经济", "学术组", "未联系", date(2026, 11, 12), "待定", "档期困难可改派代表"),
        ("第一梯队", "P0", "吴玉韶", "复旦大学老龄研究院副院长、教授（发出前核对）", "政策与产业主旨", "“十五五”养老服务与银发科技", "学术组", "未联系", date(2026, 11, 12), "待定", "学术主讲的重点邀约对象"),
        ("第一梯队", "P0", "顾捷", "傅利叶智能创始人、CEO", "产业主旨", "具身智能进入康复机构、养老院和家庭", "产业组", "未联系", date(2026, 11, 12), "待定", "本人不出席则改业务负责人"),
        ("第一梯队", "P0", "胡继刚", "住房政策研究中心秘书长；副教授级高级工程师；杨浦科企联执行会长", "核心圆桌、资源统筹", "从适老居住到智慧康养", "筹备组", "已确认", date(2026, 11, 11), "圆桌", "不另做主旨。海报待联合承办授权"),
        ("第二梯队", "P1", "朱震", "科技园及杨浦科企联相关负责人，身份待核", "成果转化", "杨浦科技创新与成果转化", "杨浦科企联", "未联系", date(2026, 11, 13), "待定", "身份未核实前不印制"),
        ("第二梯队", "P1", "赵博誉", "上海星智耀科技有限公司相关负责人", "项目对接", "成果转化机制和项目筛选", "星智耀", "未联系", date(2026, 11, 13), "短发言或对接", "压缩独立发言"),
        ("第二梯队", "P1", "徐洁平", "璞跃中国", "圆桌主持", "控场，不推介本机构", "产业组", "未联系", date(2026, 11, 13), "主持", "须确认档期和控场能力"),
        ("第二梯队", "P0", "待协会确认", "上海市养老服务和老龄产业协会负责人", "行业致辞、需求组织", "养老服务需求与机构邀请", "协会渠道", "未联系", date(2026, 11, 12), "致辞", "通过协会邀请机构，避免重复联系"),
        ("第二梯队", "P0", "待推荐", "养老运营机构负责人（第1家）", "场景需求", "康复、安全、护理和成本", "协会渠道", "未联系", date(2026, 11, 13), "需求发布", "先作为买方，不先谈赞助"),
        ("第二梯队", "P0", "待推荐", "养老或康复机构负责人（第2家）", "场景需求", "试点条件与使用流程", "协会渠道", "未联系", date(2026, 11, 14), "圆桌或对接", "有采购或试点讨论权"),
        ("第二梯队", "P1", "待确认", "康复医学或老年医学专家", "临床评价", "机器人在康复场景的适用边界", "学术组", "未联系", date(2026, 11, 16), "圆桌候选", "复旦相关附属医院按需要增补"),
        ("第二梯队", "P1", "待确认", "投资机构合伙人或投资总监", "资本观察", "规模化、付费方和回报", "招商组", "未联系", date(2026, 11, 14), "资本观察", "不邀请只负责宣传的代表"),
        ("补充", "P2", "封进", "复旦老龄研究相关方向（任职待核）", "定向交流", "银发经济与养老金融", "学术组", "未联系", date(2026, 11, 16), "否", "不占主旨，核对后再决定是否出场"),
        ("补充", "P2", "罗力", "复旦老龄研究相关方向（任职待核）", "定向交流", "健康养老与医养融合", "学术组", "未联系", date(2026, 11, 16), "否", "不占主旨，核对后再决定是否出场"),
        ("补充", "P1", "待确认", "工商银行上海市分行业务负责人", "金融支持", "养老金融或科创金融能否支持本场", "招商组", "未联系", date(2026, 11, 14), "待定", "先做大会金融权益的排他审查"),
    ]
    for i, row in enumerate(people, 1):
        r = 4 + i
        ws.cell(r, 1, i)
        for c, value in enumerate(row, 2):
            ws.cell(r, c, value)
        for c in range(1, 13):
            paint_text(ws.cell(r, c))
        paint_input(ws.cell(r, 9))
        ws.cell(r, 10).number_format = "YYYY-MM-DD"
        ws.row_dimensions[r].height = 30
    for r in range(20, 31):
        ws.cell(r, 1, r - 4)
        for c in range(1, 13):
            paint_text(ws.cell(r, c))
        paint_input(ws.cell(r, 9))
        ws.row_dimensions[r].height = 22
    list_validation(ws, "I5:I40", ["未联系", "已发邀请", "待回复", "已确认", "视频替代", "已婉拒", "改邀替补"], "邀约状态")
    list_validation(ws, "C5:C40", ["P0", "P1", "P2"], "优先级")
    status_colors(ws, "I5:I40", "已确认", "已婉拒", "待回复")
    widths(ws, {"A": 8, "B": 12, "C": 10, "D": 14, "E": 42, "F": 22, "G": 32, "H": 14, "I": 12, "J": 14, "K": 14, "L": 32})
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:L30"
    setup_page(ws)
    ws.print_title_rows = "1:4"
    return ws


def build_sales(wb):
    ws = wb.create_sheet("招商台账")
    title(ws, "招商台账", "候选方向不是已合作企业。目标金额是内部测算，签约金额和已回款只填实际数字。未确认的企业不得写入联合冠名。", 13)
    headers = ["编号", "优先级", "对象类型", "候选机构或方向", "建议档位", "目标金额", "合作切入", "责任人", "状态", "权益冲突审查", "签约金额", "已回款", "备注"]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    apply_header_row(ws, 4, 13)
    rows = [
        ("P0", "康复机器人", "傅利叶智能", "内容合作，现金另议", 20000, "先谈主讲、演示和名称授权，再谈赞助", "产业组", "未接触", "未审查", 0, 0, "未确认前不联合冠名，不把办会押在这笔款"),
        ("P1", "机器人产业链", "康复机器人及零部件企业", "产业创新或生态", 20000, "客户对接、展示和采购场景", "招商组", "未接触", "未审查", 0, 0, "方向候选，确定企业后再改名称"),
        ("P1", "数字健康", "智慧养老或数字健康企业", "生态或产业创新", 10000, "解决方案展示和技术合作", "招商组", "未接触", "未审查", 0, 0, "方向候选"),
        ("P0", "大会金融主办", "工商银行上海市分行", "资源协调", 0, "先确认现有主办权益能否支持分论坛", "招商组", "未接触", "未审查", 0, 0, "不平行销售冲突的金融冠名"),
        ("P1", "其他金融机构", "有养老金融或科技金融业务的银行", "专项生态", 10000, "确认排他边界后再接触", "招商组", "未接触", "未审查", 0, 0, "含既有渠道。不写成已经合作"),
        ("P1", "养老运营", "有试点或采购需求的运营集团", "产业创新", 20000, "先邀请作为买方，再看是否赞助", "协会渠道", "未接触", "未审查", 0, 0, "首次接触不要求高额赞助"),
        ("P2", "保险", "保险集团养老或健康业务", "产业创新或生态", 20000, "支付机制和试点资金", "招商组", "未接触", "未审查", 0, 0, "方向候选"),
        ("P1", "产业投资", "有决策权的基金合伙人或投资总监", "专项生态", 10000, "项目评价和融资对接", "招商组", "未接触", "未审查", 0, 0, "不邀请品牌宣传岗顶替"),
        ("P2", "服务置换", "摄影、物料或会务服务商", "实物或服务", 0, "置换现金支出", "执行组", "未接触", "不适用", 0, 0, "交付标准先写清"),
    ]
    for i, row in enumerate(rows, 1):
        r = 4 + i
        ws.cell(r, 1, i)
        for c, value in enumerate(row, 2):
            ws.cell(r, c, value)
        for c in range(1, 14):
            paint_text(ws.cell(r, c))
        for c in (9, 10, 11, 12):
            paint_input(ws.cell(r, c))
        for c in (6, 11, 12):
            ws.cell(r, c).number_format = '#,##0'
        ws.row_dimensions[r].height = 32
    for r in range(14, 25):
        ws.cell(r, 1, r - 4)
        for c in range(1, 14):
            paint_text(ws.cell(r, c))
        for c in (9, 10, 11, 12):
            paint_input(ws.cell(r, c))
        for c in (6, 11, 12):
            ws.cell(r, c).number_format = '#,##0'
            if ws.cell(r, c).value is None and c in (11, 12):
                ws.cell(r, c, 0)
    list_validation(ws, "I5:I40", ["未接触", "方案已发", "洽谈中", "意向", "合同审核", "已签约", "已回款", "放弃"], "招商状态")
    list_validation(ws, "J5:J40", ["未审查", "无冲突", "有冲突待处理", "不适用"], "权益审查")
    status_colors(ws, "I5:I40", "已回款", "放弃", "洽谈中")
    ws.conditional_formatting.add("J5:J40", CellIsRule(operator="equal", formula=['"有冲突待处理"'], fill=fill_rose))
    ws.conditional_formatting.add("J5:J40", CellIsRule(operator="equal", formula=['"无冲突"'], fill=fill_green))
    widths(ws, {"A": 8, "B": 10, "C": 16, "D": 32, "E": 20, "F": 12, "G": 36, "H": 12, "I": 12, "J": 16, "K": 12, "L": 12, "M": 36})
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:M24"
    setup_page(ws)
    ws.print_title_rows = "1:4"
    return ws


def build_budget(wb):
    ws = wb.create_sheet("预算控制")
    title(ws, "预算控制", "上限合计23,000元。假定场地和基础音视频免费，且没有嘉宾出场费。实际支出填黄底。结余自动计算。个人不垫付。", 7)
    headers = ["编号", "科目", "上限（元）", "是否可削减", "实际支出（元）", "结余（元）", "什么情况下才支出"]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    apply_header_row(ws, 4, 7)
    items = [
        ("主视觉、海报和电子物料", 1500, "可削减", "机构名称和标识授权完成之后"),
        ("摄影和必要视频", 3500, "可削减", "至少保留照片；精剪视回款而定"),
        ("现场执行人员", 3000, "视大会支持", "大会已提供人员则实际支出为0"),
        ("嘉宾接待和交通补助", 3000, "可削减", "不承诺出场费，本地嘉宾优先"),
        ("茶歇和饮水", 2000, "可削减", "11月20日资金未落实则取消"),
        ("设备运输和安全保障", 3500, "视回款", "企业和场馆书面确认演示后再发生"),
        ("补充技术设备", 3000, "视回款", "基础音视频不足时才使用"),
        ("预备费", 3500, "预备", "不得挪作个人垫资"),
    ]
    for i, (name, limit, flexibility, rule) in enumerate(items, 1):
        r = 4 + i
        ws.cell(r, 1, i)
        ws.cell(r, 2, name)
        ws.cell(r, 3, limit)
        ws.cell(r, 4, flexibility)
        ws.cell(r, 5, 0)
        ws.cell(r, 6, f"=C{r}-E{r}")
        ws.cell(r, 7, rule)
        for c in range(1, 8):
            paint_text(ws.cell(r, c))
        paint_input(ws.cell(r, 5))
        ws.cell(r, 3).number_format = '#,##0'
        ws.cell(r, 5).number_format = '#,##0'
        ws.cell(r, 6).number_format = '#,##0'
        ws.row_dimensions[r].height = 28
    ws.cell(13, 2, "合计")
    ws.cell(13, 3, "=SUM(C5:C12)")
    ws.cell(13, 5, "=SUM(E5:E12)")
    ws.cell(13, 6, "=C13-E13")
    for c in range(1, 8):
        cell = ws.cell(13, c)
        cell.font = Font(name="微软雅黑", size=10, bold=True, color=WHITE)
        cell.fill = fill_teal
        cell.alignment = wrap
        cell.border = thin
    for c in (3, 5, 6):
        ws.cell(13, c).number_format = '#,##0'
    ws["A15"] = "收入对照（内部，不是到账）"
    ws["A15"].font = font_bold
    ws["A16"] = "最低目标"
    ws["B16"] = 30000
    ws["A17"] = "基准目标"
    ws["B17"] = 80000
    ws["A18"] = "理想目标"
    ws["B18"] = 110000
    ws["A19"] = "基准组合"
    ws["B19"] = "首席50,000 + 产业创新20,000 + 生态10,000。不含把傅利叶现金赞助当作前提。"
    ws.merge_cells("B19:G19")
    for r in range(16, 20):
        ws.cell(r, 1).font = font_base
        ws.cell(r, 2).font = font_bold
        ws.cell(r, 2).alignment = wrap
    for r in range(16, 19):
        ws.cell(r, 2).number_format = '"¥"#,##0'
    ws["A21"] = "若8万元全额回款且支出等于上限，账面差额为57,000元。仍须扣除税费、招商服务费和书面约定的分配，不能视为可拿走的利润。场地或设备若要自付，停止使用本表并重算。"
    ws.merge_cells("A21:G21")
    ws["A21"].font = Font(name="微软雅黑", size=10, color=RED)
    ws["A21"].alignment = wrap
    ws.row_dimensions[21].height = 36
    widths(ws, {"A": 14, "B": 28, "C": 14, "D": 16, "E": 16, "F": 14, "G": 42})
    ws.freeze_panes = "A5"
    setup_page(ws, fit_height=1)
    return ws


def build_agenda(wb):
    ws = wb.create_sheet("议程执行")
    title(ws, "议程执行表", "2026年11月28日 13:00—16:30。状态改为“改为备份”时，在备份方案里写清替代内容。", 8)
    headers = ["编号", "时间", "环节", "内容", "拟邀人员", "责任", "状态", "备份方案"]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    apply_header_row(ws, 4, 8)
    rows = [
        ("13:00—13:05", "开场", "介绍背景、组织和目标", "独立主持人", "内容组", "未锁定", "主持人口径提前写好"),
        ("13:05—13:15", "致辞", "行业协会；两家联合承办统一介绍", "协会代表；授权后的承办单位", "会务组", "未锁定", "授权未完成则只介绍大会"),
        ("13:15—13:35", "学术主旨", "银发经济与主动健康", "复旦老龄研究专家1位", "学术组", "未锁定", "研究院代表或书面发言要点"),
        ("13:35—14:00", "产业主旨", "具身智能与康复养老", "傅利叶创始人或业务负责人", "产业组", "未锁定", "同类企业，或授权视频"),
        ("14:00—14:20", "场景需求", "适老居住和机构真实问题", "住策中心 + 机构代表", "场景组", "未锁定", "只保留已确认的需求"),
        ("14:20—14:40", "技术展示", "设备演示或2个筛选项目", "企业技术负责人", "技术组", "未锁定", "视频、实物和讲解"),
        ("14:40—14:50", "资本观察", "商业模式与回报", "有决策权的投资或金融代表", "招商组", "未锁定", "并入圆桌，不单设"),
        ("14:50—15:00", "茶歇", "交流和参观", "会务", "执行组", "未锁定", "资金不足则取消茶歇"),
        ("15:00—15:45", "产业圆桌", "谁应用、谁投资、如何商业化", "胡继刚等，含主持不超过6人", "主持人", "未锁定", "人员不齐就减席，不凑数"),
        ("15:45—16:15", "供需对接", "2—3个项目和机构需求", "会前配对的双方", "产业对接组", "未锁定", "没有配对表就不称有效洽谈"),
        ("16:15—16:25", "成果发布", "已落实的意向和下一步", "项目负责人", "总协调", "未锁定", "不安排空泛签约"),
        ("16:25—16:30", "闭幕", "总结和合影", "主持人", "现场总控", "未锁定", "肖像未授权不公开发布"),
    ]
    for i, row in enumerate(rows, 1):
        r = 4 + i
        ws.cell(r, 1, i)
        for c, value in enumerate(row, 2):
            ws.cell(r, c, value)
        for c in range(1, 9):
            paint_text(ws.cell(r, c))
        paint_input(ws.cell(r, 7))
        paint_input(ws.cell(r, 8))
        ws.row_dimensions[r].height = 30
    list_validation(ws, "G5:G20", ["未锁定", "已锁定", "改为备份", "取消"], "议程状态")
    status_colors(ws, "G5:G20", "已锁定", "取消", "改为备份")
    widths(ws, {"A": 8, "B": 16, "C": 14, "D": 32, "E": 32, "F": 14, "G": 12, "H": 28})
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:H16"
    setup_page(ws)
    ws.print_title_rows = "1:4"
    return ws


def build_round(wb):
    ws = wb.create_sheet("圆桌设计")
    title(ws, "核心圆桌", "15:00—15:45。主题：机器人进入养老机构。主持人独立控场。胡继刚不重复技术报告和政策报告。", 5)
    headers = ["时间", "环节", "要回答的问题", "胡继刚的讨论位置", "其他嘉宾"]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    apply_header_row(ws, 4, 5)
    rows = [
        ("15:00—15:05", "开场", "养老机器人已经具备哪些可落地功能，商业化缺少什么？", "听清问题，不先做自我介绍式发言", "主持人介绍五位嘉宾"),
        ("15:05—15:20", "技术与需求", "高校科研、适老居住和科技企业如何一起找到可持续市场？", "硬件之外，还要看空间适配、使用者、物业和养老服务、采购与付费主体", "傅利叶讲产品瓶颈；老龄专家讲需求；机构讲真实流程"),
        ("15:20—15:35", "模式与资本", "谁付费？如何评估效益？谁承担早期试点风险？", "以真实需求做小规模试点，形成效果、成本和推广数据，再谈融资", "投资或金融讨论租赁、服务付费和试点资金"),
        ("15:35—15:43", "落地动作", "每人只提一项下一步", "提出一项可以执行的场景验证或资源对接", "试点、联合研究、产品合作或融资洽谈"),
        ("15:43—15:45", "收束", "四条线如何进入会后台账", "确认自己负责跟进的事项", "主持人引导分组对接"),
    ]
    for i, row in enumerate(rows):
        r = 5 + i
        for c, value in enumerate(row, 1):
            ws.cell(r, c, value)
            paint_text(ws.cell(r, c))
        ws.row_dimensions[r].height = 48
    ws["A11"] = "海报四行"
    ws["A11"].font = font_bold
    lines = [
        "胡继刚",
        "复旦大学住房政策研究中心秘书长",
        "副教授级高级工程师",
        "上海市杨浦区科技企业联合会执行会长",
        "圆桌分享方向：《从适老居住到智慧康养：高校研究、产业场景与资本协同创新》",
    ]
    for i, line in enumerate(lines):
        ws.cell(12 + i, 1, line)
        ws.merge_cells(start_row=12 + i, start_column=1, end_row=12 + i, end_column=5)
        ws.cell(12 + i, 1).font = font_bold if i == 0 else font_base
    widths(ws, {"A": 16, "B": 16, "C": 42, "D": 46, "E": 42})
    ws.freeze_panes = "A5"
    setup_page(ws, fit_height=1)
    return ws


def build_results(wb):
    ws = wb.create_sheet("参会成果")
    title(ws, "参会结构与成果", "上表是考核，下表是席位指引。实际列请手填。考核不对外承诺成交。一家机构的多名代表，在机构数里只计一次。", 4)
    headers = ["考核指标", "目标", "实际", "说明"]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    apply_header_row(ws, 4, 4)
    kpis = [
        ("参会规模", "100—120人", "到场签到人数"),
        ("核心演讲及圆桌", "8—10人", "实际发言人数，不含未发言来宾"),
        ("科技企业", "15—20家", "按机构计数"),
        ("养老及康复运营机构", "10—15家", "按机构计数，优先有试点讨论权的单位"),
        ("金融和投资机构", "5—8家", "按机构计数"),
        ("场景或合作需求", "至少10项", "写入会后台账才计数"),
        ("路演或展示项目", "2—3个", "会前完成技术成熟度和需求表"),
        ("有效商务洽谈", "10—15组", "会前匹配，现场有记录"),
        ("进入后续讨论的项目", "2—3项", "洽谈、技术验证或试点讨论"),
        ("合作台账", "会后7个工作日内", "负责人、下一步、日期齐全"),
    ]
    for i, (name, target, note) in enumerate(kpis):
        r = 5 + i
        ws.cell(r, 1, name)
        ws.cell(r, 2, target)
        ws.cell(r, 4, note)
        for c in range(1, 5):
            paint_text(ws.cell(r, c))
        paint_input(ws.cell(r, 3))
        ws.row_dimensions[r].height = 24
    ws["A16"] = "120人席位指引（不是已确认名单）"
    ws["A16"].font = font_bold
    headers2 = ["类别", "席位", "已确认人数", "备注"]
    for i, h in enumerate(headers2, 1):
        ws.cell(17, i, h)
    apply_header_row(ws, 17, 4)
    seats = [
        ("养老运营、医疗康复和场景方", "30—35人", "买方质量优先"),
        ("康养科技企业及创新团队", "25—30人", "对应15—20家"),
        ("高校、科研及成果转化", "15—20人", "不含未授权挂名"),
        ("银行、保险及投资机构", "10—15人", "对应5—8家"),
        ("协会、联合承办和平台", "10—15人", "工作人员不计入专业观众"),
        ("合作伙伴、媒体及其他", "剩余至120人", "媒体须已落实"),
    ]
    for i, (name, seat, note) in enumerate(seats):
        r = 18 + i
        ws.cell(r, 1, name)
        ws.cell(r, 2, seat)
        ws.cell(r, 4, note)
        for c in range(1, 5):
            paint_text(ws.cell(r, c))
        paint_input(ws.cell(r, 3))
    widths(ws, {"A": 36, "B": 22, "C": 18, "D": 42})
    ws.freeze_panes = "A5"
    setup_page(ws, fit_height=1)
    return ws


def build_followup(wb):
    ws = wb.create_sheet("会后台账")
    title(ws, "会后合作台账", "11月28日现场记录，11月29日至12月5日补全。没有负责人和日期的行，不计入有效成果。", 10)
    headers = ["编号", "需求方", "供给方", "需求摘要", "合作意向", "负责人", "下一步", "节点日期", "状态", "备注"]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    apply_header_row(ws, 4, 10)
    for i in range(1, 16):
        r = 4 + i
        ws.cell(r, 1, i)
        for c in range(1, 11):
            paint_text(ws.cell(r, c))
            paint_input(ws.cell(r, c))
        paint_text(ws.cell(r, 1))
        ws.cell(r, 1, i)
        ws.cell(r, 8).number_format = "YYYY-MM-DD"
        ws.row_dimensions[r].height = 24
    list_validation(ws, "I5:I30", ["未开始", "跟进中", "已完成", "暂停"], "跟进状态")
    status_colors(ws, "I5:I30", "已完成", "暂停", "跟进中")
    widths(ws, {"A": 8, "B": 18, "C": 18, "D": 28, "E": 18, "F": 14, "G": 24, "H": 14, "I": 12, "J": 20})
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:J19"
    setup_page(ws)
    ws.print_title_rows = "1:4"
    return ws


def build_org(wb):
    ws = wb.create_sheet("组织架构")
    title(ws, "组织架构", "拟议名单。新增联合承办和学术、产业支持，取得书面授权后才能对外使用。", 4)
    headers = ["层级", "单位", "职责", "授权状态"]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    apply_header_row(ws, 4, 4)
    rows = [
        ("大会主办", "上海现代服务业联合会", "大会品牌与行业组织资源", "沿用官方名单"),
        ("大会主办", "上海市养老服务和老龄产业协会", "养老机构、采购和场景", "沿用官方名单"),
        ("大会主办", "上海丰之米会展有限公司", "场地、合同、会务和大会运营", "沿用官方名单"),
        ("大会主办", "中国工商银行上海市分行", "大会金融主办。分论坛先做权益核对", "沿用官方名单"),
        ("分论坛主办", "上海星智耀科技有限公司", "组委会对接、原有合作和成果转化", "待组委会确认"),
        ("联合承办", "上海市杨浦区科技企业联合会", "AI与机器人企业、产业资本、产业对接", "待授权"),
        ("联合承办", "复旦大学住房政策研究中心", "适老居住、社区康养、学术联络", "待授权"),
        ("学术支持", "复旦大学老龄研究院", "老龄政策、银发经济、主动健康", "待单独邀约"),
        ("产业创新", "傅利叶智能", "康复机器人、具身智能、技术展示", "待确认"),
        ("产业创新", "璞跃中国", "创新生态；徐洁平为主持候选", "待确认"),
        ("技术转移", "国家技术转移东部中心", "技术转移支持，不列指导单位", "待确认"),
        ("场景合作", "养老协会相关专委会及机构", "需求征集和专业观众", "待确认"),
    ]
    for i, row in enumerate(rows):
        r = 5 + i
        for c, value in enumerate(row, 1):
            ws.cell(r, c, value)
            paint_text(ws.cell(r, c))
        paint_input(ws.cell(r, 4))
        ws.row_dimensions[r].height = 26
    list_validation(ws, "D5:D20", ["沿用官方名单", "待组委会确认", "待授权", "待单独邀约", "待确认", "已授权", "不使用"], "授权状态")
    status_colors(ws, "D5:D20", "已授权", "不使用", "待授权")
    widths(ws, {"A": 16, "B": 36, "C": 42, "D": 18})
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:D16"
    setup_page(ws, fit_height=1)
    return ws


def build_adjust(wb):
    ws = wb.create_sheet("单位调整")
    title(ws, "原有参与方调整", "大会层级的主办单位不由分论坛单方面取消。本表只决定谁出现在11月28日这场论坛。", 4)
    headers = ["原有单位或人员", "决定", "执行方式", "对外是否出现"]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    apply_header_row(ws, 4, 4)
    rows = [
        ("上海星智耀科技有限公司", "保留", "继续成果转化和组委会协调，不另建招商决策链", "待组委会确认后使用"),
        ("国家技术转移东部中心", "改为拟邀支持", "可以参加技术转移，未授权不列指导单位", "授权前不出现"),
        ("上海辰山国家植物园", "退出主议程", "没有明确转化成果就不占核心发言", "本场不出现"),
        ("上海视觉艺术学院", "定向再邀", "有适老化或人机交互项目时再安排", "本场核心材料不出现"),
        ("长三角示范区新发展建设有限公司", "降为场景候选", "有具体养老科技项目再纳入", "没有项目就不出现"),
        ("文旅及相关学院", "不进核心名单", "留给大会其他银发文旅专题", "本场不出现"),
        ("宠物公益及一般文化艺术单位", "不列入", "维持大会原合作，不重复挂名", "本场不出现"),
        ("王红霞 / 甘薯科研", "取消上台", "没有产品、人群、合作方和合规证据，不进主议程", "本场不出现"),
        ("朱震", "建议保留", "发言围绕科创和成果转化；身份先核实", "核实前不印制"),
        ("赵博誉", "建议保留", "转化机制和项目对接，压缩独立发言", "确认出席后使用"),
        ("徐洁平", "主持候选", "圆桌主持，确认档期和职责", "确认后使用"),
        ("朱翊元及开发建设嘉宾", "不占核心圆桌", "没有即时场景项目就不安排", "默认不出现"),
        ("复旦大学老龄研究院", "重点争取", "与住房政策研究中心分开邀约", "授权前只说拟邀"),
        ("傅利叶智能", "重点新增候选", "先内容后赞助", "确认前不联合冠名"),
    ]
    for i, row in enumerate(rows):
        r = 5 + i
        for c, value in enumerate(row, 1):
            ws.cell(r, c, value)
            paint_text(ws.cell(r, c))
        ws.row_dimensions[r].height = 28
    widths(ws, {"A": 36, "B": 20, "C": 48, "D": 24})
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:D18"
    setup_page(ws, fit_height=1)
    return ws


def build_rights(wb):
    ws = wb.create_sheet("赞助权益")
    title(ws, "赞助权益交付清单", "拟议权益。组委会批准并写入合同后，才能把交付状态改为已写入合同。未批准的权益不要对客户口头答应。", 6)
    headers = ["档位", "权益", "是否须组委会批准", "交付责任", "交付状态", "备注"]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    apply_header_row(ws, 4, 6)
    rights = [
        ("首席 / 5万元 / 1家", "物料展示首席产业合作伙伴名称和标识", "是", "执行组", "未承诺"),
        ("首席 / 5万元 / 1家", "审核后10—15分钟主题报告", "是", "内容组", "未承诺"),
        ("首席 / 5万元 / 1家", "1个圆桌候选席位", "是", "主持人", "未承诺"),
        ("首席 / 5万元 / 1家", "符合安全条件的展示或演示", "是", "技术组", "未承诺"),
        ("首席 / 5万元 / 1家", "与康养机构的定向对接", "是", "产业对接组", "未承诺"),
        ("首席 / 5万元 / 1家", "6个参会名额", "是", "会务组", "未承诺"),
        ("首席 / 5万元 / 1家", "新闻稿和总结体现实际合作", "是", "内容组", "未承诺"),
        ("首席 / 5万元 / 1家", "经许可使用照片、视频和总结", "是", "执行组", "未承诺"),
        ("产业创新 / 2万元 / 2家", "品牌展示", "是", "执行组", "未承诺"),
        ("产业创新 / 2万元 / 2家", "筛选后的展示或路演", "是", "内容组", "未承诺"),
        ("产业创新 / 2万元 / 2家", "圆桌候选或交流机会", "是", "主持人", "未承诺"),
        ("产业创新 / 2万元 / 2家", "定向推荐", "是", "产业对接组", "未承诺"),
        ("产业创新 / 2万元 / 2家", "4个参会名额", "是", "会务组", "未承诺"),
        ("生态 / 1万元 / 2家", "物料和电子屏", "是", "执行组", "未承诺"),
        ("生态 / 1万元 / 2家", "供需对接", "是", "产业对接组", "未承诺"),
        ("生态 / 1万元 / 2家", "2个参会名额", "是", "会务组", "未承诺"),
        ("生态 / 1万元 / 2家", "新闻稿或鸣谢列名", "是", "内容组", "未承诺"),
        ("实物或服务", "设备、摄影、物料、茶歇或技术支持", "是", "执行组", "未承诺"),
        ("所有档位", "不承诺央媒、电视、领导会见、成交和融资", "不适用", "全体", "不适用"),
        ("所有档位", "个人名单仅在合法授权范围内提供", "不适用", "全体", "不适用"),
    ]
    for i, row in enumerate(rights):
        r = 5 + i
        for c, value in enumerate(row, 1):
            ws.cell(r, c, value)
            paint_text(ws.cell(r, c))
        paint_input(ws.cell(r, 5))
        ws.row_dimensions[r].height = 24
    list_validation(ws, "E5:E30", ["未承诺", "已写入合同", "已交付", "不适用"], "交付状态")
    status_colors(ws, "E5:E30", "已交付", "未承诺", "已写入合同")
    widths(ws, {"A": 26, "B": 42, "C": 20, "D": 16, "E": 16, "F": 28})
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:F24"
    setup_page(ws)
    ws.print_title_rows = "1:4"
    return ws


def build_risk(wb):
    ws = wb.create_sheet("风险应急")
    title(ws, "风险与应急", "状态改为已触发时，当天执行预案，并在任务表把相关任务标为阻塞或取消。", 5)
    headers = ["风险", "出现时的信号", "预案", "责任", "状态"]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    apply_header_row(ws, 4, 5)
    rows = [
        ("核心嘉宾缺席", "会前48小时仍未确认行程", "同领域替补或视频。不用未授权肖像和讲稿", "学术组 / 产业组", "未发生"),
        ("机器人无法到场", "企业或场馆未书面确认演示", "改为授权视频、案例和远程讲解", "技术组", "未发生"),
        ("赞助不足", "11月20日未到3万元或等额资源", "保留主旨、圆桌和对接；取消茶歇、非必要物料和运输", "财务", "未发生"),
        ("单位未授权", "联合承办或高校没有书面同意", "物料和新闻稿不使用该名称和标识", "总协调", "未发生"),
        ("签约不成熟", "会前没有真实合作文本", "改为需求发布和意向登记", "总协调", "未发生"),
        ("宣传超出产品能力", "讲稿含疗效、融资保证或未获批功能", "改讲稿；现场不演示未经测试的接触动作", "内容组", "未发生"),
        ("收款账户不清", "赞助要进入其他主体且没有结算协议", "停止以承办身份招商", "财务", "未发生"),
        ("要求个人垫资", "发生支出前没有预算责任主体", "不垫付，重新谈规模或退出独立承办", "总协调", "未发生"),
    ]
    for i, row in enumerate(rows):
        r = 5 + i
        for c, value in enumerate(row, 1):
            ws.cell(r, c, value)
            paint_text(ws.cell(r, c))
        paint_input(ws.cell(r, 5))
        ws.row_dimensions[r].height = 32
    list_validation(ws, "E5:E20", ["未发生", "观察中", "已触发", "已关闭"], "风险状态")
    status_colors(ws, "E5:E20", "已关闭", "已触发", "观察中")
    widths(ws, {"A": 22, "B": 36, "C": 48, "D": 20, "E": 12})
    ws.freeze_panes = "A5"
    setup_page(ws, fit_height=1)
    return ws


def build_language(wb):
    ws = wb.create_sheet("对外口径")
    title(ws, "对外可以说什么", "赞助电话、微信和面对面沟通都按这张表。拿不准时，说“以正式合同和授权为准”。", 3)
    headers = ["场景", "可以说", "不要说"]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    apply_header_row(ws, 4, 3)
    rows = [
        ("时间和地点", "11月28日下午，上海展览中心，会议室以组委会确认为准。", "未书面确认的厅室已经定稿。"),
        ("单位", "拟增补、拟邀，并说明以授权为准。", "某高校或某企业已经联合主办、联合冠名。"),
        ("嘉宾", "正在邀请老龄研究专家和康复机器人企业负责人。", "把拟邀姓名印成确认出席。"),
        ("胡继刚", "授权完成后使用三行头衔，身份是圆桌嘉宾和资源统筹。", "在授权前把校名和会标印上海报。"),
        ("赞助价格", "拟议档位5万、2万、1万和实物支持，以批准后的合同为准。", "旧方案的10万、5万、2万已经生效。"),
        ("成果", "组织需求交流、项目展示和会后跟进。", "保证采购、融资、政府审批或领导会见。"),
        ("媒体", "已经落实的媒体渠道可以写入合同。", "承诺央媒报道或电视台专访。"),
        ("名单", "在法律和授权范围内交换合作信息。", "会后提供全套个人联系方式。"),
        ("财务", "只在内部讨论分成、垫资和2.3万元预算。", "对客户谈内部分成比例或个人垫资。"),
        ("安全", "演示须通过企业和场馆确认。", "安排未测试的人形机器人接触老年体验者。"),
    ]
    for i, row in enumerate(rows):
        r = 5 + i
        for c, value in enumerate(row, 1):
            ws.cell(r, c, value)
            paint_text(ws.cell(r, c))
        ws.cell(r, 2).fill = fill_green
        ws.cell(r, 3).fill = fill_rose
        ws.row_dimensions[r].height = 32
    widths(ws, {"A": 16, "B": 62, "C": 48})
    ws.freeze_panes = "A5"
    setup_page(ws, fit_height=1)
    ws.sheet_properties.tabColor = RED
    return ws


def build_commercial(wb):
    ws = wb.create_sheet("商务条件")
    title(ws, "内部商务条件", "本表不进入赞助方案，不提交给赞助商，也不作为对组委会的报价单。它只是筹备组的谈判起点。", 4)
    ws.sheet_properties.tabColor = RED
    headers = ["项目", "谈判起点", "用法", "禁止"]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    apply_header_row(ws, 4, 4)
    rows = [
        ("优先模式", "固定策划执行费 + 招商绩效提成", "两种模式择一，避免重复计酬", "不要同时拿提成和净收益分成"),
        ("固定策划执行费", "15,000—30,000元", "覆盖策划和统筹，不覆盖无限垫资", "不要把固定费说成行业统一标准"),
        ("招商服务费", "本团队引入且已回款金额的20%—30%", "只计算实际引入并完成回款的部分", "不要按意向金额或他人带来的赞助计提"),
        ("替代模式", "可分配净收益的50%—60%", "没有固定费时使用。先扣确认过的成本、税费", "不要在毛收入上直接分成"),
        ("垫资", "不接受个人无上限垫资", "支出发生前要有预算责任主体", "不要先垫2.3万元再谈结算"),
        ("收款", "授权法人主体签约、开票、收款", "独立核算，保留台账", "不要进入个人账户"),
        ("高校和协会", "遵守各自财务制度", "挂名不能代替财务授权", "不要用研究机构名义直接收取赞助"),
        ("退路", "改为有偿单项服务", "策划、嘉宾推荐、企业邀约或招商可单项计价", "不要在没有三权时承诺整场结果"),
    ]
    for i, row in enumerate(rows):
        r = 5 + i
        for c, value in enumerate(row, 1):
            ws.cell(r, c, value)
            paint_text(ws.cell(r, c))
        ws.row_dimensions[r].height = 32
    ws["A14"] = "对组委会只需要谈清楚授权、独立核算、书面结算和不个人垫资。具体比例留到商务谈判，不写进对外赞助方案。"
    ws.merge_cells("A14:D14")
    ws["A14"].font = Font(name="微软雅黑", size=11, bold=True, color=RED)
    ws["A14"].alignment = wrap
    ws["A14"].fill = fill_rose
    for col in range(1, 5):
        ws.cell(14, col).fill = fill_rose
    ws.row_dimensions[14].height = 36
    widths(ws, {"A": 20, "B": 42, "C": 40, "D": 36})
    ws.freeze_panes = "A5"
    setup_page(ws, fit_height=1)
    return ws


def build():
    wb = Workbook()
    wb.calculation.calcMode = "auto"
    wb.properties.title = "2026康博会专题分论坛执行管理台账"
    wb.properties.creator = "专题分论坛筹备组"
    wb.properties.subject = "内部执行"
    build_overview(wb)
    build_decision(wb)
    build_tasks(wb)
    build_guests(wb)
    build_sales(wb)
    build_budget(wb)
    build_agenda(wb)
    build_round(wb)
    build_results(wb)
    build_followup(wb)
    build_org(wb)
    build_adjust(wb)
    build_rights(wb)
    build_risk(wb)
    build_language(wb)
    build_commercial(wb)
    order = [
        "总览",
        "决策检查",
        "任务倒排",
        "嘉宾邀约",
        "招商台账",
        "预算控制",
        "议程执行",
        "圆桌设计",
        "参会成果",
        "会后台账",
        "组织架构",
        "单位调整",
        "赞助权益",
        "风险应急",
        "对外口径",
        "商务条件",
    ]
    for i, name in enumerate(order):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    return OUT


if __name__ == "__main__":
    print(build())
