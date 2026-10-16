# -*- coding: utf-8 -*-
"""生成综合方案配套工作簿。

用法：
    python3 scripts/build_xlsx.py [输出路径.xlsx]
"""

import os
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.page import PageMargins

import proposal_data as D

NAVY = "0E2438"
GOLD = "C4A15A"
ICE = "F3F6F9"
SAND = "FBF6EC"
WHITE = "FFFFFF"
INK = "1C2733"
MUTED = "5C6B78"
GREEN = "E7F2EA"
YELLOW = "FFF4CC"
AMBER = "F8E8C8"
TEAL = "1F6A7A"
RED_SOFT = "F8E6E6"

thin = Border(
    left=Side(style="thin", color="E1E7EE"),
    right=Side(style="thin", color="E1E7EE"),
    top=Side(style="thin", color="E1E7EE"),
    bottom=Side(style="thin", color="E1E7EE"),
)
thick_bottom = Border(
    left=Side(style="thin", color="E1E7EE"),
    right=Side(style="thin", color="E1E7EE"),
    top=Side(style="thin", color="E1E7EE"),
    bottom=Side(style="medium", color=NAVY),
)

font_name = "微软雅黑"
title_font = Font(name=font_name, size=18, bold=True, color=NAVY)
section_font = Font(name=font_name, size=13, bold=True, color=NAVY)
header_font = Font(name=font_name, size=11, bold=True, color=WHITE)
body_font = Font(name=font_name, size=10, color=INK)
bold_font = Font(name=font_name, size=10, bold=True, color=INK)
muted_font = Font(name=font_name, size=9, italic=True, color=MUTED)
link_font = Font(name=font_name, size=11, color="1F4E79", underline="single")
input_font = Font(name=font_name, size=10, color=INK)
formula_font = Font(name=font_name, size=10, bold=True, color=NAVY)

fill_navy = PatternFill("solid", fgColor=NAVY)
fill_gold = PatternFill("solid", fgColor=GOLD)
fill_ice = PatternFill("solid", fgColor=ICE)
fill_sand = PatternFill("solid", fgColor=SAND)
fill_white = PatternFill("solid", fgColor=WHITE)
fill_green = PatternFill("solid", fgColor=GREEN)
fill_yellow = PatternFill("solid", fgColor=YELLOW)
fill_amber = PatternFill("solid", fgColor=AMBER)
fill_header_teal = PatternFill("solid", fgColor=TEAL)

wrap = Alignment(wrap_text=True, vertical="center", horizontal="left")
wrap_center = Alignment(wrap_text=True, vertical="center", horizontal="center")
left_align = Alignment(vertical="center", horizontal="left", wrap_text=True)


def apply_page(ws, title):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.horizontalCentered = True
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.7, bottom=0.6, header=0.3, footer=0.3)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddHeader.left.text = "潮起潮落 TIDELINE × 北外滩尚9·一滴水"
    ws.oddHeader.right.text = title
    ws.oddFooter.left.text = "内部研讨稿  ·  计划目标不等于已兑现结果"
    ws.oddFooter.right.text = "第 &P 页"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.sheet_properties.tabColor = NAVY
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = None
    ws.page_setup.scale = None
    ws.print_title_rows = "1:4"
    ws.sheet_format.defaultRowHeight = 18
    ws.oddHeader.center.text = ""


def write_banner(ws, title, note, last_col):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    cell = ws.cell(1, 1, title)
    cell.font = title_font
    cell.alignment = Alignment(vertical="center", horizontal="left")
    ws.row_dimensions[1].height = 28
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    note_cell = ws.cell(2, 1, note)
    note_cell.font = muted_font
    note_cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 32
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=last_col)
    ws.cell(3, 1, "").fill = fill_gold
    ws.row_dimensions[3].height = 6
    for col in range(1, last_col + 1):
        ws.cell(1, col).fill = fill_white
        ws.cell(2, col).fill = fill_white
        ws.cell(3, col).fill = fill_gold


def write_headers(ws, headers, row=4):
    for col, text in enumerate(headers, start=1):
        cell = ws.cell(row, col, text)
        cell.font = header_font
        cell.fill = fill_navy
        cell.alignment = wrap_center
        cell.border = thin
    ws.row_dimensions[row].height = 24
    ws.auto_filter.ref = f"A{row}:{get_column_letter(len(headers))}{row}"
    ws.freeze_panes = f"A{row + 1}"


def style_body_cell(cell, fill, center=False, bold=False):
    cell.font = bold_font if bold else body_font
    cell.fill = fill
    cell.alignment = wrap_center if center else wrap
    cell.border = thin


def paint_rows(ws, start, end, ncol, center_cols=None, bold_cols=None):
    center_cols = center_cols or set()
    bold_cols = bold_cols or set()
    for r in range(start, end + 1):
        fill = fill_ice if (r - start) % 2 == 0 else fill_white
        ws.row_dimensions[r].height = 32
        for c in range(1, ncol + 1):
            cell = ws.cell(r, c)
            style_body_cell(cell, fill, center=c in center_cols, bold=c in bold_cols)


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def add_table(ws, name, note, headers, rows, widths, center_cols=None, bold_cols=None, row_height=32):
    last_col = len(headers)
    write_banner(ws, name, note, last_col)
    write_headers(ws, headers)
    for r, row in enumerate(rows, start=5):
        for c, value in enumerate(row, start=1):
            ws.cell(r, c, value)
        ws.row_dimensions[r].height = row_height
    if rows:
        paint_rows(ws, 5, 4 + len(rows), last_col, center_cols, bold_cols)
        ws.auto_filter.ref = f"A4:{get_column_letter(last_col)}{4 + len(rows)}"
    set_widths(ws, widths)
    apply_page(ws, name)
    ws.auto_filter.ref = f"A4:{get_column_letter(last_col)}{4 + max(len(rows), 1)}"
    ws.freeze_panes = "A5"
    return 5 + len(rows)


def highlight_status(ws, col_idx, start, end):
    col = get_column_letter(col_idx)
    rng = f"{col}{start}:{col}{end}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'ISNUMBER(SEARCH("已过",{col}{start}))'], fill=fill_amber))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'ISNUMBER(SEARCH("进行中",{col}{start}))'], fill=fill_sand))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'ISNUMBER(SEARCH("高",{col}{start}))'], fill=fill_amber))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{col}{start}="有"'], fill=fill_green))


def build_cover(wb):
    ws = wb.active
    ws.title = "00-阅读说明"
    apply_page(ws, "阅读说明")
    ws.freeze_panes = "A8"
    ws.print_title_rows = "1:3"
    set_widths(ws, [28, 78, 36])
    write_banner(
        ws,
        "潮起潮落 TIDELINE × 北外滩尚9·一滴水",
        "青年文化消费场景创新暨长期战略合作综合方案  ·  配套工作簿  ·  对照日期 " + D.AS_OF_CN,
        3,
    )
    write_headers(ws, ["使用方式", "说明", "对应工作表"])
    rows = [
        ("给经营者", "先看商业价值、长期产品和合作模式。这三张表回答：为什么参与、以后做什么、钱怎么算。", "02、03、12"),
        ("给商务谈判", "赞助权益、场地方收入、合作模式和收益测算一起看。黄色单元格才是待填写金额。", "04、11、12、17"),
        ("给执行对接", "空间、日程、人员、安全保障、场地配合和推进时间，作为踏勘和排期清单。", "06、07、08、09、10、13"),
        ("给签约审核", "职责、评估指标和八项待确认事项。未闭合前，不把计划数字写成承诺。", "14、15、16"),
        ("口径", "客流、曝光、合作资源和经营收益都是计划目标或合作设想。赞助金额是原稿建议档位。惠迪案例来自公司自述。", "全表"),
        ("与 PPT 的关系", "PPT 用于汇报。本工作簿保留可筛选的明细，并提供一滴水本次收益的空白测算。", "全部"),
    ]
    for i, row in enumerate(rows):
        for c, value in enumerate(row, start=1):
            ws.cell(5 + i, c, value)
        ws.row_dimensions[5 + i].height = 36
    paint_rows(ws, 5, 10, 3, bold_cols={1})
    ws.auto_filter.ref = "A4:C10"

    ws.cell(12, 1, "工作表目录").font = section_font
    headers = ["序号", "工作表", "内容"]
    for c, text in enumerate(headers, start=1):
        cell = ws.cell(13, c, text)
        cell.font = header_font
        cell.fill = fill_navy
        cell.alignment = wrap_center
        cell.border = thin
    catalog = [
        ("00", "00-阅读说明", "本表。口径、目录和填写规则。"),
        ("01", "01-活动信息", "首期活动的基本事实、八项内容和六个亮点。"),
        ("02", "02-商业价值", "八项商业价值、婚庆表达、获客动作和品牌资产。"),
        ("03", "03-长期产品", "八个长期方向及婚礼、社交、摄影等产品要素。"),
        ("04", "04-赞助权益", "白银、黄金、钻石权益矩阵，以及置换和场地自有权益。"),
        ("05", "05-传播计划", "渠道目标、补充渠道、节奏、话题和核验清单。"),
        ("06", "06-空间需求", "楼层用途、功能面积和动线。合计 1,780㎡ 为需求加总。"),
        ("07", "07-活动日程", "活动日流程，以及布展撤展需要重排的原因。"),
        ("08", "08-团队案例", "惠迪自述规模、业务、案例和其他项目。"),
        ("09", "09-执行保障", "人员配置、安全保障和布撤展确认项。"),
        ("10", "10-场地配合", "需要一滴水确认的十四项配合。"),
        ("11", "11-收入与成本", "主办收入、场地方收入和成本科目。金额不在本表填写。"),
        ("12", "12-合作模式", "三种模式的适用条件和首期建议。"),
        ("13", "13-推进时间", "原稿节点相对 2026-10-16 的状态。"),
        ("14", "14-职责分工", "主办、场地方和共同职责。"),
        ("15", "15-评估指标", "客流、经营、传播、转化和交付物。"),
        ("16", "16-待确认", "签约前八项缺口、口径和责任方。"),
        ("17", "17-收益测算", "黄色单元格供商务填写。绿色为公式。未来订单不计入本次结果。"),
        ("18", "18-主题互动", "三大主题区色彩和八项互动。"),
    ]
    for i, row in enumerate(catalog):
        for c, value in enumerate(row, start=1):
            cell = ws.cell(14 + i, c, value)
        ws.row_dimensions[14 + i].height = 22
    paint_rows(ws, 14, 13 + len(catalog), 3, center_cols={1}, bold_cols={2})
    for i, row in enumerate(catalog):
        cell = ws.cell(14 + i, 2)
        cell.hyperlink = f"#'{row[1]}'!A1"
        cell.font = Font(name=font_name, size=10, bold=True, color="1F4E79", underline="single")

    ws.cell(35, 1, "填写规则").font = section_font
    rules = [
        "只有「17-收益测算」中的黄色单元格需要填写金额。其他工作表是方案内容，不把计划数字改成已经签约的结果。",
        "绿色单元格是公式。直接经营结果 = 本次直接收入 − 本次直接成本。机会成本和未来订单单独列示，不自动计入该结果。",
        "面积合计 1,780㎡ 是各功能需求相加，不是场地须提供的净面积，也不能用来判断安全容量。",
        "传播合计约 70,000+ 存在用户重叠，不能解释为 7 万名独立消费者。抖音未计入该合计。",
        "赞助 5,000 / 15,000 / 30,000 元是建议档位。高一档包含前一档权益。白银摊位未约定尺寸。",
    ]
    for i, text in enumerate(rules):
        ws.merge_cells(start_row=36 + i, start_column=1, end_row=36 + i, end_column=3)
        cell = ws.cell(36 + i, 1, text)
        cell.font = body_font
        cell.alignment = wrap
        cell.fill = fill_sand if i % 2 == 0 else fill_white
        ws.row_dimensions[36 + i].height = 28


def build_event(wb):
    ws = wb.create_sheet("01-活动信息")
    rows = [(k, v, "基本信息") for k, v in D.EVENT_FACTS]
    rows += [(num, f"{label}。{note}", "关键数字") for num, label, note in D.EVENT_METRICS]
    rows += [(title, body, "活动内容") for _, title, body in D.CONTENTS]
    rows += [(title, body, "活动亮点") for title, body in D.HIGHLIGHTS]
    rows += [(title, body, "核心客群") for title, body in D.AUDIENCES]
    rows += [(item, "到场之后可能形成的后续客户类型", "未来客户") for item in D.FUTURE_CLIENTS]
    add_table(
        ws, "首期活动信息",
        "日期、人数和主办主体尚未闭合。人数以核定容量为准。停留时间是原稿预测，需用实际数据验证。",
        ["项目", "内容", "类别"],
        rows,
        [28, 88, 16],
        center_cols={3},
        bold_cols={1},
        row_height=30,
    )


def build_value(wb):
    ws = wb.create_sheet("02-商业价值")
    rows = [(num, title, body, "八项价值") for num, title, body in D.VALUES]
    rows += [(f"{i:02d}", item, "新一代婚礼可能出现的表达", "婚庆市场") for i, item in enumerate(D.WEDDING_EXAMPLES, 1)]
    rows += [(a, b, c, "非核心档期产品") for a, b, c in D.OFFPEAK]
    rows += [(num, title, body, "现场转化动作") for num, title, body in D.ACQUISITION]
    rows += [(item, "优先由一滴水提供，收入分配写入合同", "餐饮",) for item in D.CATERING]
    rows += [(item, "取得授权后可用于宣传、推广和招商", "品牌资产") for item in D.BRAND_ASSETS]
    add_table(
        ws, "对一滴水的商业价值",
        "本表对应汇报重点。高端婚礼定位保持不变，主题婚礼是在传统产品之外增加的细分市场。",
        ["编号或项目", "内容", "说明", "类别"],
        rows,
        [28, 42, 48, 18],
        center_cols={4},
        bold_cols={1},
    )


def build_longterm(wb):
    ws = wb.create_sheet("03-长期产品")
    rows = [(num, title, body, "方向总览") for num, title, body in D.LONGTERM]
    rows += [(level, title, body, "婚礼产品") for level, title, body in D.WEDDING_LAYERS]
    rows += [(item, "预约制、小规模、主题化，优先非核心时段", "", "主题社交") for item in D.SOCIAL_FORMS]
    rows += [(item, "团队负责嘉宾、粉丝、票务和传播；一滴水提供场地及配套", "", "创作者活动") for item in D.SIGNING_FORMS]
    rows += [(item, "可与专业摄影机构做成不同档次的场地套餐", "", "摄影") for item in D.PHOTO_LINES]
    add_table(
        ws, "长期业务方向",
        "这些是试点之后才产品化的方向。使用特定商业 IP、角色形象或虚拟角色庆典前，先取得授权。",
        ["编号或名称", "内容", "经营说明", "类别"],
        rows,
        [28, 55, 40, 16],
        center_cols={4},
        bold_cols={1},
        row_height=28,
    )


def build_sponsor(wb):
    ws = wb.create_sheet("04-赞助权益")
    rows = [row for row in D.SPONSOR_MATRIX]
    end = add_table(
        ws, "赞助权益对照",
        "建议金额：白银 5,000 元，黄金 15,000 元，钻石 30,000 元。并非已确认订单。高一档包含前一档。结案报告拟含照片、触达、阅读量、互动和效果总结。",
        ["权益", "白银 5,000元", "黄金 15,000元", "钻石 30,000元", "备注"],
        rows,
        [32, 18, 18, 18, 42],
        center_cols={2, 3, 4},
        bold_cols={1},
        row_height=26,
    )
    highlight_status(ws, 2, 5, end - 1)
    highlight_status(ws, 3, 5, end - 1)
    highlight_status(ws, 4, 5, end - 1)

    start = end + 2
    ws.merge_cells(start_row=start, start_column=1, end_row=start, end_column=5)
    ws.cell(start, 1, "实物、资源置换与一滴水自有权益").font = section_font
    headers = ["类型", "内容", "", "", "类别"]
    for c, text in enumerate(["类型", "内容", "—", "—", "类别"], start=1):
        cell = ws.cell(start + 1, c, text if c < 3 or c == 5 else "")
        cell.font = header_font
        cell.fill = fill_header_teal
        cell.alignment = wrap_center
        cell.border = thin
    extra = [(a, b, "置换合作") for a, b in D.BARTER] + [(a, b, "一滴水权益") for a, b in D.VENUE_RIGHTS]
    for i, (a, b, cat) in enumerate(extra):
        r = start + 2 + i
        ws.cell(r, 1, a)
        ws.cell(r, 2, b)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.cell(r, 5, cat)
        ws.row_dimensions[r].height = 30
    paint_rows(ws, start + 2, start + 1 + len(extra), 5, center_cols={5}, bold_cols={1})


def build_media(wb):
    ws = wb.create_sheet("05-传播计划")
    rows = [(a, b, c, "计入原稿合计") for a, b, c in D.CHANNELS]
    rows += [(a, b, c, "不计入 7 万合计") for a, b, c in D.EXTRA_CHANNELS]
    rows += [(a, b, c, "宣传节奏") for a, b, c in D.RHYTHM]
    rows += [(item, "签约前由活动团队提供", "", "资源核验") for item in D.KOL_CHECKS]
    rows += [(item, "线下传播动作", "", "线下") for item in D.OFFLINE]
    rows += [(tag, "建议统一使用的话题", "", "话题") for tag in D.HASHTAGS]
    add_table(
        ws, "传播计划与目标",
        "各渠道数字是计划覆盖或曝光。合计约 70,000+ 有用户重叠，不是独立消费者。粉丝 10 万至 1,000 万、百大 UP主、卡游、杭州银行均待确认清单。",
        ["项目", "目标或时间", "内容", "类别"],
        rows,
        [28, 28, 62, 18],
        center_cols={4},
        bold_cols={1},
        row_height=26,
    )


def build_space(wb):
    ws = wb.create_sheet("06-空间需求")
    rows = [(a, b, "", "楼层用途") for a, b in D.FLOORS]
    for name, place, area in D.AREAS:
        rows.append((name, place, area, "面积需求"))
    rows.append(("合计", "需求简单加总，不是净面积，也不能判断承载", D.area_total(), "面积需求"))
    rows.append(("动线", "G层签到 → 安检 → 手环 → 摊位 → 主题互动 → 一楼宴会 → 三楼体验", "", "动线"))
    rows.append(("散场", "分批离场、单向流动、合理分流。须经场地和安全人员审核。", "", "动线"))
    end = add_table(
        ws, "空间需求",
        "使用范围以最新平面图、使用权限、消防条件和核定容量为准。三楼需同时确认电梯、楼梯、消防通道和露台安全。",
        ["区域或事项", "位置或说明", "计划面积（㎡）", "类别"],
        rows,
        [28, 70, 20, 16],
        center_cols={3, 4},
        bold_cols={1},
    )
    for r in range(5, end):
        cell = ws.cell(r, 3)
        if isinstance(cell.value, int):
            cell.number_format = '#,##0'
            cell.alignment = wrap_center
    total_row = 4 + len(rows)
    for c in range(1, 5):
        ws.cell(total_row - 2, c).fill = fill_sand
        ws.cell(total_row - 2, c).font = bold_font


def build_rundown(wb):
    ws = wb.create_sheet("07-活动日程")
    rows = [(a, b, c, "活动日") for a, b, c in D.RUNDOWN]
    rows += [
        ("布展", "原稿拟 11月26日至28日", "与 28日、29日的活动日期冲突并存，确定活动日后再排", "布撤展"),
        ("撤展", "活动结束后", "开放货梯、装卸区，完成清洁和场地恢复", "布撤展"),
        ("占用周期", "不按一天活动计算", "多楼层、多空间和布景装置，要核算完整占用", "布撤展"),
        ("人数", "原稿另列 VIP 约 200 人、普通观众分批入场", "必须按票务、空间和安全容量调整", "人数"),
    ]
    add_table(
        ws, "活动日程",
        "时间表保留原稿，部分环节并行，例如签售与夜间灯光启动重叠。正式执行时按动线重排。",
        ["时间或事项", "内容", "场地方配合或说明", "类别"],
        rows,
        [24, 42, 50, 14],
        center_cols={4},
        bold_cols={1},
        row_height=28,
    )


def build_team(wb):
    ws = wb.create_sheet("08-团队案例")
    rows = [(num, label, "公司自述，签约前核验", "自述规模") for num, label in D.COMPANY_STATS]
    rows += [(item, "公司介绍所列业务", "", "业务范围") for item in D.BUSINESSES]
    rows += [(name, "", "八个职能模块之一", "职能") for name in D.FUNCTIONS]
    rows += [(a, b, "", "管理阶段") for a, b in D.PHASE_MGMT]
    rows += [(a, b, c, "代表项目") for a, b, c in D.CASES]
    rows += [(item, "需确认是主办、承办、联合执行还是单项供应商", "", "其他案例") for item in D.OTHER_CASES]
    rows.append(("商业综合体", D.MALLS, "能否用于本项目需另行确认", "资源"))
    add_table(
        ws, "北京惠迪：自述规模与案例",
        "公司成立于 2016 年。案例人数和传播数据均来自公司材料，正式合作前核验项目角色与可核查资料。ACF 1.0 约 20 万人气、ACF 2.0 单话题 140 万以上曝光，均以可核查资料为准。",
        ["项目或指标", "材料所列内容", "核验说明", "类别"],
        rows,
        [42, 42, 42, 14],
        center_cols={4},
        bold_cols={1},
        row_height=24,
    )


def build_safety(wb):
    ws = wb.create_sheet("09-执行保障")
    rows = [(a, b, c, "人员配置") for a, b, c in D.STAFF]
    rows.append(("招商组", "摊位招募、赞助对接及合同管理", "来自品牌赞助方案，人数未单列", "人员配置"))
    rows += [(a, b, "", "安全保障") for a, b in D.SAFETY]
    rows += [(item, "布展前与场地方确认", "", "布撤展") for item in D.BUILD_ITEMS]
    add_table(
        ws, "人员与现场保障",
        "人员数量是策划阶段的初步安排。资质、专业安保和医疗配置，按场地条件、人数和管理要求确定。噪声投诉的 30 分钟、垃圾两小时清运，均为原稿提出的内部目标。",
        ["事项", "内容", "补充", "类别"],
        rows,
        [24, 55, 40, 14],
        center_cols={4},
        bold_cols={1},
    )


def build_tasks(wb):
    ws = wb.create_sheet("10-场地配合")
    rows = [(a, b, c, "待踏勘确认") for a, b, c in D.VENUE_TASKS]
    rows.append(("15", "现有设备", "确认灯光、音响、餐饮及其他配套，哪些可用现有设备，哪些需另购或租赁", "待踏勘确认"))
    add_table(
        ws, "场地方配合清单",
        "本清单用于商务谈判和现场踏勘，不表示一滴水已经承诺提供全部条件。",
        ["序号", "事项", "需要确认的内容", "状态"],
        rows,
        [10, 20, 78, 16],
        center_cols={1, 4},
        bold_cols={2},
        row_height=30,
    )


def build_money(wb):
    ws = wb.create_sheet("11-收入与成本")
    rows = [(a, b, "主办方收入，不与场地方混同") for a, b in D.HOST_INCOME]
    rows += [(a, b, "场地方收入") for a, b in D.VENUE_INCOME]
    rows += [(item, "活动主要支出科目，金额在《17-收益测算》按承担主体填写", "成本科目") for item in D.COSTS]
    rows.append(("备用金", "原稿建议预留总预算的 5%—10%", "成本科目"))
    add_table(
        ws, "收入与成本科目",
        "本表只列科目和口径。一滴水本次金额请填写在「17-收益测算」。后续订单不并入本次利润。",
        ["科目", "口径", "归属"],
        rows,
        [32, 70, 28],
        bold_cols={1},
        row_height=28,
    )


def build_modes(wb):
    ws = wb.create_sheet("12-合作模式")
    rows = []
    for name, title, what, strength, risk in D.MODES:
        rows.append((name, title, what, strength, risk, "备选" if name != "模式二" else "首期建议研究"))
    rows.append((
        "建议",
        "优惠场租＋餐饮最低消费＋场地品牌传播＋后续业务合作权",
        "具体条件按商业原则协商",
        "首期以验证商业模式为核心",
        "缺少成本和收益测算时，不承诺长期免费使用场地",
        "建议",
    ))
    end = add_table(
        ws, "三种合作方式",
        "资源置换可以研究，但要同时核算一滴水的真实场地成本和机会成本，并写明权益交付标准、后续义务，以及取消后未完成权益如何处理。",
        ["模式", "名称", "做法", "适合的情况", "需要写清的条件", "建议"],
        rows,
        [12, 28, 36, 32, 40, 16],
        center_cols={6},
        bold_cols={1, 2},
        row_height=48,
    )
    for c in range(1, 7):
        ws.cell(end - 1, c).fill = fill_sand


def build_timeline(wb):
    ws = wb.create_sheet("13-推进时间")
    rows = [row for row in D.TIMELINE]
    rows += [(item, "意向明确后由主办团队更新", "", "重排时必须明确") for item in D.REPLAN]
    end = add_table(
        ws, "推进时间",
        "对照日期 " + D.AS_OF_CN + "。状态只说明原稿节点与今天的关系，不代表已经核验团队的实际进度。",
        ["阶段或事项", "原计划时间", "主要工作", "相对今天"],
        rows,
        [24, 24, 55, 18],
        center_cols={4},
        bold_cols={1},
        row_height=26,
    )
    highlight_status(ws, 4, 5, 4 + len(D.TIMELINE))


def build_raci(wb):
    ws = wb.create_sheet("14-职责分工")
    rows = [(item, "活动主办及运营团队", "原则上负责") for item in D.HOST_DUTIES]
    rows += [(item, "一滴水", "按合同约定负责") for item in D.VENUE_DUTIES]
    rows += [(item, "双方", "共同负责") for item in D.SHARED_DUTIES]
    rows += [(item, "知识产权与合规", "正式活动前明确") for item in D.IP_POINTS]
    rows += [(item, "延期与取消", "建议写入合同") for item in D.CANCEL_POINTS]
    add_table(
        ws, "职责、合规与取消",
        "注明非官方活动，并不能替代知识产权审核。保险、场地损坏、退票和未交付权益，均写入正式合同。",
        ["事项", "责任主体", "性质"],
        rows,
        [42, 28, 22],
        center_cols={3},
        bold_cols={1},
        row_height=24,
    )


def build_eval(wb):
    ws = wb.create_sheet("15-评估指标")
    rows = [(item, "活动结束后统计", "客流") for item in D.EVAL_TRAFFIC]
    rows += [(item, "只统计一滴水实际发生的收入和成本", "经营") for item in D.EVAL_BUSINESS]
    rows += [(item, "尽量提供后台截图或平台统计", "传播") for item in D.EVAL_MEDIA]
    rows += [(item, "建议新增，用来判断活动有没有变成客户", "转化") for item in D.EVAL_CONVERT]
    rows += [(item, "商业使用需另行取得授权", "交付物") for item in D.DELIVERABLES]
    add_table(
        ws, "效果评估与交付",
        "活动结束后 3—7 天提交复盘。双方开会后再决定是否进入长期合作。转化数据是建议新增的经营指标。",
        ["指标或交付物", "口径", "类别"],
        rows,
        [36, 48, 14],
        center_cols={3},
        bold_cols={1},
        row_height=24,
    )


def build_issues(wb):
    ws = wb.create_sheet("16-待确认")
    rows = [(i, a, b, c, d, e) for i, (a, b, c, d, e) in enumerate(D.OPEN_ISSUES, start=1)]
    # OPEN_ISSUES tuples: name, level, gap, usage, owner
    rows = []
    for i, (name, level, gap, usage, owner) in enumerate(D.OPEN_ISSUES, start=1):
        rows.append((f"{i:02d}", name, level, gap, usage, owner))
    end = add_table(
        ws, "签约前必须确认的事项",
        "这八项未闭合前，方案可以用于内部研讨，不宜作为已定合作或对外宣传承诺。",
        ["序号", "事项", "优先级", "当前缺口", "签约前的使用口径", "建议确认方"],
        rows,
        [8, 16, 12, 55, 42, 18],
        center_cols={1, 3},
        bold_cols={2},
        row_height=48,
    )
    highlight_status(ws, 3, 5, end - 1)


def build_model(wb):
    ws = wb.create_sheet("17-收益测算")
    ws.sheet_properties.tabColor = GOLD
    apply_page(ws, "收益测算")
    set_widths(ws, [8, 42, 20, 28, 46])
    write_banner(
        ws,
        "一滴水本次活动收益测算",
        "黄色单元格填写金额，单位为元。绿色单元格是公式，请保留。未来订单和机会成本单独列示，不自动计入本次直接经营结果。未填写的金额按 0 计算。",
        5,
    )
    headers = ["序号", "项目", "金额（元）", "口径", "填写说明"]
    write_headers(ws, headers)

    cursor = 5
    blocks = [
        ("一、本次直接收入（计入本次结果）", D.DIRECT_INCOME_INPUTS, "income"),
        ("二、本次直接成本（从本次结果中扣减）", D.DIRECT_COST_INPUTS, "cost"),
        ("三、单独列示（不进入本次直接经营结果）", D.SEPARATE_INPUTS, "separate"),
    ]
    income_rows = []
    cost_rows = []
    separate_rows = []
    for title, items, kind in blocks:
        ws.merge_cells(start_row=cursor, start_column=1, end_row=cursor, end_column=5)
        cell = ws.cell(cursor, 1, title)
        cell.font = Font(name=font_name, size=12, bold=True, color=WHITE)
        cell.fill = fill_header_teal
        cell.alignment = Alignment(vertical="center")
        for c in range(1, 6):
            ws.cell(cursor, c).fill = fill_header_teal
            ws.cell(cursor, c).border = thin
        ws.row_dimensions[cursor].height = 24
        cursor += 1
        for i, (name, scope, remark) in enumerate(items, start=1):
            ws.cell(cursor, 1, i)
            ws.cell(cursor, 2, name)
            ws.cell(cursor, 3, None)
            ws.cell(cursor, 4, scope)
            ws.cell(cursor, 5, remark)
            for c in range(1, 6):
                style_body_cell(ws.cell(cursor, c), fill_white, center=(c in (1, 3)), bold=(c == 2))
            ws.cell(cursor, 3).fill = fill_yellow
            ws.cell(cursor, 3).number_format = '#,##0.00'
            ws.cell(cursor, 3).font = input_font
            ws.row_dimensions[cursor].height = 28
            if kind == "income":
                income_rows.append(cursor)
            elif kind == "cost":
                cost_rows.append(cursor)
            else:
                separate_rows.append(cursor)
            cursor += 1
        cursor += 1

    def sum_formula(rows):
        return "=" + "+".join(f"N(C{r})" for r in rows)

    # Result block
    result_title = cursor
    ws.merge_cells(start_row=cursor, start_column=1, end_row=cursor, end_column=5)
    cell = ws.cell(cursor, 1, "四、本次直接经营结果")
    cell.font = Font(name=font_name, size=12, bold=True, color=WHITE)
    for c in range(1, 6):
        ws.cell(cursor, c).fill = fill_navy
        ws.cell(cursor, c).font = header_font
        ws.cell(cursor, c).border = thin
    ws.row_dimensions[cursor].height = 24
    cursor += 1

    income_total_row = cursor
    cost_total_row = cursor + 1
    net_row = cursor + 2
    labels = [
        (income_total_row, "直接收入合计", sum_formula(income_rows), "只加总本次直接收入"),
        (cost_total_row, "直接成本合计", sum_formula(cost_rows), "只加总本次直接成本"),
        (net_row, "本次直接经营结果", f"=N(C{income_total_row})-N(C{cost_total_row})", "不含机会成本，不含未来订单"),
    ]
    for row, label, formula, remark in labels:
        ws.cell(row, 1, "")
        ws.cell(row, 2, label)
        ws.cell(row, 3, formula)
        ws.cell(row, 4, "公式")
        ws.cell(row, 5, remark)
        for c in range(1, 6):
            style_body_cell(ws.cell(row, c), fill_green, center=(c in (1, 3)), bold=True)
        ws.cell(row, 3).number_format = '#,##0.00'
        ws.cell(row, 3).font = formula_font
        ws.row_dimensions[row].height = 26
    cursor = net_row + 2

    ws.merge_cells(start_row=cursor, start_column=1, end_row=cursor, end_column=5)
    ws.cell(cursor, 1, "五、填写状态与决策提示").font = section_font
    cursor += 1
    income_cells = ",".join(f"C{r}" for r in income_rows + cost_rows)
    # COUNTA across the input cells
    count_parts = ",".join(f"C{r}" for r in income_rows + cost_rows)
    ws.cell(cursor, 2, "已填写的金额项")
    ws.cell(cursor, 3, f"=COUNTA({count_parts})")
    ws.cell(cursor, 4, "公式")
    ws.cell(cursor, 5, f"收入和成本共 {len(income_rows) + len(cost_rows)} 项。0 也算已填写。")
    for c in range(1, 6):
        style_body_cell(ws.cell(cursor, c), fill_ice, center=(c == 3), bold=True)
    ws.cell(cursor, 3).font = formula_font
    status_row = cursor
    cursor += 1
    ws.cell(cursor, 2, "测算状态")
    ws.cell(cursor, 3, f'=IF(C{status_row}=0,"尚未填写","已录入金额，请核对口径后再判断")')
    ws.merge_cells(start_row=cursor, start_column=3, end_row=cursor, end_column=5)
    for c in range(1, 6):
        style_body_cell(ws.cell(cursor, c), fill_sand, bold=True)
    ws.row_dimensions[cursor].height = 24
    cursor += 2
    notes = [
        "直接经营结果为正，只说明本次合同收入覆盖了本次增量成本，还要对照第三部分的机会成本。",
        "直接经营结果为负时，请写明一滴水认可的市场投入上限，并对应宣传、客资和后续合作权益。",
        "第三部分即使填写了数字，也不会改变第四部分的结果。未来订单成交后，再记入后续经营，不回写本次活动利润。",
        "品牌传播不在本表货币化。传播效果用「15-评估指标」中的可核查数据复盘。",
        "本表不替代报价单或合同。场地租金、餐饮最低消费和分成比例，以双方盖章文本为准。",
    ]
    for text in notes:
        ws.merge_cells(start_row=cursor, start_column=1, end_row=cursor, end_column=5)
        cell = ws.cell(cursor, 1, text)
        cell.font = body_font
        cell.alignment = wrap
        ws.row_dimensions[cursor].height = 22
        cursor += 1

    ws.auto_filter.ref = None
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:E4"
    # Don't filter the whole sheet because of section rows.
    ws.auto_filter.ref = None

    # Store rows for the self-test by named ranges
    wb.defined_names.add(DefinedName(name="直接经营结果", attr_text=f"'{ws.title}'!$C${net_row}"))
    wb.defined_names.add(DefinedName(name="测算状态", attr_text=f"'{ws.title}'!$C${status_row + 1}"))


def build_theme(wb):
    ws = wb.create_sheet("18-主题互动")
    unified = []
    for zone in D.ZONES:
        colors = "、".join(f"{name} {hex_color}" for hex_color, name in zone["colors"])
        unified.append(("主题区", zone["name"], zone["place"], colors, zone["set"] + zone["use"]))
    for num, title, body in D.INTERACTIONS:
        unified.append(("互动", f"{num} {title}", "—", body, "奖励或纪念内容见玩法说明"))
    add_table(
        ws, "主题区与互动",
        "色彩为原稿创意。角色形象、标识、官方素材和衍生展示，完成知识产权审核后才能使用。阿戈尔联络站的 AR 和数字内容，需要授权与技术支持。",
        ["类别", "名称或编号", "主题", "色彩或玩法", "空间作用"],
        unified,
        [12, 18, 28, 55, 36],
        center_cols={1},
        bold_cols={2},
        row_height=36,
    )


def build():
    wb = Workbook()
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.properties.title = "潮起潮落 TIDELINE × 北外滩尚9·一滴水 方案附表"
    wb.properties.subject = "综合合作方案配套工作簿"
    wb.properties.category = "内部研讨稿"
    wb.properties.keywords = "一滴水,TIDELINE,赞助,测算"
    build_cover(wb)
    build_event(wb)
    build_value(wb)
    build_longterm(wb)
    build_sponsor(wb)
    build_media(wb)
    build_space(wb)
    build_rundown(wb)
    build_team(wb)
    build_safety(wb)
    build_tasks(wb)
    build_money(wb)
    build_modes(wb)
    build_timeline(wb)
    build_raci(wb)
    build_eval(wb)
    build_issues(wb)
    build_model(wb)
    build_theme(wb)
    # Sanity: area total and required phrases
    if D.area_total() != 1780:
        raise RuntimeError("面积合计异常")
    expected = [
        "00-阅读说明", "01-活动信息", "02-商业价值", "03-长期产品", "04-赞助权益",
        "05-传播计划", "06-空间需求", "07-活动日程", "08-团队案例", "09-执行保障",
        "10-场地配合", "11-收入与成本", "12-合作模式", "13-推进时间", "14-职责分工",
        "15-评估指标", "16-待确认", "17-收益测算", "18-主题互动",
    ]
    names = wb.sheetnames
    if names != expected:
        raise RuntimeError("工作表不符：" + "、".join(names))
    return wb


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "deliverables",
        "潮起潮落TIDELINE-北外滩尚9一滴水-方案附表.xlsx",
    )
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb = build()
    wb.save(out)
    print(f"{out}\nsheets={len(wb.sheetnames)}")


if __name__ == "__main__":
    main()
