"""Generate sponsorship tracking Excel workbook.

Produces: 赞助方案/赞助归集表-2026峰会.xlsx

Sheets:
1. 仪表盘    - KPI summary with formulas referencing 赞助登记
2. 赞助登记 - Main data-entry table for prospects/contracts
3. 权益对照 - Tier rights comparison reference
4. 资源置换 - Barter partners tracker
5. 物料交付 - Asset delivery checklist
6. 收款明细 - Payment ledger
7. 嘉宾对接 - VIP intro requests
8. 字段说明 - Field/data dictionary
"""
from __future__ import annotations

import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule


OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "赞助方案")
OUT_PATH = os.path.normpath(os.path.join(OUT_DIR, "赞助归集表-2026峰会.xlsx"))


NAVY = "FF0E1F44"
BLUE = "FF1F497D"
GOLD = "FFC9A227"
LIGHT = "FFF4F6FA"
WHITE = "FFFFFFFF"
GRAY = "FF4A4A4A"


def fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)


def thin_border() -> Border:
    s = Side(style="thin", color="FFB6BCC8")
    return Border(left=s, right=s, top=s, bottom=s)


def style_header(cell, *, bg=BLUE, color=WHITE, size=11):
    cell.font = Font(name="Microsoft YaHei", bold=True, color=color, size=size)
    cell.fill = fill(bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border()


def style_body(cell, *, bold=False, color=GRAY, align="left", wrap=True, size=10):
    cell.font = Font(name="Microsoft YaHei", bold=bold, color=color, size=size)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    cell.border = thin_border()


def style_kpi_label(cell):
    cell.font = Font(name="Microsoft YaHei", bold=True, color=WHITE, size=12)
    cell.fill = fill(NAVY)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border()


def style_kpi_value(cell):
    cell.font = Font(name="Microsoft YaHei", bold=True, color=BLUE, size=20)
    cell.fill = fill(LIGHT)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border()


def set_col_widths(ws, widths: list[int]) -> None:
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def freeze(ws, cell: str) -> None:
    ws.freeze_panes = cell


# ============================================================
# 1. 赞助登记
# ============================================================
def build_sponsors(wb: Workbook):
    ws = wb.create_sheet("赞助登记")
    headers = [
        "编号", "企业名称", "简称 / Logo 名", "对接联系人", "职务", "电话",
        "微信 / 邮箱", "意向级别", "确认级别", "金额(元)", "签约状态",
        "回款状态", "回款金额(元)", "回款日期", "发票类型", "发票抬头/税号",
        "意向来源 / 引荐人", "权益备注", "物料交付", "嘉宾对接需求", "更新日期", "负责人",
    ]
    ws.append(headers)
    set_col_widths(ws, [
        6, 26, 16, 12, 12, 16, 22, 14, 14, 14, 14,
        14, 14, 14, 16, 28, 18, 28, 14, 22, 14, 12,
    ])
    ws.row_dimensions[1].height = 36
    for c in range(1, len(headers) + 1):
        style_header(ws.cell(row=1, column=c))

    sample_rows = [
        [1, "示例科技股份有限公司", "示例科技", "张总", "市场总监",
         "138-0000-0000", "wechat / market@example.com",
         "钻石(5万)", "钻石(5万)", 50000, "已签约", "已到账", 50000,
         "2026-04-20", "增值税专用发票", "示例科技股份有限公司 / 91XXXXXXXX",
         "校友黄欣引荐", "AI 硬核圆桌席位 + 颁奖授牌", "已交付", "夏春博士 1V1",
         "2026-04-22", "李三"],
        [2, "示例资本管理有限公司", "示例资本", "王合伙人", "管理合伙人",
         "139-0000-0000", "wm@example.cn",
         "铂金(3万)", "", 0, "洽谈中", "未到账", 0, "",
         "", "", "联盟内部转介", "音乐会环节鸣谢", "待提交", "曲承东对接",
         "2026-04-25", "李三"],
        [3, "示例云科技", "示例云", "刘总", "副总裁",
         "186-0000-0000", "liu@example.com",
         "总冠名", "", 0, "意向初谈", "未到账", 0, "",
         "", "", "组委会主动 BD", "需 15min 主旨演讲席",
         "未交付", "晚宴主桌锁定", "2026-04-28", "王四"],
    ]
    for row in sample_rows:
        ws.append(row)

    # Body styling for sample + 200 empty rows for direct entry
    total_rows = 200
    for r in range(2, 2 + total_rows):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            if c == 10 or c == 13:  # money
                cell.number_format = '#,##0'
                style_body(cell, align="right")
            elif c == 14 or c == 21:  # dates
                cell.number_format = 'yyyy-mm-dd'
                style_body(cell, align="center")
            elif c == 1:
                style_body(cell, align="center", bold=True)
            else:
                style_body(cell, align="left" if c in (2, 16, 17, 18, 20) else "center")

    # Data validations
    dv_intent = DataValidation(
        type="list",
        formula1='"总冠名,钻石(5万),铂金(3万),黄金(1万),资源置换,基础曝光(1500),其他"',
        allow_blank=True,
    )
    dv_intent.add(f"H2:I{1 + total_rows}")
    ws.add_data_validation(dv_intent)

    dv_sign = DataValidation(
        type="list",
        formula1='"意向初谈,洽谈中,口头确认,合同审核,已签约,已放弃"',
        allow_blank=True,
    )
    dv_sign.add(f"K2:K{1 + total_rows}")
    ws.add_data_validation(dv_sign)

    dv_pay = DataValidation(
        type="list",
        formula1='"未到账,部分到账,已到账,已开票"',
        allow_blank=True,
    )
    dv_pay.add(f"L2:L{1 + total_rows}")
    ws.add_data_validation(dv_pay)

    dv_invoice = DataValidation(
        type="list",
        formula1='"增值税专用发票,增值税普通发票,不开票"',
        allow_blank=True,
    )
    dv_invoice.add(f"O2:O{1 + total_rows}")
    ws.add_data_validation(dv_invoice)

    dv_mat = DataValidation(
        type="list",
        formula1='"未交付,部分交付,已交付,无需交付"',
        allow_blank=True,
    )
    dv_mat.add(f"S2:S{1 + total_rows}")
    ws.add_data_validation(dv_mat)

    # Conditional formatting: signed = green, abandoned = red
    ws.conditional_formatting.add(
        f"K2:K{1 + total_rows}",
        CellIsRule(operator="equal", formula=['"已签约"'],
                   fill=PatternFill("solid", fgColor="FFC6EFCE"),
                   font=Font(name="Microsoft YaHei", bold=True, color="FF006100")),
    )
    ws.conditional_formatting.add(
        f"K2:K{1 + total_rows}",
        CellIsRule(operator="equal", formula=['"已放弃"'],
                   fill=PatternFill("solid", fgColor="FFFFC7CE"),
                   font=Font(name="Microsoft YaHei", bold=True, color="FF9C0006")),
    )
    ws.conditional_formatting.add(
        f"L2:L{1 + total_rows}",
        CellIsRule(operator="equal", formula=['"已到账"'],
                   fill=PatternFill("solid", fgColor="FFC6EFCE")),
    )

    freeze(ws, "C2")


# ============================================================
# 2. 仪表盘
# ============================================================
def build_dashboard(wb: Workbook):
    ws = wb.create_sheet("仪表盘", 0)
    set_col_widths(ws, [4, 22, 18, 4, 22, 18, 4, 22, 18])

    # Title
    ws.merge_cells("B2:I3")
    t = ws["B2"]
    t.value = "重构与突围 · 2026 AI 商业化峰会  ——  赞助归集仪表盘"
    t.font = Font(name="Microsoft YaHei", bold=True, color=WHITE, size=18)
    t.fill = fill(NAVY)
    t.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("B4:I4")
    s = ws["B4"]
    s.value = "数据来源：赞助登记表  ·  字段口径见「字段说明」  ·  截止日期：5 月 18 日"
    s.font = Font(name="Microsoft YaHei", color=WHITE, size=11)
    s.fill = fill(BLUE)
    s.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 22

    # KPI cards
    kpi = [
        ("B6", "B7", "意向客户总数", '=COUNTA(赞助登记!B2:B201)'),
        ("E6", "E7", "已签约客户数",
         '=COUNTIF(赞助登记!K2:K201,"已签约")'),
        ("H6", "H7", "已到账客户数",
         '=COUNTIF(赞助登记!L2:L201,"已到账")'),

        ("B9", "B10", "意向金额合计(元)", '=SUM(赞助登记!J2:J201)'),
        ("E9", "E10", "已回款金额(元)", '=SUM(赞助登记!M2:M201)'),
        ("H9", "H10", "回款率",
         '=IFERROR(SUM(赞助登记!M2:M201)/SUM(赞助登记!J2:J201),0)'),
    ]
    for label_cell, value_cell, label, formula in kpi:
        ws[label_cell] = label
        style_kpi_label(ws[label_cell])
        ws[value_cell] = formula
        style_kpi_value(ws[value_cell])
        if "金额" in label:
            ws[value_cell].number_format = '"￥"#,##0'
        if label == "回款率":
            ws[value_cell].number_format = "0.0%"
    for r in (6, 9):
        ws.row_dimensions[r].height = 24
    for r in (7, 10):
        ws.row_dimensions[r].height = 40

    # Section header: tier breakdown
    ws.merge_cells("B12:I12")
    h = ws["B12"]
    h.value = "按级别拆分"
    h.font = Font(name="Microsoft YaHei", bold=True, color=WHITE, size=12)
    h.fill = fill(BLUE)
    h.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[12].height = 22

    tiers = [
        ("总冠名", 1, 0),
        ("钻石(5万)", 3, 50000),
        ("铂金(3万)", 5, 30000),
        ("黄金(1万)", 8, 10000),
        ("资源置换", 5, 0),
        ("基础曝光(1500)", 30, 1500),
    ]

    layout = [
        ("B", "级别", 16),
        ("C", "目标名额", 10),
        ("D", "意向数", 10),
        ("E", "已签约数", 10),
        ("F", "意向金额(元)", 16),
        ("G", "已回款(元)", 16),
        ("H", "签约率", 10),
        ("I", "回款率", 10),
    ]
    for col, label, _w in layout:
        c = ws[f"{col}13"]
        c.value = label
        style_header(c)
    ws.row_dimensions[13].height = 28

    for i, (tier, slot, _price) in enumerate(tiers):
        r = 14 + i
        ws[f"B{r}"] = tier
        ws[f"C{r}"] = slot
        ws[f"D{r}"] = f'=COUNTIF(赞助登记!H2:H201,"{tier}")'
        ws[f"E{r}"] = (f'=COUNTIFS(赞助登记!I2:I201,"{tier}",'
                      f'赞助登记!K2:K201,"已签约")')
        ws[f"F{r}"] = (f'=SUMIFS(赞助登记!J2:J201,赞助登记!I2:I201,"{tier}")')
        ws[f"G{r}"] = (f'=SUMIFS(赞助登记!M2:M201,赞助登记!I2:I201,"{tier}")')
        ws[f"H{r}"] = f'=IFERROR(E{r}/C{r},0)'
        ws[f"I{r}"] = f'=IFERROR(G{r}/F{r},0)'

        for col in "BCDEFGHI":
            cell = ws[f"{col}{r}"]
            if col in ("F", "G"):
                cell.number_format = '"￥"#,##0'
            elif col in ("H", "I"):
                cell.number_format = "0.0%"
            style_body(cell, align="center" if col != "B" else "left",
                       bold=(col == "B"))

    # Section header: signing status
    r = 14 + len(tiers) + 1
    ws.merge_cells(f"B{r}:I{r}")
    h = ws.cell(row=r, column=2)
    h.value = "按签约状态拆分"
    h.font = Font(name="Microsoft YaHei", bold=True, color=WHITE, size=12)
    h.fill = fill(BLUE)
    h.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[r].height = 22

    statuses = ["意向初谈", "洽谈中", "口头确认", "合同审核", "已签约", "已放弃"]
    r1 = r + 1
    for col, label in [("B", "签约状态"), ("C", "客户数"), ("D", "金额合计(元)"),
                        ("E", "占比")]:
        c = ws[f"{col}{r1}"]
        c.value = label
        style_header(c)
    ws.row_dimensions[r1].height = 28

    for i, st in enumerate(statuses):
        rr = r1 + 1 + i
        ws[f"B{rr}"] = st
        ws[f"C{rr}"] = f'=COUNTIF(赞助登记!K2:K201,"{st}")'
        ws[f"D{rr}"] = (f'=SUMIFS(赞助登记!J2:J201,赞助登记!K2:K201,"{st}")')
        ws[f"E{rr}"] = f'=IFERROR(C{rr}/SUM(C{r1+1}:C{r1+len(statuses)}),0)'
        for col in "BCDE":
            cell = ws[f"{col}{rr}"]
            if col == "D":
                cell.number_format = '"￥"#,##0'
            elif col == "E":
                cell.number_format = "0.0%"
            style_body(cell, align="center" if col != "B" else "left",
                       bold=(col == "B"))

    freeze(ws, "A5")


# ============================================================
# 3. 权益对照
# ============================================================
def build_rights(wb: Workbook):
    ws = wb.create_sheet("权益对照")
    headers = ["权益项", "总冠名", "钻石(5万)", "铂金(3万)", "黄金(1万)", "基础(1500)"]
    rows = [
        ["大会冠名权", "●", "—", "—", "—", "—"],
        ["独立主旨演讲（15 min）", "●", "—", "—", "—", "—"],
        ["圆桌论坛对话席", "●", "● 限 1 席", "—", "—", "—"],
        ["年度颁奖授牌", "●", "●", "—", "—", "—"],
        ["音乐会环节鸣谢", "●", "●", "●", "—", "—"],
        ["主背景板 Logo 等级", "顶级", "钻石", "铂金", "黄金", "—"],
        ["核心动线易拉宝", "●", "●", "● 专属展位", "—", "●"],
        ["大会官网 / 大屏滚动", "●", "●", "●", "●", "—"],
        ["议程手册广告", "扉页整版", "整版", "半版", "1/4 版", "尾页鸣谢"],
        ["白皮书署名", "封面联合 + 扉页", "内页 1/2", "内页 1/4", "—", "—"],
        ["媒体通稿露出", "标题级", "正文重点", "鸣谢", "—", "—"],
        ["VIP 闭门晚宴入场券", "主桌 3 人", "5 张", "3 张", "1 张", "—"],
        ["重量级嘉宾 1V1 闭门", "全程定制", "2 位", "1 位", "—", "—"],
        ["双校友产业联盟入册", "战略合作伙伴", "●", "●", "●", "—"],
        ["参考投入(元)", "面议", "50,000", "30,000", "10,000", "1,500/位"],
        ["开放名额", "1", "3", "5", "8", "不限"],
    ]
    ws.append(headers)
    for row in rows:
        ws.append(row)

    set_col_widths(ws, [28, 18, 16, 16, 16, 18])
    ws.row_dimensions[1].height = 32
    for c in range(1, len(headers) + 1):
        style_header(ws.cell(row=1, column=c))
    for r in range(2, 2 + len(rows)):
        ws.row_dimensions[r].height = 22
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            bold = (c == 1)
            color = NAVY if c == 1 else GRAY
            align = "left" if c == 1 else "center"
            style_body(cell, bold=bold, color=color, align=align)
            if cell.value == "●":
                cell.font = Font(name="Microsoft YaHei", bold=True,
                                 color=GOLD, size=12)
            if r % 2 == 0:
                cell.fill = fill(LIGHT)
    freeze(ws, "B2")


# ============================================================
# 4. 资源置换
# ============================================================
def build_barter(wb: Workbook):
    ws = wb.create_sheet("资源置换")
    headers = ["编号", "类别", "企业名称", "联系人", "电话/邮箱", "置换标的",
               "数量 / 规格", "等值预估(元)", "对应权益", "交付时间", "状态", "备注"]
    ws.append(headers)
    set_col_widths(ws, [6, 14, 24, 12, 22, 26, 16, 14, 28, 14, 12, 24])
    ws.row_dimensions[1].height = 32
    for c in range(1, len(headers) + 1):
        style_header(ws.cell(row=1, column=c))

    samples = [
        [1, "晚宴用酒", "示例酒业有限公司", "陈总", "139-xxxx-xxxx",
         "高端干红 / 单瓶零售约 800 元", "30 瓶起", 24000,
         "主桌 VIP 1 个 + 晚宴专属品鉴介绍", "5/15 前", "已确认", "需提供报关单"],
        [2, "官方伴手礼", "示例文创", "刘总", "wm@example.com",
         "VIP 精装礼盒", "50 份精装版", 25000,
         "VIP 福袋 + 数据回传", "5/16 前", "洽谈中", ""],
        [3, "独家影像", "示例影视", "Eric", "186-xxxx-xxxx",
         "双机位拍摄 + 后期精修", "8 小时全流程", 30000,
         "暖场宣传片轮播 + 联合署名", "大会当日", "意向", ""],
    ]
    for row in samples:
        ws.append(row)

    total_rows = 50
    for r in range(2, 2 + total_rows):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            if c == 8:
                cell.number_format = '#,##0'
                style_body(cell, align="right")
            elif c == 1:
                style_body(cell, align="center", bold=True)
            else:
                style_body(cell, align="left" if c in (3, 6, 9, 12) else "center")

    dv_cat = DataValidation(
        type="list",
        formula1='"晚宴用酒,官方伴手礼,独家影像,官方指定出行,媒体合作,场地/物资,其他"',
        allow_blank=True,
    )
    dv_cat.add(f"B2:B{1 + total_rows}")
    ws.add_data_validation(dv_cat)

    dv_status = DataValidation(
        type="list",
        formula1='"意向,洽谈中,已确认,已交付,已放弃"',
        allow_blank=True,
    )
    dv_status.add(f"K2:K{1 + total_rows}")
    ws.add_data_validation(dv_status)
    freeze(ws, "C2")


# ============================================================
# 5. 物料交付
# ============================================================
def build_assets(wb: Workbook):
    ws = wb.create_sheet("物料交付")
    headers = ["编号", "企业", "Logo(矢量)", "品牌简介(≤200字)", "宣传片(≤60s,1080P)",
               "易拉宝/展位 KV", "广告稿(手册)", "白皮书署名稿",
               "嘉宾出席信息", "其他", "约定截止日", "实际交付日", "状态", "对接人"]
    ws.append(headers)
    set_col_widths(ws, [6, 24, 12, 18, 18, 16, 14, 14, 16, 14, 14, 14, 12, 12])
    ws.row_dimensions[1].height = 36
    for c in range(1, len(headers) + 1):
        style_header(ws.cell(row=1, column=c))

    total_rows = 80
    check_cols = [3, 4, 5, 6, 7, 8, 9, 10]  # delivery checklist columns

    dv_check = DataValidation(
        type="list",
        formula1='"☐ 待提交,◐ 部分,☑ 已交付,— 不适用"',
        allow_blank=True,
    )
    for c in check_cols:
        dv_check.add(f"{get_column_letter(c)}2:{get_column_letter(c)}{1 + total_rows}")
    ws.add_data_validation(dv_check)

    dv_status = DataValidation(
        type="list",
        formula1='"未开始,进行中,部分完成,已完成"',
        allow_blank=True,
    )
    dv_status.add(f"M2:M{1 + total_rows}")
    ws.add_data_validation(dv_status)

    for r in range(2, 2 + total_rows):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            if c in (11, 12):
                cell.number_format = "yyyy-mm-dd"
                style_body(cell, align="center")
            elif c == 1:
                style_body(cell, align="center", bold=True)
            else:
                style_body(cell, align="left" if c == 2 else "center")
    freeze(ws, "C2")


# ============================================================
# 6. 收款明细
# ============================================================
def build_payments(wb: Workbook):
    ws = wb.create_sheet("收款明细")
    headers = ["编号", "日期", "企业名称", "对应级别", "金额(元)", "支付方式",
               "银行流水号", "发票类型", "发票号码", "已开票", "备注", "经办人"]
    ws.append(headers)
    set_col_widths(ws, [6, 14, 26, 14, 14, 14, 22, 16, 18, 10, 24, 12])
    ws.row_dimensions[1].height = 32
    for c in range(1, len(headers) + 1):
        style_header(ws.cell(row=1, column=c))

    total_rows = 100
    for r in range(2, 2 + total_rows):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            if c == 2:
                cell.number_format = "yyyy-mm-dd"
                style_body(cell, align="center")
            elif c == 5:
                cell.number_format = '"￥"#,##0'
                style_body(cell, align="right")
            elif c == 1:
                style_body(cell, align="center", bold=True)
            else:
                style_body(cell, align="left" if c in (3, 11) else "center")

    # Total row
    total_row = 2 + total_rows
    ws.cell(row=total_row, column=4).value = "合计"
    ws.cell(row=total_row, column=4).font = Font(name="Microsoft YaHei",
                                                  bold=True, color=WHITE)
    ws.cell(row=total_row, column=4).fill = fill(NAVY)
    ws.cell(row=total_row, column=4).alignment = Alignment(horizontal="center",
                                                            vertical="center")
    ws.cell(row=total_row, column=5).value = f'=SUM(E2:E{total_rows + 1})'
    ws.cell(row=total_row, column=5).number_format = '"￥"#,##0'
    ws.cell(row=total_row, column=5).font = Font(name="Microsoft YaHei",
                                                  bold=True, color=GOLD, size=14)
    ws.cell(row=total_row, column=5).alignment = Alignment(horizontal="right",
                                                            vertical="center")
    ws.cell(row=total_row, column=5).fill = fill(LIGHT)

    dv_pay = DataValidation(
        type="list",
        formula1='"对公转账,微信,支付宝,现金,资源置换"',
        allow_blank=True,
    )
    dv_pay.add(f"F2:F{1 + total_rows}")
    ws.add_data_validation(dv_pay)

    dv_invoice = DataValidation(
        type="list",
        formula1='"增值税专用发票,增值税普通发票,不开票"',
        allow_blank=True,
    )
    dv_invoice.add(f"H2:H{1 + total_rows}")
    ws.add_data_validation(dv_invoice)

    dv_yn = DataValidation(type="list", formula1='"是,否"', allow_blank=True)
    dv_yn.add(f"J2:J{1 + total_rows}")
    ws.add_data_validation(dv_yn)
    freeze(ws, "C2")


# ============================================================
# 7. 嘉宾对接
# ============================================================
def build_guests(wb: Workbook):
    ws = wb.create_sheet("嘉宾对接")
    headers = ["编号", "申请企业", "申请人", "希望对接嘉宾", "对接目的 / 议题",
               "首选时段", "是否签约赞助", "组委会负责人", "状态", "结果反馈"]
    ws.append(headers)
    set_col_widths(ws, [6, 24, 12, 22, 30, 16, 12, 14, 12, 30])
    ws.row_dimensions[1].height = 32
    for c in range(1, len(headers) + 1):
        style_header(ws.cell(row=1, column=c))

    total_rows = 60
    for r in range(2, 2 + total_rows):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            if c == 1:
                style_body(cell, align="center", bold=True)
            else:
                style_body(cell, align="left" if c in (2, 4, 5, 10) else "center")

    dv_guest = DataValidation(
        type="list",
        formula1=('"白硕,夏春博士,徐永泽(腾讯云),邵凝光(火山引擎),'
                  '吴晓东(阿里云),姚志勇教授,王维军,寇文红,张露瑶博士,'
                  '孔华威,曲承东,刘胜利,饶雪莹,马俊杰,其他"'),
        allow_blank=True,
    )
    dv_guest.add(f"D2:D{1 + total_rows}")
    ws.add_data_validation(dv_guest)

    dv_yn = DataValidation(type="list", formula1='"是,否"', allow_blank=True)
    dv_yn.add(f"G2:G{1 + total_rows}")
    ws.add_data_validation(dv_yn)

    dv_status = DataValidation(
        type="list",
        formula1='"待评估,已转介,已确认,已完成,未通过"',
        allow_blank=True,
    )
    dv_status.add(f"I2:I{1 + total_rows}")
    ws.add_data_validation(dv_status)
    freeze(ws, "C2")


# ============================================================
# 8. 字段说明
# ============================================================
def build_dictionary(wb: Workbook):
    ws = wb.create_sheet("字段说明")
    set_col_widths(ws, [4, 22, 60])
    rows = [
        ("赞助登记", "意向级别 / 确认级别",
         "下拉选项：总冠名 / 钻石(5万) / 铂金(3万) / 黄金(1万) / 资源置换 / "
         "基础曝光(1500) / 其他。\n仪表盘按此字段统计名额完成率。"),
        ("赞助登记", "金额(元)",
         "意向金额或合同金额。资源置换填等值预估，基础曝光按 1500×位数。"),
        ("赞助登记", "签约状态",
         "意向初谈 / 洽谈中 / 口头确认 / 合同审核 / 已签约 / 已放弃。"
         "仅「已签约」计入签约率分母。"),
        ("赞助登记", "回款状态",
         "未到账 / 部分到账 / 已到账 / 已开票。"),
        ("资源置换", "等值预估(元)",
         "用于估算置换价值，仅作管理参考，不计入现金回款合计。"),
        ("收款明细", "金额(元)",
         "实际到账金额。底部合计行自动求和。"),
        ("物料交付", "状态字段",
         "☐ 待提交 / ◐ 部分 / ☑ 已交付 / — 不适用。"),
        ("仪表盘", "回款率",
         "= 已回款金额合计 / 意向金额合计。仅「已签约」客户的回款被使用时建议另行筛选。"),
        ("通用", "录入规范",
         "1. 编号建议从 1 起递增；2. 日期统一 yyyy-mm-dd；"
         "3. 每张表已锁定首行，可直接筛选；4. 灰色样表数据请正式启用前清除。"),
    ]
    ws.merge_cells("B2:D2")
    t = ws["B2"]
    t.value = "字段说明 / 数据字典"
    t.font = Font(name="Microsoft YaHei", bold=True, color=WHITE, size=16)
    t.fill = fill(NAVY)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 30

    headers = ["所属表", "字段", "说明"]
    for i, h in enumerate(headers, start=2):
        c = ws.cell(row=4, column=i)
        c.value = h
        style_header(c)
    ws.row_dimensions[4].height = 24

    for i, (sheet, field, desc) in enumerate(rows, start=5):
        ws.cell(row=i, column=2, value=sheet)
        ws.cell(row=i, column=3, value=field)
        ws.cell(row=i, column=4, value=desc)
        for col in (2, 3, 4):
            cell = ws.cell(row=i, column=col)
            style_body(cell, align="left" if col == 4 else "center",
                       bold=(col == 2))
        ws.row_dimensions[i].height = 48


def main() -> None:
    wb = Workbook()
    wb.remove(wb.active)

    # Order: Dashboard placed at index 0 by build_dashboard call
    build_sponsors(wb)
    build_dashboard(wb)
    build_rights(wb)
    build_barter(wb)
    build_assets(wb)
    build_payments(wb)
    build_guests(wb)
    build_dictionary(wb)

    # Reorder: dashboard first
    wb.move_sheet("仪表盘", offset=-wb.sheetnames.index("仪表盘"))
    wb.active = 0

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    wb.save(OUT_PATH)
    print(f"[OK] Excel saved -> {OUT_PATH}")


if __name__ == "__main__":
    main()
