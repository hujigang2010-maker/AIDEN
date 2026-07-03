# -*- coding: utf-8 -*-
"""合集扩展版 Excel:合并两个项目的提资清单与报价,并扩充市场数据表。

工作表:
  1. 项目总览      两个项目 + 报价对比
  2. 提资清单(合并) 两项目提资清单合并,含"适用项目"列
  3. 报价汇总      两项目四阶段报价 + 合计对比
  4. 服务模块对照  八大模块两项目要点对照
  5. 附件一 出口TOP20
  6. 附件二 消费类出口TOP20
  7. 附件三 楼层功能布局
商务大气风格:深藏青+金色主题、横幅标题、分区表头、斑马纹、冻结窗格。
"""
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

import content_zsc as ZSC
import content_hp as HP
import content_cgc as CGC

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "deliverables")
FONT = "微软雅黑"
NAVY = "0F2A4A"
BLUE = "1F4E79"
GOLD = "C79A3B"
LIGHT = "EEF2F8"
LIGHTGOLD = "F6EEDA"
WHITE = "FFFFFF"
DARK = "222B35"

med = Side(style="thin", color="AAB4C0")
BORDER = Border(left=med, right=med, top=med, bottom=med)
WRAP = Alignment(wrap_text=True, vertical="center")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

ZSC_NAME = "广州知识城"
HP_NAME = "黄埔九佛TOD"


def setup(ws, tab_color=GOLD):
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.sheet_properties.tabColor = tab_color


def widths(ws, ws_widths):
    for j, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w


def banner(ws, ncols, title, subtitle=None):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=title)
    c.font = Font(name=FONT, size=15, bold=True, color=WHITE)
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = CENTER
    ws.row_dimensions[1].height = 34
    r = 2
    if subtitle:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
        c = ws.cell(row=2, column=1, value=subtitle)
        c.font = Font(name=FONT, size=10, italic=True, color="5A5A5A")
        c.fill = PatternFill("solid", fgColor=LIGHTGOLD)
        c.alignment = CENTER
        ws.row_dimensions[2].height = 20
        r = 3
    return r


def header(ws, row, headers, fill=GOLD):
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=j, value=h)
        c.font = Font(name=FONT, size=10.5, bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[row].height = 26


def drow(ws, row, values, center_cols=(), h=22, zebra=True, bold=False):
    fill = PatternFill("solid", fgColor=LIGHT) if (zebra and row % 2 == 0) else None
    for j, v in enumerate(values, 1):
        c = ws.cell(row=row, column=j, value=v)
        c.font = Font(name=FONT, size=10, bold=bold, color=DARK)
        c.alignment = CENTER if (j) in center_cols else LEFT
        c.border = BORDER
        if fill and not bold:
            c.fill = fill
    ws.row_dimensions[row].height = h


def total_row(ws, row, values, center_cols=()):
    for j, v in enumerate(values, 1):
        c = ws.cell(row=row, column=j, value=v)
        c.font = Font(name=FONT, size=10.5, bold=True, color=NAVY)
        c.fill = PatternFill("solid", fgColor=LIGHTGOLD)
        c.alignment = CENTER if j in center_cols else LEFT
        c.border = BORDER
    ws.row_dimensions[row].height = 24


def note(ws, row, ncols, text, h=30):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=9.5, italic=True, color="666666")
    c.alignment = LEFT
    ws.row_dimensions[row].height = h


# ------------------------------------------------------------ 1 项目总览
def sheet_overview(wb):
    ws = wb.active
    ws.title = "项目总览"
    setup(ws)
    widths(ws, [16, 52, 52])
    r = banner(ws, 3, "全球自贸365街区项目群 · 总览",
               "提供方:" + ZSC.PROVIDER_LINE + "　|　日期:2026年7月")
    header(ws, r, ["项目对比", ZSC_NAME + "「全球自贸365街区」", HP_NAME + "「全球自贸365街区」"])
    r += 1
    zsc_ind = "；".join(f"{n}:{d}" for n, d in ZSC.INDUSTRY_GROUPS)
    hp_ind = "；".join(f"{n}:{d}" for n, d in HP.INDUSTRY_GROUPS)
    data = [
        ("项目全称", ZSC.PROJECT_NAME, HP.PROJECT_NAME),
        ("委托方", ZSC.CLIENT, HP.CLIENT),
        ("核心概念", ZSC.CONCEPT_POINTS[0], HP.CONCEPT_POINTS[0]),
        ("产业方向", zsc_ind, hp_ind),
        ("服务模块", "、".join(n for n, _ in ZSC.SERVICE_MODULES),
         "、".join(n for n, _ in HP.SERVICE_MODULES)),
        ("提资清单项数", f"{len(ZSC.INFO_REQUEST_ITEMS)} 项", f"{len(HP.INFO_REQUEST_ITEMS)} 项"),
        ("报价合计(含税参考)", f"{ZSC.QUOTATION_TOTAL} 万元", f"{HP.QUOTATION_TOTAL} 万元"),
    ]
    for label, a, b in data:
        drow(ws, r, [label, a, b], center_cols=(1,), h=None, zebra=True)
        nlines = max(len(a) // 24, len(b) // 24) + 1
        ws.row_dimensions[r].height = max(24, min(nlines * 16 + 8, 160))
        r += 1
    total_row(ws, r, ["两项目报价合计",
                      f"{ZSC.QUOTATION_TOTAL + HP.QUOTATION_TOTAL} 万元（含税参考价，最终以商务洽谈为准）", ""],
              center_cols=(1,))
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws.freeze_panes = "A" + str(r - len(data))


# ------------------------------------------------------------ 2 提资清单(合并)
def sheet_info(wb):
    ws = wb.create_sheet("提资清单（合并）")
    setup(ws)
    widths(ws, [8, 14, 16, 30, 40, 9])
    r = banner(ws, 6, "提资清单（两项目合并）",
               "请委托方提供以下资料；“适用项目”标注该资料适用范围")
    header(ws, r, ["序号", "适用项目", "类别", "资料名称", "说明", "优先级"])
    r += 1
    n = 1
    for tag, C in ((ZSC_NAME, ZSC), (HP_NAME, HP)):
        for (cat, name, desc, prio) in C.INFO_REQUEST_ITEMS:
            drow(ws, r, [n, tag, cat, name, desc, prio], center_cols=(1, 2, 6), h=30)
            r += 1
            n += 1
    note(ws, r, 6, "说明：标注“高”优先级的资料请于合同签订后10个工作日内提供；"
                   "涉密资料可提供脱敏版本或安排现场查阅；暂缺资料请注明，由服务方协助补充收集。", h=40)
    ws.freeze_panes = "A" + str(r - (len(ZSC.INFO_REQUEST_ITEMS) + len(HP.INFO_REQUEST_ITEMS)))


# ------------------------------------------------------------ 3 报价汇总
def sheet_quote(wb):
    ws = wb.create_sheet("报价汇总")
    setup(ws)
    widths(ws, [14, 26, 44, 30, 12])
    r = banner(ws, 5, "策划服务报价汇总（两项目）",
               "金额为含税参考价（人民币），最终以双方商务洽谈及合同为准")
    grand = 0
    for tag, C in ((ZSC_NAME, ZSC), (HP_NAME, HP)):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        c = ws.cell(row=r, column=1, value=f"■ {tag}「全球自贸365街区」项目")
        c.font = Font(name=FONT, size=11, bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=BLUE)
        c.alignment = LEFT
        ws.row_dimensions[r].height = 24
        r += 1
        header(ws, r, ["项目", "工作阶段", "主要工作内容", "主要成果", "报价(万元)"])
        r += 1
        for (stage, work, output, fee) in C.QUOTATION_ITEMS:
            drow(ws, r, [tag, stage, work, output, fee], center_cols=(1, 5), h=40)
            r += 1
        total_row(ws, r, [tag, "小计", "", "", C.QUOTATION_TOTAL], center_cols=(5,))
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        grand += C.QUOTATION_TOTAL
        r += 2
    total_row(ws, r, ["合计", "两项目报价总计", "", "", grand], center_cols=(5,))
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    ws.row_dimensions[r].height = 26
    r += 2
    for i, nt in enumerate(ZSC.QUOTATION_NOTES):
        note(ws, r, 5, f"{i+1}. {nt}", h=22)
        r += 1


# ------------------------------------------------------------ 4 服务模块对照
def sheet_modules(wb):
    ws = wb.create_sheet("服务模块对照")
    setup(ws)
    widths(ws, [6, 20, 42, 42])
    r = banner(ws, 4, "策划服务八大模块 · 两项目对照")
    header(ws, r, ["序号", "服务模块", ZSC_NAME + "要点", HP_NAME + "要点"])
    r += 1
    for i in range(len(ZSC.SERVICE_MODULES)):
        zname, zdesc = ZSC.SERVICE_MODULES[i]
        hname, hdesc = HP.SERVICE_MODULES[i]
        module = zname if zname == hname else f"{zname} / {hname}"
        drow(ws, r, [i + 1, module, zdesc, hdesc], center_cols=(1,), h=64)
        r += 1
    ws.freeze_panes = "A" + str(r - len(ZSC.SERVICE_MODULES))


# ------------------------------------------------------------ 5 出口TOP20
def sheet_export(wb):
    ws = wb.create_sheet("附件一_出口TOP20")
    setup(ws, tab_color=BLUE)
    widths(ws, [6, 22, 22, 16, 46])
    r = banner(ws, 5, "附件一 · 2025年广州出口TOP20品类")
    header(ws, r, ["排名", "品类", "代表品牌", "广州口岸交易额", "核心说明"])
    r += 1
    for row in CGC.EXPORT_TOP20:
        drow(ws, r, list(row), center_cols=(1, 4), h=22)
        r += 1
    note(ws, r, 5, "说明：" + CGC.EXPORT_TOP20_NOTE, h=40)
    ws.freeze_panes = "A" + str(r - len(CGC.EXPORT_TOP20))


# ------------------------------------------------------------ 6 消费类TOP20
def sheet_consumer(wb):
    ws = wb.create_sheet("附件二_消费类TOP20")
    setup(ws, tab_color=BLUE)
    widths(ws, [6, 20, 12, 10, 26, 20])
    r = banner(ws, 6, "附件二 · 2025年前10月广州消费类出口20强（单位：亿元）")
    header(ws, r, ["排名", "品类", "出口额", "同比增速", "核心出口品牌", "核心市场"])
    r += 1
    for row in CGC.CONSUMER_TOP20:
        drow(ws, r, list(row), center_cols=(1, 3, 4), h=22)
        r += 1
    note(ws, r, 6, "说明：" + CGC.CONSUMER_TOP20_NOTE, h=40)
    ws.freeze_panes = "A" + str(r - len(CGC.CONSUMER_TOP20))


# ------------------------------------------------------------ 7 楼层功能布局
def sheet_floor(wb):
    ws = wb.create_sheet("附件三_楼层功能布局")
    setup(ws, tab_color=BLUE)
    widths(ws, [16, 26, 42, 40])
    r = banner(ws, 4, "附件三 · 黄埔九佛TOD「全球自贸365街区」楼层功能布局")
    header(ws, r, ["区位楼层", "功能定位", "核心业态 / 服务", "数据支撑与参考案例"])
    r += 1
    for row in CGC.FLOOR_LAYOUT:
        drow(ws, r, list(row), center_cols=(1,), h=42)
        r += 1


def main():
    wb = Workbook()
    sheet_overview(wb)
    sheet_info(wb)
    sheet_quote(wb)
    sheet_modules(wb)
    sheet_export(wb)
    sheet_consumer(wb)
    sheet_floor(wb)
    os.makedirs(OUT_DIR, exist_ok=True)
    out = os.path.join(OUT_DIR, "全球自贸365街区项目群_提资清单报价与市场数据_合并版.xlsx")
    wb.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main()
