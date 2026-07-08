# -*- coding: utf-8 -*-
"""生成《下半年三方合作计划》配套明细表（Excel）。"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties
import plan_data as D

NAVY = "2E1F47"    # 深紫（标题/深色底）
BLUE = "5B3E8E"    # 主紫
LTBLUE = "EAE3F5"  # 浅紫（表头/强调底）
GOLD = "B0841A"    # 金（数字/强调）
GREY = "F6F3FB"    # 交替行浅紫灰
WHITE = "FFFFFF"

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
FONT = "Microsoft YaHei"


def style_header(cell):
    cell.font = Font(name=FONT, bold=True, color=WHITE, size=11)
    cell.fill = PatternFill("solid", fgColor=BLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER


def style_cell(cell, wrap=True, center=False, bold=False, fill=None, color="000000"):
    cell.font = Font(name=FONT, size=10, bold=bold, color=color)
    cell.alignment = Alignment(horizontal="center" if center else "left",
                               vertical="center", wrap_text=wrap)
    cell.border = BORDER
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)


def title_row(ws, text, ncols, sub=None):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=text)
    c.font = Font(name=FONT, bold=True, color=WHITE, size=15)
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 34
    start = 2
    if sub:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
        c = ws.cell(row=2, column=1, value=sub)
        c.font = Font(name=FONT, italic=True, color="595959", size=9)
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[2].height = 18
        start = 3
    return start


wb = openpyxl.Workbook()

# ---------------------------------------------------------------- Sheet 1 总览
ws = wb.active
ws.title = "1-12场活动总览"
headers = ["序号", "拟定时间", "活动主题", "产业板块", "形式/规模", "场地建议", "报价档位", "报价(万元)"]
widths = [6, 16, 34, 20, 16, 26, 14, 11]
hr = title_row(ws, "东方枢纽 × 复旦大学 × 上海市科技企业联合会  |  下半年 12 场活动总览",
               len(headers), "以“上海市级 + 浦东新区”资源联动为核心；报价单位：人民币·万元（初步建议，最终以指定策划供应商合同为准）")
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
for j, h in enumerate(headers, 1):
    style_header(ws.cell(row=hr, column=j, value=h))
r = hr + 1
for a in D.ACTIVITIES:
    tier_letter = a["tier"].split(" ")[0]
    vals = [a["no"], a["date"], a["title"], a["sector"], a["scale"], a["venue"], tier_letter, a["price"]]
    for j, v in enumerate(vals, 1):
        center = j in (1, 2, 7, 8)
        style_cell(ws.cell(row=r, column=j, value=v), center=center,
                   fill=GREY if r % 2 else WHITE)
    ws.row_dimensions[r].height = 42
    r += 1
# 合计行
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
c = ws.cell(row=r, column=1, value="12 场合计")
style_cell(c, center=True, bold=True, fill=LTBLUE)
style_cell(ws.cell(row=r, column=8, value=D.TOTAL_PRICE), center=True, bold=True, fill=LTBLUE, color=GOLD)
ws.freeze_panes = ws.cell(row=hr + 1, column=1)

# ---------------------------------------------------------------- Sheet 2 逐场策划
ws = wb.create_sheet("2-逐场详细策划")
headers = ["序号", "主题 / 时间 / 规模", "内容建议", "拟邀嘉宾 / 资源", "招商衔接价值", "报价(万元)"]
widths = [6, 30, 42, 30, 34, 11]
hr = title_row(ws, "12 场活动 · 详细策划案", len(headers))
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
for j, h in enumerate(headers, 1):
    style_header(ws.cell(row=hr, column=j, value=h))
r = hr + 1
for a in D.ACTIVITIES:
    meta = f"{a['title']}\n\n⏰ {a['date']}\n👥 {a['scale']}\n📍 {a['venue']}"
    content = "\n".join(f"• {x}" for x in a["content"])
    guests = "\n".join(f"• {x}" for x in a["guests"])
    invest = "\n".join(f"• {x}" for x in a["invest"])
    fill = GREY if r % 2 else WHITE
    style_cell(ws.cell(row=r, column=1, value=a["no"]), center=True, bold=True, fill=fill)
    style_cell(ws.cell(row=r, column=2, value=meta), fill=fill, bold=False)
    style_cell(ws.cell(row=r, column=3, value=content), fill=fill)
    style_cell(ws.cell(row=r, column=4, value=guests), fill=fill)
    style_cell(ws.cell(row=r, column=5, value=invest), fill=fill)
    style_cell(ws.cell(row=r, column=6, value=a["price"]), center=True, bold=True, fill=fill, color=GOLD)
    ws.row_dimensions[r].height = 120
    r += 1
ws.freeze_panes = ws.cell(row=hr + 1, column=1)

# ---------------------------------------------------------------- Sheet 3 报价体系
ws = wb.create_sheet("3-报价体系")
hr = title_row(ws, "报价体系 · 三档模型与构成明细（单位：万元）", 5,
               "上不封顶原则下的建议基准价；实际以邀约名单/规模确认后核定")
headers = ["费用构成项"] + list(D.TIERS.keys())
widths = [22, 26, 22, 24]
ws.column_dimensions["A"].width = 22
for i, k in enumerate(D.TIERS.keys(), 2):
    ws.column_dimensions[get_column_letter(i)].width = 24
for j, h in enumerate(headers, 1):
    style_header(ws.cell(row=hr, column=j, value=h))
r = hr + 1
for idx, item in enumerate(D.COST_ITEMS):
    style_cell(ws.cell(row=r, column=1, value=item), fill=LTBLUE, bold=True)
    for j, k in enumerate(D.TIERS.keys(), 2):
        style_cell(ws.cell(row=r, column=j, value=D.TIERS[k][idx]), center=True,
                   fill=GREY if r % 2 else WHITE)
    r += 1
style_cell(ws.cell(row=r, column=1, value="单场合计"), bold=True, fill=GOLD, color=WHITE)
for j, k in enumerate(D.TIERS.keys(), 2):
    style_cell(ws.cell(row=r, column=j, value=D.TIER_TOTAL[k]), center=True, bold=True,
               fill=GOLD, color=WHITE)
r += 2
# 场次汇总
ws.cell(row=r, column=1, value="全年场次与预算汇总").font = Font(name=FONT, bold=True, size=12, color=NAVY)
r += 1
sum_head = ["档位", "场次", "单价(万元)", "小计(万元)"]
for j, h in enumerate(sum_head, 1):
    style_header(ws.cell(row=r, column=j, value=h))
r += 1
tier_counts = {}
for a in D.ACTIVITIES:
    tier_counts[a["tier"]] = tier_counts.get(a["tier"], 0) + 1
for k in D.TIERS.keys():
    cnt = tier_counts.get(k, 0)
    sub = round(cnt * D.TIER_TOTAL[k], 1)
    style_cell(ws.cell(row=r, column=1, value=k), fill=GREY)
    style_cell(ws.cell(row=r, column=2, value=cnt), center=True, fill=GREY)
    style_cell(ws.cell(row=r, column=3, value=D.TIER_TOTAL[k]), center=True, fill=GREY)
    style_cell(ws.cell(row=r, column=4, value=sub), center=True, bold=True, fill=GREY)
    r += 1
style_cell(ws.cell(row=r, column=1, value="合计"), bold=True, fill=NAVY, color=WHITE)
style_cell(ws.cell(row=r, column=2, value=len(D.ACTIVITIES)), center=True, bold=True, fill=NAVY, color=WHITE)
style_cell(ws.cell(row=r, column=3, value="—"), center=True, fill=NAVY, color=WHITE)
style_cell(ws.cell(row=r, column=4, value=D.TOTAL_PRICE), center=True, bold=True, fill=NAVY, color=WHITE)
r += 2
drop = round((1 - D.TOTAL_PRICE / D.PREV_TOTAL) * 100)
c = ws.cell(row=r, column=1,
            value=f"说明：按东方枢纽标准优化后，全年合计约 {D.TOTAL_PRICE} 万元，较初版 {D.PREV_TOTAL} 万元下调约 {drop}%；上不封顶原则下的合理基准价，可按规模/邀约确认后微调。")
c.font = Font(name=FONT, italic=True, size=9, color="595959")
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)

# ---------------------------------------------------------------- Sheet 4 招商价值
ws = wb.create_sheet("4-招商引资价值分析")
hr = title_row(ws, "招商引资价值分析 · 针对东方枢纽 A 片区办公项目（约 30 万㎡）", 3,
               "以下数值均为测算假设，用于展示价值逻辑，非承诺性数据")
# 漏斗
ws.column_dimensions["A"].width = 26
ws.column_dimensions["B"].width = 34
ws.column_dimensions["C"].width = 20
ws.cell(row=hr, column=1, value="一、招商转化漏斗").font = Font(name=FONT, bold=True, size=12, color=NAVY)
r = hr + 1
for j, h in enumerate(["转化阶段", "口径 / 假设", "预估数值"], 1):
    style_header(ws.cell(row=r, column=j, value=h))
r += 1
for stage, basis, val in D.FUNNEL:
    style_cell(ws.cell(row=r, column=1, value=stage), bold=True, fill=LTBLUE)
    style_cell(ws.cell(row=r, column=2, value=basis), fill=GREY if r % 2 else WHITE)
    style_cell(ws.cell(row=r, column=3, value=val), center=True, bold=True, fill=GREY if r % 2 else WHITE, color=GOLD)
    r += 1
r += 1
# 去化情景
ws.cell(row=r, column=1, value="二、去化面积测算（A 片区办公约 30 万㎡ × 情景去化率）").font = Font(name=FONT, bold=True, size=12, color=NAVY)
r += 1
for j, h in enumerate(["情景", "去化率假设", "带动去化面积"], 1):
    style_header(ws.cell(row=r, column=j, value=h))
r += 1
for name, rate, area in D.GMV_SCENARIOS:
    style_cell(ws.cell(row=r, column=1, value=name), bold=True, center=True, fill=LTBLUE)
    style_cell(ws.cell(row=r, column=2, value=rate), center=True, fill=GREY if r % 2 else WHITE)
    style_cell(ws.cell(row=r, column=3, value=area), center=True, bold=True, fill=GREY if r % 2 else WHITE, color=GOLD)
    r += 1
r += 1
# 价值支柱
ws.cell(row=r, column=1, value="三、六大价值支柱").font = Font(name=FONT, bold=True, size=12, color=NAVY)
r += 1
for j, h in enumerate(["价值支柱", "说明", ""], 1):
    style_header(ws.cell(row=r, column=j, value=h))
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
r += 1
for name, desc in D.VALUE_PILLARS:
    style_cell(ws.cell(row=r, column=1, value=name), bold=True, fill=LTBLUE)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    style_cell(ws.cell(row=r, column=2, value=desc), fill=GREY if r % 2 else WHITE)
    ws.row_dimensions[r].height = 40
    r += 1
r += 1
# ROI
ws.cell(row=r, column=1, value="四、投入产出概览").font = Font(name=FONT, bold=True, size=12, color=NAVY)
r += 1
for j, h in enumerate(["指标", "数值", ""], 1):
    style_header(ws.cell(row=r, column=j, value=h))
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
r += 1
for name, val in D.ROI_SUMMARY:
    style_cell(ws.cell(row=r, column=1, value=name), bold=True, fill=LTBLUE)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    style_cell(ws.cell(row=r, column=2, value=val), fill=GREY if r % 2 else WHITE, color=GOLD, bold=True)
    r += 1

# ---------------------------------------------------------------- Sheet 5 合作资源与执行机制
ws = wb.create_sheet("5-招商标的与执行机制")
hr = title_row(ws, "招商标的 · 合作资源背书 · 执行机制", 2)
ws.column_dimensions["A"].width = 22
ws.column_dimensions["B"].width = 72
# 招商标的
ws.cell(row=hr, column=1, value="〇、招商标的：" + D.PROJECT["name"]).font = Font(name=FONT, bold=True, size=12, color=NAVY)
r = hr + 1
style_cell(ws.cell(row=r, column=1, value="项目定位"), bold=True, center=True, fill=LTBLUE)
style_cell(ws.cell(row=r, column=2, value=D.PROJECT["position"]), fill=WHITE)
ws.row_dimensions[r].height = 32; r += 1
style_cell(ws.cell(row=r, column=1, value="体量规模"), bold=True, center=True, fill=LTBLUE)
style_cell(ws.cell(row=r, column=2, value=D.PROJECT["area"] + "（说明：招商标的为 A 片区办公项目，非“133 万方”）"), fill=WHITE, color=GOLD, bold=True)
ws.row_dimensions[r].height = 30; r += 1
style_cell(ws.cell(row=r, column=1, value="四大产品线"), bold=True, center=True, fill=LTBLUE)
style_cell(ws.cell(row=r, column=2, value="\n".join(f"• {n}：{d}" for n, d in D.PROJECT["product_lines"])), fill=WHITE)
ws.row_dimensions[r].height = 92; r += 1
style_cell(ws.cell(row=r, column=1, value="销售/租赁模式"), bold=True, center=True, fill=LTBLUE)
style_cell(ws.cell(row=r, column=2, value="\n".join("• " + x for x in D.PROJECT["model"])), fill=WHITE)
ws.row_dimensions[r].height = 66; r += 1
style_cell(ws.cell(row=r, column=1, value="活动↔产品匹配"), bold=True, center=True, fill=GOLD, color=WHITE)
style_cell(ws.cell(row=r, column=2, value=D.PROJECT["match"]), fill=WHITE)
ws.row_dimensions[r].height = 32; r += 2
# 三方
ws.cell(row=r, column=1, value="一、三方合作定位").font = Font(name=FONT, bold=True, size=12, color=NAVY)
r += 1
for name, role, desc in D.PARTIES:
    style_cell(ws.cell(row=r, column=1, value=f"{name}\n（{role}）"), bold=True, center=True, fill=LTBLUE)
    style_cell(ws.cell(row=r, column=2, value=desc), fill=GREY if r % 2 else WHITE)
    ws.row_dimensions[r].height = 48
    r += 1
style_cell(ws.cell(row=r, column=1, value="备选/补充科技组织"), bold=True, center=True, fill=GREY)
style_cell(ws.cell(row=r, column=2, value="、".join(D.ALT_TECH_ORGS)), fill=WHITE)
r += 2
# 政府资源
ws.cell(row=r, column=1, value="二、政府资源背书矩阵（市级 + 浦东新区联动）").font = Font(name=FONT, bold=True, size=12, color=NAVY)
r += 1
for group, items in D.GOV_RESOURCES.items():
    style_cell(ws.cell(row=r, column=1, value=group), bold=True, center=True, fill=BLUE, color=WHITE)
    style_cell(ws.cell(row=r, column=2, value="\n".join("• " + x for x in items)), fill=GREY if r % 2 else WHITE)
    ws.row_dimensions[r].height = 58
    r += 1
style_cell(ws.cell(row=r, column=1, value="联动逻辑"), bold=True, center=True, fill=GOLD, color=WHITE)
style_cell(ws.cell(row=r, column=2, value=D.GOV_TAGLINE), fill=WHITE)
r += 2
# 执行机制
ws.cell(row=r, column=1, value="三、执行机制与合规保障").font = Font(name=FONT, bold=True, size=12, color=NAVY)
r += 1
for j, h in enumerate(["环节", "要点"], 1):
    style_header(ws.cell(row=r, column=j, value=h))
r += 1
for name, desc in D.EXECUTION:
    style_cell(ws.cell(row=r, column=1, value=name), bold=True, center=True, fill=LTBLUE)
    style_cell(ws.cell(row=r, column=2, value=desc), fill=GREY if r % 2 else WHITE)
    ws.row_dimensions[r].height = 32
    r += 1

for ws in wb.worksheets:
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.4

wb.save("东方枢纽三方合作计划_明细表.xlsx")
print("Excel 已生成：东方枢纽三方合作计划_明细表.xlsx  (工作表数: %d)" % len(wb.sheetnames))
print("Sheets:", wb.sheetnames)
