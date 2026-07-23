# -*- coding: utf-8 -*-
"""生成《复兴岛全球创客岛收官答卷大会》执行计划 Excel（主呈现台账）。"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import content as C

OUT = Path(__file__).resolve().parent.parent / "deliverables"
OUT.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT / "复兴岛_全球创客岛_科创出海与具身智能国际大会_执行计划表.xlsx"

FONT = "Microsoft YaHei"
NAVY = "0B1F3A"
TEAL = "0A6E6A"
GOLD = "C48A2A"
LIGHT = "E6F0EE"
ROW2 = "F3F7F6"
WHITE = "FFFFFF"
INK = "1B2A44"
RED = "B33A3A"

thin = Side(style="thin", color="D5DEEB")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def hfill(color):
    return PatternFill("solid", fgColor=color)


def style_title(ws, text, ncol, color=TEAL):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    c = ws.cell(1, 1, text)
    c.font = Font(name=FONT, size=14, bold=True, color=WHITE)
    c.fill = hfill(color)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30


def style_sub(ws, text, ncol, row=2):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncol)
    c = ws.cell(row, 1, text)
    c.font = Font(name=FONT, size=10, italic=True, color="5A6B86")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 20


def header_row(ws, row, headers, widths, color=NAVY):
    for j, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row, j, h)
        c.font = Font(name=FONT, size=10.5, bold=True, color=WHITE)
        c.fill = hfill(color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[row].height = 26


def data_rows(ws, start, rows, aligns=None, base_h=24):
    for i, r in enumerate(rows):
        ri = start + i
        for j, val in enumerate(r, start=1):
            c = ws.cell(ri, j, val)
            c.font = Font(name=FONT, size=10, color=INK)
            c.fill = hfill(WHITE if i % 2 == 0 else ROW2)
            al = aligns[j - 1] if aligns else "left"
            c.alignment = Alignment(
                horizontal=al,
                vertical="center",
                wrap_text=True,
                indent=(1 if al == "left" else 0),
            )
            c.border = BORDER
        ws.row_dimensions[ri].height = base_h
    return start + len(rows)


def freeze(ws, row=5):
    ws.freeze_panes = ws.cell(row, 1)
    ws.print_title_rows = "1:4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


wb = Workbook()

# ========== 0. 总览仪表盘 ==========
ws = wb.active
ws.title = "0.总览仪表盘"
style_title(ws, f"{C.PROJECT_NAME}｜Excel 执行呈现台账", 4, NAVY)
style_sub(ws, f"{C.PROJECT_FULL}｜{C.VERSION}｜建议日期 {C.EVENT_DATE}", 4)

# KPI 卡区标题
ws.merge_cells("A4:D4")
c = ws.cell(4, 1, "一、决策速览（给领导看的关键数字）")
c.font = Font(name=FONT, size=12, bold=True, color=WHITE)
c.fill = hfill(TEAL)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[4].height = 24

kpi_cards = [
    ("建议日期", "2026-09-12（六）", "9/15前最近开市/挂匾吉日"),
    ("活动规模", "200–300 人", "岛上主场，不离岛"),
    ("国际规格", "总领事 ≥6 国", "一带一路亲自出席"),
    ("硬成果落位", "揭牌 ≥3 个", "国家厅+片区厅"),
    ("预算中值", "约 55 万元", "场地置换可压至约40万"),
    ("备选日期", "2026-09-09（三）", "工作日开市吉日"),
]
header_row(ws, 5, ["指标", "数值", "说明", "备注"], [18, 28, 36, 36], NAVY)
card_rows = []
for a, b, cnote in kpi_cards:
    card_rows.append([a, b, cnote, "见对应工作表"])
data_rows(ws, 6, card_rows, aligns=["center", "center", "left", "left"], base_h=26)

ws.merge_cells("A13:D13")
c = ws.cell(13, 1, "二、工作表导航（本文件即完整执行呈现）")
c.font = Font(name=FONT, size=12, bold=True, color=WHITE)
c.fill = hfill(GOLD)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

nav = [
    ["1.活动总览", "名称/时间/地点/主题/规格一览", "汇报封面信息"],
    ["2.择日专章", "首选/备选/规避与黄历宜忌", "日期决策依据"],
    ["3.详细议程", "全天环节、内容、责任方", "现场执行脚本"],
    ["4.嘉宾邀请", "政要/领事/产业/媒体分层", "邀约任务拆解"],
    ["5.一带一路国别池", "拟邀国家与合作侧重", "国际资源清单"],
    ["6.揭牌落位", "国家厅/片区厅/产业平台", "硬成果台账"],
    ["7.席位结构", "200–300 人席别配置", "会务排座依据"],
    ["8.组织架构", "六组专班职责", "分工责任书"],
    ["9.倒排期", "T-50 至 T+14 节点任务", "项目进度表"],
    ["10.预算测算", "分项费用区间", "资金安排"],
    ["11.KPI成效", "规格/落位/传播/转化指标", "验收标准"],
    ["12.风险预案", "风险·对策·等级", "应急管理"],
    ["13.下一步行动", "本周拍板五件事", "立即执行清单"],
]
header_row(ws, 14, ["工作表", "内容", "用途", "优先级"], [22, 40, 22, 14], NAVY)
nav_rows = [[a, b, c, "高" if i < 6 else "中"] for i, (a, b, c) in enumerate(nav)]
data_rows(ws, 15, nav_rows, aligns=["left", "left", "left", "center"], base_h=24)

ws.merge_cells("A29:D29")
c = ws.cell(29, 1, "三、一句话定位")
c.font = Font(name=FONT, size=12, bold=True, color=WHITE)
c.fill = hfill(TEAL)
ws.merge_cells("A30:D31")
c = ws.cell(30, 1, C.ONE_LINER)
c.font = Font(name=FONT, size=11, color=INK)
c.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
ws.row_dimensions[30].height = 36
ws.row_dimensions[31].height = 20

for j, w in enumerate([22, 40, 22, 36], 1):
    ws.column_dimensions[get_column_letter(j)].width = w
ws.freeze_panes = "A5"

# ========== 1 活动总览 ==========
ws = wb.create_sheet("1.活动总览")
style_title(ws, f"{C.PROJECT_NAME}｜活动总览", 4)
style_sub(ws, "PPT 封面信息的表格化呈现｜可用于请示附件", 4)
hr = 4
header_row(ws, hr, ["项目", "内容", "项目", "内容"], [18, 42, 18, 42])
pairs = []
ov = list(C.OVERVIEW)
for i in range(0, len(ov), 2):
    a = ov[i]
    b = ov[i + 1] if i + 1 < len(ov) else ("", "")
    pairs.append((a[0], a[1], b[0], b[1]))
data_rows(ws, hr + 1, pairs, aligns=["center", "left", "center", "left"], base_h=34)
freeze(ws)

# ========== 2 择日 ==========
ws = wb.create_sheet("2.择日专章")
style_title(ws, "择日专章｜9 月 15 日前最靠近收官节点的黄道吉日", 4, GOLD)
style_sub(ws, "主推 9/12；领导周六不便则改 9/9；规避 9/10 杨公忌日", 4)
hr = 4
header_row(ws, hr, ["类型", "日期", "宜/要点", "决策建议"], [12, 36, 48, 28], TEAL)
rows = [
    ["首选", C.HUANGLI["date"], "开市·挂匾·立券·交易·出行", "主推，最靠近 9/15"],
    ["备选一", C.DATE_BACKUP[0]["date"], "开市·交易·立券·出行", "工作日备选"],
    ["备选二", C.DATE_BACKUP[1]["date"], "金匮黄道日·开市·立券", "预热/分论坛"],
    ["规避", "2026-09-10（周四）", "杨公忌日·大事勿用", "不安排主会场"],
    ["越过节点", "2026-09-16（周三）", "虽为宜开市日", "不作首选"],
]
data_rows(ws, hr + 1, rows, base_h=30)
r = hr + 1 + len(rows) + 1
ws.cell(r, 1, "黄历明细").font = Font(name=FONT, size=11, bold=True, color=TEAL)
detail = [
    ["干支/值日", f"{C.HUANGLI['ganzhi']} · {C.HUANGLI['zhiri']}", "", ""],
    ["冲煞", C.HUANGLI["chong"], "", ""],
    ["宜（全文）", C.HUANGLI["yi"], "", ""],
    ["忌", C.HUANGLI["ji"], "", ""],
    ["择日理由", C.HUANGLI["why"], "", ""],
]
data_rows(ws, r + 1, detail, base_h=36)
freeze(ws)

# ========== 3 议程 ==========
ws = wb.create_sheet("3.详细议程")
style_title(ws, f"详细议程｜{C.EVENT_DATE}", 6)
style_sub(ws, "上午：规格+揭牌签约；下午：分论坛+国际对接｜状态列供现场勾选", 6)
hr = 4
header_row(ws, hr, ["时段", "时间", "环节", "内容要点", "责任方", "状态"], [10, 16, 26, 42, 14, 10])
rows = []
for t, name, detail, owner in C.AGENDA:
    period = "上午" if t[:2] in ("08", "09", "10", "11", "12") else "下午"
    rows.append([period, t, name, detail, owner, "待执行"])
data_rows(ws, hr + 1, rows, aligns=["center", "center", "left", "left", "center", "center"], base_h=28)
freeze(ws)

# ========== 4 嘉宾 ==========
ws = wb.create_sheet("4.嘉宾邀请")
style_title(ws, "嘉宾分层与邀约目标", 5)
style_sub(ws, "政治高位与国际高位必须同时在场｜可增列具体姓名后跟踪", 5)
hr = 4
header_row(ws, hr, ["层级", "邀约对象", "目标", "负责人", "确认状态"], [18, 55, 28, 12, 12])
rows = [[g["tier"], "；".join(g["targets"]), g["goal"], "", "待启动"] for g in C.GUEST_TIERS]
data_rows(ws, hr + 1, rows, base_h=52)
freeze(ws)

# ========== 5 国别池 ==========
ws = wb.create_sheet("5.一带一路国别池")
style_title(ws, "拟邀「一带一路」国家参考池（揭牌优先）", 6, TEAL)
style_sub(ws, "建议确认 6–10 国总领事出席；揭牌国须本人到场", 6)
hr = 4
header_row(ws, hr, ["序号", "国家", "合作侧重", "是否揭牌候选", "领事确认", "备注"], [8, 16, 40, 14, 12, 20])
rows = [[i + 1, n, f, "是" if i < 3 else "待定", "待邀约", ""] for i, (n, f) in enumerate(C.BRI_COUNTRY_POOL)]
data_rows(ws, hr + 1, rows, aligns=["center", "center", "left", "center", "center", "left"], base_h=26)
freeze(ws)

# ========== 6 揭牌 ==========
ws = wb.create_sheet("6.揭牌落位")
style_title(ws, "国际会议厅 / 会客厅揭牌落位计划", 6, GOLD)
style_sub(ws, "先锁意向、再上仪式；一国一厅、一厅一责", 6)
hr = 4
header_row(ws, hr, ["落位类型", "数量建议", "形式", "价值", "意向对象", "协议状态"], [28, 14, 32, 32, 16, 12])
rows = [[u["name"], u["count"], u["form"], u["value"], "", "谈判中"] for u in C.UNVEILING]
data_rows(ws, hr + 1, rows, base_h=42)
r = hr + 1 + len(rows) + 1
ws.cell(r, 1, "落位原则").font = Font(name=FONT, size=11, bold=True, color=TEAL)
for i, p in enumerate(C.UNVEILING_PRINCIPLES):
    cell = ws.cell(r + 1 + i, 1, f"· {p}")
    cell.font = Font(name=FONT, size=10, color=INK)
    ws.merge_cells(start_row=r + 1 + i, start_column=1, end_row=r + 1 + i, end_column=6)
freeze(ws)

# ========== 7 席位 ==========
ws = wb.create_sheet("7.席位结构")
style_title(ws, "席位结构（200–300 人）", 4)
style_sub(ws, "核心席质量优先于总人数冲高", 4)
hr = 4
header_row(ws, hr, ["席别", "人数", "组成", "签到通道"], [28, 14, 55, 14])
rows = [[a, b, c, "贵宾" if i < 2 else "标准"] for i, (a, b, c) in enumerate(C.SEAT_PLAN)]
data_rows(ws, hr + 1, rows, aligns=["left", "center", "left", "center"], base_h=28)
freeze(ws)

# ========== 8 组织 ==========
ws = wb.create_sheet("8.组织架构")
style_title(ws, "组织架构与职责", 4)
style_sub(ws, "建议成立六组专班，明确组长与值班机制", 4)
hr = 4
header_row(ws, hr, ["组别", "职责", "组长（待填）", "成员（待填）"], [18, 60, 14, 20])
rows = [[a, b, "", ""] for a, b in C.ORG_STRUCTURE]
data_rows(ws, hr + 1, rows, base_h=30)
freeze(ws)

# ========== 9 倒排期 ==========
ws = wb.create_sheet("9.倒排期")
style_title(ws, "执行倒排期（自即日起）", 5, TEAL)
style_sub(ws, "外事报批建议按 ≥40 天窗口提前启动｜状态列供周会更新", 5)
hr = 4
header_row(ws, hr, ["序号", "节点", "关键任务", "责任组", "状态"], [8, 26, 70, 14, 12])
owner_cycle = ["综合协调组", "外事礼宾组", "空间落位组", "产业内容组", "会务保障组", "宣传媒体组"]
rows = []
for i, (a, b) in enumerate(C.TIMELINE):
    rows.append([i + 1, a, b, owner_cycle[i % len(owner_cycle)], "待启动"])
data_rows(ws, hr + 1, rows, aligns=["center", "left", "left", "center", "center"], base_h=32)
freeze(ws)

# ========== 10 预算 ==========
ws = wb.create_sheet("10.预算测算")
style_title(ws, "预算测算（示意，单位：万元）", 5, GOLD)
style_sub(ws, C.BUDGET_NOTE, 5)
hr = 4
header_row(ws, hr, ["成本项", "预算低值", "预算高值", "中值参考", "说明"], [32, 12, 12, 12, 45])
rows = []
for a, b, c in C.BUDGET:
    if "–" in str(b) or "-" in str(b):
        parts = str(b).replace("–", "-").split("-")
        low, high = parts[0].strip(), parts[-1].strip()
        try:
            mid = f"{(float(low) + float(high.split('（')[0]) ) / 2:.0f}"
        except Exception:
            mid = b
    else:
        low = high = mid = b
    rows.append([a, low, high, mid, c])
data_rows(ws, hr + 1, rows, aligns=["left", "center", "center", "center", "left"], base_h=28)
freeze(ws)

# ========== 11 KPI ==========
ws = wb.create_sheet("11.KPI成效")
style_title(ws, "预期成效与 KPI（验收标准）", 4)
style_sub(ws, "活动结束后按此表写专报与复盘", 4)
hr = 4
header_row(ws, hr, ["维度", "量化目标", "实测结果（待填）", "是否达标"], [16, 70, 20, 12])
rows = [[a, b, "", ""] for a, b in C.KPIS]
data_rows(ws, hr + 1, rows, base_h=32)
freeze(ws)

# ========== 12 风险 ==========
ws = wb.create_sheet("12.风险预案")
style_title(ws, "风险预案", 4, RED)
style_sub(ws, "高风险项须在 T-30 前关闭或降级", 4)
hr = 4
header_row(ws, hr, ["风险", "对策", "等级", "跟踪人"], [30, 60, 10, 12])
rows = [[a, b, c, ""] for a, b, c in C.RISKS]
data_rows(ws, hr + 1, rows, aligns=["left", "left", "center", "center"], base_h=36)
freeze(ws)

# ========== 13 下一步 ==========
ws = wb.create_sheet("13.下一步行动")
style_title(ws, "下一步行动清单（本周拍板）", 4, TEAL)
style_sub(ws, "建议与 PPT 第 15 页决策清单同步使用", 4)
hr = 4
header_row(ws, hr, ["序号", "行动项", "完成时限", "状态"], [8, 70, 16, 12])
limits = ["本周内", "3 日内", "1 周内", "10 日内", "同步推进"]
rows = [[i + 1, t, limits[min(i, len(limits) - 1)], "待启动"] for i, t in enumerate(C.NEXT_STEPS)]
data_rows(ws, hr + 1, rows, aligns=["center", "left", "center", "center"], base_h=32)
freeze(ws)

# ========== 14 主题板块（补充）==========
ws = wb.create_sheet("14.主题板块")
style_title(ws, "四大主题板块协同要点", 3)
style_sub(ws, "出海为主轴，AI / 具身智能 / 低空经济板块协同", 3)
hr = 4
header_row(ws, hr, ["板块", "英文标签", "要点"], [22, 18, 70])
rows = []
for p in C.THEME_PILLARS:
    rows.append([p["name"], p["tag"], "；".join(p["points"])])
data_rows(ws, hr + 1, rows, base_h=48)
freeze(ws)

wb.save(OUT_FILE)
print(f"已生成：{OUT_FILE}（{len(wb.sheetnames)} 张工作表）")
