# -*- coding: utf-8 -*-
"""生成《复兴岛全球创客岛收官答卷大会》执行计划 Excel（V3 · 敦煌+四层结构）。"""
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
                horizontal=al, vertical="center", wrap_text=True,
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

# 0 仪表盘
ws = wb.active
ws.title = "0.总览仪表盘"
style_title(ws, f"{C.PROJECT_NAME}｜{C.PROJECT_FULL}", 4, NAVY)
style_sub(ws, f"{C.VERSION}｜杨浦AI企业互动｜四层结构｜敦煌意向｜主会日 {C.EVENT_DATE_SHORT}", 4)

ws.merge_cells("A4:D4")
c = ws.cell(4, 1, "一、决策速览")
c.font = Font(name=FONT, size=12, bold=True, color=WHITE)
c.fill = hfill(TEAL)
header_row(ws, 5, ["指标", "数值", "说明", "备注"], [18, 36, 40, 28], NAVY)
cards = [
    ("建议日期", "2026-09-12（六）", "服务9/15收官节点", "见择日专章"),
    ("活动结构", "四层组合", "主论坛+分论坛+体验层+黑客松", "见活动结构"),
    ("敦煌意向", "2000㎡/净高9–11m", "壁画IP·沉浸·机器人·待核实950万", "见敦煌项目"),
    ("黑客松", "9/11晚–9/12下午", "船台·智能体24h+文旅赛道", "见黑客松专章"),
    ("活动规模", "主会200–300人", "黑客松20–40队+体验层流动", "见席位结构"),
    ("国际规格", "总领事≥6国", "一带一路亲自出席", "见嘉宾邀请"),
    ("硬成果", "揭牌≥3个", "国家厅+片区厅+敦煌视进度", "见揭牌落位"),
    ("角色原则", "智库/顾问", "不承诺政策·一事一议·先调研", "见冷启动角色"),
]
data_rows(ws, 6, cards, aligns=["center", "center", "left", "left"], base_h=24)

ws.merge_cells("A15:D15")
c = ws.cell(15, 1, "二、工作表导航")
c.font = Font(name=FONT, size=12, bold=True, color=WHITE)
c.fill = hfill(GOLD)
nav = [
    ["1.活动总览", "名称/时间/结构/规格", "汇报封面"],
    ["2.活动结构", "四层结构与组合逻辑", "形式设计"],
    ["3.历史文脉", "时间轴·载体·衍生", "文脉依据"],
    ["4.领导寄托", "市/区要点与回应", "政治口径"],
    ["5.规划与课题", "规划对齐+六课题", "议题设计"],
    ["6.敦煌项目", "意向/空间/核实/用法", "一事一议样本"],
    ["7.冷启动角色", "原则+协同方", "对接纪律"],
    ["8.杨浦企业互动", "优刻得/智谱/苏度等", "企业邀约"],
    ["9.择日专章", "首选/备选/规避", "日期决策"],
    ["10.详细议程", "含杨浦AI力量环节", "现场脚本"],
    ["11.体验游戏层", "五关卡集章任务", "体验执行"],
    ["12.黑客松专章", "赛道/赛程/规则/KPI", "创客执行"],
    ["13.嘉宾邀请", "政要/领事/杨浦企业", "邀约台账"],
    ["14.一带一路国别池", "拟邀国家", "国际资源"],
    ["15.揭牌落位", "会客厅+敦煌意向", "硬成果"],
    ["16.席位结构", "主会席别", "会务排座"],
    ["17.组织架构", "含智库/文旅/黑客松", "分工责任"],
    ["18.倒排期", "节点任务", "进度表"],
    ["19.预算测算", "分项费用", "资金安排"],
    ["20.KPI成效", "验收标准", "复盘"],
    ["21.风险预案", "风险对策", "应急"],
    ["22.下一步行动", "企业锁定+敦煌核查", "立即执行"],
]
header_row(ws, 16, ["工作表", "内容", "用途", "优先级"], [22, 40, 22, 12], NAVY)
data_rows(ws, 17, [[a, b, c, "高" if i < 11 else "中"] for i, (a, b, c) in enumerate(nav)],
          aligns=["left", "left", "left", "center"], base_h=20)

ws.merge_cells("A39:D40")
c = ws.cell(39, 1, C.ONE_LINER)
c.font = Font(name=FONT, size=11, color=INK)
c.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
ws.row_dimensions[39].height = 48
for j, w in enumerate([22, 40, 22, 28], 1):
    ws.column_dimensions[get_column_letter(j)].width = w
ws.freeze_panes = "A5"

# 1 总览
ws = wb.create_sheet("1.活动总览")
style_title(ws, f"{C.PROJECT_NAME}｜活动总览", 4)
style_sub(ws, "PPT 封面信息的表格化呈现", 4)
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

# 2 活动结构
ws = wb.create_sheet("2.活动结构")
style_title(ws, C.EVENT_FORMAT["title"], 3, TEAL)
style_sub(ws, "主论坛定调 · 分论坛深耕 · 体验层引流 · 黑客松出苗", 3)
hr = 4
header_row(ws, hr, ["层级", "说明", "关键产出"], [22, 70, 28], NAVY)
outputs = ["领导画面/揭牌签约", "专业纪要/对接清单", "传播素材/年轻流量", "原型/人才线索"]
rows = [[a, b, outputs[i]] for i, (a, b) in enumerate(C.EVENT_FORMAT["layers"])]
data_rows(ws, hr + 1, rows, base_h=40)
r = hr + 1 + len(rows) + 1
header_row(ws, r, ["组合逻辑", "说明", "备注"], [16, 70, 20], GOLD)
data_rows(ws, r + 1, [[f"逻辑{i+1}", x, ""] for i, x in enumerate(C.EVENT_FORMAT["combo_logic"])], base_h=28)
freeze(ws)

# 3 历史文脉
ws = wb.create_sheet("3.历史文脉")
style_title(ws, "历史文脉时间轴 · 载体 · 衍生议题", 3, GOLD)
style_sub(ws, "周家嘴浅滩→定海岛→复兴岛→全球创客岛｜含敦煌衍生", 3)
hr = 4
header_row(ws, hr, ["时期", "节点", "要点"], [16, 18, 80], TEAL)
data_rows(ws, hr + 1, list(C.HISTORY_TIMELINE), base_h=28)
r = hr + 1 + len(C.HISTORY_TIMELINE) + 1
ws.cell(r, 1, "文脉载体").font = Font(name=FONT, size=11, bold=True, color=TEAL)
header_row(ws, r + 1, ["载体", "说明", "活动用法"], [28, 55, 30], NAVY)
usage = ["视察打卡/室外合影", "开幕影像符号", "展廊节点", "黑客松与分论坛场地", "低空/无人艇叙事"]
rows = [[n, d, usage[i] if i < len(usage) else "展陈"] for i, (n, d) in enumerate(C.HISTORY_HERITAGE)]
data_rows(ws, r + 2, rows, base_h=36)
r2 = r + 2 + len(rows) + 1
ws.cell(r2, 1, "衍生议题").font = Font(name=FONT, size=11, bold=True, color=GOLD)
header_row(ws, r2 + 1, ["衍生方向", "要点1", "要点2"], [36, 40, 40], NAVY)
drows = [[d["name"], d["points"][0], d["points"][1]] for d in C.HISTORY_DERIVATIVES]
data_rows(ws, r2 + 2, drows, base_h=40)
freeze(ws)

# 4 领导寄托
ws = wb.create_sheet("4.领导寄托")
style_title(ws, "市、区领导寄托与活动回应", 3)
style_sub(ws, "把领导讲话转译为可执行的活动模块", 3)
hr = 4
header_row(ws, hr, ["来源", "要点", "本场回应模块"], [36, 55, 30], NAVY)
rows = []
for block in C.LEADERSHIP:
    for p in block["points"]:
        rows.append([block["who"], p, ""])
for i, row in enumerate(rows):
    row[2] = C.LEADERSHIP_RESPONSE[i % len(C.LEADERSHIP_RESPONSE)]
data_rows(ws, hr + 1, rows, base_h=32)
freeze(ws)

# 5 规划课题
ws = wb.create_sheet("5.规划与课题")
style_title(ws, "大上海 / 杨浦规划对齐与六大课题", 3, TEAL)
style_sub(ws, "杨府发〔2026〕1号 · 量子城市 · 规划资源指导意见 · 含敦煌课题", 3)
hr = 4
header_row(ws, hr, ["规划维度", "对齐要点", "活动落点"], [18, 70, 28], NAVY)
落地 = ["主旨发布", "成果专章", "政策发布", "产业分论坛", "船台动线", "人民城市展陈"]
rows = [[a, b, 落地[i] if i < len(落地) else "议题"] for i, (a, b) in enumerate(C.PLANNING_ALIGN)]
data_rows(ws, hr + 1, rows, base_h=34)
r = hr + 1 + len(rows) + 1
header_row(ws, r, ["课题编号", "题目", "聚焦内容"], [12, 28, 60], GOLD)
data_rows(ws, r + 1, list(C.TOPIC_AGENDA), aligns=["center", "left", "left"], base_h=30)
freeze(ws)

# 6 敦煌项目
ws = wb.create_sheet("6.敦煌项目")
style_title(ws, C.DUNHUANG["name"], 4, GOLD)
style_sub(ws, C.DUNHUANG["status"], 4)
hr = 4
header_row(ws, hr, ["类别", "内容", "核实状态", "负责人"], [14, 70, 14, 14], TEAL)
rows = [["基本面", x, "待核实" if "950" in x or "待核实" in x else "已同步", "文旅沉浸组"]
        for x in C.DUNHUANG["basics"]]
rows += [["契合点", x, "共识", "智库顾问组"] for x in C.DUNHUANG["fit"]]
rows += [["活动用法", x, "待脚本", "产业内容组"] for x in C.DUNHUANG["event_use"]]
rows += [["下一步", x, "待启动", "生态转化组"] for x in C.DUNHUANG["next"]]
data_rows(ws, hr + 1, rows, aligns=["center", "left", "center", "center"], base_h=32)
freeze(ws)

# 7 冷启动角色
ws = wb.create_sheet("7.冷启动角色")
style_title(ws, "冷启动原则与协同角色", 3)
style_sub(ws, "智库/顾问定位 · 一事一议 · 先调研后成交 · 不直接承诺政策", 3)
hr = 4
header_row(ws, hr, ["序号", "原则", "执行要点"], [8, 70, 30], NAVY)
data_rows(ws, hr + 1, [[i + 1, x, "对接纪律"] for i, x in enumerate(C.COLD_START)],
          aligns=["center", "left", "center"], base_h=30)
r = hr + 1 + len(C.COLD_START) + 1
header_row(ws, r, ["角色方", "价值与职责", "备注"], [28, 70, 16], GOLD)
data_rows(ws, r + 1, [[a, b, ""] for a, b in C.PARTNERS], base_h=40)
freeze(ws)

# 8 杨浦企业互动
ws = wb.create_sheet("8.杨浦企业互动")
style_title(ws, "杨浦 AI / 具身智能优质企业互动清单", 5, TEAL)
style_sub(ws, "优刻得·智谱·苏度·卓益得等上岛同台｜主论坛发言+分论坛+体验层+黑客松导师", 5)
hr = 4
header_row(ws, hr, ["企业", "赛道", "杨浦关联", "活动互动角色", "邀约状态"], [18, 14, 32, 40, 12], NAVY)
rows = [[a, b, c, d, "待邀约"] for a, b, c, d in C.YANGPU_ENTERPRISES]
data_rows(ws, hr + 1, rows, base_h=36)
r = hr + 1 + len(rows) + 1
header_row(ws, r, ["互动机制", "说明", "责任方"], [28, 60, 18], GOLD)
data_rows(ws, r + 1, [[a, b, "科企联+产业内容组"] for a, b in C.ENTERPRISE_INTERACTION], base_h=28)
freeze(ws)

# 9 择日
ws = wb.create_sheet("9.择日专章")
style_title(ws, "择日专章｜9/15 前最靠近收官节点的黄道吉日", 4, GOLD)
style_sub(ws, "主推9/12；备选9/9；规避9/10杨公忌日；黑客松绑定周五晚开营", 4)
hr = 4
header_row(ws, hr, ["类型", "日期", "宜/要点", "决策建议"], [12, 36, 48, 28], TEAL)
rows = [
    ["首选", C.HUANGLI["date"], "开市·挂匾·立券·交易·出行", "主推，服务9/15"],
    ["黑客松窗口", C.HACKATHON_WINDOW, "创客夜→答卷日", "与主会绑定"],
    ["备选一", C.DATE_BACKUP[0]["date"], "开市·交易·立券·出行", "工作日备选"],
    ["备选二", C.DATE_BACKUP[1]["date"], "金匮黄道日", "预热/路演"],
    ["规避", "2026-09-10", "杨公忌日", "不安排主会"],
]
data_rows(ws, hr + 1, rows, base_h=30)
freeze(ws)

# 10 议程
ws = wb.create_sheet("10.详细议程")
style_title(ws, f"详细议程｜{C.EVENT_DATE}", 6)
style_sub(ws, "含「杨浦AI力量」企业环节 + 四分论坛 + 体验层 + Demo Day", 6)
hr = 4
header_row(ws, hr, ["时段", "时间", "环节", "内容要点", "责任方", "状态"], [10, 16, 26, 42, 14, 10])
rows = []
for t, name, detail, owner in C.AGENDA:
    period = "上午" if t[:2] in ("08", "09", "10", "11", "12") else "下午"
    rows.append([period, t, name, detail, owner, "待执行"])
data_rows(ws, hr + 1, rows, aligns=["center", "center", "left", "left", "center", "center"], base_h=28)
freeze(ws)

# 11 体验游戏层
ws = wb.create_sheet("11.体验游戏层")
style_title(ws, "体验 / 游戏层关卡设计（论坛×游戏化体验）", 4, TEAL)
style_sub(ws, "非电竞：沉浸关卡+集章任务；苏度/卓益得等真机互动+敦煌关卡", 4)
hr = 4
header_row(ws, hr, ["关卡", "内容", "责任组", "状态"], [22, 60, 16, 12], NAVY)
data_rows(ws, hr + 1,
          [[a, b, "文旅沉浸组" if "敦煌" in a or "沉浸" in a else ("产业内容组" if "智能体" in a else "文脉展示组"), "待脚本"]
           for a, b in C.EXPERIENCE_LAYER],
          aligns=["left", "left", "center", "center"], base_h=30)
freeze(ws)

# 12 黑客松
ws = wb.create_sheet("12.黑客松专章")
style_title(ws, f"{C.HACKATHON['name']}｜执行台账", 4, GOLD)
style_sub(ws, f"{C.HACKATHON['slogan']}｜{C.HACKATHON['window']}｜{C.HACKATHON['scale']}", 4)
hr = 4
header_row(ws, hr, ["类别", "内容", "负责人", "状态"], [14, 70, 14, 12], TEAL)
rows = [["为何举办", x, "黑客松组", "待启动"] for x in C.HACKATHON["why"]]
rows += [["赛道", f"{a}：{b}", "黑客松组", "待发布"] for a, b in C.HACKATHON["tracks"]]
rows += [["规则", x, "黑客松组", "待确认"] for x in C.HACKATHON["rules"]]
rows += [["赛程", f"{a} {b}", "黑客松组", "待执行"] for a, b in C.HACKATHON["schedule"]]
rows += [["KPI", x, "黑客松组", "待验收"] for x in C.HACKATHON["kpis"]]
rows.append(["场地", C.HACKATHON["venue"], "会务保障组", "待锁定"])
data_rows(ws, hr + 1, rows, aligns=["center", "left", "center", "center"], base_h=28)
freeze(ws)

# 13 嘉宾
ws = wb.create_sheet("13.嘉宾邀请")
style_title(ws, "嘉宾分层与邀约目标（含杨浦企业主力军）", 5)
hr = 4
header_row(ws, hr, ["层级", "邀约对象", "目标", "负责人", "确认状态"], [22, 52, 28, 12, 12])
rows = [[g["tier"], "；".join(g["targets"]), g["goal"], "", "待启动"] for g in C.GUEST_TIERS]
data_rows(ws, hr + 1, rows, base_h=52)
freeze(ws)

# 14 国别
ws = wb.create_sheet("14.一带一路国别池")
style_title(ws, "拟邀「一带一路」国家参考池", 6, TEAL)
hr = 4
header_row(ws, hr, ["序号", "国家", "合作侧重", "是否揭牌候选", "领事确认", "备注"], [8, 16, 40, 14, 12, 20])
rows = [[i + 1, n, f, "是" if i < 3 else "待定", "待邀约", ""] for i, (n, f) in enumerate(C.BRI_COUNTRY_POOL)]
data_rows(ws, hr + 1, rows, aligns=["center", "center", "left", "center", "center", "left"], base_h=26)
freeze(ws)

# 15 揭牌
ws = wb.create_sheet("15.揭牌落位")
style_title(ws, "国际会议厅 / 会客厅揭牌落位计划（含敦煌视进度）", 6, GOLD)
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

# 16 席位
ws = wb.create_sheet("16.席位结构")
style_title(ws, "席位结构（主会场 200–300 人）", 4)
hr = 4
header_row(ws, hr, ["席别", "人数", "组成", "签到通道"], [28, 14, 55, 14])
rows = [[a, b, c, "贵宾" if i < 2 else ("创客" if "黑客" in a or "创客" in a else "标准")]
        for i, (a, b, c) in enumerate(C.SEAT_PLAN)]
data_rows(ws, hr + 1, rows, aligns=["left", "center", "left", "center"], base_h=28)
freeze(ws)

# 17 组织
ws = wb.create_sheet("17.组织架构")
style_title(ws, "组织架构与职责（含智库/文旅/黑客松）", 4)
hr = 4
header_row(ws, hr, ["组别", "职责", "组长（待填）", "成员（待填）"], [18, 60, 14, 20])
rows = [[a, b, "", ""] for a, b in C.ORG_STRUCTURE]
data_rows(ws, hr + 1, rows, base_h=30)
freeze(ws)

# 18 倒排
ws = wb.create_sheet("18.倒排期")
style_title(ws, "执行倒排期", 5, TEAL)
hr = 4
header_row(ws, hr, ["序号", "节点", "关键任务", "责任组", "状态"], [8, 22, 70, 14, 12])
owners = ["综合协调组", "文旅沉浸组", "外事礼宾组", "空间落位组", "黑客松组",
          "产业内容组", "会务保障组", "宣传媒体组", "黑客松组", "综合协调组", "宣传媒体组", "文旅沉浸组"]
rows = [[i + 1, a, b, owners[i % len(owners)], "待启动"] for i, (a, b) in enumerate(C.TIMELINE)]
data_rows(ws, hr + 1, rows, aligns=["center", "left", "left", "center", "center"], base_h=30)
freeze(ws)

# 19 预算
ws = wb.create_sheet("19.预算测算")
style_title(ws, "预算测算（示意，单位：万元）", 5, GOLD)
style_sub(ws, C.BUDGET_NOTE, 5)
hr = 4
header_row(ws, hr, ["成本项", "预算低值", "预算高值", "中值参考", "说明"], [32, 12, 12, 12, 45])
rows = []
for a, b, c in C.BUDGET:
    s = str(b).replace("–", "-")
    if "-" in s and "约" not in s.split("-")[0]:
        parts = s.split("-")
        low, high = parts[0].strip(), parts[-1].strip()
        try:
            mid = f"{(float(low) + float(high.split('（')[0])) / 2:.0f}"
        except Exception:
            mid = b
    else:
        low = high = mid = b
    rows.append([a, low, high, mid, c])
data_rows(ws, hr + 1, rows, aligns=["left", "center", "center", "center", "left"], base_h=28)
freeze(ws)

# 20 KPI
ws = wb.create_sheet("20.KPI成效")
style_title(ws, "预期成效与 KPI", 4)
hr = 4
header_row(ws, hr, ["维度", "量化目标", "实测结果（待填）", "是否达标"], [14, 70, 20, 12])
data_rows(ws, hr + 1, [[a, b, "", ""] for a, b in C.KPIS], base_h=30)
freeze(ws)

# 21 风险
ws = wb.create_sheet("21.风险预案")
style_title(ws, "风险预案", 4, RED)
hr = 4
header_row(ws, hr, ["风险", "对策", "等级", "跟踪人"], [30, 60, 10, 12])
data_rows(ws, hr + 1, [[a, b, c, ""] for a, b, c in C.RISKS],
          aligns=["left", "left", "center", "center"], base_h=34)
freeze(ws)

# 22 下一步
ws = wb.create_sheet("22.下一步行动")
style_title(ws, "下一步行动清单（主题定稿 · 企业锁定 · 敦煌核查）", 4, TEAL)
hr = 4
header_row(ws, hr, ["序号", "行动项", "完成时限", "状态"], [8, 70, 18, 12])
limits = ["本周内", "对接会后48小时", "3日内", "1周内", "1周内", "10日内", "同步推进"]
rows = [[i + 1, t, limits[min(i, len(limits) - 1)], "待启动"] for i, t in enumerate(C.NEXT_STEPS)]
rows.insert(1, [0, "锁定优刻得、智谱、苏度等企业主论坛/分论坛/体验层互动名单", "本周内", "待启动"])
# renumber
for i, row in enumerate(rows):
    row[0] = i + 1
data_rows(ws, hr + 1, rows, aligns=["center", "left", "center", "center"], base_h=32)
freeze(ws)

# 23 主题板块
ws = wb.create_sheet("23.主题板块")
style_title(ws, "四大主题板块协同要点（含敦煌沉浸）", 3)
hr = 4
header_row(ws, hr, ["板块", "英文标签", "要点"], [22, 22, 70])
rows = [[p["name"], p["tag"], "；".join(p["points"])] for p in C.THEME_PILLARS]
data_rows(ws, hr + 1, rows, base_h=48)
freeze(ws)

wb.save(OUT_FILE)
print(f"已生成：{OUT_FILE}（{len(wb.sheetnames)} 张工作表）")
