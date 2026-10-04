# -*- coding: utf-8 -*-
"""上海新谷湾试单执行表。

外发：使用说明、空铺底稿、启动确认。
内部：筛选标准、企业筛选、九十天排期、责任与付费、新加坡对接、需求对照、内部口径。
"""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins
from pathlib import Path

PLUM = "2E1F47"
PURPLE = "5B3E8E"
GOLD = "C19A3A"
LAV = "ECE6F7"
WHITE = "FFFFFF"
INK = "2B2536"
GREY = "605A6B"
RED = "8C3A3A"
CREAM = "FBF7EF"
LINE = "DDD2EF"

thin = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
fill_plum = PatternFill("solid", fgColor=PLUM)
fill_purple = PatternFill("solid", fgColor=PURPLE)
fill_gold = PatternFill("solid", fgColor=GOLD)
fill_lav = PatternFill("solid", fgColor=LAV)
fill_white = PatternFill("solid", fgColor=WHITE)
fill_cream = PatternFill("solid", fgColor=CREAM)
fill_red = PatternFill("solid", fgColor="F8E8E8")
fill_input = PatternFill("solid", fgColor="FFF9EE")

font_white = Font(name="微软雅黑", size=11, color=WHITE, bold=True)
font_title = Font(name="微软雅黑", size=18, color=WHITE, bold=True)
font_head = Font(name="微软雅黑", size=12, color=PLUM, bold=True)
font_body = Font(name="微软雅黑", size=10, color=INK)
font_small = Font(name="微软雅黑", size=9, color=GREY)
font_bold = Font(name="微软雅黑", size=10, color=PLUM, bold=True)
font_gold = Font(name="微软雅黑", size=10, color=PLUM, bold=True)

wrap = Alignment(wrap_text=True, vertical="center", horizontal="left")
wrap_c = Alignment(wrap_text=True, vertical="center", horizontal="center")
left = Alignment(vertical="center", horizontal="left", wrap_text=True)


def widths(ws, pairs):
    for col, w in pairs.items():
        ws.column_dimensions[col].width = w


def paint(cell, value, font=font_body, fill=None, align=wrap, border=thin):
    cell.value = value
    cell.font = font
    cell.alignment = align
    if fill is not None:
        cell.fill = fill
    if border is not None:
        cell.border = border


def header_row(ws, row, labels, fill=fill_plum):
    for i, label in enumerate(labels, 1):
        paint(ws.cell(row, i), label, font=font_white, fill=fill, align=wrap_c)


def banner(ws, row, cols, text, fill=fill_plum, size=18):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    paint(
        ws.cell(row, 1),
        text,
        font=Font(name="微软雅黑", size=size, color=WHITE, bold=True),
        fill=fill,
        align=Alignment(vertical="center", horizontal="left", indent=1),
        border=None,
    )
    ws.row_dimensions[row].height = 32


def note(ws, row, cols, text, fill=fill_cream):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    paint(ws.cell(row, 1), text, font=font_small, fill=fill, align=wrap, border=None)
    ws.row_dimensions[row].height = 36


def page(ws, title):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.6, bottom=0.5, header=0.2, footer=0.2)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddHeader.left.text = title
    ws.oddFooter.right.text = "杨浦区科技企业联合会  ·  上海新谷湾试单  ·  第 &P 页"
    ws.oddFooter.left.text = "内部执行表。外发前只保留对方要填的表。"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.freeze_panes = "A4"
    ws.sheet_format.defaultRowHeight = 18


def dv(ws, formula, cells):
    rule = DataValidation(type="list", formula1=formula, allow_blank=True)
    rule.error = "请从列表里选"
    rule.errorTitle = "选项"
    rule.prompt = "点开选择"
    rule.promptTitle = "填写"
    ws.add_data_validation(rule)
    rule.add(cells)


wb = Workbook()

# ============================================================ 使用说明
ws = wb.active
ws.title = "使用说明"
page(ws, "使用说明")
widths(ws, {"A": 22, "B": 78, "C": 28})
banner(ws, 1, 3, "上海新谷湾试单执行表")
note(ws, 2, 3, "2026 年 10 月  ·  依据 9 月 24 日洽谈  ·  先入会，再试 90 天，不签招商对赌。面积和报价以空铺清单为准。")
header_row(ws, 3, ["表", "做什么", "能不能外发"])
rows = [
    ("空铺底稿", "对方填写。试单的启动门。小业主面积不要混进可推清单。", "可以发"),
    ("启动确认", "群里对齐对接人、清单日期、首场窗口。", "可以发"),
    ("九十天排期", "过程表。只记筛企业、沙龙、陪看、复盘。", "对过口径后可以发"),
    ("责任与付费", "和 PPT 第 10 页一致，用来避免会后各说各话。", "对过口径后可以发"),
    ("新加坡对接", "和招商分开。只记一次验证会。", "对口同事用"),
    ("筛选标准", "10–15 家怎么选。先有清单再筛人。", "内部"),
    ("企业筛选", "内部名单。整表不发到对方群。", "内部，勿外发"),
    ("需求对照", "录音原话收成需求和动作，便于内部核对。", "内部"),
    ("内部口径", "不对外说的边界：东升、宝龙 KPI、价格、创智汇。", "内部，勿外发"),
]
for i, row in enumerate(rows):
    fill = fill_lav if "内部" in row[2] else fill_white
    for c, val in enumerate(row, 1):
        paint(ws.cell(4 + i, c), val, font=font_bold if c == 1 else font_body, fill=fill)
    ws.row_dimensions[4 + i].height = 28
ws.row_dimensions[1].height = 36
note(ws, 14, 3, "发给对方时：另存，只留「空铺底稿」和「启动确认」。企业名称出现在本文件后，文件即视为内部资料。")
ws.auto_filter.ref = "A3:C12"
ws.auto_filter.ref = "A3:C12"
ws.freeze_panes = "A4"
ws.sheet_properties.tabColor = GOLD
ws.print_title_rows = "1:3"
ws.page_setup.fitToHeight = 1

# ============================================================ 空铺底稿
ws = wb.create_sheet("空铺底稿")
page(ws, "空铺底稿")
headers = [
    "序号", "楼层", "房号", "可租面积㎡", "权属", "状态", "到期日",
    "报价元/㎡/天", "是否含物业", "物业费元/㎡/月", "物业费能否谈",
    "交付标准", "能否注册", "最短租期月", "免租月", "空调", "朝向/层高",
    "谁能定租", "对接人", "电话", "备注",
]
widths(ws, {
    "A": 8, "B": 10, "C": 12, "D": 14, "E": 12, "F": 12, "G": 14,
    "H": 16, "I": 14, "J": 16, "K": 14, "L": 14, "M": 12, "N": 14,
    "O": 12, "P": 14, "Q": 16, "R": 16, "S": 12, "T": 16, "U": 28,
})
banner(ws, 1, len(headers), "自持空铺清单  ·  请新谷湾填写")
note(ws, 2, len(headers), "只填有权出租、有权定租的面积。小业主房源请另表，不要写在本表。报价不含物业。物业费洽谈口径约 20 元/㎡/月，请按实际填写，并标明能否谈。")
header_row(ws, 3, headers)
dv(ws, '"自持,小业主,待核实"', "E4:E23")
dv(ws, '"空置,在租即将到期,在谈,已租不推"', "F4:F23")
dv(ws, '"不含,含,待确认"', "I4:I23")
dv(ws, '"能谈,不能谈,未知"', "K4:K23")
dv(ws, '"毛坯,精装,现状交付,待确认"', "L4:L23")
dv(ws, '"能,不能,待确认"', "M4:M23")
dv(ws, '"独立空调,中央空调,需改造,未知"', "P4:P23")
for r in range(4, 24):
    for c in range(1, len(headers) + 1):
        paint(ws.cell(r, c), r - 3 if c == 1 else None, font=font_body, fill=fill_input, align=wrap_c if c < 18 else wrap)
    ws.row_dimensions[r].height = 22
ws.auto_filter.ref = "A3:U23"
ws.freeze_panes = "D4"
ws.auto_filter.ref = "A3:U23"
ws.sheet_properties.tabColor = "2E7D4F"
ws.oddFooter.left.text = "本表可外发。请填完发回联合会对接群。"
ws.print_title_rows = "1:3"

# ============================================================ 启动确认
ws = wb.create_sheet("启动确认")
page(ws, "启动确认")
widths(ws, {"A": 28, "B": 55, "C": 36})
banner(ws, 1, 3, "试单启动确认")
note(ws, 2, 3, "群里先填这张。清单和新加坡对接人到齐之前，不推企业，也不讨论框架协议。")
header_row(ws, 3, ["项目", "填写", "说明"])
items = [
    ("对方签约主体", "", "以营业执照全称为准，不用口头简称"),
    ("新谷湾对接人", "", "能约看盘、能把租金带回决策的人"),
    ("对接人电话", "", ""),
    ("新加坡对接人", "", "可以不是招商负责人"),
    ("新加坡对接人电话", "", "出海同事单独联系"),
    ("空铺清单提交日", "", "试单从这一天起算 90 日"),
    ("入会主体", "", "副会长，年费 1 万元"),
    ("首场沙龙窗口", "", "先量茶馆再定 30–80 人是否成立"),
    ("确认：不设税收平台", "", "填「确认」或「再议」。再议则本单暂停"),
    ("确认：不对赌面积", "", "填「确认」"),
    ("确认：只推自持清单", "", "填「确认」"),
    ("确认：创智汇另行沟通", "", "填「确认」"),
]
for i, (a, b, c) in enumerate(items):
    r = 4 + i
    paint(ws.cell(r, 1), a, font=font_bold, fill=fill_lav)
    paint(ws.cell(r, 2), b, font=font_body, fill=fill_input)
    paint(ws.cell(r, 3), c, font=font_small, fill=fill_white)
    ws.row_dimensions[r].height = 24
ws.sheet_properties.tabColor = "2E7D4F"
ws.freeze_panes = "A4"

# ============================================================ 九十天
ws = wb.create_sheet("九十天排期")
page(ws, "九十天排期")
widths(ws, {"A": 16, "B": 42, "C": 36, "D": 22, "E": 14, "F": 28})
banner(ws, 1, 6, "90 天过程排期  ·  不考核成交面积")
note(ws, 2, 6, "起算日 = 空铺清单书面提交日。电话咨询不记入带看。")
header_row(ws, 3, ["时间", "动作", "完成的样子", "负责", "状态", "记录"])
plan = [
    ("第 1 周", "收到空铺清单", "可租面积、报价、交付、能否注册填全", "新谷湾", "未开始", ""),
    ("第 1 周", "收到新加坡对接人", "姓名和电话进「新加坡对接」表", "新谷湾", "未开始", ""),
    ("第 1 周", "确认入会主体", "副会长文本写明场地和月度沙龙", "联合会", "未开始", ""),
    ("第 2–4 周", "内部筛 10–15 家", "按「筛选标准」填「企业筛选」，整表不外发", "联合会", "未开始", ""),
    ("第 1 个月", "第 1 场沙龙 + 陪看", "主题为人工智能落地或早期融资；至少 1 位早期投资人到场", "双方", "未开始", ""),
    ("第 2 个月", "第 2 场沙龙 + 陪看", "同上，另有现场记录", "双方", "未开始", ""),
    ("第 3 个月", "第 3 场沙龙 + 陪看", "同上", "双方", "未开始", ""),
    ("第 90 天", "书面复盘", "继续 / 停止 / 改成单场活动", "双方", "未开始", ""),
]
dv(ws, '"未开始,进行中,完成,暂停"', "E4:E11")
for i, row in enumerate(plan):
    r = 4 + i
    for c, val in enumerate(row, 1):
        paint(ws.cell(r, c), val, font=font_bold if c == 1 else font_body,
              fill=fill_white if i % 2 == 0 else fill_lav,
              align=wrap_c if c in (1, 5) else wrap)
    ws.row_dimensions[r].height = 32
ws.auto_filter.ref = "A3:F11"
ws.freeze_panes = "A4"
ws.sheet_properties.tabColor = PURPLE

# ============================================================ 责任与付费
ws = wb.create_sheet("责任与付费")
page(ws, "责任与付费")
widths(ws, {"A": 16, "B": 38, "C": 38, "D": 32, "E": 28})
banner(ws, 1, 5, "责任与付费  ·  与方案 PPT 一致")
note(ws, 2, 5, "试单不产生成交佣金、税收返还、传播费、新加坡落地费、外地产线费。这些要做，另写一页。")
header_row(ws, 3, ["事项", "新谷湾交付", "联合会交付", "费用", "不包含"])
fees = [
    ("入会", "以副会长单位加入", "接收会员，把场地纳入活动点", "1 万元/年，入会时支付", "不含招商佣金"),
    ("场地", "每月开放小型场地一场", "确定主题、流程和邀约", "场地不收费", "不含搭建和晚宴"),
    ("沙龙", "现场接待、开门", "试单期组织 3 场", "组织费试单期由联合会承担", "对外传播另计"),
    ("陪看", "按清单报价，授权人可定租", "筛选、预约、陪同到场", "不收费", "不写面积和税收"),
    ("挂牌", "提供墙面并配合发布", "协调复旦科技园或现有会客厅", "主体未确认则不发生", "不新设空壳"),
    ("新加坡", "指定对接人并说明选址进度", "出海同事参加一次会", "不进本方案", "不承诺代为租下站点"),
    ("产线外溢", "开放已有外地通道", "只转介自己提出需求的企业", "单案单议", "不承诺批量导入"),
]
for i, row in enumerate(fees):
    r = 4 + i
    for c, val in enumerate(row, 1):
        paint(ws.cell(r, c), val, font=font_bold if c == 1 else font_body,
              fill=fill_white if i % 2 == 0 else fill_lav)
    ws.row_dimensions[r].height = 36
ws.freeze_panes = "A4"
ws.sheet_properties.tabColor = PURPLE

# ============================================================ 新加坡
ws = wb.create_sheet("新加坡对接")
page(ws, "新加坡对接")
widths(ws, {"A": 24, "B": 50, "C": 36})
banner(ws, 1, 3, "新加坡选址  ·  与招商分开")
note(ws, 2, 3, "一次验证会。对方刚在选址；我方 TikTok 电商和小游戏已运营两年。澳大利亚今年才开，不放进这次会议的交付。")
header_row(ws, 3, ["项目", "填写", "备注"])
sg = [
    ("对方对接人", "", "会上没有留下姓名"),
    ("电话 / 微信", "", ""),
    ("选址阶段", "", "选楼 / 已看 / 已锁定 / 未知"),
    ("需要的面积", "", ""),
    ("用途", "", "办公 / 仓储 / 直播 / 其他"),
    ("希望我方谁参加", "", "出海同事，不默认招商同事"),
    ("会议日期", "", ""),
    ("会议结论", "", "能互相落地 / 再观察 / 停止"),
    ("下一步", "", "不写进招商确认函"),
]
dv(ws, '"选楼,已看,已锁定,未知"', "B6")
dv(ws, '"能互相落地,再观察,停止"', "B11")
for i, (a, b, c) in enumerate(sg):
    r = 4 + i
    paint(ws.cell(r, 1), a, font=font_bold, fill=fill_lav)
    paint(ws.cell(r, 2), b, fill=fill_input)
    paint(ws.cell(r, 3), c, font=font_small)
    ws.row_dimensions[r].height = 24
ws.freeze_panes = "A4"
ws.sheet_properties.tabColor = "3E6B8E"

# ============================================================ 筛选标准
ws = wb.create_sheet("筛选标准")
page(ws, "筛选标准")
widths(ws, {"A": 18, "B": 55, "C": 42})
banner(ws, 1, 3, "内部  ·  10–15 家怎么筛")
note(ws, 2, 3, "内部资料，不要发到对方群。顺序：先看见空铺的面积和交付，再从白皮书和人工智能库里找人。")
header_row(ws, 3, ["条件", "要满足", "不推"])
rules = [
    ("产业", "人工智能，或相邻硬科技", "传统零售；对方会用租金挡掉"),
    ("阶段", "早期，投资人还能影响选址", "中后期。国寿、兵器这类对方自己说帮助有限"),
    ("增长", "正增长，或融资正在谈", "付不起大约 1.5–2 元/㎡/天的团队"),
    ("面积", "对得上清单里的现有房源", "要求再切 1–4 人隔间的"),
    ("产品", "接受现在的大开间", "要求改公寓、改酒店的"),
    ("权属", "清单标明自持且能定租", "小业主价格战里的 1 元房源"),
    ("绑定", "尚未被免租加投资锁定", "已经确定要进免费场地加大额投资平台的"),
    ("陪看", "能约到一位早期投资人一起看", "只有朋友圈转发、没有到场的"),
    ("名录", "定向推荐，约见前口头介绍", "把企业筛选整表交给对方"),
    ("价格", "守住会员企业自己的谈判", "为了填空置，把别的园区在谈的价格打到 1 元"),
]
for i, row in enumerate(rules):
    r = 4 + i
    for c, val in enumerate(row, 1):
        paint(ws.cell(r, c), val, font=font_bold if c == 1 else font_body,
              fill=fill_red if c == 3 else (fill_white if i % 2 == 0 else fill_lav))
    ws.row_dimensions[r].height = 30
ws.freeze_panes = "A4"
ws.sheet_properties.tabColor = RED

# ============================================================ 企业筛选
ws = wb.create_sheet("企业筛选")
page(ws, "企业筛选")
headers = [
    "序号", "企业简称", "来源", "产业", "阶段", "增长或融资", "需求面积㎡",
    "承受租金元/㎡/天", "接受大开间", "要注册", "投资人", "投资人影响选址",
    "对应房号", "状态", "备注",
]
widths(ws, {
    "A": 8, "B": 16, "C": 16, "D": 14, "E": 12, "F": 16, "G": 14,
    "H": 18, "I": 14, "J": 12, "K": 16, "L": 16, "M": 14, "N": 12, "O": 28,
})
banner(ws, 1, len(headers), "内部  ·  企业筛选（整表不外发）")
note(ws, 2, len(headers), "先填完空铺底稿，再在这里写企业。没有对应房号的，状态保持「待筛」。不要把本表发进对方微信群。")
header_row(ws, 3, headers, fill=PatternFill("solid", fgColor=RED))
dv(ws, '"白皮书,人工智能库,WAIC,投行推荐,走访,其他"', "C4:C18")
dv(ws, '"人工智能,硬科技,出海,其他科技"', "D4:D18")
dv(ws, '"种子,天使,Pre-A,A轮,其他早期"', "E4:E18")
dv(ws, '"正增长,融资在谈,待核"', "F4:F18")
dv(ws, '"是,否,待核"', "I4:I18")
dv(ws, '"要,不要,待核"', "J4:J18")
dv(ws, '"是,否,未知"', "L4:L18")
dv(ws, '"待筛,可约,已陪看,暂缓,不匹配"', "N4:N18")
for r in range(4, 19):
    for c in range(1, len(headers) + 1):
        paint(ws.cell(r, c), r - 3 if c == 1 else None, fill=fill_input, align=wrap_c if c != 15 else wrap)
    ws.row_dimensions[r].height = 22
ws.auto_filter.ref = "A3:O18"
ws.freeze_panes = "C4"
ws.sheet_properties.tabColor = RED
ws.oddFooter.left.text = "内部资料。勿外发。"

# ============================================================ 需求对照
ws = wb.create_sheet("需求对照")
page(ws, "需求对照")
widths(ws, {"A": 10, "B": 22, "C": 42, "D": 36, "E": 32, "F": 24})
banner(ws, 1, 6, "内部  ·  原话、需求、动作")
note(ws, 2, 6, "用来核对方案没有把没说过的需求写进去。原话按 9 月 24 日录音整理，不是逐字稿。")
header_row(ws, 3, ["顺序", "他们要的", "录音依据", "收成的需求", "我方动作", "付费"])
mapping = [
    ("1", "把空铺填上", "出租率七十多，今年退租多；自持约 2 元以上，旁侧抛到约 1 元；收尾说到实地帮着招商", "自持部分要有能付租的科技租户。物业费空着也要交", "清单到后筛 10–15 家并陪看", "不写佣金和税收"),
    ("2", "早期资本", "收尾承认缺资本；国寿、兵器偏中后期；晶丰被免租加投资挖走", "要还能影响选址的早期投资人", "每场至少一位早期机构看盘", "到场不另收费"),
    ("3", "挂牌", "主动问能不能在这里挂牌；不少企业不知道这栋楼", "用真实机构把「靠复旦」说清楚", "优先复旦科技园或现有会客厅。主体未定不挂", "不新设机构"),
    ("4", "小活动客流", "小场地免费；活动可以谈费用。现场带看约 110 比 1", "用到场的人换带看，不换成交承诺", "每月一场，30–80 人待实测", "场地免费；传播另计"),
    ("5", "新加坡", "约下次和出海同事聊；刚在新加坡选址", "验证互相落地", "出海同事单独开会", "不进招商安排"),
    ("6", "入会", "先问入会方式和年费", "买名分和后续导流", "副会长", "1 万元/年"),
]
for i, row in enumerate(mapping):
    r = 4 + i
    for c, val in enumerate(row, 1):
        paint(ws.cell(r, c), val, font=font_bold if c in (1, 2) else font_body,
              fill=fill_white if i % 2 == 0 else fill_lav, align=wrap)
    ws.row_dimensions[r].height = 58
ws.freeze_panes = "A4"
ws.sheet_properties.tabColor = RED
ws.page_setup.fitToHeight = 1

# ============================================================ 内部口径
ws = wb.create_sheet("内部口径")
page(ws, "内部口径")
widths(ws, {"A": 24, "B": 70})
banner(ws, 1, 2, "内部口径  ·  勿外发", fill=PatternFill("solid", fgColor=RED))
note(ws, 2, 2, "这些话用来约束自己，不上 PPT，不进对方群。")
header_row(ws, 3, ["口径", "怎么把握"], fill=PatternFill("solid", fgColor=RED))
internal = [
    ("东升巨变", "不在这栋楼。在楼里的是东升另一块业务，录音里的华数科技。对外按实际入驻介绍。"),
    ("自持面积", "洽谈口径约 6,000 多㎡。整栋约 1.67 万㎡已卖给小业主。推企业只推清单里的自持部分。"),
    ("租金带", "1.5–2 元/㎡/天是我们筛企业的承受带，不是替对方出的报价。对方说自持守在 2 元以上，好项目可以降。"),
    ("宝龙强度", "宝龙是四个月、约 1,000–1,600㎡、一个月约 120 个现场客。不要把这组数字写成新谷湾的考核。"),
    ("企业库", "可以约见，不可以整表交付。同时在做宝龙、创智汇、森马，东方枢纽也在排期。"),
    ("价格纪律", "不为填这栋楼的空置，把会员企业在其他园区的价格谈到 1 元。"),
    ("改公寓", "住房政策研究中心最多口头回答一次口径。不勘察、不写进合作范围、不陪跑批文。"),
    ("机器人", "巨神智能飞地对方自己说推得慢。对方再提，回答本单主题是人工智能早期项目。"),
    ("税收", "对方老板明确不做平台。谁带来的企业，税管留在谁那里。文本里不出现分成。"),
    ("创智汇", "主谈人在创智汇待过、从富湘出来。见面认交情。合作文本不绑定中建四局和创智汇。"),
    ("澳大利亚", "2026 年才开，不写进这 90 天的交付。新加坡才是这次验证的市场。"),
    ("未读录音的设想", "园区诊断、技术需求征集、为成都批量招商、长期对赌渠道，录音里没有这些委托，不放进试单。"),
]
for i, (a, b) in enumerate(internal):
    r = 4 + i
    paint(ws.cell(r, 1), a, font=font_bold, fill=fill_red)
    paint(ws.cell(r, 2), b, font=font_body, fill=fill_white)
    ws.row_dimensions[r].height = 36
ws.freeze_panes = "A4"
ws.sheet_properties.tabColor = RED

# 打印与视图
for ws in wb.worksheets:
    ws.page_setup.horizontalCentered = True
    ws.sheet_view.view = "pageBreakPreview"
    ws.sheet_view.view = "normal"
    ws.oddHeader.center.text = ""

out = Path("/workspace/deliverables")
out.mkdir(parents=True, exist_ok=True)
target = out / "上海新谷湾试单执行表.xlsx"
wb.save(target)
print(target)
