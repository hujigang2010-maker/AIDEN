"""生成 Excel：临港 x WeTest 中日闭门会 · 转化跟踪与运营管理表（侧重会后转化）

Sheets：
  00 总览 Dashboard
  01 项目时间倒排
  02 邀约名单与跟踪
  03 嘉宾与议程
  04 现场物资与分工
  05 预算明细
  06 会前转化动作
  07 会中转化抓手
  08 会后跟进 SOP（7-30-60-90）
  09 销售/转化漏斗
  10 KPI 与考核
  11 行业方案矩阵
  12 风险预案
  13 复盘记录模板
"""
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                             NamedStyle)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import CellIsRule, FormulaRule, ColorScaleRule
from openpyxl.worksheet.datavalidation import DataValidation

# 配色
PRIMARY = "0B3D91"
ACCENT = "C0392B"
GOLD = "D4A04C"
LIGHT = "F4F6FB"
GREY = "777777"
WHITE = "FFFFFF"
DARK = "222222"
GREEN = "1E8449"
ORANGE = "E67E22"
YELLOW = "F1C40F"


def thin():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)


def medium():
    s = Side(style="medium", color="0B3D91")
    return Border(left=s, right=s, top=s, bottom=s)


def title_cell(ws, cell, text, fill=PRIMARY, color=WHITE, size=14, bold=True):
    ws[cell] = text
    ws[cell].font = Font(name="微软雅黑", size=size, bold=bold, color=color)
    ws[cell].fill = PatternFill("solid", fgColor=fill)
    ws[cell].alignment = Alignment(horizontal="left", vertical="center")


def header_row(ws, row, headers, fill=PRIMARY, color=WHITE, height=26):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(name="微软雅黑", size=11, bold=True, color=color)
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        c.border = thin()
    ws.row_dimensions[row].height = height


def write_row(ws, row, data, fill=None, bold=False, size=10,
              wrap=True, align="left"):
    for i, v in enumerate(data, 1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name="微软雅黑", size=size, bold=bold,
                      color="222222")
        c.alignment = Alignment(horizontal=align, vertical="center",
                                wrap_text=wrap)
        c.border = thin()
        if fill:
            c.fill = PatternFill("solid", fgColor=fill)


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def section_banner(ws, row, text, span=10, fill=GOLD, color=WHITE):
    ws.merge_cells(start_row=row, start_column=1, end_row=row,
                   end_column=span)
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name="微软雅黑", size=12, bold=True, color=color)
    c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 24


def freeze(ws, cell):
    ws.freeze_panes = cell


def big_title(ws, span_to_col="J", text="", sub=""):
    ws.merge_cells(f"A1:{span_to_col}1")
    c = ws["A1"]
    c.value = text
    c.font = Font(name="微软雅黑", size=18, bold=True, color=WHITE)
    c.fill = PatternFill("solid", fgColor=PRIMARY)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 36
    ws.merge_cells(f"A2:{span_to_col}2")
    c2 = ws["A2"]
    c2.value = sub
    c2.font = Font(name="微软雅黑", size=10, italic=True, color=GREY)
    c2.fill = PatternFill("solid", fgColor=LIGHT)
    c2.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 20


def build_dashboard(wb):
    ws = wb.active
    ws.title = "00 总览Dashboard"
    big_title(ws, "L",
              "临港 × WeTest 中日数字经济跨国发展闭门会 · 总览 Dashboard",
              "政策护航 × 质量守护 ｜ 会后转化是本表的灵魂")

    # 项目基本信息
    section_banner(ws, 4, "一、项目基本信息", span=12, fill=PRIMARY)
    info = [
        ["主题", "中日数字经济跨国发展与高质量交付闭门会"],
        ["主办", "上海科技企业协会"],
        ["联合主办", "临港科技城 × WeTest"],
        ["活动日期", "（待定，建议 4-6 周筹备期）"],
        ["活动地点", "临港科技城"],
        ["活动规模", "20-30 人闭门"],
        ["目标客户", "金融 / 游戏 / 电商零售 / 泛互联网 — C 级/业务 VP"],
        ["核心目标", "T+90 内首单 ≥ 3 单 · GMV ≥ 300 万元 · ROI ≥ 5×"],
    ]
    for i, (k, v) in enumerate(info, 5):
        ws.cell(row=i, column=1, value=k).font = Font(
            name="微软雅黑", bold=True, color=DARK)
        ws.cell(row=i, column=1).fill = PatternFill("solid", fgColor=LIGHT)
        ws.cell(row=i, column=1).alignment = Alignment(
            horizontal="center", vertical="center")
        ws.cell(row=i, column=1).border = thin()
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=12)
        c = ws.cell(row=i, column=2, value=v)
        c.font = Font(name="微软雅黑", color=DARK)
        c.alignment = Alignment(horizontal="left", vertical="center",
                                wrap_text=True)
        c.border = thin()

    # 核心 KPI
    section_banner(ws, 14, "二、核心 KPI（自动联动 · 仅填实际值）",
                   span=12, fill=ACCENT)
    header_row(ws, 15,
               ["维度", "指标", "目标值", "实际值", "完成率", "状态"],
               fill=PRIMARY)
    kpi = [
        ("邀约", "邀请池规模", 75),
        ("邀约", "C 级/VP 占比 (%)", 70),
        ("到场", "实际到场人数", 25),
        ("到场", "到场率 (%)", 33),
        ("现场", "圆桌发言客户数", 8),
        ("现场", "1V1 深聊预约数", 12),
        ("转化", "T+7 需求澄清会", 9),
        ("转化", "T+30 POC 立项", 5),
        ("转化", "T+90 首单数", 3),
        ("收入", "首单 GMV (元)", 3000000),
        ("ROI", "ROI 倍数", 5),
        ("品牌", "媒体露出次数", 10),
    ]
    start_row = 16
    for i, (dim, kp, target) in enumerate(kpi):
        r = start_row + i
        ws.cell(row=r, column=1, value=dim)
        ws.cell(row=r, column=2, value=kp)
        ws.cell(row=r, column=3, value=target)
        ws.cell(row=r, column=4, value=0)  # 实际
        ws.cell(row=r, column=5,
                value=f"=IFERROR(D{r}/C{r},0)")
        ws.cell(row=r, column=5).number_format = "0.0%"
        ws.cell(row=r, column=6,
                value=f'=IF(D{r}=0,"未开始",IF(E{r}>=1,"达成",IF(E{r}>=0.7,"进行中","风险")))')
        for col in range(1, 7):
            c = ws.cell(row=r, column=col)
            c.font = Font(name="微软雅黑", size=10, color=DARK)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin()
            if col == 1:
                c.fill = PatternFill("solid", fgColor=LIGHT)

    # 状态条件格式
    last = start_row + len(kpi) - 1
    rng = f"F{start_row}:F{last}"
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="equal", formula=['"达成"'],
        fill=PatternFill("solid", fgColor="D4EDDA"),
        font=Font(color="155724", bold=True)))
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="equal", formula=['"进行中"'],
        fill=PatternFill("solid", fgColor="FFF3CD"),
        font=Font(color="856404", bold=True)))
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="equal", formula=['"风险"'],
        fill=PatternFill("solid", fgColor="F8D7DA"),
        font=Font(color="721C24", bold=True)))

    # 转化漏斗
    section_banner(ws, 30, "三、转化漏斗（公式联动 · 调整目标自动更新）",
                   span=12, fill=GOLD)
    header_row(ws, 31,
               ["阶段", "动作", "目标人数", "实际人数", "目标转化率",
                "实际转化率", "状态"],
               fill=PRIMARY)
    funnel = [
        ("L0 邀请池", "定向邀约", 75, 0, ""),
        ("L1 到场（认知）", "签到 + 名片", 25, 0, "33%"),
        ("L2 强意向", "圆桌发言 / 1V1 预约", 15, 0, "60%"),
        ("L3 商机 (T+7)", "需求澄清会", 9, 0, "60%"),
        ("L4 POC (T+30)", "POC 立项", 5, 0, "55%"),
        ("L5 签单 (T+90)", "首单签订", 3, 0, "60%"),
    ]
    for i, (stage, act, t_n, a_n, t_rate) in enumerate(funnel):
        r = 32 + i
        ws.cell(row=r, column=1, value=stage)
        ws.cell(row=r, column=2, value=act)
        ws.cell(row=r, column=3, value=t_n)
        ws.cell(row=r, column=4, value=a_n)
        ws.cell(row=r, column=5, value=t_rate)
        if i == 0:
            ws.cell(row=r, column=6, value="—")
        else:
            ws.cell(row=r, column=6,
                    value=f"=IFERROR(D{r}/D{r-1},0)")
            ws.cell(row=r, column=6).number_format = "0.0%"
        ws.cell(row=r, column=7,
                value=f'=IF(D{r}>=C{r},"达成",IF(D{r}>=C{r}*0.7,"进行中","风险"))')
        for col in range(1, 8):
            c = ws.cell(row=r, column=col)
            c.font = Font(name="微软雅黑", size=10, color=DARK)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin()
            if col == 1:
                c.fill = PatternFill("solid", fgColor=LIGHT)

    rng2 = f"G32:G{32+len(funnel)-1}"
    ws.conditional_formatting.add(rng2, CellIsRule(
        operator="equal", formula=['"达成"'],
        fill=PatternFill("solid", fgColor="D4EDDA"),
        font=Font(color="155724", bold=True)))
    ws.conditional_formatting.add(rng2, CellIsRule(
        operator="equal", formula=['"进行中"'],
        fill=PatternFill("solid", fgColor="FFF3CD")))
    ws.conditional_formatting.add(rng2, CellIsRule(
        operator="equal", formula=['"风险"'],
        fill=PatternFill("solid", fgColor="F8D7DA")))

    # 使用说明
    section_banner(ws, 40, "四、使用说明", span=12, fill=PRIMARY)
    notes = [
        "1. 本工作簿为「闭门沙龙 → 转化签单」的运营底盘，每张表都对应"
        "1 个关键运营动作。",
        "2. 「02 邀约名单与跟踪」是会前的主战场，确保 75 人邀请池 / 25 人到场。",
        "3. 「08 会后跟进 SOP」是 ★ 核心 —— 活动结束 48h 启动，"
        "每条 A/B 线索必须有 Owner、动作、截止日。",
        "4. 「09 销售/转化漏斗」每日更新，A 级线索每周三同步进度。",
        "5. 「10 KPI 与考核」与销售 OKR 挂钩；本表数据每周汇总至本 Dashboard。",
        "6. 任何字段以蓝色填充列表头的均为「填写区」，其他为「自动计算区」。",
    ]
    for i, t in enumerate(notes, 41):
        ws.merge_cells(start_row=i, start_column=1, end_row=i,
                       end_column=12)
        c = ws.cell(row=i, column=1, value=t)
        c.font = Font(name="微软雅黑", size=10, color=DARK)
        c.alignment = Alignment(horizontal="left", vertical="center",
                                wrap_text=True)
        ws.row_dimensions[i].height = 22

    set_widths(ws, [16, 22, 14, 14, 14, 12, 12, 12, 12, 12, 12, 12])
    freeze(ws, "A3")


def build_timeline(wb):
    ws = wb.create_sheet("01 项目时间倒排")
    big_title(ws, "I",
              "项目时间倒排（T-30 → T+90）",
              "Owner 必须明确到人，截止日不可空")
    headers = ["阶段", "T-日", "建议日期", "关键动作", "可交付物",
               "Owner", "协作方", "状态", "备注"]
    header_row(ws, 4, headers, fill=PRIMARY)
    data = [
        ["T-30", "T-30", "", "三方 Kick-off：主题/嘉宾/预算",
         "Kick-off 会议纪要", "协会项目总", "临港 + WeTest", "未开始", ""],
        ["T-25", "T-25", "", "三方提交邀请池", "75 人初稿名单",
         "协会会员部 + 临港招商 + WeTest 销售", "—", "未开始", ""],
        ["T-21", "T-21", "", "锁定邀请名单 + NDA", "终版名单 + NDA",
         "协会项目总", "三方", "未开始", ""],
        ["T-15", "T-15", "", "首轮邀约（电话/微信）+ 邀请函",
         "首轮触达记录 / 邀请函", "三方销售", "—", "未开始",
         "嵌入「圆桌议题征集表」"],
        ["T-10", "T-10", "", "实体邀请函 + 二轮确认",
         "确认到场名单 V1", "协会会员部", "—", "未开始",
         "C 级寄送伴手礼"],
        ["T-7", "T-7", "", "嘉宾 PPT 定稿 + 圆桌议题确认",
         "议程终版 / PPT", "WeTest 技术总监", "临港政策代表",
         "未开始", ""],
        ["T-5", "T-5", "", "推送中日专项方案白皮书",
         "白皮书 PDF + 推送记录", "WeTest 市场",
         "—", "未开始", "建立技术认知"],
        ["T-3", "T-3", "", "TOP10 高优 1V1 预热电话",
         "预热脚本 + 接通记录", "WeTest 大客户经理",
         "—", "未开始", "预约会后 30 分钟深聊"],
        ["T-2", "T-2", "", "交通接待 / 嘉宾介绍卡 / 动线图推送",
         "推送记录", "协会会务", "—", "未开始", ""],
        ["T-1", "T-1", "", "现场彩排（设备、动线、主持稿）",
         "彩排清单", "三方", "—", "未开始", ""],
        ["T0", "T0", "", "活动执行 + 现场转化抓手全部触发",
         "签到表 / 圆桌纪要 / 现场视频", "三方",
         "—", "未开始", "★ 关键日"],
        ["T+1", "T+1", "", "感谢信群发 + 1V1 销售触达 + 资料包",
         "触达记录 / 资料包", "WeTest 销售", "市场",
         "未开始", "★ 黄金 48h"],
        ["T+3", "T+3", "", "CRM 录入完成 + ABC 分级",
         "CRM 截图 / 分级表", "WeTest 销售运营",
         "—", "未开始", ""],
        ["T+3", "T+3", "", "三方内部快速复盘",
         "复盘纪要", "协会项目总", "三方", "未开始", ""],
        ["T+7", "T+7", "", "A 级客户完成需求澄清会",
         "需求澄清纪要", "WeTest 大客户经理",
         "—", "未开始", "每位 A 级独立 Owner"],
        ["T+15", "T+15", "", "B 级方案 PPT 定向推送 + 二次深聊",
         "方案 PPT / 二深聊纪要", "WeTest 大客户经理",
         "—", "未开始", ""],
        ["T+15", "T+15", "", "客户视角复盘（抽样回访）",
         "回访录音 / 改进点", "协会项目总", "—", "未开始", ""],
        ["T+30", "T+30", "", "A 级 POC 立项 + B 级二次深聊",
         "POC SOW / 二深聊纪要", "WeTest 行业总监",
         "—", "未开始", ""],
        ["T+30", "T+30", "", "转化中期复盘",
         "中期复盘报告", "协会 + WeTest", "—", "未开始", ""],
        ["T+60", "T+60", "", "POC 中间汇报 + 商务洽谈",
         "POC 报告 / 商务草案", "WeTest 大客户经理",
         "—", "未开始", ""],
        ["T+90", "T+90", "", "首单签订 + 复盘 + 下一场迭代",
         "合同 / 收单复盘报告", "WeTest 行业总监",
         "三方", "未开始", "★ 收单截止"],
    ]
    for i, row in enumerate(data):
        write_row(ws, 5 + i, row, fill=LIGHT if i % 2 else None,
                  align="left")

    dv = DataValidation(type="list",
                        formula1='"未开始,进行中,已完成,延期,风险"',
                        allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"H5:H{4+len(data)}")

    set_widths(ws, [8, 6, 14, 38, 26, 22, 22, 12, 24])
    freeze(ws, "A5")


def build_invitee(wb):
    ws = wb.create_sheet("02 邀约名单与跟踪")
    big_title(ws, "S",
              "邀约名单与到场跟踪（75 人邀请池 → 25 人到场）",
              "每人均需有 Owner、首轮触达日、确认状态、行业标签")
    headers = ["#", "公司名称", "行业", "公司类型", "客户来源",
               "客户姓名", "职位/层级", "决策权",
               "联系方式", "诉求标签（出海/入华）", "诉求摘要",
               "Owner（邀请人）", "首轮触达日", "二轮确认日",
               "确认状态", "实际到场", "现场行为标签",
               "ABC 分级（会后）", "备注"]
    header_row(ws, 4, headers, fill=PRIMARY)
    sample = [
        [1, "示例 · A 出海游戏公司", "游戏", "中资出海",
         "WeTest 销售", "张总", "CTO", "拍板",
         "13800000000 / zhang@a.com", "出海日本",
         "日本市场首发 / 弱网 / 真金支付",
         "WeTest-王经理", "", "", "未确认", "未确认",
         "", "", "提前 1 周锁定"],
        [2, "示例 · B 跨境支付", "金融", "中资出海",
         "协会会员", "李总", "CIO", "强影响",
         "13900000000", "出海日本",
         "数据跨境合规 + 风控",
         "协会-赵秘书", "", "", "未确认", "未确认",
         "", "", ""],
        [3, "示例 · 日本 C 零售", "电商/零售", "日资入华",
         "临港招商", "Tanaka", "中国区总经理", "拍板",
         "tanaka@c.jp", "入华",
         "小程序生态 / 微信支付合规",
         "临港-陈顾问", "", "", "未确认", "未确认",
         "", "", "日语对接"],
    ]
    for i, row in enumerate(sample):
        write_row(ws, 5 + i, row, fill=LIGHT if i % 2 else None)
    # 留 80 行空位
    for i in range(len(sample), 80):
        for col in range(1, len(headers) + 1):
            c = ws.cell(row=5 + i, column=col,
                        value=(i + 1) if col == 1 else None)
            c.font = Font(name="微软雅黑", size=10)
            c.alignment = Alignment(horizontal="center" if col == 1
                                    else "left",
                                    vertical="center", wrap_text=True)
            c.border = thin()
            if i % 2:
                c.fill = PatternFill("solid", fgColor=LIGHT)

    # 数据验证
    dv_ind = DataValidation(type="list",
                            formula1='"金融,游戏,电商/零售,泛互联网,其他"',
                            allow_blank=True)
    ws.add_data_validation(dv_ind)
    dv_ind.add("C5:C84")

    dv_type = DataValidation(type="list",
                             formula1='"中资出海,日资入华,在华日资,其他"',
                             allow_blank=True)
    ws.add_data_validation(dv_type)
    dv_type.add("D5:D84")

    dv_src = DataValidation(type="list",
                            formula1='"协会会员,临港招商,WeTest 销售,合作伙伴,媒体推荐"',
                            allow_blank=True)
    ws.add_data_validation(dv_src)
    dv_src.add("E5:E84")

    dv_dec = DataValidation(type="list",
                            formula1='"拍板,强影响,弱影响,信息收集"',
                            allow_blank=True)
    ws.add_data_validation(dv_dec)
    dv_dec.add("H5:H84")

    dv_st = DataValidation(type="list",
                           formula1='"未确认,已确认,婉拒,改派,缺席"',
                           allow_blank=True)
    ws.add_data_validation(dv_st)
    dv_st.add("O5:O84")

    dv_at = DataValidation(type="list",
                           formula1='"未到,到场,迟到,提前离场"',
                           allow_blank=True)
    ws.add_data_validation(dv_at)
    dv_at.add("P5:P84")

    dv_abc = DataValidation(type="list",
                            formula1='"A · 高优,B · 中优,C · 培育,未分级"',
                            allow_blank=True)
    ws.add_data_validation(dv_abc)
    dv_abc.add("R5:R84")

    # 条件格式
    ws.conditional_formatting.add("O5:O84",
        CellIsRule(operator="equal", formula=['"已确认"'],
                   fill=PatternFill("solid", fgColor="D4EDDA")))
    ws.conditional_formatting.add("O5:O84",
        CellIsRule(operator="equal", formula=['"婉拒"'],
                   fill=PatternFill("solid", fgColor="F8D7DA")))
    ws.conditional_formatting.add("P5:P84",
        CellIsRule(operator="equal", formula=['"到场"'],
                   fill=PatternFill("solid", fgColor="D4EDDA")))
    ws.conditional_formatting.add("R5:R84",
        CellIsRule(operator="equal", formula=['"A · 高优"'],
                   fill=PatternFill("solid", fgColor="FAD7A0"),
                   font=Font(bold=True, color="7E5109")))

    set_widths(ws, [5, 26, 12, 12, 14, 12, 14, 12, 22, 14, 28,
                    14, 12, 12, 12, 12, 16, 14, 22])
    freeze(ws, "C5")

    # 顶部汇总
    ws.merge_cells("A3:S3")
    c = ws["A3"]
    c.value = ('实时统计： 邀请数 =COUNTA(B5:B84)   |   '
               '已确认 =COUNTIF(O5:O84,"已确认")   |   '
               '到场 =COUNTIF(P5:P84,"到场")   |   '
               'A 级 =COUNTIF(R5:R84,"A · 高优")')
    c.font = Font(name="微软雅黑", size=9, italic=True, color=GREY)
    c.fill = PatternFill("solid", fgColor=LIGHT)
    c.alignment = Alignment(horizontal="left", vertical="center")


def build_agenda(wb):
    ws = wb.create_sheet("03 嘉宾与议程")
    big_title(ws, "H", "嘉宾阵容 & 半日议程", "对应 Word/PPT，进一步细化执行")
    section_banner(ws, 4, "A. 半日议程", span=8, fill=PRIMARY)
    header_row(ws, 5, ["时间", "环节", "内容", "主讲/主持",
                       "时长(分钟)", "物料", "转化触点", "Owner"])
    agenda = [
        ["13:00-13:30", "签到 & 茶歇",
         "扫码进群 + 行业标签 + 1V1 破冰",
         "礼仪 + 协会会务", 30,
         "二维码 / 签到表 / 名片夹",
         "采集联系方式 + 行业标签", "协会"],
        ["13:30-14:30", "产业探访",
         "临港集团 + 科技城展厅参观 + 政策初探",
         "临港招商", 60,
         "讲解器 / 政策手册",
         "建立政策势能", "临港"],
        ["14:30-15:10", "宏观指引",
         "中日数字贸易趋势 + 临港自贸区跨国壁垒",
         "姚志勇教授（拟邀）", 40,
         "PPT / 笔记本",
         "权威认知背书", "协会"],
        ["15:10-15:50", "技术落地",
         "WeTest 中日跨国高质量交付与安全合规",
         "WeTest 业务线代表", 40,
         "PPT / Demo 设备 / 案例 / 限时权益券",
         "★ Demo + 案例 + 转化", "WeTest"],
        ["15:50-16:30", "闭门圆桌",
         "议题：双向拓展最大的本地化阻碍是什么？",
         "WeTest 行业总监 主持", 40,
         "议题卡 / 议程纪要本",
         "★ 1V1 资源对接 + 合作意向", "WeTest"],
        ["16:30-18:00", "晚间深聊",
         "Networking + 高优客户 1V1 30 分钟",
         "WeTest 销售 + 临港顾问", 90,
         "伴手礼 / 咨询券",
         "★ 线索→机会 升级", "WeTest"],
    ]
    for i, r in enumerate(agenda):
        write_row(ws, 6 + i, r, fill=LIGHT if i % 2 else None)

    section_banner(ws, 14, "B. 嘉宾阵容", span=8, fill=ACCENT)
    header_row(ws, 15, ["环节", "拟邀嘉宾", "身份", "联系方式",
                        "确认状态", "PPT 截止日", "Backup",
                        "Owner"])
    guests = [
        ["开场致辞", "临港科技城高管", "联合主办", "", "未确认",
         "T-3", "协会秘书长", "协会"],
        ["开场致辞", "上海科技企业协会秘书长", "主办", "", "未确认",
         "T-3", "—", "协会"],
        ["宏观指引", "姚志勇教授", "学术专家", "", "未确认",
         "T-7", "同领域权威 1 位", "协会"],
        ["技术落地", "WeTest 业务线代表", "技术专家", "", "未确认",
         "T-7", "WeTest 技术总监", "WeTest"],
        ["圆桌主持", "WeTest 行业总监", "圆桌主持", "", "未确认",
         "T-3", "WeTest CSO", "WeTest"],
        ["圆桌嘉宾", "金融行业 C 级 1 位", "客户代表", "", "未确认",
         "T-5", "—", "WeTest"],
        ["圆桌嘉宾", "游戏行业 C 级 1 位", "客户代表", "", "未确认",
         "T-5", "—", "WeTest"],
        ["圆桌嘉宾", "电商行业 C 级 1 位", "客户代表", "", "未确认",
         "T-5", "—", "WeTest"],
        ["圆桌嘉宾", "泛互联网行业 C 级 1 位", "客户代表", "", "未确认",
         "T-5", "—", "WeTest"],
    ]
    for i, r in enumerate(guests):
        write_row(ws, 16 + i, r, fill=LIGHT if i % 2 else None)

    set_widths(ws, [14, 18, 30, 22, 14, 14, 22, 12])
    freeze(ws, "A5")


def build_logistics(wb):
    ws = wb.create_sheet("04 现场物资与分工")
    big_title(ws, "H", "现场物资清单与分工", "T-3 彩排前完成全部 ✔")
    header_row(ws, 4,
               ["分类", "物资/工作项", "数量/规格", "供应方/Owner",
                "采购截止", "到场时间", "状态", "备注"])
    items = [
        ["签到", "签到台 + 桌牌", "1 套", "协会会务", "T-5",
         "13:00", "未开始", ""],
        ["签到", "签到二维码 / 行业标签自动打标", "1 套", "WeTest IT",
         "T-7", "13:00", "未开始", "企微对接"],
        ["签到", "胸卡 + 名片夹", "40 套", "协会会务", "T-5",
         "13:00", "未开始", ""],
        ["物料", "实体邀请函", "60 份", "协会市场", "T-12",
         "T-10 寄送", "未开始", ""],
        ["物料", "资料袋（政策手册+方案白皮书+权益券）", "40 套",
         "三方联制", "T-5", "13:00", "未开始", ""],
        ["物料", "桌签/议程印刷", "30 份", "协会会务", "T-3",
         "13:00", "未开始", ""],
        ["物料", "伴手礼（C 级专属）", "30 份", "协会市场", "T-7",
         "13:00", "未开始", "含权益券"],
        ["Demo", "WeTest 体验角真机（含日本节点）", "3 台 + 1 大屏",
         "WeTest 技术", "T-3", "13:00", "未开始", "★ 关键"],
        ["Demo", "海外众测看板 / CrashSight 面板",
         "演示账号", "WeTest 技术", "T-3", "13:00",
         "未开始", ""],
        ["主持", "外部主持人 / 礼仪", "1 主持 + 2 礼仪",
         "协会市场", "T-7", "12:30", "未开始", ""],
        ["影像", "摄影 + 摄像 + 精剪", "1 组", "协会市场", "T-7",
         "12:30", "未开始", "闭门禁直播"],
        ["茶歇", "茶歇 + 晚间轻餐", "30 人份", "协会会务", "T-3",
         "13:00", "未开始", ""],
        ["接待", "C 级 1V1 接送车", "10 车次", "协会 + 临港", "T-5",
         "—", "未开始", ""],
        ["交付", "圆桌议题卡 / 纪要本", "30 份", "协会会务", "T-3",
         "13:00", "未开始", ""],
        ["IT", "现场 Wi-Fi + 备用 4G/5G", "1 套", "临港 + WeTest",
         "T-1", "12:30", "未开始", "双链路冗余"],
        ["合规", "NDA 协议", "30 份", "协会法务", "T-3", "13:00",
         "未开始", "签到时签署"],
    ]
    for i, r in enumerate(items):
        write_row(ws, 5 + i, r, fill=LIGHT if i % 2 else None)

    dv = DataValidation(type="list",
                        formula1='"未开始,进行中,已完成,延期,风险"',
                        allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"G5:G{4+len(items)}")

    set_widths(ws, [10, 30, 18, 18, 12, 12, 12, 24])
    freeze(ws, "A5")


def build_budget(wb):
    ws = wb.create_sheet("05 预算明细")
    big_title(ws, "I", "预算明细", "目标 ROI ≥ 5×：首单 GMV ≥ 300 万 / 预算 ≤ 60 万")
    header_row(ws, 4,
               ["类别", "科目", "内容", "数量/标准", "单价(元)",
                "预算(元)", "实际支出(元)", "差异", "备注"])
    budget = [
        ("场地", "场地与展厅", "临港会议室+展厅+茶歇", "由临港承担", 0, 0, "由临港承担"),
        ("物料", "实体邀请函", "60 份", "60 份", 100, 6000, ""),
        ("物料", "资料袋印刷", "40 份", "40 份", 80, 3200, ""),
        ("物料", "桌签/议程印刷", "—", "1 套", 800, 800, ""),
        ("物料", "胸卡 / 名片夹", "—", "40 套", 30, 1200, ""),
        ("物料", "现场布置（KT 板/Banner）", "—", "1 套", 6800, 6800, ""),
        ("嘉宾", "学术嘉宾出席 + 交通", "1 位", "1 位", 30000, 30000, "姚教授"),
        ("伴手礼", "C 级礼盒（含权益券）", "30 × 600", "30 份", 600, 18000, ""),
        ("茶歇/餐", "下午茶", "30 人份", "30 份", 100, 3000, ""),
        ("茶歇/餐", "晚间轻餐", "30 人份", "30 份", 150, 4500, ""),
        ("影像", "摄影/摄像/精剪", "1 场", "1 场", 20000, 20000, ""),
        ("会务", "主持人 + 礼仪", "1 主持 + 2 礼仪", "1 场", 12000, 12000, ""),
        ("Demo", "日本真机/演示电脑/大屏租赁", "1 场", "1 场", 8000, 8000, ""),
        ("接待", "C 级 1V1 接送补贴", "10 人", "10 人", 1500, 15000, ""),
        ("传播", "媒体通稿 + 行业自媒体", "1 轮", "1 轮", 8000, 8000, ""),
        ("传播", "视频剪辑（闭门精华）", "1 条", "1 条", 4000, 4000, ""),
        ("机动", "机动 10%", "—", "—", 14000, 14000, ""),
    ]
    for i, (cat, item, det, qty, up, b, note) in enumerate(budget):
        r = 5 + i
        ws.cell(row=r, column=1, value=cat)
        ws.cell(row=r, column=2, value=item)
        ws.cell(row=r, column=3, value=det)
        ws.cell(row=r, column=4, value=qty)
        ws.cell(row=r, column=5, value=up)
        ws.cell(row=r, column=6, value=b)
        ws.cell(row=r, column=7, value=0)
        ws.cell(row=r, column=8, value=f"=F{r}-G{r}")
        ws.cell(row=r, column=9, value=note)
        for col in range(1, 10):
            c = ws.cell(row=r, column=col)
            c.font = Font(name="微软雅黑", size=10)
            c.alignment = Alignment(horizontal="left" if col <= 4
                                    else "right",
                                    vertical="center", wrap_text=True)
            c.border = thin()
            if i % 2:
                c.fill = PatternFill("solid", fgColor=LIGHT)
            if col in (5, 6, 7, 8):
                c.number_format = "#,##0"

    total_row = 5 + len(budget)
    write_row(ws, total_row,
              ["合计", "", "", "", "",
               f"=SUM(F5:F{total_row-1})",
               f"=SUM(G5:G{total_row-1})",
               f"=SUM(H5:H{total_row-1})", ""],
              fill=PRIMARY, bold=True)
    for col in range(1, 10):
        c = ws.cell(row=total_row, column=col)
        c.font = Font(name="微软雅黑", size=11, bold=True, color=WHITE)
        if col in (6, 7, 8):
            c.number_format = "#,##0"

    # ROI 测算
    section_banner(ws, total_row + 2, "ROI 测算", span=9, fill=GOLD)
    header_row(ws, total_row + 3,
               ["指标", "公式/来源", "目标", "实际", "状态", "", "", "", ""])
    roi = [
        ("总预算", "见上表合计", f"=F{total_row}", 0),
        ("T+90 首单数", "见 09 漏斗", 3, 0),
        ("T+90 首单 GMV (元)", "见 00 Dashboard", 3000000, 0),
        ("ROI 倍数",
         "= 首单 GMV / 总预算",
         f"=IFERROR(C{total_row+5}/C{total_row+3},0)",
         f"=IFERROR(D{total_row+5}/D{total_row+3},0)"),
    ]
    for i, (k, f, t, a) in enumerate(roi):
        r = total_row + 4 + i
        ws.cell(row=r, column=1, value=k)
        ws.cell(row=r, column=2, value=f)
        ws.cell(row=r, column=3, value=t)
        ws.cell(row=r, column=4, value=a)
        if i < 3:
            ws.cell(row=r, column=3).number_format = "#,##0"
            ws.cell(row=r, column=4).number_format = "#,##0"
        else:
            ws.cell(row=r, column=3).number_format = "0.00"
            ws.cell(row=r, column=4).number_format = "0.00"
        ws.cell(row=r, column=5,
                value=f'=IF(D{r}>=C{r},"达成",IF(D{r}>=C{r}*0.7,"进行中","风险"))')
        for col in range(1, 6):
            c = ws.cell(row=r, column=col)
            c.font = Font(name="微软雅黑", size=10)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin()
            if col == 1:
                c.fill = PatternFill("solid", fgColor=LIGHT)
                c.font = Font(name="微软雅黑", size=10, bold=True)

    set_widths(ws, [10, 24, 28, 14, 12, 14, 14, 12, 28])


def build_pre_event(wb):
    ws = wb.create_sheet("06 会前转化动作")
    big_title(ws, "I", "会前转化动作清单",
              "在 T-21 → T0 之间锁定客户期待 + 收集需求 + 预约会后深聊")
    header_row(ws, 4,
               ["#", "时间", "动作", "目的", "Owner", "对象",
                "可交付物", "完成", "备注"])
    actions = [
        [1, "T-21", "三方对齐邀约名单（含 NDA）",
         "建立 75 人邀请池基础", "协会项目总", "三方",
         "终版名单", "否", "去重 + 高优标星"],
        [2, "T-15", "首轮邀约（邀请函含「圆桌议题征集表」）",
         "提前收集客户痛点 → 形成定制化方案",
         "三方销售", "75 人邀请池",
         "邀请函 / 议题征集表 / 首轮记录", "否",
         "★ 关键：议题征集表直接决定 A 级线索"],
        [3, "T-10", "实体邀请函 + 伴手礼（C 级）",
         "提升 C 级到场仪式感", "协会会员部", "C 级 20 人",
         "邮寄确认 / 收件确认", "否", "EMS 顺丰 1V1"],
        [4, "T-7", "议程定稿 + 嘉宾介绍卡推送",
         "提升期待 + 锁定到场", "WeTest 市场", "全部邀请池",
         "议程 PDF / 嘉宾卡", "否", ""],
        [5, "T-5", "推送《WeTest 中日专项方案白皮书》",
         "建立技术认知 → 缩短 Demo 沟通时间",
         "WeTest 市场", "75 人",
         "白皮书 PDF / 推送记录", "否", ""],
        [6, "T-3", "TOP 10 高优 1V1 预热电话",
         "★ 预约会后 30 分钟深聊（强转化锚点）",
         "WeTest 大客户经理", "TOP 10 高优",
         "预热脚本 / 接通记录 / 预约时间表", "否",
         "每位至少 15 分钟"],
        [7, "T-2", "交通接待 / 动线图推送",
         "提升到场率", "协会会务", "已确认到场者",
         "推送记录", "否", ""],
        [8, "T-1", "二轮到场确认（短信/电话）",
         "防止临时缺席", "协会会员部", "已确认到场者",
         "确认表 V2", "否", ""],
        [9, "T-1", "现场彩排 + 应急预案 review",
         "确保现场零意外", "三方", "—",
         "彩排清单 / 应急预案", "否", ""],
    ]
    for i, r in enumerate(actions):
        write_row(ws, 5 + i, r, fill=LIGHT if i % 2 else None)

    dv = DataValidation(type="list", formula1='"是,否,延期"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"H5:H{4+len(actions)}")

    set_widths(ws, [5, 8, 38, 30, 18, 22, 30, 8, 22])
    freeze(ws, "A5")


def build_onsite(wb):
    ws = wb.create_sheet("07 会中转化抓手")
    big_title(ws, "I", "★ 会中转化抓手清单（执行 checklist）",
              "8 个抓手 → 把现场每个环节都做成销售场")
    header_row(ws, 4,
               ["#", "抓手", "落地形式", "对应议程", "目标",
                "Owner", "成功标准", "状态", "备注"])
    items = [
        [1, "扫码进群 + 行业标签",
         "签到二维码 + 企微自动打标",
         "13:00 签到",
         "100% 入群 + 100% 打标",
         "WeTest IT + 市场",
         "群人数 ≥ 25 / 标签覆盖率 100%",
         "未开始", ""],
        [2, "WeTest 体验角",
         "UDT 日本真机 / 海外众测看板 / CrashSight",
         "茶歇 + 圆桌间隙",
         "★ 让客户看见、摸到、用上",
         "WeTest 技术 SA",
         "≥ 15 人体验 + ≥ 8 条具体咨询",
         "未开始", "需 3 名 SA 在场"],
        [3, "限时权益券",
         "30 天免费 POC / 5 万元代金券",
         "技术落地环节宣布",
         "★ 制造紧迫感 → 推动签约",
         "WeTest 销售",
         "≥ 10 张领取 / ≥ 5 张激活意向",
         "未开始", "印在伴手礼内"],
        [4, "圆桌 1V1 承接",
         "每位发言客户配 1 名 WeTest 行业 SA",
         "15:50 圆桌",
         "★ 当场记录需求 → 24h 内输出 PPT",
         "WeTest 行业总监",
         "≥ 8 位客户发言 / 8 份纪要",
         "未开始", "纪要本编号 1-8"],
        [5, "闭门客户专属通道",
         "微信/钉钉直连中日团队 + 临港顾问",
         "圆桌结束前 5 分钟",
         "★ 建立长期可触达通道",
         "WeTest CSO",
         "≥ 20 人加入专属通道",
         "未开始", ""],
        [6, "晚间深聊预约",
         "高优客户 1V1 30 分钟",
         "16:30 → 18:00",
         "★ 线索 → 机会 升级",
         "WeTest 大客户经理",
         "≥ 12 位高优完成深聊",
         "未开始", ""],
        [7, "伴手礼附咨询券",
         "1V1 闭门技术咨询券 + 政策对接券",
         "离场",
         "建立二次触达由头",
         "协会会务",
         "100% 发放 + ≥ 8 张回访激活",
         "未开始", ""],
        [8, "现场拍摄客户证言",
         "短视频 + 案例素材",
         "圆桌结束后 / 深聊间隙",
         "为后续传播积累素材",
         "协会市场",
         "≥ 5 段证言 / ≥ 3 段可剪短视频",
         "未开始", "需提前沟通签字"],
    ]
    for i, r in enumerate(items):
        write_row(ws, 5 + i, r, fill=LIGHT if i % 2 else None)

    dv = DataValidation(type="list",
                        formula1='"未开始,进行中,已完成,未达标"',
                        allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"H5:H{4+len(items)}")

    set_widths(ws, [5, 24, 30, 18, 28, 22, 26, 12, 18])
    freeze(ws, "A5")


def build_post_event(wb):
    ws = wb.create_sheet("08 会后跟进SOP")
    big_title(ws, "M",
              "★ 会后跟进 SOP（7-30-60-90 天节奏）",
              "活动结束 48h 启动；每条 A/B 线索必须有 Owner + 截止日")

    section_banner(ws, 4, "A. ABC 三级分级标准", span=13, fill=PRIMARY)
    header_row(ws, 5,
               ["级别", "判定标准", "目标动作", "负责人",
                "复诊频次", "退出标准", "", "", "", "", "", "", ""])
    grades = [
        ["A · 高优",
         "C 级到场 + 圆桌发言 + 明确预算/项目 + 四大行业",
         "T+7 需求澄清；T+30 POC；T+90 首单",
         "WeTest 行业总监 + 大客户经理",
         "每周三同步", "签单 / 明确流失"],
        ["B · 中优",
         "业务 VP 到场 + 有兴趣但未明确预算",
         "T+15 二深聊；T+45 输出方案；T+90 POC",
         "WeTest 大客户经理", "每两周同步",
         "升 A / 降 C / 明确流失"],
        ["C · 培育",
         "代表到场 / 决策权弱 / 暂无项目",
         "月度内容运营；季度活动邀约",
         "WeTest 市场 + 内容运营", "月度同步",
         "升级 / 二次活动到场"],
    ]
    for i, r in enumerate(grades):
        write_row(ws, 6 + i, r + ["", "", "", "", "", "", ""],
                  fill=LIGHT if i % 2 else None)

    section_banner(ws, 11, "B. 7-30-60-90 跟进里程碑", span=13, fill=ACCENT)
    header_row(ws, 12,
               ["公司", "客户", "ABC", "Owner",
                "T+1 触达", "T+3 录入",
                "T+7 需求澄清", "T+15 方案",
                "T+30 POC/二深聊",
                "T+60 POC 进度",
                "T+90 签单/POC 验收",
                "当前阶段", "下一步动作"])

    # 留 30 行
    for i in range(30):
        r = 13 + i
        write_row(ws, r,
                  ["", "", "未分级", "", "未完成", "未完成",
                   "未完成", "未完成", "未完成", "未完成", "未完成",
                   "L1 到场", ""],
                  fill=LIGHT if i % 2 else None)

    # 数据验证
    dv_abc = DataValidation(type="list",
                            formula1='"A · 高优,B · 中优,C · 培育,未分级"',
                            allow_blank=True)
    ws.add_data_validation(dv_abc)
    dv_abc.add("C13:C42")

    dv_step = DataValidation(type="list",
                             formula1='"未完成,进行中,已完成,延期,N/A"',
                             allow_blank=True)
    ws.add_data_validation(dv_step)
    for col in ["E", "F", "G", "H", "I", "J", "K"]:
        dv_step.add(f"{col}13:{col}42")

    dv_stage = DataValidation(type="list",
                              formula1='"L1 到场,L2 强意向,L3 商机,L4 POC,L5 签单,流失"',
                              allow_blank=True)
    ws.add_data_validation(dv_stage)
    dv_stage.add("L13:L42")

    # 条件格式：已完成
    for col in ["E", "F", "G", "H", "I", "J", "K"]:
        ws.conditional_formatting.add(
            f"{col}13:{col}42",
            CellIsRule(operator="equal", formula=['"已完成"'],
                       fill=PatternFill("solid", fgColor="D4EDDA")))
        ws.conditional_formatting.add(
            f"{col}13:{col}42",
            CellIsRule(operator="equal", formula=['"延期"'],
                       fill=PatternFill("solid", fgColor="F8D7DA")))
    ws.conditional_formatting.add("C13:C42",
        CellIsRule(operator="equal", formula=['"A · 高优"'],
                   fill=PatternFill("solid", fgColor="FAD7A0"),
                   font=Font(bold=True, color="7E5109")))
    ws.conditional_formatting.add("L13:L42",
        CellIsRule(operator="equal", formula=['"L5 签单"'],
                   fill=PatternFill("solid", fgColor="D4EDDA"),
                   font=Font(bold=True, color="155724")))
    ws.conditional_formatting.add("L13:L42",
        CellIsRule(operator="equal", formula=['"流失"'],
                   fill=PatternFill("solid", fgColor="F8D7DA"),
                   font=Font(color="721C24")))

    set_widths(ws, [22, 12, 12, 18, 11, 11, 13, 11, 16, 14, 18, 14, 28])
    freeze(ws, "D13")


def build_funnel(wb):
    ws = wb.create_sheet("09 转化漏斗")
    big_title(ws, "H", "销售/转化漏斗（详细版）",
              "每周三 17:00 销售例会同步")
    header_row(ws, 4,
               ["阶段", "动作", "目标人数", "实际人数",
                "目标转化率（环比上阶段）", "实际转化率",
                "状态", "备注"])
    funnel = [
        ("L0 邀请池", "定向邀约", 75, 0, "—"),
        ("L1 到场（认知）", "签到 + 名片", 25, 0, "33%"),
        ("L2 强意向", "圆桌发言 / 1V1 预约", 15, 0, "60%"),
        ("L3 商机 (T+7)", "需求澄清会", 9, 0, "60%"),
        ("L4 POC (T+30)", "POC 立项", 5, 0, "55%"),
        ("L5 签单 (T+90)", "首单签订", 3, 0, "60%"),
    ]
    for i, (s, a, t, ac, rate) in enumerate(funnel):
        r = 5 + i
        ws.cell(row=r, column=1, value=s)
        ws.cell(row=r, column=2, value=a)
        ws.cell(row=r, column=3, value=t)
        ws.cell(row=r, column=4, value=ac)
        ws.cell(row=r, column=5, value=rate)
        if i == 0:
            ws.cell(row=r, column=6, value="—")
        else:
            ws.cell(row=r, column=6,
                    value=f"=IFERROR(D{r}/D{r-1},0)")
            ws.cell(row=r, column=6).number_format = "0.0%"
        ws.cell(row=r, column=7,
                value=f'=IF(D{r}>=C{r},"达成",IF(D{r}>=C{r}*0.7,"进行中","风险"))')
        ws.cell(row=r, column=8, value="")
        for col in range(1, 9):
            c = ws.cell(row=r, column=col)
            c.font = Font(name="微软雅黑", size=11)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin()
            if col == 1:
                c.fill = PatternFill("solid", fgColor=LIGHT)
                c.font = Font(name="微软雅黑", size=11, bold=True)

    # 总转化率
    r = 5 + len(funnel) + 1
    ws.cell(row=r, column=1, value="整体转化率（L0 → L5）").font = Font(
        name="微软雅黑", bold=True)
    ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor=GOLD)
    ws.cell(row=r, column=1).font = Font(name="微软雅黑", bold=True,
                                          color=WHITE)
    ws.cell(row=r, column=1).border = thin()
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws.cell(row=r, column=2,
            value=f"目标：=ROUND(D10/D5,3)").alignment = Alignment(
        horizontal="center", vertical="center")
    ws.cell(row=r, column=4,
            value=f"=IFERROR(D10/D5,0)")
    ws.cell(row=r, column=4).number_format = "0.0%"
    ws.cell(row=r, column=4).font = Font(name="微软雅黑", bold=True,
                                          color=ACCENT)

    # 状态条件格式
    ws.conditional_formatting.add(
        f"G5:G{5+len(funnel)-1}",
        CellIsRule(operator="equal", formula=['"达成"'],
                   fill=PatternFill("solid", fgColor="D4EDDA")))
    ws.conditional_formatting.add(
        f"G5:G{5+len(funnel)-1}",
        CellIsRule(operator="equal", formula=['"风险"'],
                   fill=PatternFill("solid", fgColor="F8D7DA")))

    section_banner(ws, 14, "★ 销售周例会模板（每周三 17:00）",
                   span=8, fill=PRIMARY)
    weekly = [
        "1. 本周新增 A/B 级线索数 vs 目标",
        "2. A 级线索逐条 review：当前阶段、阻塞、下一步、Owner",
        "3. B 级线索抽样 review（≥ 30%）",
        "4. POC 进度与阻塞（A 级专项）",
        "5. 流失线索原因分析（≥ 1 条结构化复盘）",
        "6. 下周关键动作 + Owner + 截止日",
        "7. 需协会/临港支持事项",
    ]
    for i, t in enumerate(weekly, 15):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=8)
        c = ws.cell(row=i, column=1, value=t)
        c.font = Font(name="微软雅黑", size=10, color=DARK)
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[i].height = 22

    set_widths(ws, [16, 22, 14, 14, 22, 16, 12, 28])
    freeze(ws, "A5")


def build_kpi(wb):
    ws = wb.create_sheet("10 KPI与考核")
    big_title(ws, "G", "KPI 与销售考核挂钩", "活动效果不是 “办了” 而是 “成单”")
    header_row(ws, 4,
               ["维度", "指标", "目标", "权重",
                "实际", "完成率", "得分"])
    kpis = [
        ("邀约", "邀请池规模", 75, 0.08),
        ("邀约", "C 级/VP 占比 (%)", 70, 0.08),
        ("到场", "到场率 (%)", 33, 0.08),
        ("到场", "实际到场人数", 25, 0.08),
        ("现场", "圆桌发言客户数", 8, 0.08),
        ("现场", "1V1 深聊预约数", 12, 0.08),
        ("转化", "T+7 需求澄清会", 9, 0.12),
        ("转化", "T+30 POC 立项", 5, 0.12),
        ("转化", "T+90 首单数", 3, 0.15),
        ("收入", "首单 GMV (元)", 3000000, 0.10),
        ("品牌", "媒体露出次数", 10, 0.03),
    ]
    for i, (d, k, t, w) in enumerate(kpis):
        r = 5 + i
        ws.cell(row=r, column=1, value=d)
        ws.cell(row=r, column=2, value=k)
        ws.cell(row=r, column=3, value=t)
        ws.cell(row=r, column=4, value=w)
        ws.cell(row=r, column=4).number_format = "0%"
        ws.cell(row=r, column=5, value=0)
        ws.cell(row=r, column=6, value=f"=IFERROR(E{r}/C{r},0)")
        ws.cell(row=r, column=6).number_format = "0.0%"
        ws.cell(row=r, column=7, value=f"=MIN(F{r},1.2)*D{r}*100")
        ws.cell(row=r, column=7).number_format = "0.0"
        for col in range(1, 8):
            c = ws.cell(row=r, column=col)
            c.font = Font(name="微软雅黑", size=10)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin()
            if col == 1:
                c.fill = PatternFill("solid", fgColor=LIGHT)

    # 总分
    total_r = 5 + len(kpis)
    write_row(ws, total_r,
              ["", "总分（满分 100，120% 封顶）", "",
               "=SUM(D5:D{0})".format(total_r - 1),
               "", "",
               f"=SUM(G5:G{total_r-1})"],
              fill=PRIMARY, bold=True)
    for col in range(1, 8):
        c = ws.cell(row=total_r, column=col)
        c.font = Font(name="微软雅黑", size=11, bold=True, color=WHITE)
    ws.cell(row=total_r, column=4).number_format = "0%"
    ws.cell(row=total_r, column=7).number_format = "0.0"

    section_banner(ws, total_r + 2, "考核挂钩规则", span=7, fill=GOLD)
    rules = [
        "≥ 90 分：项目奖金 100%；销售提成正常上浮。",
        "70-89 分：项目奖金 80%；销售按目标提成。",
        "50-69 分：项目奖金 50%；销售提成按 70% 计。",
        "< 50 分：项目奖金 0；触发复盘机制 + 责任倒查。",
    ]
    for i, t in enumerate(rules, total_r + 3):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=7)
        c = ws.cell(row=i, column=1, value=t)
        c.font = Font(name="微软雅黑", size=10, color=DARK)
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[i].height = 22

    set_widths(ws, [10, 28, 18, 10, 14, 12, 10])
    freeze(ws, "A5")


def build_matrix(wb):
    ws = wb.create_sheet("11 行业方案矩阵")
    big_title(ws, "H", "四大行业 × 方案矩阵",
              "每位 A/B 级客户进入时，先对齐该矩阵选首选方案包")
    header_row(ws, 4,
               ["行业", "出海/入华主诉求", "首选方案包",
                "次选方案包", "标杆案例", "首单切入产品",
                "话术钩子", "Owner"])
    rows = [
        ["金融", "数据跨境 + 风控 + 隐私 + 防篡改",
         "③ 跨境合规安全包", "① 日本出海加速包",
         "银联云小程序安全扫描",
         "小程序合规扫描 + 数据跨境审计",
         "「过审无忧」「合规托底」", "WeTest 金融组"],
        ["游戏", "日本市场首发 + 网络 / 设备 + 真金支付 + 防作弊",
         "① 日本出海加速包", "③ 跨境合规安全包",
         "跨境电商 UDT + 海外众测",
         "海外众测 + UDT 日本真机",
         "「即开即用」「免封号」", "WeTest 游戏组"],
        ["电商/零售", "本地支付 + 防黑产 + 跨境物流 + 风控封号",
         "① + ③", "② 日资入华质量包",
         "跨境电商 UDT",
         "海外支付验收 + 安全网关",
         "「真金支付测试」「免封号」", "WeTest 电商组"],
        ["泛互联网", "本地化 / 多语言 / 网络 / SDK 兼容",
         "① 日本出海加速包", "PerfDog 单品",
         "PerfDog + 海外众测",
         "PerfDog + 兼容/众测",
         "「即插即用」「全平台」", "WeTest 互联网组"],
        ["日资入华", "小程序兼容 / 微信隐私合规 / 移动支付适配",
         "② 日资入华质量包", "③ 跨境合规安全包",
         "微信小程序合规扫描",
         "小程序兼容 + 隐私合规",
         "「中国独有生态质控」", "WeTest 中日团队"],
    ]
    for i, r in enumerate(rows):
        write_row(ws, 5 + i, r, fill=LIGHT if i % 2 else None)

    set_widths(ws, [12, 30, 22, 22, 22, 24, 22, 18])


def build_risk(wb):
    ws = wb.create_sheet("12 风险预案")
    big_title(ws, "H", "风险评估与应急预案",
              "概率 × 影响 → 自动判定优先级")
    header_row(ws, 4,
               ["#", "类别", "风险", "概率(1-5)",
                "影响(1-5)", "风险值", "预案", "Owner"])
    rows = [
        [1, "到场率", "目标客户临时缺席", 3, 5,
         "T-2/T-1 二轮确认 + 备选 5 人 + 协会内部直播链路", "协会"],
        [2, "嘉宾", "学术嘉宾日程冲突", 2, 5,
         "锁定 1 位 backup（同领域权威） + 预演 PPT", "协会"],
        [3, "设备", "日本真机网络不稳定", 3, 3,
         "本地录屏备份 Demo + 双链路网络冗余", "WeTest IT"],
        [4, "转化", "WeTest 销售跟进不及时",
         3, 5,
         "T+1 强制 1V1 触达 SOP + 周例会督办 + KPI 考核挂钩",
         "WeTest 销售运营"],
        [5, "合规", "闭门内容外泄", 2, 3,
         "签到处 NDA 签署 + 现场无直播 + 协会统一精剪发布",
         "协会法务"],
        [6, "不可抗力", "线下临时改线上", 2, 5,
         "腾讯会议直播方案 + 物料邮寄 + 远程 Demo 备份",
         "三方"],
        [7, "签到", "签到拥堵 / 二维码失败",
         2, 2,
         "纸质签到表备份 + 2 名礼仪 + 备用 Wi-Fi",
         "协会会务"],
        [8, "议程", "圆桌冷场 / 客户不愿发言",
         3, 4,
         "提前与 4 位行业代表对齐 ”首发言“ + 主持人引导脚本",
         "WeTest 行业总监"],
    ]
    for i, r in enumerate(rows):
        write_row(ws, 5 + i, r[:5] + [f"=D{5+i}*E{5+i}"] + r[5:],
                  fill=LIGHT if i % 2 else None)
        ws.cell(row=5 + i, column=6).number_format = "0"

    # 风险值色阶
    ws.conditional_formatting.add(
        f"F5:F{4+len(rows)}",
        ColorScaleRule(start_type="num", start_value=1, start_color="D4EDDA",
                       mid_type="num", mid_value=10, mid_color="FFF3CD",
                       end_type="num", end_value=25, end_color="F8D7DA"))

    set_widths(ws, [5, 12, 28, 10, 10, 10, 50, 18])


def build_review(wb):
    ws = wb.create_sheet("13 复盘记录")
    big_title(ws, "H", "复盘记录模板",
              "T+3 / T+15 / T+30 / T+90 四阶段复盘")
    section_banner(ws, 4,
                   "T+3 内部快速复盘（90 分钟，三方）", span=8, fill=PRIMARY)
    header_row(ws, 5,
               ["维度", "目标", "实际", "GAP",
                "原因分析", "改进动作", "Owner", "截止日"])
    rows = [
        ["邀约质量", "C 级 / VP 占比 ≥ 70%", "", "", "", "", "", ""],
        ["到场率", "≥ 33%", "", "", "", "", "", ""],
        ["现场执行", "8 个抓手 100% 触发", "", "", "", "", "", ""],
        ["销售触达", "T+1 1V1 触达 100%", "", "", "", "", "", ""],
        ["客户满意度", "NPS ≥ 8", "", "", "", "", "", ""],
    ]
    for i, r in enumerate(rows):
        write_row(ws, 6 + i, r, fill=LIGHT if i % 2 else None)

    section_banner(ws, 12,
                   "T+15 客户视角复盘（抽样 5 位 A 级回访）", span=8,
                   fill=GOLD)
    header_row(ws, 13,
               ["客户", "好评点", "改进点", "对方案包评价",
                "下一步", "Owner", "回访日", "状态"])
    for i in range(5):
        write_row(ws, 14 + i, ["", "", "", "", "", "", "", ""],
                  fill=LIGHT if i % 2 else None)

    section_banner(ws, 20,
                   "T+30 转化中期复盘 + T+90 收单复盘（合并模板）",
                   span=8, fill=ACCENT)
    header_row(ws, 21,
               ["阶段", "目标", "实际", "ROI", "成单亮点",
                "流失原因", "下一场迭代建议", "Owner"])
    write_row(ws, 22,
              ["T+30 中期", "POC ≥ 5", "", "", "", "", "", ""],
              fill=LIGHT)
    write_row(ws, 23,
              ["T+90 收单", "首单 ≥ 3 / GMV ≥ 300 万",
               "", "", "", "", "", ""])

    set_widths(ws, [14, 24, 16, 16, 28, 28, 14, 12])


def main():
    wb = Workbook()
    build_dashboard(wb)
    build_timeline(wb)
    build_invitee(wb)
    build_agenda(wb)
    build_logistics(wb)
    build_budget(wb)
    build_pre_event(wb)
    build_onsite(wb)
    build_post_event(wb)
    build_funnel(wb)
    build_kpi(wb)
    build_matrix(wb)
    build_risk(wb)
    build_review(wb)

    # 设置默认字体（隐式由我们逐 cell 控制）
    out = ("/workspace/deliverables/"
           "临港x WeTest 中日闭门会_转化运营管理表_V1.0.xlsx")
    wb.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
