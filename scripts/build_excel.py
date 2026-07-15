# -*- coding: utf-8 -*-
"""
生成《复旦 × 腾讯云 2026 秋季 AI 大会》执行计划 Excel。
含：活动总览 / 主论坛议程 / 分论坛议程 / 嘉宾邀请 / 黑客松 /
互动体验 / 预算明细 / 赞助权益 / 执行倒排期 / 物料清单 / KPI。
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

FONT = "Microsoft YaHei"
NAVY = "0B1F3A"; TENCENT = "006EFF"; CYAN = "22C1E6"; GOLD = "F5B50C"
FUDAN = "C01F2E"; LIGHT = "EAF1FB"; GREY = "8A97AD"; ROW2 = "F4F7FC"
WHITE = "FFFFFF"; INK = "1B2A44"

thin = Side(style="thin", color="D5DEEB")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()


def hfill(color):
    return PatternFill("solid", fgColor=color)


def style_title(ws, text, ncol, color=TENCENT):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    c = ws.cell(1, 1, text)
    c.font = Font(name=FONT, size=15, bold=True, color=WHITE)
    c.fill = hfill(color)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30


def style_sub(ws, text, ncol, row=2, color=NAVY):
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


def data_rows(ws, start, rows, aligns=None, base_h=22):
    for i, r in enumerate(rows):
        ri = start + i
        for j, val in enumerate(r, start=1):
            c = ws.cell(ri, j, val)
            c.font = Font(name=FONT, size=10, color=INK)
            c.fill = hfill(WHITE if i % 2 == 0 else ROW2)
            al = (aligns[j-1] if aligns else "left")
            c.alignment = Alignment(horizontal=al, vertical="center",
                                    wrap_text=True, indent=(1 if al == "left" else 0))
            c.border = BORDER
        ws.row_dimensions[ri].height = base_h
    return start + len(rows)


def freeze_and_filter(ws, header_row_idx, ncol):
    ws.freeze_panes = ws.cell(header_row_idx + 1, 1)
    ws.auto_filter.ref = f"A{header_row_idx}:{get_column_letter(ncol)}{header_row_idx}"


# ============================================================ 1. 活动总览
ws = wb.active; ws.title = "1.活动总览"
style_title(ws, "复旦大学住房政策研究中心 · 杨浦区科技企业联合会  ×  腾讯云｜2026 秋季 AI 主题大会", 4)
style_sub(ws, "执行计划表（建议稿）｜以 5·22「2026 人工智能商业化落地峰会」为蓝图｜聚焦 Agent · 算力 · Token", 4)
hr = 4
header_row(ws, hr, ["项目", "内容", "项目", "内容"], [18, 40, 18, 40])
ov = [
    ("活动名称", "2026 秋季 AI 主题大会（暂定）", "主题", "智能体涌现 · 商业落地"),
    ("主办单位", "复旦大学住房政策研究中心、杨浦区科技企业联合会", "联合主办", "腾讯云"),
    ("建议时间", "2026 年 8—9 月（单日，可延展至 1.5 天）", "建议地点", "上海·杨浦（高校 / 园区会场）"),
    ("核心议题", "Agent 智能体、算力、Token 经济", "活动形式", "1 主论坛 + 多分论坛 + 黑客松 + 体验层"),
    ("预计规模", "现场 ≥ 500 人；线上 ≥ 50,000 人次", "预算区间", "约 41—81 万元（可由赞助/置换分摊）"),
    ("核心环节", "大咖卡位分享 / 黑客松 / 创意集市 / 集章集票 / 开放麦", "知识产出", "联合白皮书、落地榜单、案例库、社群"),
    ("活动目标", "品牌共建 / 商业落地 / 人才生态 / 知识闭环", "呈现物", "策划方案 PPT + 执行计划 Excel"),
]
data_rows(ws, hr + 1, ov, aligns=["center", "left", "center", "left"], base_h=30)

# ============================================================ 2. 主论坛议程
ws = wb.create_sheet("2.主论坛议程")
style_title(ws, "主论坛议程（上午）", 5, TENCENT)
style_sub(ws, "顶级大咖卡位分享 · 趋势发布 · 重磅签约 · 颁证", 5)
hr = 4
header_row(ws, hr, ["时间", "时长", "环节", "拟邀嘉宾 / 主体", "备注"],
           [14, 8, 30, 28, 26])
main = [
    ("09:00-09:30", "30'", "签到 · 参会护照领取 · 展区开放", "全体 / 志愿者", "动线引导、媒体接待"),
    ("09:30-09:50", "20'", "开场致辞", "主办双方领导 + 政府/园区", "双方各 1 位 + 政府 1 位"),
    ("09:50-10:25", "35'", "主旨演讲一：AI 与产业趋势", "复旦教授 / 院士级学者", "学术卡位、权威背书"),
    ("10:25-11:00", "35'", "主旨演讲二：腾讯云大模型/Agent/算力战略", "腾讯云技术高管", "腾讯云核心卡位 + 发布"),
    ("11:00-11:15", "15'", "重磅发布 / 趋势报告", "腾讯云 + 复旦联合", "趋势/产品/白皮书预告"),
    ("11:15-11:50", "35'", "高峰对话：Agent 商业化圆桌", "学者+高管+企业+投资人", "主持人主持，4-5 位嘉宾"),
    ("11:50-12:00", "10'", "战略签约 / 联合倡议 / 颁证仪式", "主办双方 + 合作方", "证书与签约同步拍摄"),
]
data_rows(ws, hr + 1, main, aligns=["center", "center", "left", "left", "left"], base_h=28)
freeze_and_filter(ws, hr, 5)

# ============================================================ 3. 分论坛议程
ws = wb.create_sheet("3.分论坛议程")
style_title(ws, "分论坛议程（下午 · 三场平行）", 6, CYAN)
style_sub(ws, "13:30—17:00 三个分论坛平行进行，参会者按兴趣自由选择", 6)
hr = 4
header_row(ws, hr, ["分论坛", "时间", "议题", "形式", "拟邀嘉宾方向", "产出"],
           [16, 13, 26, 14, 22, 18])
sub = [
    ("A · Agent 落地", "13:30-14:10", "从 PoC 到生产级 Agent", "主题演讲", "头部企业技术负责人", "案例集"),
    ("A · Agent 落地", "14:10-14:50", "多智能体协作与编排", "主题演讲", "Agent 框架团队", "案例集"),
    ("A · Agent 落地", "14:50-15:30", "企业落地 ROI 实战", "主题演讲", "甲方数字化负责人", "案例集"),
    ("A · Agent 落地", "15:40-16:40", "Agent 商业化圆桌", "圆桌", "厂商 + 甲方 + 投资", "观点纪要"),
    ("B · 算力基础设施", "13:30-14:10", "训练/推理算力优化", "主题演讲", "云/芯片架构师", "实践指南"),
    ("B · 算力基础设施", "14:10-14:50", "成本、调度与弹性", "主题演讲", "平台/运维专家", "实践指南"),
    ("B · 算力基础设施", "14:50-15:30", "异构与国产化算力", "主题演讲 + Demo", "算力厂商", "实践指南"),
    ("B · 算力基础设施", "15:40-16:40", "算力降本圆桌", "圆桌", "云厂商 + 用户", "观点纪要"),
    ("C · Token 与应用", "13:30-14:10", "大模型成本结构与定价", "主题演讲", "模型/商业化专家", "白皮书章节"),
    ("C · Token 与应用", "14:10-14:50", "Token 消耗优化策略", "主题演讲", "应用架构师", "白皮书章节"),
    ("C · Token 与应用", "14:50-16:40", "创业项目路演 + 投资点评", "路演", "初创团队 + 投资人", "项目榜单"),
]
data_rows(ws, hr + 1, sub, aligns=["left", "center", "left", "center", "left", "left"], base_h=26)
freeze_and_filter(ws, hr, 6)

# ============================================================ 4. 嘉宾邀请清单
ws = wb.create_sheet("4.嘉宾邀请清单")
style_title(ws, "嘉宾邀请清单与跟进（模板）", 8, FUDAN)
style_sub(ws, "卡位嘉宾设 A/B 备选；负责人持续跟进确认状态", 8)
hr = 4
header_row(ws, hr, ["序号", "拟邀嘉宾/单位", "类别", "环节/卡位", "邀请方", "对接人", "状态", "备注"],
           [6, 24, 14, 20, 14, 12, 12, 20])
guests = [
    (1, "复旦大学 教授/学者", "学术大咖", "主论坛主旨演讲一", "复旦", "（待填）", "待邀约", "权威背书"),
    (2, "腾讯云 技术高管", "产业大咖", "主论坛主旨演讲二", "腾讯云", "（待填）", "待邀约", "战略发布"),
    (3, "头部企业 CEO/CTO", "企业代表", "主论坛圆桌", "双方", "（待填）", "待邀约", "落地案例"),
    (4, "知名投资人", "投资视角", "主论坛圆桌", "双方", "（待填）", "待邀约", "资本视角"),
    (5, "政府 / 园区领导", "政府", "开场致辞", "复旦/联合会", "（待填）", "待邀约", "政策环境"),
    (6, "Agent 方向专家", "技术专家", "分论坛 A", "双方", "（待填）", "待邀约", ""),
    (7, "算力 / 架构专家", "技术专家", "分论坛 B", "腾讯云", "（待填）", "待邀约", ""),
    (8, "模型商业化专家", "技术专家", "分论坛 C", "双方", "（待填）", "待邀约", ""),
    (9, "黑客松评委 ×3-5", "评委", "黑客松路演", "双方", "（待填）", "待邀约", "技术+投资+学术"),
    (10, "主持人 ×2", "主持", "主论坛/分论坛", "复旦", "（待填）", "待邀约", "专业主持"),
]
data_rows(ws, hr + 1, guests, aligns=["center", "left", "center", "left", "center", "center", "center", "left"], base_h=24)
freeze_and_filter(ws, hr, 8)

# ============================================================ 5. 黑客松
ws = wb.create_sheet("5.黑客松")
style_title(ws, "黑客松（Hackathon）执行计划", 4, FUDAN)
style_sub(ws, "Agent 实战挑战赛｜赛前招募 → 现场冲刺 → 路演评审", 4)
hr = 4
header_row(ws, hr, ["模块", "内容", "负责方", "备注"], [16, 44, 16, 24])
hack = [
    ("赛题方向", "基于腾讯云能力构建可落地 Agent 应用（效率办公/行业垂类/创意应用）", "双方", "赛题 T-6 周公布"),
    ("参赛规模", "20—30 支队伍 / 80—120 人", "复旦组织", "线上预报名审核"),
    ("赛程", "T-6 招募 → T-2 选手确认 → 现场 8—12h 冲刺 → 路演", "双方", "全天并行主会场"),
    ("技术资源", "腾讯云算力 / API / Token 额度 + 技术导师驻场", "腾讯云", "代金券/Token 包"),
    ("场地资源", "比赛场地、网络、电源、餐饮、志愿者", "复旦", "独立分区"),
    ("评审维度", "创新性 / 技术完成度 / 商业可行性 / 落地价值 / 现场路演", "评委组", "各 20 分"),
    ("奖项设置", "冠/亚/季军 + 最佳创意 + 最佳商业潜力", "双方", "证书+奖杯"),
    ("奖励", "奖金 + 腾讯云资源代金券 + 孵化/投资对接", "双方", "优秀作品入案例库"),
    ("产出", "作品纳入会后案例集与榜单，二次传播", "双方", "知识闭环"),
]
data_rows(ws, hr + 1, hack, aligns=["center", "left", "center", "left"], base_h=26)
freeze_and_filter(ws, hr, 4)

# ============================================================ 6. 互动体验
ws = wb.create_sheet("6.互动体验区")
style_title(ws, "互动体验区清单（贯穿全场）", 5, GOLD)
style_sub(ws, "展览/展位 · 创意集市 · 集票集章 · 开放麦", 5)
hr = 4
header_row(ws, hr, ["板块", "内容", "目标", "负责方", "物料/资源"],
           [16, 34, 20, 14, 24])
exp = [
    ("展览区/腾讯云展区", "腾讯云能力展示 + Demo 体验 + 答疑", "技术心智 + 转化", "腾讯云", "展架/屏幕/设备"),
    ("企业展位", "合作企业标准/精品展位分级", "招商收益", "双方", "展位/桌椅/电源"),
    ("创意集市", "AI 文创/周边/初创产品试用摊位", "氛围 + 自由交流", "复旦", "摊位/水牌"),
    ("集票·集章打卡", "参会护照 + 打卡地图，逛展听会盖章集票", "全场动线引导", "复旦", "护照/印章/兑奖台"),
    ("礼品/抽奖", "集满兑换礼品 + 抽奖资格", "留存与活跃", "双方", "礼品/抽奖系统"),
    ("临时开放麦 Open Mic", "随时上台 3—5 分钟即兴分享/自荐/招募", "即兴交流连接", "双方", "舞台/音响/计时"),
    ("茶歇社交区", "茶歇 + 自由社交 + 名片墙", "商机撮合", "复旦", "茶歇/桌椅"),
]
data_rows(ws, hr + 1, exp, aligns=["center", "left", "left", "center", "left"], base_h=26)
freeze_and_filter(ws, hr, 5)

# ============================================================ 7. 预算明细
ws = wb.create_sheet("7.预算明细")
style_title(ws, "预算明细（建议区间，单位：万元）", 5, TENCENT)
style_sub(ws, "可通过赞助权益分级、生态共建与资源置换（腾讯云算力/Token）显著降低净支出", 5)
hr = 4
header_row(ws, hr, ["序号", "费用项", "下限", "上限", "说明"], [6, 24, 10, 10, 36])
bud = [
    (1, "场地 & 搭建", 8, 15, "主会场+分会场+黑客松+体验区"),
    (2, "舞台/视觉/设备", 6, 12, "主视觉、舞美、灯光音响、直播"),
    (3, "嘉宾差旅接待", 4, 8, "差旅、住宿、接待"),
    (4, "黑客松", 6, 12, "奖金、算力/Token、运营、物料"),
    (5, "餐饮 & 茶歇", 5, 10, "午餐、茶歇、嘉宾餐"),
    (6, "传播 & 物料", 5, 10, "设计、投放、媒体、印刷"),
    (7, "互动体验/礼品", 3, 6, "集市、礼品、抽奖、护照印章"),
    (8, "执行 & 不可预见", 4, 8, "执行人力、保险、应急"),
]
end = data_rows(ws, hr + 1, bud, aligns=["center", "left", "center", "center", "left"], base_h=24)
# 合计行
tr = end
ws.cell(tr, 1, "")
ws.merge_cells(start_row=tr, start_column=1, end_row=tr, end_column=2)
c = ws.cell(tr, 1, "合计（参考区间）"); c.font = Font(name=FONT, size=11, bold=True, color=WHITE)
c.fill = hfill(NAVY); c.alignment = Alignment(horizontal="center", vertical="center"); c.border = BORDER
for col, formula in [(3, f"=SUM(C{hr+1}:C{end-1})"), (4, f"=SUM(D{hr+1}:D{end-1})")]:
    cc = ws.cell(tr, col, formula)
    cc.font = Font(name=FONT, size=11, bold=True, color=GOLD)
    cc.fill = hfill(NAVY); cc.alignment = Alignment(horizontal="center", vertical="center"); cc.border = BORDER
cc = ws.cell(tr, 5, "净支出可由分级赞助/资源置换进一步下探")
cc.font = Font(name=FONT, size=10, italic=True, color=WHITE)
cc.fill = hfill(NAVY); cc.alignment = Alignment(horizontal="left", vertical="center", indent=1); cc.border = BORDER
ws.row_dimensions[tr].height = 26

# ============================================================ 8. 赞助权益
ws = wb.create_sheet("8.赞助权益")
style_title(ws, "招商与赞助权益分级（延续 5·22 模式）", 5, FUDAN)
style_sub(ws, "延续为腾讯云提供品牌展示/证书/展位的成功做法，分级吸引生态伙伴共建", 5)
hr = 4
header_row(ws, hr, ["权益项", "联合主办/首席", "钻石赞助", "生态合作", "备注"],
           [20, 18, 16, 16, 20])
spon = [
    ("主论坛卡位/发布", "核心卡位+发布窗口", "—", "—", "黄金时段"),
    ("分论坛冠名/演讲", "✔", "✔（冠名）", "—", ""),
    ("主视觉品牌露出", "首位", "✔", "Logo", "贯穿物料"),
    ("权威证书", "联合颁发", "荣誉证书", "参会证书", "权威背书"),
    ("展位/集市", "精品展位+集市位", "标准展位", "摊位", "现场展示"),
    ("黑客松权益", "冠名权", "导师席", "—", ""),
    ("开放麦优先", "✔", "✔", "优先席", ""),
    ("白皮书署名", "联合署名", "鸣谢", "鸣谢", "知识资产"),
    ("传播权益", "全渠道头条", "通稿", "社群推荐", ""),
]
data_rows(ws, hr + 1, spon, aligns=["left", "center", "center", "center", "left"], base_h=24)
freeze_and_filter(ws, hr, 5)

# ============================================================ 9. 执行倒排期
ws = wb.create_sheet("9.执行倒排期")
style_title(ws, "执行倒排期与责任分工", 6, TENCENT)
style_sub(ws, "以活动日 T-0 倒排；双方设对接人，定期联合例会对齐进度", 6)
hr = 4
header_row(ws, hr, ["阶段", "时间节点", "关键任务", "复旦/联合会", "腾讯云", "状态"],
           [12, 12, 30, 16, 16, 10])
tl = [
    ("立项", "T-8 周", "确认主题/规模/预算/分工，成立联合筹备组", "牵头", "确认", "待启动"),
    ("启动", "T-7 周", "锁定时间地点，启动招商与嘉宾初邀", "场地/嘉宾", "招商/嘉宾", "待启动"),
    ("框架", "T-6 周", "主视觉+议程框架，黑客松赛题与招募启动", "议程/组织", "赛题/技术", "待启动"),
    ("确认", "T-4 周", "嘉宾确认、赞助签约、传播预热上线", "嘉宾/媒体", "赞助/渠道", "待启动"),
    ("定稿", "T-3 周", "议程定稿、报名开放、物料制作", "物料/报名", "内容/物料", "待启动"),
    ("筹备", "T-2 周", "黑客松选手确认、现场动线与彩排方案", "动线/志愿", "导师/资源", "待启动"),
    ("彩排", "T-1 周", "全流程彩排、物料到位、应急预案确认", "现场统筹", "技术保障", "待启动"),
    ("执行", "T-0", "活动执行、现场指挥、直播传播", "现场总控", "技术/直播", "待启动"),
    ("收口", "T+2 周", "白皮书/榜单、复盘、商机跟进", "学术编纂", "内容/案例", "待启动"),
]
data_rows(ws, hr + 1, tl, aligns=["center", "center", "left", "center", "center", "center"], base_h=26)
freeze_and_filter(ws, hr, 6)

# ============================================================ 10. 物料清单
ws = wb.create_sheet("10.物料清单")
style_title(ws, "物料与场地清单（Checklist）", 5, CYAN)
style_sub(ws, "用于现场执行核对，状态可勾选：待办/进行中/完成", 5)
hr = 4
header_row(ws, hr, ["类别", "物料/项目", "数量/规格", "负责方", "状态"],
           [14, 30, 18, 14, 12])
mat = [
    ("主视觉", "主背景板 / 主KV / 导视系统", "按场地", "双方", "待办"),
    ("舞台", "舞台/灯光/音响/LED 屏", "主会场", "复旦", "待办"),
    ("直播", "直播设备/导播/推流", "1 套", "腾讯云", "待办"),
    ("展区", "展架/展位/电源/网络", "按招商", "双方", "待办"),
    ("证书", "卡位/荣誉/参会证书", "按名单", "复旦", "待办"),
    ("互动", "参会护照/打卡地图/印章/兑奖台", "按规模", "复旦", "待办"),
    ("开放麦", "舞台/麦克风/计时器/水牌", "1 区", "双方", "待办"),
    ("黑客松", "桌椅/插排/网络/茶水/计时", "按队伍", "复旦", "待办"),
    ("礼品", "伴手礼/抽奖礼品/周边", "按规模", "双方", "待办"),
    ("接待", "签到台/胸牌/指引/志愿者服", "按规模", "复旦", "待办"),
]
data_rows(ws, hr + 1, mat, aligns=["center", "left", "center", "center", "center"], base_h=24)
freeze_and_filter(ws, hr, 5)

# ============================================================ 11. KPI
ws = wb.create_sheet("11.KPI与产出")
style_title(ws, "KPI 与知识产出目标", 4, GOLD)
style_sub(ws, "用于事前对齐预期、事后复盘评估", 4)
hr = 4
header_row(ws, hr, ["维度", "指标", "目标值", "说明"], [16, 26, 16, 30])
kpi = [
    ("到场", "现场参会人数", "≥ 500 人", "高质量定向邀约"),
    ("线上", "直播观看人次", "≥ 50,000", "双方渠道联合直播"),
    ("传播", "媒体/自媒体报道", "≥ 30 篇", "预热+爆发+长尾"),
    ("社群", "沉淀参会名单", "≥ 1,000", "商机池蓄水"),
    ("黑客松", "报名队伍", "≥ 25 队", "开发者生态"),
    ("招商", "赞助/生态伙伴", "≥ 8 家", "分级权益"),
    ("产出", "联合白皮书", "1 部", "双方署名"),
    ("产出", "落地案例/榜单", "≥ 1 份", "可二次传播"),
    ("满意度", "嘉宾/参会满意度", "≥ 90%", "问卷回收"),
]
data_rows(ws, hr + 1, kpi, aligns=["center", "left", "center", "left"], base_h=24)
freeze_and_filter(ws, hr, 4)

# ---------------- 打印页面设置：横向、按宽度铺满一页 ----------------
for sh in wb.worksheets:
    sh.page_setup.orientation = "landscape"
    sh.page_setup.fitToWidth = 1
    sh.page_setup.fitToHeight = 0
    sh.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    sh.page_margins.left = sh.page_margins.right = 0.3
    sh.page_margins.top = sh.page_margins.bottom = 0.4

# ---------------- 保存 ----------------
out_dir = "/workspace/deliverables"
os.makedirs(out_dir, exist_ok=True)
path = os.path.join(out_dir, "复旦x腾讯云_2026秋季AI大会_执行计划表.xlsx")
wb.save(path)
print("Saved:", path, "sheets:", len(wb.sheetnames))
