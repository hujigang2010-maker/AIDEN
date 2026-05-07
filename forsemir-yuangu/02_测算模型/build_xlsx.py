"""Build the Yuangu cooperation revenue model `.xlsx`.

The workbook contains:
  Sheet 1  封面与说明
  Sheet 2  收入结构总览（保守 / 基础 / 乐观三档）
  Sheet 3  月费 Retainer 测算（团队成本反算）
  Sheet 4  招商佣金阶梯测算（按面积分档）
  Sheet 5  活动 + 奖项激励测算
  Sheet 6  股权分红与三年累计预测
  Sheet 7  敏感性分析（招商进度 × 月费）
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = Path(__file__).with_name("合作收益测算模型.xlsx")

PRIMARY = "FF142C5E"
ACCENT = "FFF27E2D"
LIGHT = "FFEAEEF5"
ALT = "FFF8F9FB"
WHITE = "FFFFFFFF"

THIN = Side(style="thin", color="FFB0BEC5")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

H1 = Font(name="微软雅黑", size=18, bold=True, color="FFFFFFFF")
H2 = Font(name="微软雅黑", size=12, bold=True, color="FFFFFFFF")
H3 = Font(name="微软雅黑", size=11, bold=True, color=PRIMARY[2:])
BODY = Font(name="微软雅黑", size=11, color="FF1F2A44")
NOTE = Font(name="微软雅黑", size=10, italic=True, color="FF55607A")
NUM = Font(name="Consolas", size=11, color="FF1F2A44")
NUM_BOLD = Font(name="Consolas", size=11, bold=True, color=PRIMARY[2:])

FILL_PRIMARY = PatternFill("solid", fgColor=PRIMARY[2:])
FILL_ACCENT = PatternFill("solid", fgColor=ACCENT[2:])
FILL_LIGHT = PatternFill("solid", fgColor=LIGHT[2:])
FILL_ALT = PatternFill("solid", fgColor=ALT[2:])

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)


def set_widths(ws, widths: list[float]) -> None:
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_header(ws, row: int, headers: list[str]) -> None:
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=j, value=h)
        c.font = H2
        c.fill = FILL_PRIMARY
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[row].height = 28


def write_row(ws, row: int, vals: list, num_cols: list[int] | None = None, bold_total: bool = False) -> None:
    num_cols = num_cols or []
    for j, v in enumerate(vals, start=1):
        c = ws.cell(row=row, column=j, value=v)
        c.border = BORDER
        if j in num_cols:
            c.number_format = '#,##0'
            c.alignment = RIGHT
            c.font = NUM_BOLD if bold_total else NUM
        else:
            c.alignment = LEFT
            c.font = BODY
        if bold_total:
            c.fill = FILL_ACCENT
            c.font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFFFF")
        elif row % 2 == 0:
            c.fill = FILL_ALT
        else:
            c.fill = FILL_LIGHT


def title_bar(ws, text: str, span: int) -> None:
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    c = ws.cell(row=1, column=1, value=text)
    c.font = H1
    c.fill = FILL_PRIMARY
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 40


def main() -> None:
    wb = Workbook()

    # ============= Sheet 1 封面 =============
    ws = wb.active
    ws.title = "01 封面与说明"
    title_bar(ws, "元谷项目 · 胡教授团队 × 森马 联合运营合作收益测算模型 v1.0", 8)
    set_widths(ws, [4, 22, 22, 22, 22, 22, 22, 22])

    notes = [
        ("", ""),
        ("文件版本", "v1.0"),
        ("适用项目", "森马（上海）国际运营中心 元谷项目"),
        ("适用主体", "森马集团 × 危总团队 × 胡教授团队 联合运营合资公司"),
        ("货币单位", "人民币 元（除非另行注明）"),
        ("测算口径", "首年现金口径；不含税；不含合资公司行政固定成本（由合资公司主体承担）"),
        ("Sheet 索引", "02 收入结构总览 / 03 月费测算 / 04 招商佣金 / 05 活动+奖项 / 06 股权分红+三年累计 / 07 敏感性分析"),
        ("", ""),
        ("【输入资源资产化逻辑】", ""),
        ("北欧创新国际会客厅", "国际外事 / IP 引入 → 进 4# 楼，按月费 + 活动收入贡献"),
        ("福布斯系列奖项", "品牌势能 → 按挂牌奖项一次性激励"),
        ("科技开放麦", "活动 IP → 单场利润分成 + 招商导流"),
        ("AI 腾讯", "技术导流 → 招商漏斗 + 共享设计中心分润"),
        ("仲量联行爬楼大数据", "招商弹药 → 拉高佣金转化率（已投入 26,000 元）"),
        ("追觅科技基金", "返投落地 → 中下游产能锁定 + 招商佣金加成"),
        ("", ""),
        ("【收入结构】", ""),
        ("固定月费 Retainer", "覆盖团队 + 顾问 + 数据接口 → 见 Sheet 03"),
        ("专项奖项激励", "牌照 / 福布斯 / 小镇奖项 → 见 Sheet 05"),
        ("活动运营收入", "科技开放麦 + 潮玩大赛 + 北欧外事 → 见 Sheet 05"),
        ("招商佣金", "实际成交年租金的 1–2.5 个月 → 见 Sheet 04"),
        ("股权分红", "合资公司净利润按股比分配 → 见 Sheet 06"),
    ]
    for i, (k, v) in enumerate(notes, start=2):
        if not k and not v:
            continue
        a = ws.cell(row=i, column=2, value=k); a.font = H3 if k.startswith("【") else BODY; a.alignment = LEFT
        b = ws.cell(row=i, column=3, value=v); b.font = BODY; b.alignment = LEFT
        ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=8)

    # ============= Sheet 2 收入结构总览 =============
    ws = wb.create_sheet("02 收入结构总览")
    title_bar(ws, "首年收入结构总览（保守 / 基础 / 乐观 三档）", 6)
    set_widths(ws, [4, 28, 18, 18, 18, 32])

    write_header(ws, 3, ["", "收入类别", "保守场景 (元)", "基础场景 (元)", "乐观场景 (元)", "测算口径备注"])
    rows = [
        ("固定月费 Retainer (年化)",     3_600_000,  4_800_000,  6_000_000, "30 / 40 / 50 万元 × 12 月"),
        ("招商佣金 Commission",          3_000_000,  6_000_000, 12_000_000, "见 Sheet 04 招商佣金阶梯"),
        ("活动运营收入",                 1_800_000,  2_400_000,  3_600_000, "10–15 场科技开放麦 + 1 届潮玩大赛"),
        ("专项奖项 / 挂牌激励",            600_000,  1_500_000,  2_000_000, "牌照 / 福布斯 / 小镇奖项 (Sheet 05)"),
        ("股权分红 (合资公司净利)",        600_000,  1_200_000,  2_400_000, "合资公司净利润 × 30% (Sheet 06)"),
    ]
    sub_total_conservative = sum(r[1] for r in rows)
    sub_total_base = sum(r[2] for r in rows)
    sub_total_optimistic = sum(r[3] for r in rows)
    for i, (cat, c1, c2, c3, note) in enumerate(rows, start=4):
        write_row(ws, i, [i - 3, cat, c1, c2, c3, note], num_cols=[3, 4, 5])
    last = 4 + len(rows)
    write_row(ws, last, ["合计", "首年合计 (元)", sub_total_conservative, sub_total_base, sub_total_optimistic, "≈ 1,150 万 / 1,590 万 / 2,600 万"], num_cols=[3, 4, 5], bold_total=True)
    write_row(ws, last + 1, ["折合 (万元)", "首年合计 (万元)", round(sub_total_conservative / 10000), round(sub_total_base / 10000), round(sub_total_optimistic / 10000), "保守 ≈ 960 万 / 基础 ≈ 1,590 万 / 乐观 ≈ 2,600 万"], num_cols=[3, 4, 5], bold_total=True)

    # 收入占比小表
    pct_row = last + 3
    ws.cell(row=pct_row, column=2, value="基础场景下各收入占比").font = H3
    write_header(ws, pct_row + 1, ["", "收入类别", "金额 (元)", "占比", "", ""])
    for i, (cat, _c1, c2, _c3, _note) in enumerate(rows, start=pct_row + 2):
        pct = c2 / sub_total_base
        write_row(ws, i, [i - (pct_row + 1), cat, c2, f"{pct*100:.1f}%", "", ""], num_cols=[3])

    # ============= Sheet 3 月费测算 =============
    ws = wb.create_sheet("03 月费测算")
    title_bar(ws, "固定月费 Retainer · 团队成本反算（建议月费区间）", 6)
    set_widths(ws, [4, 26, 14, 14, 18, 32])

    write_header(ws, 3, ["", "成本项", "人数", "月度单价 (元)", "月度合计 (元)", "说明"])
    cost_items = [
        ("胡教授（CSO 顾问费）",   1, 60_000, "每周 2 个工作日，外加战略评审"),
        ("产业招商经理",          1, 35_000, "底薪 25K + 招商绩效 10K"),
        ("国际合作 & 活动策划",    1, 30_000, "底薪 22K + 活动绩效 8K"),
        ("基金投后 & 政府关系",    1, 30_000, "底薪 22K + 牌照绩效 8K"),
        ("行政 / 财务 (共享)",     1, 12_000, "合资公司共享岗位的 50% 摊销"),
        ("仲量联行爬楼数据接口",   1, 5_000,  "在 26,000 元已购基础上的运维费摊销"),
        ("差旅 / 接待 / 物料",     1, 18_000, "北欧外事接待、爬楼差旅、物料制作"),
        ("管理与协调费",           1, 10_000, "合资公司沟通、月度汇报、外部协调"),
    ]
    total_monthly = 0
    for i, (name, n, price, note) in enumerate(cost_items, start=4):
        sub = n * price
        total_monthly += sub
        write_row(ws, i, [i - 3, name, n, price, sub, note], num_cols=[3, 4, 5])
    last = 4 + len(cost_items)
    write_row(ws, last, ["合计", "团队月度成本", "", "", total_monthly, "胡教授团队月度刚性支出"], num_cols=[5], bold_total=True)

    margin_rows = [
        ("月费保底 (建议)", 400_000, "推荐方案：覆盖成本 + 约 40% 安全垫"),
        ("月费区间下限",   300_000, "保守方案：仅覆盖核心 5 人 + 接口"),
        ("月费区间上限",   500_000, "乐观方案：扩配 + 数据 + 国际接待"),
    ]
    ws.cell(row=last + 2, column=2, value="月费谈判区间").font = H3
    write_header(ws, last + 3, ["", "档位", "金额 (元/月)", "年化 (元)", "毛利率 (相对成本)", "说明"])
    for i, (name, m, note) in enumerate(margin_rows, start=last + 4):
        annual = m * 12
        margin = (m - total_monthly) / m
        write_row(ws, i, [i - (last + 3), name, m, annual, f"{margin*100:.1f}%", note], num_cols=[3, 4])

    # ============= Sheet 4 招商佣金阶梯 =============
    ws = wb.create_sheet("04 招商佣金阶梯")
    title_bar(ws, "招商佣金阶梯测算（按面积分档；以年租金 X 个月计提）", 7)
    set_widths(ws, [4, 24, 14, 14, 14, 16, 28])

    write_header(ws, 3, ["", "面积档位", "日租金 (元/㎡/天)", "年租金 (元/㎡)", "佣金月数", "佣金 (元/㎡)", "说明"])
    daily = [4.5, 5.0, 5.5]
    tiers = [
        ("≤ 2,000㎡ 小型",      daily[0], 1.0, "小型潮玩 / 服务机构"),
        ("2,001–5,000㎡ 中型",   daily[1], 1.5, "中型潮玩运营企业"),
        ("> 5,000㎡ 头部",       daily[2], 2.5, "头部央企 / 行业协会 / 中型以上品牌"),
    ]
    for i, (tier, d, m, note) in enumerate(tiers, start=4):
        annual = round(d * 365)
        commission = round(annual * m / 12)
        write_row(ws, i, [i - 3, tier, d, annual, m, commission, note], num_cols=[4, 6])

    # 三档情境总佣金
    write_header(ws, 9, ["", "场景", "成交面积 (㎡)", "平均佣金 (元/㎡)", "成交家数", "佣金合计 (元)", "说明"])
    scenarios = [
        ("保守", 6_000,  500, 12, "首年成交不足；以中小型为主"),
        ("基础", 12_000, 500, 22, "稳态推进；中小+中型 + 1 家头部"),
        ("乐观", 20_000, 600, 35, "返投基金 + 牌照拉动；头部签约 3 家"),
    ]
    for i, (s, area, avg, n, note) in enumerate(scenarios, start=10):
        total = area * avg
        write_row(ws, i, [i - 9, s, area, avg, n, total, note], num_cols=[3, 4, 5, 6])

    ws.cell(row=14, column=2, value="说明").font = H3
    ws.cell(row=15, column=2, value="① 仲量联行爬楼大数据可使佣金转化率提升 30%；②追觅基金返投 1:1.5，锁定企业额外计 0.5 个月佣金加成；③ 招商主导权由合资公司独家持有 5 年。").font = NOTE
    ws.merge_cells(start_row=15, start_column=2, end_row=15, end_column=7)
    ws.row_dimensions[15].height = 38

    # ============= Sheet 5 活动 + 奖项 =============
    ws = wb.create_sheet("05 活动+奖项")
    title_bar(ws, "活动运营收入 + 专项奖项激励 测算", 7)
    set_widths(ws, [4, 28, 12, 14, 14, 16, 28])

    write_header(ws, 3, ["", "活动 / 奖项", "频次/年", "单场收入 (元)", "成本 (元)", "净收入 (元)", "说明"])
    activities = [
        ("科技开放麦 (基础场)",     12, 80_000,  30_000, "赞助 + 票务 + 园区分摊"),
        ("科技开放麦 (大场)",        2, 250_000, 80_000, "森马联合发布 / 国际嘉宾"),
        ("北欧创新国际会客厅外事接待", 6, 120_000, 50_000, "外事接待 + 路演"),
        ("全国潮玩设计大赛 (年度)",   1, 800_000, 300_000, "省级补贴 + 头部赞助"),
        ("福布斯榜单 元谷专场发布",   1, 600_000, 200_000, "联合森马 + 福布斯方"),
        ("AI 共享设计中心 工作坊",   12, 30_000,  10_000, "AI 腾讯背书 + 中型企业付费"),
    ]
    total_net = 0
    for i, (name, f, rev, cost, note) in enumerate(activities, start=4):
        net = (rev - cost) * f
        total_net += net
        write_row(ws, i, [i - 3, name, f, rev, cost, net, note], num_cols=[4, 5, 6])
    write_row(ws, 4 + len(activities), ["合计", "活动年度净收入", "", "", "", total_net, "≈ 240–360 万元，与 Sheet 02 对应"], num_cols=[6], bold_total=True)

    # 奖项激励
    base = 4 + len(activities) + 2
    ws.cell(row=base, column=2, value="专项奖项 / 挂牌激励 (一次性)").font = H3
    write_header(ws, base + 1, ["", "奖项 / 挂牌", "数量", "单项激励 (元)", "合计 (元)", "", "说明"])
    awards = [
        ("AI 潮玩产业基地 牌照挂牌",      1, 300_000, "中国动漫集团"),
        ("潮玩次元商业专委会 牌照挂牌",   1, 300_000, "中国百货商业协会"),
        ("上海科技时尚特色小镇 政府奖项",  1, 400_000, "市/区级"),
        ("福布斯榜单上榜 / 评选挂牌",     2, 200_000, "U30 / 创新榜 / 园区榜"),
        ("国家级科技 / 文创奖项",         1, 300_000, "如国家文创基金 / 工信部专项"),
    ]
    award_total = 0
    for i, (n, qty, price, note) in enumerate(awards, start=base + 2):
        total = qty * price
        award_total += total
        write_row(ws, i, [i - (base + 1), n, qty, price, total, "", note], num_cols=[4, 5])
    write_row(ws, base + 2 + len(awards), ["合计", "首年挂牌奖励合计", "", "", award_total, "", "≈ 150–200 万元"], num_cols=[5], bold_total=True)

    # ============= Sheet 6 股权分红 + 三年累计 =============
    ws = wb.create_sheet("06 股权分红+三年累计")
    title_bar(ws, "股权分红预测 + 三年累计现金流（基础场景）", 7)
    set_widths(ws, [4, 28, 16, 16, 16, 16, 28])

    write_header(ws, 3, ["", "项目", "T+1 年 (元)", "T+2 年 (元)", "T+3 年 (元)", "三年合计 (元)", "假设"])

    rows_3y = [
        ("固定月费 Retainer", 4_800_000, 5_400_000, 6_000_000, "月费 40 → 45 → 50 万"),
        ("招商佣金",          6_000_000, 8_400_000, 10_800_000, "成交面积 1.2 / 1.6 / 2.0 万㎡"),
        ("活动运营净收入",     2_400_000, 3_000_000,  3_600_000, "活动+1场/年 + 大赛升级"),
        ("奖项 / 挂牌激励",    1_500_000, 1_000_000,  1_500_000, "牌照前置，第二年保持"),
        ("合资公司分红 (30%)", 1_200_000, 2_400_000,  3_900_000, "合资公司直营业态进入正现金流"),
    ]
    sums = [0, 0, 0]
    for i, (n, a, b, c, note) in enumerate(rows_3y, start=4):
        s = a + b + c
        sums[0] += a; sums[1] += b; sums[2] += c
        write_row(ws, i, [i - 3, n, a, b, c, s, note], num_cols=[3, 4, 5, 6])
    last = 4 + len(rows_3y)
    write_row(ws, last, ["合计", "胡教授团队年度收入 (元)", sums[0], sums[1], sums[2], sum(sums), "三年合计现金口径"], num_cols=[3, 4, 5, 6], bold_total=True)
    write_row(ws, last + 1, ["折合", "胡教授团队年度收入 (万元)", round(sums[0] / 10000), round(sums[1] / 10000), round(sums[2] / 10000), round(sum(sums) / 10000), "三年累计 ≈ 6,000+ 万元"], num_cols=[3, 4, 5, 6], bold_total=True)

    # 合资公司利润反推 (说明分红 30% 来源)
    base = last + 3
    ws.cell(row=base, column=2, value="合资公司净利润反推 (基础场景)").font = H3
    write_header(ws, base + 1, ["", "项目", "T+1 (元)", "T+2 (元)", "T+3 (元)", "", "说明"])
    jv_rows = [
        ("合资公司收入",         12_000_000, 18_000_000, 26_000_000, "招商分润 + 直营业态毛利"),
        ("合资公司经营成本",      8_000_000, 10_000_000, 12_000_000, "团队 + 运营 + 租赁分摊"),
        ("合资公司净利润",        4_000_000,  8_000_000, 13_000_000, "基础场景"),
        ("胡教授团队 30% 分红",   1_200_000,  2_400_000,  3_900_000, "= 上述净利润 × 30%"),
    ]
    for i, (n, a, b, c, note) in enumerate(jv_rows, start=base + 2):
        write_row(ws, i, [i - (base + 1), n, a, b, c, "", note], num_cols=[3, 4, 5])

    # ============= Sheet 7 敏感性分析 =============
    ws = wb.create_sheet("07 敏感性分析")
    title_bar(ws, "敏感性分析：月费 × 招商成交面积 → 首年总收入", 9)
    set_widths(ws, [4, 26, 14, 14, 14, 14, 14, 14, 14])

    monthly_options = [200_000, 300_000, 400_000, 500_000, 600_000]
    area_options = [4_000, 8_000, 12_000, 16_000, 20_000, 24_000]

    write_header(ws, 3, ["", "月费 \\ 成交面积 (㎡)"] + [f"{a:,}" for a in area_options])
    for i, m in enumerate(monthly_options, start=4):
        row = [i - 3, f"月费 {m:,} 元"]
        for a in area_options:
            avg_commission = 500
            commission = a * avg_commission
            activity_award = 3_900_000
            dividend = 1_200_000
            total = m * 12 + commission + activity_award + dividend
            row.append(total)
        write_row(ws, i, row, num_cols=list(range(3, 3 + len(area_options))))

    note = "说明：表内单元格 = 月费 × 12 + 成交面积 × 500 元/㎡（佣金平均） + 活动 + 奖项 (390 万) + 分红 (120 万)。仅作敏感性参考，不替代逐项测算。"
    ws.cell(row=4 + len(monthly_options) + 1, column=2, value=note).font = NOTE
    ws.merge_cells(start_row=4 + len(monthly_options) + 1, start_column=2, end_row=4 + len(monthly_options) + 1, end_column=2 + len(area_options))
    ws.row_dimensions[4 + len(monthly_options) + 1].height = 36

    wb.save(OUT)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
