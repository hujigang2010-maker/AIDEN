#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""10 月 31 日王德峰教授讲座：综合多方意见后的最终定价与 4 人落地执行方案。"""

import math
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

OUT = Path(__file__).resolve().parents[1] / "output"
DOCX_PATH = OUT / "10月31日讲座_最终定价与落地执行方案.docx"
CHECK_PATH = OUT / "10月31日讲座_测算核对.txt"

NAVY = RGBColor(0x1F, 0x3A, 0x5F)
GREY = RGBColor(0x55, 0x55, 0x55)
FONT = "微软雅黑"

# 实收率：扣除票务平台手续费、开票税费与少量退款后的到手比例
NET_RATE = 0.94
CAPACITY = 200
FEES_WAN = [5, 8, 10, 12, 15, 18]

# 推荐票档结构：(名称, 价格, 占比, 说明)
MIX_RECOMMENDED = [
    ("早鸟席", 499, 0.30, "10 月 20 日 24:00 前，限 60 张"),
    ("标准席", 699, 0.50, "对外主价格"),
    ("企业团体", 599, 0.10, "同一单位 5 张起，按人计价"),
    ("前排席", 999, 0.10, "限 20 张，前三排＋纸质讲义；签名书确认后再加"),
]
MIX_LOW = [("早鸟", 399, 0.30, ""), ("标准", 499, 0.60, ""), ("VIP", 999, 0.10, "")]
MIX_HIGH = [("早鸟", 699, 0.30, ""), ("标准", 899, 0.50, ""), ("双人", 799, 0.10, ""), ("VIP", 1599, 0.10, "")]


def avg_price(mix):
    return sum(p * w for _, p, w, _ in mix)


def breakeven(fee_wan, price):
    return math.ceil(fee_wan * 10000 / (price * NET_RATE))


def set_font(run, size=10.5, bold=False, color=None):
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def para(doc, text, size=10.5, bold=False, color=None, after=4, align=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.25
    if align is not None:
        p.alignment = align
    set_font(p.add_run(text), size, bold, color)
    return p


def h1(doc, text):
    return para(doc, text, size=14, bold=True, color=NAVY, after=6)


def h2(doc, text):
    return para(doc, text, size=11.5, bold=True, color=NAVY, after=3)


def bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(2)
    set_font(p.add_run(text), 10.5)
    return p


def table(doc, header, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]
        c.text = ""
        set_font(c.paragraphs[0].add_run(h), 9.5, True, NAVY)
    for r in rows:
        cells = t.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = ""
            set_font(cells[i].paragraphs[0].add_run(str(v)), 9.5)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def yuan(x):
    return f"{round(x):,} 元"


def build():
    rec_avg = avg_price(MIX_RECOMMENDED)
    low_avg = avg_price(MIX_LOW)
    high_avg = avg_price(MIX_HIGH)
    rec_net = rec_avg * NET_RATE

    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Cm(2)
    sec.left_margin = sec.right_margin = Cm(2.2)

    para(doc, "王德峰教授 10 月 31 日《传习录》与阳明心学讲座", size=16, bold=True, color=NAVY,
         align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
    para(doc, "最终定价与 4 人落地执行方案（综合 Muse / Grok / Gemini / DeepSeek / Cursor / 汇总意见）",
         size=10.5, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, after=10)

    h1(doc, "一、结论")
    para(doc, "标准席 699 元/人，早鸟 499 元/人（10 月 20 日前、限 60 张）。"
              "另设企业团体 599 元/人（5 张起）和限量 20 张前排席 999 元。海报主价格只写“699 元，早鸟 499 元”。",
         bold=True)
    para(doc, f"按这个结构，平均票面约 {round(rec_avg)} 元，到手约 {round(rec_net)} 元/人。"
              f"讲师费 10 万元时，需要 {breakeven(10, rec_avg)} 人付费才能保本，"
              f"约占 {CAPACITY} 座的 {breakeven(10, rec_avg) / CAPACITY:.0%}。")
    para(doc, "这个价格成立有两个前提：讲师费不超过约 10 万元，且讲师费可以分期、老师取消全额退。"
              "10–12 万元要靠企业团体票托底；超过 12 万元、或对方坚持单场 2,500 元，"
              "就不应靠公开门票来扛，见第三部分的决策表。")

    h1(doc, "二、为什么是 699，而不是 499、899 或 1,680")
    h2(doc, "1. 成本结构变了，499 已经不够")
    para(doc, "此前对外沟通的 499 元，前提是讲师费按票务分成结算，主办方不先付固定现金。"
              "现在的前提是：场地、物料等其他费用为零，唯一的现金支出是一笔固定讲师费。"
              "固定成本只有一项，保本人数就完全由“讲师费 ÷ 每张票到手金额”决定。")
    rows = []
    for p in [399, 499, 599, 699, 899]:
        rows.append([f"{p} 元"] + [f"{breakeven(f, p)} 人" for f in FEES_WAN])
    table(doc, ["全场单一票价"] + [f"讲师费 {f} 万" for f in FEES_WAN], rows)
    para(doc, f"说明：到手按票面 {NET_RATE:.0%} 计，已扣票务平台手续费、开票税费和少量退款。"
              f"若场地约 {CAPACITY} 座、实际能卖七到八成（140–160 人），"
              f"讲师费 10 万元时 499 元需要 {breakeven(10, 499)} 人，超过整个场地；"
              f"699 元需要 {breakeven(10, 699)} 人，刚好落在可卖范围内。", size=9.5, color=GREY)

    h2(doc, "2. 三种票档结构对比（以 200 座为例）")
    rows = []
    for name, mix, avg in [
        ("399 / 499（Cursor、DeepSeek 路线）", MIX_LOW, low_avg),
        ("499 / 699（本方案）", MIX_RECOMMENDED, rec_avg),
        ("699 / 899（Muse 路线）", MIX_HIGH, high_avg),
    ]:
        net = avg * NET_RATE
        rows.append([name, f"{round(avg)} 元", f"{breakeven(10, avg)} 人",
                     yuan(150 * net), yuan(180 * net)])
    table(doc, ["票档结构", "平均票面", "讲师费 10 万保本", "卖 150 人到手", "卖 180 人到手"], rows,
          widths=[5.6, 2.2, 3.0, 3.0, 3.0])
    para(doc, "699 / 899 账面更好看，但前提是卖得动。上海普通听众对“纯听 3 小时讲座”的心理价位多在 500–800 元，"
              "899 已经超出这个区间。21 天的销售窗口不允许先定高价、卖不动再降价：降价会伤害已付款的人。"
              "所以把 699 作为主价格，用 499 早鸟带动开售速度，再用团体票和前排席抬高平均票面。")

    h2(doc, "3. 上海当前的消费环境")
    bullet(doc, "国家统计局 2026 年上半年数据：城镇居民人均可支配收入中位数 26,389 元，折合每月约 4,398 元。"
                "699 元约占月收入中位数的 15.9%，499 元约占 11.3%。上海收入高于全国，但消费者对非刚需项目会比价。")
    bullet(doc, "真正的主力买家不是大众粉丝，而是企业经营者、管理者和专业人士。"
                "这批人里有相当一部分可以按企业培训报销，对 699 元不敏感，但对“值不值”很敏感。")
    bullet(doc, "王德峰教授网上公开视频很多，普通粉丝愿意为“现场听”付费，但上限不高。"
                "这也是不把 1,680–1,980 元（Grok 建议）作为普通席价格的原因。")

    h2(doc, "4. 对各方意见的取舍")
    table(doc, ["来源", "建议", "采纳", "不采纳"], [
        ["Muse", "699 早鸟 / 899 标准 / 1,599 VIP", "成本倒推法；先理顺票种", "899 主价格偏高"],
        ["Grok", "1,680–1,980 标准", "企业团体折扣", "3 小时讲座撑不起这个价"],
        ["Gemini", "580–780 普通 / 1,280–1,680 VIP", "按座位区分档；复旦名义合规提醒", "“学费改叫会务费”规避监管"],
        ["DeepSeek", "480–680 主力；先测后定", "单场讲座定位；授权优先", "讲师费 10–30 万的估计没有依据"],
        ["Cursor", "统一 499", "海报只出一个主价格", "固定讲师费下保不了本"],
    ], widths=[2.0, 4.6, 5.0, 5.0])

    h1(doc, "三、按讲师费决定票价（决策表）")
    para(doc, "讲师费是唯一的变量。先拿到书面数字，再按下表选方案。")
    table(doc, ["讲师费（含税现金）", "票价方案", "保本人数", "判断"], [
        ["按票务分成，或 ≤ 6 万", "统一 499（早鸟 399）", f"约 {breakeven(6, low_avg)} 人以内", "维持此前对外口径，优先扩大影响"],
        ["6–10 万", "本方案：699 / 早鸟 499", f"{breakeven(6, rec_avg)}–{breakeven(10, rec_avg)} 人", "推荐，200 座卖七到八成即可"],
        ["10–12 万", "本方案＋团体票托底", f"{breakeven(10, rec_avg)}–{breakeven(12, rec_avg)} 人",
         "要卖到八成以上，开售前先锁定 30 张以上企业团体票"],
        ["12–18 万", "本方案＋赞助", f"{breakeven(12, rec_avg)}–{breakeven(18, rec_avg)} 人", "超出 200 座，赞助签约到账后再开售"],
        ["> 18 万，或坚持单场 2,500 元", "不做公开售票", "—", "改为分成，或只做宣传支持"],
    ], widths=[4.0, 4.2, 3.0, 5.4])
    para(doc, "讲师费付款建议：签约付 30% 定金，活动结束后 3 个工作日内付清；合同写明老师因故取消，已付款项 100% 退还。"
              "不要在开售前一次性付清。", size=10)

    h1(doc, "四、开售前必须先改掉的四处问题")
    bullet(doc, "删除“半天专题席位”。整场只有一个上午，不存在“半天票”。席位只按座位区分：标准席、前排席。")
    bullet(doc, "海报上的“（上）”“（下）”改成“第一讲”“第二讲”，明确是同一上午的两段，避免被误解为两天课。")
    bullet(doc, "复旦大学住房政策研究中心的名义：如果没有拿到校内书面同意，就列为“学术支持单位”，不作为收费主体。"
                "收款与开票由首乾书院或杨浦区科技企业联合会其中一家负责，提前定好。")
    bullet(doc, "签名书、合影、问答在书面确认前不写进任何票档权益。前排席先只承诺“前三排座位＋纸质讲义”。")

    h1(doc, "五、4 人分工")
    table(doc, ["角色", "负责事项", "关键交付"], [
        ["A 总负责", "讲师费合同与付款、首乾书院对接、复旦与企业联合会授权、赞助、最终拍板",
         "10 月 12 日前拿到讲师费、授权、收款主体三份书面确认"],
        ["B 内容传播", "海报定稿、公众号推文、朋友圈素材、倒计时海报、社群内容",
         "10 月 13 日开售推文；每 2–3 天一条更新"],
        ["C 票务客服", "报名小程序（小鹅通或活动行）、收款、开票、退款、群答疑、名单与座位",
         "每天 20:00 报销售日报；10 月 29 日锁定名单"],
        ["D 渠道现场", "企业联合会会员企业逐家邀约团体票；场地、设备、签到、茶歇、动线",
         "团体票目标 20 张以上；10 月 30 日完成彩排"],
    ], widths=[2.4, 8.0, 6.2])
    para(doc, "4 个人足够，但必须把讲师、授权、收款这三件事集中到一个人手里，其他三人只管卖票和交付。", size=10)

    h1(doc, "六、21 天执行时间表（10 月 10 日—10 月 31 日）")
    be = breakeven(10, rec_avg)
    table(doc, ["时间", "要做的事", "检查点"], [
        ["10/10—10/12", "锁定讲师费与付款方式、场地容量、授权与收款主体；按决策表定票价；改海报",
         "三份书面确认齐全，否则不开售"],
        ["10/13", "开售：公众号首发、朋友圈、首乾书院群、企业联合会会员群同步", "首日售出 ≥ 30 张"],
        ["10/13—10/20", "早鸟期。D 逐家打电话邀约会员企业；B 每两天发一次进度和讲座看点",
         f"10/17 达到保本人数的 30%（约 {math.ceil(be * 0.3)} 人）"],
        ["10/20 24:00", "早鸟截止，自动切换 699", f"累计达到保本人数的 60%（约 {math.ceil(be * 0.6)} 人）"],
        ["10/22", "是否继续的决策点",
         f"不足保本人数的 50%（约 {math.ceil(be * 0.5)} 人）：与讲师方改为分成或降低讲师费；谈不成就取消活动，全额退款"],
        ["10/21—10/28", "标准价期。推企业团体票、老学员带新人；发参会须知", f"10/28 达到 {be} 人以上"],
        ["10/29", "停售、锁定名单与座位，发入场二维码和交通指引", "名单与座位表定稿"],
        ["10/30", "场地彩排：音响、投影、话筒、签到、茶歇点位、引导牌", "设备全部试过一次"],
        ["10/31", "07:30 到场；08:15 签到；09:00 开场；12:00 结束", "现场零差错"],
        ["11/1—11/7", "讲师尾款、开票、退款收尾、复盘；社群继续运营", "复盘：到场率、渠道来源、满意度"],
    ], widths=[2.4, 8.4, 5.8])

    h2(doc, "当天议程（建议）")
    table(doc, ["时间", "内容"], [
        ["09:00—09:05", "主持开场，介绍主办单位"],
        ["09:05—10:20", "第一讲：走进王阳明——心即理、致良知"],
        ["10:20—10:40", "茶歇交流（现场才有的交流，是和看网上视频的区别）"],
        ["10:40—11:40", "第二讲：知行合一——在不确定的时代践行心学"],
        ["11:40—12:00", "现场问答（问题提前在群里征集，主持人筛选）"],
    ], widths=[3.0, 13.6])

    h1(doc, "七、退费规则（写在报名页）")
    bullet(doc, "10 月 24 日 24:00 前：全额退款。")
    bullet(doc, "10 月 25 日—10 月 29 日：退 80%，或免费转让给他人。")
    bullet(doc, "10 月 30 日起：不退款，可转让。")
    bullet(doc, "因讲师或主办方原因取消或改期：全额退款。")

    h1(doc, "八、海报票务区文字")
    para(doc, "标准席　699 元/人　｜　早鸟 499 元（10 月 20 日前，限 60 席）", bold=True, after=2)
    para(doc, "企业团体 5 人起 599 元/人　｜　前排席 999 元（限 20 席）", after=2)
    para(doc, "2026 年 10 月 31 日（星期六）上午 09:00—12:00　｜　扫码报名", after=8)

    h1(doc, "九、说明")
    para(doc, "讲师费、场地容量与上座率都是待确认的假设，不是已核实数据。"
              "本方案中的保本人数按“唯一现金成本是讲师费”计算；"
              "茶歇、讲义、签到物料等若需少量支出，建议预留讲师费的 5% 作为备用金。",
         size=9.5, color=GREY)

    OUT.mkdir(parents=True, exist_ok=True)
    doc.save(DOCX_PATH)

    lines = [f"实收率 {NET_RATE}；场地 {CAPACITY} 座", ""]
    lines.append("单一票价保本人数（讲师费 " + "/".join(str(f) for f in FEES_WAN) + " 万）")
    for p in [399, 499, 599, 699, 899]:
        lines.append(f"  {p} 元: " + ", ".join(str(breakeven(f, p)) for f in FEES_WAN))
    lines.append("")
    for name, mix in [("399/499", MIX_LOW), ("499/699 推荐", MIX_RECOMMENDED), ("699/899", MIX_HIGH)]:
        a = avg_price(mix)
        lines.append(f"{name}: 平均票面 {a:.1f}，到手 {a * NET_RATE:.1f}；保本 "
                     + ", ".join(str(breakeven(f, a)) for f in FEES_WAN)
                     + f"；150 人到手 {150 * a * NET_RATE:,.0f}，180 人到手 {180 * a * NET_RATE:,.0f}")
    CHECK_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    build()
    print(DOCX_PATH)
    print(CHECK_PATH.read_text(encoding="utf-8"))
