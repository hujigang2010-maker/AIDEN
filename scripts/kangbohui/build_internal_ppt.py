"""内部汇报：承办决策与执行策划。"""

from __future__ import annotations

import os

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

from theme import (
    CARD_GOLD,
    CARD_TEAL,
    CREAM,
    GOLD,
    INK,
    LINE,
    MIST,
    MUTED,
    PAPER,
    ROSE,
    ROSE_SOFT,
    SLATE,
    SW,
    TEAL,
    TEAL_DEEP,
    WHITE,
    add_table,
    add_text,
    blank,
    content_header,
    notes,
    rect,
    rrect,
)

TOTAL = 18
OUT = os.path.normpath(
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "..",
        "deliverables",
        "2026康博会专题分论坛",
        "01-内部-承办决策与执行策划.pptx",
    )
)


def new_prs():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = "2026康博会专题分论坛承办决策与执行策划"
    prs.core_properties.subject = "内部执行稿 V3.0"
    prs.core_properties.category = "内部"
    prs.core_properties.author = "专题分论坛筹备组"
    return prs


def cover(prs):
    s = blank(prs)
    rect(s, 0, 0, SW, 7.5, INK)
    rect(s, 0, 0, 0.14, 7.5, GOLD)
    add_text(s, "内部执行稿  ·  V3.0  ·  2026年11月11日", 0.62, 0.38, 8, 0.32, size=13, color=GOLD)
    add_text(s, "不对外发送", 10.3, 0.38, 2.5, 0.32, size=13, bold=True, color=GOLD, align=PP_ALIGN.RIGHT)
    add_text(s, "2026上海国际智慧康养产业博览会", 0.62, 1.55, 12, 0.38, size=18, color=CREAM)
    add_text(s, "专题分论坛", 0.58, 2.05, 12, 0.78, size=48, bold=True, color=WHITE)
    add_text(s, "承办决策与执行策划", 0.62, 2.88, 12, 0.58, size=32, bold=True, color=GOLD)
    rect(s, 0.62, 3.62, 1.4, 0.045, GOLD)
    add_text(
        s,
        "“十五五”银发科技与具身智能康养创新论坛",
        0.62,
        3.85,
        12,
        0.42,
        size=20,
        color=WHITE,
    )
    add_text(
        s,
        "AI养老场景应用  ·  主动健康  ·  科技成果转化与产业资本对接",
        0.62,
        4.32,
        12,
        0.34,
        size=15,
        color=CREAM,
    )
    rect(s, 0.62, 5.15, 12.1, 0.015, TEAL)
    metas = [
        ("11月28日", "星期六  13:00—16:30"),
        ("上海展览中心", "中央大厅会议室C，待确认"),
        ("100—120人", "闭门对接，不做重礼仪峰会"),
    ]
    for i, (a, b) in enumerate(metas):
        x = 0.62 + i * 4.05
        add_text(s, a, x, 5.4, 3.8, 0.36, size=18, bold=True, color=WHITE)
        add_text(s, b, x, 5.8, 3.8, 0.3, size=13, color=CREAM)
    add_text(
        s,
        "拟提交大会组委会审议    ·    结论：有条件承接，重新策划",
        0.62,
        6.7,
        12,
        0.34,
        size=14,
        color=GOLD,
    )
    notes(s, "本套幻灯片仅供筹备组决策和与组委会谈判。赞助商只接收对外赞助方案。")


def slide_judgment(prs):
    s = blank(prs)
    content_header(s, "01    总体结论", "有条件承接，不按原方案直接办", 2, TOTAL)
    rrect(s, 0.42, 1.18, 12.5, 0.78, CARD_TEAL, radius=0.08)
    add_text(
        s,
        "值得争取的是康博会平台上的长期合作网络，而不是一场120人的常规会议。距离11月28日只剩17天，改为小规模、有买方、有项目成果的产业对接。",
        0.62,
        1.3,
        12.1,
        0.56,
        size=15,
        color=INK,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    cards = [
        (TEAL, "值得争取", "大会场地和协会入口可以降低独立办会成本。高校、科创企业、投资机构与养老场景之间，存在真实的供需合作。"),
        (GOLD, "必须重构", "嘉宾、演讲和签约主体仍待确认。大会已有AI融合峰会和投资风向峰会，本场不能再重复同一题目。"),
        (ROSE, "时间已紧", "原招商函要求11月15日前反馈。现在不适合重礼仪、重嘉宾数量的传统峰会，也不适合追逐不确定的高额现金赞助。"),
    ]
    for i, (color, title, body) in enumerate(cards):
        x = 0.42 + i * 4.18
        rrect(s, x, 2.14, 4.02, 2.55, PAPER, line=LINE, radius=0.08)
        rect(s, x, 2.14, 0.08, 2.55, color)
        add_text(s, title, x + 0.28, 2.32, 3.55, 0.4, size=18, bold=True, color=color)
        add_text(s, body, x + 0.28, 2.82, 3.52, 1.65, size=14, color=SLATE, spacing=1.15)
    pills = [
        ("大会平台", "降低场地与组织成本"),
        ("产业入口", "机构、医院、金融与科技供需"),
        ("已有资源", "高校、科创、投资与协会"),
        ("会后延伸", "转化、融资顾问与季度对接"),
    ]
    for i, (a, b) in enumerate(pills):
        x = 0.42 + i * 3.14
        rrect(s, x, 4.86, 3.02, 0.95, MIST, radius=0.1)
        add_text(s, a, x + 0.16, 4.96, 2.7, 0.32, size=14, bold=True, color=INK)
        add_text(s, b, x + 0.16, 5.3, 2.7, 0.36, size=12, color=MUTED)
    rrect(s, 0.42, 5.98, 12.5, 0.98, INK, radius=0.08)
    add_text(s, "决策原则", 0.64, 6.1, 2, 0.28, size=12, bold=True, color=GOLD)
    add_text(
        s,
        "有策划权、有收益机制、有成本边界，就接。只有执行责任、没有授权和结算保障，就不以独立总承办身份接手。",
        0.64,
        6.4,
        12.05,
        0.44,
        size=15,
        color=WHITE,
    )
    notes(s, "结论对应建议书：有条件接，但必须重构。个人更适合做总策划、资源整合和商业合作，而不是只执行别人已定的方案。")


def slide_rights(prs):
    s = blank(prs)
    content_header(s, "02    承接前提", "先签承办协议，再对外招商", 3, TOTAL)
    cards = [
        ("策划权与执行权", "主题、议程、嘉宾、合作单位、主持人和赞助方案由项目团队主导，组委会做必要审批。现场执行、嘉宾联络和对接服务可以统筹。"),
        ("招商权与收益权", "可销售的冠名、展示、发言和采购对接权益，必须与大会已售权益核对。回款归属、结算和利润分配写成文字。"),
        ("成本边界", "场地、基础音视频、现场人员和官方宣传由大会落实。不承担大会整体招商缺口，不接受个人无上限垫资。"),
    ]
    for i, (title, body) in enumerate(cards):
        x = 0.42 + i * 4.18
        rrect(s, x, 1.22, 4.02, 3.15, PAPER, line=LINE, radius=0.08)
        rrect(s, x + 0.22, 1.42, 0.62, 0.42, TEAL, radius=0.12)
        add_text(s, f"0{i+1}", x + 0.22, 1.46, 0.62, 0.34, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(s, title, x + 0.22, 2.02, 3.58, 0.7, size=18, bold=True, color=INK)
        add_text(s, body, x + 0.22, 2.78, 3.58, 1.4, size=14, color=SLATE, spacing=1.12)
    rrect(s, 0.42, 4.54, 12.5, 2.42, CARD_GOLD, radius=0.08)
    add_text(s, "内部谈判口径", 0.66, 4.7, 6, 0.32, size=14, bold=True, color=GOLD)
    add_text(
        s,
        "优先谈“固定策划执行费 + 招商绩效提成”。建议固定费1.5万—3万元，另加由本团队引入且完成回款部分的20%—30%作为招商服务费。",
        0.66,
        5.08,
        12.0,
        0.7,
        size=15,
        color=INK,
    )
    add_text(
        s,
        "大会若不支付固定费，则论坛独立核算、费用可审计、赞助收入透明，执行团队争取可分配净收益的50%—60%。两种方式只选一种，避免重复计酬。比例是谈判起点，不是行业统一标准。",
        0.66,
        5.78,
        12.0,
        0.95,
        size=14,
        color=SLATE,
    )
    notes(s, "这页不进入对外赞助方案，也不宜原样提交给尚未建立信任的组委会。对组委会只强调：授权、独立核算、书面结算、不个人垫资。")


def slide_gates(prs):
    s = blank(prs)
    content_header(s, "03    时间", "14日决定能否发布，20日决定投入多少", 4, TOTAL)
    steps = [
        ("11/11—13", "定条件", "授权、结算、场地"),
        ("11/14", "闸门一", "嘉宾渠道与技术条件"),
        ("11/16—20", "锁资源", "合同、嘉宾、项目材料"),
        ("11/20", "闸门二", "3万元或等额资源"),
        ("11/28", "举办", "对接记录，不做空签约"),
        ("12/05前", "收口", "结算、核销、合作台账"),
    ]
    for i, (when, name, detail) in enumerate(steps):
        x = 0.42 + i * 2.12
        rrect(s, x, 1.28, 2.0, 1.72, INK if i in (1, 3) else MIST, radius=0.1)
        color = GOLD if i in (1, 3) else TEAL
        title_c = WHITE if i in (1, 3) else INK
        body_c = CREAM if i in (1, 3) else MUTED
        add_text(s, when, x + 0.1, 1.4, 1.8, 0.28, size=12, bold=True, color=color, align=PP_ALIGN.CENTER)
        add_text(s, name, x + 0.1, 1.74, 1.8, 0.4, size=16, bold=True, color=title_c, align=PP_ALIGN.CENTER)
        add_text(s, detail, x + 0.12, 2.22, 1.76, 0.6, size=12, color=body_c, align=PP_ALIGN.CENTER)
    blocks = [
        (CARD_TEAL, "11月14日以前必须明确", "主办授权、主旨嘉宾邀约渠道、傅利叶合作意向、场地技术条件。有一项仍不清楚，就不大规模对外发布，也不使用任何未授权的单位名称和标识。"),
        (ROSE_SOFT, "11月20日以前必须落袋", "至少3万元实际到账，或取得正式合同保障的等额资源。达不到，就改为低成本研讨，取消茶歇、非必要物料和演示运输。必要费用也没有时，暂停支出并重谈规模。"),
    ]
    for i, (fill, title, body) in enumerate(blocks):
        x = 0.42 + i * 6.34
        rrect(s, x, 3.24, 6.16, 2.15, fill, radius=0.08)
        add_text(s, title, x + 0.24, 3.4, 5.7, 0.38, size=16, bold=True, color=INK)
        add_text(s, body, x + 0.24, 3.88, 5.7, 1.3, size=14, color=SLATE, spacing=1.12)
    add_text(
        s,
        "若大会交给的是整组分论坛而不只是11月28日下午这一场：先争取整体策划和招商统筹权，今年实际重点操盘不超过两到三个精品场次。",
        0.42,
        5.58,
        12.5,
        0.7,
        size=14,
        color=MUTED,
    )
    rrect(s, 0.42, 6.28, 12.5, 0.72, PAPER, line=LINE, radius=0.08)
    add_text(
        s,
        "部分条件达不到时，可以只做有偿的策划、嘉宾推荐、科技企业邀约或招商服务，不对整场结果负责。",
        0.62,
        6.4,
        12.1,
        0.48,
        size=14,
        color=INK,
        anchor=MSO_ANCHOR.MIDDLE,
    )


def slide_position(prs):
    s = blank(prs)
    content_header(s, "04    定位", "从学术交流改为场景、转化与资本对接", 5, TOTAL)
    rrect(s, 0.42, 1.2, 12.5, 2.05, INK, radius=0.08)
    add_text(s, "正式名称", 0.68, 1.36, 3, 0.28, size=12, bold=True, color=GOLD)
    add_text(s, "“十五五”银发科技与具身智能康养创新论坛", 0.68, 1.68, 11.9, 0.46, size=22, bold=True, color=WHITE)
    add_text(
        s,
        "副标题    AI养老场景应用 · 主动健康 · 科技成果转化与产业资本对接",
        0.68,
        2.24,
        11.9,
        0.34,
        size=14,
        color=CREAM,
    )
    add_text(s, "传播口号    科技赋能康养，场景驱动创新，资本助力转化。", 0.68, 2.66, 11.9, 0.34, size=14, color=GOLD)
    shifts = [
        ("01", "行业交流", "银发科技产业创新"),
        ("02", "技术展示", "真实养老场景应用"),
        ("03", "学术研讨", "科研、产业、资本和项目落地"),
    ]
    for i, (n, a, b) in enumerate(shifts):
        x = 0.42 + i * 4.18
        rrect(s, x, 3.46, 4.02, 1.85, MIST, radius=0.08)
        add_text(s, n, x + 0.2, 3.62, 1, 0.3, size=13, bold=True, color=TEAL)
        add_text(s, a, x + 0.2, 4.0, 3.6, 0.34, size=14, color=MUTED)
        add_text(s, b, x + 0.2, 4.4, 3.6, 0.64, size=16, bold=True, color=INK)
    add_text(
        s,
        "备选题目仅在资源结构变化时使用：科技投资型、成果转化型、商业合作型。对外统一使用本页正式名称，避免同一活动多个名字。",
        0.42,
        5.5,
        12.5,
        0.7,
        size=14,
        color=MUTED,
    )
    rrect(s, 0.42, 6.22, 12.5, 0.76, CARD_TEAL, radius=0.08)
    add_text(
        s,
        "传统康养地产、一般养生产品和泛文化休闲不作为主线。植物科研、视觉艺术和区域开发项目，没有具体转化成果就不进核心议程。",
        0.62,
        6.34,
        12.1,
        0.52,
        size=14,
        color=INK,
        anchor=MSO_ANCHOR.MIDDLE,
    )


def slide_diff(prs):
    s = blank(prs)
    content_header(s, "05    差异化", "不重复大会已经公布的两场峰会", 6, TOTAL)
    rrect(s, 0.42, 1.2, 6.05, 2.55, MIST, radius=0.08)
    add_text(s, "大会已有安排", 0.64, 1.36, 5.5, 0.32, size=14, bold=True, color=MUTED)
    add_text(s, "智慧康养与AI融合创新峰会", 0.64, 1.78, 5.6, 0.55, size=16, bold=True, color=INK)
    add_text(s, "智慧康养产业政策与投资风向峰会", 0.64, 2.4, 5.6, 0.55, size=16, bold=True, color=INK)
    add_text(s, "本场若再讲一遍“AI+康养”或“康养投资趋势”，就没有独立价值。", 0.64, 3.1, 5.6, 0.45, size=13, color=ROSE)
    rrect(s, 6.67, 1.2, 6.25, 2.55, TEAL, radius=0.08)
    add_text(s, "本场只做四件事", 6.9, 1.36, 5.8, 0.32, size=14, bold=True, color=CARD_GOLD)
    items = ["项目能不能落地", "谁有采购和试点需求", "成果如何转化", "资本如何撮合"]
    for i, item in enumerate(items):
        add_text(s, f"0{i+1}    {item}", 6.9, 1.82 + i * 0.42, 5.7, 0.38, size=16, bold=True, color=WHITE)
    add_text(s, "内容配比", 0.42, 3.95, 3, 0.3, size=13, bold=True, color=TEAL)
    bars = [
        (0.50, "50%  科技应用", "智能护理、康复机器人、外骨骼、监护、AI健康与居家养老", TEAL),
        (0.30, "30%  场景与医疗服务", "慢病早筛、康复、认知、医养衔接、有验证基础的健康管理", GOLD),
        (0.20, "20%  投融资与成果转化", "融资、并购、技术转移、场景采购与准入", INK),
    ]
    y = 4.32
    for ratio, label, desc, color in bars:
        add_text(s, label, 0.42, y, 3.3, 0.32, size=13, bold=True, color=INK)
        rect(s, 3.8, y + 0.08, 8.7 * ratio, 0.18, color)
        rect(s, 3.8 + 8.7 * ratio, y + 0.08, 8.7 * (1 - ratio), 0.18, LINE)
        add_text(s, desc, 0.42, y + 0.32, 12.3, 0.28, size=12, color=MUTED)
        y += 0.72
    notes(s, "政策依据只在口头补充：2026年国家推动人工智能进入养老机构、社区和家庭；上海市民政局2026年首期智慧养老供需对接已有机构与企业结对意向。幻灯片不展开未经核验的具体数字，避免对外误用。")


def slide_topics(prs):
    s = blank(prs)
    content_header(s, "06    内容主线", "三个议题撑起整场，第三题是商业卖点", 7, TOTAL)
    topics = [
        ("01", "具身智能进入养老场景", "从康复机器人到AI养老", "拟邀傅利叶智能分享康复、陪护和规模化部署。优先争取适合会议室的设备演示。未确认前，不把企业名称用于联合冠名。"),
        ("02", "老龄研究与银发经济", "“十五五”新需求与产业机遇", "拟邀复旦大学老龄研究院专家，围绕人口老龄化、主动健康、老年友好和银发经济。住房政策研究中心负责适老居住，不与老龄研究院混为一家。"),
        ("03", "谁为养老机器人买单", "采购、运营、融资与回报", "养老机构、机器人企业、金融或保险、产业投资人同台。回答谁采购、谁运营、谁承担试点成本。这是圆桌，也是招商时真正能卖的权益。"),
    ]
    for i, (n, title, sub, body) in enumerate(topics):
        x = 0.42 + i * 4.18
        rrect(s, x, 1.22, 4.02, 4.55, PAPER, line=LINE, radius=0.08)
        rect(s, x, 1.22, 4.02, 0.1, TEAL if i < 2 else GOLD)
        add_text(s, n, x + 0.22, 1.48, 3.5, 0.36, size=18, bold=True, color=TEAL)
        add_text(s, title, x + 0.22, 1.95, 3.58, 0.85, size=20, bold=True, color=INK)
        add_text(s, sub, x + 0.22, 2.85, 3.58, 0.55, size=14, color=GOLD)
        add_text(s, body, x + 0.22, 3.5, 3.58, 2.0, size=14, color=SLATE, spacing=1.12)
    add_text(
        s,
        "安全边界：不安排未经现场测试的大型人形机器人自主行走、负重护理，也不安排设备直接接触老年体验者。无法到场时，改用授权视频、实物和讲解。",
        0.42,
        5.95,
        12.5,
        0.95,
        size=14,
        color=MUTED,
    )


def slide_org(prs):
    s = blank(prs)
    content_header(s, "07    架构", "五种身份分开，新增单位必须先授权", 8, TOTAL)
    data = [
        ["层级", "拟定单位", "职责"],
        ["大会主办", "沿用康博会官方名单", "品牌、场地和公共资源，不擅自改动"],
        ["分论坛主办", "上海星智耀科技有限公司", "对接组委会，承接原有合作，待最终确认"],
        ["联合承办", "上海市杨浦区科技企业联合会", "AI与机器人企业、产业资本、产业对接"],
        ["联合承办", "复旦大学住房政策研究中心", "适老居住、社区康养、学术资源协调"],
        ["学术支持", "复旦大学老龄研究院", "老龄政策、银发经济、主动健康，单独邀约"],
        ["产业合作", "傅利叶智能、璞跃中国等", "技术展示与创新生态，以确认结果为准"],
        ["场景合作", "市养老服务和老龄产业协会", "采购需求、试点场景、机构邀约"],
    ]
    add_table(s, 0.42, 1.18, 12.5, 4.85, data, [1.7, 4.5, 6.3], size=12)
    add_text(
        s,
        "总协调及产业资源统筹：胡继刚。头衔统一为：复旦大学住房政策研究中心秘书长、副教授级高级工程师、上海市杨浦区科技企业联合会执行会长。",
        0.42,
        6.15,
        12.5,
        0.8,
        size=14,
        color=INK,
    )
    notes(s, "复旦体系机构名称、校名和Logo不能凭个人头衔使用。老龄研究院与住房政策研究中心分开邀约。工商银行上海市分行保留在大会主办层级，分论坛金融招商先做排他审查。")


def slide_adjust(prs):
    s = blank(prs)
    content_header(s, "08    名单", "只保留真正承担工作的机构", 9, TOTAL)
    data = [
        ["原有安排", "处理", "原因"],
        ["上海星智耀", "保留", "继续做成果转化和组委会协调，不另起一套决策链"],
        ["国家技术转移东部中心", "改为拟邀支持", "有技术转移价值，未授权前不列指导单位"],
        ["上海辰山国家植物园", "退出主议程", "与养老科技主线弱，除非另有转化成果"],
        ["上海视觉艺术学院", "定向再邀", "适老化或人机交互有项目时再安排"],
        ["长三角示范区新发展建设", "降为场景候选", "没有确定的养老科技场景就不占席"],
        ["王红霞 / 甘薯科研", "取消上台", "保留在成果储备，不进本场主议程"],
        ["文旅、宠物及泛文化单位", "不列入本论坛", "维持大会原有合作即可，不重复挂名"],
        ["傅利叶智能", "重点新增候选", "先谈内容和演示，不把办会押在赞助款上"],
    ]
    add_table(s, 0.42, 1.16, 12.5, 5.15, data, [3.55, 2.15, 6.8], size=12)
    add_text(s, "原则：宁可5家真做事的机构，也不要20家只有标识的挂名单位。", 0.42, 6.42, 12.5, 0.4, size=15, bold=True, color=TEAL)


def slide_guests_strategy(prs):
    s = blank(prs)
    content_header(s, "09    邀约", "先锁采购方，再锁企业和赞助", 10, TOTAL)
    left = [
        ("第一", "有采购权或试点决策权的养老机构、康复医院、社区运营方。没有买方，论坛就难形成交易。"),
        ("第二", "已有养老、康复、居家照护产品的科技企业负责人，优先创始人或业务负责人。"),
        ("第三", "产业基金合伙人、投资总监、上市公司战略投资负责人，不请只负责宣传的代表。"),
        ("第四", "高校和科研院所里负责成果转化的团队。"),
    ]
    for i, (k, body) in enumerate(left):
        y = 1.18 + i * 1.15
        rrect(s, 0.42, y, 7.55, 1.05, MIST, radius=0.1)
        add_text(s, k, 0.6, y + 0.14, 1.3, 0.76, size=16, bold=True, color=TEAL, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, body, 1.9, y + 0.14, 5.85, 0.78, size=13, color=SLATE, anchor=MSO_ANCHOR.MIDDLE)
    rrect(s, 8.15, 1.18, 4.76, 4.6, INK, radius=0.08)
    add_text(s, "舞台只留8—10人", 8.4, 1.4, 4.3, 0.4, size=18, bold=True, color=WHITE)
    seats = [
        "1位行业协会致辞",
        "1位复旦老龄研究主旨",
        "1位傅利叶主题分享",
        "1位养老机构需求发布",
        "1位资本或金融机构",
        "圆桌4—5人，尽量复用",
    ]
    for i, line in enumerate(seats):
        add_text(s, line, 8.4, 2.0 + i * 0.42, 4.3, 0.38, size=14, color=CREAM)
    add_text(
        s,
        "方法：两位核心人物带动，六位行业专家补充，十家采购机构支撑。方案里的姓名都按拟邀处理，出席、题目和单位授权要重新确认。",
        0.42,
        5.9,
        12.5,
        0.95,
        size=14,
        color=MUTED,
    )


def slide_guest_list(prs):
    s = blank(prs)
    content_header(s, "10    拟邀名单", "学术保证1位，产业保证1位，场景和资本各1位", 11, TOTAL)
    data = [
        ["嘉宾", "身份（邀约前再核对）", "建议角色"],
        ["彭希哲", "复旦大学老龄研究院院长", "学术主旨，可改派代表"],
        ["吴玉韶", "复旦大学老龄研究院副院长、教授", "政策与产业主旨，重点邀约"],
        ["顾捷", "傅利叶智能创始人、CEO", "产业主旨，档期不允许则改业务负责人"],
        ["胡继刚", "住房政策研究中心秘书长；杨浦科企联执行会长", "核心圆桌，资源统筹，不另做主旨"],
        ["朱震", "科技园及杨浦科企联相关负责人，身份待核", "科创与成果转化，建议保留"],
        ["赵博誉", "上海星智耀相关负责人", "转化机制与项目对接，压缩独立发言"],
        ["徐洁平", "璞跃中国，拟邀", "独立主持圆桌，须确认档期和控场"],
        ["场景 / 资本", "协会推荐的机构负责人；投资决策人", "需求发布与商业模式，席位待定"],
    ]
    add_table(s, 0.32, 1.14, 12.7, 5.15, data, [1.7, 6.5, 4.5], size=11)
    add_text(
        s,
        "定向补充、不占主旨：封进看银发经济与养老金融，罗力看健康养老与医养融合。任职和时间以正式邀约核对为准。朱翊元如无即时场景项目，不占核心圆桌。",
        0.4,
        6.38,
        12.5,
        0.6,
        size=13,
        color=MUTED,
    )


def slide_agenda(prs):
    s = blank(prs)
    content_header(s, "11    议程", "11月28日  13:00—16:30  ·  210分钟", 12, TOTAL)
    rows = [
        ("13:00", "开场", "主持人介绍定位、组织和目标"),
        ("13:05", "致辞", "行业协会；两家联合承办统一介绍"),
        ("13:15", "学术主旨", "复旦老龄研究：银发经济与主动健康"),
        ("13:35", "产业主旨", "傅利叶：具身智能与康复养老"),
        ("14:00", "场景需求", "适老居住，以及机构的真实问题"),
        ("14:20", "技术展示", "康复设备，或2个经筛选的项目"),
        ("14:40", "资本观察", "商业模式、采购与投资回报"),
        ("14:50", "茶歇", "参观与自由交流"),
        ("15:00", "产业圆桌", "谁来应用、谁来投资、如何商业化"),
        ("15:45", "供需对接", "预先匹配的分组洽谈"),
        ("16:15", "成果发布", "只发布已经落实的意向和下一步"),
        ("16:25", "闭幕", "总结与合影"),
    ]
    for i, (tm, name, detail) in enumerate(rows):
        col = 0 if i < 6 else 1
        r = i if i < 6 else i - 6
        x = 0.42 + col * 6.45
        y = 1.16 + r * 0.9
        highlight = name == "产业圆桌"
        rrect(s, x, y, 6.28, 0.82, CARD_GOLD if highlight else MIST, radius=0.1)
        add_text(s, tm, x + 0.14, y + 0.12, 1.15, 0.56, size=13, bold=True, color=TEAL, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, name, x + 1.3, y + 0.08, 4.8, 0.32, size=14, bold=True, color=INK)
        add_text(s, detail, x + 1.3, y + 0.4, 4.8, 0.32, size=12, color=MUTED)
    notes(s, "没有成熟签约项目时，成果发布改为试点发布和意向交流。所有路演项目会前收集技术成熟度、场景、采购或融资需求、联系人。")


def slide_roundtable(prs):
    s = blank(prs)
    content_header(s, "12    圆桌", "15:00—15:45   机器人进入养老机构", 13, TOTAL)
    add_text(
        s,
        "从实验室到养老院：谁来应用、谁来投资、如何实现商业化？主持人独立控场。胡继刚不另设主题演讲。",
        0.42,
        1.16,
        12.5,
        0.4,
        size=14,
        color=MUTED,
    )
    people = [
        ("主持", "徐洁平", "璞跃中国，待确认"),
        ("嘉宾", "胡继刚", "研究、场景与资本"),
        ("嘉宾", "傅利叶代表", "产品与产业化瓶颈"),
        ("嘉宾", "老龄研究院专家", "需求、人本与政策"),
        ("嘉宾", "养老机构负责人", "采购、护理与验证"),
        ("嘉宾", "投资或金融", "资金、模式与规模化"),
    ]
    for i, (role, name, point) in enumerate(people):
        x = 0.42 + (i % 6) * 2.12
        rrect(s, x, 1.68, 2.02, 1.55, INK if name == "胡继刚" else PAPER, line=None if name == "胡继刚" else LINE, radius=0.1)
        name_c = WHITE if name == "胡继刚" else INK
        sub_c = GOLD if name == "胡继刚" else TEAL
        body_c = CREAM if name == "胡继刚" else MUTED
        add_text(s, role, x + 0.1, 1.78, 1.82, 0.26, size=11, bold=True, color=sub_c, align=PP_ALIGN.CENTER)
        add_text(s, name, x + 0.08, 2.08, 1.86, 0.55, size=13, bold=True, color=name_c, align=PP_ALIGN.CENTER)
        add_text(s, point, x + 0.1, 2.64, 1.82, 0.46, size=11, color=body_c, align=PP_ALIGN.CENTER)
    rounds = [
        ("15:05", "技术与需求", "每人约3分钟，只回答一个具体问题。"),
        ("15:20", "模式与资本", "谁付费，如何评估效益，谁承担早期试点风险。"),
        ("15:35", "落地动作", "每人只提出一项可执行的下一步。"),
    ]
    for i, (tm, title, body) in enumerate(rounds):
        x = 0.42 + i * 4.18
        rrect(s, x, 3.42, 4.02, 1.45, MIST, radius=0.08)
        add_text(s, tm, x + 0.18, 3.54, 3.6, 0.26, size=12, bold=True, color=TEAL)
        add_text(s, title, x + 0.18, 3.82, 3.66, 0.32, size=15, bold=True, color=INK)
        add_text(s, body, x + 0.18, 4.18, 3.66, 0.52, size=13, color=SLATE)
    rrect(s, 0.42, 5.04, 12.5, 1.92, CARD_TEAL, radius=0.08)
    add_text(s, "胡继刚的讨论位置", 0.64, 5.18, 6, 0.3, size=13, bold=True, color=TEAL)
    add_text(
        s,
        "海报口径：《从适老居住到智慧康养：高校研究、产业场景与资本协同创新》",
        0.64,
        5.5,
        12.0,
        0.32,
        size=14,
        bold=True,
        color=INK,
    )
    add_text(
        s,
        "第一轮重点：高校科研、适老居住和科技企业如何一起为养老机器人找到可持续市场。第二轮重点：以真实需求组织小规模试点，形成可评估的效果、成本和推广方式，再谈融资。不重复傅利叶的技术分享，也不重复老龄专家的政策报告。",
        0.64,
        5.9,
        12.0,
        0.88,
        size=13,
        color=SLATE,
    )


def slide_sponsor_tiers(prs):
    s = blank(prs)
    content_header(s, "13    招商", "卖场景对接，不卖标识面积", 14, TOTAL)
    tiers = [
        ("首席产业合作伙伴", "50,000元", "1家", "核心展示、审核后的主题分享、圆桌候选、演示空间、定向对接、6个名额"),
        ("产业创新合作伙伴", "20,000元", "2家", "品牌展示、筛选后的路演、圆桌或交流机会、定向推荐、4个名额"),
        ("专项生态合作伙伴", "10,000元", "2家", "电子屏与物料、供需对接、2个名额、新闻稿鸣谢"),
        ("展示及服务支持", "实物或服务", "若干", "设备、摄影、物料、茶歇。按实际贡献协商，不单列虚高现金价"),
    ]
    for i, (name, price, quota, rights) in enumerate(tiers):
        y = 1.18 + i * 1.15
        rrect(s, 0.42, y, 12.5, 1.05, PAPER, line=LINE, radius=0.1)
        rect(s, 0.42, y, 0.08, 1.05, GOLD if i == 0 else TEAL)
        add_text(s, name, 0.7, y + 0.12, 4.3, 0.36, size=16, bold=True, color=INK)
        add_text(s, rights, 0.7, y + 0.52, 7.3, 0.4, size=13, color=MUTED)
        add_text(s, price, 8.3, y + 0.22, 2.7, 0.58, size=18, bold=True, color=TEAL, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, quota, 11.15, y + 0.28, 1.5, 0.48, size=14, bold=True, color=INK, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    add_text(
        s,
        "拟议价格，须经组委会批准，且不得与已售权益冲突。不承诺央媒、电视、领导会见、融资成功、采购成交，也不共享没有合法依据的个人名单。",
        0.42,
        5.85,
        12.5,
        0.95,
        size=14,
        color=ROSE,
    )


def slide_money(prs):
    s = blank(prs)
    content_header(s, "14    财务", "基准看8万元收入、2.3万元支出", 15, TOTAL)
    targets = [("最低", "3万", "继续办的现金底线"), ("基准", "8万", "5万 + 2万 + 1万"), ("理想", "11万", "不作为对外承诺")]
    for i, (name, num, desc) in enumerate(targets):
        x = 0.42 + i * 2.55
        rrect(s, x, 1.18, 2.42, 1.55, INK if i == 1 else MIST, radius=0.1)
        add_text(s, name, x + 0.12, 1.28, 2.18, 0.28, size=12, bold=True, color=GOLD if i == 1 else TEAL, align=PP_ALIGN.CENTER)
        add_text(s, num, x + 0.12, 1.56, 2.18, 0.58, size=26, bold=True, color=WHITE if i == 1 else INK, align=PP_ALIGN.CENTER)
        add_text(s, desc, x + 0.12, 2.2, 2.18, 0.38, size=12, color=CREAM if i == 1 else MUTED, align=PP_ALIGN.CENTER)
    rrect(s, 8.1, 1.18, 4.82, 1.55, CARD_GOLD, radius=0.1)
    add_text(s, "执行预算上限", 8.28, 1.3, 4.5, 0.28, size=12, bold=True, color=GOLD)
    add_text(s, "23,000元", 8.28, 1.58, 4.5, 0.5, size=26, bold=True, color=INK)
    add_text(s, "假定场地和基础设备免费", 8.28, 2.16, 4.5, 0.36, size=13, color=SLATE)
    lines = [
        "物料1,500  ·  摄影3,500  ·  现场人员3,000",
        "嘉宾接待3,000  ·  茶歇2,000  ·  设备运输3,500",
        "补充技术3,000  ·  预备3,500",
    ]
    for i, line in enumerate(lines):
        add_text(s, line, 0.42, 2.9 + i * 0.32, 12.4, 0.3, size=14, color=SLATE)
    rrect(s, 0.42, 3.95, 12.5, 1.45, ROSE_SOFT, radius=0.08)
    add_text(s, "三条约定", 0.64, 4.08, 4, 0.28, size=13, bold=True, color=ROSE)
    add_text(s, "8万元若回款、支出2.3万元，剩余5.7万元仍要扣税、招商服务费和各方分配，不是个人利润。", 0.64, 4.4, 12.0, 0.36, size=14, color=INK)
    add_text(s, "傅利叶先作为内容和演示合作方。未确认前不使用其名称联合冠名，也不把办会押在它是否付款。", 0.64, 4.78, 12.0, 0.48, size=14, color=INK)
    add_text(
        s,
        "收款、开票和合同由经授权的法人主体办理。不宜由个人先垫付这2.3万元。高校和协会的收支必须符合各自财务制度。",
        0.42,
        5.55,
        12.5,
        0.7,
        size=14,
        color=MUTED,
    )
    notes(s, "内部测算还有三档情景：保守收入4万成本2.5万；基准收入9万成本3.5万；乐观收入16万成本5万。这些是建议书早期模型，执行以本页8万/2.3万为准，且都不是收入承诺。")


def slide_plan(prs):
    s = blank(prs)
    content_header(s, "15    执行", "八项任务，责任写到主体", 16, TOTAL)
    rows = [
        ["11/11—13", "架构与授权", "两家联合承办授权、结算方式、场地清单", "星智耀 + 杨浦科企联 + 住策中心"],
        ["11/12—14", "核心邀约", "老龄专家、傅利叶、两家养老场景机构", "学术组 + 产业组"],
        ["11/13—16", "赞助谈判", "10—15家定向发送，先锁1家主赞助", "招商组"],
        ["11/16—20", "合同回款", "至少3万元到账，再锁定支出", "商务与财务"],
        ["11/18—23", "人员与项目", "议程、采购方、投资人、路演材料", "会务组 + 协会"],
        ["11/23—26", "现场锁定", "演示、主持稿、座席、安全和备份", "执行组 + 技术组"],
        ["11/27—28", "联调举办", "接待、签到、记录需求与洽谈", "现场总控"],
        ["11/29—12/5", "成果结算", "影像、权益核销、费用和合作台账", "项目负责人 + 财务"],
    ]
    add_table(s, 0.32, 1.16, 12.7, 5.0, [["时间", "任务", "完成标准", "责任"] , *rows], [1.7, 1.7, 5.3, 4.0], size=11)
    add_text(s, "详细勾选、状态下拉和自动结论在《执行管理台账》。", 0.42, 6.3, 12.4, 0.4, size=14, color=TEAL)


def slide_close(prs):
    s = blank(prs)
    content_header(s, "16    收口", "先有内容和买方，再拿着资源去招商", 17, TOTAL)
    metrics = [
        ("10项", "可讨论的场景或合作需求"),
        ("2—3个", "有基础的科技项目"),
        ("10—15组", "有效商务洽谈"),
        ("2—3项", "进入验证或试点讨论"),
    ]
    for i, (n, label) in enumerate(metrics):
        x = 0.42 + i * 3.2
        rrect(s, x, 1.18, 3.05, 1.2, MIST, radius=0.1)
        add_text(s, n, x + 0.12, 1.28, 2.8, 0.42, size=20, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
        add_text(s, label, x + 0.12, 1.74, 2.8, 0.46, size=13, color=SLATE, align=PP_ALIGN.CENTER)
    asks = [
        ("定位", "改为银发科技与具身智能产业化论坛，保留主动健康。"),
        ("架构", "增补两家联合承办。老龄研究院与傅利叶须另行确认。"),
        ("议程", "删去与主线不符的报告。核心换成老龄研究、机器人、场景需求和资本圆桌。"),
        ("财务", "独立招商、独立预算、书面结算。投入由到账和已落实资源决定。"),
    ]
    for i, (title, body) in enumerate(asks):
        col = i % 2
        row = i // 2
        x = 0.42 + col * 6.45
        y = 2.56 + row * 1.15
        rrect(s, x, y, 6.28, 1.05, PAPER, line=LINE, radius=0.1)
        add_text(s, title, x + 0.18, y + 0.12, 1.2, 0.8, size=15, bold=True, color=TEAL, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, body, x + 1.4, y + 0.14, 4.7, 0.78, size=13, color=SLATE, anchor=MSO_ANCHOR.MIDDLE)
    rrect(s, 0.42, 4.98, 12.5, 1.95, INK, radius=0.08)
    add_text(s, "本周只做三件事", 0.66, 5.12, 6, 0.3, size=13, bold=True, color=GOLD)
    add_text(s, "第一，先谈承办和结算，不先找赞助。", 0.66, 5.48, 12.0, 0.32, size=15, color=WHITE)
    add_text(s, "第二，先锁定愿意谈试点的养老机构，再接触机器人和数字医疗企业。", 0.66, 5.84, 12.0, 0.32, size=15, color=WHITE)
    add_text(s, "第三，删掉无法控制的央媒、电视、名单共享和未授权的主办身份。", 0.66, 6.2, 12.0, 0.5, size=15, color=WHITE)


def slide_end(prs):
    s = blank(prs)
    rect(s, 0, 0, SW, 7.5, INK)
    rect(s, 0, 0, 0.14, 7.5, GOLD)
    add_text(s, "内部执行稿", 0.7, 1.35, 10, 0.3, size=14, bold=True, color=GOLD)
    add_text(s, "以研究识别需求", 0.68, 1.9, 12, 0.55, size=28, bold=True, color=WHITE)
    add_text(s, "以科技回应需求", 0.68, 2.55, 12, 0.55, size=28, bold=True, color=WHITE)
    add_text(s, "以场景验证技术", 0.68, 3.2, 12, 0.55, size=28, bold=True, color=WHITE)
    add_text(s, "以资本促进转化", 0.68, 3.85, 12, 0.55, size=28, bold=True, color=GOLD)
    add_text(
        s,
        "配套文件：活动执行策划案  ·  执行管理台账  ·  对外赞助方案",
        0.7,
        5.15,
        12,
        0.4,
        size=16,
        color=CREAM,
    )
    add_text(
        s,
        "招商合同和财务结算由指定法人主体管理。对外身份、内容安排和商业责任分开。",
        0.7,
        5.7,
        11.5,
        0.7,
        size=15,
        color=CREAM,
    )
    add_text(s, "18  /  18", 11.2, 6.9, 1.6, 0.3, size=12, bold=True, color=GOLD, align=PP_ALIGN.RIGHT)


def build():
    prs = new_prs()
    cover(prs)
    slide_judgment(prs)
    slide_rights(prs)
    slide_gates(prs)
    slide_position(prs)
    slide_diff(prs)
    slide_topics(prs)
    slide_org(prs)
    slide_adjust(prs)
    slide_guests_strategy(prs)
    slide_guest_list(prs)
    slide_agenda(prs)
    slide_roundtable(prs)
    slide_sponsor_tiers(prs)
    slide_money(prs)
    slide_plan(prs)
    slide_close(prs)
    slide_end(prs)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    prs.save(OUT)
    return OUT


if __name__ == "__main__":
    print(build())
