"""对外赞助介绍。不含内部预算、分成和退出条件。"""

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
    SLATE,
    SW,
    TEAL,
    WHITE,
    add_table,
    add_text,
    blank,
    content_header,
    notes,
    rect,
    rrect,
)

TOTAL = 13
OUT = os.path.normpath(
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "..",
        "deliverables",
        "2026康博会专题分论坛",
        "05-对外-赞助合作介绍.pptx",
    )
)


def new_prs():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = "2026康博会专题分论坛赞助及产业合作方案"
    prs.core_properties.subject = "对外拟议版"
    prs.core_properties.category = "对外"
    prs.core_properties.author = "专题分论坛筹备组"
    return prs


def cover(prs):
    s = blank(prs)
    rect(s, 0, 0, SW, 7.5, INK)
    rect(s, 0, 0, 0.14, 7.5, GOLD)
    add_text(s, "对外拟议版    ·    2026年11月11日", 0.62, 0.4, 8, 0.3, size=14, color=GOLD)
    add_text(s, "以正式合同为准", 9.4, 0.4, 3.4, 0.3, size=14, bold=True, color=GOLD, align=PP_ALIGN.RIGHT)
    add_text(s, "2026上海国际智慧康养产业博览会", 0.62, 1.45, 12, 0.36, size=18, color=CREAM)
    add_text(s, "专题分论坛", 0.58, 1.95, 12, 0.7, size=44, bold=True, color=WHITE)
    add_text(s, "赞助及产业合作方案", 0.62, 2.75, 12, 0.55, size=30, bold=True, color=GOLD)
    rect(s, 0.62, 3.5, 1.35, 0.045, GOLD)
    add_text(s, "“十五五”银发科技与具身智能康养创新论坛", 0.62, 3.75, 12, 0.42, size=20, color=WHITE)
    add_text(
        s,
        "AI养老场景应用  ·  主动健康  ·  科技成果转化与产业资本对接",
        0.62,
        4.25,
        12,
        0.34,
        size=15,
        color=CREAM,
    )
    chips = [("11月28日", "13:00—16:30"), ("上海展览中心", "会议室以组委会确认为准"), ("约120人", "定向邀请")]
    for i, (a, b) in enumerate(chips):
        x = 0.62 + i * 4.1
        add_text(s, a, x, 5.35, 3.8, 0.34, size=18, bold=True, color=WHITE)
        add_text(s, b, x, 5.75, 3.8, 0.3, size=13, color=CREAM)
    add_text(
        s,
        "联合承办单位、学术机构和产业企业名称，须取得正式授权后方可使用。",
        0.62,
        6.55,
        12,
        0.4,
        size=14,
        color=GOLD,
    )
    notes(s, "发出前请填写授权联系人。未授权时，口头介绍可说拟邀方向，不要说某企业或高校已经加盟。")


def slide_why(prs):
    s = blank(prs)
    content_header(s, "合作价值", "一场把技术、场景和资本放在同一张桌上的对接", 2, TOTAL, internal=False)
    points = [
        ("看见真实需求", "养老运营、康复医疗和社区服务机构直接说明护理、安全和设备使用中的问题。"),
        ("展示可落地的技术", "聚焦具身智能、康复机器人、智慧养老和数字健康，不做泛主题展览。"),
        ("讨论谁来买单", "把采购、租赁、服务付费、试点资金和投资回报放进同一场对话。"),
        ("留下可跟进的项目", "会后形成需求、方案、联系人和下一步，而不是只交换名片。"),
    ]
    for i, (title, body) in enumerate(points):
        col = i % 2
        row = i // 2
        x = 0.42 + col * 6.45
        y = 1.25 + row * 2.15
        rrect(s, x, y, 6.28, 2.0, PAPER, line=LINE, radius=0.08)
        rect(s, x, y, 0.08, 2.0, GOLD if i == 2 else TEAL)
        add_text(s, f"0{i+1}", x + 0.3, y + 0.22, 1.2, 0.32, size=14, bold=True, color=TEAL)
        add_text(s, title, x + 0.3, y + 0.58, 5.7, 0.4, size=20, bold=True, color=INK)
        add_text(s, body, x + 0.3, y + 1.1, 5.7, 0.7, size=14, color=SLATE)
    add_text(
        s,
        "科技赋能康养，场景驱动创新，资本助力转化。",
        0.42,
        5.7,
        12.5,
        0.4,
        size=16,
        bold=True,
        color=TEAL,
    )
    add_text(
        s,
        "本方案是论坛专项合作的拟议文本。具体身份、品牌露出、发言和展示，以组委会批准的合同及附件为准。",
        0.42,
        6.15,
        12.5,
        0.7,
        size=14,
        color=MUTED,
    )


def slide_facts(prs):
    s = blank(prs)
    content_header(s, "活动概况", "时间、地点和规模一次说清", 3, TOTAL, internal=False)
    facts = [
        ("时间", "2026年11月28日\n星期六  13:00—16:30"),
        ("地点", "上海展览中心\n中央大厅会议室C（拟）"),
        ("规模", "100—120人\n定向邀请，不对公众售票"),
        ("形式", "主旨报告、技术展示\n产业圆桌、分组对接"),
    ]
    for i, (k, v) in enumerate(facts):
        x = 0.42 + i * 3.2
        rrect(s, x, 1.22, 3.05, 2.15, MIST, radius=0.08)
        add_text(s, k, x + 0.18, 1.38, 2.7, 0.3, size=13, bold=True, color=TEAL)
        add_text(s, v, x + 0.18, 1.78, 2.7, 1.35, size=16, bold=True, color=INK, spacing=1.15)
    add_text(s, "拟议中的组织基础", 0.42, 3.58, 6, 0.32, size=14, bold=True, color=TEAL)
    orgs = [
        ("大会平台", "沿用2026上海国际智慧康养产业博览会主办体系"),
        ("联合承办", "上海市杨浦区科技企业联合会；复旦大学住房政策研究中心"),
        ("学术方向", "拟邀复旦大学老龄研究院提供老龄研究支持"),
        ("产业方向", "拟邀康复机器人、具身智能及智慧养老企业参与"),
    ]
    for i, (k, v) in enumerate(orgs):
        y = 4.0 + i * 0.68
        add_text(s, k, 0.42, y, 1.8, 0.5, size=14, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, v, 2.3, y, 10.4, 0.5, size=14, color=SLATE, anchor=MSO_ANCHOR.MIDDLE)
    notes(s, "傅利叶智能、璞跃中国等具体企业名称，只有在对方书面确认后才写入对外终稿。本页用产业方向表述。")


def slide_topics(prs):
    s = blank(prs)
    content_header(s, "讨论什么", "五个问题，服务同一条产业化路径", 4, TOTAL, internal=False)
    topics = [
        ("01", "银发经济与主动健康", "“十五五”时期的养老服务需求，以及科技进入机构、社区和家庭的机会。"),
        ("02", "具身智能与康复养老", "康复机器人、人机交互、智能陪护如何从技术展示走到场景应用。"),
        ("03", "适老居住与社区场景", "智慧居家、嵌入式养老、适老化改造，以及设备如何进入真实空间。"),
        ("04", "机构采购的真实约束", "护理人力、康复训练、安全监测、设备投资和长期运营成本。"),
        ("05", "商业化与产业金融", "采购、租赁、按服务付费，以及银行、保险和产业基金如何参与早期试点。"),
    ]
    for i, (n, title, body) in enumerate(topics):
        y = 1.18 + i * 1.05
        rrect(s, 0.42, y, 12.5, 0.95, PAPER if i % 2 == 0 else MIST, line=LINE, radius=0.1)
        add_text(s, n, 0.62, y + 0.22, 0.8, 0.5, size=16, bold=True, color=TEAL, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, title, 1.55, y + 0.12, 4.3, 0.7, size=16, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, body, 5.9, y + 0.14, 6.75, 0.68, size=14, color=SLATE, anchor=MSO_ANCHOR.MIDDLE)


def slide_agenda(prs):
    s = blank(prs)
    content_header(s, "议程", "趋势、需求、技术、资本，然后进入对接", 5, TOTAL, internal=False)
    data = [
        ["时间", "环节", "合作企业可以参与的方式"],
        ["13:00—13:15", "开场与行业致辞", "了解论坛定位与组织单位"],
        ["13:15—14:00", "学术与产业主旨", "首席伙伴可申请审核后的主题分享"],
        ["14:00—14:40", "场景需求与技术展示", "发布需求，或进行筛选后的案例展示"],
        ["14:40—15:00", "资本观察与茶歇", "与机构、投资人做初步交流"],
        ["15:00—15:45", "产业圆桌", "讨论采购、运营和试点成本"],
        ["15:45—16:25", "分组对接与成果", "按事先匹配的需求进行洽谈"],
    ]
    add_table(s, 0.42, 1.2, 12.5, 4.55, data, [2.3, 3.3, 6.9], size=13)
    add_text(
        s,
        "发言、圆桌和展示都需要通过内容和安全审核。没有成熟签约基础时，论坛发布合作意向和后续安排，不安排形式化签约。",
        0.42,
        5.95,
        12.5,
        0.85,
        size=14,
        color=MUTED,
    )


def slide_audience(prs):
    s = blank(prs)
    content_header(s, "邀请结构", "按买方、技术和资本来组织席位", 6, TOTAL, internal=False)
    add_text(s, "以下是邀约目标，不是已经确认的到场名单。", 0.42, 1.16, 12, 0.32, size=14, color=MUTED)
    rows = [
        ("养老与康复机构", "10—15家", "提出场景、采购和试点条件"),
        ("康养科技企业", "15—20家", "提供可讨论的产品与合作需求"),
        ("金融和投资机构", "5—8家", "讨论资金支持与商业模式"),
        ("高校与成果转化", "定向邀请", "连接研究、技术转移和应用"),
        ("行业协会与平台", "定向邀请", "组织专业观众和需求征集"),
    ]
    for i, (name, num, desc) in enumerate(rows):
        y = 1.6 + i * 0.9
        rrect(s, 0.42, y, 12.5, 0.8, MIST, radius=0.1)
        add_text(s, name, 0.64, y + 0.16, 3.4, 0.48, size=16, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, num, 4.2, y + 0.16, 2.3, 0.48, size=16, bold=True, color=TEAL, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, desc, 6.6, y + 0.16, 6.0, 0.48, size=15, color=SLATE, anchor=MSO_ANCHOR.MIDDLE)
    notes(s, "不要把目标人数说成已报名人数。养老机构首先作为买方邀请，不在初次接触时要求高额赞助。")


def slide_levels(prs):
    s = blank(prs)
    content_header(s, "合作级别", "四个层级，按对接深度而不是标识大小区分", 7, TOTAL, internal=False)
    levels = [
        ("首席产业合作伙伴", "50,000元", "1家", "主题分享、展示、圆桌候选和专项对接"),
        ("产业创新合作伙伴", "20,000元", "2家", "案例展示、定向推荐和交流席位"),
        ("专项生态合作伙伴", "10,000元", "2家", "品牌鸣谢、会场展示和商务交流"),
        ("展示及服务支持伙伴", "实物或服务", "协商", "设备、物料、摄影、茶歇或技术支持"),
    ]
    for i, (name, price, quota, point) in enumerate(levels):
        y = 1.22 + i * 1.2
        rrect(s, 0.42, y, 12.5, 1.08, INK if i == 0 else PAPER, line=None if i == 0 else LINE, radius=0.1)
        name_c = WHITE if i == 0 else INK
        sub_c = CREAM if i == 0 else MUTED
        price_c = GOLD if i == 0 else TEAL
        add_text(s, name, 0.68, y + 0.14, 6.4, 0.38, size=18, bold=True, color=name_c)
        add_text(s, point, 0.68, y + 0.56, 6.6, 0.36, size=13, color=sub_c)
        add_text(s, price, 7.6, y + 0.22, 3.1, 0.62, size=20, bold=True, color=price_c, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, quota, 10.9, y + 0.28, 1.7, 0.5, size=16, bold=True, color=name_c, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def slide_chief(prs):
    s = blank(prs)
    content_header(s, "首席产业合作伙伴", "50,000元    ·    1家", 8, TOTAL, internal=False)
    rights = [
        ("品牌", "物料展示“首席产业合作伙伴”名称和标识。可按主营业务设置专项称谓。"),
        ("内容", "经议题审核，1位企业代表做10—15分钟主题报告，并获得1个圆桌候选席位。"),
        ("场景", "在满足场馆和安全条件时安排展示或演示，并组织与康养机构的定向对接。"),
        ("名额与资料", "6个参会名额。新闻稿和成果总结体现实际合作，照片视频经许可后使用。"),
    ]
    for i, (title, body) in enumerate(rights):
        col = i % 2
        row = i // 2
        x = 0.42 + col * 6.45
        y = 1.22 + row * 2.15
        rrect(s, x, y, 6.28, 2.0, CARD_GOLD if i == 0 else MIST, radius=0.08)
        add_text(s, title, x + 0.24, y + 0.22, 5.8, 0.4, size=18, bold=True, color=INK)
        add_text(s, body, x + 0.24, y + 0.75, 5.8, 1.0, size=15, color=SLATE, spacing=1.12)
    add_text(s, "称谓示例：具身智能康养创新合作伙伴，或智慧养老场景创新合作伙伴。", 0.42, 5.7, 12.5, 0.4, size=15, color=TEAL)
    add_text(s, "主题报告和圆桌席位以议程审核结果为准，不自动等于产品背书。", 0.42, 6.15, 12.5, 0.55, size=14, color=MUTED)


def slide_other(prs):
    s = blank(prs)
    content_header(s, "创新伙伴与生态伙伴", "2万元和1万元两档", 9, TOTAL, internal=False)
    rrect(s, 0.42, 1.2, 6.2, 4.55, PAPER, line=LINE, radius=0.08)
    add_text(s, "产业创新合作伙伴", 0.66, 1.4, 5.7, 0.36, size=18, bold=True, color=INK)
    add_text(s, "20,000元    ·    2家", 0.66, 1.84, 5.7, 0.32, size=14, bold=True, color=TEAL)
    items = [
        "以产业创新合作伙伴身份展示品牌",
        "提交项目，经筛选后参与展示或路演",
        "1个圆桌候选或产业交流机会",
        "定向推荐相关机构或科技企业资源",
        "4个参会名额",
        "现场及新闻稿体现品牌，提供经许可的照片和总结",
    ]
    for i, item in enumerate(items):
        add_text(s, item, 0.66, 2.35 + i * 0.5, 5.7, 0.42, size=14, color=SLATE)
    rrect(s, 6.82, 1.2, 6.1, 4.55, TEAL, radius=0.08)
    add_text(s, "专项生态合作伙伴", 7.06, 1.4, 5.6, 0.36, size=18, bold=True, color=WHITE)
    add_text(s, "10,000元    ·    2家", 7.06, 1.84, 5.6, 0.32, size=14, bold=True, color=GOLD)
    items2 = [
        "宣传物料和现场电子屏展示",
        "参加专项供需对接和商务交流",
        "2个嘉宾或企业代表名额",
        "可提交产品介绍和合作需求",
        "新闻稿或鸣谢材料列明合作身份",
        "获得成果总结及可合法共享的合作信息",
    ]
    for i, item in enumerate(items2):
        add_text(s, item, 7.06, 2.35 + i * 0.5, 5.6, 0.42, size=14, color=WHITE)
    add_text(s, "两档的发言和席位都要经过议程审核。名额用于企业代表，不对外转让。", 0.42, 5.95, 12.5, 0.75, size=14, color=MUTED)


def slide_inkind(prs):
    s = blank(prs)
    content_header(s, "展示及服务支持", "用设备、物料和服务参与", 10, TOTAL, internal=False)
    kinds = [
        ("设备展示", "康复设备、机器人或相关硬件。须先确认运输、电源、通道、操作人员和安全方案。"),
        ("会务支持", "摄影摄像、物料设计和现场服务。交付标准在合作前写明。"),
        ("接待支持", "茶歇、饮用水或伴手礼。不因实物支持改变论坛的内容和嘉宾结构。"),
        ("专业支持", "项目路演辅导或技术讲解。内容须符合产品实际能力和宣传规范。"),
    ]
    for i, (title, body) in enumerate(kinds):
        y = 1.2 + i * 1.15
        rrect(s, 0.42, y, 12.5, 1.05, MIST, radius=0.1)
        add_text(s, f"0{i+1}", 0.64, y + 0.28, 0.7, 0.48, size=16, bold=True, color=TEAL, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, title, 1.5, y + 0.14, 2.5, 0.76, size=16, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, body, 4.15, y + 0.16, 8.45, 0.74, size=14, color=SLATE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, "权益根据贡献、资源价值和论坛需求协商，不套用现金档位的全部权益。", 0.42, 5.9, 12.5, 0.7, size=15, color=INK)


def slide_who(prs):
    s = blank(prs)
    content_header(s, "欢迎哪些伙伴", "五类企业最匹配这场论坛", 11, TOTAL, internal=False)
    groups = [
        ("具身智能与康复设备", "机器人整机、外骨骼、智能护理、传感器和康复器械企业。"),
        ("数字健康与智慧养老", "AI慢病管理、远程监护、智慧养老平台和健康管理企业。"),
        ("养老运营与医疗康复", "有试点、采购或品牌合作需求的机构集团和康复服务企业。"),
        ("银行、保险与产业基金", "养老金融、科技金融、保险养老和有决策权的投资机构。"),
        ("核心配套与科技服务", "机器人零部件、医疗器械、技术转移和专业服务机构。"),
    ]
    for i, (title, body) in enumerate(groups):
        y = 1.18 + i * 1.02
        rrect(s, 0.42, y, 12.5, 0.92, PAPER, line=LINE, radius=0.1)
        rect(s, 0.42, y, 0.08, 0.92, GOLD if i == 3 else TEAL)
        add_text(s, title, 0.72, y + 0.1, 4.5, 0.72, size=16, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, body, 5.3, y + 0.12, 7.3, 0.68, size=14, color=SLATE, anchor=MSO_ANCHOR.MIDDLE)


def slide_boundary(prs):
    s = blank(prs)
    content_header(s, "合作边界", "写进合同之前，先把不能承诺的事情说清楚", 12, TOTAL, internal=False)
    rows = [
        ("对接是撮合", "论坛组织需求交流、项目展示和资本见面。不承诺采购成交、投资收益、融资结果或政府审批。"),
        ("不构成背书", "赞助不等于高校、科研机构或专家对产品质量、疗效和投资价值的证明。"),
        ("媒体看落实", "报道、采访和视频传播以已经落实的媒体资源为准。没有书面约定的报道不写入权益。"),
        ("名单要合法", "个人联系方式和参会名单，只在符合法律规定并取得必要授权时提供。"),
        ("展示要合规", "医疗器械和养老机器人的功能宣传应符合产品实际。现场演示须通过安全和场馆审批。"),
        ("名称要授权", "高校校名、机构名称和企业标识，取得书面同意后才能用于物料、新闻稿和招商。"),
    ]
    for i, (title, body) in enumerate(rows):
        col = i % 2
        row = i // 2
        x = 0.42 + col * 6.45
        y = 1.18 + row * 1.8
        rrect(s, x, y, 6.28, 1.66, CARD_TEAL if i % 2 == 0 else CARD_GOLD, radius=0.08)
        add_text(s, title, x + 0.2, y + 0.16, 5.9, 0.36, size=16, bold=True, color=INK)
        add_text(s, body, x + 0.2, y + 0.58, 5.9, 0.9, size=13, color=SLATE)


def slide_next(prs):
    s = blank(prs)
    content_header(s, "下一步", "确认意向后，由授权联系人对接合同", 13, TOTAL, internal=False)
    steps = [
        ("01", "提交兴趣", "说明希望的合作级别、展示需求和参会人数。"),
        ("02", "核对权益", "组委会确认该权益尚未售出，且与大会赞助不冲突。"),
        ("03", "签署合同", "由具备履约能力的主体签约、开票和收款。"),
        ("04", "准备对接", "提交产品说明、目标场景，以及采购或融资需求。"),
    ]
    for i, (n, title, body) in enumerate(steps):
        x = 0.42 + i * 3.2
        rrect(s, x, 1.22, 3.05, 2.35, MIST, radius=0.08)
        add_text(s, n, x + 0.16, 1.38, 2.7, 0.32, size=14, bold=True, color=TEAL)
        add_text(s, title, x + 0.16, 1.78, 2.7, 0.4, size=18, bold=True, color=INK)
        add_text(s, body, x + 0.16, 2.28, 2.72, 1.05, size=13, color=SLATE)
    rrect(s, 0.42, 3.78, 12.5, 2.15, INK, radius=0.08)
    add_text(s, "合作咨询", 0.68, 3.96, 4, 0.3, size=13, bold=True, color=GOLD)
    add_text(s, "联系人    ________________________", 0.68, 4.4, 7.5, 0.36, size=16, color=WHITE)
    add_text(s, "电话        ________________________", 0.68, 4.84, 7.5, 0.36, size=16, color=WHITE)
    add_text(s, "邮箱        ________________________", 0.68, 5.28, 7.5, 0.4, size=16, color=WHITE)
    add_text(s, "发出前填写。\n只使用组委会\n指定的联系人。", 8.6, 4.25, 3.9, 1.3, size=14, color=CREAM)
    add_text(
        s,
        "诚邀产业创新企业、金融机构及生态伙伴参与。以前沿科技推动养老服务创新，以真实场景促成产业合作。",
        0.42,
        6.1,
        12.5,
        0.7,
        size=14,
        color=MUTED,
    )


def build():
    prs = new_prs()
    cover(prs)
    slide_why(prs)
    slide_facts(prs)
    slide_topics(prs)
    slide_agenda(prs)
    slide_audience(prs)
    slide_levels(prs)
    slide_chief(prs)
    slide_other(prs)
    slide_inkind(prs)
    slide_who(prs)
    slide_boundary(prs)
    slide_next(prs)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    prs.save(OUT)
    return OUT


if __name__ == "__main__":
    print(build())
