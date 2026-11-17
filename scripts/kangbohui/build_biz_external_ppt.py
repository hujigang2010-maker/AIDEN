"""对外可发送的赞助与展位方案。不含佣金、垫资和小型展台报价。"""

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
    set_cell_fill,
)

TOTAL = 12
OUT = os.path.normpath(
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "..",
        "deliverables",
        "2026康博会专题分论坛",
        "09-对外-赞助与展位合作方案.pptx",
    )
)


def new_prs():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = "2026康博会产业创新赞助与展位合作方案"
    prs.core_properties.subject = "对外拟议版"
    prs.core_properties.category = "对外"
    prs.core_properties.author = "专题分论坛筹备组"
    return prs


def cover(prs):
    s = blank(prs)
    rect(s, 0, 0, SW, 7.5, INK)
    rect(s, 0, 0, 0.14, 7.5, GOLD)
    add_text(s, "对外合作方案    ·    2026年11月17日", 0.62, 0.38, 8, 0.3, size=14, color=GOLD)
    add_text(s, "以正式合同为准", 9.2, 0.38, 3.6, 0.3, size=14, bold=True, color=GOLD, align=PP_ALIGN.RIGHT)
    add_text(s, "2026上海国际智慧康养产业博览会", 0.62, 1.28, 12, 0.34, size=18, color=CREAM)
    add_text(s, "AI重塑养老未来", 0.58, 1.78, 12, 0.72, size=40, bold=True, color=WHITE)
    add_text(s, "银发科技与具身智能产业创新论坛", 0.62, 2.58, 12, 0.46, size=26, bold=True, color=GOLD)
    rect(s, 0.62, 3.2, 1.35, 0.045, GOLD)
    add_text(s, "两种参与方式：产业创新赞助，或三天标准展位。", 0.62, 3.42, 12, 0.36, size=18, color=WHITE)
    chips = [
        ("论坛", "11月28日  13:00—16:30"),
        ("展期", "11月27日—29日"),
        ("地点", "上海展览中心"),
    ]
    for i, (a, b) in enumerate(chips):
        x = 0.62 + i * 4.1
        add_text(s, a, x, 4.55, 3.8, 0.28, size=14, bold=True, color=GOLD)
        add_text(s, b, x, 4.9, 3.8, 0.36, size=18, bold=True, color=WHITE)
    add_text(
        s,
        "拟增补联合承办：上海市杨浦区科技企业联合会、复旦大学住房政策研究中心。\n正式署名以授权和组委会确认为准。企业名称在书面确认前均为邀约方向。",
        0.62,
        5.85,
        12,
        0.7,
        size=15,
        color=CREAM,
        spacing=1.15,
    )
    notes(s, "本套可发给企业。发给瑞健未来以外的对象时，先删去“瑞健未来”专页。不要同时发送较早的赞助介绍，以免名额口径不一致。")


def prices(prs):
    s = blank(prs)
    content_header(s, "建议报价", "赞助主推两万元，展位采用大会公开标准", 2, TOTAL, internal=False)
    cards = [
        (TEAL, "方案一  ·  产业创新赞助", "¥20,000", "每家", "建议主推。品牌展示、专业分享申请、产业对接。"),
        (GOLD, "方案二  ·  标准展位", "¥16,800", "每个9㎡  ·  展期三天", "大会公开的国内标准展位价格。双开口另加10%。"),
        (INK, "二十个展位的规划总额", "¥336,000", "按满额、单价16,800元测算", "这是规划销售额，不是某一家企业的费用。"),
    ]
    for i, (accent, kicker, price, unit, body) in enumerate(cards):
        x = 0.4 + i * 4.28
        rrect(s, x, 1.22, 4.12, 4.15, PAPER, line=LINE, radius=0.08)
        rect(s, x, 1.22, 4.12, 0.08, accent)
        add_text(s, kicker, x + 0.22, 1.48, 3.7, 0.55, size=15, bold=True, color=accent)
        add_text(s, price, x + 0.2, 2.15, 3.7, 0.7, size=36, bold=True, color=INK)
        add_text(s, unit, x + 0.22, 2.9, 3.7, 0.55, size=14, color=MUTED)
        add_text(s, body, x + 0.22, 3.65, 3.7, 1.35, size=15, color=SLATE, spacing=1.15)
    add_text(
        s,
        "展位价格采用大会官方公开标准。20个展位的数量、位置及销售授权，仍须由组委会确认。两种方式可以单独选择。",
        0.45,
        5.55,
        12.4,
        0.7,
        size=15,
        color=INK,
    )
    notes(s, "336000不要念成对方要付的钱。双开口单价18480，需要时再报，不在主推价格里。")


def forum(prs):
    s = blank(prs)
    content_header(s, "方案一", "一场把银发科技和具身智能放到同一张桌上的论坛", 3, TOTAL, internal=False)
    add_text(
        s,
        "聚焦“十五五”银发经济、人工智能、具身智能、主动健康、智慧物联网与科技成果转化。规模100—120人，以产业资源交流和供需对接为主。",
        0.45,
        1.18,
        12.4,
        0.7,
        size=16,
        color=SLATE,
    )
    items = [
        ("学术交流", "拟邀复旦大学相关专家，说明老龄社会真正需要技术解决的问题。"),
        ("技术展示", "拟邀康复机器人、具身智能和智慧康养企业，展示已经能进场景的能力。"),
        ("产业对话", "养老运营、社区服务、科技企业同场讨论谁采购、谁运营、谁维护。"),
        ("项目合作", "会后保留需求、方案和联系人，方便继续对接，而不是只交换名片。"),
    ]
    for i, (title, body) in enumerate(items):
        col, row = i % 2, i // 2
        x = 0.42 + col * 6.45
        y = 2.05 + row * 1.85
        rrect(s, x, y, 6.25, 1.7, MIST if row == 0 else PAPER, line=LINE, radius=0.08)
        add_text(s, title, x + 0.28, y + 0.22, 5.7, 0.38, size=18, bold=True, color=TEAL)
        add_text(s, body, x + 0.28, y + 0.7, 5.7, 0.75, size=15, color=SLATE)
    add_text(
        s,
        "傅利叶智能、眸深智能、方舟健客、瑞健未来等均为邀约方向，书面确认前不视为已经参会。",
        0.45,
        5.9,
        12.4,
        0.45,
        size=14,
        color=MUTED,
    )


def tiers(prs):
    s = blank(prs)
    content_header(s, "方案一", "三档赞助，建议主推产业创新合作伙伴", 4, TOTAL, internal=False)
    table = add_table(
        s,
        0.55,
        1.28,
        12.2,
        2.55,
        [
            ["合作级别", "赞助金额", "建议名额", "这份方案里怎么用"],
            ["首席产业合作伙伴", "50,000元", "1家", "完整权益以大会确认的清单为准"],
            ["产业创新合作伙伴", "20,000元", "2—3家", "本方案主推，权益见下一页"],
            ["生态支持伙伴", "10,000元", "3—5家", "完整权益以大会确认的清单为准"],
        ],
        [3.3, 2.2, 1.8, 4.9],
        size=15,
    )
    for cell in table.table.rows[2].cells:
        set_cell_fill(cell, CARD_GOLD)
    add_text(
        s,
        "两万元这一档，适合希望同时获得品牌露出、一次专业分享申请和产业对接的企业。分享时长按10—15分钟申请，不是自动获得的演讲席位。",
        0.55,
        4.15,
        12.2,
        0.85,
        size=16,
        color=INK,
    )
    rrect(s, 0.55, 5.15, 12.2, 1.35, CARD_TEAL, radius=0.08)
    add_text(
        s,
        "大会另有公开的论坛演讲或产品推介标准：12,000元/人/20分钟。该价格与本方案的赞助权益分开计算。选择赞助，不自动包含这场官方演讲。",
        0.78,
        5.38,
        11.75,
        0.95,
        size=16,
        color=INK,
        anchor=MSO_ANCHOR.MIDDLE,
    )


def benefits(prs):
    s = blank(prs)
    content_header(s, "产业创新合作伙伴  ·  20,000元", "品牌、分享申请、对接和成果，四项写进权益", 5, TOTAL, internal=False)
    blocks = [
        ("01", "荣誉及品牌", "以“论坛产业创新合作伙伴”参与。名称和标识用于背景板、邀请函及指定宣传材料。主持人在合作伙伴介绍环节鸣谢。"),
        ("02", "论坛内容", "可推荐高管或技术负责人作为嘉宾候选。可提交与主题相关的案例，申请10—15分钟专业分享。须经内容审核和议程确认。"),
        ("03", "产业交流", "参加论坛组织的合作交流，对接具身智能、智慧康养、社区服务、养老运营和科技投资等相关资源。提供4个企业参会席位。"),
        ("04", "宣传与成果", "经审核的新闻稿体现企业实际参与情况。提供可授权使用的活动照片和宣传素材。产品与合作需求可纳入后续对接。"),
    ]
    for i, (n, title, body) in enumerate(blocks):
        col, row = i % 2, i // 2
        x = 0.42 + col * 6.45
        y = 1.2 + row * 2.35
        rrect(s, x, y, 6.28, 2.2, PAPER, line=LINE, radius=0.08)
        add_text(s, n, x + 0.24, y + 0.18, 1.2, 0.32, size=14, bold=True, color=GOLD)
        add_text(s, title, x + 0.24, y + 0.5, 5.8, 0.38, size=18, bold=True, color=INK)
        add_text(s, body, x + 0.24, y + 1.0, 5.8, 1.0, size=14, color=SLATE, spacing=1.08)
    notes(s, "不要口头增加央媒、电视、领导会见、采购成交或融资成功。审核不通过时，分享改为资料交流。")


def ruijian(prs):
    s = blank(prs)
    content_header(s, "重点邀约  ·  待企业确认", "瑞健未来：社区长寿科技与这次论坛同题", 6, TOTAL, internal=False)
    rrect(s, 0.42, 1.18, 8.15, 5.35, PAPER, line=LINE, radius=0.08)
    add_text(s, "拟邀企业", 0.68, 1.38, 3, 0.28, size=13, bold=True, color=TEAL)
    add_text(s, "瑞健未来", 0.68, 1.68, 7.5, 0.42, size=28, bold=True, color=INK)
    add_text(s, "拟邀代表    杨磊    首席运营官（待企业最终确认）", 0.68, 2.2, 7.6, 0.36, size=15, color=SLATE)
    add_text(s, "建议分享题目", 0.68, 2.75, 7.4, 0.28, size=13, bold=True, color=TEAL)
    add_text(
        s,
        "让长寿科技走进社区：\n从高端设备到家门口的健康服务新商业",
        0.68,
        3.1,
        7.5,
        0.85,
        size=20,
        bold=True,
        color=INK,
        spacing=1.05,
    )
    add_text(
        s,
        "建议结合社区门店、主动健康设备、适老化技术和服务运营，讲长寿科技怎样从产品走到社区市场，并讨论与养老机构、适老化居住项目的合作。",
        0.68,
        4.15,
        7.55,
        1.15,
        size=15,
        color=SLATE,
    )
    add_text(s, "发言须经内容审核，不默认占有演讲席位。", 0.68, 5.5, 7.5, 0.4, size=14, bold=True, color=ROSE)
    rrect(s, 8.75, 1.18, 4.15, 2.5, INK, radius=0.08)
    add_text(s, "建议合作", 9.0, 1.38, 3.7, 0.28, size=13, bold=True, color=GOLD)
    add_text(s, "¥20,000", 8.98, 1.78, 3.75, 0.6, size=32, bold=True, color=WHITE)
    add_text(s, "产业创新合作伙伴", 9.0, 2.5, 3.7, 0.7, size=16, color=CREAM)
    rrect(s, 8.75, 3.85, 4.15, 2.68, CARD_GOLD, radius=0.08)
    add_text(s, "另一种选择", 9.0, 4.05, 3.7, 0.28, size=13, bold=True, color=GOLD)
    add_text(s, "¥16,800", 8.98, 4.4, 3.75, 0.5, size=28, bold=True, color=INK)
    add_text(s, "9㎡标准展位\n11月27日—29日", 9.0, 5.05, 3.7, 0.7, size=15, color=SLATE)
    add_text(s, "两种方式可单独选择。", 9.0, 5.85, 3.7, 0.4, size=14, bold=True, color=INK)
    notes(s, "本页只在发给瑞健未来时保留。此前若已邀请分享且没有谈费用，赞助是可选合作，不是把演讲改成必须付费。")


def sponsor_close(prs):
    s = blank(prs)
    content_header(s, "方案一", "赞助合作按授权主体、对公账户和确认清单执行", 7, TOTAL, internal=False)
    steps = [
        ("01", "签约主体", "协议由具备授权和履约能力的主体签署、开票、收款。"),
        ("02", "付款时间", "资金按合同约定，在活动开始前付至指定对公账户。"),
        ("03", "权益清单", "品牌、分享、对接和宣传，以双方签署并经大会确认的清单为准。"),
        ("04", "内容边界", "品牌露出、舞台分享、产品宣传和场地展示，符合大会规范及法律法规。"),
    ]
    for i, (n, title, body) in enumerate(steps):
        y = 1.22 + i * 1.15
        rrect(s, 0.42, y, 12.5, 1.05, PAPER, line=LINE, radius=0.08)
        add_text(s, n, 0.62, y + 0.28, 0.7, 0.45, size=16, bold=True, color=TEAL, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, title, 1.5, y + 0.14, 3.2, 0.75, size=18, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, body, 4.8, y + 0.18, 7.8, 0.7, size=16, color=SLATE, anchor=MSO_ANCHOR.MIDDLE)
    notes(s, "联系人尚未填入。发出前补上授权联系人，不要留空发送。")


def zone(prs):
    s = blank(prs)
    content_header(s, "方案二", "银发科技与AI智慧康养创新成果展示区", 8, TOTAL, internal=False)
    facts = [
        ("展会", "2026上海国际\n智慧康养产业博览会"),
        ("展期", "11月27日—29日\n共三天"),
        ("地点", "上海展览中心\n位置以平面图为准"),
        ("规模", "拟定20个标准展位\n合计180㎡净面积"),
    ]
    for i, (k, v) in enumerate(facts):
        x = 0.42 + i * 3.22
        rrect(s, x, 1.2, 3.08, 1.85, MIST, radius=0.08)
        add_text(s, k, x + 0.16, 1.34, 2.75, 0.28, size=13, bold=True, color=TEAL)
        add_text(s, v, x + 0.16, 1.7, 2.75, 1.1, size=16, bold=True, color=INK, spacing=1.08)
    add_text(s, "展区希望集中展示的方向", 0.45, 3.25, 8, 0.32, size=14, bold=True, color=TEAL)
    tags = [
        "具身智能与养老机器人",
        "智慧健康与智能康复",
        "主动健康与数字医疗",
        "社区长寿科技",
        "智慧物联网",
        "适老产品",
    ]
    for i, tag in enumerate(tags):
        col, row = i % 3, i // 3
        x = 0.42 + col * 4.28
        y = 3.7 + row * 0.85
        rrect(s, x, y, 4.1, 0.7, PAPER, line=LINE, radius=0.1)
        add_text(s, tag, x + 0.2, y + 0.16, 3.7, 0.4, size=16, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)
    add_text(
        s,
        "20个展位是否可售、具体位置和布展权限，以组委会确认的展馆平面图和销售授权为准。180㎡不含公共通道和疏散空间。",
        0.45,
        5.55,
        12.4,
        0.7,
        size=15,
        color=MUTED,
    )


def booth_price(prs):
    s = blank(prs)
    content_header(s, "方案二", "标准展位9平方米，公开报价16,800元", 9, TOTAL, internal=False)
    rrect(s, 0.42, 1.18, 6.3, 3.55, INK, radius=0.08)
    add_text(s, "国内标准展位", 0.68, 1.4, 5.8, 0.3, size=14, bold=True, color=GOLD)
    add_text(s, "¥16,800", 0.64, 1.85, 5.8, 0.7, size=40, bold=True, color=WHITE)
    add_text(s, "3米 × 3米  ·  9平方米  ·  三天展期", 0.68, 2.65, 5.8, 0.35, size=16, color=CREAM)
    add_text(s, "双开口按公开标准加收10%\n18,480元 / 个", 0.68, 3.2, 5.8, 0.7, size=16, color=WHITE, spacing=1.1)
    add_text(s, "特殊用电、运输、特装、网络另行确认", 0.68, 4.1, 5.8, 0.35, size=14, color=CREAM)
    configs = ["标准围板及企业楣板", "产品展示空间", "基础桌椅", "基础照明与常规用电接口"]
    for i, item in enumerate(configs):
        y = 1.18 + i * 0.72
        rrect(s, 6.95, y, 5.95, 0.64, MIST, radius=0.08)
        add_text(s, item, 7.2, y + 0.12, 5.5, 0.4, size=16, color=INK, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, "以上配置以主办方正式展位清单为准。", 6.95, 4.15, 5.95, 0.4, size=13, color=MUTED)
    rrect(s, 0.42, 4.95, 12.48, 1.5, CARD_GOLD, radius=0.08)
    add_text(
        s,
        "展位费不包含论坛主旨演讲、圆桌席位、赞助冠名或专属媒体报道。若另走大会论坛演讲或产品推介，公开标准为12,000元/人/20分钟，与展位费分开结算。",
        0.65,
        5.2,
        12.05,
        1.05,
        size=16,
        color=INK,
        anchor=MSO_ANCHOR.MIDDLE,
    )


def layout(prs):
    s = blank(prs)
    content_header(s, "方案二", "六个方向，二十个展位，企业都待确认", 10, TOTAL, internal=False)
    add_table(
        s,
        0.42,
        1.2,
        12.5,
        4.55,
        [
            ["展示专区", "展位", "邀约方向"],
            ["具身智能与养老机器人", "4", "傅利叶、眸深、智元及机器人产业链"],
            ["AI与智慧物联网", "4", "智能感知、居家监护、智慧家庭"],
            ["主动健康与数字医疗", "4", "方舟健客、数字医疗、智能康复、健康管理"],
            ["社区长寿科技与适老化", "4", "瑞健未来、康养设备、社区健康服务"],
            ["银发金融及养老服务", "2", "银行、保险、养老运营机构"],
            ["创新成果与科技转化", "2", "高校科研成果及创新企业"],
        ],
        [4.3, 1.3, 6.9],
        size=14,
    )
    add_text(s, "上表是招商方向，不是已确定的参展名单。合计20个。", 0.45, 5.95, 12.4, 0.4, size=15, color=MUTED)


def rights(prs):
    s = blank(prs)
    content_header(s, "方案二", "参展权益围绕展示和对接，不附带论坛席位", 11, TOTAL, internal=False)
    rights_list = [
        ("展位", "一个经确认的9平方米标准展位，使用期为11月27日至29日。"),
        ("展示", "在规定范围内展示品牌、设备、产品和解决方案。"),
        ("名录", "企业信息可纳入经组委会批准的展商资料或展示信息。"),
        ("论坛", "可参加开放的同期产业论坛和商务交流，不因此获得演讲席位。"),
        ("对接", "可申请供需对接和产业交流。"),
        ("服务", "获得必要的参展服务和展馆管理信息。"),
    ]
    for i, (title, body) in enumerate(rights_list):
        col, row = i % 2, i // 2
        x = 0.42 + col * 6.45
        y = 1.18 + row * 1.45
        rrect(s, x, y, 6.28, 1.32, PAPER, line=LINE, radius=0.08)
        add_text(s, title, x + 0.24, y + 0.16, 5.8, 0.34, size=16, bold=True, color=TEAL)
        add_text(s, body, x + 0.24, y + 0.56, 5.8, 0.58, size=14, color=SLATE)
    add_text(
        s,
        "医疗器械、氧疗设备和机器人动态演示，参展前须完成资质、宣传合规和现场安全审核。",
        0.45,
        5.65,
        12.4,
        0.55,
        size=15,
        color=ROSE,
    )


def close(prs):
    s = blank(prs)
    content_header(s, "下一步", "选定一种方式后，按确认过的合同执行", 12, TOTAL, internal=False)
    choices = [
        ("选赞助", "产业创新合作伙伴", "20,000元 / 家", "品牌、分享申请、4个席位、对接与成果素材。"),
        ("选展位", "9㎡标准展位", "16,800元 / 个", "三天展示。双开口18,480元。不含演讲和冠名。"),
    ]
    for i, (k, title, price, body) in enumerate(choices):
        x = 0.42 + i * 6.45
        rrect(s, x, 1.2, 6.25, 2.55, INK if i == 0 else PAPER, line=None if i == 0 else LINE, radius=0.08)
        title_c = GOLD if i == 0 else TEAL
        body_c = CREAM if i == 0 else SLATE
        price_c = WHITE if i == 0 else INK
        add_text(s, k, x + 0.28, 1.38, 5.7, 0.3, size=14, bold=True, color=title_c)
        add_text(s, title, x + 0.28, 1.75, 5.7, 0.4, size=20, bold=True, color=price_c)
        add_text(s, price, x + 0.28, 2.25, 5.7, 0.45, size=22, bold=True, color=title_c)
        add_text(s, body, x + 0.28, 2.85, 5.7, 0.65, size=15, color=body_c)
    add_text(s, "授权联系人  ·  发出前填写", 0.48, 4.0, 6, 0.32, size=14, bold=True, color=TEAL)
    for i, label in enumerate(["单位", "联系人", "电话", "邮箱"]):
        x = 0.42 + (i % 2) * 6.45
        yy = 4.45 + (i // 2) * 0.85
        rrect(s, x, yy, 6.25, 0.72, MIST, radius=0.08)
        add_text(s, label, x + 0.22, yy + 0.16, 1.3, 0.4, size=14, bold=True, color=MUTED, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, "待填写", x + 1.6, yy + 0.16, 4.3, 0.4, size=16, color=INK, anchor=MSO_ANCHOR.MIDDLE)
    notes(s, "联系人留空时不要外发。不要把大会秘书处电话印成我们的招商电话。")


def build():
    prs = new_prs()
    cover(prs)
    prices(prs)
    forum(prs)
    tiers(prs)
    benefits(prs)
    ruijian(prs)
    sponsor_close(prs)
    zone(prs)
    booth_price(prs)
    layout(prs)
    rights(prs)
    close(prs)
    if len(prs.slides) != TOTAL:
        raise RuntimeError(f"页数{len(prs.slides)}与TOTAL {TOTAL}不一致")
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    prs.save(OUT)
    return OUT


if __name__ == "__main__":
    print(build())
