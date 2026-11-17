# -*- coding: utf-8 -*-
"""生成 16:9 合作洽谈演示。学院讨论用，不作为招生简章。"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from coop_data import AI_PLAN, EXISTING_AIGC, META

NAVY = RGBColor(0x13, 0x29, 0x4B)
NAVY2 = RGBColor(0x1C, 0x3A, 0x64)
CRIMSON = RGBColor(0x8C, 0x1D, 0x2C)
GOLD = RGBColor(0xB0, 0x8A, 0x45)
INK = RGBColor(0x1C, 0x24, 0x30)
MUTED = RGBColor(0x5E, 0x6A, 0x78)
PAPER = RGBColor(0xF6, 0xF3, 0xEE)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT = RGBColor(0xF8, 0xF1, 0xF2)
LINE = RGBColor(0xE4, 0xDD, 0xD2)
CREAM = RGBColor(0xFB, 0xF8, 0xF3)

FONT = "微软雅黑"
SW = Inches(13.333)
SH = Inches(7.5)

prs = Presentation()
prs.slide_width = SW
prs.slide_height = SH
prs.core_properties.title = "复旦终身教育学院与杨浦科联 2026—2027年合作洽谈"
prs.core_properties.author = META["compiler"]
prs.core_properties.subject = META["nature"]
BLANK = prs.slide_layouts[6]


def _font(run, size, color, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.append(el)
        el.set("typeface", FONT)


def _shape(slide, kind, x, y, w, h, fill, line=None):
    sh = slide.shapes.add_shape(kind, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1)
    sh.shadow.inherit = False
    return sh


def rect(slide, x, y, w, h, fill):
    return _shape(slide, MSO_SHAPE.RECTANGLE, x, y, w, h, fill)


def rrect(slide, x, y, w, h, fill, radius=0.08):
    sh = _shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill)
    try:
        sh.adjustments[0] = radius
    except Exception:
        pass
    return sh


def tb(slide, x, y, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.vertical_anchor = anchor
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    if isinstance(lines, tuple):
        lines = [lines]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        runs = line if isinstance(line, list) else [line]
        p.alignment = align
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        for run_spec in runs:
            text, size, color, bold = run_spec[:4]
            gap = run_spec[4] if len(run_spec) > 4 else 0
            r = p.add_run()
            r.text = text
            _font(r, size, color, bold)
            if gap:
                p.space_after = Pt(gap)
        if p.line_spacing is None:
            p.line_spacing = 1.0
    return box


def new(bg=PAPER):
    s = prs.slides.add_slide(BLANK)
    fill = s.background.fill
    fill.solid()
    fill.fore_color.rgb = bg
    return s


def light(kicker, title):
    s = new(PAPER)
    rect(s, 0, 0, Inches(0.10), SH, CRIMSON)
    tb(s, Inches(0.48), Inches(0.26), Inches(10), Inches(0.28),
       [(kicker, 12, CRIMSON, True)])
    tb(s, Inches(0.46), Inches(0.50), Inches(12.2), Inches(0.48),
       [(title, 26, NAVY, True)])
    rect(s, Inches(0.48), Inches(1.08), Inches(1.15), Inches(0.045), GOLD)
    return s


def foot(s, page, dark=False):
    color = RGBColor(0xC5, 0xCD, 0xD6) if dark else MUTED
    accent = GOLD if dark else CRIMSON
    tb(s, Inches(0.48), Inches(7.12), Inches(9.6), Inches(0.26),
       [(f"{META['party_a']}  ×  {META['party_b']}    {META['version']}", 10, color, False)])
    tb(s, Inches(11.6), Inches(7.12), Inches(1.25), Inches(0.26),
       [(f"{page:02d}", 11, accent, True)], align=PP_ALIGN.RIGHT)


def card_title(s, x, y, w, label, title, body, label_color=CRIMSON):
    tb(s, x, y, w, Inches(0.24), [(label, 11, label_color, True)])
    tb(s, x, y + Inches(0.26), w, Inches(0.36), [(title, 16, NAVY, True)])
    tb(s, x, y + Inches(0.66), w, Inches(1.15), [(body, 12, MUTED, False)])


# ---------- slides ----------

def s01():
    s = new(NAVY)
    rect(s, 0, 0, SW, Inches(0.12), CRIMSON)
    rect(s, 0, Inches(7.38), SW, Inches(0.12), CRIMSON)
    tb(s, Inches(0.72), Inches(0.85), Inches(11), Inches(0.32),
       [("合作洽谈稿    2026年11月", 14, GOLD, True)])
    tb(s, Inches(0.70), Inches(1.45), Inches(12), Inches(1.7), [
        [("2026—2027年", 40, WHITE, True)],
        [("战略合作框架建议", 40, WHITE, True)],
    ])
    rect(s, Inches(0.74), Inches(3.35), Inches(1.4), Inches(0.05), GOLD)
    tb(s, Inches(0.72), Inches(3.58), Inches(11.5), Inches(0.7), [
        [(META["party_a"], 18, WHITE, True)],
        [("×   " + META["party_b"], 18, WHITE, True)],
    ])
    tb(s, Inches(0.72), Inches(4.55), Inches(11.2), Inches(0.4),
       [("高校教育资源  +  产业组织平台  +  市场化运营", 16, RGBColor(0xD5, 0xDE, 0xEA), False)])
    chips = ["委托培训打底", "产业参访立差异", "科创短训做增长"]
    for i, t in enumerate(chips):
        x = Inches(0.72) + Inches(i * 3.15)
        rrect(s, x, Inches(5.45), Inches(2.95), Inches(0.48), NAVY2, 0.2)
        tb(s, x, Inches(5.54), Inches(2.95), Inches(0.32),
           [(t, 13, WHITE, True)], align=PP_ALIGN.CENTER)
    tb(s, Inches(0.72), Inches(6.35), Inches(11), Inches(0.55), [
        [("供学院内部讨论。产品、价格、证书和分成均未批准。", 13, RGBColor(0xC5, 0xCD, 0xD6), False)],
        [("正式合作以学校审批、合作备忘录和项目合同为准。", 13, RGBColor(0xC5, 0xCD, 0xD6), False)],
    ])


def s02():
    s = light("阅读说明", "先分清：纪要、建议、尚未核准")
    items = [
        (CRIMSON, "来自洽谈", "合作方向、客户线索和参访菜单，来自11月6日、11月10日三份纪要。学院要市场转化能力，并已提出先整理参访资源。"),
        (NAVY, "本次建议", "六项业务、三类模式、产品方案、价格测试、排期和评价指标，是联合会整理的实施建议，双方尚未确认。"),
        (GOLD, "不能写成已批准", "内部管理费、税务口径、票据路径、证书和校友权益，都还没有正式核准。正式文件只写真实服务和校方审批。"),
    ]
    for i, (color, title, body) in enumerate(items):
        x = Inches(0.48) + Inches(i * 4.2)
        rrect(s, x, Inches(1.5), Inches(4.0), Inches(3.35), WHITE, 0.06)
        rect(s, x, Inches(1.5), Inches(4.0), Inches(0.08), color)
        tb(s, x + Inches(0.24), Inches(1.8), Inches(3.5), Inches(0.4),
           [(title, 18, NAVY, True)])
        tb(s, x + Inches(0.24), Inches(2.4), Inches(3.5), Inches(2.1),
           [(body, 14, INK, False)])
    files = [
        ("Word", "完整框架和四份附件，可会后细读"),
        ("PPT", "这次沟通只用决策信息"),
        ("Excel", "可筛选台账。金色格是假设，内部跟进表不要外发"),
    ]
    for i, (k, v) in enumerate(files):
        x = Inches(0.48) + Inches(i * 4.2)
        tb(s, x, Inches(5.1), Inches(4.0), Inches(0.3), [(k, 13, CRIMSON, True)])
        tb(s, x, Inches(5.42), Inches(4.0), Inches(0.7), [(v, 13, MUTED, False)])
    foot(s, 2)


def s03():
    s = light("合作方向", "三次洽谈，收成一件事")
    tb(s, Inches(0.48), Inches(1.35), Inches(12.3), Inches(0.85),
       [("以学院的教育、师资和培训交付为核心，以联合会的企业客户、市场渠道和活动运营为支撑，共同做面向科技企业、央国企、政府园区和企业家的非学历教育与产业交流。", 16, INK, False)])
    cards = [
        ("11月6日", "终身教育业务", "学院的 B 端委托培训比较成熟，C 端招生和新产品仍在探索。联合会要补上的是开拓、转化和产业组织，而不是只递几家企业名单。"),
        ("11月6日", "教育圈资源对接", "作为三份背景材料之一。没有在纪要里单独核准的机构、证书和收益安排，不带入本次合作范围。"),
        ("11月10日", "科创培训项目", "学院要求整理参访资源菜单，并表示符合价格体系后可以委托。先形成备忘录和基础框架，2027年再逐步推进项目。"),
    ]
    for i, (date, title, body) in enumerate(cards):
        y = Inches(2.4) + Inches(i * 1.45)
        rrect(s, Inches(0.48), y, Inches(12.35), Inches(1.32), WHITE, 0.06)
        rect(s, Inches(0.48), y, Inches(0.08), Inches(1.32), CRIMSON if i == 2 else NAVY)
        tb(s, Inches(0.8), y + Inches(0.16), Inches(2.2), Inches(0.95), [
            [(date, 12, CRIMSON, True)],
            [(title, 16, NAVY, True)],
        ])
        tb(s, Inches(3.2), y + Inches(0.22), Inches(9.3), Inches(0.95),
           [(body, 14, INK, False)], anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 3)


def s04():
    s = light("总结构", "一个定位，两方分工，六项业务")
    blocks = [
        ("01", "一个定位", "高校资源\n产业平台\n市场运营"),
        ("02", "两大主体", "学院管教学\n联合会管\n市场与产业"),
        ("03", "六项业务", "培训、人才\n参访、招生\n咨询、特色"),
        ("04", "三类模式", "直接委托\n需求整合\n联合开发"),
        ("05", "一套机制", "双负责人\n月度会\n品牌授权"),
    ]
    for i, (n, t, b) in enumerate(blocks):
        x = Inches(0.42) + Inches(i * 2.56)
        rrect(s, x, Inches(1.55), Inches(2.4), Inches(3.15), WHITE, 0.08)
        rect(s, x, Inches(1.55), Inches(2.4), Inches(0.08), GOLD if i % 2 else CRIMSON)
        tb(s, x + Inches(0.16), Inches(1.8), Inches(2.1), Inches(0.4), [(n, 18, CRIMSON, True)])
        tb(s, x + Inches(0.16), Inches(2.25), Inches(2.1), Inches(0.4), [(t, 16, NAVY, True)])
        tb(s, x + Inches(0.16), Inches(2.8), Inches(2.1), Inches(1.5),
           [(line, 14, MUTED, False, 4) for line in b.split("\n")])
    tb(s, Inches(0.5), Inches(5.0), Inches(12.2), Inches(0.9), [
        [("主线", 13, CRIMSON, True)],
        [("委托培训形成业务基础，产业参访形成差异，科创短训形成增长，长期靠资源共享把合作留住。", 16, NAVY, True)],
    ])
    tb(s, Inches(0.5), Inches(6.15), Inches(12.2), Inches(0.7),
       [("具体项目仍以学校审批为准。国学研究院等机构化设想单独立题，不写入这次已经确定的事项。", 13, MUTED, False)])
    foot(s, 4)


def s05():
    s = light("目标", "年底定边界，明年做试点")
    goals = [
        ("2026年底", "近期", "建立对接机制，签署合作备忘录，完成参访资源库和首批课程方案。"),
        ("2027年", "中期", "企业委托培训、人工智能研修班和产业参访按审批进度落地。"),
        ("2027年以后", "长期", "做出能持续招生的科创人才培养机制，再扩展到长三角、园区和政府委托。"),
    ]
    for i, (when, tag, body) in enumerate(goals):
        x = Inches(0.48) + Inches(i * 4.2)
        rrect(s, x, Inches(1.5), Inches(4.0), Inches(3.15), WHITE, 0.06)
        tb(s, x + Inches(0.24), Inches(1.7), Inches(3.5), Inches(0.28), [(tag, 12, CRIMSON, True)])
        tb(s, x + Inches(0.24), Inches(2.05), Inches(3.5), Inches(0.5), [(when, 22, NAVY, True)])
        tb(s, x + Inches(0.24), Inches(2.7), Inches(3.5), Inches(1.6), [(body, 14, INK, False)])
    tb(s, Inches(0.5), Inches(4.95), Inches(12.2), Inches(0.35),
       [("合作原则", 13, CRIMSON, True)])
    principles = ["优势互补", "专业分工", "项目化运作", "依法合规", "市场导向", "成果共享"]
    for i, p in enumerate(principles):
        x = Inches(0.48) + Inches(i * 2.12)
        rrect(s, x, Inches(5.45), Inches(2.0), Inches(0.48), NAVY, 0.2)
        tb(s, x, Inches(5.55), Inches(2.0), Inches(0.3),
           [(p, 13, WHITE, True)], align=PP_ALIGN.CENTER)
    tb(s, Inches(0.5), Inches(6.15), Inches(12.2), Inches(0.6),
       [("以上时间为建议排期。立项、招生和学校审批的实际进度，优先于这张时间表。", 13, MUTED, False)])
    foot(s, 5)


def s06():
    s = light("主体", "学院主导教学，联合会负责市场和产业")
    cols = [
        ("甲方  ·  学院（拟）", [
            "教育项目论证、立项和校内审批",
            "课程、教学质量、师资组织",
            "教学资源、场地和学员服务",
            "经批准的考核与证书",
            "政府、国央企项目的教学交付",
            "教育品牌和教学标准",
        ]),
        ("乙方  ·  联合会（拟）", [
            "科技企业、金融机构、园区和政府客户",
            "培训需求、市场调研和项目引荐",
            "社群、论坛和公开课的市场组织",
            "参访资源、商务协调和活动服务",
            "产业专家、企业家和应用场景",
            "宣传、客户维护和经批准的配套服务",
        ]),
    ]
    for i, (title, lines) in enumerate(cols):
        x = Inches(0.48) + Inches(i * 6.4)
        rrect(s, x, Inches(1.45), Inches(6.15), Inches(4.35), WHITE, 0.05)
        rect(s, x, Inches(1.45), Inches(6.15), Inches(0.5), NAVY if i == 0 else CRIMSON)
        tb(s, x + Inches(0.24), Inches(1.54), Inches(5.7), Inches(0.36),
           [(title, 15, WHITE, True)])
        y = 2.15
        for line in lines:
            tb(s, x + Inches(0.28), Inches(y), Inches(5.6), Inches(0.48),
               [("·  " + line, 14, INK, False)])
            y += 0.55
    tb(s, Inches(0.5), Inches(5.95), Inches(12.3), Inches(0.95),
       [("可再引入课程策划、招生、班务或参访机构。赵海等人员只作为候选合作人员，不当然取得校方岗位或授权。院内对接由刘老师协调，授权以正式指定为准；联合会由执行会长牵头，协调朱总及相关人员。", 13, MUTED, False)])
    foot(s, 6)


def s07():
    s = light("业务", "六项业务，只先做最高优先级里的三件")
    items = [
        ("01", "委托培训", "最高", "央国企、金融机构、科技企业的定制课程"),
        ("02", "科创人才", "最高", "企业家 AI 战略、治理、投融资短训"),
        ("03", "产业参访", "最先启动", "标杆企业菜单和长三角现场教学"),
        ("04", "市场招生", "高", "公开课、论坛，再导入正式报名"),
        ("05", "智库咨询", "中高", "微诊断、课题和产业研究"),
        ("06", "特色项目", "中期", "银发经济、国学、国际交流"),
    ]
    for i, (n, title, tag, body) in enumerate(items):
        col, row = i % 3, i // 3
        x = Inches(0.45) + Inches(col * 4.25)
        y = Inches(1.42) + Inches(row * 2.55)
        rrect(s, x, y, Inches(4.08), Inches(2.38), WHITE, 0.07)
        tb(s, x + Inches(0.22), y + Inches(0.18), Inches(1.2), Inches(0.36), [(n, 18, CRIMSON, True)])
        rrect(s, x + Inches(2.15), y + Inches(0.22), Inches(1.7), Inches(0.32), SOFT, 0.3)
        tb(s, x + Inches(2.15), y + Inches(0.24), Inches(1.7), Inches(0.28),
           [(tag, 11, CRIMSON, True)], align=PP_ALIGN.CENTER)
        tb(s, x + Inches(0.22), y + Inches(0.7), Inches(3.6), Inches(0.4), [(title, 18, NAVY, True)])
        tb(s, x + Inches(0.22), y + Inches(1.25), Inches(3.6), Inches(0.85), [(body, 14, MUTED, False)])
    foot(s, 7)


def s08():
    s = light("试点", "资源先集中在三个项目")
    items = [
        ("项目一", "科技企业标杆参访", "联合会整理 10—20 家候选机构，核实接待、费用和预约条件。学院用现有委托班次试用。"),
        ("项目二", "央国企委托培训", "优先约国泰海通、国药控股，或已有联系的金融机构。学院按真实需求出课程建议。"),
        ("项目三", "AI 领军人才研修班", "学院牵头课程、师资和预算。联合会先测试企业家是否愿意报名。立项和招生审批完成后再开班。"),
    ]
    for i, (k, t, b) in enumerate(items):
        y = Inches(1.38) + Inches(i * 1.35)
        rrect(s, Inches(0.48), y, Inches(8.35), Inches(1.22), WHITE, 0.06)
        tb(s, Inches(0.7), y + Inches(0.14), Inches(7.9), Inches(0.28), [(k, 12, CRIMSON, True)])
        tb(s, Inches(0.7), y + Inches(0.40), Inches(7.9), Inches(0.3), [(t, 16, NAVY, True)])
        tb(s, Inches(0.7), y + Inches(0.74), Inches(7.9), Inches(0.38), [(b, 12, MUTED, False)])
    rrect(s, Inches(9.05), Inches(1.38), Inches(3.8), Inches(4.0), NAVY, 0.06)
    tb(s, Inches(9.28), Inches(1.58), Inches(3.4), Inches(0.4), [("推进顺序", 14, GOLD, True)])
    order = [
        ("1", "先有订单", "三家培训负责人引荐给刘老师"),
        ("2", "先交菜单", "参访条件核实到可以报价"),
        ("3", "先验证", "访谈决策者，再决定是否招生"),
    ]
    for i, (n, t, b) in enumerate(order):
        y = 2.15 + i * 1.0
        tb(s, Inches(9.28), Inches(y), Inches(3.4), Inches(0.28), [(n + "   " + t, 15, WHITE, True)])
        tb(s, Inches(9.28), Inches(y + 0.32), Inches(3.4), Inches(0.45), [(b, 12, RGBColor(0xD5, 0xDE, 0xEA), False)])
    foot(s, 8)


def s09():
    s = light("客户", "第一轮只引荐三家培训负责人")
    cards = [
        ("P0", "国泰海通证券", "培训部门已有直接联系。先问 2027 年培训计划、科技金融和 AI 需求。"),
        ("P0", "国药控股", "已提到负责培训的人员。先了解医药经营、组织发展和数字化转型。"),
        ("P0", "中远相关企业", "先核实是集团、下属公司还是培训机构，再谈专题。主体未核前不定项目。"),
    ]
    for i, (tag, name, body) in enumerate(cards):
        x = Inches(0.45) + Inches(i * 4.25)
        rrect(s, x, Inches(1.45), Inches(4.08), Inches(3.15), WHITE, 0.06)
        rect(s, x, Inches(1.45), Inches(0.08), Inches(3.15), CRIMSON)
        tb(s, x + Inches(0.28), Inches(1.65), Inches(3.5), Inches(0.3), [(tag, 13, CRIMSON, True)])
        tb(s, x + Inches(0.28), Inches(2.05), Inches(3.5), Inches(0.7), [(name, 20, NAVY, True)])
        tb(s, x + Inches(0.28), Inches(2.9), Inches(3.5), Inches(1.4), [(body, 14, INK, False)])
    tb(s, Inches(0.5), Inches(4.85), Inches(12.3), Inches(0.35),
       [("第二梯队，核实后再排期", 13, CRIMSON, True)])
    second = ["招商银行上海分行", "交通银行相关单位", "城商行", "保险机构", "中国宝武"]
    for i, name in enumerate(second):
        x = Inches(0.48) + Inches(i * 2.52)
        rrect(s, x, Inches(5.35), Inches(2.4), Inches(0.48), WHITE, 0.2)
        tb(s, x, Inches(5.45), Inches(2.4), Inches(0.3),
           [(name, 12, NAVY, True)], align=PP_ALIGN.CENTER)
    tb(s, Inches(0.5), Inches(6.05), Inches(12.3), Inches(0.75),
       [("外发清单不写未经确认的负责人姓名、电话和预算。三家里有一家形成真实需求，备忘录就有了具体落脚点。", 13, MUTED, False)])
    foot(s, 9)


def s10():
    s = light("方向一", "委托培训：按企业真实需求设计")
    themes = [
        "生成式人工智能与企业数字化转型",
        "企业战略、公司治理与组织创新",
        "科技成果转化、技术商业化与资本运作",
        "产业投资、并购重组与融资实务",
        "先进制造、具身智能与新质生产力",
        "国央企科技创新与产业转型",
    ]
    for i, t in enumerate(themes):
        col, row = i % 2, i // 2
        x = Inches(0.48) + Inches(col * 6.4)
        y = Inches(1.45) + Inches(row * 1.15)
        rrect(s, x, y, Inches(6.15), Inches(1.0), WHITE, 0.08)
        tb(s, x + Inches(0.22), y + Inches(0.28), Inches(0.7), Inches(0.4),
           [(f"{i + 1:02d}", 16, CRIMSON, True)])
        tb(s, x + Inches(1.0), y + Inches(0.28), Inches(4.9), Inches(0.45),
           [(t, 15, NAVY, True)], anchor=MSO_ANCHOR.MIDDLE)
    tb(s, Inches(0.5), Inches(5.05), Inches(12.3), Inches(1.5), [
        [("产品形态    3—5 天高管班、专题研修、企业内训。", 15, INK, False, 6)],
        [("合作方式    联合会对接培训部门并带回需求，学院设计课程、组织师资、实施培训，按各自承担的服务签项目协议。", 15, INK, False)],
    ])
    foot(s, 10)


def s11():
    s = light("方向三", "参访菜单是最近可以交付的成果")
    stats = [("37", "条候选资源"), ("12", "个优先核实"), ("5", "条教学线路"), ("3", "档服务标准")]
    for i, (n, label) in enumerate(stats):
        x = Inches(0.48) + Inches(i * 3.2)
        rrect(s, x, Inches(1.42), Inches(3.02), Inches(1.25), WHITE, 0.08)
        tb(s, x, Inches(1.52), Inches(3.02), Inches(0.55),
           [(n, 26, CRIMSON, True)], align=PP_ALIGN.CENTER)
        tb(s, x, Inches(2.12), Inches(3.02), Inches(0.35),
           [(label, 13, MUTED, False)], align=PP_ALIGN.CENTER)
    groups = [
        ("A 类", "腾讯、阿里、抖音、美团、B站、小红书、叠纸、华为。有历史对接经验，窗口要重新确认。"),
        ("B 类", "百度、张江机器人谷、张江药谷、宝武、江南造船、国泰海通、银行和康养机构。重点拟对接。"),
        ("C 类", "张江 AI 应用商店、人形机器人展示中心、基因岛等公开展示点。团体教学接待另行确认。"),
    ]
    for i, (t, b) in enumerate(groups):
        x = Inches(0.48) + Inches(i * 4.2)
        tb(s, x, Inches(2.95), Inches(4.0), Inches(0.32), [(t, 14, CRIMSON, True)])
        tb(s, x, Inches(3.35), Inches(4.0), Inches(1.35), [(b, 13, INK, False)])
    rrect(s, Inches(0.48), Inches(4.9), Inches(12.35), Inches(1.85), WHITE, 0.05)
    tb(s, Inches(0.72), Inches(5.08), Inches(11.9), Inches(0.3),
       [("第一批先核实", 13, CRIMSON, True)])
    tb(s, Inches(0.72), Inches(5.45), Inches(11.9), Inches(0.45),
       [("腾讯、美团、B站、张江 AI 应用商店、张江机器人谷、张江药谷。", 16, NAVY, True)])
    tb(s, Inches(0.72), Inches(6.0), Inches(11.9), Inches(0.5),
       [("每家做成资料卡：概况、主题、对象、人数、时长、讲解、费用、开放条件和预约流程。清单不是接待承诺函。", 13, MUTED, False)])
    foot(s, 11)


def s12():
    s = light("现场教学", "五条线路，三种时长")
    routes = [
        ("01", "人工智能与数字化", "标杆企业 + AI 场景 + 专题研讨"),
        ("02", "机器人与未来产业", "机器人谷 + 代表企业或展示平台"),
        ("03", "生物医药与医疗科技", "张江药谷 + 创新平台 + 企业"),
        ("04", "先进制造与新能源", "宝武或新能源场景 + 案例研讨"),
        ("05", "科技金融与产业投资", "金融机构 + 科技企业 + 课堂"),
    ]
    for i, (n, t, b) in enumerate(routes):
        y = Inches(1.38) + Inches(i * 0.78)
        rrect(s, Inches(0.45), y, Inches(8.3), Inches(0.7), WHITE, 0.1)
        tb(s, Inches(0.65), y + Inches(0.16), Inches(0.6), Inches(0.38), [(n, 14, CRIMSON, True)])
        tb(s, Inches(1.35), y + Inches(0.08), Inches(3.3), Inches(0.52), [(t, 14, NAVY, True)], anchor=MSO_ANCHOR.MIDDLE)
        tb(s, Inches(4.7), y + Inches(0.08), Inches(3.8), Inches(0.52), [(b, 13, MUTED, False)], anchor=MSO_ANCHOR.MIDDLE)
    tiers = [
        ("基础", "60—90 分钟", "介绍、开放区域、问答"),
        ("专题", "90—150 分钟", "参观、案例、专题交流"),
        ("深度", "半天至一天", "授课、参访、管理层交流"),
    ]
    for i, (t, time, b) in enumerate(tiers):
        y = Inches(1.38) + Inches(i * 1.45)
        rrect(s, Inches(8.95), y, Inches(3.9), Inches(1.32), NAVY if i == 1 else WHITE, 0.08)
        title_c = WHITE if i == 1 else NAVY
        body_c = RGBColor(0xD5, 0xDE, 0xEA) if i == 1 else MUTED
        tb(s, Inches(9.15), y + Inches(0.14), Inches(3.5), Inches(0.3), [(t, 14, GOLD if i == 1 else CRIMSON, True)])
        tb(s, Inches(9.15), y + Inches(0.46), Inches(3.5), Inches(0.28), [(time, 14, title_c, True)])
        tb(s, Inches(9.15), y + Inches(0.8), Inches(3.5), Inches(0.35), [(b, 12, body_c, False)])
    foot(s, 12)


def s13():
    s = light("方向二", "新班补战略和落地，不重复工具课")
    rrect(s, Inches(0.45), Inches(1.42), Inches(6.05), Inches(4.35), WHITE, 0.05)
    rrect(s, Inches(6.75), Inches(1.42), Inches(6.1), Inches(4.35), NAVY, 0.05)
    tb(s, Inches(0.7), Inches(1.62), Inches(5.6), Inches(0.3), [("学院已有产品", 13, CRIMSON, True)])
    tb(s, Inches(0.7), Inches(2.0), Inches(5.6), Inches(0.7),
       [("AIGC 应用技能研修班", 20, NAVY, True)])
    left = [
        EXISTING_AIGC["term"],
        EXISTING_AIGC["price_note"] + "，每期 40 席",
        "三个教学日，邯郸校区",
        "工具、商业场景、品牌与 AI 领导力",
        "项目编号 " + EXISTING_AIGC["codes"],
    ]
    for i, line in enumerate(left):
        tb(s, Inches(0.7), Inches(2.85) + Inches(i * 0.42), Inches(5.5), Inches(0.38),
           [(line, 14, INK, False)])
    tb(s, Inches(7.0), Inches(1.62), Inches(5.6), Inches(0.3), [("建议的新位置", 13, GOLD, True)])
    tb(s, Inches(7.0), Inches(2.0), Inches(5.6), Inches(0.85),
       [("企业家战略与业务落地", 20, WHITE, True)])
    right = [
        "AI 战略、组织变革、投融资",
        "真实经营课题，而非提示词训练",
        "标杆企业现场教学",
        "每位学员一份实施计划",
        "与技能班做成进阶，避免校内重复",
    ]
    for i, line in enumerate(right):
        tb(s, Inches(7.0), Inches(3.0) + Inches(i * 0.42), Inches(5.5), Inches(0.38),
           [(line, 14, RGBColor(0xE4, 0xEA, 0xF2), False)])
    tb(s, Inches(0.5), Inches(5.95), Inches(12.3), Inches(0.85),
       [("公开页面已核对。该班结业证书属于这个已批准项目，新合作不能自行承诺同样的证书、校友身份或学历学位。", 13, MUTED, False)])
    foot(s, 13)


def s14():
    s = light("产品草案", AI_PLAN["name"])
    nums = [
        (str(AI_PLAN["target_n"]) + "人", "首期测算班额"),
        ("16,800", "元/人，测试价"),
        (AI_PLAN["months"], "周末集中授课"),
        (str(AI_PLAN["total_days"]) + "天", "6 天课堂 + 2 天现场"),
    ]
    for i, (n, label) in enumerate(nums):
        x = Inches(0.42) + Inches(i * 3.22)
        rrect(s, x, Inches(1.4), Inches(3.05), Inches(1.45), WHITE, 0.08)
        tb(s, x, Inches(1.52), Inches(3.05), Inches(0.6),
           [(n, 26, CRIMSON, True)], align=PP_ALIGN.CENTER)
        tb(s, x, Inches(2.2), Inches(3.05), Inches(0.4),
           [(label, 13, MUTED, False)], align=PP_ALIGN.CENTER)
    modules = [
        "AI 战略与产业趋势",
        "企业应用与价值创造",
        "科技成果转化",
        "投融资与企业成长",
        "组织与数字化管理",
        "企业 AI 项目工作坊",
        "标杆企业现场教学 2 天",
    ]
    for i, name in enumerate(modules):
        x = Inches(0.42) + Inches((i % 4) * 3.22)
        y = Inches(3.15) + Inches((i // 4) * 0.85)
        rrect(s, x, y, Inches(3.05), Inches(0.72), NAVY if i == 6 else WHITE, 0.12)
        color = WHITE if i == 6 else NAVY
        tb(s, x + Inches(0.12), y + Inches(0.16), Inches(2.8), Inches(0.42),
           [(name, 13, color, True)], align=PP_ALIGN.CENTER)
    tb(s, Inches(0.5), Inches(5.85), Inches(12.3), Inches(0.95),
       [("交付物是一份《企业 AI 应用机会与实施计划》，包含场景、条件和 90 天行动。不承诺培训之后必然降本、增收或融资。名称、学费和证书以学院审批为准。", 14, MUTED, False)])
    foot(s, 14)


def s15():
    s = light("招生", "先访谈决策者，再谈 25 人开班")
    rrect(s, Inches(0.45), Inches(1.4), Inches(12.4), Inches(1.15), NAVY, 0.06)
    tb(s, Inches(0.7), Inches(1.58), Inches(11.9), Inches(0.85),
       [("约 2 万人的历史社群，是累计资源，不是有效生源。第一期先从会员、论坛和重点企业里筛出 30—50 位真正要做 AI 转型的决策者。", 16, WHITE, False)])
    steps = [("30+", "需求访谈"), ("120", "合法联系"), ("70", "有效线索"), ("45", "正式咨询"), ("25", "审核缴费")]
    for i, (n, label) in enumerate(steps):
        x = Inches(0.45) + Inches(i * 2.55)
        rrect(s, x, Inches(2.85), Inches(2.4), Inches(1.55), WHITE, 0.08)
        tb(s, x, Inches(3.0), Inches(2.4), Inches(0.55),
           [(n, 24, CRIMSON, True)], align=PP_ALIGN.CENTER)
        tb(s, x, Inches(3.65), Inches(2.4), Inches(0.4),
           [(label, 13, NAVY, True)], align=PP_ALIGN.CENTER)
    mix = [("会员定向 8人", "32%"), ("央国企团体 6人", "24%"), ("论坛公开课 5人", "20%"), ("私域咨询 3人", "12%"), ("学院渠道 3人", "12%")]
    for i, (name, pct) in enumerate(mix):
        x = Inches(0.45) + Inches(i * 2.55)
        tb(s, x, Inches(4.65), Inches(2.4), Inches(0.28), [(pct, 14, CRIMSON, True)], align=PP_ALIGN.CENTER)
        tb(s, x, Inches(4.98), Inches(2.4), Inches(0.5), [(name, 12, MUTED, False)], align=PP_ALIGN.CENTER)
    tb(s, Inches(0.5), Inches(5.7), Inches(12.3), Inches(1.05),
       [("上表是渠道配置目标，不是已经测到的转化率。报名、收费和宣传走学院批准的渠道。未审批前，不以学校名义收费或承诺开班。", 14, MUTED, False)])
    foot(s, 15)


def s16():
    s = light("测算", "16,800 元只是测试价，余额不是利润")
    headers = ["人数", "理论学费", "假设可支配", "假设余额"]
    rows = [("20人", "33.6万", "16.8万", "1.3万"), ("25人", "42.0万", "21.0万", "5.5万"), ("30人", "50.4万", "25.2万", "9.7万")]
    xs = [0.5, 3.5, 6.6, 9.7]
    widths = [2.7, 2.8, 2.8, 2.8]
    for x, w, h in zip(xs, widths, headers):
        tb(s, Inches(x), Inches(1.45), Inches(w), Inches(0.35), [(h, 13, CRIMSON, True)])
    for r, row in enumerate(rows):
        y = Inches(1.95) + Inches(r * 0.72)
        bg = WHITE if r != 1 else SOFT
        rrect(s, Inches(0.45), y, Inches(12.4), Inches(0.64), bg, 0.08)
        for x, w, val in zip(xs, widths, row):
            bold = r == 1
            tb(s, Inches(x), y + Inches(0.14), Inches(w), Inches(0.38),
               [(val, 16, NAVY if not bold else CRIMSON, True)])
    rrect(s, Inches(0.45), Inches(4.25), Inches(4.0), Inches(1.55), NAVY, 0.08)
    tb(s, Inches(0.68), Inches(4.42), Inches(3.6), Inches(0.3), [("假设平衡点", 13, GOLD, True)])
    tb(s, Inches(0.68), Inches(4.8), Inches(3.6), Inches(0.55), [("约 19 人", 28, WHITE, True)])
    tb(s, Inches(0.68), Inches(5.4), Inches(3.6), Inches(0.25), [("按当前假设", 12, RGBColor(0xD5, 0xDE, 0xEA), False)])
    tb(s, Inches(4.7), Inches(4.3), Inches(8.1), Inches(1.7), [
        [("假设条件", 13, CRIMSON, True, 4)],
        [("可支配资金按学费的 50% 估算。班均直接成本暂按 15.5 万元，且不随人数变动。", 14, INK, False, 4)],
        [("这两项都不是学院确认的管理费或成本。余额不能分配，正式立项前由学院重做预算。", 14, INK, False)],
    ])
    tb(s, Inches(0.5), Inches(6.05), Inches(12.3), Inches(0.75),
       [("工作表里这三个假设是金色输入格，改一个数，情景和平衡点会重算。公开课测试价 1,980—3,980 元，企业定制班单独报价。", 13, MUTED, False)])
    foot(s, 16)


def s17():
    s = light("商业安排", "三种合同路径，优先直接委托学院")
    modes = [
        ("模式 A", "优先", "客户直接委托学院", "央国企、政府和金融机构具备直接签约条件时使用。学院教学、交付和统一收费。联合会只做书面约定的调研、参访或运营，按工作和成果结算。"),
        ("模式 B", "有条件", "联合会整合后委托", "联合会自身有合法签约资质、真实采购关系时使用。教学部分按学校规定采购或委托。双方按真实服务分别签约，不用虚构主体规避审批。"),
        ("模式 C", "联合开发", "短训、公开课、参访", "立项前分开核算教学、招生、场地、参访和宣传。授权主体收款，按合同验收。盈利安排要经过双方财务和职能部门。"),
    ]
    for i, (name, tag, title, body) in enumerate(modes):
        x = Inches(0.42) + Inches(i * 4.28)
        rrect(s, x, Inches(1.42), Inches(4.1), Inches(4.55), WHITE, 0.06)
        rect(s, x, Inches(1.42), Inches(4.1), Inches(0.08), CRIMSON if i == 0 else NAVY)
        tb(s, x + Inches(0.22), Inches(1.65), Inches(2.3), Inches(0.3), [(name, 14, CRIMSON, True)])
        tb(s, x + Inches(2.2), Inches(1.65), Inches(1.7), Inches(0.3), [(tag, 13, GOLD, True)], align=PP_ALIGN.RIGHT)
        tb(s, x + Inches(0.22), Inches(2.15), Inches(3.65), Inches(0.7), [(title, 18, NAVY, True)])
        tb(s, x + Inches(0.22), Inches(3.0), Inches(3.65), Inches(2.6), [(body, 13, INK, False)])
    tb(s, Inches(0.5), Inches(6.15), Inches(12.3), Inches(0.65),
       [("会议里讨论过的内部分成，不预先写成已经批准的比例。", 14, MUTED, False)])
    foot(s, 17)


def s18():
    s = light("合规", "正式文件只保留可以执行的合作")
    tb(s, Inches(0.5), Inches(1.38), Inches(12.2), Inches(0.45),
       [("采用真实服务、真实合同、依法开票、校方审批。", 18, NAVY, True)])
    donts = [
        ("管理费", "不把学院内部管理费比例写进合作文件。"),
        ("票据", "不用不同票据、个人账户或费用拆分调节利润。"),
        ("证书", "不承诺学历学位，也不承诺超出规定的校友权益。"),
        ("品牌", "未获授权，不使用校名、标识，也不发布招生宣传。"),
    ]
    for i, (t, b) in enumerate(donts):
        col, row = i % 2, i // 2
        x = Inches(0.45) + Inches(col * 6.4)
        y = Inches(2.05) + Inches(row * 1.7)
        rrect(s, x, y, Inches(6.15), Inches(1.52), WHITE, 0.07)
        rect(s, x, y, Inches(0.08), Inches(1.52), CRIMSON)
        tb(s, x + Inches(0.32), y + Inches(0.22), Inches(5.5), Inches(0.35), [(t, 16, CRIMSON, True)])
        tb(s, x + Inches(0.32), y + Inches(0.68), Inches(5.5), Inches(0.6), [(b, 15, INK, False)])
    tb(s, Inches(0.5), Inches(5.65), Inches(12.3), Inches(1.15),
       [("政府项目涉及采购和招投标的，走法定程序。研究咨询要先写明承担主体和工作范围。参访报价分开列示接待费、会务交通费和资源协调费。", 14, MUTED, False)])
    foot(s, 18)


def s19():
    s = light("机制", "每个项目有负责人，每月对一次进度")
    items = [
        ("01", "联合工作小组", "双方各一名项目负责人，配置市场、教学、产业和财务对接。"),
        ("02", "项目负责人", "需求、立项、预算、交付、验收、复盘一条线走完。"),
        ("03", "共享清单", "客户、课程、师资方向、参访、渠道和项目台账。个人信息按保密和授权处理。"),
        ("04", "月度沟通", "每月推进会，每季度复盘，决定新增课程和资源投入。"),
        ("05", "品牌与质量", "校名、标识、证书、校友身份和校园参访，都以正式授权为准。"),
    ]
    for i, (n, t, b) in enumerate(items):
        y = Inches(1.38) + Inches(i * 1.02)
        tb(s, Inches(0.5), y, Inches(0.7), Inches(0.4), [(n, 16, CRIMSON, True)])
        tb(s, Inches(1.35), y, Inches(3.3), Inches(0.4), [(t, 16, NAVY, True)])
        tb(s, Inches(4.8), y, Inches(8.0), Inches(0.7), [(b, 14, INK, False)])
    foot(s, 19)


def s20():
    s = light("排期", "2026年11月到2027年下半年")
    rows = [
        ("2026.11", "梳理需求和资源", "框架、参访菜单、客户线索"),
        ("2026.12", "明确试点和边界", "备忘录草案、课程和预算"),
        ("2027 Q1", "按审批启动试点", "1个培训、1个参访、1次招生测试"),
        ("2027 Q2", "复制已经跑通的", "标准课程、固定线路、项目台账"),
        ("2027 下半年", "再扩大地域和客户", "长三角、地方政府、特色研修"),
    ]
    for i, (when, task, result) in enumerate(rows):
        y = Inches(1.4) + Inches(i * 1.02)
        rrect(s, Inches(0.45), y, Inches(12.4), Inches(0.9), WHITE, 0.1)
        tb(s, Inches(0.65), y + Inches(0.22), Inches(2.3), Inches(0.45), [(when, 15, CRIMSON, True)], anchor=MSO_ANCHOR.MIDDLE)
        tb(s, Inches(3.1), y + Inches(0.22), Inches(4.3), Inches(0.45), [(task, 16, NAVY, True)], anchor=MSO_ANCHOR.MIDDLE)
        tb(s, Inches(7.6), y + Inches(0.22), Inches(4.9), Inches(0.45), [(result, 14, MUTED, False)], anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 20)


def s21():
    s = light("2027", "先完成四项核心业务")
    cores = [
        ("委托培训", "2—3 个项目"),
        ("AI 研修班", "1 期，成熟后再开"),
        ("标杆参访", "6—10 批"),
        ("公开课论坛", "4—6 场"),
    ]
    for i, (t, b) in enumerate(cores):
        x = Inches(0.45) + Inches(i * 3.2)
        rrect(s, x, Inches(1.42), Inches(3.05), Inches(1.7), NAVY, 0.08)
        tb(s, x + Inches(0.18), Inches(1.58), Inches(2.7), Inches(0.55), [(t, 16, WHITE, True)])
        tb(s, x + Inches(0.18), Inches(2.25), Inches(2.7), Inches(0.55), [(b, 14, GOLD, True)])
    tb(s, Inches(0.5), Inches(3.35), Inches(12), Inches(0.3), [("重点培育", 13, CRIMSON, True)])
    grows = ["地方政府和园区人才培训", "科技金融与投融资专题", "银发经济与智慧康养"]
    for i, t in enumerate(grows):
        x = Inches(0.45) + Inches(i * 4.25)
        rrect(s, x, Inches(3.75), Inches(4.08), Inches(0.7), WHITE, 0.12)
        tb(s, x, Inches(3.9), Inches(4.08), Inches(0.4), [(t, 14, NAVY, True)], align=PP_ALIGN.CENTER)
    tb(s, Inches(0.5), Inches(4.65), Inches(12), Inches(0.3), [("储备，不作为今年收入考核", 13, CRIMSON, True)])
    holds = ["企业微咨询", "长三角与国际团组", "中国哲学与人文领导力"]
    for i, t in enumerate(holds):
        x = Inches(0.45) + Inches(i * 4.25)
        rrect(s, x, Inches(5.05), Inches(4.08), Inches(0.62), WHITE, 0.12)
        tb(s, x, Inches(5.18), Inches(4.08), Inches(0.38), [(t, 14, MUTED, False)], align=PP_ALIGN.CENTER)
    tb(s, Inches(0.5), Inches(5.9), Inches(12.3), Inches(0.9),
       [("试点看完成的班次、真实需求、复购、教学反馈和资源沉淀。工作表里有一套约 96.6 万元的规模示例，只供讨论量级，不是任务，也不归任何一方。", 14, MUTED, False)])
    foot(s, 21)


def s22():
    s = light("备忘录", "一年试点，非排他，先写意向")
    items = [
        "合作主体、范围和目标",
        "资源投入和责任边界",
        "立项、审批和授权程序",
        "客户对接和成果归属",
        "培训、参访、运营的合同",
        "结算、费用和风险承担",
        "宣传、知识产权和保密",
        "学员、质量、证书和售后",
        "期限、退出和争议解决",
        "未审批不得对外承诺",
    ]
    for i, t in enumerate(items):
        col, row = i % 2, i // 2
        x = Inches(0.48) + Inches(col * 6.4)
        y = Inches(1.4) + Inches(row * 0.85)
        tb(s, x, y, Inches(0.55), Inches(0.4), [(f"{i + 1:02d}", 14, CRIMSON, True)])
        tb(s, x + Inches(0.6), y, Inches(5.5), Inches(0.4), [(t, 16, NAVY, True)])
    tb(s, Inches(0.5), Inches(5.85), Inches(12.3), Inches(0.95),
       [("费用、品牌使用、招生承诺和业务交付，放进后续专项合同。备忘录本身不承载这些承诺。", 15, MUTED, False)])
    foot(s, 22)


def s23():
    s = light("近期", "资料交换之后，再定文本和试点")
    rrect(s, Inches(0.45), Inches(1.42), Inches(6.15), Inches(4.55), WHITE, 0.05)
    rrect(s, Inches(6.8), Inches(1.42), Inches(6.05), Inches(4.55), WHITE, 0.05)
    rect(s, Inches(0.45), Inches(1.42), Inches(6.15), Inches(0.5), CRIMSON)
    rect(s, Inches(6.8), Inches(1.42), Inches(6.05), Inches(0.5), NAVY)
    tb(s, Inches(0.68), Inches(1.52), Inches(5.7), Inches(0.34), [("联合会这次带来", 15, WHITE, True)])
    tb(s, Inches(7.02), Inches(1.52), Inches(5.6), Inches(0.34), [("请学院协调", 15, WHITE, True)])
    left = [
        "参访资源清单和 12 个优先项",
        "客户线索，以及三家 P0 的引荐安排",
        "研修班的渠道、定位和价格测试",
        "2027 年核心、重点、储备项目表",
    ]
    right = [
        "现有委托培训和课程产品简介",
        "立项、合同、招生、证书的管理要求",
        "安排学院负责人讨论合作边界",
        "对参访和课程提出教学与费用要求",
    ]
    for i, line in enumerate(left):
        tb(s, Inches(0.7), Inches(2.15) + Inches(i * 0.85), Inches(5.6), Inches(0.65),
           [(f"{i + 1}    {line}", 15, INK, False)])
    for i, line in enumerate(right):
        tb(s, Inches(7.05), Inches(2.15) + Inches(i * 0.85), Inches(5.5), Inches(0.65),
           [(f"{i + 1}    {line}", 15, INK, False)])
    foot(s, 23)


def s24():
    s = light("请学院确认", "下次见面，五件事定下来就能往下走")
    questions = [
        ("01", "签约主体", "备忘录由哪个校内主体签署，审批路径经过哪些环节。"),
        ("02", "参访采购", "现场教学按什么价格口径进入现有培训项目，谁来验收。"),
        ("03", "产品关系", "企业家研修班与现有 AIGC 技能班如何分工，是否联合拓展。"),
        ("04", "授权边界", "校名、标识、证书、招生宣传和校园活动，各自授权到哪一步。"),
        ("05", "共同访谈", "国泰海通、国药、中远三家，学院是否派人一起做第一轮需求访谈。"),
    ]
    for i, (n, t, b) in enumerate(questions):
        y = Inches(1.35) + Inches(i * 1.05)
        tb(s, Inches(0.5), y, Inches(0.7), Inches(0.35), [(n, 16, CRIMSON, True)])
        tb(s, Inches(1.3), y, Inches(2.5), Inches(0.35), [(t, 16, NAVY, True)])
        tb(s, Inches(3.9), y, Inches(8.8), Inches(0.7), [(b, 15, INK, False)])
    foot(s, 24)


def s25():
    s = new(NAVY)
    rect(s, 0, 0, SW, Inches(0.12), CRIMSON)
    tb(s, Inches(0.72), Inches(0.85), Inches(11), Inches(0.3),
       [("框架核心", 14, GOLD, True)])
    tb(s, Inches(0.7), Inches(1.4), Inches(11.8), Inches(2.2), [
        [("以委托培训形成业务基础", 28, WHITE, True)],
        [("以产业参访形成差异化", 28, WHITE, True)],
        [("以科创短训形成增长空间", 28, WHITE, True)],
    ])
    rect(s, Inches(0.74), Inches(3.85), Inches(1.4), Inches(0.05), GOLD)
    tb(s, Inches(0.72), Inches(4.15), Inches(11.5), Inches(1.1),
       [("下一次沟通最有用的三份成果：合作备忘录草案、企业参访资源菜单、人工智能与科创企业家研修班方案。今天这套材料，就是这三份成果的洽谈稿。", 16, RGBColor(0xD5, 0xDE, 0xEA), False)])
    tb(s, Inches(0.72), Inches(5.7), Inches(11), Inches(0.7), [
        [(META["compiler"] + "  编制", 14, WHITE, True)],
        [(META["compiled"] + "    " + META["version"], 13, RGBColor(0xC5, 0xCD, 0xD6), False)],
    ])
    foot(s, 25, dark=True)


def build(path: Path):
    for fn in (s01, s02, s03, s04, s05, s06, s07, s08, s09, s10,
               s11, s12, s13, s14, s15, s16, s17, s18, s19, s20,
               s21, s22, s23, s24, s25):
        fn()
    prs.save(path)


if __name__ == "__main__":
    out = Path(__file__).resolve().parents[1] / "output"
    out.mkdir(exist_ok=True)
    target = out / "复旦终身教育学院与杨浦科联合作洽谈.pptx"
    build(target)
    print(target, len(prs.slides))
