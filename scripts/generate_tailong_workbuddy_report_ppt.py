#!/usr/bin/env python3
"""生成《泰隆银行上海分行与腾讯 WorkBuddy 合作推进情况汇报》精简 PPT（5页）。"""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt


# 银行汇报风格：深蓝主色 + 金色点缀
NAVY = RGBColor(0x0B, 0x2F, 0x5B)
NAVY_DARK = RGBColor(0x06, 0x1E, 0x3C)
GOLD = RGBColor(0xC4, 0xA3, 0x5A)
LIGHT = RGBColor(0xF3, 0xF6, 0xFA)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1F, 0x2A, 0x37)
GRAY = RGBColor(0x5C, 0x6B, 0x7A)
SOFT = RGBColor(0xE8, 0xEE, 0xF5)
ACCENT_RED = RGBColor(0xA6, 0x3D, 0x2F)

FONT = "Microsoft YaHei"
FOOTER_TEXT = "泰隆银行上海分行 × 腾讯 WorkBuddy · 合作推进情况汇报 · 2026-08-03"


def add_rect(slide, x, y, w, h, fill, *, line=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
    shp.shadow.inherit = False
    return shp


def add_round(slide, x, y, w, h, fill):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    shp.adjustments[0] = 0.08
    return shp


def set_run(run, text, *, size=16, bold=False, color=DARK, font=FONT):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def add_text(
    slide,
    x,
    y,
    w,
    h,
    text,
    *,
    size=16,
    bold=False,
    color=DARK,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    font=FONT,
):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        set_run(p.add_run(), line, size=size, bold=bold, color=color, font=font)
    return tb


def add_bullets(slide, x, y, w, h, items, *, size=15, color=DARK, bullet_color=GOLD, spacing=1.25):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = spacing
        set_run(p.add_run(), "●  ", size=size, color=bullet_color)
        set_run(p.add_run(), item, size=size, color=color)
    return tb


def header(slide, title, subtitle=None):
    add_rect(slide, 0, 0, SW, Inches(0.95), NAVY)
    add_rect(slide, 0, Inches(0.95), SW, Inches(0.05), GOLD)
    add_text(slide, Inches(0.5), Inches(0.2), Inches(12.2), Inches(0.45), title, size=24, bold=True, color=WHITE)
    if subtitle:
        add_text(slide, Inches(0.5), Inches(0.58), Inches(12.2), Inches(0.3), subtitle, size=12, color=LIGHT)


def footer(slide, page, total=5):
    add_text(slide, Inches(0.5), Inches(7.12), Inches(10.5), Inches(0.28), FOOTER_TEXT, size=10, color=GRAY)
    add_text(
        slide,
        Inches(11.5),
        Inches(7.12),
        Inches(1.3),
        Inches(0.28),
        f"{page} / {total}",
        size=10,
        color=GRAY,
        align=PP_ALIGN.RIGHT,
    )


def build_ppt(output_path: Path) -> None:
    global SW, SH
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    SW, SH = prs.slide_width, prs.slide_height
    blank = prs.slide_layouts[6]

    # ========== 第1页：封面 / 总体判断 ==========
    s = prs.slides.add_slide(blank)
    add_rect(s, 0, 0, SW, SH, NAVY_DARK)
    add_rect(s, 0, 0, Inches(0.18), SH, GOLD)
    add_text(s, Inches(0.8), Inches(1.35), Inches(11.5), Inches(0.4), "行长专题汇报 · 内部材料", size=14, color=GOLD)
    add_text(
        s,
        Inches(0.8),
        Inches(1.9),
        Inches(11.5),
        Inches(1.3),
        "泰隆银行上海分行与腾讯 WorkBuddy\n合作推进情况汇报",
        size=36,
        bold=True,
        color=WHITE,
    )
    add_rect(s, Inches(0.8), Inches(3.45), Inches(2.2), Inches(0.05), GOLD)
    add_text(
        s,
        Inches(0.8),
        Inches(3.75),
        Inches(11.5),
        Inches(1.1),
        "两轮洽谈后，双方已明确以「积分商城兑换 + 开户赠送AI权益」为主要合作方向；\n"
        "建议原则同意推进，1,500万元仅为方案框架，先合规审查与小范围试点，再分阶段采购。",
        size=16,
        color=LIGHT,
    )
    add_text(
        s,
        Inches(0.8),
        Inches(5.5),
        Inches(11.5),
        Inches(0.4),
        "汇报对象：泰隆银行上海分行行长　　汇报日期：2026年8月3日",
        size=14,
        color=GOLD,
    )
    add_text(
        s,
        Inches(0.8),
        Inches(6.2),
        Inches(11.5),
        Inches(0.35),
        "状态分层：已达成共识｜初步方案｜尚待确认",
        size=13,
        color=GRAY,
    )

    # ========== 第2页：合作方向与模式 ==========
    s = prs.slides.add_slide(blank)
    add_rect(s, 0, 0, SW, SH, WHITE)
    header(s, "一、合作方向与拟议模式", "服务中小微企业 · 纳入客户权益体系 · 助力新增对公客户")

    # 背景要点条
    add_round(s, Inches(0.45), Inches(1.2), Inches(12.4), Inches(0.85), SOFT)
    add_text(
        s,
        Inches(0.7),
        Inches(1.35),
        Inches(12.0),
        Inches(0.55),
        "背景：上海经营约16年、近100万客户，传统产品同质化明显；需引入贴近企业经营的非金融增值服务，提高开户吸引力与存量黏性。\n"
        "对方：WorkBuddy 为企业级AI应用及技能平台（约2.8万家企业客户，企业版日活超10万），适合中小微低成本体验入口。",
        size=12,
        color=DARK,
    )

    cards = [
        ("已达成共识", "积分商城兑换", "集中采购账号/额度/权益\n上架积分商城兑换或购买\n丰富积分消耗场景"),
        ("已达成共识", "开户/拜访赠礼", "以AI账号或额度替代部分\n实物礼品；兑换码自助激活\n提升拜访体验与开户转化"),
        ("初步方案", "联合推广", "依托130余网点与客户经理\n面向制造业、商贸等客群试点\n腾讯负责产品与技术支持"),
        ("后续延伸", "平台专区等", "开设泰隆银行专区\n联合开发行业AI技能\n与开户/结算/代发结合"),
    ]
    for i, (tag, title, body) in enumerate(cards):
        x = Inches(0.45 + i * 3.2)
        y = Inches(2.3)
        add_round(s, x, y, Inches(3.0), Inches(4.2), NAVY if i < 2 else SOFT)
        tag_color = GOLD if i < 2 else (NAVY if i == 2 else GRAY)
        title_color = WHITE if i < 2 else NAVY
        body_color = LIGHT if i < 2 else DARK
        add_text(s, x + Inches(0.2), y + Inches(0.25), Inches(2.6), Inches(0.35), tag, size=12, bold=True, color=tag_color)
        add_text(s, x + Inches(0.2), y + Inches(0.7), Inches(2.6), Inches(0.7), title, size=20, bold=True, color=title_color)
        add_text(s, x + Inches(0.2), y + Inches(1.6), Inches(2.6), Inches(2.2), body, size=13, color=body_color)
    footer(s, 2)

    # ========== 第3页：采购框架与预期价值 ==========
    s = prs.slides.add_slide(blank)
    add_rect(s, 0, 0, SW, SH, WHITE)
    header(s, "二、采购方案设想与预期价值", "1,500万元为测算框架，非采购承诺；首期建议300万/500万两档论证")

    # 左：采购框架
    add_round(s, Inches(0.45), Inches(1.25), Inches(6.3), Inches(5.5), SOFT)
    add_text(s, Inches(0.7), Inches(1.45), Inches(5.8), Inches(0.4), "采购框架（初步方案）", size=18, bold=True, color=NAVY)
    metrics = [
        ("最高约1,500万", "总体方案设计上限"),
        ("三阶段 × 约500万", "按阶段推进与验收"),
        ("首期300–500万", "待报价与试点测算"),
    ]
    for i, (big, small) in enumerate(metrics):
        y = Inches(2.05 + i * 1.05)
        add_round(s, Inches(0.75), y, Inches(5.7), Inches(0.9), WHITE)
        add_rect(s, Inches(0.75), y, Inches(0.12), Inches(0.9), GOLD)
        add_text(s, Inches(1.1), y + Inches(0.12), Inches(5.1), Inches(0.4), big, size=18, bold=True, color=NAVY)
        add_text(s, Inches(1.1), y + Inches(0.5), Inches(5.1), Inches(0.3), small, size=12, color=GRAY)

    add_text(
        s,
        Inches(0.75),
        Inches(5.35),
        Inches(5.7),
        Inches(1.1),
        "尚待腾讯提供：账号/License/Token数量、有效期、\n阶梯折扣、未激活处理、扩容续费退款、付款与验收。",
        size=12,
        color=DARK,
    )

    # 右：预期价值
    add_round(s, Inches(7.0), Inches(1.25), Inches(5.85), Inches(5.5), NAVY)
    add_text(s, Inches(7.3), Inches(1.45), Inches(5.3), Inches(0.4), "预期价值", size=18, bold=True, color=GOLD)
    values = [
        "带动新增对公客户：差异化开户/拜访权益",
        "提升存量客户黏性：触达老板/财务关键人",
        "促进积分活跃：补充企业服务类权益",
        "带动结算与资金沉淀：商城/续费走我行账户",
        "塑造科技服务品牌：金融+经营赋能形象",
    ]
    add_bullets(s, Inches(7.3), Inches(2.1), Inches(5.3), Inches(4.0), values, size=14, color=WHITE, bullet_color=GOLD, spacing=1.55)
    footer(s, 3)

    # ========== 第4页：风险与共识 ==========
    s = prs.slides.add_slide(blank)
    add_rect(s, 0, 0, SW, SH, WHITE)
    header(s, "三、风险事项与初步共识", "合规、采购、技术、结算仍待确认；优先推进两类场景")

    # 左侧风险
    add_text(s, Inches(0.5), Inches(1.2), Inches(6.2), Inches(0.35), "尚待确认 · 主要风险", size=16, bold=True, color=ACCENT_RED)
    risks = [
        ("安全合规", "公有云SaaS不接行内网，但仍需审查数据隔离、访问权限、日志与责任机制；首期仅用于外部客户权益"),
        ("采购成本", "缺兑换/激活/使用/转化数据，忌一次性大额采购；建议分期交付与效果验收"),
        ("技术交付", "批量兑换码、自助注册激活、分组与续费等需完整流程测试"),
        ("结算税务", "商城主体、发票与资金链路需财务/法务确认"),
        ("宣传口径", "联合品牌与效果表述须授权并经宣传审核"),
    ]
    for i, (t, d) in enumerate(risks):
        y = Inches(1.6 + i * 0.95)
        add_round(s, Inches(0.45), y, Inches(6.3), Inches(0.85), SOFT)
        add_text(s, Inches(0.65), y + Inches(0.1), Inches(5.9), Inches(0.3), t, size=13, bold=True, color=NAVY)
        add_text(s, Inches(0.65), y + Inches(0.4), Inches(5.9), Inches(0.4), d, size=11, color=DARK)

    # 右侧共识
    add_text(s, Inches(7.1), Inches(1.2), Inches(5.7), Inches(0.35), "已达成共识（两轮谈判）", size=16, bold=True, color=NAVY)
    add_round(s, Inches(7.0), Inches(1.6), Inches(5.85), Inches(5.15), SOFT)
    consensus = [
        "优先推进：积分商城兑换 + 开户赠礼",
        "采用兑换码/License，客户自助激活",
        "暂不印实体卡，后续再定电子/纸质形式",
        "按季度或半年度套餐及阶梯价设计",
        "腾讯提交完整产品及商务方案",
        "8月5日下午与分行领导进一步沟通",
        "上海先行试点，效果好再扩围",
    ]
    add_bullets(s, Inches(7.3), Inches(1.9), Inches(5.4), Inches(4.6), consensus, size=14, color=DARK, spacing=1.45)
    footer(s, 4)

    # ========== 第5页：请示事项与下一步 ==========
    s = prs.slides.add_slide(blank)
    add_rect(s, 0, 0, SW, SH, WHITE)
    header(s, "四、建议请示事项与下一步", "原则同意立项论证 · 授权继续谈判 · 先试点后扩量")

    decisions = [
        ("01", "原则同意继续推进", "作为中小微增值服务创新方向继续论证"),
        ("02", "同意上海地区试点", "选定网点/客群，设定兑换激活获客指标"),
        ("03", "启动跨部门评估", "公司金融、权益、科技、合规、法务、财务、采购"),
        ("04", "授权继续商务谈判", "1,500万为测算框架；首期金额另行审批"),
    ]
    for i, (num, title, desc) in enumerate(decisions):
        x = Inches(0.45 + (i % 2) * 6.4)
        y = Inches(1.25 + (i // 2) * 1.55)
        add_round(s, x, y, Inches(6.15), Inches(1.4), SOFT)
        add_round(s, x + Inches(0.2), y + Inches(0.3), Inches(0.8), Inches(0.8), NAVY)
        add_text(
            s,
            x + Inches(0.2),
            y + Inches(0.3),
            Inches(0.8),
            Inches(0.8),
            num,
            size=18,
            bold=True,
            color=GOLD,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text(s, x + Inches(1.2), y + Inches(0.3), Inches(4.7), Inches(0.45), title, size=16, bold=True, color=NAVY)
        add_text(s, x + Inches(1.2), y + Inches(0.75), Inches(4.7), Inches(0.45), desc, size=12, color=DARK)

    # 底部结论条
    add_round(s, Inches(0.45), Inches(4.5), Inches(12.4), Inches(2.2), NAVY)
    add_text(s, Inches(0.75), Inches(4.7), Inches(11.8), Inches(0.35), "汇报结论 / 下一步", size=14, bold=True, color=GOLD)
    add_text(
        s,
        Inches(0.75),
        Inches(5.15),
        Inches(11.8),
        Inches(1.3),
        "建议行长原则同意立项论证并授权继续谈判，先试点、后扩量，以实际客户转化和使用效果作为后续采购依据。\n"
        "下一步：① 8/5前腾讯提交完整方案　② 明确300万/500万两档权益　③ 完成兑换激活全流程演示\n"
        "　　　　④ 启动法务合规科技财务采购预审　⑤ 设计试点方案与量化指标　⑥ 达标后再启动二三阶段采购",
        size=13,
        color=WHITE,
    )
    footer(s, 5)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    print(f"已生成：{output_path}")


if __name__ == "__main__":
    out = Path(__file__).resolve().parents[1] / "deliverables" / "泰隆银行上海分行与腾讯WorkBuddy合作推进情况汇报_20260803.pptx"
    build_ppt(out)
