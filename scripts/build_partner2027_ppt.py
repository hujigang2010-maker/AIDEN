# -*- coding: utf-8 -*-
"""生成《泰隆银行·科创企业成长伙伴计划（2027）》商务沟通稿 PPT。"""

import os
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Emu, Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import partner2027_data as D

NAVY = RGBColor(0x0C, 0x1E, 0x33)
NAVY_MID = RGBColor(0x1A, 0x3A, 0x5C)
INK = RGBColor(0x1B, 0x28, 0x36)
MUTED = RGBColor(0x5C, 0x68, 0x74)
CREAM = RGBColor(0xF6, 0xF3, 0xEE)
PAPER = RGBColor(0xFF, 0xFC, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GOLD = RGBColor(0xA6, 0x84, 0x4A)
GOLD_SOFT = RGBColor(0xF4, 0xEC, 0xDC)
LINE = RGBColor(0xE4, 0xDB, 0xCE)
RED = RGBColor(0x8E, 0x2F, 0x2F)
RED_SOFT = RGBColor(0xF8, 0xEE, 0xEC)
NAVY_SOFT = RGBColor(0xE7, 0xEE, 0xF4)
OK_BG = RGBColor(0xE8, 0xF0, 0xEA)
OK = RGBColor(0x1E, 0x5A, 0x45)

FONT = "微软雅黑"
PAGE_W = Inches(13.333)
PAGE_H = Inches(7.5)
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "deliverables", "泰隆银行科创企业成长伙伴计划2027_商务沟通稿.pptx")

PRI_FILL = {
    "最高": RED,
    "高": NAVY,
    "中高": GOLD,
    "中": RGBColor(0xD9, 0xD2, 0xC6),
}
PRI_FG = {
    "最高": WHITE,
    "高": WHITE,
    "中高": NAVY,
    "中": INK,
}


def style_run(run, size, color, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", FONT)


def add_text(slide, x, y, w, h, paragraphs, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, spacing=1.05):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    for i, para in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        runs = [para] if isinstance(para, tuple) else para
        for text, size, color, bold in runs:
            r = p.add_run()
            r.text = text
            style_run(r, size, color, bold)
    return tb


def add_rect(slide, x, y, w, h, fill, line=None, line_w=0.75, radius=0.08, shape=None):
    sp = slide.shapes.add_shape(shape or MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    if sp.adjustments and shape is None:
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    return sp


def shape_text(sp, paragraphs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, spacing=1.02):
    tf = sp.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(0.12)
    tf.margin_right = Inches(0.12)
    tf.margin_top = Inches(0.06)
    tf.margin_bottom = Inches(0.06)
    for i, para in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        runs = [para] if isinstance(para, tuple) else para
        for text, size, color, bold in runs:
            r = p.add_run()
            r.text = text
            style_run(r, size, color, bold)
    return sp


def bg(slide, color=CREAM):
    add_rect(slide, 0, 0, PAGE_W, PAGE_H, color, radius=0, shape=MSO_SHAPE.RECTANGLE)


def header(slide, kicker, title, sub=None):
    add_rect(slide, 0, 0, PAGE_W, Inches(0.08), NAVY, radius=0, shape=MSO_SHAPE.RECTANGLE)
    add_rect(slide, Inches(0.48), Inches(0.26), Inches(0.22), Inches(0.08), GOLD,
             radius=0, shape=MSO_SHAPE.RECTANGLE)
    add_text(slide, Inches(0.78), Inches(0.18), Inches(10.5), Inches(0.26),
             [(kicker, 11, GOLD, True)])
    add_text(slide, Inches(0.46), Inches(0.42), Inches(12.4), Inches(0.46),
             [(title, 26, NAVY, True)])
    line_y = Inches(1.02)
    if sub:
        add_text(slide, Inches(0.48), Inches(0.92), Inches(12.3), Inches(0.30),
                 [(sub, 13, MUTED, False)])
        line_y = Inches(1.28)
    add_rect(slide, Inches(0.48), line_y, Inches(1.25), Pt(2.25), GOLD,
             radius=0, shape=MSO_SHAPE.RECTANGLE)
    return line_y + Inches(0.16)


def footer(slide, n, total):
    add_rect(slide, Inches(0.48), Inches(7.16), Inches(12.38), Pt(0.75), LINE,
             radius=0, shape=MSO_SHAPE.RECTANGLE)
    add_text(slide, Inches(0.48), Inches(7.20), Inches(9.2), Inches(0.24),
             [("科创企业成长伙伴计划（2027）· 商务沟通稿 · 预算、接待与品牌授权均待正式确认", 9, MUTED, False)])
    add_text(slide, Inches(10.7), Inches(7.20), Inches(2.16), Inches(0.24),
             [(f"{n:02d}  /  {total:02d}", 9, NAVY, True)], align=PP_ALIGN.RIGHT)


def pill(slide, x, y, w, h, text, fill, fg, size=11):
    sp = add_rect(slide, x, y, w, h, fill, radius=0.5)
    shape_text(sp, [(text, size, fg, True)], align=PP_ALIGN.CENTER)
    return sp


def set_cell_border(cell, color_hex="E4DBCE", width_pt=0.6):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    ln_w = str(int(width_pt * 12700))
    for edge in ("lnL", "lnR", "lnT", "lnB"):
        tag = qn(f"a:{edge}")
        old = tcPr.find(tag)
        if old is not None:
            tcPr.remove(old)
        ln = OxmlElement(f"a:{edge}")
        ln.set("w", ln_w)
        ln.set("cap", "flat")
        ln.set("cmpd", "sng")
        ln.set("algn", "ctr")
        sf = OxmlElement("a:solidFill")
        c = OxmlElement("a:srgbClr")
        c.set("val", color_hex)
        sf.append(c)
        ln.append(sf)
        tcPr.append(ln)


def cell_text(cell, text, size, color, bold=False, fill=None, align=PP_ALIGN.LEFT):
    cell.text = ""
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    else:
        cell.fill.background()
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = cell.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.06)
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.03)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    style_run(r, size, color, bold)
    set_cell_border(cell)


def add_table(slide, x, y, w, h, rows, col_w, font=12, header=True, aligns=None, bold_col=None):
    tbl_shape = slide.shapes.add_table(len(rows), len(col_w), x, y, w, h)
    tbl = tbl_shape.table
    for i, cw in enumerate(col_w):
        tbl.columns[i].width = cw
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            is_header = header and r == 0
            fill = NAVY if is_header else (PAPER if r % 2 else CREAM)
            fg = WHITE if is_header else INK
            bold = is_header or (bold_col is not None and c == bold_col and not is_header)
            al = PP_ALIGN.LEFT
            if aligns and c < len(aligns):
                al = aligns[c]
            cell_text(tbl.cell(r, c), val, font if not is_header else font, fg, bold or is_header, fill, al)
    for r in range(len(rows)):
        tbl.rows[r].height = int(h / len(rows))
    return tbl_shape


def s01(slide, n, total):
    bg(slide, NAVY)
    add_rect(slide, 0, 0, Inches(0.16), PAGE_H, GOLD, radius=0, shape=MSO_SHAPE.RECTANGLE)
    add_text(slide, Inches(0.72), Inches(0.48), Inches(8), Inches(0.32),
             [("商务沟通稿", 14, GOLD, True)])
    add_text(slide, Inches(0.72), Inches(0.82), Inches(6), Inches(0.28),
             [(D.DATE, 13, RGBColor(0xC9, 0xD3, 0xDE), False)])
    add_rect(slide, Inches(0.72), Inches(1.28), Inches(1.4), Pt(1.5), GOLD,
             radius=0, shape=MSO_SHAPE.RECTANGLE)
    add_text(slide, Inches(0.70), Inches(1.55), Inches(11.5), Inches(0.42),
             [("泰隆银行", 20, GOLD, True)])
    add_text(slide, Inches(0.68), Inches(2.05), Inches(12.0), Inches(0.85),
             [("科创企业成长伙伴计划", 40, WHITE, True)])
    add_text(slide, Inches(0.72), Inches(2.95), Inches(6), Inches(0.45),
             [("2027", 28, GOLD, True)])
    add_text(slide, Inches(0.72), Inches(3.65), Inches(11.4), Inches(0.85),
             [
                 [("以企业参访切入，以科技金融为核心。", 16, RGBColor(0xE7, 0xE2, 0xD6), False)],
                 [("以高校科创资源为差异化优势，把培训、产业对接和投融资做成长期服务。", 16, RGBColor(0xE7, 0xE2, 0xD6), False)],
             ], spacing=1.15)
    add_rect(slide, Inches(0.72), Inches(4.85), Inches(11.9), Inches(1.85), NAVY_MID, radius=0.06)
    cols = [
        ("建议合作单位", "泰隆银行\n上海市杨浦区科技企业联合会"),
        ("拟联动单位", "复旦科技园、高校继续教育\n及相关科研、园区资源"),
        ("文件性质", "洽谈稿，非正式合同\n预算、接待与授权均待确认"),
    ]
    for i, (h, b) in enumerate(cols):
        x = Inches(0.95 + i * 3.9)
        add_text(slide, x, Inches(5.05), Inches(3.5), Inches(0.30), [(h, 12, GOLD, True)])
        lines = [[(line, 14, WHITE, False)] for line in b.split("\n")]
        add_text(slide, x, Inches(5.40), Inches(3.55), Inches(0.95), lines, spacing=1.15)


def s02(slide, n, total):
    bg(slide)
    header(slide, "合作背景", "两次洽谈已经谈到的，和还不能写成定论的",
           "2026年11月6日企业服务洽谈，11月9日泰隆银行、复旦科技园及相关方交流")
    footer(slide, n, total)
    cards = [
        (RED, "11月6日", "企业服务合作洽谈", [
            "参访资源清单、预算和组织执行，拟由项目执行方承担",
            "年度参访经费曾口头讨论十几万至二十几万元",
            "该数字尚未审批，不能视为已批准预算",
            "银行关注客户子女的高校研学，入学权益不在合作范围",
        ]),
        (NAVY, "11月9日", "三方交流", [
            "虹口支行介绍了国际业务，以及航运、进口贸易客户基础",
            "11月18日拟举办约100人的科创路演",
            "泰隆方面考虑支持茶歇、会务和场地，并设置企业对接",
            "活动是否举行、费用是否获批，都还需要核实",
        ]),
    ]
    for i, (accent, date, name, bullets) in enumerate(cards):
        x = Inches(0.48 + i * 6.40)
        add_rect(slide, x, Inches(1.62), Inches(6.18), Inches(4.05), PAPER, LINE, 0.75, 0.06)
        add_rect(slide, x, Inches(1.62), Inches(0.08), Inches(4.05), accent, radius=0, shape=MSO_SHAPE.RECTANGLE)
        add_text(slide, x + Inches(0.28), Inches(1.78), Inches(5.6), Inches(0.28),
                 [(date, 12, accent, True)])
        add_text(slide, x + Inches(0.28), Inches(2.08), Inches(5.6), Inches(0.36),
                 [(name, 20, NAVY, True)])
        body = [[("·  " + b, 14, INK, False)] for b in bullets]
        add_text(slide, x + Inches(0.28), Inches(2.58), Inches(5.65), Inches(2.85), body, spacing=1.2)
    add_rect(slide, Inches(0.48), Inches(5.84), Inches(12.38), Inches(1.16), NAVY, radius=0.06)
    add_text(slide, Inches(0.72), Inches(6.00), Inches(11.9), Inches(0.28),
             [("据此形成的合作判断", 12, GOLD, True)])
    add_text(slide, Inches(0.72), Inches(6.30), Inches(11.9), Inches(0.52),
             [("争取有年度经费支持的科创企业综合服务机制：参访是入口，科技金融是核心，培训、产业对接和投融资形成持续合作。", 15, WHITE, False)])


def s03(slide, n, total):
    bg(slide)
    header(slide, "一页纸共识草案", "先把能一起确认的写在一页上",
           "供本次沟通使用。确认后仍须落入银行采购、合同和财务流程")
    footer(slide, n, total)
    rows = [
        ("项目名称", "泰隆银行·科创企业成长伙伴计划"),
        ("合作架构", "1个联合服务平台 + 3类资源 + 6项业务"),
        ("近期动作", "核实11月18日路演及支持费用；12月完成1场20—30人参访试点"),
        ("2027年节奏", "参访原则上每月1场；路演按4场重点+4场小型专项另行安排"),
        ("第一轮经费建议", "参访年度15万—20万元；路演年度14万元另案。均待审批，按实核定"),
        ("合作边界", "未授权不冠名；不承诺入学；不对个人按贷款、开户或转介结果付佣"),
    ]
    y = 1.58
    for i, (k, v) in enumerate(rows):
        fill = PAPER if i % 2 == 0 else GOLD_SOFT
        add_rect(slide, Inches(0.48), Inches(y), Inches(12.38), Inches(0.72), fill, radius=0.04)
        add_text(slide, Inches(0.68), Inches(y), Inches(2.35), Inches(0.72),
                 [(k, 14, NAVY, True)], anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(3.15), Inches(y), Inches(9.45), Inches(0.72),
                 [(v, 15, INK, False)], anchor=MSO_ANCHOR.MIDDLE)
        y += 0.78
    add_text(slide, Inches(0.50), Inches(6.32), Inches(12.3), Inches(0.70),
             [[("本页是洽谈共识草案，不构成合同、采购承诺或预算批准。", 13, RED, True)],
              [("承办主体、费用标准、发票、审批和结算，以银行正式流程为准。", 13, MUTED, False)]],
             spacing=1.15)


def s04(slide, n, total):
    bg(slide)
    header(slide, "总体定位", "1个平台，3类资源，6项业务", D.THESIS)
    footer(slide, n, total)
    add_rect(slide, Inches(0.48), Inches(1.70), Inches(12.38), Inches(1.05), NAVY, radius=0.06)
    add_text(slide, Inches(0.72), Inches(1.82), Inches(11.9), Inches(0.28),
             [("联合服务平台", 12, GOLD, True)])
    add_text(slide, Inches(0.72), Inches(2.12), Inches(11.9), Inches(0.42),
             [("泰隆银行  ×  杨浦区科技企业联合会，联动复旦科技园、高校继续教育与区域产业园区", 16, WHITE, False)])
    labels = [("企业资源", "商协会、园区、科技企业"),
              ("教育资源", "高校专家、课程、校友"),
              ("金融资源", "银行、投资机构、产业基金")]
    for i, (h, b) in enumerate(labels):
        x = Inches(0.48 + i * 4.18)
        add_rect(slide, x, Inches(2.92), Inches(4.02), Inches(1.22), PAPER, LINE, 0.75, 0.08)
        add_text(slide, x + Inches(0.2), Inches(3.05), Inches(3.6), Inches(0.32), [(h, 16, NAVY, True)])
        add_text(slide, x + Inches(0.2), Inches(3.42), Inches(3.6), Inches(0.48), [(b, 14, MUTED, False)])
    for i, name in enumerate(D.BUSINESSES):
        x = Inches(0.48 + (i % 6) * 2.08)
        add_rect(slide, x, Inches(4.34), Inches(1.98), Inches(1.35), PAPER, LINE, 0.75, 0.1)
        add_text(slide, x, Inches(4.48), Inches(1.98), Inches(0.36),
                 [(f"0{i+1}", 12, GOLD, True)], align=PP_ALIGN.CENTER)
        add_text(slide, x + Inches(0.08), Inches(4.86), Inches(1.82), Inches(0.68),
                 [(name, 14, NAVY, True)], align=PP_ALIGN.CENTER)
    add_text(slide, Inches(0.48), Inches(5.86), Inches(12.38), Inches(1.05),
             [[("服务闭环", 12, GOLD, True)],
              [(D.LOOP, 16, NAVY, True)],
              [("复旦科技园和学校单位先作为资源方。联合主办、品牌、课程和证书，单独取得授权后再写入对外材料。", 13, MUTED, False)]],
             spacing=1.08)


def s05(slide, n, total):
    bg(slide)
    header(slide, "六个方向", "第一项和第二项，放到2026年年底至2027年第一季度",
           "大型活动赞助保留合作位置，不作为这一轮的主要收入来源")
    footer(slide, n, total)
    y = 1.62
    for d in D.DIRECTIONS:
        add_rect(slide, Inches(0.48), Inches(y), Inches(12.38), Inches(0.78), PAPER, LINE, 0.6, 0.08)
        add_text(slide, Inches(0.62), Inches(y), Inches(0.55), Inches(0.78),
                 [(d["no"], 18, GOLD, True)], anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(1.20), Inches(y + 0.08), Inches(3.15), Inches(0.62),
                 [[(d["name"], 15, NAVY, True)], [(d["window"], 11, MUTED, False)]],
                 anchor=MSO_ANCHOR.MIDDLE, spacing=1.0)
        add_text(slide, Inches(4.45), Inches(y), Inches(6.15), Inches(0.78),
                 [(d["detail"], 13, INK, False)], anchor=MSO_ANCHOR.MIDDLE)
        pill(slide, Inches(10.85), Inches(y + 0.20), Inches(1.75), Inches(0.38),
             d["priority"], PRI_FILL[d["priority"]], PRI_FG[d["priority"]], 12)
        y += 0.84


def s06(slide, n, total):
    bg(slide)
    header(slide, "角色", "做综合服务的发起和运营方",
           "银行需要的是可持续的企业服务场景，单场到场人数只是其中一项记录")
    footer(slide, n, total)
    add_rect(slide, Inches(0.48), Inches(1.62), Inches(5.95), Inches(5.32), PAPER, LINE, 0.75, 0.06)
    add_text(slide, Inches(0.72), Inches(1.80), Inches(5.5), Inches(0.70),
             [[("项目执行方承担", 12, GOLD, True)],
              [("发起、策划与运营", 22, NAVY, True)]])
    assets = [
        ("01", "科技企业与商协会", "把参访和企业邀约做成可重复交付的年度产品"),
        ("02", "高校与复旦相关资源", "协调研修、研学和成果转化，授权范围单列"),
        ("03", "科技项目与投资机构", "把路演、产业对接和银行债权服务接到同一场活动"),
    ]
    yy = 2.70
    for num, title, desc in assets:
        add_text(slide, Inches(0.72), Inches(yy), Inches(5.4), Inches(0.28),
                 [(num + "   " + title, 15, NAVY, True)])
        add_text(slide, Inches(0.72), Inches(yy + 0.32), Inches(5.4), Inches(0.55),
                 [(desc, 13, MUTED, False)])
        yy += 1.15
    add_text(slide, Inches(6.70), Inches(1.62), Inches(6.1), Inches(0.36),
             [("请行长明确的五件事", 16, NAVY, True)])
    five_short = [
        ("年度预算量级", "口头量级先进入审批，不写成已批预算"),
        ("首场参访项目", "京东、美团、傅里叶、张江中落实1—2个"),
        ("承办主体与结算", "写明签约主体、发票、验收和付款"),
        ("年度活动计划", "12月提交参访年历，路演单独成项"),
        ("客户服务与资源权益", "银行经营客户，执行方运营参访产品"),
    ]
    for i, (title, desc) in enumerate(five_short):
        y = 2.12 + i * 0.92
        add_rect(slide, Inches(6.68), Inches(y), Inches(6.18), Inches(0.84), PAPER, LINE, 0.6, 0.08)
        sp = add_rect(slide, Inches(6.84), Inches(y + 0.22), Inches(0.40), Inches(0.40), NAVY, radius=0.5)
        shape_text(sp, [(str(i + 1), 13, WHITE, True)], align=PP_ALIGN.CENTER)
        add_text(slide, Inches(7.38), Inches(y + 0.08), Inches(5.3), Inches(0.30),
                 [(title, 14, NAVY, True)])
        add_text(slide, Inches(7.38), Inches(y + 0.40), Inches(5.3), Inches(0.38),
                 [(desc, 11, MUTED, False)])


def s07(slide, n, total):
    bg(slide)
    header(slide, "费用", "三档报价供谈判，第一轮先报参访",
           "以下不是泰隆银行已确认的采购价格")
    footer(slide, n, total)
    cards = [
        ("单场标杆参访", "5,000—10,000", "元 / 场", "策划、协调及组织执行服务费"),
        ("路演或专题活动", "10,000—30,000", "元 / 场", "活动承办与会务服务费"),
        ("年度综合服务", "15万—25万", "元 / 年", "分项目采购、验收与结算"),
    ]
    for i, (name, price, unit, note) in enumerate(cards):
        x = Inches(0.48 + i * 4.18)
        add_rect(slide, x, Inches(1.64), Inches(4.02), Inches(2.15), PAPER, LINE, 0.75, 0.06)
        add_rect(slide, x, Inches(1.64), Inches(4.02), Inches(0.08), GOLD if i < 2 else NAVY,
                 radius=0, shape=MSO_SHAPE.RECTANGLE)
        add_text(slide, x + Inches(0.2), Inches(1.84), Inches(3.6), Inches(0.32),
                 [(name, 14, MUTED, True)])
        add_text(slide, x + Inches(0.2), Inches(2.20), Inches(3.6), Inches(0.55),
                 [(price, 26, NAVY, True)])
        add_text(slide, x + Inches(0.2), Inches(2.78), Inches(3.6), Inches(0.28),
                 [(unit, 12, GOLD, True)])
        add_text(slide, x + Inches(0.2), Inches(3.16), Inches(3.6), Inches(0.42),
                 [(note, 13, INK, False)])
    add_rect(slide, Inches(0.48), Inches(3.98), Inches(12.38), Inches(1.42), NAVY, radius=0.06)
    add_text(slide, Inches(0.72), Inches(4.12), Inches(11.9), Inches(0.28),
             [("第一轮建议报价", 12, GOLD, True)])
    add_text(slide, Inches(0.72), Inches(4.44), Inches(11.9), Inches(0.78),
             [[("参访年度合作 15万—20万元。", 18, WHITE, True),
               ("与11月6日口头讨论的十几万至二十几万元相衔接，仍待审批。", 15, RGBColor(0xE7, 0xE2, 0xD6), False)],
              [("路演年度 14万元，作为第二项独立方案，不并入参访经费。", 15, RGBColor(0xE7, 0xE2, 0xD6), False)]],
             spacing=1.12)
    notes = [
        ("按实核定", "人数、企业接待费、专家、交通和场地单场确认，服务费不等于全部成本"),
        ("符合采购", "一页纸可以先共识；主体、发票、审批和结算走银行流程"),
        ("不付个人佣金", "不按贷款审批、开户金额或转介结果向个人支付报酬"),
    ]
    for i, (h, b) in enumerate(notes):
        x = Inches(0.48 + i * 4.18)
        add_rect(slide, x, Inches(5.58), Inches(4.02), Inches(1.38), PAPER, LINE, 0.6, 0.08)
        add_text(slide, x + Inches(0.18), Inches(5.70), Inches(3.66), Inches(0.30), [(h, 14, RED, True)])
        add_text(slide, x + Inches(0.18), Inches(6.04), Inches(3.66), Inches(0.76), [(b, 13, INK, False)])


def s08(slide, n, total):
    bg(slide)
    header(slide, "推进顺序", "先有一场能落地的活动，再谈年度框架",
           "三步都在2026年11月至12月内启动")
    footer(slide, n, total)
    steps = [
        ("01", "确认11月18日路演", "这是眼下最近的试点。核实活动是否按计划举行，茶歇、会务、场地支持是否获批，以及银行在现场的对接环节。"),
        ("02", "落实1—2个参访资源", "优先协调京东、美团、傅里叶和张江机器人。用一场20—30人的真实活动证明组织能力，而不是先交一份名单。"),
        ("03", "签订年度合作框架", "参访作为基础项目，路演作为独立扩展。写清预算、双方权益、承办主体和结算方式。"),
    ]
    for i, (num, title, body) in enumerate(steps):
        x = Inches(0.48 + i * 4.18)
        add_rect(slide, x, Inches(1.68), Inches(4.02), Inches(4.55), PAPER, LINE, 0.75, 0.06)
        add_rect(slide, x, Inches(1.68), Inches(4.02), Inches(0.10), GOLD if i < 2 else NAVY,
                 radius=0, shape=MSO_SHAPE.RECTANGLE)
        add_text(slide, x + Inches(0.22), Inches(1.95), Inches(3.5), Inches(0.50),
                 [(num, 28, GOLD, True)])
        add_text(slide, x + Inches(0.22), Inches(2.55), Inches(3.58), Inches(0.80),
                 [(title, 20, NAVY, True)])
        add_text(slide, x + Inches(0.22), Inches(3.50), Inches(3.58), Inches(2.30),
                 [(body, 14, INK, False)], spacing=1.2)
    add_rect(slide, Inches(0.48), Inches(6.40), Inches(12.38), Inches(0.58), GOLD_SOFT, radius=0.06)
    add_text(slide, Inches(0.70), Inches(6.40), Inches(12.0), Inches(0.58),
             [("参访资源要成为执行方的长期产品。每一次约访都进入年度服务，不单独做无偿引荐。", 14, NAVY, True)],
             anchor=MSO_ANCHOR.MIDDLE)


def s09(slide, n, total):
    bg(slide)
    header(slide, "标杆参访", "最先签约、最快形成收入的产品",
           "2027年原则上每月1场。排期按接待确认情况滚动调整")
    footer(slide, n, total)
    metrics = [("12场", "年度参访"), ("20—30人", "单场高管规模"), ("240—360", "年度服务人次"), ("半天", "标准活动时长")]
    for i, (n_, l) in enumerate(metrics):
        x = Inches(0.48 + i * 3.18)
        add_rect(slide, x, Inches(1.62), Inches(3.05), Inches(1.15), NAVY if i == 0 else PAPER,
                 None if i == 0 else LINE, 0.7, 0.08)
        add_text(slide, x + Inches(0.16), Inches(1.72), Inches(2.75), Inches(0.48),
                 [(n_, 22, WHITE if i == 0 else NAVY, True)])
        add_text(slide, x + Inches(0.16), Inches(2.24), Inches(2.75), Inches(0.36),
                 [(l, 13, GOLD if i == 0 else MUTED, False)])
    for i, (name, body) in enumerate(D.SERIES):
        x = Inches(0.48 + i * 4.18)
        add_rect(slide, x, Inches(2.96), Inches(4.02), Inches(2.05), PAPER, LINE, 0.7, 0.08)
        add_text(slide, x + Inches(0.18), Inches(3.10), Inches(3.66), Inches(0.28),
                 [(f"系列 {i+1}", 12, GOLD, True)])
        add_text(slide, x + Inches(0.18), Inches(3.40), Inches(3.66), Inches(0.55),
                 [(name, 16, NAVY, True)])
        add_text(slide, x + Inches(0.18), Inches(4.05), Inches(3.66), Inches(0.75),
                 [(body, 13, INK, False)])
    add_rect(slide, Inches(0.48), Inches(5.20), Inches(12.38), Inches(1.75), PAPER, LINE, 0.7, 0.06)
    add_text(slide, Inches(0.70), Inches(5.36), Inches(11.9), Inches(0.30),
             [("单场怎么结束", 13, GOLD, True)])
    add_text(slide, Inches(0.70), Inches(5.72), Inches(11.9), Inches(1.00),
             [[("参访结束后安排企业需求交流，并由银行客户经理做一对一服务。", 15, INK, False)],
              [("每场留下活动总结、需求清单和对接记录。开户、授信和资金业务的转化，由银行自行管理和评价。", 14, MUTED, False)]],
             spacing=1.15)


def s10(slide, n, total):
    bg(slide)
    header(slide, "参访资源 · A类", "七家重点推进，全部尚未确认接待",
           "A类指洽谈中已有线索或可协助约访，不是企业已经答应2027年档期")
    footer(slide, n, total)
    rows = [["序号", "拟参访单位", "核心主题", "资源状态"]]
    for v in D.VISITS:
        if v["cat"] != "A":
            continue
        rows.append([f"{v['no']:02d}", v["org"], v["theme"], v["status"]])
    add_table(
        slide, Inches(0.42), Inches(1.62), Inches(12.48), Inches(4.85),
        rows,
        [Inches(0.7), Inches(3.5), Inches(4.7), Inches(3.58)],
        font=12,
        aligns=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT],
        bold_col=1,
    )
    add_text(slide, Inches(0.48), Inches(6.58), Inches(12.3), Inches(0.45),
             [("优先落地顺序：京东、美团、傅里叶、张江机器人。魔力社区、魔丸熊、博展科技保留在清单内，按银行线索和档期替换插入。", 13, MUTED, False)])


def s11(slide, n, total):
    bg(slide)
    header(slide, "参访资源 · B类与C类", "历史经验可以再用，接待必须重谈",
           "B类是此前有参访经历的方向；C类是建议开发的产业，还没有落实接待单位")
    footer(slide, n, total)
    b = [v for v in D.VISITS if v["cat"] == "B"]
    for i, v in enumerate(b):
        col = 0 if i < 6 else 1
        row = i if i < 6 else i - 6
        x = Inches(0.48 + col * 6.40)
        y = 1.58 + row * 0.52
        add_rect(slide, x, Inches(y), Inches(6.22), Inches(0.46), PAPER, LINE, 0.5, 0.15)
        add_text(slide, x + Inches(0.12), Inches(y), Inches(0.48), Inches(0.46),
                 [(f"{v['no']:02d}", 12, GOLD, True)], anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, x + Inches(0.58), Inches(y), Inches(5.5), Inches(0.46),
                 [(v["org"] + "  ·  " + v["track"], 12, INK, False)], anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(0.48), Inches(4.78), Inches(6), Inches(0.28),
             [("C类 · 建议拓展", 13, GOLD, True)])
    for i, (name, kind, value, month) in enumerate(D.VISIT_DIRS):
        x = Inches(0.48 + (i % 3) * 4.18)
        y = 5.12 + (i // 3) * 0.90
        add_rect(slide, x, Inches(y), Inches(4.02), Inches(0.82), GOLD_SOFT, radius=0.08)
        add_text(slide, x + Inches(0.14), Inches(y + 0.06), Inches(3.74), Inches(0.30),
                 [(name, 13, NAVY, True)])
        add_text(slide, x + Inches(0.14), Inches(y + 0.38), Inches(3.74), Inches(0.36),
                 [(month + "  ·  " + value, 11, MUTED, False)])


def s12(slide, n, total):
    bg(slide)
    header(slide, "2027参访年历", "十二场储备方案，按月滚动",
           "6月、具体企业二选一的月份，以先确认接待的单位为准")
    footer(slide, n, total)
    for i, (month, theme, org) in enumerate(D.VISIT_CALENDAR):
        col, row = i % 4, i // 4
        x = Inches(0.48 + col * 3.18)
        y = Inches(1.58 + row * 1.78)
        add_rect(slide, x, y, Inches(3.05), Inches(1.66), PAPER, LINE, 0.7, 0.08)
        add_rect(slide, x, y, Inches(0.08), Inches(1.66), GOLD, radius=0, shape=MSO_SHAPE.RECTANGLE)
        add_text(slide, x + Inches(0.22), y + Inches(0.12), Inches(2.7), Inches(0.32),
                 [(month, 14, GOLD, True)])
        add_text(slide, x + Inches(0.22), y + Inches(0.46), Inches(2.72), Inches(0.62),
                 [(theme, 13, NAVY, True)])
        add_text(slide, x + Inches(0.22), y + Inches(1.12), Inches(2.7), Inches(0.40),
                 [(org, 12, MUTED, False)])


def s13(slide, n, total):
    bg(slide)
    header(slide, "参访怎么做", "标准半天流程，三方分工写清楚",
           "如需晚餐、园区走访或融资对接，另做半日或一日方案")
    footer(slide, n, total)
    add_text(slide, Inches(0.48), Inches(1.58), Inches(6.2), Inches(0.30),
             [("20—30人  ·  约半天", 14, NAVY, True)])
    for i, (tm, act) in enumerate(D.VISIT_FLOW):
        y = 1.98 + i * 0.72
        add_rect(slide, Inches(0.48), Inches(y), Inches(7.15), Inches(0.64), PAPER, LINE, 0.55, 0.1)
        add_text(slide, Inches(0.64), Inches(y), Inches(1.85), Inches(0.64),
                 [(tm, 13, GOLD, True)], anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(2.55), Inches(y), Inches(4.9), Inches(0.64),
                 [(act, 14, INK, False)], anchor=MSO_ANCHOR.MIDDLE)
    roles = [
        (NAVY, "泰隆银行", "客户筛选与邀约，客户经理配置，金融服务，采购和费用审批"),
        (RED, "联合会及执行方", "企业资源、方案、行程、现场组织和活动总结"),
        (GOLD, "复旦科技园等", "按项目参与邀约和园区协同。主办名称以授权文件为准"),
    ]
    for i, (color, title, body) in enumerate(roles):
        y = 1.98 + i * 1.45
        add_rect(slide, Inches(7.85), Inches(y), Inches(5.00), Inches(1.32), PAPER, LINE, 0.6, 0.08)
        add_rect(slide, Inches(7.85), Inches(y), Inches(0.08), Inches(1.32), color,
                 radius=0, shape=MSO_SHAPE.RECTANGLE)
        add_text(slide, Inches(8.12), Inches(y + 0.12), Inches(4.55), Inches(0.32),
                 [(title, 15, NAVY, True)])
        add_text(slide, Inches(8.12), Inches(y + 0.50), Inches(4.55), Inches(0.68),
                 [(body, 13, INK, False)])


def s14(slide, n, total):
    bg(slide)
    header(slide, "投融资路演", "4场重点产业路演 + 4场小型对接",
           "6月、9月、12月留给复访、项目推进和复盘，不强制加场")
    footer(slide, n, total)
    stats = [("8场", "全年活动"), ("32—40", "路演项目"), ("40—60家", "重点跟踪库")]
    for i, (a, b) in enumerate(stats):
        x = Inches(0.48 + i * 4.18)
        add_rect(slide, x, Inches(1.58), Inches(4.02), Inches(0.78), NAVY if i == 0 else PAPER,
                 None if i == 0 else LINE, 0.6, 0.1)
        add_text(slide, x + Inches(0.18), Inches(1.58), Inches(1.7), Inches(0.78),
                 [(a, 18, WHITE if i == 0 else NAVY, True)], anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, x + Inches(1.9), Inches(1.58), Inches(1.95), Inches(0.78),
                 [(b, 13, GOLD if i == 0 else MUTED, False)], anchor=MSO_ANCHOR.MIDDLE)
    rows = [["时间", "活动", "规模", "类型"]]
    for month, name, kind, scale, _focus in D.ROADSHOWS:
        rows.append([month, name, scale, "重点路演" if kind == "重点" else "小型对接"])
    add_table(
        slide, Inches(0.42), Inches(2.54), Inches(12.48), Inches(4.35),
        rows,
        [Inches(0.9), Inches(6.5), Inches(1.7), Inches(3.38)],
        font=12,
        aligns=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.CENTER],
        bold_col=1,
    )


def s15(slide, n, total):
    bg(slide)
    header(slide, "四场重点专题", "产业要窄，项目要有真实需求",
           "重点路演控制在每场5—6家企业，保证点评和对接时间")
    footer(slide, n, total)
    for i, t in enumerate(D.THEMES):
        col, row = i % 2, i // 2
        x = Inches(0.48 + col * 6.40)
        y = Inches(1.60 + row * 2.62)
        add_rect(slide, x, y, Inches(6.22), Inches(2.48), PAPER, LINE, 0.7, 0.06)
        add_rect(slide, x, y, Inches(0.08), Inches(2.48), GOLD if row == 0 else NAVY,
                 radius=0, shape=MSO_SHAPE.RECTANGLE)
        add_text(slide, x + Inches(0.28), y + Inches(0.12), Inches(5.75), Inches(0.28),
                 [(f"专题 {i+1}", 12, GOLD, True)])
        add_text(slide, x + Inches(0.28), y + Inches(0.40), Inches(5.75), Inches(0.36),
                 [(t["name"], 18, NAVY, True)])
        add_text(slide, x + Inches(0.28), y + Inches(0.86), Inches(5.75), Inches(0.70),
                 [(t["tracks"], 13, INK, False)])
        add_text(slide, x + Inches(0.28), y + Inches(1.60), Inches(5.75), Inches(0.72),
                 [(t["value"], 12, MUTED, False)])


def s16(slide, n, total):
    bg(slide)
    header(slide, "路演现场与项目库", "60—80人的重点场，和20—30人的闭门场分开设计",
           "企业是否符合授信或投资标准，由银行和投资机构独立决定")
    footer(slide, n, total)
    for i, (tm, act) in enumerate(D.ROADSHOW_FLOW):
        y = 1.58 + i * 0.62
        add_rect(slide, Inches(0.48), Inches(y), Inches(7.35), Inches(0.56), PAPER, LINE, 0.5, 0.12)
        add_text(slide, Inches(0.62), Inches(y), Inches(1.7), Inches(0.56),
                 [(tm, 12, GOLD, True)], anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(2.35), Inches(y), Inches(5.3), Inches(0.56),
                 [(act, 13, INK, False)], anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(8.05), Inches(1.58), Inches(4.8), Inches(0.30),
             [("三类项目库", 14, NAVY, True)])
    for i, (name, desc) in enumerate(D.POOLS):
        y = 2.00 + i * 1.15
        add_rect(slide, Inches(8.05), Inches(y), Inches(4.80), Inches(1.05), PAPER, LINE, 0.6, 0.08)
        add_text(slide, Inches(8.22), Inches(y + 0.08), Inches(4.5), Inches(0.30),
                 [(name, 14, NAVY, True)])
        add_text(slide, Inches(8.22), Inches(y + 0.42), Inches(4.5), Inches(0.52),
                 [(desc, 12, MUTED, False)])
    add_text(slide, Inches(8.05), Inches(5.55), Inches(4.8), Inches(1.35),
             [[("小型对接会", 13, GOLD, True)],
              [("3—4个项目，20—30人，闭门深谈。", 13, INK, False)],
              [("一般重点场安排5—6家路演，避免项目过多。", 13, INK, False)]],
             spacing=1.08)


def s17(slide, n, total):
    bg(slide)
    header(slide, "两套预算分开列", "参访和路演不重复列支",
           "若银行只设一个企业服务年度预算，需要重新分配，而不是把两套数字相加后直接报批")
    footer(slide, n, total)
    add_rect(slide, Inches(0.48), Inches(1.62), Inches(5.95), Inches(5.30), PAPER, LINE, 0.7, 0.06)
    add_text(slide, Inches(0.70), Inches(1.78), Inches(5.5), Inches(0.30),
             [("参访 · 第一轮主方案", 14, GOLD, True)])
    add_text(slide, Inches(0.70), Inches(2.15), Inches(5.5), Inches(0.55),
             [("15万—20万元 / 年", 28, NAVY, True)])
    lines = [
        "对应11月6日口头讨论的十几万至二十几万元，尚未批准",
        "5,000—10,000元/场只计策划、协调和组织执行服务费",
        "企业接待、交通、专家、场地按单场实际成本另核",
        "12场服务费区间约6万—12万元，年度建议高于此数，是因为包含全年资源维护和按实成本",
        "建议2026年12月先做1场试点，再转入常态化",
    ]
    add_text(slide, Inches(0.70), Inches(2.85), Inches(5.5), Inches(3.7),
             [[("·  " + t, 13, INK, False)] for t in lines], spacing=1.18)
    add_rect(slide, Inches(6.62), Inches(1.62), Inches(6.24), Inches(5.30), NAVY, radius=0.06)
    add_text(slide, Inches(6.86), Inches(1.78), Inches(5.8), Inches(0.28),
             [("路演 · 第二项独立方案", 14, GOLD, True)])
    add_text(slide, Inches(6.86), Inches(2.12), Inches(5.8), Inches(0.48),
             [("140,000元", 28, WHITE, True)])
    items = [
        ("重点产业路演 × 4", "80,000"),
        ("小型专项对接 × 4", "48,000"),
        ("年度协调与复盘 × 1", "12,000"),
    ]
    y = 2.75
    for name, money in items:
        add_text(slide, Inches(6.86), Inches(y), Inches(3.6), Inches(0.40),
                 [(name, 14, RGBColor(0xE7, 0xE2, 0xD6), False)], anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(10.30), Inches(y), Inches(2.2), Inches(0.40),
                 [(money, 16, WHITE, True)], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        y += 0.48
    add_rect(slide, Inches(6.86), Inches(y + 0.08), Inches(5.75), Pt(1), GOLD,
             radius=0, shape=MSO_SHAPE.RECTANGLE)
    add_text(slide, Inches(6.86), Inches(4.45), Inches(5.75), Inches(2.1),
             [[("含会务、场地、物料、执行、邀约、筛选和成果整理。", 13, RGBColor(0xE7, 0xE2, 0xD6), False)],
              [("园区若提供免费场地，银行采购额可以下调。", 13, RGBColor(0xE7, 0xE2, 0xD6), False)],
              [("税费口径以采购清单和合同为准。", 13, RGBColor(0xE7, 0xE2, 0xD6), False)]],
             spacing=1.15)


def s18(slide, n, total):
    bg(slide)
    header(slide, "执行节奏", "2026年11月到2027年第一季度",
           "绩效同时看有效企业、需求条数、银行对接、满意度和授权，不只看人数")
    footer(slide, n, total)
    for i, (when, task, result) in enumerate(D.TIMELINE):
        y = 1.60 + i * 1.02
        add_rect(slide, Inches(0.48), Inches(y), Inches(12.38), Inches(0.92), PAPER, LINE, 0.6, 0.08)
        add_rect(slide, Inches(0.48), Inches(y), Inches(0.08), Inches(0.92), GOLD if i < 3 else NAVY,
                 radius=0, shape=MSO_SHAPE.RECTANGLE)
        add_text(slide, Inches(0.74), Inches(y), Inches(3.15), Inches(0.92),
                 [(when, 13, NAVY, True)], anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(3.95), Inches(y + 0.10), Inches(4.3), Inches(0.72),
                 [(task, 15, INK, True)], anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(8.30), Inches(y + 0.10), Inches(4.35), Inches(0.72),
                 [(result, 13, MUTED, False)], anchor=MSO_ANCHOR.MIDDLE)


def s19(slide, n, total):
    bg(slide)
    header(slide, "绩效与边界", "记录服务过程，转化结果由银行评价",
           "客户信息的交换和使用，须取得适当授权")
    footer(slide, n, total)
    for i, (when, what) in enumerate(D.TRACK):
        x = Inches(0.48 + i * 4.18)
        add_rect(slide, x, Inches(1.60), Inches(4.02), Inches(1.35), PAPER, LINE, 0.7, 0.08)
        add_text(slide, x + Inches(0.18), Inches(1.72), Inches(3.66), Inches(0.30),
                 [(when, 14, GOLD, True)])
        add_text(slide, x + Inches(0.18), Inches(2.08), Inches(3.66), Inches(0.70),
                 [(what, 13, INK, False)])
    for i, item in enumerate(D.KPI):
        col, row = i % 4, i // 4
        x = Inches(0.48 + col * 3.18)
        y = 3.15 + row * 0.62
        add_rect(slide, x, Inches(y), Inches(3.05), Inches(0.54), NAVY_SOFT, radius=0.12)
        add_text(slide, x + Inches(0.12), Inches(y), Inches(2.82), Inches(0.54),
                 [(item, 12, NAVY, True)], anchor=MSO_ANCHOR.MIDDLE)
    bounds = [
        ("预算", "口头量级不等于批复"),
        ("接待", "名单不等于企业已同意"),
        ("品牌", "复旦及相关单位授权前不冠名"),
        ("合规", "不承诺入学，不对个人付结果佣金"),
    ]
    for i, (h, b) in enumerate(bounds):
        x = Inches(0.48 + i * 3.18)
        add_rect(slide, x, Inches(5.55), Inches(3.05), Inches(1.38), PAPER, LINE, 0.6, 0.08)
        add_text(slide, x + Inches(0.14), Inches(5.68), Inches(2.76), Inches(0.30),
                 [(h, 14, RED, True)])
        add_text(slide, x + Inches(0.14), Inches(6.04), Inches(2.76), Inches(0.70),
                 [(b, 13, INK, False)])


def s20(slide, n, total):
    bg(slide, NAVY)
    add_rect(slide, 0, 0, Inches(0.16), PAGE_H, GOLD, radius=0, shape=MSO_SHAPE.RECTANGLE)
    add_text(slide, Inches(0.70), Inches(0.38), Inches(10), Inches(0.28),
             [("本次沟通希望带走的决定", 13, GOLD, True)])
    add_text(slide, Inches(0.68), Inches(0.72), Inches(12), Inches(0.55),
             [("先定参访，再扩展路演、研修和出海", 28, WHITE, True)])
    asks = [
        "核实11月18日路演是否举行，以及茶歇、会务、场地支持是否获批",
        "确认参访年度预算按15万—20万元进入审批口径，路演14万元另案",
        "指定预算和采购负责人，明确承办主体、发票和结算",
        "在京东、美团、傅里叶、张江机器人中确定首场试点",
        "约定客户服务由银行负责，参访产品由执行方持续运营",
    ]
    for i, text in enumerate(asks):
        y = 1.55 + i * 0.72
        sp = add_rect(slide, Inches(0.72), Inches(y), Inches(0.46), Inches(0.46), GOLD, radius=0.5)
        shape_text(sp, [(str(i + 1), 14, NAVY, True)], align=PP_ALIGN.CENTER)
        add_text(slide, Inches(1.36), Inches(y), Inches(11.2), Inches(0.50),
                 [(text, 16, WHITE, False)], anchor=MSO_ANCHOR.MIDDLE)
    add_rect(slide, Inches(0.70), Inches(5.30), Inches(11.95), Inches(1.55), NAVY_MID, radius=0.06)
    add_text(slide, Inches(0.94), Inches(5.46), Inches(11.5), Inches(0.28),
             [("建议的经营顺序", 12, GOLD, True)])
    add_text(slide, Inches(0.94), Inches(5.82), Inches(11.5), Inches(0.80),
             [("先把企业参访做成稳定的年度服务，再延伸到投融资路演、复旦终身教育和企业出海。三份文件配套使用：本简报用于沟通，Word稿用于提交，Excel表用于排期、报价和会后跟踪。", 14, WHITE, False)],
             spacing=1.15)
    add_text(slide, Inches(11.2), Inches(7.05), Inches(1.6), Inches(0.28),
             [(f"{n:02d}  /  {total:02d}", 11, GOLD, True)], align=PP_ALIGN.RIGHT)


def build():
    slides = [s01, s02, s03, s04, s05, s06, s07, s08, s09, s10,
              s11, s12, s13, s14, s15, s16, s17, s18, s19, s20]
    prs = Presentation()
    prs.slide_width = PAGE_W
    prs.slide_height = PAGE_H
    prs.core_properties.title = D.TITLE
    prs.core_properties.subject = "商务沟通稿"
    prs.core_properties.author = "上海市杨浦区科技企业联合会"
    prs.core_properties.category = "洽谈稿"
    blank = prs.slide_layouts[6]
    total = len(slides)
    for i, fn in enumerate(slides, 1):
        slide = prs.slides.add_slide(blank)
        fn(slide, i, total)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    prs.save(OUT)
    print(OUT)
    print("slides", total)


if __name__ == "__main__":
    build()
