#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 2026 年 10 月 31 日王德峰讲座的定价综合结论（内部 Word）。"""

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output"

NAVY = RGBColor(0x1A, 0x3A, 0x5C)
NAVY_HEX = "1A3A5C"
GOLD = RGBColor(0x8B, 0x69, 0x14)
INK = RGBColor(0x33, 0x33, 0x33)
MUTED = RGBColor(0x66, 0x66, 0x66)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED = RGBColor(0x8B, 0x1E, 0x1E)
RED_HEX = "8B1E1E"
CREAM_HEX = "F7F3EA"
SOFT_HEX = "EEF3F8"
AMBER_HEX = "F8F1E3"
GREEN_HEX = "E8F2EA"


def set_run_font(run, *, name="微软雅黑", size=11, bold=False, color=None):
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:ascii"), name)
    run._element.rPr.rFonts.set(qn("w:hAnsi"), name)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    if color is not None:
        run.font.color.rgb = color


def shade_cell(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def set_table_widths(table, widths_cm):
    table.autofit = False
    table.allow_autofit = False
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(int(sum(widths_cm) * 567)))
    tbl_w.set(qn("w:type"), "dxa")
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            cell.width = Cm(widths_cm[idx])
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(int(widths_cm[idx] * 567)))
            tc_w.set(qn("w:type"), "dxa")
            v_align = OxmlElement("w:vAlign")
            v_align.set(qn("w:val"), "center")
            tc_pr.append(v_align)


def prevent_row_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    tr_pr.append(OxmlElement("w:cantSplit"))


def add_page_number(paragraph):
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)
    set_run_font(run, size=9, color=MUTED)


def setup_section(doc):
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.1)
    section.right_margin = Cm(2.1)
    section.header_distance = Cm(0.8)
    section.footer_distance = Cm(0.8)

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = hp.add_run("内部评估  ·  不直接发给学员、赞助商或合作方")
    set_run_font(run, size=9, color=RED, bold=True)
    p_pr = hp._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), RED_HEX)
    pbdr.append(bottom)
    p_pr.append(pbdr)

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    left = fp.add_run("10月31日讲座定价综合结论  ·  第 ")
    set_run_font(left, size=9, color=MUTED)
    add_page_number(fp)
    right = fp.add_run(" 页")
    set_run_font(right, size=9, color=MUTED)


def add_para(
    doc,
    text,
    *,
    size=11,
    bold=False,
    color=INK,
    align=None,
    space_before=0,
    space_after=6,
    first_line=False,
):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = 1.25
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    if first_line:
        pf.first_line_indent = Cm(0.74)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    return p


def add_h1(doc, text):
    return add_para(doc, text, size=14, bold=True, color=NAVY, space_before=12, space_after=6)


def add_h2(doc, text):
    return add_para(doc, text, size=12, bold=True, color=GOLD, space_before=8, space_after=4)


def set_cell_text(cell, text, *, bold=False, size=9.5, color=INK, fill=None, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.05
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    if fill:
        shade_cell(cell, fill)


def add_table(doc, headers, rows, widths):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for i, header in enumerate(headers):
        set_cell_text(
            table.rows[0].cells[i],
            header,
            bold=True,
            size=9,
            color=WHITE,
            fill=NAVY_HEX,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        )
    for r, row in enumerate(rows):
        fill = "F5F7FA" if r % 2 else "FFFFFF"
        for c, value in enumerate(row):
            set_cell_text(table.rows[r + 1].cells[c], value, size=9, fill=fill)
    set_table_widths(table, widths)
    for row in table.rows:
        prevent_row_split(row)
    add_para(doc, "", size=4, space_after=2)
    return table


def build():
    doc = Document()
    setup_section(doc)

    add_para(
        doc,
        "2026年10月31日讲座",
        size=12,
        bold=True,
        color=GOLD,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=2,
    )
    add_para(
        doc,
        "定价综合结论",
        size=22,
        bold=True,
        color=NAVY,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_before=0,
        space_after=2,
    )
    add_para(
        doc,
        "《传习录》与阳明心学  ·  一场上午专题",
        size=11,
        color=MUTED,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=2,
    )
    add_para(
        doc,
        "评估日期：2026年10月10日  ·  本页取代同日较早的“公开价 499 元”版本",
        size=10,
        color=MUTED,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=8,
    )

    add_para(doc, "执行结论", size=14, bold=True, color=NAVY, space_before=2, space_after=4)
    add_para(
        doc,
        "海报标准价 699 元/人。报名页另设早鸟 599 元，10月20日 24:00 前，限 40 张。企业前区席 1,280 元，只通过上海市杨浦区科技企业联合会邀约，不印上主海报。",
        size=12,
        bold=True,
        color=RED,
        space_after=6,
    )
    add_para(
        doc,
        "本场卖的是 10 月 31 日 09:00—12:00 的一场讲座。699 元取公开听讲价格带的上沿，并配一份纸质讲义和散场茶歇。1,280 元只卖座位差异，卖给可以走培训或商务费用的企业人员。签名书、合影和闭门午宴本次不上架。",
        first_line=True,
    )

    add_h1(doc, "一、五方意见如何收成这一个价")
    add_para(
        doc,
        "五方给出的主力价分别是 499 元、480–680 元、580–780 元、699–899 元、1,680–1,980 元。这些数字不能相加再平均。下表只记录每一条意见里被本次执行价吸收的部分。",
        first_line=True,
    )
    add_table(
        doc,
        ["意见", "原建议", "本次怎么用"],
        [
            ["Cursor 上一版", "公开价 499 元", "保留为退路：只有听讲、师资按分成、现金成本约 5 万元时，改回 499 元"],
            ["DeepSeek", "主力 480–680 元；引流可到 199–299 元", "公开价取这一档上沿。199–299 元只适用于把本场当长课入口"],
            ["Gemini", "标准 580–780 元；企业席 1,280–1,680 元", "标准价定为 699 元。企业前区定为 1,280 元。闭门午宴不上架"],
            ["Muse", "早鸟 699、标准 899、VIP 1,599", "讲义和茶歇写入交付。标准价放在 699：105 人 × 899 元 = 94,395 元，低于其假设的 10 万元成本下限"],
            ["Grok", "标准 1,680–1,980 元", "不作为公开标准价。3 小时公开讲座不按全日研修出售"],
        ],
        [3.2, 5.4, 7.8],
    )
    add_para(
        doc,
        "Muse 的成本区间 10–18 万元、Grok 的学术出场费 3–5 万元、Gemini 的课酬 10–15 万元、DeepSeek 的总成本 17–45 万元，都是估计，不是首乾书院或场地的报价。正式定价所用的成本，以书面报价替换这些区间。",
        first_line=True,
    )

    add_h1(doc, "二、只按座位分档")
    add_para(
        doc,
        "活动时间就是 10 月 31 日 09:00—12:00，没有另一场可以叫作“全天”的内容。海报精简稿里的“（上）”和“（下）”并成一段课程介绍，仍放在这一个上午里。席位差异只来自购买时间和座位区域。",
        first_line=True,
    )
    add_table(
        doc,
        ["席位", "价格", "出现位置", "权益"],
        [
            ["早鸟席", "599 元/人", "报名页；10月20日 24:00 前；限 40 张", "与标准席相同"],
            ["标准席", "699 元/人", "海报、公众号、朋友圈、报名页", "上午听讲、纸质讲义、散场茶歇"],
            ["企业前区席", "1,280 元/人", "联合会邀请函、报名页", "前区座位，其余与标准席相同"],
        ],
        [2.8, 2.8, 5.6, 5.2],
    )
    add_para(
        doc,
        "699 元的最低交付：入场发放一份纸质讲义；课程 12:00 结束，茶歇留到 12:30。场地合同若 12:00 必须清场，茶歇改为 08:30—09:00 入场茶歇。讲义和茶歇都写不进承办清单时，标准价改回 499 元，企业前区席同时撤下。",
        first_line=True,
    )
    add_para(
        doc,
        "本次不上架的项目：半天专题席，1,599 元 VIP，2,480 元以上互动席，3,800 元闭门午宴。签名书、合影、午宴都还没有王德峰教授或权利方的书面授权。企业前区席还要求场地能分出明确的前区；会场是平层、没有前区差别时，这一档不开售。",
        first_line=True,
    )

    add_h1(doc, "三、699 元对应哪一档成本")
    add_para(
        doc,
        "下表的票面保本价 = 假设总成本 ÷ 到场人数，还没扣除支付手续费和退款。它用来看价格和成本假设是否同档，不能当成已核实的盈亏点。",
        first_line=True,
    )
    add_table(
        doc,
        ["假设总成本", "100 人", "150 人", "200 人"],
        [
            ["5 万元", "500 元", "333 元", "250 元"],
            ["10 万元", "1,000 元", "667 元", "500 元"],
            ["18 万元", "1,800 元", "1,200 元", "900 元"],
            ["25 万元", "2,500 元", "1,667 元", "1,250 元"],
        ],
        [4.0, 4.0, 4.0, 4.4],
    )
    add_para(
        doc,
        "只卖 699 元标准席：105 人票面 73,395 元；150 人票面 104,850 元，按 94% 实收为 98,559 元。覆盖 10 万元票面大约需要 144 人，扣掉约 6% 后大约需要 153 人。105 人 × 899 元 = 94,395 元，到不了 10 万元，所以标准价不放到 899 元。",
        first_line=True,
    )
    add_para(
        doc,
        "内部销量目标：标准席 120 张 × 699 元 = 83,880 元，企业前区 25 张 × 1,280 元 = 32,000 元，合计票面 115,880 元，按 94% 约 108,927 元。这组目标对齐的是“总成本约 10 万元”这一档估计。总成本若到 18 万元或更高，差额用赞助或师资分成补上，标准价保持 699 元。",
        first_line=True,
    )
    add_para(
        doc,
        "收入参照：国家统计局 2026 年 7 月 15 日发布，2026 年上半年城镇居民人均可支配收入中位数 26,389 元，月均约 4,398 元。699 元约占 15.9%，1,280 元约占 29.1%，1,680 元约占 38.2%。1,280 元面向企业报销，不作为朋友圈主价格。",
        first_line=True,
    )

    add_h1(doc, "四、已经按 2,399 元承诺的人")
    add_para(
        doc,
        "内部材料里的 12 人、每人 2,399 元是测算口径，不是本结论确认的到账名单。2,399 元接近第三方招生页上系统课每个授课日约 2,475 元的锚点，不作为本场公开价。",
        first_line=True,
    )
    add_table(
        doc,
        ["状态", "收费", "权益写法"],
        [
            ["尚无报价、尚未收款", "699 元，或符合条件的 599 / 1,280 元", "按对应席位"],
            ["口头提过 2,399 元，没有书面承诺，也未收款", "改按标准席 699 元重新确认", "上午听讲、讲义、茶歇"],
            ["已有书面承诺或已收款", "维持 2,399 元", "入场，加上当时写明的额外项目"],
        ],
        [5.4, 5.2, 5.8],
    )

    add_h1(doc, "五、复旦署名先由校方确认")
    add_para(
        doc,
        "海报若把“复旦大学住房政策研究中心”印成联合主办，票款又不进入复旦账户，开售前由该中心按《普通高等学校举办非学历教育管理规定（试行）》（教职成厅函〔2021〕23号）确认两件事：这个名称能否用于面向社会的收费研修；票款由谁收取。",
        first_line=True,
    )
    add_para(
        doc,
        "该规定第九条写明：校内非实体性质的单位、职能管理部门、群团组织及教职员工个人不得以高校名义举办非学历教育。第二十二条写明：非学历教育办学所有收入纳入学校预算，统一核算；严禁合作方以任何名义收取费用。本结论不认定本场已经适用该规定。",
        first_line=True,
    )
    add_para(
        doc,
        "学校若认定本场不是复旦的办学项目，海报署名改为学术支持，联合主办改为首乾书院与上海市杨浦区科技企业联合会。学校若认定这是本校非学历教育项目，则按学校的立项、审批和财务规定办理。",
        first_line=True,
    )

    add_h1(doc, "六、可以直接贴进物料的句子")
    add_h2(doc, "海报票务区")
    add_para(doc, "标准席  699元/人", size=12, bold=True, color=NAVY, space_after=2)
    add_para(doc, "10月31日上午  09:00—12:00", space_after=2)
    add_para(doc, "含纸质讲义与茶歇", space_after=2)
    add_para(doc, "扫码报名", space_after=2)
    add_para(doc, "课程安排及退费规则，以报名页面为准。", size=10, color=MUTED)
    add_h2(doc, "报名页")
    add_para(
        doc,
        "早鸟席 599 元/人，10 月 20 日 24:00 前，限 40 张，权益与标准席相同。标准席 699 元/人，含 10 月 31 日上午讲座、纸质讲义、茶歇。企业前区席 1,280 元/人，含前区座位，其余与标准席相同；场地不能区分前区时不开售。建议退费：10 月 24 日（含）前全额退款；10 月 25 日至 29 日退还票价的 50%；10 月 30 日起可以转让，不退款。",
        first_line=True,
    )
    add_h2(doc, "朋友圈文案末尾")
    add_para(doc, "标准席 699 元/人，活动报名及详情请见海报二维码。", bold=True)

    add_h1(doc, "七、开售前核对")
    add_para(doc, "1. 讲义印制和茶歇时段写入承办清单。写不进去，标准价改回 499 元。", space_after=3)
    add_para(doc, "2. 用师资和场地的书面报价替换本文的成本估计。", space_after=3)
    add_para(doc, "3. 三家机构确认名称和标识；复旦署名按第五节办理。", space_after=3)
    add_para(doc, "4. 王德峰教授授权使用姓名、肖像和本场课程信息。签名书、合影、午宴未授权前不上架。", space_after=8)
    add_para(
        doc,
        "本结论用于内部定价。对外票价、权益和退费，以报名页及主办方正式通知为准。",
        size=10,
        color=MUTED,
    )

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "10月31日讲座_每人收费评估.docx"
    doc.save(path)
    return path


if __name__ == "__main__":
    print(build())
