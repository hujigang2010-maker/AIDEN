"""Generate sponsorship agreement Word document.

Produces: 赞助方案/赞助协议-2026峰会.docx
"""
from __future__ import annotations

import os

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "赞助方案")
OUT_PATH = os.path.normpath(os.path.join(OUT_DIR, "赞助协议-2026峰会.docx"))


def set_cn_font(run, font_name: str = "宋体", size: int = 11, bold: bool | None = None,
                color: tuple[int, int, int] | None = None) -> None:
    run.font.name = font_name
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        from docx.oxml import OxmlElement
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color is not None:
        run.font.color.rgb = RGBColor(*color)


def add_heading(doc: Document, text: str, level: int = 1) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if level == 0:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        set_cn_font(run, "黑体", 22, bold=True)
    elif level == 1:
        run = p.add_run(text)
        set_cn_font(run, "黑体", 14, bold=True, color=(31, 73, 125))
    elif level == 2:
        run = p.add_run(text)
        set_cn_font(run, "黑体", 12, bold=True)
    else:
        run = p.add_run(text)
        set_cn_font(run, "黑体", 11, bold=True)


def add_para(doc: Document, text: str, *, bold: bool = False, size: int = 11,
             align: int = WD_ALIGN_PARAGRAPH.LEFT, indent_first: bool = True) -> None:
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    pf.space_after = Pt(2)
    if indent_first:
        pf.first_line_indent = Pt(size * 2)
    run = p.add_run(text)
    set_cn_font(run, "宋体", size, bold=bold)


def add_kv_table(doc: Document, rows: list[tuple[str, str]], *, col_widths=(4, 12)) -> None:
    table = doc.add_table(rows=len(rows), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for i, (k, v) in enumerate(rows):
        c0, c1 = table.rows[i].cells
        c0.width = Cm(col_widths[0])
        c1.width = Cm(col_widths[1])
        for cell, text, bold in ((c0, k, True), (c1, v, False)):
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            cell.paragraphs[0].text = ""
            run = cell.paragraphs[0].add_run(text)
            set_cn_font(run, "宋体", 11, bold=bold)


def add_grid_table(doc: Document, header: list[str], rows: list[list[str]]) -> None:
    table = doc.add_table(rows=1 + len(rows), cols=len(header))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(header):
        hdr_cells[i].paragraphs[0].text = ""
        run = hdr_cells[i].paragraphs[0].add_run(h)
        set_cn_font(run, "黑体", 11, bold=True, color=(255, 255, 255))
        # shading
        from docx.oxml import OxmlElement
        tc_pr = hdr_cells[i]._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "1F497D")
        tc_pr.append(shd)
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, text in enumerate(row):
            cell = table.rows[r_idx].cells[c_idx]
            cell.paragraphs[0].text = ""
            run = cell.paragraphs[0].add_run(text)
            set_cn_font(run, "宋体", 10)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def main() -> None:
    doc = Document()
    # default margins
    section = doc.sections[0]
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.4)
    section.right_margin = Cm(2.4)

    # ===== Cover-ish title =====
    add_heading(doc, "重构与突围", level=0)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cn_font(p.add_run("2026 人工智能商业化落地与硬核投资破局峰会"),
                "黑体", 16, bold=True, color=(31, 73, 125))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cn_font(p.add_run("赞助合作协议（Sponsorship Agreement）"),
                "黑体", 14, bold=True)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cn_font(p.add_run("协议编号：__________________   签订日期：______年______月______日"),
                "宋体", 11)

    doc.add_paragraph()

    # ===== Parties =====
    add_heading(doc, "甲乙双方信息", level=1)
    add_kv_table(doc, [
        ("甲方（主办方）", "北京大学经济学院上海校友会 / 复旦大学住房政策研究中心"),
        ("通讯地址", "上海市______区______路______号"),
        ("项目联系人", "________________   联系电话：________________"),
        ("电子邮箱", "________________"),
        ("乙方（赞助方）", "________________________________"),
        ("统一社会信用代码", "________________________________"),
        ("注册地址", "________________________________"),
        ("授权代表", "________________   职务：____________"),
        ("项目联系人", "________________   联系电话：________________"),
        ("电子邮箱", "________________"),
    ])

    doc.add_paragraph()
    add_para(doc, "鉴于甲方拟于 2026 年 5 月（具体日期以正式邀请函为准）在上海主办"
                  "“重构与突围——2026 人工智能商业化落地与硬核投资破局峰会”"
                  "（以下简称“本次峰会”），乙方愿意成为本次峰会的赞助合作方，"
                  "双方本着平等、自愿、诚实信用的原则，依据《中华人民共和国民法典》"
                  "等相关法律法规，就乙方赞助本次峰会事宜达成如下协议，以资共同遵守。")

    # ===== Articles =====
    add_heading(doc, "第一条  合作内容与赞助级别", level=1)
    add_para(doc, "1.1 经双方协商一致，乙方选择以下赞助级别（请在所选级别前打“√”）：")

    add_grid_table(doc,
                   header=["选择", "级别", "投入金额（人民币）", "名额", "核心权益概述"],
                   rows=[
                       ["□", "独家总冠名", "面议", "1 席", "大会冠名 + 15 分钟独立主旨演讲 + 圆桌定制席 + 晚宴主桌（3 人）+ 全域霸屏"],
                       ["□", "钻石赞助", "￥50,000", "3 席", "圆桌论坛对话席 + 年度颁奖授牌 + VIP 晚宴入场券 5 张"],
                       ["□", "铂金赞助", "￥30,000", "5 席", "音乐会环节鸣谢 + 专属展位 + VIP 晚宴入场券 3 张"],
                       ["□", "黄金赞助", "￥10,000", "8 席", "官网/大屏滚动展示 + 手册内页广告 + VIP 晚宴入场券 1 张"],
                       ["□", "资源置换合伙人", "实物/服务", "限定", "晚宴用酒/伴手礼/影像/出行/媒体（详见附件一）"],
                       ["□", "基础品牌曝光", "￥1,500/位", "不限", "核心动线易拉宝广告位 1 个"],
                   ])

    doc.add_paragraph()
    add_para(doc, "1.2 乙方本次赞助投入：人民币（大写）______________________元整（小写：￥__________元）。")
    add_para(doc, "1.3 资源置换标的（如适用）：__________________________________________________________。")
    add_para(doc, "1.4 双方就乙方专属定制权益的具体清单与执行细则，详见本协议《附件一·赞助权益清单》。"
                  "附件一与本协议正文具有同等法律效力。")

    add_heading(doc, "第二条  甲方义务", level=1)
    for t in [
        "2.1 甲方应按照《附件一·赞助权益清单》载明的权益项、规格、数量与时间节点，"
        "保质保量交付乙方所享有的全部权益，包括但不限于品牌曝光、议程参与、"
        "晚宴接待、嘉宾对接、白皮书署名等。",
        "2.2 甲方应在签约后 5 个工作日内向乙方提供物料规格清单（Logo、KV、宣传片等"
        "技术参数与提交规范），并指定专人对接物料交付与上线。",
        "2.3 甲方应在峰会结束后 7 个自然日内向乙方出具《赞助权益执行回执》，"
        "并附现场图片、媒体链接、嘉宾合影等证明材料。",
        "2.4 甲方应妥善保管乙方提供的物料、商标及其他商业资料，未经乙方书面同意，"
        "不得用于本次峰会之外的任何用途。",
    ]:
        add_para(doc, t)

    add_heading(doc, "第三条  乙方义务", level=1)
    for t in [
        "3.1 乙方应按本协议第四条约定的时间与方式足额支付赞助款项，或按时交付资源置换标的。",
        "3.2 乙方应在签约后 5 个工作日内向甲方交付权益落地所需物料，"
        "包括但不限于：矢量 Logo（AI/EPS）、品牌简介（不超过 200 字）、宣传片"
        "（不超过 60 秒，1080P 以上 MP4）、易拉宝 / 展位 KV 设计稿等。",
        "3.3 乙方保证其提供的所有物料、文字、商标、视频不侵犯任何第三方知识产权"
        "及其他合法权益，否则由此产生的一切法律责任由乙方自行承担，"
        "并赔偿甲方因此遭受的全部损失。",
        "3.4 乙方应遵守峰会现场及晚宴的相关秩序与流程安排，配合甲方完成议程互动。",
    ]:
        add_para(doc, t)

    add_heading(doc, "第四条  款项支付与发票", level=1)
    for t in [
        "4.1 乙方应于本协议签订之日起 5 个工作日内，将全部赞助款项一次性汇入甲方指定账户：",
    ]:
        add_para(doc, t)
    add_kv_table(doc, [
        ("收款单位", "________________________________"),
        ("开户银行", "________________________________"),
        ("银行账号", "________________________________"),
        ("税号", "________________________________"),
    ])
    add_para(doc, "4.2 甲方在款项到账后 10 个工作日内向乙方开具等额合法有效的"
                  "增值税普通发票 / 专用发票（发票内容：会议服务费 / 赞助费）。")
    add_para(doc, "4.3 资源置换标的之交付时间、规格、验收标准详见《附件一》，"
                  "乙方逾期未交付的，视为放弃相应权益，甲方有权调整或回收。")

    add_heading(doc, "第五条  知识产权与品牌使用", level=1)
    for t in [
        "5.1 乙方授权甲方在本次峰会的宣传、报道、白皮书、官网、大屏、"
        "议程手册、媒体通稿、回顾视频等场景中使用乙方商标与企业简介，"
        "授权范围限于本次峰会及其衍生宣传内容，授权期限为本协议签订之日"
        "起至大会结束后 12 个月。",
        "5.2 甲方拥有本次峰会名称、视觉系统及衍生品的完整知识产权。"
        "乙方在使用大会名称、Logo、嘉宾形象等元素进行二次宣传前，"
        "应取得甲方书面同意，并不得作出有损大会形象或误导性的表述。",
        "5.3 双方对在合作过程中接触到的对方商业秘密、客户名录、"
        "嘉宾联系方式等信息负有保密义务，保密期限自本协议签订之日起 3 年。",
    ]:
        add_para(doc, t)

    add_heading(doc, "第六条  违约责任", level=1)
    for t in [
        "6.1 乙方未按约定时间支付赞助款项的，每逾期一日按未付款项的 0.5% 向甲方支付违约金；"
        "逾期超过 10 日的，甲方有权解除本协议并不退还任何已收款项。",
        "6.2 甲方未按《附件一》约定提供权益且无正当理由的，应按未履行权益对应价值的 100%"
        "向乙方退款或在后续合作中以同等价值补偿。",
        "6.3 因不可抗力（包括但不限于自然灾害、政府管制、公共卫生事件等）"
        "导致大会延期或取消的，双方互不承担违约责任，已付款项可顺延至甲方"
        "下一届同类活动或按未执行权益比例退还。",
    ]:
        add_para(doc, t)

    add_heading(doc, "第七条  争议解决", level=1)
    add_para(doc, "7.1 因本协议引起的或与本协议有关的任何争议，双方应首先友好协商解决；"
                  "协商不成的，任何一方均有权向甲方所在地人民法院提起诉讼。")
    add_para(doc, "7.2 本协议适用中华人民共和国法律。")

    add_heading(doc, "第八条  其他", level=1)
    for t in [
        "8.1 本协议自双方加盖公章（合同章）并由授权代表签字之日起生效。",
        "8.2 本协议一式肆份，甲乙双方各执贰份，具有同等法律效力。",
        "8.3 本协议未尽事宜，双方可另行签订书面补充协议，补充协议与本协议具有同等法律效力。",
    ]:
        add_para(doc, t)

    doc.add_paragraph()
    doc.add_paragraph()

    # ===== Signature block =====
    add_heading(doc, "签署页", level=1)
    sig = doc.add_table(rows=6, cols=2)
    sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig.style = "Table Grid"
    sig_rows = [
        ("甲方（盖章）：", "乙方（盖章）："),
        ("", ""),
        ("授权代表（签字）：", "授权代表（签字）："),
        ("", ""),
        ("日期：    年    月    日", "日期：    年    月    日"),
        ("", ""),
    ]
    for i, (a, b) in enumerate(sig_rows):
        for cell, text in zip(sig.rows[i].cells, (a, b)):
            cell.paragraphs[0].text = ""
            run = cell.paragraphs[0].add_run(text)
            set_cn_font(run, "宋体", 11, bold=True if i in (0, 2) else False)
            cell.width = Cm(8)

    doc.add_page_break()

    # ===== Annex 1: Rights list =====
    add_heading(doc, "附件一  赞助权益清单（Sponsorship Rights List）", level=1)
    add_para(doc, "本附件用于明确乙方在本次峰会中享有的具体权益。请甲乙双方共同确认后于每项后打勾或补充。",
             indent_first=False)

    add_grid_table(doc,
                   header=["序号", "权益项", "规格 / 描述", "数量", "确认"],
                   rows=[
                       ["1", "主背景板 Logo 展示", "等级：□总冠名 □钻石 □铂金 □黄金", "1", "□"],
                       ["2", "核心动线易拉宝", "尺寸 80×200cm，物料：□甲方代制 □乙方提供", "____ 个", "□"],
                       ["3", "议程手册广告", "□扉页整版 □整版 □半版 □1/4 版 □尾页鸣谢", "1", "□"],
                       ["4", "白皮书署名", "□封面联合 □扉页整版 □内页 1/2 □内页 1/4", "—", "□"],
                       ["5", "官网 / 大屏滚动展示", "时段：大会前 30 日至大会结束后 30 日", "—", "□"],
                       ["6", "主旨演讲 / 圆桌席位", "□15min 主旨演讲 □AI 硬核圆桌 □投资大圆桌", "1 人", "□"],
                       ["7", "年度颁奖授牌", "颁奖典礼现场领奖 + 媒体合影", "1", "□"],
                       ["8", "VIP 闭门晚宴入场券", "□主桌 3 人 □5 张 □3 张 □1 张", "____ 张", "□"],
                       ["9", "嘉宾 1V1 闭门沟通", "组委会定向引荐重量级嘉宾", "____ 位", "□"],
                       ["10", "媒体通稿露出", "□标题级 □正文重点 □鸣谢", "—", "□"],
                       ["11", "双校友产业联盟入册", "永久入册，享后续活动优先合作权", "—", "□"],
                       ["12", "其他定制权益", "________________________________", "—", "□"],
                   ])

    doc.add_paragraph()
    add_para(doc, "本附件经双方签字盖章后与协议正文具有同等法律效力。", indent_first=False)

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    doc.save(OUT_PATH)
    print(f"[OK] Word saved -> {OUT_PATH}")


if __name__ == "__main__":
    main()
