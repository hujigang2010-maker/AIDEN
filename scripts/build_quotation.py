"""Generate a single-page-style service quotation organized by the 4 modules
the organizer提出 (按甲方反馈大纲，不拆解子项，便于一次性盖章与整体验收)。"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


GREEN = RGBColor(0x00, 0x6B, 0x3F)
DARK_GREEN = RGBColor(0x00, 0x4A, 0x2C)
GOLD = RGBColor(0xC8, 0xA2, 0x5B)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
GREY = RGBColor(0x55, 0x55, 0x55)
LIGHT_GREEN = RGBColor(0xE8, 0xF1, 0xEC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def _force_font(run, font="宋体"):
    rpr = run._element.get_or_add_rPr()
    rfonts = OxmlElement("w:rFonts")
    rfonts.set(qn("w:eastAsia"), font)
    rfonts.set(qn("w:ascii"), font)
    rfonts.set(qn("w:hAnsi"), font)
    rpr.append(rfonts)


def set_default_font(doc, name="宋体"):
    style = doc.styles["Normal"]
    style.font.name = name
    style.font.size = Pt(11)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), name)
    rfonts.set(qn("w:ascii"), name)
    rfonts.set(qn("w:hAnsi"), name)


def heading(doc, text, size=20, color=DARK_GREEN, center=True, font="黑体", bold=True):
    p = doc.add_paragraph()
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.color.rgb = color
    _force_font(run, font=font)
    return p


def para(doc, text, bold=False, size=11, color=DARK, align=None,
         first_indent=True, font="宋体", line_spacing=1.4):
    p = doc.add_paragraph()
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.line_spacing = line_spacing
    if first_indent:
        p.paragraph_format.first_line_indent = Pt(size * 2)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.bold = bold
    _force_font(run, font=font)
    return p


def blank(doc, n=1):
    for _ in range(n):
        doc.add_paragraph()


def shade(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tcPr.append(shd)


def cell_text(cell, text, bold=False, color=DARK, size=10.5,
              align_center=False, font="宋体", line_spacing=1.3):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    cell.text = ""
    p = cell.paragraphs[0]
    if align_center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = line_spacing
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.bold = bold
    _force_font(run, font=font)


def kv_table(doc, rows, col_widths=(3.8, 12.5)):
    t = doc.add_table(rows=len(rows), cols=2)
    t.style = "Table Grid"
    t.autofit = False
    for row in t.rows:
        row.cells[0].width = Cm(col_widths[0])
        row.cells[1].width = Cm(col_widths[1])
    for i, (k, v) in enumerate(rows):
        cell_text(t.rows[i].cells[0], k, bold=True, size=10.5)
        shade(t.rows[i].cells[0], "E8F1EC")
        cell_text(t.rows[i].cells[1], v, size=10.5)
    return t


def build():
    doc = Document()
    set_default_font(doc)

    section = doc.sections[0]
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.4)
    section.right_margin = Cm(2.4)

    # ===== 抬头 =====
    heading(doc, "服务报价单", size=24, color=DARK_GREEN)
    heading(doc,
            "（绿城中国 | 绿城·潮鸣外滩 · 晚宴冠名战略合作伙伴 · 10 万元）",
            size=13, color=GREY, bold=False)
    para(doc, "Service Quotation — Dinner Title Strategic Partnership",
         align="center", size=10, color=GREY, first_indent=False)
    blank(doc)

    para(doc, "报价单编号：QT-2026-GTC-潮鸣外滩-001       报价日期：2026 年 5 月 ____ 日       有效期：本报价单签发之日起 10 个自然日内有效",
         first_indent=False, size=10.5)
    blank(doc)

    # ===== 双方信息 =====
    para(doc, "一、双方信息", bold=True, first_indent=False, size=12, color=GREEN)
    kv_table(doc, rows=[
        ["报价方（甲方）", "北京大学经济学院上海校友会 / 复旦大学住房政策研究中心（联合主办）"],
        ["甲方联系人 / 电话", "________________ / ________________"],
        ["受报价方（乙方）", "绿城中国控股有限公司 / ____________________________（项目主体）"],
        ["项目品牌", "绿城·潮鸣外滩"],
        ["乙方联系人 / 电话", "Aiden（绿城方总对接） / ________________"],
    ])
    blank(doc)

    # ===== 合作活动 =====
    para(doc, "二、合作活动与合作身份", bold=True, first_indent=False, size=12, color=GREEN)
    kv_table(doc, rows=[
        ["合作活动", "重构与突围 · 2026 人工智能商业化落地与硬核投资破局峰会"],
        ["时间 / 地点", "2026 年 5 月 22 日 · 上海·北外滩·一滴水"],
        ["现场规模", "500+ 位高净值嘉宾"],
        ["合作身份", "晚宴冠名战略合作伙伴（Dinner Title Strategic Partner）"],
    ])
    blank(doc)

    # ===== 报价金额 =====
    para(doc, "三、服务费用合计", bold=True, first_indent=False, size=12, color=GREEN)
    t = doc.add_table(rows=1, cols=2)
    t.style = "Table Grid"
    t.autofit = False
    t.rows[0].cells[0].width = Cm(8.0)
    t.rows[0].cells[1].width = Cm(8.5)
    cell_text(t.rows[0].cells[0], "人民币壹拾万元整（含税）",
              bold=True, color=WHITE, size=14, align_center=True)
    shade(t.rows[0].cells[0], "006B3F")
    cell_text(t.rows[0].cells[1], "￥100,000.00 元",
              bold=True, color=WHITE, size=16, align_center=True)
    shade(t.rows[0].cells[1], "006B3F")
    blank(doc)

    # ===== 服务内容（4 大模块，不再拆解子项） =====
    para(doc, "四、服务内容（按 4 大模块整体打包，不再拆解子项以利整体验收）",
         bold=True, first_indent=False, size=12, color=GREEN)

    rows = [
        ["1", "晚宴冠名和专属权益",
         "晚宴主题冠名、专场 PPT 宣讲、视频轮播、主持人口播、席卡等物料植入",
         "￥55,000.00"],
        ["2", "主会场权益",
         "1 号位展台、项目物料植入参会嘉宾手拎袋、视频轮播奖项颁发",
         "￥30,000.00"],
        ["3", "专场项目参观邀约",
         "论坛/颁奖结束后主持人鸣谢口播 + 现场动线引导，定向邀约意向嘉宾前往项目案场",
         "￥5,000.00"],
        ["4", "宣发配合",
         "现场摄影、朋友圈九宫格、回顾视频项目鸣谢、执行回执（图片 + 合影 + 执行报告）",
         "￥10,000.00"],
    ]

    body = doc.add_table(rows=1 + len(rows) + 1, cols=4)
    body.style = "Table Grid"
    body.autofit = False
    widths = [1.2, 4.0, 8.3, 3.0]
    for i, w in enumerate(widths):
        for row in body.rows:
            row.cells[i].width = Cm(w)
    headers = ["序号", "服务模块", "服务内容描述", "金额（元）"]
    for i, h in enumerate(headers):
        cell_text(body.rows[0].cells[i], h, bold=True, color=WHITE,
                  size=11, align_center=True)
        shade(body.rows[0].cells[i], "006B3F")
    for r, row in enumerate(rows):
        cells = body.rows[r + 1].cells
        cell_text(cells[0], row[0], bold=True, size=11, align_center=True)
        cell_text(cells[1], row[1], bold=True, size=11)
        cell_text(cells[2], row[2], size=10.5)
        cell_text(cells[3], row[3], bold=True, size=11, align_center=True)
        body.rows[r + 1].height = Cm(1.4)
    total_cells = body.rows[-1].cells
    cell_text(total_cells[0], "合计", bold=True, color=WHITE, size=12, align_center=True)
    cell_text(total_cells[1], "人民币壹拾万元整（含税）", bold=True, color=WHITE, size=12)
    cell_text(total_cells[2], "—", color=WHITE, align_center=True)
    cell_text(total_cells[3], "￥100,000.00", bold=True, color=WHITE, size=13, align_center=True)
    for ci in range(4):
        shade(total_cells[ci], "006B3F")
    blank(doc)

    # ===== 验收方式 =====
    para(doc, "五、验收方式（整体打包验收，不逐项贴照片）",
         bold=True, first_indent=False, size=12, color=GREEN)
    para(doc, "1. 上述四大服务模块作为一个整体进行交付与验收；甲方在峰会执行过程中按模块归档典型证据材料（如现场代表性照片、印刷成品成型照、屏幕画面截图、回顾视频片段等），不再就每一子项逐一贴附照片。",
         size=10.5)
    para(doc, "2. 大会结束后 7 个自然日内，甲方向乙方一次性出具《赞助权益执行回执》一份，附本报价四大模块对应的代表性证据材料、媒体链接及嘉宾合影，作为整体验收依据。",
         size=10.5)
    para(doc, "3. 若个别细项因现场实际情况调整（如场地像素比、嘉宾名单等），双方仅就调整事项进行书面确认（含电子邮件、企业微信），不影响整体报价与验收。",
         size=10.5)
    blank(doc)

    # ===== 说明与排除项 =====
    para(doc, "六、说明与排除项",
         bold=True, first_indent=False, size=12, color=GREEN)
    para(doc, "1. 本报价为整体打包含税总价；发票内容：会议服务费 / 赞助费；增值税普通发票或专用发票。",
         size=10.5)
    para(doc, "2. 付款方式：合同（含本报价单签字盖章后视同确认）签订之日起 5 个工作日内一次性付款至甲方指定账户；甲方在款项到账后 10 个工作日内向乙方开具等额合法有效发票。",
         size=10.5)
    para(doc, "3. 本次合作不纳入或已弱化的事项：（1）白皮书扉页联合署名 — 不纳入；（2）媒体通稿「不少于 5 家头部财经/地产媒体」「标题级或副标题级」等硬性数量与位置承诺 — 不纳入；（3）双校长三角校友产业联盟战略合作伙伴永久入册 + 牌匾交接 — 不纳入；（4）同行业排他承诺范围 — 限定为「晚宴冠名战略合作伙伴」身份及主背景板核心 logo 位，其他级别合作不在排他范围。",
         size=10.5)
    para(doc, "4. 下列事项由乙方自办，不计入本报价 100,000 元，亦不构成甲方义务：（1）项目案场接驳（华为尊界 8–10 辆 + 考斯特补位）；（2）项目案场接待（沙盘 + 样板间，含夜场灯光秀如适用）；（3）现场销售人员（统一品牌服装与胸卡）；（4）展位升级包装（背景墙、特装等，在甲方提供的基础包装上叠加）。",
         size=10.5)
    para(doc, "5. 本报价单经双方加盖公章（合同章）后，即作为本次合作的正式商业要约与服务交付依据；与《附件四·晚宴冠名战略合作伙伴专项合作协议》共同执行，金额与服务范围以本报价单为准。",
         size=10.5)
    blank(doc)

    # ===== 收款账户 =====
    para(doc, "七、收款账户", bold=True, first_indent=False, size=12, color=GREEN)
    kv_table(doc, rows=[
        ["收款单位", "________________________________"],
        ["开户银行", "________________________________"],
        ["银行账号", "________________________________"],
        ["税号", "________________________________"],
    ])
    blank(doc, 2)

    # ===== 签字盖章 =====
    para(doc, "八、双方确认（签字 / 盖章）",
         bold=True, first_indent=False, size=12, color=GREEN)
    sig = doc.add_table(rows=4, cols=2)
    sig.style = "Table Grid"
    sig.autofit = False
    for row in sig.rows:
        row.cells[0].width = Cm(8.0)
        row.cells[1].width = Cm(8.0)
    cell_text(sig.rows[0].cells[0],
              "甲方（盖章）：\n北京大学经济学院上海校友会 /\n复旦大学住房政策研究中心",
              bold=True, size=11)
    cell_text(sig.rows[0].cells[1],
              "乙方（盖章）：\n绿城中国控股有限公司 /\n____________________________（项目主体）",
              bold=True, size=11)
    cell_text(sig.rows[1].cells[0], "\n\n\n（公章 / 合同章位置）\n\n",
              size=10.5, align_center=True)
    cell_text(sig.rows[1].cells[1], "\n\n\n（公章 / 合同章位置）\n\n",
              size=10.5, align_center=True)
    cell_text(sig.rows[2].cells[0],
              "授权代表（签字）：______________________\n职务：__________________", size=11)
    cell_text(sig.rows[2].cells[1],
              "授权代表（签字）：______________________\n职务：__________________", size=11)
    cell_text(sig.rows[3].cells[0], "签字日期：______年______月______日", size=11)
    cell_text(sig.rows[3].cells[1], "签字日期：______年______月______日", size=11)

    blank(doc)
    para(doc, "本报价单一式肆份，甲乙双方各执贰份，具有同等法律效力。",
         align="center", first_indent=False, size=10, color=GREY)

    out = "/workspace/deliverables/绿城中国-潮鸣外滩-10万元晚宴冠名服务报价单.docx"
    doc.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    build()
