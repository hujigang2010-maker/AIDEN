"""Generate the Greentown China 100,000 RMB sponsorship Word document."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


GREEN = RGBColor(0x00, 0x6B, 0x3F)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
GREY = RGBColor(0x55, 0x55, 0x55)


def set_default_font(doc, name="微软雅黑"):
    style = doc.styles["Normal"]
    style.font.name = name
    style.font.size = Pt(10.5)
    rpr = style.element.rPr
    rfonts = rpr.rFonts if rpr is not None and rpr.rFonts is not None else None
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            style.element.append(rpr)
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), name)
    rfonts.set(qn("w:ascii"), name)
    rfonts.set(qn("w:hAnsi"), name)


def add_heading(doc, text, level=1, color=GREEN):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    if level == 0:
        run.font.size = Pt(22)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif level == 1:
        run.font.size = Pt(16)
    elif level == 2:
        run.font.size = Pt(13)
    else:
        run.font.size = Pt(11.5)
    run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rfonts = OxmlElement("w:rFonts")
    rfonts.set(qn("w:eastAsia"), "微软雅黑")
    rfonts.set(qn("w:ascii"), "微软雅黑")
    rpr.append(rfonts)
    return p


def add_para(doc, text, bold=False, size=10.5, color=DARK, align=None, indent=False):
    p = doc.add_paragraph()
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.74)
    p.paragraph_format.line_spacing = 1.4
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = OxmlElement("w:rFonts")
    rfonts.set(qn("w:eastAsia"), "微软雅黑")
    rfonts.set(qn("w:ascii"), "微软雅黑")
    rpr.append(rfonts)
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.line_spacing = 1.4
    run = p.add_run(text)
    run.font.size = Pt(10.5)
    rpr = run._element.get_or_add_rPr()
    rfonts = OxmlElement("w:rFonts")
    rfonts.set(qn("w:eastAsia"), "微软雅黑")
    rfonts.set(qn("w:ascii"), "微软雅黑")
    rpr.append(rfonts)
    return p


def shade_cell(cell, color_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)


def set_cell_text(cell, text, bold=False, color=DARK, size=10, align_center=False):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    cell.text = ""
    p = cell.paragraphs[0]
    if align_center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = OxmlElement("w:rFonts")
    rfonts.set(qn("w:eastAsia"), "微软雅黑")
    rfonts.set(qn("w:ascii"), "微软雅黑")
    rpr.append(rfonts)


def add_table(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Light Grid Accent 1"
    table.autofit = False
    if col_widths:
        for i, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Cm(w)
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        set_cell_text(hdr[i], h, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), size=10.5, align_center=True)
        shade_cell(hdr[i], "006B3F")
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            set_cell_text(table.rows[r + 1].cells[c], str(val), size=10)
    return table


def build():
    doc = Document()
    set_default_font(doc)

    section = doc.sections[0]
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.4)
    section.right_margin = Cm(2.4)

    add_heading(doc, "绿城中国 | 绿城·潮鸣外滩", level=0, color=GREEN)
    add_para(
        doc,
        "「重构与突围 · 2026 人工智能商业化落地与硬核投资破局峰会」",
        bold=True, size=12, color=DARK, align="center",
    )
    add_para(
        doc,
        "晚宴冠名战略合作伙伴 · 10 万元赞助专项设计方案",
        bold=True, size=14, color=GREEN, align="center",
    )
    add_para(doc, "提报日期：2026 年 5 月", size=10, color=GREY, align="center")
    doc.add_paragraph()

    # 一、合作概要
    add_heading(doc, "一、合作概要 Cooperation Summary", level=1)
    add_table(
        doc,
        headers=["项目", "内容"],
        rows=[
            ["赞助方", "绿城中国控股有限公司"],
            ["项目品牌", "绿城·潮鸣外滩"],
            ["合作身份", "晚宴冠名战略合作伙伴（Dinner Title Strategic Partner）"],
            ["赞助金额", "人民币 100,000 元（壹拾万元整）"],
            ["合作峰会", "重构与突围 · 2026 AI 商业化落地与硬核投资破局峰会"],
            ["主办方", "北京大学经济学院上海校友会 · 复旦大学住房政策研究中心"],
            ["大会时间", "2026 年 5 月 22 日"],
            ["大会地点", "上海·北外滩·一滴水"],
            ["现场规模", "500+ 位高净值嘉宾（25% 上市公司/独角兽高管、20% 一二级市场基金合伙人、25% AI 创业团队创始人、20% 双校核心校友、10% 政府与媒体）"],
        ],
        col_widths=[4.5, 12.5],
    )

    # 二、定制权益概览
    add_heading(doc, "二、10 万元晚宴冠名战略合作伙伴 · 定制权益总览", level=1)
    add_para(
        doc,
        "本级权益在原峰会「钻石赞助 50,000 元」基础上，叠加「晚宴冠名」+「项目专场宣讲」+「场地展位+尊界接驳参观」三大核心独家权益，整体价值评估高于现有钻石+铂金组合，为绿城·潮鸣外滩在 500 人高净值客群中提供从「会前预热—会中渗透—晚宴主场—会后传播—项目到访」的全链路品牌植入。",
        indent=True,
    )
    add_table(
        doc,
        headers=["模块", "核心权益", "形式 / 物料", "执行方"],
        rows=[
            ["A. 晚宴冠名", "「晚宴冠名战略合作伙伴：绿城中国 | 绿城·潮鸣外滩」全程统称权", "晚宴 KV 主视觉、席卡、桌卡、菜单 logo 植入；主持人开场+结束口播鸣谢", "主办 + 活动公司"],
            ["B. 主会场植入", "主背景板 logo、议程手册广告、会场宣传片轮播、500 人手拎袋折页", "主背景板钻石级 logo、议程手册整版、手拎袋内 285/310 户型折页", "主办 + 印刷单位"],
            ["C. 现场展位", "电梯口 + 展场内 2 处专属品牌展位，各派销售 2 人", "易拉宝、户型图 KT 板、纸质物料、桌台轻包装；销售物料台", "绿城方 + 活动公司"],
            ["D. 项目专场宣讲", "晚宴开场前 15 分钟，川总「绿城·潮鸣外滩」专场 PPT 演讲", "PPT、话筒、音响、LED 大屏；主持人引荐", "绿城方 + 活动公司"],
            ["E. 项目到访接驳", "华为尊界接驳 8–10 辆 + 考斯特补位，定向引导嘉宾到访", "尊界品牌车队 + 现场引导员 + 项目案场接待", "绿城方"],
            ["F. 论坛后口播", "圆桌结束后主持人口播赞助方并引导嘉宾前往项目", "主持人定制口播稿；现场动线引导员", "主办"],
            ["G. 媒体宣发", "媒体通稿、九宫格植入、白皮书署名、回顾视频片头鸣谢", "新闻通稿、朋友圈九宫格图、回顾短片", "主办媒体组"],
            ["H. 长效圈层", "双校长三角校友产业联盟战略合作伙伴永久入册", "联盟通讯录、后续活动优先合作权", "主办"],
        ],
        col_widths=[2.5, 5.5, 5.5, 3.5],
    )

    # 三、晚宴现场设计
    add_heading(doc, "三、晚宴现场冠名设计方案", level=1)

    add_heading(doc, "3.1 视觉植入位与责任分工", level=2, color=DARK)
    add_table(
        doc,
        headers=["序号", "物料", "logo 植入方式", "尺寸建议", "责任方"],
        rows=[
            ["1", "晚宴 KV 主视觉（背板/LED 大屏）", "「晚宴冠名战略合作伙伴 绿城中国 | 绿城·潮鸣外滩」横向锁屏", "建议向场地方索取 LED 像素比，常见 1920×1080（16:9）或 6144×1080（横向 LED 条屏）", "主办（绿城提供矢量 logo + slogan）"],
            ["2", "席卡（座位卡）", "卡面右下角 logo + 一行品牌主张", "正面 90×55mm；竖式或横式以场地最终方案为准", "主办"],
            ["3", "桌卡（桌号牌）", "桌号 + 「绿城·潮鸣外滩 · 晚宴冠名」字样", "标准 A5 双面 148×210mm 折立；或亚克力 200×150mm", "活动公司印刷"],
            ["4", "菜单", "封面整版 logo + 内页页脚 logo", "对开 A4（210×285mm）或 285×210mm 单页", "活动公司印刷"],
            ["5", "晚宴期间 LED 轮播", "项目宣传片 + 静帧画面循环", "16:9 1920×1080，30s 宣传片 + 5 张静帧 5s/张", "绿城提供素材，主办上屏"],
            ["6", "主持人口播", "开宴鸣谢 + 川总宣讲引荐 + 散场再次鸣谢", "口播稿 3 段，每段约 30 秒", "主办 + 绿城"],
            ["7", "晚宴桌花卡 / 餐巾纸条", "（可选）项目主 slogan 烫金字条", "150×30mm 条形", "活动公司"],
        ],
        col_widths=[1.0, 3.5, 5.0, 4.8, 3.0],
    )
    add_para(
        doc,
        "说明：席卡 + 晚宴 KV 由主办方统一加上「晚宴冠名战略合作伙伴：绿城中国 | 绿城·潮鸣外滩」字样及 logo；桌卡（桌号）、菜单的项目 logo 植入由活动公司印刷单位负责；其他手册类物料（议程手册、白皮书、手拎袋折页等）由主办方统一在印刷端植入。",
        size=10, color=GREY, indent=True,
    )

    add_heading(doc, "3.2 视频与 PPT 尺寸建议（待场地方最终确认像素比后微调）", level=2, color=DARK)
    add_table(
        doc,
        headers=["素材类型", "推荐比例", "推荐分辨率", "时长 / 帧率", "备注"],
        rows=[
            ["主会场宣传片（轮播）", "16:9", "1920×1080（4K 备份 3840×2160）", "≤ 60s，25/30 fps", "已与主办确认主会场 16:9 无问题"],
            ["晚宴 LED 宣传片", "16:9 横屏（如条屏另出 6144×1080）", "1920×1080，H.264 .mp4 ≤ 200MB", "30–60s 循环", "请场地方提供准确像素比与刷新率"],
            ["川总专场宣讲 PPT", "16:9", "1920×1080 (40cm×22.5cm)", "≤ 25 张，控制 15 分钟内", "建议预留封面/封底/项目沙盘 3D 页"],
            ["晚宴 KV 主视觉", "依场地像素比", "建议先按 1920×1080 出主稿，再按场地比例延展", "静帧 PNG/JPG + 矢量源文件 AI", "导出 RGB 色域"],
            ["项目 logo", "矢量", "AI / EPS / SVG（无白底）", "—", "同时提供 PNG 透明底 4 套尺寸"],
            ["户型折页（285/310）", "印刷品", "印刷 297×210mm 三折页；出血 3mm；CMYK，300 dpi", "—", "由印刷单位统一排版"],
            ["项目易拉宝", "印刷品", "800×2000mm，出血 5mm，CMYK，150 dpi", "—", "电梯口 + 展场内各 1 组"],
            ["朋友圈九宫格", "1:1", "1080×1080 PNG，9 张", "—", "媒体组统一发布，绿城同步素材"],
        ],
        col_widths=[3.5, 3.0, 4.5, 3.0, 3.5],
    )

    # 四、现场动线
    add_heading(doc, "四、现场展位与动线设计", level=1)
    add_para(doc, "4.1 展位布置（共 2 处）", bold=True, size=11.5)
    add_bullet(doc, "电梯口展位：嘉宾抵达第一触点。配置 1 张品牌桌台（含轻包装）、2 个易拉宝、户型折页/纸质物料展示架、销售 2 人（着统一品牌服装，胸卡）。")
    add_bullet(doc, "展场内展位：主会场前厅近茶歇区。配置 1 张洽谈圆桌、2 个易拉宝、285/310 户型 KT 板 1 组、销售 2 人，承接现场意向客户登记。")
    add_bullet(doc, "由于设计制作时间仅 1 天，保底交付：4 个易拉宝 + 2 张品牌桌台轻包装 + 项目纸质物料；如设计时间允许，升级背景挂幔/KT 板背景墙。")

    add_para(doc, "4.2 主会场植入", bold=True, size=11.5)
    add_bullet(doc, "主背景板：钻石级 logo 位（已含在权益内），同时晚宴冠名身份在会议结束后切换为「晚宴冠名战略合作伙伴」副标题位。")
    add_bullet(doc, "宣传片：嘉宾入场及中场休息时段在主舞台 LED 16:9 轮播绿城·潮鸣外滩项目宣传片。")
    add_bullet(doc, "手拎袋：500 份手拎袋内统一植入项目折页 + 285/310 户型图（双折页或三折页），由印刷单位与活动公司对接装袋。")

    add_para(doc, "4.3 论坛 → 晚宴衔接 & 项目到访动线", bold=True, size=11.5)
    add_bullet(doc, "17:55–18:10 颁奖典礼结束后，主持人口播「感谢晚宴冠名战略合作伙伴 绿城中国 · 绿城·潮鸣外滩，欢迎嘉宾前往项目案场实地体验」。")
    add_bullet(doc, "现场设引导员 3 名，引导意向嘉宾至电梯口尊界车队登车。")
    add_bullet(doc, "华为尊界接驳 8–10 辆 + 考斯特补位，往返时间预估 60–90 分钟。")
    add_bullet(doc, "案场接待：项目售楼处沙盘讲解 20 分钟 + 样板间参观 20 分钟，结束后由尊界车队送回晚宴现场。")

    add_heading(doc, "4.4 关于「项目参观与晚宴衔接时间紧张」的处理建议", level=2, color=DARK)
    add_para(
        doc,
        "经测算，往返车程 + 案场参观 ≥ 60 分钟，与现版本议程（17:55 颁奖结束 → 18:10 晚宴开始）存在 45 分钟以上的时间冲突。建议主办方与活动公司沟通采取以下任一方案：",
        indent=True,
    )
    add_table(
        doc,
        headers=["方案", "做法", "优劣"],
        rows=[
            [
                "方案 A · 推荐",
                "现场参观调整为「会后专场」：18:10–20:30 晚宴期间正常进行项目宣讲与轮播；20:30 晚宴结束后，由尊界车队直接将意向嘉宾送至项目案场进行夜场参观（灯光秀+样板间），再分送回酒店。",
                "优势：不打断现有议程；项目方夜场氛围更具记忆点；尊界车队利用率高。劣势：嘉宾时间较晚，需提前确认意向名单。",
            ],
            [
                "方案 B",
                "议程内插入「项目专场参观」：把 17:55 颁奖结束后到晚宴开始之间拉长为 90 分钟参观窗口，将晚宴推迟到 19:30 开始。",
                "优势：参观时间充裕。劣势：影响主办原议程及晚宴餐厅档期，需主办方协调。",
            ],
            [
                "方案 C",
                "「云参观 + 邀约到访」：晚宴期间用 3D VR 沙盘 + 川总 15 分钟现场宣讲完成线上参观；现场提供「次日专车到访」预约登记。",
                "优势：完全不影响议程；可批量收集意向客户。劣势：现场参观感受弱，更像招商方式。",
            ],
        ],
        col_widths=[2.5, 9.5, 5.0],
    )
    add_para(
        doc,
        "建议方案 A 作为首选；最终时间窗以主办、活动公司、绿城三方在 5 月 19 日前确认为准，绿城据此安排尊界车队与案场接待人员。",
        size=10, color=GREY, indent=True,
    )

    # 五、晚宴流程
    add_heading(doc, "五、晚宴现场流程脚本（建议版）", level=1)
    add_table(
        doc,
        headers=["时间", "环节", "内容 / 物料", "音视频需求"],
        rows=[
            ["18:00–18:10", "嘉宾入场", "晚宴 KV 主视觉点亮；席卡引位；项目宣传片轮播（无声/低音乐）", "LED 16:9 视频；背景音乐"],
            ["18:10–18:15", "主持人开场 + 冠名鸣谢", "主持人口播「晚宴冠名战略合作伙伴：绿城中国 | 绿城·潮鸣外滩」", "话筒×2（手持/领夹各 1）"],
            ["18:15–18:30", "川总「绿城·潮鸣外滩」专场宣讲", "项目 PPT 15 分钟 + Q&A 5 分钟", "1080P 16:9 PPT；话筒×1；切换器准备"],
            ["18:30–18:35", "联合祝酒辞", "主办 + 绿城联合致酒辞", "话筒×2"],
            ["18:35–19:30", "正餐第一轮", "项目宣传片循环；菜单植入 logo；桌卡植入 logo", "宣传片循环；背景音乐"],
            ["19:30–20:00", "圈层社交 + 校友联盟介绍", "插入「双校长三角校友产业联盟」介绍片；绿城战略合作伙伴永久入册仪式（牌匾交接）", "LED 视频；摄影"],
            ["20:00–20:25", "正餐第二轮 / 自由交流", "项目销售在场内进行 1V1 沟通", "背景音乐"],
            ["20:25–20:30", "结束鸣谢 & 引导到访", "主持人再次口播鸣谢；引导意向嘉宾至尊界车队", "话筒×1"],
        ],
        col_widths=[2.0, 3.8, 7.5, 3.7],
    )

    # 六、宣发
    add_heading(doc, "六、宣发与回执", level=1)
    add_bullet(doc, "媒体通稿：标题级或副标题级出现「绿城·潮鸣外滩」+ 晚宴冠名身份；主流财经/地产头部媒体不少于 5 家。")
    add_bullet(doc, "朋友圈九宫格：晚宴现场 9 张视觉，至少 3 张含项目 logo / KV，文案口径与绿城品牌部对齐。")
    add_bullet(doc, "白皮书：扉页联合署名「晚宴冠名战略合作伙伴 绿城中国 · 绿城·潮鸣外滩」。")
    add_bullet(doc, "回顾视频：片头 3s 鸣谢、片尾 logo 墙重点位、川总宣讲及参观画面剪入。")
    add_bullet(doc, "大会后 7 个自然日内，由主办向绿城出具《赞助权益执行回执》，附现场图片、媒体链接、嘉宾合影等证明材料。")

    # 七、物料与时间节点
    add_heading(doc, "七、时间节点 Timeline", level=1)
    add_table(
        doc,
        headers=["日期", "节点", "绿城方", "主办 / 活动公司"],
        rows=[
            ["5/18", "确认合作框架及金额", "盖章合同回传", "出具盖章合同"],
            ["5/19", "物料源文件交付", "提供 logo（AI/EPS/PNG）、项目宣传片、川总宣讲 PPT、户型折页源文件、品牌简介", "提供晚宴 KV、桌卡、菜单、席卡设计稿初稿"],
            ["5/20", "设计稿确认", "确认设计稿一轮，提出修改意见", "完成印刷物料终稿"],
            ["5/21", "印刷下单 & 物流", "—", "桌卡、菜单、易拉宝、户型折页全部下印"],
            ["5/22 上午", "现场布展", "销售 4 人到场对接展位包装", "活动公司完成展位、KV、LED 调试"],
            ["5/22 全天", "大会执行", "执行+陪同+案场对接", "执行+主持口播+宣传片轮播"],
            ["5/23–5/29", "宣发 & 回执", "同步绿城自媒体矩阵传播", "通稿/九宫格发布；7 日内出回执"],
        ],
        col_widths=[2.0, 3.5, 5.5, 6.0],
    )

    # 八、付款与发票
    add_heading(doc, "八、付款与发票", level=1)
    add_bullet(doc, "赞助金额：人民币 100,000 元（壹拾万元整）。")
    add_bullet(doc, "支付方式：合同签订后 5 个工作日内一次性付款至主办方指定账户。")
    add_bullet(doc, "发票内容：会议服务费 / 赞助费，开具等额合法有效增值税普通发票或专用发票。")
    add_bullet(doc, "执行依据：参照《2026 人工智能商业化落地与硬核投资破局峰会赞助合作协议》及其《附件一·赞助权益清单》，本方案作为定制权益的执行细则附件。")

    # 九、对接人
    add_heading(doc, "九、对接人 Contact", level=1)
    add_table(
        doc,
        headers=["角色", "对接事项", "对接人", "备注"],
        rows=[
            ["绿城方·总对接", "整体方案、付款、合同", "Aiden", "合作框架内植入物料对接"],
            ["绿城方·项目宣讲", "川总专场 PPT、品牌口径", "川总团队", "确认 PPT 终稿"],
            ["绿城方·现场销售", "电梯口 & 展场内 2 处展位", "销售负责人", "销售 4 人，统一物料"],
            ["主办·组委会招商", "权益落地、媒体宣发", "组委会招商组 13262607888", "微信 13262607888"],
            ["主办·现场执行", "场地、KV、印刷、流程", "wangsheng", "确认晚宴 KV 像素比、印刷物料植入"],
            ["活动公司·印刷", "桌卡、菜单、手册物料 logo 植入", "印刷负责人", "5/21 前完成印刷"],
        ],
        col_widths=[3.5, 5.5, 4.5, 3.5],
    )

    add_para(doc, " ", size=8)
    add_para(
        doc,
        "本方案为「重构与突围 · 2026 AI 商业化峰会」与绿城中国 · 绿城·潮鸣外滩 10 万元晚宴冠名战略合作伙伴的执行细则，最终以双方签署的赞助合作协议正文及附件为准。",
        size=9, color=GREY, align="center",
    )

    out = "/workspace/deliverables/绿城中国-潮鸣外滩-10万元晚宴冠名赞助专项设计方案.docx"
    doc.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    build()
