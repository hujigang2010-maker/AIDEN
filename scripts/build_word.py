"""生成 Word：临港 x WeTest 中日数字经济闭门沙龙 · 活动策划方案（侧重会后转化）"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_cell_bg(cell, color_hex: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)


def add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.name = "微软雅黑"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
        if level == 0:
            run.font.color.rgb = RGBColor(0x0B, 0x3D, 0x91)
        elif level == 1:
            run.font.color.rgb = RGBColor(0x0B, 0x3D, 0x91)
        else:
            run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    return h


def add_p(doc, text, bold=False, size=10.5, color=None, align=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    run = p.add_run(text)
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    return p


def add_bullets(doc, items, style="List Bullet"):
    for it in items:
        p = doc.add_paragraph(style=style)
        run = p.add_run(it)
        run.font.name = "微软雅黑"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
        run.font.size = Pt(10.5)


def add_table(doc, headers, rows, header_fill="0B3D91", first_col_fill=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Light Grid Accent 1"
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        p = hdr[i].paragraphs[0]
        r = p.add_run(h)
        r.bold = True
        r.font.name = "微软雅黑"
        r._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_cell_bg(hdr[i], header_fill)
    for ri, row in enumerate(rows):
        cells = table.rows[ri + 1].cells
        for ci, val in enumerate(row):
            cells[ci].text = ""
            p = cells[ci].paragraphs[0]
            r = p.add_run(str(val))
            r.font.name = "微软雅黑"
            r._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
            r.font.size = Pt(9.5)
            if first_col_fill and ci == 0:
                set_cell_bg(cells[ci], first_col_fill)
    return table


def main():
    doc = Document()
    # 全局样式
    style = doc.styles["Normal"]
    style.font.name = "微软雅黑"
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")

    # 页面设置
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(2.2)

    # ============ 封面 ============
    add_p(doc, "")
    add_p(doc, "")
    add_p(doc, "临港 × WeTest", bold=True, size=26,
          color=(0x0B, 0x3D, 0x91), align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p(doc, "中日数字经济跨国发展与高质量交付闭门会", bold=True, size=20,
          color=(0x0B, 0x3D, 0x91), align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p(doc, "活动策划方案（侧重会后转化）", size=14,
          color=(0x55, 0x55, 0x55), align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p(doc, "")
    add_p(doc, "政策护航（临港新片区）  ×  质量守护（WeTest）",
          bold=True, size=12, color=(0x33, 0x33, 0x33),
          align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p(doc, "")
    add_p(doc, "—— 一站式中日出海/入华解决方案沙龙 ——",
          size=11, color=(0x99, 0x99, 0x99),
          align=WD_ALIGN_PARAGRAPH.CENTER)
    for _ in range(8):
        add_p(doc, "")
    add_p(doc, "主办方：上海科技企业协会", size=11,
          align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p(doc, "联合主办：临港科技城  |  WeTest 一站式质量开放平台", size=11,
          align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p(doc, "版本：V1.0    编制：活动策划组",
          size=10, color=(0x99, 0x99, 0x99),
          align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()

    # ============ 目录说明 ============
    add_heading(doc, "目录", level=1)
    toc_items = [
        "一、活动背景与战略意义",
        "二、核心主题与价值主张",
        "三、目标人群画像与邀约策略",
        "四、活动概览与议程详细设计",
        "五、嘉宾阵容与分享设计",
        "六、★ 转化核心策略（会前 / 会中 / 会后）",
        "七、会后跟进 SOP（7-30-60-90 天节奏）",
        "八、中日专项测试方案（业务包装）",
        "九、四大行业精准切入策略",
        "十、资源、分工与时间倒排",
        "十一、预算明细",
        "十二、风险评估与应急预案",
        "十三、KPI 与复盘机制",
        "附录 A：邀约话术与会后跟进话术模板",
        "附录 B：现场转化抓手清单",
    ]
    for i, it in enumerate(toc_items, 1):
        add_p(doc, f"  {it}", size=11)
    doc.add_page_break()

    # ============ 一、活动背景 ============
    add_heading(doc, "一、活动背景与战略意义", level=1)

    add_heading(doc, "1.1 宏观背景", level=2)
    add_p(doc,
          "中日两国数字经济双向流动正在加速：中资互联网/游戏/金融/电商企业加快"
          "面向日本市场出海布局，日资企业亦寻求借助小程序、移动支付等中国独有"
          "数字生态实现入华增长。两条业务流共同面对"
          "「数据跨境合规、本地化支付与设备适配、网络与风控、隐私与安全审计」"
          "等深度痛点。")

    add_heading(doc, "1.2 临港新片区的政策红利", level=2)
    add_bullets(doc, [
        "数据跨境流动：自贸区特殊数据跨境审批机制，是金融/游戏/电商企业的最大刚需。",
        "金融创新：跨境支付、跨境投融资便利化，匹配跨境业务结算与回流。",
        "税收优惠与人才落户：减轻企业出海/入华初期的人才与运营成本。",
        "国际化营商：与日本经济产业省、JETRO 等的桥梁更顺畅。",
    ])

    add_heading(doc, "1.3 WeTest 的落地能力", level=2)
    add_bullets(doc, [
        "腾讯官方出品的一站式质量开放平台，覆盖兼容、性能、安全、自动化、众测等。",
        "UDT 全球建站能力，可在日本接入本地真机进行远程调试与共享调度。",
        "海外众测：覆盖真实网络（弱网/丢包/延迟）、本地真金支付（信用卡/电子钱包/Apple/GP）、本地语言功能验收。",
        "CrashSight 海外数据合规上报、小程序合规扫描、防篡改/防爬虫等安全能力。",
    ])

    add_heading(doc, "1.4 战略契合点", level=2)
    add_p(doc, "本次活动以 “政策护航 × 质量守护” 为价值主线，为受邀企业提供"
          "「宏观政策合规 → 中观产业战略 → 微观技术落地」的一站式跨国增长解决方案。"
          "活动定位为高规格、定向邀请的闭门生态沙龙，确保信号密度高、转化路径短。",
          bold=True)

    doc.add_page_break()

    # ============ 二、核心主题与价值主张 ============
    add_heading(doc, "二、核心主题与价值主张", level=1)

    add_heading(doc, "2.1 主题定位", level=2)
    add_p(doc, "中日数字经济跨国发展与高质量交付闭门会", bold=True, size=14,
          color=(0x0B, 0x3D, 0x91))
    add_p(doc, "副标题：政策红利 × 测试护航——跨国业务的合规、落地与转化")

    add_heading(doc, "2.2 一句话价值主张", level=2)
    add_p(doc,
          "「来临港，看政策；用 WeTest，跑通日本/中国市场最后一公里。」",
          bold=True, color=(0xC0, 0x39, 0x2B))

    add_heading(doc, "2.3 三方价值矩阵", level=2)
    add_table(doc,
        ["角色", "核心诉求", "本次活动带来的价值"],
        [
            ["受邀企业 C 级/业务 VP",
             "跨国合规、落地速度、风险可控",
             "一次性获取政策路径 + 学术趋势 + 技术解决方案"],
            ["临港科技城",
             "高质量企业落地、产业集聚",
             "精准触达 20-30 家有跨国业务诉求的目标企业"],
            ["WeTest 业务线",
             "中高净值客户线索、签单转化",
             "在权威场域中完成产品演示与圆桌共创、获取强意向线索"],
        ])

    doc.add_page_break()

    # ============ 三、目标人群 ============
    add_heading(doc, "三、目标人群画像与邀约策略", level=1)

    add_heading(doc, "3.1 目标行业（四大行业）", level=2)
    add_table(doc,
        ["行业", "典型痛点", "WeTest 抓手"],
        [
            ["金融",
             "数据跨境合规、风控、隐私扫描、防篡改",
             "CrashSight 海外合规 + 安全加固 + 小程序合规扫描"],
            ["游戏",
             "海外弱网、设备碎片化、支付与防作弊",
             "UDT 日本真机 + 海外众测 + 真金支付测试 + 防外挂"],
            ["电商/零售",
             "本地支付通路、防黑产、跨境物流接口、风控封号",
             "海外多渠道支付验收 + 安全网关 + 小程序兼容"],
            ["泛互联网",
             "本地化功能 / 多语言 / 网络适配 / SDK 兼容",
             "兼容/性能/众测 + 自动化巡检 + 远程真机"],
        ])

    add_heading(doc, "3.2 目标决策人画像", level=2)
    add_bullets(doc, [
        "中资出海方：CTO / CIO / 海外业务 VP / 海外质量负责人 / 合规负责人。",
        "日资入华方：日本总部 CTO / 中国区总经理 / 中国区技术负责人 / 数字化负责人。",
        "决策权：单笔 50 万 - 500 万人民币测试/合规预算的拍板人或强影响人。",
    ])

    add_heading(doc, "3.3 邀约渠道与配比（共邀 60-80 人，确认到场 25 人）", level=2)
    add_table(doc,
        ["渠道", "邀请数量", "预计到场转化率", "到场人数"],
        [
            ["上海科技企业协会会员定向邀请", "30", "40%", "12"],
            ["临港科技城既有企业池 / 招商线索", "15", "35%", "5"],
            ["WeTest 销售线索池（中日有出海/入华诉求）", "20", "30%", "6"],
            ["合作伙伴/媒体推荐（含日资 JETRO 等）", "10", "20%", "2"],
            ["合计", "75", "—", "25"],
        ])

    add_heading(doc, "3.4 邀约节奏（T 为活动日）", level=2)
    add_bullets(doc, [
        "T-21：协会/临港/WeTest 三方对齐邀约名单（NDA 内部锁定）。",
        "T-15：定向 1V1 电话/微信首轮触达 + 邀请函发出（含议程亮点）。",
        "T-10：寄出实体邀请函/伴手礼（针对 C 级）+ 二轮确认。",
        "T-5：到场确认 + 议程二次推送 + 圆桌议题征集（关键转化动作）。",
        "T-2：交通/接待方案推送 + 司机/会务对接。",
        "T-0：现场签到 + 名片扫码进群 + 议题预热。",
    ])

    doc.add_page_break()

    # ============ 四、活动概览与议程 ============
    add_heading(doc, "四、活动概览与议程详细设计", level=1)

    add_heading(doc, "4.1 活动概览", level=2)
    add_table(doc,
        ["项目", "内容"],
        [
            ["主题", "中日数字经济跨国发展与高质量交付闭门会"],
            ["时间", "下午 13:30 - 16:30（半日闭门）+ 16:30 - 18:00 自由对接 / 一对一深聊"],
            ["地点", "临港科技城（具体会议室与展厅由临港统一安排）"],
            ["形式", "闭门圆桌 + 政策宣讲 + 产业探访 + 案例分享 + 现场对接"],
            ["规模", "20-30 人（C 级 / 业务 VP 为主，确保信号密度与转化深度）"],
            ["语言", "中文为主，日语 / 英文资料同步提供"],
        ])

    add_heading(doc, "4.2 半日议程", level=2)
    add_table(doc,
        ["时间", "环节", "内容", "转化触点"],
        [
            ["13:00-13:30", "签到 & 茶歇",
             "签到表 + 名片扫码自动入群 + 一对一线下破冰",
             "采集联系方式与意向标签"],
            ["13:30-14:30", "【产业探访】",
             "临港科技城参观 / 展厅 / 政策初探",
             "建立 “场” 的政策势能"],
            ["14:30-15:10", "【宏观指引】",
             "姚志勇教授（拟邀）：中日数字贸易趋势 + 上海/临港自贸区跨国壁垒",
             "建立权威认知"],
            ["15:10-15:50", "【技术落地】",
             "WeTest 业务线分享：中日跨国业务的高质量交付与安全合规",
             "★ 关键转化触点：现场 Demo + 案例 + 限定权益"],
            ["15:50-16:30", "【闭门圆桌】",
             "议题：跨国双向拓展中最大的本地化阻碍是什么？",
             "★ 现场 1V1 资源对接 + 测试服务初步合作意向"],
            ["16:30-18:00", "【晚间深聊】",
             "Networking / 一对一深聊 / 高意向客户拉群预约下次专项会谈",
             "★ 完成 “线索→机会” 转化升级"],
        ])

    doc.add_page_break()

    # ============ 五、嘉宾阵容 ============
    add_heading(doc, "五、嘉宾阵容与分享设计", level=1)

    add_table(doc,
        ["环节", "拟邀嘉宾", "身份", "分享要点", "时长"],
        [
            ["开场致辞", "临港科技城高管 + 协会秘书长",
             "主办方 / 联合主办",
             "政策护航 × 质量守护 主线宣讲",
             "10 分钟"],
            ["宏观指引", "姚志勇教授（或同等权威）",
             "学术专家",
             "中日数字贸易趋势 + 临港自贸区构建跨国商业壁垒",
             "40 分钟"],
            ["技术落地", "WeTest 业务线代表",
             "技术专家",
             "UDT 日本真机 / 海外众测 / 真金支付 / 合规扫描 全栈方案",
             "40 分钟"],
            ["圆桌主持", "WeTest CSO 或行业总监",
             "圆桌主持",
             "引导四大行业代表表达痛点 + 当场承接",
             "40 分钟"],
            ["圆桌嘉宾", "金融/游戏/电商/泛互联网各 1 位 C 级",
             "客户代表",
             "现身说法 + 痛点共鸣 + 行业引领",
             "40 分钟"],
        ])

    doc.add_page_break()

    # ============ 六、转化核心策略 ============
    add_heading(doc, "六、★ 转化核心策略（会前 / 会中 / 会后）", level=1)
    add_p(doc,
          "「会议本身不是目的，会议是转化的开关。」全部内容设计、嘉宾配置、"
          "现场动线、伴手礼、议程节奏都必须服务于一条主线 —— "
          "把高净值客户从「认知」推进到「机会」直到「签单」。",
          bold=True, color=(0xC0, 0x39, 0x2B))

    add_heading(doc, "6.1 转化漏斗设计（目标）", level=2)
    add_table(doc,
        ["阶段", "动作", "目标转化率", "目标人数"],
        [
            ["L0 邀请池", "定向邀约", "—", "75"],
            ["L1 到场（已认知）", "签到 + 名片扫码", "33%", "25"],
            ["L2 强意向（兴趣）", "圆桌发言 / 现场预约下次深聊", "60%", "15"],
            ["L3 商机（机会）", "T+7 内完成需求澄清会", "60%", "9"],
            ["L4 POC / 试用", "T+30 内立项 POC", "55%", "5"],
            ["L5 签单", "T+90 内首单签订", "60%", "3"],
        ])
    add_p(doc, "整体「邀请池 → 签单」目标转化率 ≥ 4%（3 单），ROI 目标 ≥ 5×。",
          bold=True)

    add_heading(doc, "6.2 会前转化（建立期待 + 锁定需求）", level=2)
    add_bullets(doc, [
        "邀请函中嵌入「圆桌议题征集表」：要求受邀人提交 1 个跨国痛点 —— 提前掌握线索画像。",
        "T-5 推送《WeTest 中日专项测试方案白皮书》预读版，建立技术认知。",
        "T-3 由 WeTest 销售对前 10 名高优客户做 1V1 电话预热，确认到场并预约会后 30 分钟深聊。",
        "T-2 推送嘉宾介绍卡片 + 当日动线图，提升仪式感与到场率。",
    ])

    add_heading(doc, "6.3 会中转化（高密度抓手）", level=2)
    add_bullets(doc, [
        "签到处「扫码进群 + 选择行业标签」：自动完成线索分级与企微触达。",
        "茶歇区设「WeTest 体验角」：演示 UDT 日本真机远程调试、海外众测看板、CrashSight 海外合规面板。",
        "现场资料袋含三件套：临港政策手册 + WeTest 中日专项方案 + 限时权益券（30 天免费 POC 名额 / 5 万元测试代金券）。",
        "圆桌环节由 WeTest 行业总监主持，每位发言客户配 1 名 WeTest 行业 SA 现场记录 + 当场承接需求。",
        "结束前 5 分钟：现场宣布「闭门客户专属通道」—— 微信/钉钉直连 WeTest 中日团队 + 临港落地顾问。",
        "晚间深聊：高优客户预约成功者赠送伴手礼（含 1V1 闭门技术咨询券）。",
    ])

    add_heading(doc, "6.4 会后转化（48 小时黄金期）", level=2)
    add_p(doc, "活动结束后 48 小时是转化的「黄金窗口」，必须按 SOP 执行：",
          bold=True, color=(0xC0, 0x39, 0x2B))
    add_bullets(doc, [
        "T+1（次日 10:00 前）：群发感谢信 + 现场资料包 + 1V1 销售触达分级名单。",
        "T+2：完成所有到场客户的 CRM 录入，按 ABC 三级分级（详见附录）。",
        "T+3 至 T+7：A 级客户完成需求澄清会（线下/视频），B 级客户完成方案 PPT 定向推送，C 级客户进入月度内容运营池。",
        "T+15：发出活动复盘报告 + 行业洞察简报（建立持续触达节奏）。",
        "T+30：A 级客户启动 POC；B 级客户完成第二次深聊。",
        "T+60/90：A 级客户力争首单签订；B 级客户进入 POC；C 级客户进入二次活动邀约池。",
    ])

    doc.add_page_break()

    # ============ 七、会后跟进 SOP ============
    add_heading(doc, "七、会后跟进 SOP（7-30-60-90 天节奏）", level=1)

    add_heading(doc, "7.1 客户分级标准（ABC 三级）", level=2)
    add_table(doc,
        ["级别", "判定标准", "目标动作", "负责人"],
        [
            ["A 级（高优）",
             "C 级到场 + 圆桌发言 + 明确表达预算/项目 + 公司符合四大行业",
             "T+7 完成需求澄清；T+30 启动 POC；T+90 力争首单",
             "WeTest 行业总监 + 大客户经理"],
            ["B 级（中优）",
             "业务 VP 到场 + 表达兴趣但未明确预算 / 项目方向需挖掘",
             "T+15 二次深聊；T+45 输出针对性方案；T+90 进入 POC",
             "WeTest 大客户经理"],
            ["C 级（培育）",
             "代表到场 / 决策权弱 / 暂无明确项目",
             "进入月度内容运营池；季度二次活动邀约",
             "WeTest 市场 + 内容运营"],
        ])

    add_heading(doc, "7.2 会后跟进里程碑", level=2)
    add_table(doc,
        ["时间", "动作", "A 级", "B 级", "C 级"],
        [
            ["T+1", "感谢信 + 资料包 + 1V1 触达", "电话", "微信", "邮件"],
            ["T+3", "CRM 录入完成 / 分级", "完成", "完成", "完成"],
            ["T+7", "需求澄清会", "完成", "排期", "—"],
            ["T+15", "方案 PPT 定向推送", "—", "完成", "推送"],
            ["T+30", "POC 立项 / 二次深聊", "POC 启动", "二次深聊", "月度内容"],
            ["T+60", "POC 中间汇报 / 商务", "进行中", "POC 启动", "季度活动"],
            ["T+90", "首单签订 / POC 验收", "首单签订", "POC 验收", "—"],
        ])

    add_heading(doc, "7.3 跟进工具与协同机制", level=2)
    add_bullets(doc, [
        "工具：企业微信 + 飞书 CRM + WeTest 内部销售看板（每周三同步进度）。",
        "协同：临港科技城招商顾问 + 协会会员服务 + WeTest 销售 三方周例会，"
        "共同推进高优客户的政策落地与技术 POC 联动。",
        "KPI 责任：每条 A/B 级线索需有明确 Owner + 周更新动作 + 月度复盘。",
    ])

    doc.add_page_break()

    # ============ 八、中日专项测试方案（业务包装）============
    add_heading(doc, "八、中日专项测试方案（业务包装）", level=1)

    add_p(doc,
          "为提升活动后的转化效率，在沙龙举办前后将 WeTest 服务包装为 3 个"
          "「即买即用」的中日专项产品包，降低客户决策门槛。",
          bold=True)

    add_table(doc,
        ["产品包", "包含模块", "目标客户", "起价（人民币）", "首期 POC 周期"],
        [
            ["日本出海加速包",
             "UDT 日本真机 + 海外众测（弱网/真金支付）+ CrashSight 海外合规",
             "出海日本的金融/游戏/电商/泛互联网",
             "¥98,000 / 月",
             "2 周"],
            ["日资入华质量包",
             "小程序兼容扫描 + 微信隐私合规检测 + 支付宝/微信小程序安全加固",
             "入华日资企业（特别是金融、零售）",
             "¥128,000 / 月",
             "2 周"],
            ["跨境合规安全包",
             "数据跨境合规扫描 + 安全网关（防爬/防黑产）+ 渗透测试",
             "金融、电商、政务",
             "¥168,000 / 单项目",
             "3 周"],
        ])

    add_heading(doc, "8.1 包装核心卖点（用于现场宣讲与销售话术）", level=2)
    add_bullets(doc, [
        "「即开即用」：合同签订 5 个工作日内开通，POC 不超过 2 周出报告。",
        "「免封号」：海外真金支付测试，规避账号/卡审查引起的封号与风控。",
        "「双向闭环」：中→日 与 日→中 互补打通，一份方案两个市场。",
        "「合规托底」：CrashSight 海外数据合规 + 小程序隐私扫描，过审无忧。",
    ])

    doc.add_page_break()

    # ============ 九、四大行业切入策略 ============
    add_heading(doc, "九、四大行业精准切入策略", level=1)

    add_table(doc,
        ["行业", "出海/入华主诉求", "首选方案包", "首选案例展示", "首单切入产品"],
        [
            ["金融",
             "数据跨境 + 风控 + 隐私合规",
             "跨境合规安全包",
             "银联云小程序安全扫描案例",
             "小程序合规扫描 + 数据跨境审计"],
            ["游戏",
             "日本市场首发 + 网络/设备适配 + 真金支付",
             "日本出海加速包",
             "跨境电商 UDT + 海外众测案例",
             "海外众测 + UDT 日本真机"],
            ["电商/零售",
             "本地支付 + 防黑产 + 多端兼容",
             "日本出海加速包 + 跨境合规安全包",
             "跨境电商 UDT 案例",
             "海外支付验收 + 安全网关"],
            ["泛互联网",
             "出海工具/SaaS 本地化、APP 性能",
             "日本出海加速包",
             "PerfDog + 海外众测案例",
             "PerfDog + 兼容/众测"],
        ])

    doc.add_page_break()

    # ============ 十、资源、分工与时间倒排 ============
    add_heading(doc, "十、资源、分工与时间倒排", level=1)

    add_heading(doc, "10.1 三方分工", level=2)
    add_table(doc,
        ["主体", "核心职责"],
        [
            ["上海科技企业协会",
             "总策划 / 邀请名单把关 / 现场主持 / 媒体宣发 / 复盘报告"],
            ["临港科技城",
             "场地与展厅 / 政策宣讲嘉宾 / 招商顾问 / 落地政策对接"],
            ["WeTest 业务线",
             "技术分享嘉宾 / 案例 / 体验角 / 销售跟进 / 转化执行"],
        ])

    add_heading(doc, "10.2 关键里程碑（倒排）", level=2)
    add_table(doc,
        ["阶段", "时间", "关键动作", "责任方"],
        [
            ["T-30", "立项启动", "三方 Kick-off，确定主题/嘉宾/预算", "协会"],
            ["T-25", "名单初稿", "三方提交各自邀请池", "三方"],
            ["T-21", "锁定名单", "去重 + NDA 锁定", "协会"],
            ["T-15", "首轮邀约", "电话/微信 + 邀请函", "三方销售"],
            ["T-10", "二轮邀约", "实体邀请函 + 二次确认", "协会"],
            ["T-7", "议程定稿", "嘉宾 PPT 确认 + 圆桌议题征集", "WeTest + 临港"],
            ["T-5", "白皮书推送", "技术预读 + 1V1 预热", "WeTest"],
            ["T-3", "现场彩排", "动线 / 设备 / 主持稿 / 应急预案", "三方"],
            ["T-2", "交通接待", "推送 + 司机/酒店对接", "协会"],
            ["T0", "活动执行", "现场转化抓手全部触发", "三方"],
            ["T+1", "感谢信", "群发 + 1V1 触达", "WeTest 销售"],
            ["T+7", "需求澄清", "A 级完成需求澄清会", "WeTest"],
            ["T+30", "POC", "A 级启动 POC / B 级二深聊", "WeTest"],
            ["T+90", "首单", "首单签订 / 季度复盘", "三方"],
        ])

    doc.add_page_break()

    # ============ 十一、预算 ============
    add_heading(doc, "十一、预算明细（参考，单位：人民币元）", level=1)

    budget = [
        ["场地与展厅", "临港会议室 + 展厅参观 + 茶歇", "由临港承担", 0],
        ["邀请函/印刷物料", "实体邀请函 + 资料袋 + 政策手册", "60 份", 18000],
        ["嘉宾费用", "学术嘉宾出席费 + 交通", "1 位", 30000],
        ["伴手礼", "C 级专属礼盒（含限时权益券）", "30 份 × 600", 18000],
        ["茶歇 & 餐饮", "下午茶 + 晚间深聊轻餐", "30 人 × 250", 7500],
        ["影像/直播", "摄影/摄像/精剪 + 闭门精华片", "1 场", 20000],
        ["主持/会务", "外部主持 + 礼仪 + 签到", "1 场", 12000],
        ["现场 Demo 设备", "日本真机/演示电脑/大屏租赁", "1 场", 8000],
        ["交通接待", "C 级 1V1 接送 + 高铁/酒店补贴", "10 人", 15000],
        ["会后传播", "媒体通稿 + 行业自媒体 + 视频剪辑", "1 轮", 12000],
        ["机动 10%", "应急 / 升级", "—", 14000],
    ]
    total = sum(x[3] for x in budget)
    rows = [[a, b, c, f"¥{d:,}"] for a, b, c, d in budget]
    rows.append(["合计", "", "", f"¥{total:,}"])
    add_table(doc, ["科目", "内容", "数量/标准", "预算"], rows)
    add_p(doc, "* 场地、政策嘉宾由临港承担；WeTest 业务线承担技术 Demo 与会后客户 SA 资源。",
          size=10, color=(0x77, 0x77, 0x77))

    doc.add_page_break()

    # ============ 十二、风险预案 ============
    add_heading(doc, "十二、风险评估与应急预案", level=1)
    add_table(doc,
        ["风险类别", "具体风险", "概率", "影响", "预案"],
        [
            ["到场率", "目标客户临时缺席", "中", "高",
             "T-2 / T-1 二轮确认 + 备选名单 5 人 + 直播链路（仅限协会内部）"],
            ["嘉宾", "学术嘉宾日程冲突", "低", "高",
             "提前锁定 1 位 backup（同领域权威）+ 预演 PPT"],
            ["现场设备", "日本真机网络不稳定", "中", "中",
             "本地录屏备份 Demo + 双链路网络冗余"],
            ["转化执行", "WeTest 销售跟进不及时", "中", "高",
             "T+1 强制 1V1 触达 SOP + 周例会督办"],
            ["合规/媒体", "闭门内容外泄", "低", "中",
             "签到处 NDA 签署 + 现场无直播 + 精剪后由协会统一发布"],
            ["疫情/不可抗力", "线下临时改线上", "低", "高",
             "Teams/腾讯会议直播方案 + 物料邮寄"],
        ])

    doc.add_page_break()

    # ============ 十三、KPI ============
    add_heading(doc, "十三、KPI 与复盘机制", level=1)

    add_heading(doc, "13.1 量化 KPI（活动 + 转化）", level=2)
    add_table(doc,
        ["维度", "指标", "目标值"],
        [
            ["邀约", "邀请池规模", "≥ 75 人"],
            ["邀约", "C 级 / 业务 VP 占比", "≥ 70%"],
            ["到场", "实际到场", "≥ 25 人"],
            ["到场", "到场率", "≥ 33%"],
            ["现场", "圆桌发言客户数", "≥ 8"],
            ["现场", "高优客户预约下次深聊数", "≥ 12"],
            ["转化", "T+7 完成需求澄清会数", "≥ 9"],
            ["转化", "T+30 POC 立项数", "≥ 5"],
            ["转化", "T+90 首单数", "≥ 3"],
            ["收入", "活动直接归因 GMV（首单）", "≥ ¥3,000,000"],
            ["ROI", "活动 ROI（首单 GMV / 预算）", "≥ 5×"],
            ["品牌", "媒体/自媒体露出次数", "≥ 10"],
        ])

    add_heading(doc, "13.2 复盘机制", level=2)
    add_bullets(doc, [
        "T+3 内部三方快速复盘（90 分钟）：到场质量、现场执行、销售触达完成度。",
        "T+15 客户视角复盘：抽样回访 5 位 A 级客户，收集体验改进点。",
        "T+30 转化中期复盘：POC 进度、销售漏斗健康度、流失原因分析。",
        "T+90 转化收单复盘：ROI 计算、案例沉淀、下一场迭代计划。",
        "建立常态化季度沙龙机制：以本场为模板，每季度迭代行业/主题。",
    ])

    doc.add_page_break()

    # ============ 附录 A：话术 ============
    add_heading(doc, "附录 A：邀约话术与会后跟进话术模板", level=1)

    add_heading(doc, "A.1 邀约话术（电话/微信）", level=2)
    add_p(doc,
          "「X 总您好，我是上海科技企业协会的 XX。本月 XX 日下午我们在临港科技城举办一场闭门沙龙，"
          "主题是「中日数字经济跨国发展与高质量交付」。临港会现场宣讲数据跨境与跨境支付的最新政策，"
          "姚志勇教授会做中日数字贸易的趋势分享，WeTest 团队会现场演示日本真机调试与海外真金支付测试。"
          "全场只邀请 25 位左右 C 级高管，您贵公司在 XX（行业）的跨国布局非常有代表性，"
          "我们特意为您预留了一个席位，希望您能拨冗出席。」",
          size=11)

    add_heading(doc, "A.2 会后 T+1 感谢信关键句", level=2)
    add_p(doc,
          "「感谢您出席「中日数字经济跨国发展与高质量交付闭门会」。"
          "随信附上：① 临港政策手册；② WeTest 中日专项方案白皮书；"
          "③ 您在圆桌中提到的 X 项痛点对应的初步解决思路（PDF）。"
          "我们的 X 总（WeTest 行业 SA）将在 48 小时内与您预约 30 分钟需求澄清会，"
          "并基于您的实际情况输出一份「日本出海加速包 / 跨境合规安全包」的定制 POC 方案。」",
          size=11)

    add_heading(doc, "A.3 会后 T+7 需求澄清会议程模板", level=2)
    add_bullets(doc, [
        "① 业务概览：客户跨国业务现状与未来 12 个月计划（10 分钟）。",
        "② 痛点深挖：合规 / 性能 / 安全 / 本地化 四象限对照（15 分钟）。",
        "③ WeTest 方案匹配：从 3 个专项包中选定主方案（15 分钟）。",
        "④ POC 设计：范围、周期、验收标准、商务条款（15 分钟）。",
        "⑤ Next Step：双方 Owner、时间表、下次会议（5 分钟）。",
    ])

    doc.add_page_break()

    # ============ 附录 B：现场转化抓手清单 ============
    add_heading(doc, "附录 B：现场转化抓手清单（执行 checklist）", level=1)

    add_table(doc,
        ["序号", "抓手", "落地形式", "责任人"],
        [
            ["1", "扫码进群 + 行业标签", "签到二维码 + 企微自动打标", "市场 + IT"],
            ["2", "WeTest 体验角", "UDT 日本真机 / 海外众测看板 / CrashSight 面板", "WeTest SA"],
            ["3", "限时权益券", "30 天免费 POC / 5 万元代金券", "WeTest 销售"],
            ["4", "圆桌一对一承接", "每位发言客户配 1 名 WeTest 行业 SA", "WeTest"],
            ["5", "闭门客户专属通道", "微信/钉钉直连中日团队 + 临港顾问", "三方"],
            ["6", "晚间深聊预约", "高优客户 1V1 30 分钟", "WeTest"],
            ["7", "伴手礼附咨询券", "1V1 闭门技术咨询券 + 政策对接券", "协会"],
            ["8", "现场拍摄客户证言", "短视频 / 案例素材", "市场"],
        ])

    add_p(doc, "")
    add_p(doc, "—— 本方案为 V1.0，需在三方 Kick-off 后完成最终对齐与细化 ——",
          size=10, color=(0x99, 0x99, 0x99), align=WD_ALIGN_PARAGRAPH.CENTER)

    out = "/workspace/deliverables/临港x WeTest 中日闭门会_活动策划方案_V1.0.docx"
    doc.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
