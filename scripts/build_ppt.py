"""生成 PPT：临港 x WeTest 中日数字经济跨国发展闭门会 · 宣讲与策划"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# 主色
PRIMARY = RGBColor(0x0B, 0x3D, 0x91)      # 深蓝
ACCENT = RGBColor(0xC0, 0x39, 0x2B)       # 红
GOLD = RGBColor(0xD4, 0xA0, 0x4C)         # 金
LIGHT = RGBColor(0xF4, 0xF6, 0xFB)        # 浅蓝
DARK = RGBColor(0x22, 0x22, 0x22)         # 文字深灰
GREY = RGBColor(0x77, 0x77, 0x77)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def set_font(run, name="微软雅黑", size=18, bold=False, color=DARK):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("a:ea"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("a:ea"), {"typeface": name})
        rPr.append(rFonts)
    else:
        rFonts.set("typeface", name)


def add_rect(slide, left, top, width, height, fill=PRIMARY, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
    shape.shadow.inherit = False
    return shape


def add_text(slide, left, top, width, height, text, size=18, bold=False,
             color=DARK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             font="微软雅黑"):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    if isinstance(text, str):
        text = [text]
    for i, t in enumerate(text):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = t
        set_font(run, name=font, size=size, bold=bold, color=color)
    return tb


def add_bullet_list(slide, left, top, width, height, items, size=14,
                    color=DARK, bullet="•", line_spacing=1.25):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    for i, it in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = line_spacing
        run = p.add_run()
        run.text = f"{bullet}  {it}"
        set_font(run, size=size, color=color)
    return tb


def slide_header(slide, title, subtitle=None, page_no=None, total=None):
    # 顶部条
    add_rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.45),
             fill=PRIMARY)
    add_rect(slide, Inches(0), Inches(0.45), Inches(13.333), Inches(0.05),
             fill=GOLD)
    # 标题
    add_text(slide, Inches(0.5), Inches(0.55), Inches(10), Inches(0.6),
             title, size=24, bold=True, color=PRIMARY,
             anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_text(slide, Inches(0.5), Inches(1.05), Inches(12), Inches(0.35),
                 subtitle, size=12, color=GREY, anchor=MSO_ANCHOR.MIDDLE)
    if page_no and total:
        add_text(slide, Inches(11.5), Inches(7.0), Inches(1.5), Inches(0.3),
                 f"{page_no} / {total}", size=10, color=GREY,
                 align=PP_ALIGN.RIGHT)


def make_card(slide, left, top, width, height, title, body_items,
              accent=PRIMARY, title_size=14, body_size=11):
    add_rect(slide, left, top, width, height, fill=LIGHT)
    add_rect(slide, left, top, Inches(0.08), height, fill=accent)
    add_text(slide, left + Inches(0.2), top + Inches(0.1),
             width - Inches(0.3), Inches(0.4),
             title, size=title_size, bold=True, color=accent)
    add_bullet_list(slide, left + Inches(0.2), top + Inches(0.55),
                    width - Inches(0.3), height - Inches(0.6),
                    body_items, size=body_size, color=DARK)


def make_table(slide, left, top, width, height, headers, rows,
               header_color=PRIMARY, font_size=10):
    shape = slide.shapes.add_table(len(rows) + 1, len(headers),
                                   left, top, width, height)
    table = shape.table
    # 表头
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_color
        cell.margin_left = Inches(0.08)
        cell.margin_right = Inches(0.08)
        cell.margin_top = Inches(0.05)
        cell.margin_bottom = Inches(0.05)
        tf = cell.text_frame
        tf.text = ""
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h
        set_font(r, size=font_size + 1, bold=True, color=WHITE)
    # 内容
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = table.cell(ri + 1, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if ri % 2 == 0 else LIGHT
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            cell.margin_top = Inches(0.04)
            cell.margin_bottom = Inches(0.04)
            tf = cell.text_frame
            tf.text = ""
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            r = p.add_run()
            r.text = str(val)
            set_font(r, size=font_size, color=DARK)
    return table


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    TOTAL = 25
    page = [0]

    def new_slide():
        page[0] += 1
        return prs.slides.add_slide(blank), page[0]

    # ========== 1 封面 ==========
    s, p = new_slide()
    # 背景
    add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), fill=PRIMARY)
    # 装饰条
    add_rect(s, Inches(0), Inches(3.6), Inches(13.333), Inches(0.04), fill=GOLD)
    # 主标题
    add_text(s, Inches(0.8), Inches(2.0), Inches(11.5), Inches(1.0),
             "临港  ×  WeTest", size=54, bold=True, color=WHITE,
             align=PP_ALIGN.LEFT)
    add_text(s, Inches(0.8), Inches(3.8), Inches(11.5), Inches(0.8),
             "中日数字经济跨国发展与高质量交付闭门会",
             size=32, bold=True, color=WHITE)
    add_text(s, Inches(0.8), Inches(4.6), Inches(11.5), Inches(0.5),
             "政策护航  ×  质量守护  —— 一站式中日出海/入华解决方案",
             size=16, color=GOLD)
    add_text(s, Inches(0.8), Inches(6.2), Inches(11.5), Inches(0.4),
             "活动策划 · 现场宣讲 · 转化路径  |  V1.0",
             size=12, color=WHITE)
    add_text(s, Inches(0.8), Inches(6.7), Inches(11.5), Inches(0.4),
             "主办：上海科技企业协会   联合主办：临港科技城 × WeTest",
             size=12, color=WHITE)

    # ========== 2 战略逻辑 ==========
    s, p = new_slide()
    slide_header(s, "战略逻辑", "政策红利 × 质量护航 —— 一次活动，三重转化",
                 p, TOTAL)
    make_card(s, Inches(0.5), Inches(1.5), Inches(4), Inches(5.4),
              "1  临港新片区（场）",
              ["数据跨境流动审批",
               "金融创新 / 跨境支付",
               "税收 / 人才落户红利",
               "对日 JETRO 桥梁"], accent=PRIMARY, title_size=18)
    make_card(s, Inches(4.7), Inches(1.5), Inches(4), Inches(5.4),
              "2  姚志勇教授（势）",
              ["中日数字贸易宏观趋势",
               "自贸区跨国商业壁垒",
               "学术权威背书",
               "建立 C 级共识"], accent=GOLD, title_size=18)
    make_card(s, Inches(8.9), Inches(1.5), Inches(4), Inches(5.4),
              "3  WeTest（落地）",
              ["UDT 日本真机 / 海外众测",
               "真金支付 / 弱网 / 合规",
               "CrashSight 海外合规",
               "小程序安全 / 防篡改"], accent=ACCENT, title_size=18)

    # ========== 3 为什么是现在 ==========
    s, p = new_slide()
    slide_header(s, "为什么是现在？",
                 "中日数字经济双向流动加速，跨国质量痛点集中爆发", p, TOTAL)
    add_bullet_list(s, Inches(0.6), Inches(1.6), Inches(12), Inches(5),
                    [
                        "中资互联网/游戏/金融/电商加速对日出海，遇到「数据跨境、本地支付、设备适配、风控封号」四大刚需。",
                        "日资企业入华，对中国独有的小程序生态、移动支付与隐私合规几乎为零经验。",
                        "临港新片区在「数据跨境 + 跨境金融」上的政策窗口期已开启。",
                        "WeTest 是腾讯官方一站式质量平台，可在跨国「合规 + 测试 + 安全」三个维度提供端到端能力。",
                        "→ 现在用 1 场闭门沙龙，把「政策」「学术权威」「技术落地」一次性交付给 25 位 C 级决策人，转化效率最高。",
                    ], size=18, line_spacing=1.45)

    # ========== 4 目标客户画像 ==========
    s, p = new_slide()
    slide_header(s, "目标客户画像", "四大行业 × C 级决策人 × 中日双向流", p, TOTAL)
    make_table(s, Inches(0.5), Inches(1.5), Inches(12.3), Inches(4.5),
               ["行业", "典型痛点", "WeTest 抓手", "决策人"],
               [
                   ["金融", "数据跨境合规 / 风控 / 隐私 / 防篡改",
                    "CrashSight 合规 + 安全加固 + 小程序扫描",
                    "CTO / 合规 VP"],
                   ["游戏", "海外弱网 / 设备碎片化 / 真金支付 / 防作弊",
                    "UDT 日本真机 + 海外众测 + 真金支付",
                    "CTO / 出海总经理"],
                   ["电商/零售", "本地支付 / 防黑产 / 跨境物流 / 风控封号",
                    "海外多渠道支付验收 + 安全网关",
                    "CTO / 海外业务 VP"],
                   ["泛互联网", "本地化 / 多语言 / 网络 / SDK 兼容",
                    "兼容/性能/众测 + 自动化巡检",
                    "CTO / 海外质量负责人"],
               ], font_size=13)

    # ========== 5 活动概览 ==========
    s, p = new_slide()
    slide_header(s, "活动概览", "20-30 人闭门 · 半日 · 临港科技城", p, TOTAL)
    make_table(s, Inches(1.0), Inches(1.6), Inches(11.3), Inches(4.5),
               ["项目", "内容"],
               [
                   ["主题", "中日数字经济跨国发展与高质量交付闭门会"],
                   ["时间", "下午 13:30-16:30 + 16:30-18:00 自由对接 / 1V1 深聊"],
                   ["地点", "临港科技城"],
                   ["形式", "闭门圆桌 + 政策宣讲 + 产业探访 + 案例 + 现场对接"],
                   ["规模", "20-30 人，C 级 / 业务 VP 为主"],
                   ["语言", "中文为主，日英资料同步"],
               ], font_size=14)

    # ========== 6 详细议程 ==========
    s, p = new_slide()
    slide_header(s, "半日议程（含转化触点）",
                 "每个环节都对应一个明确的转化抓手", p, TOTAL)
    make_table(s, Inches(0.3), Inches(1.5), Inches(12.7), Inches(5.5),
               ["时间", "环节", "内容", "转化触点"],
               [
                   ["13:00-13:30", "签到 & 茶歇", "扫码进群 + 1V1 破冰",
                    "采集联系方式 + 行业标签"],
                   ["13:30-14:30", "产业探访", "临港展厅 + 政策初探",
                    "建立 “场” 的政策势能"],
                   ["14:30-15:10", "宏观指引", "姚志勇教授：中日数字贸易 + 临港壁垒",
                    "权威认知背书"],
                   ["15:10-15:50", "技术落地", "WeTest：中日跨国高质量交付",
                    "★ Demo + 案例 + 权益券"],
                   ["15:50-16:30", "闭门圆桌", "议题：双向拓展最大痛点是什么？",
                    "★ 1V1 资源对接 + 合作意向"],
                   ["16:30-18:00", "晚间深聊", "Networking + 高优 1V1",
                    "★ 线索→机会 升级"],
               ], font_size=12)

    # ========== 7 嘉宾阵容 ==========
    s, p = new_slide()
    slide_header(s, "嘉宾阵容", "学术 + 政策 + 技术 + 客户 四维背书", p, TOTAL)
    make_table(s, Inches(0.3), Inches(1.5), Inches(12.7), Inches(5.2),
               ["环节", "拟邀嘉宾", "身份", "分享要点"],
               [
                   ["开场致辞", "临港科技城高管 + 协会秘书长", "主办 / 联合主办",
                    "政策护航 × 质量守护 主线宣讲"],
                   ["宏观指引", "姚志勇教授（或同等权威）", "学术专家",
                    "中日数字贸易趋势 + 临港自贸区跨国壁垒"],
                   ["技术落地", "WeTest 业务线代表", "技术专家",
                    "UDT / 海外众测 / 真金支付 / 合规扫描"],
                   ["圆桌主持", "WeTest 行业总监", "圆桌主持",
                    "四大行业代表表达痛点 + 当场承接"],
                   ["圆桌嘉宾", "金融/游戏/电商/泛互联网 各 1 位 C 级", "客户代表",
                    "现身说法 + 痛点共鸣 + 行业引领"],
               ], font_size=12)

    # ========== 8 ★ 转化漏斗（核心 slide） ==========
    s, p = new_slide()
    slide_header(s, "★ 转化漏斗设计", "从 75 人邀请池到 3 单签约：目标 4% 转化率",
                 p, TOTAL)
    funnel = [
        ("L0 邀请池", "75", PRIMARY),
        ("L1 到场（认知）", "25 · 33%", PRIMARY),
        ("L2 强意向", "15 · 60%", GOLD),
        ("L3 商机 (T+7)", "9 · 60%", GOLD),
        ("L4 POC (T+30)", "5 · 55%", ACCENT),
        ("L5 签单 (T+90)", "3 · 60%", ACCENT),
    ]
    base_w = 11.0
    base_left = 1.2
    top = 1.6
    h = 0.65
    gap = 0.08
    for i, (lbl, val, c) in enumerate(funnel):
        w = base_w - i * 1.4
        left = base_left + (base_w - w) / 2
        add_rect(s, Inches(left), Inches(top + i * (h + gap)),
                 Inches(w), Inches(h), fill=c)
        add_text(s, Inches(left), Inches(top + i * (h + gap)),
                 Inches(w), Inches(h),
                 f"{lbl}      |      {val}",
                 size=16, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(0.6), Inches(6.7), Inches(12), Inches(0.5),
             "目标：T+90 内首单 GMV ≥ 300 万元，ROI ≥ 5 ×",
             size=15, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)

    # ========== 9 会前转化 ==========
    s, p = new_slide()
    slide_header(s, "会前转化", "建立期待 + 锁定需求（T-21 → T0）", p, TOTAL)
    add_bullet_list(s, Inches(0.7), Inches(1.6), Inches(12), Inches(5.2),
                    [
                        "T-21：三方对齐邀约名单（去重 + NDA），WeTest 销售贡献已沟通过的线索。",
                        "T-15：邀请函中嵌入「圆桌议题征集表」 —— 要求受邀人提交 1 个跨国痛点，提前锁定线索画像。",
                        "T-10：实体邀请函 + 伴手礼寄出（C 级），仪式感即话语权。",
                        "T-5：推送《WeTest 中日专项方案白皮书》预读版，建立技术认知。",
                        "T-3：WeTest 销售对 TOP 10 高优客户 1V1 电话预热 + 预约会后 30 分钟深聊。",
                        "T-2：交通接待方案 / 嘉宾介绍卡 / 当日动线图推送，到场率 +10%。",
                    ], size=16, line_spacing=1.4)

    # ========== 10 ★ 会中转化抓手 ==========
    s, p = new_slide()
    slide_header(s, "★ 会中转化抓手", "高密度触点，每 30 分钟一个钩子", p, TOTAL)
    cards = [
        ("扫码进群 + 行业标签",
         ["签到二维码自动打标",
          "企微 1V1 推送对应行业方案",
          "实时观察互动热度"]),
        ("WeTest 体验角",
         ["UDT 日本真机现场遥控",
          "海外众测看板演示",
          "CrashSight 合规面板"]),
        ("限时权益券",
         ["30 天免费 POC 名额",
          "5 万元测试代金券",
          "仅限闭门客户专享"]),
        ("圆桌 1V1 承接",
         ["每位发言客户配 1 名 SA",
          "现场记录需求 → 24h 内 PPT",
          "圆桌即销售场"]),
        ("闭门客户专属通道",
         ["微信/钉钉直连中日团队",
          "+ 临港政策落地顾问",
          "建立长期触达"]),
        ("晚间深聊",
         ["高优客户 1V1 30 分钟",
          "技术咨询券 + 伴手礼",
          "线索→机会 升级"]),
    ]
    col_w = 4.05
    row_h = 2.4
    for i, (t, items) in enumerate(cards):
        r, c = i // 3, i % 3
        left = 0.4 + c * (col_w + 0.2)
        top = 1.5 + r * (row_h + 0.15)
        make_card(s, Inches(left), Inches(top), Inches(col_w), Inches(row_h),
                  t, items, accent=ACCENT, title_size=13, body_size=11)

    # ========== 11 ★ 会后转化 ==========
    s, p = new_slide()
    slide_header(s, "★ 会后转化 · 48 小时黄金期",
                 "活动结束就是销售开始", p, TOTAL)
    make_table(s, Inches(0.3), Inches(1.5), Inches(12.7), Inches(5),
               ["时间", "动作", "A 级（高优）", "B 级（中优）", "C 级（培育）"],
               [
                   ["T+1", "感谢信 + 资料包", "电话 1V1", "微信 1V1", "邮件"],
                   ["T+3", "CRM 录入 + 分级", "完成", "完成", "完成"],
                   ["T+7", "需求澄清会", "完成", "排期", "—"],
                   ["T+15", "方案 PPT 定向推送", "—", "完成", "推送"],
                   ["T+30", "POC / 二次深聊", "POC 启动", "二次深聊", "月度内容"],
                   ["T+60", "POC 中间汇报 / 商务", "进行中", "POC 启动", "季度活动"],
                   ["T+90", "首单签订", "首单签订", "POC 验收", "—"],
               ], font_size=12)
    add_text(s, Inches(0.3), Inches(6.7), Inches(12.7), Inches(0.4),
             "★ 每条 A/B 级线索必须有明确 Owner + 周更新 + 月复盘",
             size=14, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)

    # ========== 12 客户分级 ==========
    s, p = new_slide()
    slide_header(s, "客户分级标准（ABC 三级）",
                 "用清晰的标准让销售团队不在线索上空转", p, TOTAL)
    make_table(s, Inches(0.3), Inches(1.5), Inches(12.7), Inches(4.5),
               ["级别", "判定标准", "目标动作", "负责人"],
               [
                   ["A · 高优",
                    "C 级到场 + 圆桌发言 + 明确预算/项目 + 四大行业",
                    "T+7 需求澄清；T+30 POC；T+90 首单",
                    "行业总监 + 大客户经理"],
                   ["B · 中优",
                    "业务 VP 到场 + 有兴趣未明确预算",
                    "T+15 二深聊；T+45 定方案；T+90 POC",
                    "大客户经理"],
                   ["C · 培育",
                    "代表到场 / 决策权弱 / 暂无项目",
                    "月度内容运营；季度活动",
                    "市场 + 内容运营"],
               ], font_size=13)

    # ========== 13 中日专项产品包 ==========
    s, p = new_slide()
    slide_header(s, "中日专项产品包（销售即买即用）",
                 "把 62 页方案 → 浓缩为 3 个 SKU", p, TOTAL)
    make_table(s, Inches(0.3), Inches(1.5), Inches(12.7), Inches(4.5),
               ["产品包", "包含模块", "目标客户", "起价", "POC"],
               [
                   ["① 日本出海加速包",
                    "UDT 日本真机 + 海外众测（弱网/真金支付）+ CrashSight 合规",
                    "出海日本：金融/游戏/电商/泛互联网",
                    "¥98K/月", "2 周"],
                   ["② 日资入华质量包",
                    "小程序兼容 + 微信隐私合规 + 安全加固",
                    "入华日资：金融/零售",
                    "¥128K/月", "2 周"],
                   ["③ 跨境合规安全包",
                    "数据跨境合规 + 安全网关 + 渗透测试",
                    "金融/电商/政务",
                    "¥168K/项目", "3 周"],
               ], font_size=12)
    add_text(s, Inches(0.3), Inches(6.4), Inches(12.7), Inches(0.6),
             "卖点：即开即用 · 免封号 · 双向闭环 · 合规托底",
             size=18, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)

    # ========== 14 四大行业切入 ==========
    s, p = new_slide()
    slide_header(s, "四大行业精准切入", "每个行业有一个首选方案 + 一个标杆案例",
                 p, TOTAL)
    make_table(s, Inches(0.3), Inches(1.5), Inches(12.7), Inches(4.8),
               ["行业", "主诉求", "首选方案包", "标杆案例", "首单切入"],
               [
                   ["金融", "数据跨境 + 风控 + 隐私",
                    "③ 跨境合规安全包", "银联云小程序安全扫描",
                    "小程序合规 + 数据跨境审计"],
                   ["游戏", "日本首发 + 网络/设备 + 真金支付",
                    "① 日本出海加速包", "跨境电商 UDT + 海外众测",
                    "海外众测 + UDT 日本真机"],
                   ["电商/零售", "本地支付 + 防黑产 + 多端",
                    "① + ③", "跨境电商 UDT",
                    "海外支付验收 + 安全网关"],
                   ["泛互联网", "本地化 + 性能 + SDK 兼容",
                    "① 日本出海加速包", "PerfDog + 海外众测",
                    "PerfDog + 兼容/众测"],
               ], font_size=12)

    # ========== 15 邀约策略 ==========
    s, p = new_slide()
    slide_header(s, "邀约策略", "75 邀请池 → 25 到场 · 4 渠道精准配比",
                 p, TOTAL)
    make_table(s, Inches(0.6), Inches(1.5), Inches(12.1), Inches(4),
               ["渠道", "邀请数", "到场转化率", "到场人数"],
               [
                   ["上海科技企业协会会员", "30", "40%", "12"],
                   ["临港科技城企业池 / 招商线索", "15", "35%", "5"],
                   ["WeTest 中日有出海/入华诉求线索", "20", "30%", "6"],
                   ["合作伙伴/媒体推荐（含 JETRO 等）", "10", "20%", "2"],
                   ["合计", "75", "—", "25"],
               ], font_size=14)
    add_text(s, Inches(0.6), Inches(6.0), Inches(12), Inches(0.5),
             "C 级 / 业务 VP 占比 ≥ 70%（强决策权是高转化前提）",
             size=14, bold=True, color=PRIMARY)

    # ========== 16 资源分工 ==========
    s, p = new_slide()
    slide_header(s, "三方分工", "协会 + 临港 + WeTest", p, TOTAL)
    make_card(s, Inches(0.5), Inches(1.5), Inches(4), Inches(5.2),
              "上海科技企业协会（总策划）",
              ["总策划 / 邀请名单把关",
               "现场主持 / 媒体宣发",
               "复盘报告 / 季度迭代",
               "会员价值放大"], accent=PRIMARY, title_size=16, body_size=13)
    make_card(s, Inches(4.7), Inches(1.5), Inches(4), Inches(5.2),
              "临港科技城（政策与场域）",
              ["场地 / 展厅参观",
               "政策宣讲嘉宾",
               "招商顾问 1V1 对接",
               "落地政策包"], accent=GOLD, title_size=16, body_size=13)
    make_card(s, Inches(8.9), Inches(1.5), Inches(4), Inches(5.2),
              "WeTest（技术与转化）",
              ["技术分享 + Demo",
               "案例与体验角",
               "销售跟进与 POC",
               "签单交付"], accent=ACCENT, title_size=16, body_size=13)

    # ========== 17 时间倒排 ==========
    s, p = new_slide()
    slide_header(s, "时间倒排（T-30 → T+90）", "关键里程碑", p, TOTAL)
    make_table(s, Inches(0.3), Inches(1.5), Inches(12.7), Inches(5.5),
               ["阶段", "时间", "关键动作", "责任方"],
               [
                   ["T-30", "立项", "三方 Kick-off 主题/嘉宾/预算", "协会"],
                   ["T-21", "锁名单", "去重 + NDA 锁定", "协会"],
                   ["T-15", "首轮邀约", "电话 + 邀请函", "三方销售"],
                   ["T-10", "二轮邀约", "实体邀请函 + 二次确认", "协会"],
                   ["T-7", "议程定稿", "嘉宾 PPT + 圆桌议题征集", "WeTest+临港"],
                   ["T-5", "白皮书推送", "技术预读 + 1V1 预热", "WeTest"],
                   ["T-3", "现场彩排", "动线/设备/主持/应急", "三方"],
                   ["T0", "活动执行", "现场转化抓手全部触发", "三方"],
                   ["T+1", "感谢信", "群发 + 1V1 触达", "WeTest"],
                   ["T+7", "需求澄清", "A 级完成需求澄清会", "WeTest"],
                   ["T+30", "POC", "A 级启动 POC / B 级二深聊", "WeTest"],
                   ["T+90", "首单", "首单签订 / 季度复盘", "三方"],
               ], font_size=11)

    # ========== 18 预算 ==========
    s, p = new_slide()
    slide_header(s, "预算明细（参考）", "总预算约 ¥154,500（含 10% 机动）",
                 p, TOTAL)
    make_table(s, Inches(0.3), Inches(1.5), Inches(12.7), Inches(5),
               ["科目", "内容", "数量/标准", "预算 (¥)"],
               [
                   ["场地展厅", "临港会议室+展厅+茶歇", "由临港承担", "0"],
                   ["邀请函/物料", "实体邀请函 + 资料袋", "60 份", "18,000"],
                   ["嘉宾费用", "学术嘉宾出席 + 交通", "1 位", "30,000"],
                   ["伴手礼", "C 级礼盒（含权益券）", "30 × 600", "18,000"],
                   ["茶歇/餐饮", "下午茶 + 晚间轻餐", "30 × 250", "7,500"],
                   ["影像/直播", "摄影摄像 + 精剪", "1 场", "20,000"],
                   ["主持/会务", "外部主持 + 礼仪", "1 场", "12,000"],
                   ["Demo 设备", "日本真机 + 大屏", "1 场", "8,000"],
                   ["交通接待", "C 级 1V1 接送补贴", "10 人", "15,000"],
                   ["会后传播", "媒体通稿 + 自媒体", "1 轮", "12,000"],
                   ["机动 10%", "应急/升级", "—", "14,000"],
                   ["合计", "", "", "154,500"],
               ], font_size=11)

    # ========== 19 KPI ==========
    s, p = new_slide()
    slide_header(s, "KPI · 量化目标", "活动 + 转化 + ROI 三层 KPI", p, TOTAL)
    make_table(s, Inches(0.5), Inches(1.5), Inches(12.3), Inches(5),
               ["维度", "指标", "目标值"],
               [
                   ["邀约", "邀请池规模", "≥ 75 人"],
                   ["邀约", "C 级 / VP 占比", "≥ 70%"],
                   ["到场", "实际到场", "≥ 25 人"],
                   ["到场", "到场率", "≥ 33%"],
                   ["现场", "圆桌发言客户", "≥ 8"],
                   ["现场", "1V1 深聊预约", "≥ 12"],
                   ["转化", "T+7 需求澄清会", "≥ 9"],
                   ["转化", "T+30 POC 立项", "≥ 5"],
                   ["转化", "T+90 首单数", "≥ 3"],
                   ["收入", "首单 GMV", "≥ ¥3,000,000"],
                   ["ROI", "首单 GMV / 预算", "≥ 5 ×"],
                   ["品牌", "媒体露出次数", "≥ 10"],
               ], font_size=12)

    # ========== 20 风险预案 ==========
    s, p = new_slide()
    slide_header(s, "风险预案", "6 类风险 × 预案兜底", p, TOTAL)
    make_table(s, Inches(0.3), Inches(1.5), Inches(12.7), Inches(5),
               ["类别", "风险", "概率", "影响", "预案"],
               [
                   ["到场率", "目标客户临时缺席", "中", "高",
                    "T-2/T-1 二轮确认 + 备选 5 人 + 协会内部直播"],
                   ["嘉宾", "学术嘉宾日程冲突", "低", "高",
                    "锁定 1 位 backup + 预演 PPT"],
                   ["设备", "日本真机网络不稳", "中", "中",
                    "录屏备份 + 双链路网络冗余"],
                   ["转化", "WeTest 销售跟进不及时", "中", "高",
                    "T+1 强制 1V1 SOP + 周例会督办"],
                   ["合规", "闭门内容外泄", "低", "中",
                    "现场 NDA + 无直播 + 协会统一精剪发布"],
                   ["不可抗力", "线下临改线上", "低", "高",
                    "腾讯会议直播 + 物料邮寄"],
               ], font_size=11)

    # ========== 21 复盘机制 ==========
    s, p = new_slide()
    slide_header(s, "复盘机制", "T+3 / T+15 / T+30 / T+90 四阶段", p, TOTAL)
    add_bullet_list(s, Inches(0.8), Inches(1.6), Inches(12), Inches(5),
                    [
                        "T+3 三方内部快速复盘（90 分钟）：到场质量 + 现场执行 + 销售触达完成度。",
                        "T+15 客户视角复盘：抽样回访 5 位 A 级，收集体验改进点。",
                        "T+30 转化中期复盘：POC 进度 + 销售漏斗健康度 + 流失原因分析。",
                        "T+90 转化收单复盘：ROI 计算 + 案例沉淀 + 下一场迭代计划。",
                        "建立季度沙龙机制：以本场为模板，每季度迭代行业 / 主题。",
                    ], size=17, line_spacing=1.5)

    # ========== 22 一句话价值主张 ==========
    s, p = new_slide()
    add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), fill=PRIMARY)
    add_rect(s, Inches(0), Inches(3.6), Inches(13.333), Inches(0.04), fill=GOLD)
    add_text(s, Inches(0.5), Inches(2.4), Inches(12.3), Inches(1.5),
             "「来临港，看政策；用 WeTest，",
             size=48, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.5), Inches(3.8), Inches(12.3), Inches(1.5),
             "跑通日本 / 中国市场最后一公里。」",
             size=48, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.5), Inches(5.8), Inches(12.3), Inches(0.6),
             "—— 中日数字经济跨国发展与高质量交付闭门会 · 一句话主张 ——",
             size=16, color=WHITE, align=PP_ALIGN.CENTER)

    # ========== 23 下一步 Next Step ==========
    s, p = new_slide()
    slide_header(s, "Next Step", "我们需要您的拍板，将策划推进到落地", p, TOTAL)
    add_bullet_list(s, Inches(0.8), Inches(1.6), Inches(12), Inches(5),
                    [
                        "1. 确认主题、议程、规模、日期（建议预留 4-6 周筹备期）。",
                        "2. 三方 Kick-off：协会、临港、WeTest 对齐资源、嘉宾、预算。",
                        "3. 启动邀约：T-21 锁定 75 人邀请池，T-15 首轮触达。",
                        "4. 同步业务包装：将 3 个中日专项产品包 SKU 化，准备销售话术与白皮书。",
                        "5. 建立 CRM 转化看板，所有动作回流统一仪表盘。",
                        "6. 执行后按 7-30-90 节奏推进转化，确保首单在 T+90 内落地。",
                    ], size=16, line_spacing=1.5)

    # ========== 24 团队/分工速览 ==========
    s, p = new_slide()
    slide_header(s, "项目团队建议", "明确 Owner，避免责任稀释", p, TOTAL)
    make_table(s, Inches(0.3), Inches(1.5), Inches(12.7), Inches(5),
               ["角色", "归属", "核心职责", "投入"],
               [
                   ["项目总负责人", "协会", "整体方案 + 三方协调", "30%"],
                   ["邀约负责人", "协会", "邀请名单 + 到场率", "全职"],
                   ["政策嘉宾对接", "临港", "学术嘉宾 + 政策讲师 + 场地", "兼职"],
                   ["技术嘉宾", "WeTest", "PPT + Demo + 体验角", "全职"],
                   ["销售跟进负责人", "WeTest", "T+1 → T+90 转化执行", "全职"],
                   ["市场 / 传播", "协会 + WeTest", "通稿 + 短视频 + 客户证言", "兼职"],
                   ["IT / 设备", "WeTest", "签到系统 + 真机 + 网络", "兼职"],
                   ["CRM / 数据", "WeTest", "线索分级 + 漏斗看板", "全职"],
               ], font_size=12)

    # ========== 25 致谢 ==========
    s, p = new_slide()
    add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), fill=PRIMARY)
    add_rect(s, Inches(0), Inches(3.7), Inches(13.333), Inches(0.04), fill=GOLD)
    add_text(s, Inches(0.5), Inches(2.5), Inches(12.3), Inches(1.5),
             "Thank You.",
             size=80, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.5), Inches(4.2), Inches(12.3), Inches(0.8),
             "临港 × WeTest  中日数字经济跨国发展与高质量交付闭门会",
             size=20, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.5), Inches(5.0), Inches(12.3), Inches(0.6),
             "—— 政策护航 × 质量守护 ——",
             size=16, color=GOLD, align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.5), Inches(6.6), Inches(12.3), Inches(0.4),
             "活动策划组 · V1.0 · 期待与您共建",
             size=12, color=WHITE, align=PP_ALIGN.CENTER)

    out = "/workspace/deliverables/临港x WeTest 中日闭门会_宣讲与策划PPT_V1.0.pptx"
    prs.save(out)
    print(f"Saved: {out}, slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
