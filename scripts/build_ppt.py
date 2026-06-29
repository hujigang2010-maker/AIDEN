# -*- coding: utf-8 -*-
"""
生成《复旦大学住房政策研究中心 × 腾讯云 2026 秋季 AI 主题大会》策划方案 PPT。
以 5·22「2026 人工智能商业化落地峰会」为蓝图，聚焦 Agent / 算力 / Token。
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------------- 主题色 ----------------
NAVY      = RGBColor(0x0B, 0x1F, 0x3A)   # 深海军蓝（背景）
NAVY2     = RGBColor(0x12, 0x2A, 0x4D)
TENCENT   = RGBColor(0x00, 0x6E, 0xFF)   # 腾讯云蓝
FUDAN     = RGBColor(0xC0, 0x1F, 0x2E)   # 复旦红
CYAN      = RGBColor(0x22, 0xC1, 0xE6)   # 科技青
GOLD      = RGBColor(0xF5, 0xB5, 0x0C)   # 强调金
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT     = RGBColor(0xE8, 0xEE, 0xF7)
GREY      = RGBColor(0x9A, 0xA7, 0xBD)
CARD      = RGBColor(0x16, 0x31, 0x57)
CARD2     = RGBColor(0xF4, 0xF7, 0xFC)
INK       = RGBColor(0x1B, 0x2A, 0x44)

FONT = "Microsoft YaHei"   # 终端在 PowerPoint 打开时使用；Linux 预览自动回退到文泉驿

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def set_cjk(run, name=FONT):
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", name)


def add_slide(bg=NAVY):
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    r.fill.solid(); r.fill.fore_color.rgb = bg
    r.line.fill.background()
    r.shadow.inherit = False
    return s


def rect(s, x, y, w, h, color, line=None, line_w=None, shape=MSO_SHAPE.RECTANGLE):
    sp = s.shapes.add_shape(shape, x, y, w, h)
    if color is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = color
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = line_w or Pt(1)
    sp.shadow.inherit = False
    return sp


def text(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         space_after=4, line_spacing=1.0):
    """runs: list of paragraphs; each paragraph is list of (text,size,color,bold)."""
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Pt(2)
    tf.margin_top = tf.margin_bottom = Pt(2)
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        p.line_spacing = line_spacing
        for (t, size, color, bold) in para:
            r = p.add_run(); r.text = t
            r.font.size = Pt(size); r.font.color.rgb = color; r.font.bold = bold
            set_cjk(r)
    return tb


def header(s, kicker, title, idx):
    rect(s, 0, 0, SW, Inches(1.25), NAVY2)
    rect(s, 0, Inches(1.25), SW, Pt(3), TENCENT)
    rect(s, Inches(0.55), Inches(0.34), Pt(6), Inches(0.62), GOLD)
    text(s, Inches(0.75), Inches(0.22), Inches(10.5), Inches(0.9),
         [[(kicker, 12, CYAN, True)], [(title, 26, WHITE, True)]], space_after=2)
    text(s, Inches(11.7), Inches(0.42), Inches(1.2), Inches(0.5),
         [[("%02d" % idx, 24, GOLD, True)]], align=PP_ALIGN.RIGHT)


def footer(s):
    text(s, Inches(0.75), Inches(7.05), Inches(9), Inches(0.35),
         [[("复旦大学住房政策研究中心 · 杨浦区科技企业联合会  ×  腾讯云", 9, GREY, False)]])
    text(s, Inches(10.5), Inches(7.05), Inches(2.1), Inches(0.35),
         [[("2026 秋季 AI 主题大会", 9, GREY, False)]], align=PP_ALIGN.RIGHT)


def chip(s, x, y, w, label, color=TENCENT, h=Inches(0.42)):
    c = rect(s, x, y, w, h, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x, y, w, h, [[(label, 12, WHITE, True)]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return c


def card(s, x, y, w, h, title, lines, accent=TENCENT, tcolor=WHITE, body=LIGHT,
         bg=CARD, tsize=15, bsize=11):
    rect(s, x, y, w, h, bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, y, w, Pt(5), accent, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    runs = [[(title, tsize, tcolor, True)]]
    for ln in lines:
        runs.append([("· " + ln, bsize, body, False)])
    text(s, x + Inches(0.22), y + Inches(0.22), w - Inches(0.44), h - Inches(0.4),
         runs, space_after=4, line_spacing=1.02)


# ============================================================== 01 封面
s = add_slide(NAVY)
rect(s, 0, 0, SW, SH, NAVY)
# 装饰条
rect(s, 0, Inches(5.55), SW, Pt(3), TENCENT)
rect(s, Inches(0.0), Inches(5.62), Inches(4.4), Pt(3), FUDAN)
for i, c in enumerate([TENCENT, CYAN, GOLD]):
    rect(s, Inches(11.0)+Inches(0.55)*i, Inches(0.7), Inches(0.34), Inches(0.34), c,
         shape=MSO_SHAPE.OVAL)
text(s, Inches(0.9), Inches(1.35), Inches(11.5), Inches(0.5),
     [[("联合主办 ｜ 复旦大学住房政策研究中心 · 杨浦区科技企业联合会  ×  腾讯云", 14, CYAN, True)]])
text(s, Inches(0.9), Inches(2.0), Inches(11.7), Inches(2.4),
     [[("2026 秋季 AI 主题大会", 50, WHITE, True)],
      [("Agent · 算力 · Token —— 人工智能商业化落地 2.0", 24, GOLD, True)]],
     space_after=10)
text(s, Inches(0.9), Inches(4.15), Inches(11.5), Inches(1.0),
     [[("以 5·22「2026 人工智能商业化落地峰会」为蓝图  ｜  1 主论坛 + 多分论坛 + 黑客松 + 创意集市", 14, LIGHT, False)],
      [("活动建议时间：2026 年 8—9 月    ｜    建议地点：上海·杨浦", 13, GREY, False)]],
     space_after=6)
text(s, Inches(0.9), Inches(5.75), Inches(11.5), Inches(0.6),
     [[("合作策划方案（建议稿）", 16, WHITE, True)],
      [("呈现物：策划方案 PPT + 执行计划 Excel", 11, GREY, False)]], space_after=4)

# ============================================================== 02 目录
s = add_slide(NAVY); header(s, "CONTENTS", "方案目录", 2); footer(s)
items = [
    ("01  合作回顾与标杆案例", "8 场活动 · 500+ 引荐 · 5·22 峰会"),
    ("02  活动目标与双方优势", "复旦学术品牌 × 腾讯云技术生态"),
    ("03  主题定位与三大核心", "Agent · 算力 · Token"),
    ("04  活动整体架构", "1 主论坛 + 多分论坛"),
    ("05  嘉宾与环节设置", "大咖卡位 · 黑客松 · 集市 · 开放麦"),
    ("06  议程时间表", "全天流程一览"),
    ("07  知识产出与闭环", "白皮书 · 榜单 · 社群"),
    ("08  资源分工与赞助权益", "品牌 · 证书 · 展位"),
    ("09  传播 · 预算 · 倒排期", "影响力与落地保障"),
    ("10  风险预案与下一步", "推进建议"),
]
x0, y0 = Inches(0.75), Inches(1.7)
for i, (t, d) in enumerate(items):
    col, row = i % 2, i // 2
    x = x0 + col * Inches(6.15)
    y = y0 + row * Inches(1.02)
    card(s, x, y, Inches(5.85), Inches(0.86), t, [d], accent=CYAN, tsize=15, bsize=11)

# ============================================================== 03 合作回顾
s = add_slide(NAVY); header(s, "REVIEW · 合作回顾", "半年八场，沉淀深度信任", 3); footer(s)
stats = [("8", "场", "连续联合举办的系列活动\n（去年 11 月 — 今年 5 月）"),
         ("500+", "位", "向腾讯云引荐的行业\n专业人士 / 决策者"),
         ("1", "场", "共同推动 5·22\n人工智能商业化落地峰会"),
         ("100%", "免费", "为腾讯云提供品牌展示、\n证书与展位支持")]
for i, (n, u, d) in enumerate(stats):
    x = Inches(0.75) + i * Inches(3.05)
    rect(s, x, Inches(1.75), Inches(2.8), Inches(2.1), CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, Inches(1.75), Inches(2.8), Pt(5), GOLD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x, Inches(2.0), Inches(2.8), Inches(0.9),
         [[(n, 40, CYAN, True), ("  " + u, 16, WHITE, True)]], align=PP_ALIGN.CENTER)
    text(s, x + Inches(0.2), Inches(2.95), Inches(2.4), Inches(0.9),
         [[(line, 11, LIGHT, False)] for line in d.split("\n")],
         align=PP_ALIGN.CENTER, space_after=2)
card(s, Inches(0.75), Inches(4.15), Inches(11.85), Inches(2.45),
     "合作机制已跑通：可复制、可放大",
     ["资源互补：复旦提供学术公信力、政产学研人脉与场地组织能力；腾讯云提供技术、产品与生态资源。",
      "信任沉淀：8 场活动建立了稳定的双边协作流程与默契，沟通成本低、交付确定性高。",
      "价值验证：5·22 峰会双方收益显著，腾讯云主动提出 8—9 月再度合办——本方案即为承接。",
      "本次升级：在峰会成功模式上做「内容更深、互动更强、产出更实」的 2.0 版本。"],
     accent=TENCENT, tsize=17, bsize=12.5)

# ============================================================== 04 标杆案例 5·22
s = add_slide(NAVY); header(s, "BENCHMARK · 标杆案例", "5·22「2026 人工智能商业化落地峰会」", 4); footer(s)
card(s, Inches(0.75), Inches(1.7), Inches(5.85), Inches(2.3), "峰会概况",
     ["时间：2026 年 5 月 22 日",
      "主办：腾讯云 × 复旦大学住房政策研究中心 联合举办",
      "定位：聚焦人工智能商业化落地的年度行业盛会",
      "公开可查：相关信息已在网络公开传播"], accent=FUDAN, tsize=16)
card(s, Inches(6.75), Inches(1.7), Inches(5.85), Inches(2.3), "我方为腾讯云提供的支持（免费）",
     ["品牌展示：主视觉、主舞台与传播物料中的品牌曝光",
      "权威证书：联合颁发的合作 / 荣誉证书背书",
      "展位支持：现场展位与展示空间",
      "组织保障：嘉宾邀约、现场组织与传播协同"], accent=GOLD, tsize=16)
card(s, Inches(0.75), Inches(4.2), Inches(11.85), Inches(2.4),
     "为何要做 2.0 升级版",
     ["从「单一峰会」升级为「主论坛 + 多分论坛」矩阵，覆盖更多细分人群与议题。",
      "从「听会」升级为「参与」：新增黑客松、创意集市、集章集票、开放麦等强互动环节。",
      "从「一次曝光」升级为「持续闭环」：沉淀白皮书 / 榜单 / 社群，形成可延续的知识资产。"],
     accent=CYAN, tsize=17, bsize=12.5)

# ============================================================== 05 目标与优势
s = add_slide(NAVY); header(s, "GOALS · 目标与优势", "充分发挥复旦 × 腾讯云双重优势", 5); footer(s)
card(s, Inches(0.75), Inches(1.7), Inches(5.85), Inches(2.15),
     "复旦大学住房政策研究中心 · 杨浦科技企业联合会",
     ["学术公信力与权威背书", "政产学研人脉与高质量人群", "会议组织与本地资源整合能力",
      "区域产业（杨浦）落地承接力"], accent=FUDAN, tsize=15)
card(s, Inches(6.75), Inches(1.7), Inches(5.85), Inches(2.15),
     "腾讯云",
     ["大模型 / Agent / 算力技术与产品矩阵", "丰富的开发者与企业客户生态",
      "品牌影响力与传播资源", "云资源、API、Token 等可投入的硬通货"], accent=TENCENT, tsize=15)
text(s, Inches(0.75), Inches(4.05), Inches(11.85), Inches(0.5),
     [[("四大活动目标", 16, GOLD, True)]])
goals = [("品牌共建", "强化双方在 AI 领域的联合行业心智"),
         ("商业落地", "促成技术与场景、供给与需求的真实对接"),
         ("人才生态", "以黑客松汇聚开发者与创新团队"),
         ("知识闭环", "沉淀白皮书 / 榜单 / 案例，持续运营")]
for i, (t, d) in enumerate(goals):
    x = Inches(0.75) + i * Inches(3.0)
    card(s, x, Inches(4.55), Inches(2.8), Inches(1.7), t, [d], accent=CYAN, tsize=14, bsize=11.5)

# ============================================================== 06 主题定位 + 三大核心
s = add_slide(NAVY); header(s, "POSITIONING · 主题定位", "聚焦 Agent · 算力 · Token", 6); footer(s)
text(s, Inches(0.75), Inches(1.55), Inches(11.85), Inches(0.95),
     [[("主题（建议）：", 16, WHITE, True), ("「智能体涌现 · 商业落地」", 22, GOLD, True)],
      [("Slogan：让 Agent 真正干活，让算力与 Token 用在刀刃上", 13, LIGHT, False)]], space_after=6)
cores = [("Agent 智能体", TENCENT,
          ["从 Demo 到生产级落地", "多智能体协作与编排", "企业场景 ROI 与案例"]),
         ("算力 Compute", CYAN,
          ["训练 / 推理算力优化", "成本、调度与弹性", "国产化与异构算力"]),
         ("Token 经济", GOLD,
          ["大模型成本结构与定价", "Token 消耗优化策略", "商业模式与计量计费"])]
for i, (t, c, lines) in enumerate(cores):
    x = Inches(0.75) + i * Inches(4.05)
    card(s, x, Inches(2.75), Inches(3.8), Inches(3.3), t, lines, accent=c, tsize=18, bsize=13)
    rect(s, x + Inches(1.5), Inches(2.95), Inches(0.8), Inches(0.8), c, shape=MSO_SHAPE.OVAL)
    text(s, x + Inches(1.5), Inches(2.95), Inches(0.8), Inches(0.8),
         [[(["A", "C", "T"][i], 28, WHITE, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# ============================================================== 07 整体架构
s = add_slide(NAVY); header(s, "ARCHITECTURE · 活动架构", "1 个主论坛 + 多个分论坛", 7); footer(s)
# 主论坛
rect(s, Inches(3.4), Inches(1.6), Inches(6.5), Inches(0.95), TENCENT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, Inches(3.4), Inches(1.6), Inches(6.5), Inches(0.95),
     [[("主论坛 ｜ 顶级大咖卡位分享 · 趋势发布 · 重磅签约", 15, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
# 连接线
rect(s, Inches(6.6), Inches(2.55), Pt(3), Inches(0.35), CYAN)
subs = [("分论坛 A\nAgent 落地", TENCENT), ("分论坛 B\n算力基础设施", CYAN),
        ("分论坛 C\nToken 与大模型应用", GOLD), ("黑客松\nHackathon", FUDAN)]
for i, (t, c) in enumerate(subs):
    x = Inches(0.75) + i * Inches(3.05)
    rect(s, x, Inches(2.95), Inches(2.8), Inches(1.25), CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x, Inches(2.95), Inches(2.8), Pt(5), c, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x, Inches(2.95), Inches(2.8), Inches(1.25),
         [[(line, 13, WHITE, True)] for line in t.split("\n")],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=2)
# 体验层
rect(s, Inches(0.75), Inches(4.55), Inches(11.85), Inches(1.7), CARD2, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
rect(s, Inches(0.75), Inches(4.55), Inches(11.85), Pt(5), GOLD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, Inches(1.0), Inches(4.7), Inches(11.4), Inches(0.4),
     [[("贯穿全场的体验层", 15, INK, True)]])
exp = ["展览区 / 企业展位", "创意集市", "集票 · 集章打卡", "临时开放麦 Open Mic", "茶歇 · 自由社交"]
for i, e in enumerate(exp):
    x = Inches(1.0) + i * Inches(2.32)
    chip(s, x, Inches(5.45), Inches(2.15), e, color=TENCENT)
text(s, Inches(0.75), Inches(6.45), Inches(11.85), Inches(0.5),
     [[("设计原则：主论坛立势、分论坛深耕、黑客松聚才、体验层引爆现场，四层联动形成知识产出闭环。", 12, GREY, False)]])

# ============================================================== 08 主论坛设计
s = add_slide(NAVY); header(s, "MAIN FORUM · 主论坛", "顶级大咖卡位分享", 8); footer(s)
card(s, Inches(0.75), Inches(1.7), Inches(5.85), Inches(2.5), "嘉宾阵容（卡位邀请方向）",
     ["复旦大学教授 / 院士级学者：AI 与产业趋势主旨演讲",
      "腾讯云技术高管：大模型 / Agent / 算力战略发布",
      "头部企业 CTO / CEO：商业化落地实战案例",
      "知名投资人：AI 赛道资本视角",
      "政府 / 园区领导：政策与产业环境致辞"], accent=FUDAN, tsize=15, bsize=12)
card(s, Inches(6.75), Inches(1.7), Inches(5.85), Inches(2.5), "主论坛环节",
     ["开场致辞（主办双方 + 政府）",
      "主旨演讲 ×2（学术 + 产业）",
      "腾讯云重磅趋势 / 产品发布",
      "高峰对话：Agent 商业化圆桌",
      "战略签约 / 联合倡议 / 颁证仪式"], accent=TENCENT, tsize=15, bsize=12)
card(s, Inches(0.75), Inches(4.4), Inches(11.85), Inches(1.85),
     "「卡位分享」机制说明",
     ["卡位 = 为重点嘉宾 / 赞助方预留主舞台黄金时段与议题位置，形成强曝光与强背书。",
      "每位卡位嘉宾配套：主视觉露出 + 演讲 15—20 分钟 + 官方证书 + 传播通稿 + 集锦视频。",
      "腾讯云作为联合主办享有主论坛核心卡位与发布窗口；复旦提供学术卡位与权威背书。"],
     accent=GOLD, tsize=16, bsize=12.5)

# ============================================================== 09 分论坛设计
s = add_slide(NAVY); header(s, "SUB-FORUMS · 分论坛", "三大平行分论坛，深耕细分", 9); footer(s)
subf = [("分论坛 A · Agent 落地", TENCENT,
         ["主题：从 PoC 到生产", "形式：4 场演讲 + 1 圆桌", "人群：企业技术 / 产品负责人", "产出：Agent 落地案例集"]),
        ("分论坛 B · 算力基础设施", CYAN,
         ["主题：算力降本增效", "形式：4 场演讲 + Demo", "人群：架构师 / 运维 / CTO", "产出：算力优化实践指南"]),
        ("分论坛 C · Token 与应用", GOLD,
         ["主题：Token 经济与定价", "形式：4 场演讲 + 路演", "人群：创业者 / 投资人", "产出：商业模式白皮书章节"])]
for i, (t, c, lines) in enumerate(subf):
    x = Inches(0.75) + i * Inches(4.05)
    card(s, x, Inches(1.75), Inches(3.8), Inches(3.0), t, lines, accent=c, tsize=15.5, bsize=12)
card(s, Inches(0.75), Inches(4.95), Inches(11.85), Inches(1.3),
     "排期建议",
     ["上午主论坛聚势；下午三个分论坛平行进行，参会者按兴趣自由选择，黑客松全天并行。",
      "每个分论坛配独立主持、速记与现场金句采集，便于会后内容沉淀与二次传播。"],
     accent=FUDAN, tsize=15, bsize=12)

# ============================================================== 10 黑客松
s = add_slide(NAVY); header(s, "HACKATHON · 黑客松", "Agent 实战黑客松挑战赛", 10); footer(s)
card(s, Inches(0.75), Inches(1.7), Inches(5.85), Inches(2.4), "赛制设计",
     ["赛题：基于腾讯云能力构建可落地 Agent 应用",
      "赛道：效率办公 / 行业垂类 / 创意应用",
      "形式：赛前招募 → 现场 8—12h 冲刺 → 路演评审",
      "规模：建议 20—30 支队伍 / 80—120 人"], accent=FUDAN, tsize=15, bsize=12)
card(s, Inches(6.75), Inches(1.7), Inches(5.85), Inches(2.4), "资源与激励",
     ["腾讯云提供：算力 / API / Token 额度 + 技术导师",
      "复旦提供：场地、评委、学术指导、媒体",
      "奖项：冠亚季军 + 最佳创意 / 最佳商业潜力",
      "奖励：奖金 + 云资源代金券 + 孵化对接"], accent=TENCENT, tsize=15, bsize=12)
text(s, Inches(0.75), Inches(4.3), Inches(11.85), Inches(0.45),
     [[("评审维度", 15, GOLD, True)]])
dims = ["创新性", "技术完成度", "商业可行性", "落地价值", "现场路演"]
for i, d in enumerate(dims):
    x = Inches(0.75) + i * Inches(2.37)
    chip(s, x, Inches(4.85), Inches(2.2), d, color=CYAN)
card(s, Inches(0.75), Inches(5.55), Inches(11.85), Inches(0.95),
     "价值",
     ["以赛聚才：汇聚开发者生态，沉淀腾讯云技术心智，优秀作品可纳入会后案例与孵化对接。"],
     accent=TENCENT, tsize=14, bsize=12)

# ============================================================== 11 互动体验
s = add_slide(NAVY); header(s, "EXPERIENCE · 互动体验", "展览 · 集市 · 集章 · 开放麦", 11); footer(s)
exps = [("展览区 & 企业展位", TENCENT,
         ["腾讯云能力展示区 + 合作企业展位", "标准 / 精品展位分级", "Demo 体验 + 现场答疑"]),
        ("创意集市", GOLD,
         ["AI 文创 / 周边 / 开发者好物", "初创团队产品试用摊位", "轻松氛围促进自由交流"]),
        ("集票 · 集章打卡", CYAN,
         ["参会护照 + 打卡地图", "逛展 / 听会 / 互动盖章集票", "集满兑换礼品 / 抽奖资格"]),
        ("临时开放麦 Open Mic", FUDAN,
         ["随时上台 3—5 分钟即兴分享", "项目自荐 / 观点碰撞 / 招募", "促进即兴交流与连接"])]
for i, (t, c, lines) in enumerate(exps):
    col, row = i % 2, i // 2
    x = Inches(0.75) + col * Inches(6.1)
    y = Inches(1.7) + row * Inches(2.35)
    card(s, x, y, Inches(5.85), Inches(2.15), t, lines, accent=c, tsize=16, bsize=12)

# ============================================================== 12 议程时间表
s = add_slide(NAVY); header(s, "AGENDA · 议程", "全天议程时间表（建议）", 12); footer(s)
rows = [
    ("09:00–09:30", "签到 · 参会护照领取 · 展区开放", "全体"),
    ("09:30–10:00", "开场致辞（主办双方 + 政府 / 园区）", "主论坛"),
    ("10:00–11:10", "主旨演讲：学术趋势 + 腾讯云战略发布", "主论坛"),
    ("11:10–12:00", "高峰对话 + 战略签约 / 颁证仪式", "主论坛"),
    ("12:00–13:30", "午餐 · 创意集市 · 展位巡览 · 集章打卡", "体验层"),
    ("13:30–17:00", "分论坛 A/B/C 平行进行（演讲 + 圆桌 + 路演）", "分论坛"),
    ("10:00–18:00", "黑客松：现场冲刺（全天并行）", "黑客松"),
    ("16:00–17:30", "临时开放麦 Open Mic（体验区舞台）", "体验层"),
    ("17:30–18:30", "黑客松路演 + 颁奖 + 集票抽奖", "收尾"),
    ("18:30–", "自由交流 · 合影 · 散场", "全体"),
]
y = Inches(1.65)
rect(s, Inches(0.75), y, Inches(11.85), Inches(0.5), TENCENT)
for cx, w, t in [(Inches(0.75), Inches(2.4), "时间"), (Inches(3.15), Inches(7.4), "环节"), (Inches(10.55), Inches(2.05), "板块")]:
    text(s, cx + Inches(0.15), y, w, Inches(0.5), [[(t, 13, WHITE, True)]], anchor=MSO_ANCHOR.MIDDLE)
y = y + Inches(0.5)
for i, (tm, ev, bl) in enumerate(rows):
    bg = CARD if i % 2 == 0 else NAVY2
    rect(s, Inches(0.75), y, Inches(11.85), Inches(0.45), bg)
    text(s, Inches(0.9), y, Inches(2.3), Inches(0.45), [[(tm, 11, GOLD, True)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(3.3), y, Inches(7.2), Inches(0.45), [[(ev, 11, LIGHT, False)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(10.7), y, Inches(1.9), Inches(0.45), [[(bl, 11, CYAN, True)]], anchor=MSO_ANCHOR.MIDDLE)
    y = y + Inches(0.45)

# ============================================================== 13 知识产出闭环
s = add_slide(NAVY); header(s, "OUTPUT · 知识产出", "形成有效知识产出与闭环", 13); footer(s)
loop = [("会前", CYAN, ["议题征集 / 嘉宾招募", "黑客松赛队招募", "预热传播 + 报名沉淀"]),
        ("会中", TENCENT, ["主 / 分论坛内容生产", "现场金句 / Demo 采集", "黑客松作品产出"]),
        ("会后", GOLD, ["联合白皮书 / 案例集", "AI 落地榜单发布", "社群运营 + 商机跟进"])]
for i, (t, c, lines) in enumerate(loop):
    x = Inches(0.75) + i * Inches(4.05)
    card(s, x, Inches(1.8), Inches(3.5), Inches(2.6), t, lines, accent=c, tsize=18, bsize=13)
    if i < 2:
        text(s, x + Inches(3.55), Inches(2.6), Inches(0.5), Inches(0.8),
             [[("→", 30, WHITE, True)]], align=PP_ALIGN.CENTER)
card(s, Inches(0.75), Inches(4.65), Inches(11.85), Inches(1.7),
     "可沉淀的知识资产",
     ["《复旦 × 腾讯云 AI 商业化落地白皮书 2026（秋）》——双方联合署名，权威背书。",
      "「AI 落地优秀案例 / 黑客松作品」榜单与案例库，可持续运营与二次传播。",
      "高质量参会者社群与商机池，为下一场活动与长期合作蓄水。"],
     accent=FUDAN, tsize=16, bsize=12.5)

# ============================================================== 14 资源分工
s = add_slide(NAVY); header(s, "ROLES · 资源分工", "双方资源投入与分工建议", 14); footer(s)
y = Inches(1.65)
rect(s, Inches(0.75), y, Inches(11.85), Inches(0.5), NAVY2)
cols = [(Inches(0.75), Inches(3.0), "板块"), (Inches(3.75), Inches(4.45), "复旦中心 / 联合会"), (Inches(8.2), Inches(4.4), "腾讯云")]
for cx, w, t in cols:
    text(s, cx + Inches(0.15), y, w, Inches(0.5), [[(t, 13, CYAN, True)]], anchor=MSO_ANCHOR.MIDDLE)
data = [
    ("主办背书", "学术品牌 / 权威证书 / 政府对接", "技术品牌 / 生态资源"),
    ("嘉宾", "学者 / 政产学研人脉", "技术高管 / 客户案例"),
    ("场地组织", "场地 / 现场执行 / 志愿者", "技术布展 / Demo"),
    ("黑客松", "评委 / 学术指导 / 场地", "算力 / API / Token / 导师 / 奖励"),
    ("传播", "校友 / 本地媒体 / 社群", "官方渠道 / 开发者社区"),
    ("产出", "白皮书学术编纂", "技术内容 / 案例供给"),
]
y = y + Inches(0.5)
for i, (a, b, c) in enumerate(data):
    bg = CARD if i % 2 == 0 else NAVY2
    rect(s, Inches(0.75), y, Inches(11.85), Inches(0.62), bg)
    text(s, Inches(0.9), y, Inches(2.8), Inches(0.62), [[(a, 12, GOLD, True)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(3.9), y, Inches(4.3), Inches(0.62), [[(b, 11, LIGHT, False)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(8.35), y, Inches(4.25), Inches(0.62), [[(c, 11, LIGHT, False)]], anchor=MSO_ANCHOR.MIDDLE)
    y = y + Inches(0.62)

# ============================================================== 15 赞助权益
s = add_slide(NAVY); header(s, "SPONSORSHIP · 赞助权益", "招商与赞助权益（延续 5·22 模式）", 15); footer(s)
tiers = [("联合主办 / 首席", FUDAN, ["主论坛核心卡位 + 发布窗口", "主视觉首位品牌露出", "联合颁发权威证书", "精品展位 + 创意集市位",
                                "黑客松冠名权", "白皮书联合署名"]),
         ("钻石赞助", TENCENT, ["分论坛冠名 + 演讲席位", "主视觉品牌露出", "荣誉证书", "标准展位",
                            "黑客松导师席", "传播通稿"]),
         ("生态合作", CYAN, ["展位 / 集市摊位", "Logo 露出", "参会证书", "开放麦优先席", "社群推荐"])]
for i, (t, c, lines) in enumerate(tiers):
    x = Inches(0.75) + i * Inches(4.05)
    card(s, x, Inches(1.7), Inches(3.8), Inches(3.7), t, lines, accent=c, tsize=16, bsize=11.5)
card(s, Inches(0.75), Inches(5.6), Inches(11.85), Inches(0.95),
     "说明",
     ["延续 5·22 为腾讯云提供品牌展示 / 证书 / 展位的成功做法；本次以分级权益吸引更多生态伙伴共建，分摊成本、放大声量。"],
     accent=GOLD, tsize=14, bsize=12)

# ============================================================== 16 传播
s = add_slide(NAVY); header(s, "PROMOTION · 传播", "全周期传播与影响力", 16); footer(s)
phases = [("预热期（T-30→T-7）", CYAN, ["双方官方渠道联合预告", "嘉宾 / 议题海报矩阵", "报名 H5 + 黑客松招募", "校友 / 社群 / 媒体扩散"]),
          ("爆发期（活动当天）", GOLD, ["现场直播 + 图文快讯", "金句卡片实时产出", "KOL / 媒体现场报道", "话题与短视频传播"]),
          ("长尾期（T+1→T+30）", TENCENT, ["白皮书 / 榜单发布", "演讲实录 + 集锦视频", "案例稿 + 复盘报道", "社群持续运营"])]
for i, (t, c, lines) in enumerate(phases):
    x = Inches(0.75) + i * Inches(4.05)
    card(s, x, Inches(1.8), Inches(3.8), Inches(2.9), t, lines, accent=c, tsize=15.5, bsize=12)
card(s, Inches(0.75), Inches(4.95), Inches(11.85), Inches(1.3),
     "传播目标（建议 KPI）",
     ["到场 ≥ 500 人；线上观看 ≥ 50,000 人次；媒体 / 自媒体报道 ≥ 30 篇；黑客松报名 ≥ 25 队；沉淀名单 ≥ 1,000。"],
     accent=FUDAN, tsize=15, bsize=12.5)

# ============================================================== 17 预算概览
s = add_slide(NAVY); header(s, "BUDGET · 预算", "预算概览（建议区间，单位：万元）", 17); footer(s)
budget = [("场地 & 搭建", "8 – 15"), ("舞台 / 视觉 / 设备", "6 – 12"), ("嘉宾差旅接待", "4 – 8"),
          ("黑客松（奖金 / 算力 / 运营）", "6 – 12"), ("餐饮 & 茶歇", "5 – 10"),
          ("传播 & 物料", "5 – 10"), ("互动体验 / 礼品", "3 – 6"), ("执行 & 不可预见", "4 – 8")]
y = Inches(1.7)
for i, (a, b) in enumerate(budget):
    col, row = i % 2, i // 2
    x = Inches(0.75) + col * Inches(6.1)
    yy = y + row * Inches(0.78)
    rect(s, x, yy, Inches(5.85), Inches(0.64), CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x + Inches(0.25), yy, Inches(3.9), Inches(0.64), [[(a, 13, LIGHT, True)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, x + Inches(4.0), yy, Inches(1.7), Inches(0.64), [[(b, 14, GOLD, True)]], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
rect(s, Inches(0.75), Inches(5.7), Inches(11.85), Inches(0.85), TENCENT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, Inches(1.0), Inches(5.7), Inches(7.5), Inches(0.85),
     [[("预算总区间（参考）", 16, WHITE, True)]], anchor=MSO_ANCHOR.MIDDLE)
text(s, Inches(8.5), Inches(5.7), Inches(3.85), Inches(0.85),
     [[("约 41 – 81 万元", 20, WHITE, True)]], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
text(s, Inches(0.75), Inches(6.65), Inches(11.85), Inches(0.4),
     [[("注：可通过赞助权益分级、生态共建与资源置换（如腾讯云算力 / Token）显著降低净支出。", 11, GREY, False)]])

# ============================================================== 18 倒排期
s = add_slide(NAVY); header(s, "TIMELINE · 倒排期", "筹备里程碑（倒排）", 18); footer(s)
ms = [("T-8 周", "立项 · 双方确认主题 / 规模 / 预算 / 分工"),
      ("T-7 周", "确定时间地点 · 启动招商 · 嘉宾初邀"),
      ("T-6 周", "主视觉与议程框架 · 黑客松赛题与招募启动"),
      ("T-4 周", "嘉宾确认 · 赞助签约 · 传播预热上线"),
      ("T-3 周", "议程定稿 · 报名开放 · 物料制作"),
      ("T-2 周", "黑客松选手确认 · 现场动线与彩排方案"),
      ("T-1 周", "全流程彩排 · 物料到位 · 应急预案确认"),
      ("T-0", "活动执行 · 现场指挥 · 直播传播"),
      ("T+2 周", "内容沉淀 · 白皮书 / 榜单 · 复盘与商机跟进")]
y = Inches(1.7)
for i, (t, d) in enumerate(ms):
    col, row = i % 3, i // 3
    x = Inches(0.75) + col * Inches(4.05)
    yy = y + row * Inches(1.5)
    card(s, x, yy, Inches(3.8), Inches(1.3), t, [d], accent=[CYAN, TENCENT, GOLD, FUDAN][i % 4], tsize=15, bsize=11.5)

# ============================================================== 19 风险预案
s = add_slide(NAVY); header(s, "RISK · 风险预案", "风险识别与应对", 19); footer(s)
risks = [("嘉宾变动", "核心嘉宾设 A/B 备选，提前确认行程与备稿"),
         ("到场不足", "精准定向邀约 + 报名审核 + 预约提醒机制"),
         ("黑客松参与度", "提前招募 + 算力激励 + 导师陪跑 + 孵化对接"),
         ("现场动线 / 安全", "分区动线设计 + 安保 / 医疗 + 应急通道"),
         ("天气 / 不可抗力", "室内主场 + 备用时间窗 + 线上同步直播"),
         ("预算超支", "分级赞助 + 资源置换 + 预留不可预见费")]
for i, (t, d) in enumerate(risks):
    col, row = i % 2, i // 2
    x = Inches(0.75) + col * Inches(6.1)
    y = Inches(1.75) + row * Inches(1.55)
    card(s, x, y, Inches(5.85), Inches(1.35), t, [d], accent=FUDAN, tsize=15, bsize=12)

# ============================================================== 20 下一步 / 结语
s = add_slide(NAVY); header(s, "NEXT · 下一步", "推进建议与下一步", 20); footer(s)
steps = [("1 · 共识", "确认主题、时间（8—9 月）、规模与预算框架"),
         ("2 · 立项", "成立联合筹备组，明确双方对接人与职责"),
         ("3 · 启动", "锁定场地与日期，启动嘉宾邀约与招商"),
         ("4 · 落地", "按倒排期推进，定期联合例会对齐进度")]
for i, (t, d) in enumerate(steps):
    x = Inches(0.75) + i * Inches(3.0)
    card(s, x, Inches(1.8), Inches(2.8), Inches(1.9), t, [d], accent=CYAN, tsize=16, bsize=12)
rect(s, Inches(0.75), Inches(4.1), Inches(11.85), Inches(2.2), CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
rect(s, Inches(0.75), Inches(4.1), Inches(11.85), Pt(5), GOLD, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, Inches(1.1), Inches(4.45), Inches(11.2), Inches(1.6),
     [[("期待与腾讯云再度携手", 24, WHITE, True)],
      [("在 8—9 月共同打造一场「内容更深、互动更强、产出更实」的 AI 主题大会，", 15, LIGHT, False)],
      [("延续 5·22 峰会的成功，把复旦 × 腾讯云的联合品牌做成行业标杆。", 15, LIGHT, False)]],
     space_after=8)
text(s, Inches(1.1), Inches(6.0), Inches(11.2), Inches(0.4),
     [[("配套呈现：本策划方案 PPT + 执行计划 Excel（议程 / 嘉宾 / 黑客松 / 预算 / 倒排期 / 物料 / KPI）", 12, CYAN, True)]])

# ---------------- 保存 ----------------
out_dir = "/workspace/deliverables"
os.makedirs(out_dir, exist_ok=True)
path = os.path.join(out_dir, "复旦x腾讯云_2026秋季AI大会_策划方案.pptx")
prs.save(path)
print("Saved:", path, "slides:", len(prs.slides._sldIdLst))
