#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""创智汇 A4 楼书宣传册（PDF）— 沟通 / 楼盘介绍用
用法: python3 build_brochure_a4.py
"""

from pathlib import Path

from reportlab.lib.colors import Color, HexColor, white, black
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

HERE = Path(__file__).resolve().parent
ART = Path("/opt/cursor/artifacts")
FONT_PATH = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"

# 注册字体
pdfmetrics.registerFont(TTFont("CN", FONT_PATH, subfontIndex=0))
FONT = "CN"

# 品牌色
PURPLE = HexColor("#1A1140")
PURPLE2 = HexColor("#3A206E")
PURPLE3 = HexColor("#2A1E55")
GOLD = HexColor("#C9A46A")
GOLD2 = HexColor("#E9C27C")
INK = HexColor("#1A1140")
MUTED = HexColor("#5A5470")
LIGHT = HexColor("#F7F4FC")
LINE = HexColor("#D9D0E8")
ACCENT = HexColor("#6B5BCC")
GREEN = HexColor("#2E8B6A")

W, H = A4  # 595.27 x 841.89


def draw_bg(c, tone="light"):
    if tone == "dark":
        c.setFillColor(PURPLE)
        c.rect(0, 0, W, H, fill=1, stroke=0)
        c.setFillColor(PURPLE2)
        c.setFillColorRGB(0.23, 0.125, 0.43, alpha=0.55)
        c.circle(-30, H + 40, 180, fill=1, stroke=0)
        c.setFillColorRGB(0.45, 0.28, 0.75, alpha=0.25)
        c.circle(W + 40, 120, 160, fill=1, stroke=0)
    else:
        c.setFillColor(LIGHT)
        c.rect(0, 0, W, H, fill=1, stroke=0)
        # top brand bar
        c.setFillColor(PURPLE)
        c.rect(0, H - 14 * mm, W, 14 * mm, fill=1, stroke=0)
        c.setFillColor(GOLD)
        c.rect(0, H - 15.2 * mm, W, 1.2 * mm, fill=1, stroke=0)


def footer(c, page, total, dark=False):
    c.setFillColor(GOLD if dark else MUTED)
    c.setFont(FONT, 8)
    c.drawString(16 * mm, 8 * mm, "上海创智汇 · AI+数字内容无界共创港")
    c.drawRightString(W - 16 * mm, 8 * mm, f"{page} / {total}")
    if not dark:
        c.setStrokeColor(LINE)
        c.setLineWidth(0.6)
        c.line(16 * mm, 12 * mm, W - 16 * mm, 12 * mm)


def section_title(c, en, zh, y):
    c.setFillColor(GOLD)
    c.setFont(FONT, 9)
    c.drawString(16 * mm, y, en)
    c.setFillColor(INK)
    c.setFont(FONT, 18)
    c.drawString(16 * mm, y - 8 * mm, zh)
    c.setStrokeColor(GOLD)
    c.setLineWidth(2)
    c.line(16 * mm, y - 11 * mm, 42 * mm, y - 11 * mm)
    return y - 18 * mm


def card(c, x, y, w, h, title=None, fill=white, border=LINE):
    c.setFillColor(fill)
    c.setStrokeColor(border)
    c.setLineWidth(0.8)
    c.roundRect(x, y - h, w, h, 4, fill=1, stroke=1)
    if title:
        c.setFillColor(PURPLE)
        c.setFont(FONT, 11)
        c.drawString(x + 4 * mm, y - 7 * mm, title)


def bullet(c, x, y, text, size=9, color=MUTED, max_w=160 * mm):
    c.setFillColor(GOLD)
    c.setFont(FONT, size)
    c.drawString(x, y, "●")
    c.setFillColor(color)
    # simple wrap
    chars = list(text)
    line, lines = "", []
    for ch in chars:
        trial = line + ch
        if c.stringWidth(trial, FONT, size) > max_w:
            lines.append(line)
            line = ch
        else:
            line = trial
    if line:
        lines.append(line)
    for i, ln in enumerate(lines):
        c.drawString(x + 4 * mm, y - i * (size + 3), ln)
    return len(lines) * (size + 3) + 2


# ===================== pages =====================

def page_cover(c, page, total):
    draw_bg(c, "dark")
    # gold accent bar
    c.setFillColor(GOLD)
    c.rect(0, H - 3 * mm, W, 3 * mm, fill=1, stroke=0)

    c.setFillColor(GOLD2)
    c.setFont(FONT, 11)
    c.drawString(22 * mm, H - 35 * mm, "PROPERTY BROCHURE  ·  A4")

    c.setFillColor(white)
    c.setFont(FONT, 32)
    c.drawString(22 * mm, H - 55 * mm, "上海创智汇")
    c.setFont(FONT, 16)
    c.setFillColor(GOLD2)
    c.drawString(22 * mm, H - 68 * mm, "AI + 数字内容无界共创港")

    c.setStrokeColor(GOLD)
    c.setLineWidth(1.5)
    c.line(22 * mm, H - 76 * mm, 70 * mm, H - 76 * mm)

    c.setFillColor(HexColor("#C6BAEA"))
    c.setFont(FONT, 11)
    for i, t in enumerate([
        "五角场片区城市更新 · 创智天地核心圈",
        "3F 孵化办公 ≈2850㎡　｜　5F 展厅贸易 ≈3670㎡",
        "合计约 6600㎡　·　产业空间招商楼书",
    ]):
        c.drawString(22 * mm, H - 90 * mm - i * 7 * mm, t)

    # key stats
    stats = [
        ("6600㎡", "合作面积"),
        ("2 层", "主力楼层"),
        ("五角场", "城市副中心"),
        ("YOUNG", "政策核心圈"),
    ]
    for i, (n, l) in enumerate(stats):
        x = 22 * mm + i * 42 * mm
        y = 55 * mm
        c.setFillColor(HexColor("#2A1E55"))
        c.roundRect(x, y, 38 * mm, 28 * mm, 3, fill=1, stroke=0)
        c.setStrokeColor(GOLD)
        c.setLineWidth(0.8)
        c.roundRect(x, y, 38 * mm, 28 * mm, 3, fill=0, stroke=1)
        c.setFillColor(GOLD2)
        c.setFont(FONT, 14)
        c.drawCentredString(x + 19 * mm, y + 14 * mm, n)
        c.setFillColor(HexColor("#C6BAEA"))
        c.setFont(FONT, 8)
        c.drawCentredString(x + 19 * mm, y + 6 * mm, l)

    c.setFillColor(HexColor("#9486C4"))
    c.setFont(FONT, 8)
    c.drawString(22 * mm, 22 * mm, "业主方协同 · 中建四局体系　｜　沟通展示版　｜　价格与政策以当期洽谈及政府文件为准")
    footer(c, page, total, dark=True)


def page_overview(c, page, total):
    draw_bg(c)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 8)
    c.drawString(16 * mm, H - 9 * mm, "SHANGHAI CHUANGZHIHUI")
    c.setFillColor(white)
    c.setFont(FONT, 9)
    c.drawRightString(W - 16 * mm, H - 9 * mm, "楼书 · 项目总览")

    y = section_title(c, "01  PROJECT OVERVIEW", "项目总览", H - 28 * mm)

    # intro card
    c.setFillColor(white)
    c.setStrokeColor(LINE)
    c.roundRect(16 * mm, y - 38 * mm, W - 32 * mm, 38 * mm, 4, fill=1, stroke=1)
    c.setFillColor(PURPLE)
    c.setFont(FONT, 11)
    c.drawString(20 * mm, y - 8 * mm, "定位一句话")
    c.setFillColor(MUTED)
    c.setFont(FONT, 9.5)
    lines = [
        "以人工智能与数字内容为产业主轴，打造「上海创智汇 · AI+数字内容无界共创港」——",
        "环同济经济圈超级链接器：IP 为内容、AI 为工具、空间为载体，服务科创企业入驻、",
        "展示交易、政策兑现与出海成长。",
    ]
    for i, ln in enumerate(lines):
        c.drawString(20 * mm, y - 16 * mm - i * 5.5 * mm, ln)
    y -= 46 * mm

    # info table
    rows = [
        ("项目名称", "上海创智汇（一期）OPC + AI + IP 创新中心"),
        ("项目位置", "上海市杨浦区五角场 · 创智天地片区"),
        ("建筑面积", "约 6600㎡（3 楼 ≈2850㎡ + 5 楼 ≈3670㎡）"),
        ("产业主轴", "人工智能 · 数字内容 · OPC · IP 潮玩文创"),
        ("空间属性", "产业办公 + 孵化加速 + 展贸交易 + 社群培训"),
        ("政策卡位", "杨浦 YOUNG立方内容集聚区核心圈"),
        ("业主协同", "中建四局体系（城市更新 / 产业空间）"),
        ("适合客群", "AI/AIGC、微短剧、数字IP、潮玩玩具、科技型中小、成果转化"),
    ]
    for i, (k, v) in enumerate(rows):
        yy = y - i * 9 * mm
        c.setFillColor(PURPLE3 if i % 2 == 0 else white)
        c.rect(16 * mm, yy - 6.5 * mm, W - 32 * mm, 9 * mm, fill=1, stroke=0)
        c.setFillColor(GOLD if i % 2 == 0 else PURPLE)
        c.setFont(FONT, 9)
        c.drawString(20 * mm, yy - 3.5 * mm, k)
        c.setFillColor(white if i % 2 == 0 else INK)
        c.drawString(52 * mm, yy - 3.5 * mm, v)

    y = y - len(rows) * 9 * mm - 10 * mm
    c.setFillColor(PURPLE)
    c.setFont(FONT, 11)
    c.drawString(16 * mm, y, "本册目录")
    toc = [
        "02  区位交通　　五角场城市副中心 · 轨交商圈 · 高校环绕",
        "03  楼层与价格　3F/5F 产品切分 · 参考租金与组合",
        "04  优惠政策　　YOUNG立方 / 长阳秀带 / 高企专精特新等",
        "05  业主介绍　　中建四局体系与项目价值",
        "06  楼层平面图　3F 孵化办公 · 5F 展厅贸易示意",
        "07  入驻指南　　客群适配 · 联系方式",
    ]
    for i, t in enumerate(toc):
        c.setFillColor(MUTED)
        c.setFont(FONT, 9.5)
        c.drawString(20 * mm, y - 8 * mm - i * 6.5 * mm, t)

    footer(c, page, total)


def page_location(c, page, total):
    draw_bg(c)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 8)
    c.drawString(16 * mm, H - 9 * mm, "LOCATION & TRANSIT")
    c.setFillColor(white)
    c.setFont(FONT, 9)
    c.drawRightString(W - 16 * mm, H - 9 * mm, "楼书 · 区位交通")

    y = section_title(c, "02  LOCATION", "区位交通", H - 28 * mm)

    # map-like panel
    c.setFillColor(PURPLE)
    c.roundRect(16 * mm, y - 52 * mm, W - 32 * mm, 52 * mm, 4, fill=1, stroke=0)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 12)
    c.drawString(22 * mm, y - 10 * mm, "杨浦 · 五角场城市副中心")
    c.setFillColor(white)
    c.setFont(FONT, 10)
    c.drawString(22 * mm, y - 20 * mm, "创智天地片区 · 环同济经济圈核心")
    c.setFillColor(HexColor("#C6BAEA"))
    c.setFont(FONT, 9)
    for i, t in enumerate([
        "东临大学路创新街区 · 南接五角场商圈 · 北依江湾知识社区",
        "正处 YOUNG立方「一楼（V聚场）· 一街（大学路）· 一园（B站新世代）· 一区」核心腹地",
        "复旦、同济、上海财大、上理工高校群环绕，人才与创业项目密度领先",
    ]):
        c.drawString(22 * mm, y - 30 * mm - i * 6 * mm, "▸  " + t)
    y -= 60 * mm

    # transit cards
    cards = [
        ("轨交 TOD", [
            "邻近地铁 10 号线（五角场 / 江湾体育场方向）",
            "多线路换乘可达全市主要商务区",
            "适合客户到访、人才通勤、活动集散",
        ]),
        ("商圈配套", [
            "五角场万达 / 百联等成熟商圈",
            "酒店、餐饮、会议、停车配套完善",
            "利于展贸接待与日常商务活动",
        ]),
        ("创新网络", [
            "大学路 · 创智天地 · 国定路科创带",
            "紧邻内容产业与互联网集聚区",
            "便于高校成果转化与企校联动",
        ]),
    ]
    for i, (title, items) in enumerate(cards):
        x = 16 * mm + i * 60 * mm
        c.setFillColor(white)
        c.setStrokeColor(LINE)
        c.roundRect(x, y - 55 * mm, 56 * mm, 55 * mm, 3, fill=1, stroke=1)
        c.setFillColor(GOLD)
        c.rect(x, y - 1 * mm, 56 * mm, 1.2 * mm, fill=1, stroke=0)
        c.setFillColor(PURPLE)
        c.setFont(FONT, 11)
        c.drawString(x + 3 * mm, y - 8 * mm, title)
        c.setFillColor(MUTED)
        c.setFont(FONT, 8)
        yy = y - 16 * mm
        for it in items:
            # wrap
            line = ""
            for ch in it:
                if c.stringWidth(line + ch, FONT, 8) > 48 * mm:
                    c.drawString(x + 3 * mm, yy, line)
                    yy -= 4.2 * mm
                    line = ch
                else:
                    line += ch
            if line:
                c.drawString(x + 3 * mm, yy, line)
                yy -= 5.5 * mm

    y -= 68 * mm
    c.setFillColor(PURPLE)
    c.setFont(FONT, 11)
    c.drawString(16 * mm, y, "五区联动优势")
    c.setFillColor(MUTED)
    c.setFont(FONT, 9)
    c.drawString(16 * mm, y - 8 * mm, "杨浦推动大学校区 + 科技园区 + 公共社区融合；五角场定位为社区、校区、园区、营区、商区联动创新高地。")
    c.drawString(16 * mm, y - 14 * mm, "创智汇承接「空间改造 + 产业升级 + 商业激活」完整路径，适合作为 AI+数字内容示范载体。")

    # tags
    tags = ["环同济", "五角场副中心", "YOUNG立方", "内容产业圈", "高校人才池"]
    for i, t in enumerate(tags):
        x = 16 * mm + i * 35 * mm
        c.setFillColor(PURPLE3)
        c.roundRect(x, y - 32 * mm, 32 * mm, 9 * mm, 2, fill=1, stroke=0)
        c.setFillColor(GOLD2)
        c.setFont(FONT, 8)
        c.drawCentredString(x + 16 * mm, y - 29 * mm, t)

    footer(c, page, total)


def page_price(c, page, total):
    draw_bg(c)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 8)
    c.drawString(16 * mm, H - 9 * mm, "FLOOR & PRICING")
    c.setFillColor(white)
    c.setFont(FONT, 9)
    c.drawRightString(W - 16 * mm, H - 9 * mm, "楼书 · 楼层与价格")

    y = section_title(c, "03  FLOOR & PRICE", "楼层与价格", H - 28 * mm)

    # two floors
    for i, (fl, area, axis, pos, color) in enumerate([
        ("3 楼", "≈ 2850㎡", "AI / OPC 主轴", "孵化器 + 产业办公", PURPLE),
        ("5 楼", "≈ 3670㎡", "IP 内容主轴", "展厅 + 贸易集群", PURPLE2),
    ]):
        x = 16 * mm + i * 92 * mm
        c.setFillColor(color)
        c.roundRect(x, y - 42 * mm, 88 * mm, 42 * mm, 4, fill=1, stroke=0)
        c.setFillColor(GOLD2)
        c.setFont(FONT, 16)
        c.drawString(x + 5 * mm, y - 12 * mm, fl)
        c.setFillColor(white)
        c.setFont(FONT, 12)
        c.drawString(x + 28 * mm, y - 12 * mm, area)
        c.setFont(FONT, 10)
        c.setFillColor(GOLD2)
        c.drawString(x + 5 * mm, y - 22 * mm, axis)
        c.setFillColor(HexColor("#C6BAEA"))
        c.setFont(FONT, 9)
        c.drawString(x + 5 * mm, y - 30 * mm, pos)
        c.drawString(x + 5 * mm, y - 37 * mm, "可分割单元 · 支持整层/分区租赁")

    y -= 52 * mm
    c.setFillColor(PURPLE)
    c.setFont(FONT, 11)
    c.drawString(16 * mm, y, "参考租赁价格（沟通口径）")
    c.setFillColor(MUTED)
    c.setFont(FONT, 8)
    c.drawString(16 * mm, y - 5.5 * mm, "以下为招商沟通参考，最终以物业条件、装修标准、租期与优惠政策书面确认为准。")

    y -= 12 * mm
    headers = ["产品类型", "建议面积", "参考租金", "物业/备注"]
    col_w = [42 * mm, 38 * mm, 48 * mm, 50 * mm]
    # header
    c.setFillColor(PURPLE)
    c.rect(16 * mm, y - 8 * mm, sum(col_w), 9 * mm, fill=1, stroke=0)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 9)
    x = 16 * mm
    for htxt, w in zip(headers, col_w):
        c.drawString(x + 2 * mm, y - 5.5 * mm, htxt)
        x += w

    rows = [
        ("3F 标准小单元", "80–150㎡", "约 2.3–3.0 元/㎡/天", "精装交付面议"),
        ("3F 成长型单元", "200–350㎡", "约 2.3–2.8 元/㎡/天", "适合研发/设计团队"),
        ("3F OPC / 工位", "按工位", "工位月租面议", "社群配套共享"),
        ("3F 直播/展示位", "共享时段", "按次 / 包月", "可共享运营位"),
        ("5F 集群展位", "270–900㎡", "约 2.0–2.8 元/㎡/天", "展贸属性可谈扣点"),
        ("5F 培训/沙龙", "120–650㎡", "租金或分成二选一", "活动经营性收入"),
        ("整层 / 大客户", "整层或≥1000㎡", "一事一议", "装修与免租期可谈"),
    ]
    for i, row in enumerate(rows):
        yy = y - 9 * mm - i * 8.5 * mm
        c.setFillColor(white if i % 2 else HexColor("#EFEAF8"))
        c.rect(16 * mm, yy - 6.5 * mm, sum(col_w), 8.5 * mm, fill=1, stroke=0)
        c.setFillColor(INK)
        c.setFont(FONT, 8.5)
        x = 16 * mm
        for val, w in zip(row, col_w):
            c.drawString(x + 2 * mm, yy - 3.5 * mm, val)
            x += w

    y = y - 9 * mm - len(rows) * 8.5 * mm - 8 * mm
    c.setFillColor(HexColor("#FFF8E7"))
    c.setStrokeColor(GOLD)
    c.roundRect(16 * mm, y - 28 * mm, W - 32 * mm, 28 * mm, 3, fill=1, stroke=1)
    c.setFillColor(PURPLE)
    c.setFont(FONT, 10)
    c.drawString(20 * mm, y - 8 * mm, "价格说明")
    c.setFillColor(MUTED)
    c.setFont(FONT, 8.5)
    c.drawString(20 * mm, y - 15 * mm, "· 参考租金锚定片区产业办公成本与业主运营测算（约 2.3 元/㎡/天量级），优质客户可叠加免租期 / 装修补贴。")
    c.drawString(20 * mm, y - 21 * mm, "· 5F 展贸可采用「租金」或「订单扣点」模式，避免重复取酬；战略客户、协会挂牌单位一事一议。")

    footer(c, page, total)


def page_policy(c, page, total):
    draw_bg(c)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 8)
    c.drawString(16 * mm, H - 9 * mm, "INCENTIVES & POLICY")
    c.setFillColor(white)
    c.setFont(FONT, 9)
    c.drawRightString(W - 16 * mm, H - 9 * mm, "楼书 · 优惠政策")

    y = section_title(c, "04  POLICY PACKAGE", "优惠政策（入驻礼包）", H - 28 * mm)

    c.setFillColor(MUTED)
    c.setFont(FONT, 8.5)
    c.drawString(16 * mm, y, "以下为企业侧可对接的主要政策抓手（以各级政府最新文件及认定结果为准），由科技企业服务中心协助申报。")
    y -= 8 * mm

    policies = [
        ("YOUNG立方 · 内容产业", [
            "入孵场景：房租物业减免阶梯支持（政策适用主体以认定为准）",
            "直播券：单企年度综合补贴最高约 50 万",
            "活动补贴：按投入最高 50%，同年累计最高 200 万",
            "人才：优质创作者购房/租房支持（最高档可达 200 万购房补贴）",
        ]),
        ("长阳秀带 · AI / 大数据", [
            "AI 及大数据企业房租补贴：最高约 100 万/年（最长约 3 年）",
            "技术攻关研发补贴、应用场景资助等可叠加对接",
            "适合 3F AI 应用、智能体、内容工具类企业",
        ]),
        ("企业资质奖励", [
            "高新技术企业首次认定：约 20 万",
            "专精特新（市级）约 10 万；国家小巨人约 30 万",
            "区科技小巨人 / 双创小巨人等梯度支持",
            "科技创新券：企业每年最高约 30 万额度（抵付服务）",
        ]),
        ("载体与服务加持", [
            "拟挂牌杨浦科技企业服务中心，政务申报快通道",
            "成果转化服务平台运营支持（经认定）",
            "协会 / 校友会 / IP 机构挂牌背书，助力获客与信用",
            "算力补贴、云资源折扣可作为招商组合拳",
        ]),
    ]
    for i, (title, items) in enumerate(policies):
        col = i % 2
        row = i // 2
        x = 16 * mm + col * 92 * mm
        yy = y - row * 58 * mm
        c.setFillColor(white)
        c.setStrokeColor(LINE)
        c.roundRect(x, yy - 54 * mm, 88 * mm, 54 * mm, 3, fill=1, stroke=1)
        c.setFillColor(PURPLE)
        c.rect(x, yy - 1.5 * mm, 88 * mm, 1.5 * mm, fill=1, stroke=0)
        c.setFillColor(PURPLE)
        c.setFont(FONT, 10)
        c.drawString(x + 3 * mm, yy - 8 * mm, title)
        c.setFillColor(MUTED)
        c.setFont(FONT, 8)
        ty = yy - 15 * mm
        for it in items:
            c.setFillColor(GOLD)
            c.drawString(x + 3 * mm, ty, "▸")
            c.setFillColor(MUTED)
            c.drawString(x + 7 * mm, ty, it[:28] + ("…" if len(it) > 28 else ""))
            # full text with wrap if needed - keep short lines
            if len(it) > 28:
                c.drawString(x + 7 * mm, ty - 3.8 * mm, it[28:56] + ("…" if len(it) > 56 else ""))
                ty -= 8.5 * mm
            else:
                ty -= 5.5 * mm

    y = y - 2 * 58 * mm - 4 * mm
    c.setFillColor(PURPLE3)
    c.roundRect(16 * mm, y - 22 * mm, W - 32 * mm, 22 * mm, 3, fill=1, stroke=0)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 9)
    c.drawString(20 * mm, y - 8 * mm, "入驻即享服务")
    c.setFillColor(white)
    c.setFont(FONT, 8.5)
    c.drawString(20 * mm, y - 15 * mm, "政策诊断 → 申报陪跑 → 高企/专精特新培育 → 创新券对接 → 出海与活动资源　｜　具体补贴以审批拨付为准")

    footer(c, page, total)


def page_owner(c, page, total):
    draw_bg(c)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 8)
    c.drawString(16 * mm, H - 9 * mm, "OWNER & OPERATOR")
    c.setFillColor(white)
    c.setFont(FONT, 9)
    c.drawRightString(W - 16 * mm, H - 9 * mm, "楼书 · 业主介绍")

    y = section_title(c, "05  OWNER", "业主与运营协同", H - 28 * mm)

    c.setFillColor(PURPLE)
    c.roundRect(16 * mm, y - 48 * mm, W - 32 * mm, 48 * mm, 4, fill=1, stroke=0)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 14)
    c.drawString(22 * mm, y - 12 * mm, "中建四局（中国建筑第四工程局）")
    c.setFillColor(white)
    c.setFont(FONT, 9.5)
    texts = [
        "世界 500 强中国建筑旗下骨干企业，具备大型城市更新、产业空间与公建综合体建设运营能力。",
        "本项目位于五角场片区城市更新范畴，业主方以高品质空间交付与长期资产运营为导向。",
        "创智汇定位为产业空间合作项目——空间硬件由业主侧保障，产业招商与企业服务由专业运营方协同。",
    ]
    for i, t in enumerate(texts):
        c.drawString(22 * mm, y - 22 * mm - i * 6.5 * mm, t)
    y -= 58 * mm

    c.setFillColor(PURPLE)
    c.setFont(FONT, 11)
    c.drawString(16 * mm, y, "业主价值 · 给入驻企业的确定性")
    y -= 6 * mm
    vals = [
        ("空间品质", "城市更新高标准改造，机电消防物业体系完整，适合长期办公与展贸"),
        ("信用背书", "央企背景业主，合同履约与资产管理规范，降低选址风险"),
        ("产业协同", "可对接中建体系工程、供应链与外延园区资源（视合作进度）"),
        ("长期运营", "不以短炒为导向，愿与优质产业运营方共建可复制样板"),
    ]
    for i, (t, d) in enumerate(vals):
        yy = y - i * 16 * mm
        c.setFillColor(white)
        c.setStrokeColor(LINE)
        c.roundRect(16 * mm, yy - 13 * mm, W - 32 * mm, 14 * mm, 2, fill=1, stroke=1)
        c.setFillColor(GOLD)
        c.setFont(FONT, 10)
        c.drawString(20 * mm, yy - 6 * mm, t)
        c.setFillColor(MUTED)
        c.setFont(FONT, 8.5)
        c.drawString(48 * mm, yy - 6 * mm, d)

    y = y - 4 * 16 * mm - 6 * mm
    c.setFillColor(HexColor("#FFF8E7"))
    c.setStrokeColor(GOLD)
    c.roundRect(16 * mm, y - 32 * mm, W - 32 * mm, 32 * mm, 3, fill=1, stroke=1)
    c.setFillColor(PURPLE)
    c.setFont(FONT, 10)
    c.drawString(20 * mm, y - 8 * mm, "运营服务协同（非业主本体）")
    c.setFillColor(MUTED)
    c.setFont(FONT, 8.5)
    c.drawString(20 * mm, y - 15 * mm, "产业招商、活动媒体、政策申报、科创服务中心挂牌等，由同普会 / 同浦汇与杨浦区科技企业服务中心等专业方协同落地。")
    c.drawString(20 * mm, y - 21 * mm, "入驻企业可同时获得「央企业主空间」+「协会级政策服务」双重支持，形成选址与成长闭环。")
    c.drawString(20 * mm, y - 27 * mm, "＊业主主体名称、产权与招商业态以最终公示及租赁合同约定为准。")

    footer(c, page, total)


def draw_zone(c, x, y, w, h, title, area, fill, text_color=white):
    c.setFillColor(fill)
    c.setStrokeColor(GOLD)
    c.setLineWidth(0.6)
    c.roundRect(x, y, w, h, 2, fill=1, stroke=1)
    c.setFillColor(text_color)
    c.setFont(FONT, 8)
    c.drawCentredString(x + w / 2, y + h / 2 + 2, title)
    if area:
        c.setFont(FONT, 7)
        c.setFillColor(GOLD2)
        c.drawCentredString(x + w / 2, y + h / 2 - 6, area)


def page_plan_3f(c, page, total):
    draw_bg(c)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 8)
    c.drawString(16 * mm, H - 9 * mm, "FLOOR PLAN · 3F")
    c.setFillColor(white)
    c.setFont(FONT, 9)
    c.drawRightString(W - 16 * mm, H - 9 * mm, "楼书 · 三楼平面")

    y = section_title(c, "06  FLOOR PLAN", "3 楼平面示意 · OPC + AI + IP ≈2850㎡", H - 28 * mm)

    # canvas area
    fx, fy, fw, fh = 16 * mm, y - 95 * mm, 120 * mm, 95 * mm
    c.setFillColor(PURPLE)
    c.roundRect(fx, fy, fw, fh, 3, fill=1, stroke=0)

    pad = 3 * mm
    ix, iy = fx + pad, fy + pad
    iw, ih = fw - 2 * pad, fh - 2 * pad
    lw = iw * 0.56
    gp = 2 * mm
    rx = ix + lw + gp
    rw = iw - lw - gp

    draw_zone(c, ix, iy + ih * 0.62, lw, ih * 0.36, "AI+IP 产业办公", "600㎡", HexColor("#3D2A6E"))
    draw_zone(c, ix, iy + ih * 0.48, lw, ih * 0.12, "接待·前台·休闲·茶水", "配套", HexColor("#5A3A7A"))
    draw_zone(c, ix, iy, lw, ih * 0.46, "AI+IP 产业办公", "1150㎡", HexColor("#2E2058"))
    draw_zone(c, rx, iy + ih * 0.72, rw, ih * 0.26, "直播间", "250㎡", HexColor("#4A3A9A"))
    draw_zone(c, rx, iy + ih * 0.52, rw, ih * 0.18, "AI展示运营", "150㎡", HexColor("#2E6B5A"))
    draw_zone(c, rx, iy, rw, ih * 0.50, "OPC 产业办公", "700㎡", HexColor("#5B3FA8"))

    # side legend
    lx = 140 * mm
    c.setFillColor(white)
    c.setStrokeColor(LINE)
    c.roundRect(lx, fy, W - 16 * mm - lx, fh, 3, fill=1, stroke=1)
    c.setFillColor(PURPLE)
    c.setFont(FONT, 10)
    c.drawString(lx + 3 * mm, fy + fh - 8 * mm, "楼层要点")
    points = [
        "OPC + AI + IP 创新中心整层运营",
        "约 9 个可分割办公单元",
        "含直播间 + AI 展示双运营位",
        "主轴对齐「人工智能+」",
        "适配 AI 初创 / 研发设计 / 孵化",
        "建议 4–6 人设计条件",
        "不做纯联合办公散租",
    ]
    c.setFont(FONT, 8)
    for i, p in enumerate(points):
        c.setFillColor(GOLD)
        c.drawString(lx + 3 * mm, fy + fh - 16 * mm - i * 7 * mm, "●")
        c.setFillColor(MUTED)
        c.drawString(lx + 7 * mm, fy + fh - 16 * mm - i * 7 * mm, p)

    y = fy - 8 * mm
    c.setFillColor(PURPLE)
    c.setFont(FONT, 10)
    c.drawString(16 * mm, y, "功能分区一览")
    rows = [
        ("AI+IP 产业办公", "301–310 等", "约 1750㎡", "主力可分割办公"),
        ("OPC 产业办公", "304 / 305", "约 700㎡", "社群与项目办公"),
        ("直播间 + AI 展示", "314 / 316", "约 400㎡", "内容生产与展示"),
        ("接待配套", "—", "配套", "前台 / 休闲 / 茶水"),
    ]
    y -= 4 * mm
    c.setFillColor(PURPLE)
    c.rect(16 * mm, y - 7 * mm, W - 32 * mm, 8 * mm, fill=1, stroke=0)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 8)
    for i, htxt in enumerate(["分区", "房号参考", "面积", "说明"]):
        c.drawString(18 * mm + i * 45 * mm, y - 4.5 * mm, htxt)
    for i, row in enumerate(rows):
        yy = y - 8 * mm - i * 7.5 * mm
        c.setFillColor(white if i % 2 else HexColor("#EFEAF8"))
        c.rect(16 * mm, yy - 5.5 * mm, W - 32 * mm, 7.5 * mm, fill=1, stroke=0)
        c.setFillColor(INK)
        c.setFont(FONT, 8)
        for j, val in enumerate(row):
            c.drawString(18 * mm + j * 45 * mm, yy - 3 * mm, val)

    c.setFillColor(MUTED)
    c.setFont(FONT, 7.5)
    c.drawString(16 * mm, 16 * mm, "＊平面为方案示意复原，分区/房号/面积以物业实测与设计院图纸为准。")
    footer(c, page, total)


def page_plan_5f(c, page, total):
    draw_bg(c)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 8)
    c.drawString(16 * mm, H - 9 * mm, "FLOOR PLAN · 5F")
    c.setFillColor(white)
    c.setFont(FONT, 9)
    c.drawRightString(W - 16 * mm, H - 9 * mm, "楼书 · 五楼平面")

    y = section_title(c, "06  FLOOR PLAN", "5 楼平面示意 · 综合集群展厅 ≈3670㎡", H - 28 * mm)

    fx, fy, fw, fh = 16 * mm, y - 88 * mm, W - 32 * mm, 88 * mm
    c.setFillColor(PURPLE)
    c.roundRect(fx, fy, fw, fh, 3, fill=1, stroke=0)
    pad = 2.5 * mm
    ix, iy = fx + pad, fy + pad
    iw, ih = fw - 2 * pad, fh - 2 * pad
    g = 1.5 * mm
    # 5 columns approximate
    widths = [0.20, 0.09, 0.22, 0.24, 0.20]
    xs = []
    x = ix
    for i, r in enumerate(widths):
        w = iw * r - g
        xs.append((x, w))
        x += w + g

    draw_zone(c, xs[0][0], iy, xs[0][1], ih, "AI+IP展示中心", "700㎡", HexColor("#6B3A7A"))
    draw_zone(c, xs[1][0], iy, xs[1][1], ih, "沙龙", "", HexColor("#4A3A9A"))
    draw_zone(c, xs[2][0], iy + ih * 0.62, xs[2][1], ih * 0.36, "东莞潮玩", "330㎡", HexColor("#2E6B5A"))
    draw_zone(c, xs[2][0], iy, xs[2][1], ih * 0.58, "汕头玩具集群", "900㎡", HexColor("#5B3FA8"))
    draw_zone(c, xs[3][0], iy + ih * 0.68, xs[3][1], ih * 0.30, "AI智能集群", "270㎡", HexColor("#3D2A6E"))
    draw_zone(c, xs[3][0], iy, xs[3][1], ih * 0.64, "扬州毛绒集群", "700㎡", HexColor("#2E2058"))
    draw_zone(c, xs[4][0], iy + ih * 0.72, xs[4][1], ih * 0.26, "仓储", "120㎡", HexColor("#4A4570"))
    draw_zone(c, xs[4][0], iy, xs[4][1], ih * 0.68, "OPC培训中心", "650㎡", HexColor("#3A2A7A"))

    y = fy - 8 * mm
    c.setFillColor(PURPLE)
    c.setFont(FONT, 10)
    c.drawString(16 * mm, y, "5F 产品与客群")
    items = [
        ("综合 / IP 展示", "700–900㎡", "50+ IP 轮展、博物馆文创、景区文创"),
        ("产业集群展位", "270–900㎡/区", "汕头玩具 · 扬州毛绒 · 东莞潮玩 · AI 智能"),
        ("OPC 培训中心", "约 650㎡", "社群培训、沙龙、小红书招募转化"),
        ("顶层活动", "可临建/档期", "漫展、发布会、峰会借势（门票+赞助）"),
    ]
    y -= 4 * mm
    for i, (a, b, d) in enumerate(items):
        yy = y - i * 11 * mm
        c.setFillColor(white if i % 2 else HexColor("#EFEAF8"))
        c.roundRect(16 * mm, yy - 8 * mm, W - 32 * mm, 10 * mm, 2, fill=1, stroke=0)
        c.setFillColor(PURPLE)
        c.setFont(FONT, 9)
        c.drawString(20 * mm, yy - 4.5 * mm, a)
        c.setFillColor(GOLD)
        c.drawString(58 * mm, yy - 4.5 * mm, b)
        c.setFillColor(MUTED)
        c.setFont(FONT, 8)
        c.drawString(95 * mm, yy - 4.5 * mm, d)

    c.setFillColor(MUTED)
    c.setFont(FONT, 7.5)
    c.drawString(16 * mm, 16 * mm, "＊平面为方案示意复原，分区/面积以物业实测与设计院图纸为准；展贸可谈租金或订单扣点模式。")
    footer(c, page, total)


def page_guide(c, page, total):
    draw_bg(c)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 8)
    c.drawString(16 * mm, H - 9 * mm, "HOW TO ENTER")
    c.setFillColor(white)
    c.setFont(FONT, 9)
    c.drawRightString(W - 16 * mm, H - 9 * mm, "楼书 · 入驻指南")

    y = section_title(c, "07  NEXT", "入驻指南与联系", H - 28 * mm)

    steps = [
        ("01", "需求对接", "确认面积、业态、预算与入住时间"),
        ("02", "看场选址", "参观 3F/5F，匹配单元或展位"),
        ("03", "政策诊断", "服务中心评估可申报补贴与资质路径"),
        ("04", "商务签约", "租金/免租期/装修条款书面确认"),
        ("05", "入驻启航", "注册落位、装修进场、申报陪跑启动"),
    ]
    for i, (n, t, d) in enumerate(steps):
        x = 16 * mm + (i % 5) * 35.5 * mm
        yy = y - (i // 5) * 0
        c.setFillColor(PURPLE if i % 2 == 0 else PURPLE2)
        c.roundRect(x, y - 28 * mm, 33.5 * mm, 28 * mm, 3, fill=1, stroke=0)
        c.setFillColor(GOLD2)
        c.setFont(FONT, 12)
        c.drawCentredString(x + 16.75 * mm, y - 10 * mm, n)
        c.setFillColor(white)
        c.setFont(FONT, 8)
        c.drawCentredString(x + 16.75 * mm, y - 17 * mm, t)
        c.setFillColor(HexColor("#C6BAEA"))
        c.setFont(FONT, 6.5)
        # wrap d
        c.drawCentredString(x + 16.75 * mm, y - 23 * mm, d[:10])
        c.drawCentredString(x + 16.75 * mm, y - 26.5 * mm, d[10:])

    y -= 40 * mm
    c.setFillColor(PURPLE)
    c.setFont(FONT, 11)
    c.drawString(16 * mm, y, "重点适配客群")
    tags = [
        ("AI / AIGC / 智能体", "3F"),
        ("微短剧 / 内容工具", "3F·5F"),
        ("数字 IP / 潮玩玩具", "5F"),
        ("科技型中小企业", "3F"),
        ("高校成果转化", "3F"),
        ("跨境展销贸易", "5F"),
    ]
    for i, (t, fl) in enumerate(tags):
        x = 16 * mm + (i % 3) * 60 * mm
        yy = y - 10 * mm - (i // 3) * 14 * mm
        c.setFillColor(white)
        c.setStrokeColor(LINE)
        c.roundRect(x, yy - 10 * mm, 56 * mm, 12 * mm, 2, fill=1, stroke=1)
        c.setFillColor(INK)
        c.setFont(FONT, 8.5)
        c.drawString(x + 3 * mm, yy - 6 * mm, t)
        c.setFillColor(GOLD)
        c.drawRightString(x + 53 * mm, yy - 6 * mm, fl)

    y -= 48 * mm
    c.setFillColor(PURPLE)
    c.roundRect(16 * mm, y - 55 * mm, W - 32 * mm, 55 * mm, 4, fill=1, stroke=0)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 12)
    c.drawString(22 * mm, y - 12 * mm, "联系我们 · 预约看场")
    c.setFillColor(white)
    c.setFont(FONT, 9.5)
    lines = [
        "项目　　上海创智汇 · 杨浦五角场 · 创智天地片区",
        "业主协同　中建四局体系",
        "产业运营　同普会 / 同浦汇（协同）",
        "政策服务　杨浦区科技企业服务中心",
        "联系人　　____________________　　电话　____________________",
        "邮箱/微信　____________________　　看场预约以双方确认为准",
    ]
    for i, t in enumerate(lines):
        c.drawString(22 * mm, y - 22 * mm - i * 5.2 * mm, t)

    y -= 64 * mm
    c.setFillColor(MUTED)
    c.setFont(FONT, 7.5)
    c.drawString(16 * mm, y, "免责声明：本楼书为沟通展示材料，平面、价格、政策均为示意或公开政策摘要，不构成要约；最终以合同、图纸及政府批文为准。")
    c.drawString(16 * mm, y - 5 * mm, "© 上海创智汇项目沟通材料 · 内部/商务使用")

    footer(c, page, total)


def build():
    ART.mkdir(parents=True, exist_ok=True)
    out_cn = HERE / "创智汇A4楼书宣传册.pdf"
    out_en = HERE / "chuangzhihui-a4-brochure.pdf"

    pages = [
        page_cover,
        page_overview,
        page_location,
        page_price,
        page_policy,
        page_owner,
        page_plan_3f,
        page_plan_5f,
        page_guide,
    ]
    total = len(pages)

    for path in (out_cn, out_en):
        c = canvas.Canvas(str(path), pagesize=A4)
        c.setTitle("上海创智汇 · A4楼书宣传册")
        c.setAuthor("Chuangzhihui Project")
        for i, fn in enumerate(pages, 1):
            fn(c, i, total)
            c.showPage()
        c.save()

    for path in (out_cn, out_en):
        dest = ART / path.name
        dest.write_bytes(path.read_bytes())

    print(f"Wrote {out_cn} ({total} pages)")
    print(f"Copied to {ART}")
    return out_cn


if __name__ == "__main__":
    build()
