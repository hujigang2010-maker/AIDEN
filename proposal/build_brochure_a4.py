#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""创智汇 A4 楼书宣传册（PDF）v2
- 满版铺开、少留白
- 产业招服：同浦汇
- 业主：杨浦科创集团（杨浦国资委）
- 联合协办：上海市杨浦区科技企业联合会
- 科技服务：上海市杨浦区科技企业服务中心
- 国家级孵化器：上海市云计算创新基地
- 学术支持：复旦大学住房政策研究中心
- 联系人：高辰辰 15339617481
- 补充火山引擎园区政策
用法: python3 build_brochure_a4.py
"""

from pathlib import Path

from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

HERE = Path(__file__).resolve().parent
ART = Path("/opt/cursor/artifacts")
FONT_PATH = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
pdfmetrics.registerFont(TTFont("CN", FONT_PATH, subfontIndex=0))
FONT = "CN"

PURPLE = HexColor("#140C32")
PURPLE2 = HexColor("#2E1A5C")
PURPLE3 = HexColor("#241848")
PURPLE4 = HexColor("#1C123E")
GOLD = HexColor("#C9A46A")
GOLD2 = HexColor("#E9C27C")
INK = HexColor("#140C32")
MUTED = HexColor("#4A4460")
LIGHT = HexColor("#F3EFFA")
LINE = HexColor("#D4CBE8")
SOFT = HexColor("#B8AED4")

W, H = A4
M = 12 * mm  # tighter margins for full bleed content


def bg_dark(c):
    c.setFillColor(PURPLE)
    c.rect(0, 0, W, H, fill=1, stroke=0)
    c.setFillColorRGB(0.28, 0.14, 0.48, alpha=0.55)
    c.circle(-20, H + 20, 200, fill=1, stroke=0)
    c.setFillColorRGB(0.42, 0.22, 0.72, alpha=0.22)
    c.circle(W + 50, 80, 190, fill=1, stroke=0)
    c.setFillColor(GOLD)
    c.rect(0, H - 2.2 * mm, W, 2.2 * mm, fill=1, stroke=0)
    c.rect(0, 0, W, 2.2 * mm, fill=1, stroke=0)


def bg_light(c):
    c.setFillColor(LIGHT)
    c.rect(0, 0, W, H, fill=1, stroke=0)
    # full-bleed side wash
    c.setFillColor(PURPLE)
    c.rect(0, 0, 4.5 * mm, H, fill=1, stroke=0)
    c.setFillColor(GOLD)
    c.rect(4.5 * mm, 0, 1.2 * mm, H, fill=1, stroke=0)
    # top band full width
    c.setFillColor(PURPLE)
    c.rect(0, H - 16 * mm, W, 16 * mm, fill=1, stroke=0)
    c.setFillColor(GOLD)
    c.rect(0, H - 17.4 * mm, W, 1.4 * mm, fill=1, stroke=0)
    # bottom band
    c.setFillColor(PURPLE)
    c.rect(0, 0, W, 11 * mm, fill=1, stroke=0)


def foot(c, page, total, dark=False):
    c.setFillColor(GOLD2 if dark else GOLD2)
    c.setFont(FONT, 7.5)
    left = M if dark else 10 * mm
    c.drawString(left, 3.8 * mm if not dark else 4.5 * mm,
                 "上海创智汇 · AI+数字内容无界共创港　｜　同浦汇产业招服")
    c.drawRightString(W - M, 3.8 * mm if not dark else 4.5 * mm, f"{page} / {total}")


def head_bar(c, en, zh):
    c.setFillColor(GOLD2)
    c.setFont(FONT, 8)
    c.drawString(10 * mm, H - 7 * mm, en)
    c.setFillColor(white)
    c.setFont(FONT, 10)
    c.drawRightString(W - M, H - 7 * mm, zh)


def title(c, en, zh, y=None):
    if y is None:
        y = H - 28 * mm
    c.setFillColor(GOLD)
    c.setFont(FONT, 8.5)
    c.drawString(10 * mm, y, en)
    c.setFillColor(INK)
    c.setFont(FONT, 17)
    c.drawString(10 * mm, y - 7.5 * mm, zh)
    c.setStrokeColor(GOLD)
    c.setLineWidth(2.2)
    c.line(10 * mm, y - 10.5 * mm, 48 * mm, y - 10.5 * mm)
    return y - 16 * mm


def panel(c, x, y, w, h, fill=white, stroke=LINE, r=3):
    c.setFillColor(fill)
    c.setStrokeColor(stroke)
    c.setLineWidth(0.7)
    c.roundRect(x, y, w, h, r, fill=1, stroke=1)


def wrap_draw(c, text, x, y, max_w, size=8.5, color=MUTED, leading=None):
    if leading is None:
        leading = size + 2.8
    c.setFont(FONT, size)
    c.setFillColor(color)
    line, n = "", 0
    for ch in text:
        if c.stringWidth(line + ch, FONT, size) > max_w:
            c.drawString(x, y - n * leading, line)
            n += 1
            line = ch
        else:
            line += ch
    if line:
        c.drawString(x, y - n * leading, line)
        n += 1
    return n * leading


# ================= pages =================

def page_cover(c, page, total):
    bg_dark(c)

    # eyebrow
    c.setFillColor(GOLD)
    c.setFont(FONT, 8)
    c.drawString(M, H - 22 * mm, "PROPERTY BROCHURE")
    c.setFillColor(SOFT)
    c.setFont(FONT, 8)
    c.drawString(M + 40 * mm, H - 22 * mm, "·  A4 产业空间招商楼书")

    # HERO — 大气字号层级
    c.setFillColor(white)
    c.setFont(FONT, 48)
    c.drawString(M, H - 48 * mm, "上海创智汇")
    c.setFillColor(GOLD2)
    c.setFont(FONT, 16)
    c.drawString(M, H - 60 * mm, "AI + 数字内容无界共创港")
    c.setStrokeColor(GOLD)
    c.setLineWidth(3)
    c.line(M, H - 65 * mm, M + 62 * mm, H - 65 * mm)

    c.setFillColor(white)
    c.setFont(FONT, 13)
    c.drawString(M, H - 76 * mm, "五角场 · 创智天地")
    c.setFillColor(SOFT)
    c.setFont(FONT, 9)
    c.drawString(M + 48 * mm, H - 76 * mm, "环同济经济圈核心　｜　杨浦城市更新产业空间")

    # KPI 条铺满中部，消灭留白
    kpis = [
        ("6600㎡", "合作面积"),
        ("2850㎡", "3F 孵化办公"),
        ("3670㎡", "5F 展厅贸易"),
        ("国资业主", "杨浦科创集团"),
        ("火山引擎", "园区算力政策"),
    ]
    kw = (W - 2 * M - 4 * 2 * mm) / 5
    for i, (n, l) in enumerate(kpis):
        x = M + i * (kw + 2 * mm)
        c.setFillColor(PURPLE3)
        c.roundRect(x, H - 108 * mm, kw, 24 * mm, 2, fill=1, stroke=0)
        c.setStrokeColor(GOLD)
        c.setLineWidth(0.9)
        c.roundRect(x, H - 108 * mm, kw, 24 * mm, 2, fill=0, stroke=1)
        c.setFillColor(GOLD2)
        c.setFont(FONT, 12)
        c.drawCentredString(x + kw / 2, H - 94 * mm, n)
        c.setFillColor(SOFT)
        c.setFont(FONT, 7)
        c.drawCentredString(x + kw / 2, H - 102 * mm, l)

    # 一句话卖点条
    c.setFillColor(HexColor("#2A1850"))
    c.roundRect(M, H - 124 * mm, W - 2 * M, 12 * mm, 2, fill=1, stroke=0)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 8.5)
    c.drawCentredString(W / 2, H - 119.5 * mm,
                        "国资业主 × 联合会协办 × 服务中心科技服务 × 国家级孵化器 × 复旦学术支持 × 同浦汇产业招服")

    # partners — 拉高铺满至联系条，消除下半留白
    partners = [
        ("业主", "杨浦科创集团", "杨浦国资委体系"),
        ("联合协办", "杨浦区科技企业联合会", "协会背书·政企通道"),
        ("科技服务", "杨浦区科技企业服务中心", "政策申报·企业服务"),
        ("国家级孵化器", "上海市云计算创新基地", "算力与云孵化载体"),
        ("学术支持", "复旦大学住房政策研究中心", "政策研究与智力支持"),
        ("产业招服", "同浦汇", "招商运营·产业服务"),
    ]
    y_top = H - 130 * mm
    contact_top = 31 * mm
    gap = 2.2 * mm
    cw = (W - 2 * M - 4 * mm) / 3
    ch = (y_top - contact_top - gap) / 2
    for i, (role, name, sub) in enumerate(partners):
        col, row = i % 3, i // 3
        x = M + col * (cw + 2 * mm)
        y = y_top - row * (ch + gap)
        c.setFillColor(PURPLE3)
        c.roundRect(x, y - ch, cw, ch, 2.5, fill=1, stroke=0)
        c.setStrokeColor(GOLD)
        c.setLineWidth(0.85)
        c.roundRect(x, y - ch, cw, ch, 2.5, fill=0, stroke=1)
        c.setFillColor(GOLD)
        c.setFont(FONT, 8)
        c.drawString(x + 3.5 * mm, y - 8 * mm, role)
        c.setFillColor(white)
        c.setFont(FONT, 11)
        if c.stringWidth(name, FONT, 11) > cw - 7 * mm:
            mid = (len(name) + 1) // 2
            c.drawString(x + 3.5 * mm, y - 18 * mm, name[:mid])
            c.drawString(x + 3.5 * mm, y - 25 * mm, name[mid:])
        else:
            c.drawString(x + 3.5 * mm, y - 20 * mm, name)
        c.setFillColor(SOFT)
        c.setFont(FONT, 7.5)
        c.drawString(x + 3.5 * mm, y - ch + 6 * mm, sub)

    # contact bar
    c.setFillColor(HexColor("#3A2468"))
    c.rect(0, 13 * mm, W, 18 * mm, fill=1, stroke=0)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 10)
    c.drawString(M, 24 * mm, "预约看场")
    c.setFillColor(white)
    c.setFont(FONT, 18)
    c.drawString(M + 26 * mm, 22.5 * mm, "高辰辰　15339617481")
    c.setFillColor(SOFT)
    c.setFont(FONT, 8)
    c.drawRightString(W - M, 23 * mm, "同浦汇 · 产业招服")

    foot(c, page, total, dark=True)


def page_overview(c, page, total):
    bg_light(c)
    head_bar(c, "SHANGHAI CHUANGZHIHUI  ·  OVERVIEW", "楼书 · 项目总览")
    y = title(c, "01  PROJECT OVERVIEW", "项目总览")

    # hero statement full width
    panel(c, 10 * mm, y - 34 * mm, W - 16 * mm, 34 * mm, fill=PURPLE, stroke=PURPLE)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 9)
    c.drawString(14 * mm, y - 8 * mm, "定位")
    c.setFillColor(white)
    c.setFont(FONT, 11)
    wrap_draw(c,
              "打造「上海创智汇 · AI+数字内容无界共创港」——环同济超级链接器：IP 为内容、AI 为工具、空间为载体，服务科创企业入驻、展示交易、政策兑现与出海成长。",
              14 * mm, y - 16 * mm, W - 28 * mm, size=10, color=white, leading=13)
    y -= 40 * mm

    rows = [
        ("项目名称", "上海创智汇（一期）OPC + AI + IP 创新中心"),
        ("项目位置", "上海市杨浦区五角场 · 创智天地片区"),
        ("建筑面积", "约 6600㎡（3 楼 ≈2850㎡ + 5 楼 ≈3670㎡）"),
        ("产业主轴", "人工智能 · 数字内容 · OPC · IP 潮玩文创"),
        ("业主单位", "杨浦科创集团（杨浦区国资委体系）"),
        ("联合协办", "上海市杨浦区科技企业联合会"),
        ("科技服务", "上海市杨浦区科技企业服务中心"),
        ("国家级孵化器", "上海市云计算创新基地"),
        ("学术支持", "复旦大学住房政策研究中心"),
        ("产业招服", "同浦汇"),
        ("政策卡位", "杨浦 YOUNG立方内容集聚区核心圈 + 火山引擎园区政策"),
        ("适合客群", "AI/AIGC、微短剧、数字IP、潮玩玩具、科技型中小、成果转化"),
    ]
    rh = 8.2 * mm
    for i, (k, v) in enumerate(rows):
        yy = y - i * rh
        c.setFillColor(PURPLE3 if i % 2 == 0 else white)
        c.rect(10 * mm, yy - rh + 1.5 * mm, W - 16 * mm, rh, fill=1, stroke=0)
        c.setFillColor(GOLD2 if i % 2 == 0 else GOLD)
        c.setFont(FONT, 8.5)
        c.drawString(13 * mm, yy - 4.8 * mm, k)
        c.setFillColor(white if i % 2 == 0 else INK)
        c.setFont(FONT, 8.5)
        c.drawString(48 * mm, yy - 4.8 * mm, v)

    y = y - len(rows) * rh - 4 * mm
    # toc fills rest
    panel(c, 10 * mm, 14 * mm, W - 16 * mm, y - 14 * mm, fill=white)
    c.setFillColor(PURPLE)
    c.setFont(FONT, 10)
    c.drawString(14 * mm, y - 7 * mm, "本册目录")
    toc = [
        "02 区位交通　五角场副中心 · 轨交商圈 · 高校环绕",
        "03 楼层价格　3F/5F 产品 · 参考租金",
        "04 优惠政策　YOUNG立方 / 长阳秀带 / 火山引擎 / 高企专精特新",
        "05 业主与协同　杨浦科创集团 · 联合会 · 服务中心 · 云创基地 · 复旦",
        "06–07 平面图　3F 孵化办公 · 5F 展厅贸易",
        "08 入驻指南　流程 · 客群 · 联系人高辰辰",
    ]
    for i, t in enumerate(toc):
        c.setFillColor(MUTED)
        c.setFont(FONT, 9)
        c.drawString(14 * mm, y - 16 * mm - i * 6.2 * mm, "▸  " + t)
    foot(c, page, total)


def page_location(c, page, total):
    bg_light(c)
    head_bar(c, "LOCATION & TRANSIT", "楼书 · 区位交通")
    y = title(c, "02  LOCATION", "区位交通")

    # big map panel — tall to fill
    panel(c, 10 * mm, y - 48 * mm, W - 16 * mm, 48 * mm, fill=PURPLE, stroke=PURPLE)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 14)
    c.drawString(15 * mm, y - 12 * mm, "杨浦 · 五角场城市副中心")
    c.setFillColor(white)
    c.setFont(FONT, 11)
    c.drawString(15 * mm, y - 22 * mm, "创智天地片区 · 环同济经济圈核心")
    c.setFillColor(SOFT)
    c.setFont(FONT, 9)
    for i, t in enumerate([
        "东临大学路创新街区 · 南接五角场商圈 · 北依江湾知识社区",
        "正处 YOUNG立方「一楼（V聚场）· 一街（大学路）· 一园（B站新世代）· 一区」核心腹地",
        "复旦、同济、上海财大、上理工高校群环绕；人才与创业项目密度领先",
        "邻近轨交 TOD，商圈酒店餐饮会议配套成熟，利于展贸接待与日常通勤",
    ]):
        c.drawString(15 * mm, y - 32 * mm - i * 5.2 * mm, "●  " + t)
    y -= 54 * mm

    cards = [
        ("轨交 TOD", ["邻近地铁 10 号线方向", "多线路换乘达全市商务区", "客户到访 / 人才通勤便利", "活动集散效率高"]),
        ("商圈配套", ["五角场万达 / 百联成熟商圈", "酒店餐饮会议停车完善", "利于展贸与商务接待", "生活配套一站可达"]),
        ("创新网络", ["大学路 · 创智天地科创带", "内容产业与互联网集聚", "企校联动成果转化便利", "YOUNG立方政策核心圈"]),
    ]
    ch = 52 * mm
    cw = (W - 16 * mm - 4 * mm) / 3
    for i, (t, items) in enumerate(cards):
        x = 10 * mm + i * (cw + 2 * mm)
        panel(c, x, y - ch, cw, ch)
        c.setFillColor(GOLD)
        c.rect(x, y - 1.2 * mm, cw, 1.2 * mm, fill=1, stroke=0)
        c.setFillColor(PURPLE)
        c.setFont(FONT, 11)
        c.drawString(x + 3 * mm, y - 9 * mm, t)
        c.setFont(FONT, 8.5)
        for j, it in enumerate(items):
            c.setFillColor(GOLD)
            c.drawString(x + 3 * mm, y - 18 * mm - j * 7.5 * mm, "▸")
            c.setFillColor(MUTED)
            c.drawString(x + 7 * mm, y - 18 * mm - j * 7.5 * mm, it)
    y -= ch + 4 * mm

    # fill remaining with 五区联动 + tags
    panel(c, 10 * mm, 14 * mm, W - 16 * mm, y - 14 * mm, fill=PURPLE3, stroke=PURPLE3)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 11)
    c.drawString(14 * mm, y - 9 * mm, "五区联动 · 选址价值")
    c.setFillColor(white)
    c.setFont(FONT, 9)
    wrap_draw(c,
              "杨浦推动大学校区 + 科技园区 + 公共社区融合；五角场定位为社区、校区、园区、营区、商区联动创新高地。创智汇承接「空间改造 + 产业升级 + 商业激活」路径，适合作为 AI+数字内容示范载体与国家级孵化器协同落位空间。",
              14 * mm, y - 18 * mm, W - 28 * mm, size=9, color=white, leading=12)
    tags = ["环同济", "五角场副中心", "YOUNG立方", "云计算创新基地", "高校人才池", "国资业主"]
    for i, t in enumerate(tags):
        x = 14 * mm + i * 29 * mm
        c.setFillColor(PURPLE2)
        c.roundRect(x, 20 * mm, 27 * mm, 8 * mm, 2, fill=1, stroke=0)
        c.setFillColor(GOLD2)
        c.setFont(FONT, 7)
        c.drawCentredString(x + 13.5 * mm, 22.5 * mm, t)
    foot(c, page, total)


def page_price(c, page, total):
    bg_light(c)
    head_bar(c, "FLOOR & PRICING", "楼书 · 楼层与价格")
    y = title(c, "03  FLOOR & PRICE", "楼层与价格")

    for i, (fl, area, axis, pos) in enumerate([
        ("3 楼", "≈ 2850㎡", "AI / OPC 主轴", "孵化器 + 产业办公 · 可分割单元"),
        ("5 楼", "≈ 3670㎡", "IP 内容主轴", "展厅 + 贸易集群 · 展贸/扣点可选"),
    ]):
        x = 10 * mm + i * 96 * mm
        c.setFillColor(PURPLE if i == 0 else PURPLE2)
        c.roundRect(x, y - 36 * mm, 92 * mm, 36 * mm, 3, fill=1, stroke=0)
        c.setFillColor(GOLD2)
        c.setFont(FONT, 18)
        c.drawString(x + 5 * mm, y - 12 * mm, fl)
        c.setFillColor(white)
        c.setFont(FONT, 13)
        c.drawString(x + 28 * mm, y - 12 * mm, area)
        c.setFillColor(GOLD2)
        c.setFont(FONT, 10)
        c.drawString(x + 5 * mm, y - 22 * mm, axis)
        c.setFillColor(SOFT)
        c.setFont(FONT, 8.5)
        c.drawString(x + 5 * mm, y - 30 * mm, pos)
    y -= 42 * mm

    c.setFillColor(PURPLE)
    c.setFont(FONT, 10)
    c.drawString(10 * mm, y, "参考租赁价格（沟通口径）")
    c.setFillColor(MUTED)
    c.setFont(FONT, 7.5)
    c.drawString(10 * mm, y - 5 * mm, "最终以物业条件、装修标准、租期与优惠政策书面确认为准")
    y -= 8 * mm

    headers = ["产品类型", "建议面积", "参考租金", "备注"]
    col_w = [48 * mm, 40 * mm, 50 * mm, 50 * mm]
    c.setFillColor(PURPLE)
    c.rect(10 * mm, y - 8 * mm, sum(col_w), 9 * mm, fill=1, stroke=0)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 8.5)
    x = 10 * mm
    for htxt, w in zip(headers, col_w):
        c.drawString(x + 2 * mm, y - 5.2 * mm, htxt)
        x += w

    rows = [
        ("3F 标准小单元", "80–150㎡", "约 2.3–3.0 元/㎡/天", "精装交付面议"),
        ("3F 成长型单元", "200–350㎡", "约 2.3–2.8 元/㎡/天", "研发/设计团队"),
        ("3F OPC / 工位", "按工位", "工位月租面议", "社群共享配套"),
        ("3F 直播/展示位", "共享时段", "按次 / 包月", "可共享运营位"),
        ("5F 集群展位", "270–900㎡", "约 2.0–2.8 元/㎡/天", "可谈订单扣点"),
        ("5F 培训/沙龙", "120–650㎡", "租金或分成二选一", "经营性收入"),
        ("整层 / 大客户", "整层或≥1000㎡", "一事一议", "免租期/装修可谈"),
    ]
    for i, row in enumerate(rows):
        yy = y - 9 * mm - i * 8.8 * mm
        c.setFillColor(white if i % 2 else HexColor("#EBE4F6"))
        c.rect(10 * mm, yy - 6.8 * mm, sum(col_w), 8.8 * mm, fill=1, stroke=0)
        c.setFillColor(INK)
        c.setFont(FONT, 8.5)
        x = 10 * mm
        for val, w in zip(row, col_w):
            c.drawString(x + 2 * mm, yy - 3.5 * mm, val)
            x += w

    y = y - 9 * mm - len(rows) * 8.8 * mm - 3 * mm
    panel(c, 10 * mm, 14 * mm, W - 16 * mm, y - 14 * mm, fill=HexColor("#FFF6E4"), stroke=GOLD)
    c.setFillColor(PURPLE)
    c.setFont(FONT, 10)
    c.drawString(14 * mm, y - 8 * mm, "价格与模式说明")
    c.setFillColor(MUTED)
    c.setFont(FONT, 8.5)
    notes = [
        "参考租金锚定片区产业办公成本（约 2.3 元/㎡/天量级），优质客户可叠加免租期 / 装修支持。",
        "5F 展贸可采用「租金」或「订单扣点」模式；战略客户、协会挂牌单位一事一议。",
        "产业招服由同浦汇对接；政策礼包由科技企业服务中心协同诊断（含火山引擎算力包）。",
    ]
    for i, t in enumerate(notes):
        wrap_draw(c, "● " + t, 14 * mm, y - 16 * mm - i * 10 * mm, W - 28 * mm, size=8.5, color=MUTED, leading=11)
    foot(c, page, total)


def page_policy(c, page, total):
    bg_light(c)
    head_bar(c, "INCENTIVES & VOLCANO ENGINE", "楼书 · 优惠政策")
    y = title(c, "04  POLICY PACKAGE", "优惠政策（含火山引擎）")

    c.setFillColor(MUTED)
    c.setFont(FONT, 8)
    c.drawString(10 * mm, y + 2 * mm, "企业侧政策抓手由上海市杨浦区科技企业服务中心协助申报；以最新文件及认定结果为准。")
    y -= 2 * mm

    # volcano featured full width
    panel(c, 10 * mm, y - 48 * mm, W - 16 * mm, 48 * mm, fill=PURPLE, stroke=PURPLE)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 12)
    c.drawString(14 * mm, y - 8 * mm, "火山引擎园区独立政策（优于火山工坊）")
    c.setFillColor(white)
    c.setFont(FONT, 9)
    c.drawString(14 * mm, y - 16 * mm, "① 无门槛半年免费：入园企业按预估半年费用一次性发放代金券，无消费门槛。")
    c.drawString(14 * mm, y - 22 * mm, "② 大客户额外折扣（代金券之外，累计消费越高折扣越优）：")
    # mini discount table
    discs = [("0–10万", "5–7折"), ("10–30万", "4.5–5折"), ("30–50万", "4–4.5折"),
             ("50–100万", "3.5–4折"), ("100–300万", "3–3.5折"), ("300万+", "2.5–3折")]
    for i, (a, b) in enumerate(discs):
        x = 14 * mm + (i % 6) * 30 * mm
        c.setFillColor(PURPLE3)
        c.roundRect(x, y - 36 * mm, 28 * mm, 11 * mm, 1.5, fill=1, stroke=0)
        c.setFillColor(GOLD2)
        c.setFont(FONT, 7)
        c.drawCentredString(x + 14 * mm, y - 29 * mm, a)
        c.setFillColor(white)
        c.setFont(FONT, 8)
        c.drawCentredString(x + 14 * mm, y - 34 * mm, b)
    c.setFillColor(SOFT)
    c.setFont(FONT, 8)
    c.drawString(14 * mm, y - 44 * mm, "叠加杨浦算力补贴（凭票最高 50%）→ 打造「创智汇专有算力政策包」")
    y -= 54 * mm

    # 4 policy cards fill rest
    policies = [
        ("YOUNG立方 · 内容", [
            "直播券单企最高约 50 万/年",
            "活动补贴：投入×50%，封顶 200 万",
            "人才购房/租房支持可叠加",
            "对齐 5F IP / 微短剧业态",
        ]),
        ("长阳秀带 · AI", [
            "AI/大数据房租补贴≤100万/年",
            "最长约 3 年（按文件执行）",
            "研发攻关 / 场景资助可对接",
            "对齐 3F AI 应用与工具企业",
        ]),
        ("企业资质奖励", [
            "高企首认约 20 万",
            "专精特新市级约 10 万",
            "国家小巨人约 30 万",
            "创新券企业≤30 万/年额度",
        ]),
        ("载体与服务加持", [
            "科技企业服务中心申报通道",
            "联合会政企资源协同",
            "云计算创新基地孵化协同",
            "复旦住房政策中心学术支持",
        ]),
    ]
    ch = y - 14 * mm
    cw = (W - 16 * mm - 6 * mm) / 2
    card_h = (ch - 2 * mm) / 2
    for i, (t, items) in enumerate(policies):
        col, row = i % 2, i // 2
        x = 10 * mm + col * (cw + 2 * mm)
        yy = y - row * (card_h + 2 * mm)
        panel(c, x, yy - card_h, cw, card_h)
        c.setFillColor(PURPLE)
        c.setFont(FONT, 10)
        c.drawString(x + 3 * mm, yy - 7 * mm, t)
        c.setFont(FONT, 8.5)
        for j, it in enumerate(items):
            c.setFillColor(GOLD)
            c.drawString(x + 3 * mm, yy - 16 * mm - j * 7 * mm, "▸")
            c.setFillColor(MUTED)
            c.drawString(x + 7 * mm, yy - 16 * mm - j * 7 * mm, it)
    foot(c, page, total)


def page_owner(c, page, total):
    bg_light(c)
    head_bar(c, "OWNER & PARTNERS", "楼书 · 业主与协同单位")
    y = title(c, "05  OWNER & PARTNERS", "业主与协同单位")

    # owner hero
    panel(c, 10 * mm, y - 40 * mm, W - 16 * mm, 40 * mm, fill=PURPLE, stroke=PURPLE)
    c.setFillColor(GOLD)
    c.setFont(FONT, 8)
    c.drawString(14 * mm, y - 8 * mm, "业主单位")
    c.setFillColor(GOLD2)
    c.setFont(FONT, 16)
    c.drawString(14 * mm, y - 18 * mm, "杨浦科创集团")
    c.setFillColor(white)
    c.setFont(FONT, 9.5)
    wrap_draw(c,
              "隶属于杨浦区国资委体系，承担区域科创空间、产业载体与城市更新相关资产运营职能。本项目以国资信用为底色，坚持高品质空间交付与长期产业运营，为企业提供稳定、可信的落位环境。",
              14 * mm, y - 26 * mm, W - 28 * mm, size=9, color=white, leading=12)
    y -= 46 * mm

    units = [
        ("联合协办单位", "上海市杨浦区科技企业联合会",
         "协会品牌背书、会员网络与政企对接；支持挂牌、活动与产业资源导入。"),
        ("科技服务单位", "上海市杨浦区科技企业服务中心",
         "负责政策申报、高企/专精特新陪跑、创新券与企业服务落地。"),
        ("国家级孵化器", "上海市云计算创新基地",
         "国家级孵化器资质；算力、云资源与硬科技企业孵化协同。"),
        ("学术支持单位", "复旦大学住房政策研究中心",
         "提供政策研究、智库支持与学术背书，助力项目高质量发展。"),
        ("产业招服", "同浦汇",
         "产业招商与企业服务运营；看场接待、商务谈判、活动与出海协同。"),
        ("联系人", "高辰辰　15339617481",
         "预约看场、资料索取、政策诊断对接窗口。"),
    ]
    # fill remaining page completely
    avail = y - 14 * mm
    ch = avail / 3 - 1.5 * mm
    cw = (W - 16 * mm - 2 * mm) / 2
    for i, (role, name, desc) in enumerate(units):
        col, row = i % 2, i // 2
        x = 10 * mm + col * (cw + 2 * mm)
        yy = y - row * (ch + 1.5 * mm)
        panel(c, x, yy - ch, cw, ch)
        c.setFillColor(GOLD)
        c.setFont(FONT, 8)
        c.drawString(x + 3 * mm, yy - 6.5 * mm, role)
        c.setFillColor(PURPLE)
        c.setFont(FONT, 10)
        wrap_draw(c, name, x + 3 * mm, yy - 14 * mm, cw - 6 * mm, size=10, color=PURPLE, leading=12)
        wrap_draw(c, desc, x + 3 * mm, yy - 28 * mm, cw - 6 * mm, size=8, color=MUTED, leading=10.5)
    foot(c, page, total)


def draw_zone(c, x, y, w, h, title, area, fill):
    c.setFillColor(fill)
    c.setStrokeColor(GOLD)
    c.setLineWidth(0.55)
    c.roundRect(x, y, w, h, 2, fill=1, stroke=1)
    c.setFillColor(white)
    c.setFont(FONT, 8)
    c.drawCentredString(x + w / 2, y + h / 2 + 2, title)
    if area:
        c.setFillColor(GOLD2)
        c.setFont(FONT, 7)
        c.drawCentredString(x + w / 2, y + h / 2 - 6, area)


def page_plan_3f(c, page, total):
    bg_light(c)
    head_bar(c, "FLOOR PLAN · 3F", "楼书 · 三楼平面")
    y = title(c, "06  FLOOR PLAN", "3 楼平面示意 · ≈2850㎡")

    # large plan fills most of page
    fh = y - 78 * mm
    fx, fy, fw = 10 * mm, 78 * mm, W - 16 * mm
    c.setFillColor(PURPLE)
    c.roundRect(fx, fy, fw, fh, 3, fill=1, stroke=0)

    pad = 3 * mm
    ix, iy = fx + pad, fy + pad
    iw, ih = fw - 2 * pad, fh - 2 * pad
    lw, gp = iw * 0.58, 2 * mm
    rx, rw = ix + lw + gp, iw - lw - gp

    draw_zone(c, ix, iy + ih * 0.62, lw, ih * 0.36, "AI+IP 产业办公", "600㎡", HexColor("#3D2A6E"))
    draw_zone(c, ix, iy + ih * 0.48, lw, ih * 0.12, "接待·前台·休闲·茶水", "配套", HexColor("#5A3A7A"))
    draw_zone(c, ix, iy, lw, ih * 0.46, "AI+IP 产业办公", "1150㎡", HexColor("#2E2058"))
    draw_zone(c, rx, iy + ih * 0.72, rw, ih * 0.26, "直播间", "250㎡", HexColor("#4A3A9A"))
    draw_zone(c, rx, iy + ih * 0.52, rw, ih * 0.18, "AI展示运营", "150㎡", HexColor("#2E6B5A"))
    draw_zone(c, rx, iy, rw, ih * 0.50, "OPC 产业办公", "700㎡", HexColor("#5B3FA8"))

    # bottom table fills rest
    c.setFillColor(PURPLE)
    c.setFont(FONT, 9)
    c.drawString(10 * mm, 70 * mm, "功能分区")
    rows = [
        ("AI+IP 产业办公", "301–310 等", "约 1750㎡", "主力可分割办公"),
        ("OPC 产业办公", "304 / 305", "约 700㎡", "社群与项目办公"),
        ("直播间 + AI 展示", "314 / 316", "约 400㎡", "内容生产与展示"),
        ("接待配套", "—", "配套", "前台 / 休闲 / 茶水"),
        ("楼层要点", "整层运营", "≈2850㎡", "对齐「人工智能+」· 不做纯联合办公散租"),
    ]
    c.setFillColor(PURPLE)
    c.rect(10 * mm, 62 * mm, W - 16 * mm, 7 * mm, fill=1, stroke=0)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 8)
    for i, htxt in enumerate(["分区", "房号参考", "面积", "说明"]):
        c.drawString(13 * mm + i * 46 * mm, 64.2 * mm, htxt)
    for i, row in enumerate(rows):
        yy = 62 * mm - (i + 1) * 8.5 * mm
        c.setFillColor(white if i % 2 else HexColor("#EBE4F6"))
        c.rect(10 * mm, yy, W - 16 * mm, 8.5 * mm, fill=1, stroke=0)
        c.setFillColor(INK)
        c.setFont(FONT, 8)
        for j, val in enumerate(row):
            c.drawString(13 * mm + j * 46 * mm, yy + 2.5 * mm, val)
    c.setFillColor(MUTED)
    c.setFont(FONT, 7)
    c.drawString(10 * mm, 15.5 * mm, "＊平面为方案示意，以物业实测与设计院图纸为准。产业招服：同浦汇")
    foot(c, page, total)


def page_plan_5f(c, page, total):
    bg_light(c)
    head_bar(c, "FLOOR PLAN · 5F", "楼书 · 五楼平面")
    y = title(c, "06  FLOOR PLAN", "5 楼平面示意 · ≈3670㎡")

    fh = y - 72 * mm
    fx, fy, fw = 10 * mm, 72 * mm, W - 16 * mm
    c.setFillColor(PURPLE)
    c.roundRect(fx, fy, fw, fh, 3, fill=1, stroke=0)
    pad, g = 2.5 * mm, 1.5 * mm
    ix, iy = fx + pad, fy + pad
    iw, ih = fw - 2 * pad, fh - 2 * pad
    widths = [0.20, 0.09, 0.22, 0.24, 0.20]
    xs, x = [], ix
    for r in widths:
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

    items = [
        ("综合 / IP 展示", "700–900㎡", "50+ IP 轮展、博物馆/景区文创"),
        ("产业集群展位", "270–900㎡/区", "汕头玩具 · 扬州毛绒 · 东莞潮玩 · AI 智能"),
        ("OPC 培训中心", "约 650㎡", "社群培训、沙龙、内容转化"),
        ("顶层活动", "档期临建", "漫展 / 发布会 / 峰会（门票+赞助）"),
        ("租赁模式", "租金或扣点", "展贸可谈订单扣点，避免重复取酬"),
    ]
    c.setFillColor(PURPLE)
    c.setFont(FONT, 9)
    c.drawString(10 * mm, 64 * mm, "5F 产品与客群")
    for i, (a, b, d) in enumerate(items):
        yy = 56 * mm - i * 8.2 * mm
        c.setFillColor(white if i % 2 else HexColor("#EBE4F6"))
        c.rect(10 * mm, yy, W - 16 * mm, 8.2 * mm, fill=1, stroke=0)
        c.setFillColor(PURPLE)
        c.setFont(FONT, 8.5)
        c.drawString(13 * mm, yy + 2.5 * mm, a)
        c.setFillColor(GOLD)
        c.drawString(52 * mm, yy + 2.5 * mm, b)
        c.setFillColor(MUTED)
        c.drawString(88 * mm, yy + 2.5 * mm, d)
    c.setFillColor(MUTED)
    c.setFont(FONT, 7)
    c.drawString(10 * mm, 15.5 * mm, "＊平面为方案示意，以物业实测与设计院图纸为准。联合协办：杨浦区科技企业联合会")
    foot(c, page, total)


def page_guide(c, page, total):
    bg_light(c)
    head_bar(c, "HOW TO ENTER", "楼书 · 入驻指南")
    y = title(c, "07  NEXT", "入驻指南与联系")

    steps = [
        ("01", "需求对接", "面积业态预算"),
        ("02", "看场选址", "3F/5F 匹配"),
        ("03", "政策诊断", "服务中心评估"),
        ("04", "商务签约", "同浦汇对接"),
        ("05", "入驻启航", "落位与申报"),
    ]
    cw = (W - 16 * mm - 8 * mm) / 5
    for i, (n, t, d) in enumerate(steps):
        x = 10 * mm + i * (cw + 2 * mm)
        c.setFillColor(PURPLE if i % 2 == 0 else PURPLE2)
        c.roundRect(x, y - 26 * mm, cw, 26 * mm, 2.5, fill=1, stroke=0)
        c.setFillColor(GOLD2)
        c.setFont(FONT, 12)
        c.drawCentredString(x + cw / 2, y - 9 * mm, n)
        c.setFillColor(white)
        c.setFont(FONT, 8.5)
        c.drawCentredString(x + cw / 2, y - 16 * mm, t)
        c.setFillColor(SOFT)
        c.setFont(FONT, 7)
        c.drawCentredString(x + cw / 2, y - 22 * mm, d)
    y -= 32 * mm

    c.setFillColor(PURPLE)
    c.setFont(FONT, 10)
    c.drawString(10 * mm, y, "重点适配客群")
    tags = [
        ("AI / AIGC / 智能体", "3F"),
        ("微短剧 / 内容工具", "3F·5F"),
        ("数字 IP / 潮玩玩具", "5F"),
        ("科技型中小企业", "3F"),
        ("高校成果转化", "3F"),
        ("跨境展销贸易", "5F"),
    ]
    for i, (t, fl) in enumerate(tags):
        x = 10 * mm + (i % 3) * 62 * mm
        yy = y - 8 * mm - (i // 3) * 12 * mm
        panel(c, x, yy - 9 * mm, 58 * mm, 10 * mm)
        c.setFillColor(INK)
        c.setFont(FONT, 8.5)
        c.drawString(x + 2.5 * mm, yy - 5.5 * mm, t)
        c.setFillColor(GOLD)
        c.drawRightString(x + 55 * mm, yy - 5.5 * mm, fl)
    y -= 36 * mm

    # contact fills rest
    panel(c, 10 * mm, 14 * mm, W - 16 * mm, y - 14 * mm, fill=PURPLE, stroke=PURPLE)
    c.setFillColor(GOLD2)
    c.setFont(FONT, 13)
    c.drawString(14 * mm, y - 10 * mm, "预约看场 · 同浦汇")
    c.setFillColor(white)
    c.setFont(FONT, 20)
    c.drawString(14 * mm, y - 22 * mm, "高辰辰　15339617481")
    c.setFont(FONT, 9)
    lines = [
        "项目　　上海创智汇 · 杨浦五角场 · 创智天地片区",
        "业主　　杨浦科创集团（杨浦区国资委体系）",
        "联合协办　上海市杨浦区科技企业联合会",
        "科技服务　上海市杨浦区科技企业服务中心",
        "国家级孵化器　上海市云计算创新基地",
        "学术支持　复旦大学住房政策研究中心",
        "产业招服　同浦汇",
    ]
    for i, t in enumerate(lines):
        c.setFillColor(SOFT if i % 2 else white)
        c.drawString(14 * mm, y - 34 * mm - i * 5.5 * mm, t)

    c.setFillColor(GOLD)
    c.setFont(FONT, 7)
    c.drawString(14 * mm, 18 * mm, "本楼书为沟通展示材料，平面/价格/政策为示意或公开摘要，不构成要约；最终以合同、图纸及政府批文为准。")
    foot(c, page, total)


def build():
    ART.mkdir(parents=True, exist_ok=True)
    out_cn = HERE / "创智汇A4楼书宣传册.pdf"
    out_en = HERE / "chuangzhihui-a4-brochure.pdf"
    pages = [
        page_cover, page_overview, page_location, page_price,
        page_policy, page_owner, page_plan_3f, page_plan_5f, page_guide,
    ]
    total = len(pages)
    for path in (out_cn, out_en):
        c = canvas.Canvas(str(path), pagesize=A4)
        c.setTitle("上海创智汇 · A4楼书宣传册")
        c.setAuthor("同浦汇")
        for i, fn in enumerate(pages, 1):
            fn(c, i, total)
            c.showPage()
        c.save()
        (ART / path.name).write_bytes(path.read_bytes())
    print(f"Wrote {out_cn} ({total} pages)")
    return out_cn


if __name__ == "__main__":
    build()
