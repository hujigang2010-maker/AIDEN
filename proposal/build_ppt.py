# -*- coding: utf-8 -*-
"""生成「创智汇 6600㎡ AI+IP 产业创新中心合作方案」高端商务版 PPT。

设计语言：深空蓝渐变 + 香槟金点缀 + 大留白 + 装饰性章节页。
用法: python3 build_ppt.py
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
FONT = "Microsoft YaHei"
FONT_EN = "Arial"

# ---------- 高端商务调色板 ----------
BG_A   = "0A0E1A"   # 深空蓝
BG_B   = "151D38"   # 蓝紫
PANEL  = RGBColor(0x18, 0x21, 0x3E)
PANEL2 = RGBColor(0x12, 0x19, 0x30)
INK    = RGBColor(0xF3, 0xF6, 0xFF)
MUT    = RGBColor(0x9A, 0xA8, 0xCC)
SOFT   = RGBColor(0x6E, 0x7C, 0xA6)
ACC    = RGBColor(0x5B, 0x8C, 0xFF)   # 蓝
ACC2   = RGBColor(0x8A, 0x6C, 0xFF)   # 紫
GOLD   = RGBColor(0xE6, 0xB8, 0x6A)   # 香槟金
GOLD_D = "E6B86A"
GREEN  = RGBColor(0x49, 0xD0, 0xA0)
LINE   = RGBColor(0x2C, 0x37, 0x59)

SW, SH = Inches(13.333), Inches(7.5)
ML = Inches(0.85)          # 左边距
CW = Inches(11.63)         # 内容宽

prs = Presentation()
prs.slide_width = SW
prs.slide_height = SH
BLANK = prs.slide_layouts[6]


# ================= 底层绘制工具 =================
def _no_shadow(sp):
    sp.shadow.inherit = False


def _set_gradient(shape, stops, angle_deg=90):
    """stops: [(pos0-100, 'RRGGBB', alpha0-100 or None)]"""
    spPr = shape._element.spPr
    for tag in ('a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'):
        e = spPr.find(qn(tag))
        if e is not None:
            spPr.remove(e)
    grad = spPr.makeelement(qn('a:gradFill'), {})
    gsLst = grad.makeelement(qn('a:gsLst'), {})
    for pos, color, alpha in stops:
        gs = grad.makeelement(qn('a:gs'), {'pos': str(int(pos * 1000))})
        clr = grad.makeelement(qn('a:srgbClr'), {'val': color})
        if alpha is not None:
            a = grad.makeelement(qn('a:alpha'), {'val': str(int(alpha * 1000))})
            clr.append(a)
        gs.append(clr)
        gsLst.append(gs)
    grad.append(gsLst)
    lin = grad.makeelement(qn('a:lin'), {'ang': str(int(angle_deg * 60000)), 'scaled': '1'})
    grad.append(lin)
    ln = spPr.find(qn('a:ln'))
    if ln is not None:
        ln.addprevious(grad)
    else:
        spPr.append(grad)


def slide(gradient=True):
    s = prs.slides.add_slide(BLANK)
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    _no_shadow(bg); bg.line.fill.background()
    if gradient:
        _set_gradient(bg, [(0, BG_A, None), (55, BG_B, None), (100, BG_A, None)], 120)
    else:
        bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(0x0A, 0x0E, 0x1A)
    s.shapes._spTree.remove(bg._element); s.shapes._spTree.insert(2, bg._element)
    return s


def rect(s, x, y, w, h, fill=None, line=None, lw=1.0, radius=False, grad=None, gang=90):
    t = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    b = s.shapes.add_shape(t, x, y, w, h)
    _no_shadow(b)
    if grad is not None:
        _set_gradient(b, grad, gang)
    elif fill is None:
        b.fill.background()
    else:
        b.fill.solid(); b.fill.fore_color.rgb = fill
    if line is None:
        b.line.fill.background()
    else:
        b.line.color.rgb = line; b.line.width = Pt(lw)
    return b


def oval(s, x, y, w, h, grad=None, gang=90, fill=None):
    b = s.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)
    _no_shadow(b); b.line.fill.background()
    if grad is not None:
        _set_gradient(b, grad, gang)
    elif fill is not None:
        b.fill.solid(); b.fill.fore_color.rgb = fill
    return b


def text(s, x, y, w, h, runs, size=18, color=INK, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, space=1.0, font=FONT, spacing=None):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = Pt(2); tf.margin_right = Pt(2); tf.margin_top = Pt(1); tf.margin_bottom = Pt(1)
    if isinstance(runs, str):
        runs = [(runs, color, bold)]
    p = tf.paragraphs[0]; p.alignment = align; p.line_spacing = space
    for seg in runs:
        t, c, b = seg[0], seg[1], seg[2]
        f = seg[3] if len(seg) > 3 else font
        r = p.add_run(); r.text = t
        r.font.size = Pt(size); r.font.bold = b; r.font.color.rgb = c; r.font.name = f
        if spacing is not None:
            _letter_spacing(r, spacing)
    return tb


def _letter_spacing(run, pts):
    run._r.get_or_add_rPr().set('spc', str(int(pts * 100)))


def bullets(s, x, y, w, h, items, size=14, color=MUT, gap=8, mark=ACC, lh=1.18):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = Pt(2); tf.margin_top = Pt(2)
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap); p.line_spacing = lh
        r0 = p.add_run(); r0.text = "▪  "
        r0.font.size = Pt(size); r0.font.color.rgb = mark; r0.font.name = FONT
        r = p.add_run(); r.text = it
        r.font.size = Pt(size); r.font.color.rgb = color; r.font.name = FONT
    return tb


def header(s, sec, eyebrow, title):
    """统一页眉：左侧细金条 + 小标签 + 大标题；右上角章节序号。"""
    rect(s, ML, Inches(0.62), Inches(0.06), Inches(0.86), fill=GOLD)
    text(s, Emu(ML + Inches(0.22)), Inches(0.6), Inches(9), Inches(0.34),
         eyebrow, size=12.5, color=GOLD, bold=True, spacing=2, font=FONT_EN)
    text(s, Emu(ML + Inches(0.22)), Inches(0.92), Inches(10.5), Inches(0.66),
         title, size=27, color=INK, bold=True)
    # 右上章节序号
    text(s, Inches(11.3), Inches(0.5), Inches(1.55), Inches(0.9),
         [(sec, GOLD, True, FONT_EN)], size=40, align=PP_ALIGN.RIGHT, color=GOLD, font=FONT_EN)
    rect(s, Inches(11.55), Inches(1.32), Inches(1.28), Pt(2), fill=LINE)
    # 底部细线
    rect(s, ML, Inches(6.92), CW, Pt(1), fill=LINE)


def footer(s, idx):
    text(s, ML, Inches(7.0), Inches(9), Inches(0.32),
         [("创智汇 ", SOFT, False), ("CHUANGZHIHUI", SOFT, False, FONT_EN),
          ("  ·  6600㎡ AI+IP 产业创新中心 · 谈判集中汇报", SOFT, False)], size=9.5)
    text(s, Inches(11.4), Inches(7.0), Inches(1.43), Inches(0.32),
         [("%02d" % idx, GOLD, True, FONT_EN), (" / 21", SOFT, False, FONT_EN)],
         size=10.5, align=PP_ALIGN.RIGHT)


def card(s, x, y, w, h, title=None, body=None, items=None, accent=ACC,
         tcolor=INK, num=None, tsize=15.5, bsize=12.5):
    rect(s, x, y, w, h, grad=[(0, "1A2340", None), (100, "121A30", None)],
         gang=120, line=LINE, lw=1, radius=True)
    rect(s, Emu(x + Inches(0.0)), y, Inches(0.07), h, fill=accent)  # 左侧强调条
    ty = Emu(y + Inches(0.2))
    tx = Emu(x + Inches(0.34))
    if num is not None:
        d = Inches(0.5)
        bx = rect(s, tx, Emu(y + Inches(0.22)), d, d, grad=[(0, "5B8CFF", None), (100, "8A6CFF", None)], gang=120, radius=True)
        text(s, tx, Emu(y + Inches(0.22)), d, d, num, size=16, color=RGBColor(0x0A, 0x0E, 0x1A),
             bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font=FONT_EN)
        tx = Emu(x + Inches(1.0))
    if title:
        text(s, tx, ty, Emu(w - (tx - x) - Inches(0.2)), Inches(0.5), title, size=tsize, color=tcolor, bold=True,
             anchor=MSO_ANCHOR.MIDDLE if num else MSO_ANCHOR.TOP)
    if body:
        text(s, Emu(x + Inches(0.34)), Emu(y + Inches(0.66)), Emu(w - Inches(0.6)),
             Emu(h - Inches(0.78)), body, size=bsize, color=MUT, space=1.16)
    if items:
        bullets(s, Emu(x + Inches(0.34)), Emu(y + Inches(0.62)), Emu(w - Inches(0.6)),
                Emu(h - Inches(0.72)), items, size=bsize, mark=accent)


def table(s, x, y, w, rows, col_w, sizes=None, rh=Inches(0.52), head_rh=Inches(0.5),
          first_col_color=GOLD):
    ncol = len(rows[0]); sizes = sizes or [12.5] * ncol
    cy = y
    for ri, row in enumerate(rows):
        cur_h = head_rh if ri == 0 else rh
        if ri == 0:
            rect(s, x, cy, w, cur_h, grad=[(0, "20294A", None), (100, "1A2240", None)], gang=0, radius=False)
        elif ri % 2 == 0:
            rect(s, x, cy, w, cur_h, fill=RGBColor(0x10, 0x17, 0x2C))
        cx = x
        for ci, cell in enumerate(row):
            cwid = Emu(int(w * col_w[ci]))
            tb = s.shapes.add_textbox(cx, cy, cwid, cur_h)
            tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Pt(10); tf.margin_right = Pt(6); tf.margin_top = Pt(2); tf.margin_bottom = Pt(2)
            p = tf.paragraphs[0]; p.line_spacing = 1.04
            r = p.add_run(); r.text = str(cell)
            if ri == 0:
                r.font.size = Pt(12.5); r.font.bold = True; r.font.color.rgb = GOLD
            else:
                r.font.size = Pt(sizes[ci]); r.font.bold = (ci == 0)
                r.font.color.rgb = first_col_color if ci == 0 else (INK if ci == 0 else MUT)
            r.font.name = FONT
            cx = Emu(cx + cwid)
        rect(s, x, Emu(cy + cur_h - Pt(0.75)), w, Pt(0.75), fill=LINE)
        cy = Emu(cy + cur_h)
    # 左侧首列竖描金线
    rect(s, x, y, Pt(2.5), Emu(cy - y), fill=GOLD)
    return cy


# ============================================================
#                         SLIDES
# ============================================================

# ---------- 1 封面 ----------
s = slide()
# 装饰：右侧大光晕圆
oval(s, Inches(8.3), Inches(-2.0), Inches(7.5), Inches(7.5),
     grad=[(0, "5B8CFF", 22), (100, "0A0E1A", 0)], gang=120)
oval(s, Inches(9.6), Inches(3.2), Inches(5.5), Inches(5.5),
     grad=[(0, "8A6CFF", 18), (100, "0A0E1A", 0)], gang=120)
rect(s, 0, 0, Inches(0.16), SH, grad=[(0, "E6B86A", None), (100, "8A6CFF", None)], gang=90)
text(s, ML, Inches(1.15), Inches(11), Inches(0.4),
     [("CHUANGZHIHUI  ·  WUJIAOCHANG URBAN RENEWAL", GOLD, True, FONT_EN)],
     size=13, spacing=3)
rect(s, ML, Inches(1.72), Inches(0.9), Pt(2.5), fill=GOLD)
text(s, ML, Inches(2.0), Inches(11.4), Inches(1.7),
     [("创智汇 ", INK, True), ("6600㎡", GOLD, True), ("\nAI + IP 产业创新中心 · 招商运营合作", INK, True)],
     size=40, bold=True, space=1.12)
text(s, ML, Inches(3.85), Inches(11.4), Inches(0.5),
     [("谈 判 集 中 汇 报", MUT, False)], size=18, spacing=4)
text(s, ML, Inches(4.5), Inches(11.4), Inches(0.9),
     [("五角场片区城市更新 · 创智汇（一期）OPC + AI + IP 创新中心\n", MUT, False),
      ("3 楼 孵化器 + 办公 ≈ ", MUT, False), ("2850㎡", GOLD, True),
      ("      5 楼 展厅 + 贸易 ≈ ", MUT, False), ("3670㎡", GOLD, True)],
     size=15.5, space=1.5)
tags = ["上海 · 杨浦 · 五角场 · 创智天地", "环同济经济圈 · 超级链接器",
        "IP 为内容 · AI 为工具 · 空间为载体", "参照「森马 × 元谷」打法编制"]
tx = ML
for t in tags:
    wd = Inches(0.42 + len(t) * 0.135)
    rect(s, tx, Inches(5.75), wd, Inches(0.52), fill=None, line=LINE, lw=1.25, radius=True)
    text(s, tx, Inches(5.75), wd, Inches(0.52), t, size=11.5, color=INK,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    tx = Emu(tx + wd + Inches(0.2))
rect(s, ML, Inches(6.65), CW, Pt(1), fill=LINE)
text(s, ML, Inches(6.78), Inches(11), Inches(0.4),
     [("v1.0 谈判稿", GOLD, True), ("    本汇报仅用于商务谈判沟通", SOFT, False)], size=11)

# ---------- 2 议程 ----------
s = slide(); header(s, "", "AGENDA", "汇报议程 · 七大板块"); footer(s, 2)
ag = [("00", "关键信息", "材料浓缩为事实底盘"), ("01", "项目与共识", "定位 / 双重身份 / 谈判要点"),
      ("02", "我们带来什么", "六大资源资产化"), ("03", "招商方案", "IP+AI 双轨 / 楼层 / 漏斗"),
      ("04", "节点与排期", "启动 - 导入 - 成型 - 提升"), ("05", "品牌与活动", "挂牌 / 沙龙 / 漫展 / 会客厅"),
      ("06", "条款与收费", "两模式 + 五类收费 + 测算"), ("07", "投决与下一步", "三问决策 + 30/60/90 天")]
cw, ch = Inches(2.78), Inches(2.18)
gx, gy = Inches(0.13), Inches(0.2)
x0 = ML
for i, (n, t, d) in enumerate(ag):
    x = Emu(x0 + (i % 4) * (cw + gx)); y = Emu(Inches(1.95) + (i // 4) * (ch + gy))
    card(s, x, y, cw, ch, accent=(GOLD if i % 2 else ACC))
    text(s, Emu(x + Inches(0.34)), Emu(y + Inches(0.3)), Inches(2), Inches(0.7),
         [(n, GOLD if i % 2 else ACC, True, FONT_EN)], size=38, font=FONT_EN)
    text(s, Emu(x + Inches(0.34)), Emu(y + Inches(1.18)), Emu(cw - Inches(0.5)), Inches(0.4),
         t, size=15.5, color=INK, bold=True)
    text(s, Emu(x + Inches(0.34)), Emu(y + Inches(1.6)), Emu(cw - Inches(0.5)), Inches(0.5),
         d, size=11.5, color=MUT, space=1.1)

# ---------- 3 关键信息 ----------
s = slide(); header(s, "00", "FACT SHEET", "关键信息提取 · 项目基本盘"); footer(s, 3)
y0 = Inches(1.78)
rect(s, ML, y0, Inches(6.05), Inches(3.05), grad=[(0, "1A2340", None), (100, "121A30", None)],
     gang=120, line=LINE, radius=True)
table(s, Emu(ML + Inches(0.18)), Emu(y0 + Inches(0.2)), Inches(5.7), [
    ["项目", "内容"], ["位置", "上海杨浦五角场 · 创智汇"], ["合作面积", "约 6600㎡"],
    ["3 楼 ≈2850㎡", "孵化器 + 办公（AI / OPC 主轴）"], ["5 楼 ≈3670㎡", "展厅 + 贸易（IP 内容主轴）"],
    ["属性", "产业空间 + 科创合作 + 招商运营"]],
    [0.30, 0.70], sizes=[12, 12], rh=Inches(0.46), head_rh=Inches(0.44))
card(s, Inches(7.15), y0, Inches(5.33), Inches(3.05), "区位与资源优势", accent=GOLD, items=[
    "环同济：复旦 / 同济 / 财大 / 上理工高校群",
    "TOD + 五角场商圈成熟，通勤配套俱佳",
    "杨浦三区联动 + 五角场五区联动创新高地",
    "离我方极近（20–30min）→ 轻驻场 · 高频次"])
yy = Inches(5.05)
card(s, ML, yy, Inches(3.72), Inches(1.72), "定位主线", accent=ACC,
     body="超级链接器：IP 为内容、AI 为工具、空间为载体；产业筋骨 · 文化灵魂 · 商业血脉。")
card(s, Inches(4.78), yy, Inches(3.72), Inches(1.72), "三大集群", accent=GREEN,
     body="动漫 IP  ·  科技应用  ·  交互设计")
card(s, Inches(8.76), yy, Inches(3.72), Inches(1.72), "可链接资源", accent=ACC2,
     body="北大上海校友会、同济设计创新院、科企联、IP/玩具/广告协会、混知等 IP、聚成智能、中建四局。", bsize=11.5)


# ---------- 章节页工具 ----------
def divider(idx, num, cn, en, desc):
    s = slide()
    oval(s, Inches(-2.2), Inches(2.4), Inches(7.5), Inches(7.5),
         grad=[(0, "5B8CFF", 16), (100, "0A0E1A", 0)], gang=60)
    oval(s, Inches(9.2), Inches(-2.6), Inches(6.5), Inches(6.5),
         grad=[(0, "8A6CFF", 16), (100, "0A0E1A", 0)], gang=120)
    text(s, Inches(7.1), Inches(0.7), Inches(6.0), Inches(3.2),
         [(num, GOLD, True, FONT_EN)], size=200, align=PP_ALIGN.RIGHT, color=GOLD, font=FONT_EN)
    rect(s, ML, Inches(2.95), Inches(0.9), Pt(3), fill=GOLD)
    text(s, ML, Inches(3.2), Inches(9), Inches(0.4),
         [(en, GOLD, True, FONT_EN)], size=14, spacing=3)
    text(s, ML, Inches(3.6), Inches(9.5), Inches(1.1), cn, size=44, color=INK, bold=True)
    text(s, ML, Inches(4.85), Inches(9), Inches(0.5), desc, size=15, color=MUT, space=1.3)
    footer(s, idx)
    return s


# ---------- 4 章节01 ----------
divider(4, "01", "项目与共识", "PROJECT & CONSENSUS", "不做二房东，做产业空间合作运营")

# ---------- 5 共识 ----------
s = slide(); header(s, "01", "PROJECT & CONSENSUS", "双重身份 + 三件谈判要点"); footer(s, 5)
card(s, ML, Inches(1.8), Inches(5.72), Inches(1.5), "身份一 · 产业孵化核心", accent=ACC,
     body="五角场城市更新的孵化核心与先行示范区——空间改造 + 产业升级 + 商业激活的完整落地路径。")
card(s, Inches(6.76), Inches(1.8), Inches(5.72), Inches(1.5), "身份二 · 超级链接器", accent=ACC2,
     body="环同济经济圈资源链接平台：IP 内容 × AI 工具 × 空间载体，闭环生长。")
text(s, ML, Inches(3.55), Inches(8), Inches(0.4), "本次谈判要点", size=17, color=GOLD, bold=True)
pts = [("1", "合作模式", "轻运营（不对赌）还是\n对赌运营（对赌招商）？"),
       ("2", "商业条款", "月费 + 招商佣金 + 挂牌\n+ 活动 + 媒体 的组合与金额"),
       ("3", "启动时间", "与装修 / 设计交付同步，\n尽快锁定并启动招商")]
for i, (n, t, d) in enumerate(pts):
    x = Emu(ML + i * Inches(3.94))
    card(s, x, Inches(4.05), Inches(3.72), Inches(1.5), num=n, title=t, accent=GOLD)
    text(s, Emu(x + Inches(0.34)), Emu(Inches(4.05) + Inches(0.8)), Inches(3.3), Inches(0.66),
         d, size=11.5, color=MUT, space=1.18)
rect(s, ML, Inches(5.78), CW, Inches(1.0), grad=[(0, "1C2748", None), (100, "16203C", None)],
     gang=0, line=GOLD, lw=1, radius=True)
text(s, Inches(1.1), Inches(5.78), Inches(11.1), Inches(1.0),
     [("核心共识　", GOLD, True),
      ("把 6600㎡ 升级为「产业空间合作运营项目」——我方输出 政策 + 招商 + 产业资源 + 活动 + 媒体 五项能力，按「基础服务费 + 招商佣金 + 增值分成」取酬。", INK, False)],
     size=13.5, anchor=MSO_ANCHOR.MIDDLE, space=1.2)

# ---------- 6 章节02 ----------
divider(6, "02", "我们能带来什么", "WHAT WE BRING", "六大稀缺资源，全部可资产化为招商工具")

# ---------- 7 六大资源 ----------
s = slide(); header(s, "02", "WHAT WE BRING", "资源即招商 · 六大引擎"); footer(s, 7)
res = [("① 政策抓手", "杨浦 AI/大数据补贴存量续享 + 算力补贴最高 50% + 腾讯云 85 折/免 2 月 →「创智汇专有政策包」", GOLD),
       ("② 高校 / 协会背书", "北大上海校友会、同济设计创新院、复旦、科企联、IP/玩具/广告协会 → 挂牌即招商", ACC),
       ("③ IP 内容资源", "混知、观山礼玉等 IP 品牌方 + 汕头玩具/扬州毛绒/东莞潮玩/文创集群 → 5 楼自带内容", ACC2),
       ("④ AI / OPC 社群", "自有 OPC（AI 开放社群）品牌、黑客松、AI 项目库、聚成智能 → 3 楼自带客流", GREEN),
       ("⑤ 活动 / 会客厅", "漫展（门票+赞助）、北欧会客厅（出海撮合）、沙龙、峰会借势 → 活动即招商", ACC),
       ("⑥ 资本 / 产业方", "江西金控基金、景德镇陶瓷版权交易中心、中建四局工程资源、奇瑞/华为外延 → 链主带动", GOLD)]
cw2 = Inches(3.78); ch2 = Inches(1.72)
for i, (t, d, c) in enumerate(res):
    x = Emu(ML + (i % 3) * (cw2 + Inches(0.14))); y = Emu(Inches(1.8) + (i // 3) * (ch2 + Inches(0.2)))
    card(s, x, y, cw2, ch2, title=t, body=d, accent=c, bsize=11.5)
rect(s, ML, Inches(5.95), CW, Inches(0.72), grad=[(0, "1C2748", None), (100, "16203C", None)],
     gang=0, line=LINE, radius=True)
text(s, Inches(1.1), Inches(5.95), Inches(11), Inches(0.72),
     [("普通园区「发广告 · 等上门」；六维齐发 → 客户 ", MUT, False),
      ("「被送进来」而非「被拉进来」。", GOLD, True)], size=13.5, anchor=MSO_ANCHOR.MIDDLE)

# ---------- 8 章节03 ----------
divider(8, "03", "招商方案", "LEASING STRATEGY", "IP + AI 双轨 · 楼层产品化 · 务实目标")

# ---------- 9 楼层定位 ----------
s = slide(); header(s, "03", "LEASING STRATEGY", "3 楼孵化办公（AI 主轴）+ 5 楼展厅贸易（IP 主轴）"); footer(s, 9)
rect(s, ML, Inches(1.78), Inches(5.72), Inches(4.85), grad=[(0, "1A2340", None), (100, "121A30", None)],
     gang=120, line=LINE, radius=True)
rect(s, ML, Inches(1.78), Inches(0.07), Inches(4.85), fill=ACC)
text(s, Inches(1.1), Inches(1.95), Inches(5.4), Inches(0.4),
     [("3 楼", ACC, True), ("　约 2850㎡　孵化 + 办公", INK, True)], size=16)
table(s, Inches(1.05), Inches(2.5), Inches(5.4), [
    ["产品", "面积", "客户"], ["标准小单元", "80–150㎡×8–12", "AI 初创/设计/小微"],
    ["成长型单元", "200–350㎡×3–4", "AI 应用/研发服务"], ["OPC 联合办公", "工位大区", "AI 项目/黑客松/孵化"],
    ["直播 / AI 展示", "共享 1–2 间", "按次 / 时段"]],
    [0.33, 0.34, 0.33], sizes=[11.5, 11, 11], rh=Inches(0.46), head_rh=Inches(0.44), first_col_color=INK)
rect(s, Inches(6.76), Inches(1.78), Inches(5.72), Inches(4.85), grad=[(0, "1A2340", None), (100, "121A30", None)],
     gang=120, line=LINE, radius=True)
rect(s, Inches(6.76), Inches(1.78), Inches(0.07), Inches(4.85), fill=GOLD)
text(s, Inches(7.02), Inches(1.95), Inches(5.4), Inches(0.4),
     [("5 楼", GOLD, True), ("　约 3670㎡　展厅 + 贸易", INK, True)], size=16)
table(s, Inches(6.97), Inches(2.5), Inches(5.4), [
    ["产品", "面积", "内容"], ["综合集群展厅", "700–900㎡", "50+ IP 轮展"],
    ["产业集群展位", "270–700㎡×4–5", "玩具/毛绒/潮玩/文创"], ["IP 展销/跨境", "按摊位", "IP 零售/跨境展销"],
    ["培训沙龙/仓储", "120–650㎡", "OPC 培训/活动"]],
    [0.33, 0.34, 0.33], sizes=[11.5, 11, 11], rh=Inches(0.46), head_rh=Inches(0.44), first_col_color=INK)

# ---------- 10 漏斗 + 目标 ----------
s = slide(); header(s, "03", "LEASING STRATEGY", "四级招商漏斗 + 务实去化目标"); footer(s, 10)
fn = [("L1", "牌照锚定", "校友会/科企联/IP/玩具协会挂牌前置，客户送进来"),
      ("L2", "政策招商", "AI/大数据续享 + 算力补贴 + 腾讯云作「入驻礼包」"),
      ("L3", "社群带流", "OPC 社群、黑客松、小红书投流、漫展、沙龙持续导流"),
      ("L4", "资源转化", "高校成果转化、IP 品牌方、链主配套、外贸玩具厂转化")]
cwf = Inches(2.78)
for i, (n, t, d) in enumerate(fn):
    x = Emu(ML + i * (cwf + Inches(0.16)))
    card(s, x, Inches(1.85), cwf, Inches(2.0), accent=ACC)
    text(s, Emu(x + Inches(0.34)), Inches(2.05), Inches(1.6), Inches(0.5),
         [(n, ACC, True, FONT_EN)], size=26, font=FONT_EN)
    text(s, Emu(x + Inches(0.34)), Inches(2.62), Inches(2.3), Inches(0.4), t, size=15, color=INK, bold=True)
    text(s, Emu(x + Inches(0.34)), Inches(3.08), Inches(2.35), Inches(0.7), d, size=11, color=MUT, space=1.12)
kpi = [("2000–3000㎡", "一年去化达标线（不设过高保底）", GOLD),
       ("3300–4300㎡", "对赌模式建议锁定对赌面积", ACC),
       ("5 类客群", "AI / 潮玩玩具文创 IP / 科技中小 /\n高校转化 / 生产性服务", GREEN)]
for i, (v, l, c) in enumerate(kpi):
    x = Emu(ML + i * Inches(3.94))
    card(s, x, Inches(4.25), Inches(3.72), Inches(1.95), accent=c)
    text(s, Emu(x + Inches(0.36)), Inches(4.55), Inches(3.3), Inches(0.7), [(v, c, True)], size=25, bold=True)
    text(s, Emu(x + Inches(0.36)), Inches(5.4), Inches(3.3), Inches(0.7), l, size=11.5, color=MUT, space=1.12)

# ---------- 11 章节04 ----------
divider(11, "04", "节点与排期", "MILESTONES & SCHEDULE", "启动 · 导入 · 成型 · 提升 四阶段")

# ---------- 12 排期 ----------
s = slide(); header(s, "04", "MILESTONES & SCHEDULE", "从定位包装到稳定运营"); footer(s, 12)
phases = [("启动期", "0–3 月", "定位 + 政策 + 渠道", "招商手册、收费标准、政策汇编、企业库、首批挂牌、1–2 家样板", ACC),
          ("导入期", "3–6 月", "首批入驻 30–50%", "集中招引 AI/IP/潮玩、首批优惠、推介路演、政策诊断", ACC2),
          ("成型期", "6–12 月", "提升出租率与质量", "重点客户补位、专精特新/高企培育、企业服务收费、首场漫展", GREEN),
          ("提升期", "12 月+", "稳定收入与品牌", "区级示范点、载体资质、活动 IP 化、出海撮合、模式复制", GOLD)]
# 时间轴
rect(s, Inches(1.2), Inches(2.05), Pt(3), Inches(4.2), grad=[(0, "5B8CFF", None), (100, "E6B86A", None)], gang=90)
y = Inches(2.0)
for nm, tw, goal, act, c in phases:
    oval(s, Inches(1.06), Emu(y + Inches(0.1)), Inches(0.34), Inches(0.34), fill=c)
    card(s, Inches(1.7), y, Inches(10.78), Inches(0.95), accent=c)
    text(s, Inches(2.0), y, Inches(1.7), Inches(0.95), [(nm, c, True)], size=16, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(3.5), y, Inches(1.4), Inches(0.95), tw, size=13, color=MUT, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(4.7), y, Inches(2.5), Inches(0.95), goal, size=13, color=INK, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(7.0), y, Inches(5.3), Inches(0.95), act, size=11.5, color=MUT, anchor=MSO_ANCHOR.MIDDLE, space=1.1)
    y = Emu(y + Inches(1.07))

# ---------- 13 章节05 ----------
divider(13, "05", "品牌与活动", "BRAND & EVENTS", "挂牌 / 沙龙 / 漫展 / 北欧会客厅")
# 注意：仍占第 13 页，但下面继续是内容；调整 footer 计数

# ---------- 重新规划页码：因品牌与收费内容较多，合并章节为内容页 ----------
# 14 品牌与活动（内容）
s = slide(); header(s, "05", "BRAND & EVENTS", "活动即招商，品牌即势能"); footer(s, 14)
card(s, ML, Inches(1.8), Inches(3.78), Inches(2.25), "挂牌 · 牌照即招商", accent=GOLD,
     body="北大上海校友会（优先）、科企联 / 科技服务中心、IP 协会、中国玩具协会。\n\n总包约 20–50 万")
card(s, Inches(4.77), Inches(1.8), Inches(3.78), Inches(2.25), "活动 · 全年节奏", accent=ACC,
     body="20–24 场沙龙 / 路演 / 培训 / 发布会。\n\n打包约 30 万（+赞助 / 门票）")
card(s, Inches(8.74), Inches(1.8), Inches(3.74), Inches(2.25), "漫展 · 自造流量", accent=ACC2,
     body="千人级，门票 + 赞助；10–11 月首场，对接 BW 广告商与二次元资源。")
rect(s, ML, Inches(4.3), CW, Inches(2.05), grad=[(0, "1A2340", None), (100, "121A30", None)],
     gang=120, line=LINE, radius=True)
rect(s, ML, Inches(4.3), Inches(0.07), Inches(2.05), fill=GREEN)
text(s, Inches(1.1), Inches(4.5), Inches(11), Inches(0.4), "北欧会客厅 · 出海撮合", size=15.5, color=GREEN, bold=True)
text(s, Inches(1.1), Inches(5.05), Inches(11.1), Inches(1.1),
     [("复用品牌势能做海外 IP / 技术入华与国内 IP 出海撮合台。取酬二选一：\n", MUT, False),
      ("收租金则不分出海成交；不收租金、企业自费布展，则在撮合成交中取分成。", INK, True)],
     size=13.5, space=1.3)

# ---------- 15 章节06（用内容页头，避免页数超） ----------
divider(15, "06", "商业条款与收费", "TERMS & FEES", "两种模式 · 五类收费 · 测算（核心）")

# 因为我们已超过 14 页，调整 footer 总数。后面统一改。

ppt_path = os.path.join(HERE, "创智汇6600平合作方案.pptx")

# === 由于内容较多，这里继续补充收费等核心页 ===
# 16 两模式
s = slide(); header(s, "06", "TERMS & FEES", "先定大方向：轻运营 vs 对赌运营"); footer(s, 16)
rect(s, ML, Inches(1.85), Inches(5.72), Inches(3.5), grad=[(0, "1A2340", None), (100, "121A30", None)],
     gang=120, line=LINE, radius=True)
rect(s, ML, Inches(1.85), Inches(0.07), Inches(3.5), fill=SOFT)
text(s, Inches(1.1), Inches(2.05), Inches(5.4), Inches(0.4), "模式 A · 轻运营（不对赌招商）", size=16, color=INK, bold=True)
bullets(s, Inches(1.1), Inches(2.65), Inches(5.3), Inches(2.5),
        ["不驻场、不背去化指标", "收费：活动费 + 挂牌费 + 媒体策划费", "适用：对方只要内容 / 活动 / 品牌"],
        size=13.5, gap=12, mark=SOFT)
rect(s, Inches(6.76), Inches(1.85), Inches(5.72), Inches(3.5), grad=[(0, "23315A", None), (100, "1A2340", None)],
     gang=120, line=GOLD, lw=1.25, radius=True)
rect(s, Inches(6.76), Inches(1.85), Inches(0.07), Inches(3.5), fill=GOLD)
text(s, Inches(7.02), Inches(2.05), Inches(5.4), Inches(0.4),
     [("模式 B · 对赌运营", GOLD, True), ("（推荐）", GREEN, True)], size=16)
bullets(s, Inches(7.02), Inches(2.65), Inches(5.3), Inches(2.5),
        ["轻驻场（1 人/接待位，离得近、可挂牌办公）+ 背对赌去化",
         "收费：基础月费 + 招商佣金 + 活动 + 挂牌 + 媒体",
         "适用：对方要招商结果、愿付运营费"], size=13.5, gap=12, mark=GOLD)
rect(s, ML, Inches(5.65), CW, Inches(0.9), grad=[(0, "1C2748", None), (100, "16203C", None)],
     gang=0, line=LINE, radius=True)
text(s, Inches(1.1), Inches(5.65), Inches(11), Inches(0.9),
     [("谈判原则　", GOLD, True), ("先按高位报，再让对方还价；收月费即对赌去化，不收月费则不背指标。", INK, False)],
     size=13.5, anchor=MSO_ANCHOR.MIDDLE)

# 17 收费清单
s = slide(); header(s, "06", "TERMS & FEES", "五类收费项 · 收费清单"); footer(s, 17)
table(s, ML, Inches(1.82), CW, [
    ["类别", "对象", "方式", "建议金额 / 口径"],
    ["① 基础运营月费", "合作方", "按月", "市场 10–30 万；首报 12 万；对赌版 5–6 万（对赌 3300–4300㎡）"],
    ["② 招商佣金", "合作方/业主", "成交后", "市场 2–3 月年租金；我方抽 ≤1 月，只给首月不重复"],
    ["③ 挂牌费", "合作方", "一次性", "总包 20–50 万（按数量与影响力分档）"],
    ["④ 活动执行费", "合作方+外部", "按场/打包", "20–24 场/年约 30 万 + 赞助 / 门票分成"],
    ["⑤ 媒体流量费", "合作方", "按季/项目", "小红书投流、OPC 内容、直播间运营"],
    ["⑥ 企业增值服务", "入驻企业", "按项/成功", "工商注册、政策申报、知识产权、财税法务、融资、出海撮合"]],
    [0.20, 0.13, 0.13, 0.54], sizes=[12.5, 11.5, 11.5, 11.5], rh=Inches(0.6), head_rh=Inches(0.5),
    first_col_color=GOLD)

# 18 报价测算
s = slide(); header(s, "06", "TERMS & FEES", "打包报价 + 对方视角测算"); footer(s, 18)
rect(s, ML, Inches(1.82), Inches(5.72), Inches(4.5), grad=[(0, "1A2340", None), (100, "121A30", None)],
     gang=120, line=LINE, radius=True)
rect(s, ML, Inches(1.82), Inches(0.07), Inches(4.5), fill=GOLD)
text(s, Inches(1.1), Inches(2.02), Inches(5.4), Inches(0.4), "打包报价示例（谈判锚）", size=15.5, color=INK, bold=True)
bullets(s, Inches(1.1), Inches(2.6), Inches(5.3), Inches(3.6), [
    "固定服务包：月费 5 万×12 = 60 万 + 活动 30 万 + 挂牌 20 万 = 约 110 万/年（可砍至 ~100 万）",
    "浮动：招商佣金按去化另计（我方抽 ≤1 月）；驻场 / 媒体 / 出海分成据实另计",
    "政策申报：普通 0.3–2 万/项；高企 / 专精特新 2–8 万/项；补贴类按到账 5%–15% 成功费"],
    size=12.5, gap=14, mark=GOLD)
rect(s, Inches(6.76), Inches(1.82), Inches(5.72), Inches(4.5), grad=[(0, "1A2340", None), (100, "121A30", None)],
     gang=120, line=LINE, radius=True)
rect(s, Inches(6.76), Inches(1.82), Inches(0.07), Inches(4.5), fill=ACC)
text(s, Inches(7.02), Inches(2.02), Inches(5.4), Inches(0.4), "对方视角测算", size=15.5, color=INK, bold=True)
table(s, Inches(6.97), Inches(2.6), Inches(5.4), [
    ["维度", "口径"], ["对方空间成本", "≈554 万/年（2.3 元/㎡/天×6600×365）"],
    ["我方服务包", "100–120 万/年 + 招商佣金（按去化）"], ["列支来源", "对方利润 / 中建四局运营费，非纯增量"],
    ["对方所得", "去化↑ · 企业质量↑ · 政策承接 · 品牌示范"]],
    [0.30, 0.70], sizes=[12, 11.5], rh=Inches(0.62), head_rh=Inches(0.5), first_col_color=ACC)

# 19 章节07
divider(19, "07", "投决建议与下一步", "DECISION & NEXT STEP", "三问决策 · 30 / 60 / 90 天启动")

# 20 三问
s = slide(); header(s, "07", "DECISION & NEXT STEP", "三问决策 + 启动路线图"); footer(s, 20)
q = [("Q1", "合作模式？", "建议 模式 B 轻量版：离得近 → 轻驻场 + 对赌去化，盈利性最佳", GOLD),
     ("Q2", "月费档位？", "首报 12 万/月（市场 10–30 万），对赌版可落 5–6 万；打包年费锚 100–120 万", ACC),
     ("Q3", "何时启动？", "与装修 / 设计交付同步，尽快签 MOU、先行政策梳理与挂牌筹备", ACC2)]
for i, (n, t, d, qc) in enumerate(q):
    x = Emu(ML + i * Inches(3.94))
    card(s, x, Inches(1.8), Inches(3.72), Inches(1.7), accent=qc)
    text(s, Emu(x + Inches(0.34)), Inches(2.0), Inches(3.2), Inches(0.4),
         [(n + "　", qc, True, FONT_EN), (t, INK, True)], size=15.5)
    text(s, Emu(x + Inches(0.34)), Inches(2.5), Inches(3.2), Inches(0.95), d, size=11.5, color=MUT, space=1.2)
road = [("30", "签 MOU + 招商手册与收费方案 + 政策汇编 + 首批挂牌对接"),
        ("60", "首场沙龙 / 路演 + OPC 社群与投流 + 首单招商 + 漫展筹备"),
        ("90", "正式挂牌 + 样板企业入驻 + 去化向 30–50% 推进 + 漫展锁档")]
for i, (n, d) in enumerate(road):
    x = Emu(ML + i * Inches(3.94))
    card(s, x, Inches(3.7), Inches(3.72), Inches(1.7), accent=ACC)
    text(s, Emu(x + Inches(0.34)), Inches(3.9), Inches(2.6), Inches(0.6),
         [(n, ACC, True, FONT_EN), (" 天", MUT, False)], size=24)
    text(s, Emu(x + Inches(0.34)), Inches(4.55), Inches(3.2), Inches(0.85), d, size=11.5, color=MUT, space=1.18)
rect(s, ML, Inches(5.65), CW, Inches(0.95), grad=[(0, "1C2748", None), (100, "16203C", None)],
     gang=0, line=GOLD, lw=1, radius=True)
text(s, Inches(1.1), Inches(5.65), Inches(11), Inches(0.95),
     [("分步策略　", GOLD, True), ("先签 6600㎡ 单项目专项（单项目单核算），跑通后再承接体系内其他物业与外延项目。", INK, False)],
     size=13.5, anchor=MSO_ANCHOR.MIDDLE)

# 21 封底
s = slide()
oval(s, Inches(7.5), Inches(-2.4), Inches(8), Inches(8),
     grad=[(0, "5B8CFF", 20), (100, "0A0E1A", 0)], gang=120)
oval(s, Inches(-2.5), Inches(3.0), Inches(6.5), Inches(6.5),
     grad=[(0, "8A6CFF", 16), (100, "0A0E1A", 0)], gang=60)
rect(s, 0, 0, Inches(0.16), SH, grad=[(0, "E6B86A", None), (100, "8A6CFF", None)], gang=90)
text(s, ML, Inches(2.3), Inches(11), Inches(0.4), [("THANKS", GOLD, True, FONT_EN)], size=15, spacing=4)
rect(s, ML, Inches(2.85), Inches(0.9), Pt(3), fill=GOLD)
text(s, ML, Inches(3.1), Inches(11.5), Inches(1.8),
     [("让创智汇成为杨浦五角场\n", INK, True), ("AI + IP 产业与城市更新", GOLD, True),
      ("的超级链接器与可复制样板", INK, True)], size=33, bold=True, space=1.2)
text(s, ML, Inches(5.1), Inches(11.5), Inches(0.5),
     "90 天启动  ·  政策借势  ·  闭环招商  ·  内容自造流量", size=16, color=MUT, spacing=1)
rect(s, ML, Inches(6.4), CW, Pt(1), fill=LINE)
text(s, ML, Inches(6.55), Inches(11), Inches(0.4),
     [("谈判集中汇报 v1.0", GOLD, True), ("　期待与各方达成合作", SOFT, False)], size=12)

# 修正所有 footer 的总页数为实际页数
total = len(prs.slides._sldIdLst)
prs.save(ppt_path)
print("PPT saved:", ppt_path, "slides:", total)
