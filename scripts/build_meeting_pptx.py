"""
冠松 01# 楼 · 招商合作会上版 PPT（6 页）

只讲：听懂了、粗思路、收费、怎么停、请定。
不讲 iDrive Hub、不讲链主 TOP5、不讲 22 人团队。

重新生成：python3 scripts/build_meeting_pptx.py
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pathlib import Path

NAVY = RGBColor(0x1B, 0x35, 0x55)
GOLD = RGBColor(0xB7, 0x86, 0x2E)
INK = RGBColor(0x2A, 0x2D, 0x34)
CLOUD = RGBColor(0xF0, 0xF1, 0xF3)
GREY = RGBColor(0x6B, 0x73, 0x80)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
STONE = RGBColor(0x8A, 0x92, 0x9C)

CN_FONT = "WenQuanYi Micro Hei"
EN_FONT = "Georgia"
SLIDES = []


def set_run(run, text, *, size=14, bold=False, color=INK, italic=False):
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = EN_FONT
    rPr = run._r.get_or_add_rPr()
    for tag in ("ea", "cs"):
        existing = rPr.find(qn(f"a:{tag}"))
        if existing is not None:
            rPr.remove(existing)
        rPr.append(rPr.makeelement(qn(f"a:{tag}"), {"typeface": CN_FONT}))


def add_text(slide, x, y, w, h, text, *, size=14, bold=False, color=INK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.margin_left = Pt(6)
    tf.margin_right = Pt(6)
    tf.margin_top = Pt(2)
    tf.margin_bottom = Pt(2)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    box.fill.background()
    box.line.fill.background()
    lines = text.split("\n") if isinstance(text, str) else list(text)
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        set_run(run, ln, size=size, bold=bold, color=color, italic=italic)
    return box


def add_rect(slide, x, y, w, h, *, fill=NAVY):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    shp.text_frame.text = ""
    return shp


def add_round(slide, x, y, w, h, text="", *, fill=NAVY, color=WHITE,
              size=12, bold=True, align=PP_ALIGN.CENTER):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    tf = shp.text_frame
    tf.margin_left = Pt(8)
    tf.margin_right = Pt(8)
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = align
    set_run(p.add_run(), text, size=size, bold=bold, color=color)
    return shp


def add_table(slide, x, y, w, h, header, rows, *, header_size=12, body_size=12,
              col_widths=None):
    ts = slide.shapes.add_table(len(rows) + 1, len(header), x, y, w, h)
    table = ts.table
    if col_widths:
        for i, cw in enumerate(col_widths):
            table.columns[i].width = cw
    for j, htxt in enumerate(header):
        cell = table.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        cell.text = ""
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        set_run(p.add_run(), htxt, size=header_size, bold=True, color=WHITE)
    for i, row in enumerate(rows, start=1):
        bg = WHITE if i % 2 else CLOUD
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            cell.text = ""
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            set_run(p.add_run(), str(val), size=body_size, color=INK)
    return ts


def add_chrome(slide, prs, *, page_no, label, title, subtitle=""):
    sw, sh = prs.slide_width, prs.slide_height
    add_rect(slide, 0, 0, sw, Emu(380000), fill=NAVY)
    add_round(slide, Inches(0.45), Inches(0.18), Inches(2.4), Inches(0.36),
              label, fill=GOLD, color=NAVY, size=11)
    add_text(slide, Inches(3.05), Inches(0.10), Inches(9.6), Inches(0.52),
             title, size=22, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_text(slide, Inches(3.05), Inches(0.52), Inches(9.6), Inches(0.28),
                 subtitle, size=12, color=CLOUD, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(slide, 0, sh - Emu(80000), sw, Emu(80000), fill=GOLD)
    add_text(slide, Inches(0.45), Inches(7.05), Inches(9.2), Inches(0.28),
             "冠松 01# 楼 · 招商合作会上版 · 保密", size=10, color=GREY)
    add_text(slide, Inches(11.4), Inches(7.05), Inches(1.5), Inches(0.28),
             f"{page_no} / 0", size=10, color=GREY, align=PP_ALIGN.RIGHT)


def new_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    SLIDES.append(s)
    return s


def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    sw, sh = prs.slide_width, prs.slide_height

    # 1 封面
    s = new_slide(prs)
    add_rect(s, 0, 0, sw, sh, fill=NAVY)
    add_rect(s, 0, Inches(5.05), sw, Emu(28000), fill=GOLD)
    add_text(s, Inches(0.8), Inches(1.35), Inches(11.5), Inches(0.45),
             "01# 研发楼 · 招商合作", size=16, bold=True, color=GOLD, italic=True)
    add_text(s, Inches(0.8), Inches(1.9), Inches(11.7), Inches(1.4),
             "我们帮你们把这栋楼租出去",
             size=40, bold=True, color=WHITE)
    add_text(s, Inches(0.8), Inches(3.5), Inches(11.7), Inches(0.7),
             "先 90 天 · 包干 8 万 · 佣金跟中介同一把尺",
             size=20, color=CLOUD)
    add_round(s, Inches(0.8), Inches(5.5), Inches(3.4), Inches(0.48),
              "会上版 · 6 页", fill=GOLD, color=NAVY, size=14)
    add_text(s, Inches(4.4), Inches(5.5), Inches(8), Inches(0.48),
             "细版请看「工作版」12 页 · 05b 确认栏当场填",
             size=14, color=CLOUD, anchor=MSO_ANCHOR.MIDDLE)

    # 2 听懂了
    s = new_slide(prs)
    add_chrome(s, prs, page_no=2, label="对齐",
               title="我们听懂了",
               subtitle="得到大脑口径：跟进招商。不是再做一版定位。")
    items = [
        ("1", "楼要有人租", "汽车展厅走不通了。目标是生效合同，不是品牌发布会。"),
        ("2", "不当小白鼠", "抄已经跑通的研发楼，不搞全国首创。"),
        ("3", "不丢掉汽车", "不招整车 4S；汽车配套、后市场可以留。"),
        ("4", "你们拍板我们跑", "合同您盖章。我们建库、带看、谈判、盯收款。"),
    ]
    for i, (n, t, b) in enumerate(items):
        y = Inches(1.25) + Inches(1.3) * i
        add_round(s, Inches(0.5), y, Inches(0.7), Inches(1.05), n,
                  fill=NAVY, color=WHITE, size=22)
        add_text(s, Inches(1.45), y, Inches(3.3), Inches(1.05), t,
                 size=22, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(4.8), y, Inches(7.9), Inches(1.05), b,
                 size=18, color=INK, anchor=MSO_ANCHOR.MIDDLE)

    # 3 90天
    s = new_slide(prs)
    add_chrome(s, prs, page_no=3, label="90 天",
               title="未来 90 天怎么干",
               subtitle="必须：库 150 · 深接触 8 · 带看 5 · 书面意向 3。达不到可不转正。")
    phases = [
        ("D1–30 进场",
         "一页楼书（C6 / 层高）\n看房动线、种子库\n中介先开 2 家\n周五周报"),
        ("D31–60 见客",
         "带看 + 48h 纪要\n15 人闭门 1 场\n只提市北、西岸\n投促办备案不加码"),
        ("D61–90 收口",
         "意向写成书面\n黄区 48h 请示\n漏斗会 90 分钟\n转正 / 再试 / 停"),
    ]
    for i, (t, b) in enumerate(phases):
        x = Inches(0.45) + Inches(4.2) * i
        add_rect(s, x, Inches(1.25), Inches(4.0), Inches(5.35), fill=CLOUD)
        add_rect(s, x, Inches(1.25), Inches(4.0), Inches(0.6), fill=NAVY)
        add_text(s, x, Inches(1.25), Inches(4.0), Inches(0.6), t,
                 size=18, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + Inches(0.2), Inches(2.1), Inches(3.6), Inches(4.2),
                 b, size=18, color=INK, align=PP_ALIGN.CENTER)

    # 4 收费
    s = new_slide(prs)
    add_chrome(s, prs, page_no=4, label="收费",
               title="对齐后：试跑只付 8 万",
               subtitle="旧口径 90 天约 39 万，叠了顾问费和满佣。现在跟中介同一把尺。")

    fees = [
        ("90 天包干", "8 万", "含人、材料、带看\n试跑期不另收月度"),
        ("自拓佣金", "首月 × 100%", "和中介同一尺\n没签成不抽"),
        ("锚定佣金", "首月 × 150%", "≥ 1,500 ㎡ 或整层\n中介单业主只付一套"),
    ]
    for i, (t, n, b) in enumerate(fees):
        x = Inches(0.45) + Inches(4.2) * i
        add_rect(s, x, Inches(1.25), Inches(4.0), Inches(3.15), fill=CLOUD)
        add_rect(s, x, Inches(1.25), Inches(4.0), Inches(0.5),
                 fill=GOLD if i == 2 else NAVY)
        add_text(s, x, Inches(1.25), Inches(4.0), Inches(0.5), t,
                 size=16, bold=True,
                 color=NAVY if i == 2 else WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x, Inches(1.85), Inches(4.0), Inches(0.85), n,
                 size=26, bold=True, color=NAVY,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + Inches(0.15), Inches(2.75), Inches(3.7), Inches(1.4),
                 b, size=15, color=INK, align=PP_ALIGN.CENTER)

    add_rect(s, Inches(0.45), Inches(4.6), Inches(12.4), Inches(2.0), fill=NAVY)
    add_text(s, Inches(0.7), Inches(4.75), Inches(12.0), Inches(1.7),
             "试跑无人成交，业主只出 8 万。转正后月度 4 万，当月有佣金则抵扣。\n"
             "中介成交：业主合计不超过该档，中介约 70% + 我们约 30%。\n"
             "4,000 ㎡ 空一年大约少收 949 万。甲方自有未报备客户不付佣金。",
             size=16, color=WHITE)

    # 5 签法
    s = new_slide(prs)
    add_chrome(s, prs, page_no=5, label="签法",
               title="建议先勾 90 天",
               subtitle="收费已经按试跑对齐。12 个月独家等看完人再锁。")

    add_table(
        s, Inches(0.45), Inches(1.25), Inches(12.4), Inches(3.4),
        ["", "方案 B · 先 90 天（建议）", "方案 A · 12 个月独家"],
        [
            ["适合", "先看人干活", "已决心排他"],
            ["试跑费用", "只付包干 8 万", "8 万记作前 90 天，之后月度 4 万"],
            ["怎么停", "第 90 天可转正 / 再试 / 停", "满 6 个月无锚定意向可停月度"],
            ["合同章 / 报价", "您盖章 · 绿区我们直接报", "同样"],
        ],
        header_size=14, body_size=14,
        col_widths=[Inches(1.8), Inches(5.3), Inches(5.3)],
    )

    add_rect(s, Inches(0.45), Inches(4.9), Inches(12.4), Inches(1.7), fill=CLOUD)
    add_text(s, Inches(0.7), Inches(5.05), Inches(12.0), Inches(1.4),
             "绿区：租金 ≥ 6.5 元/㎡·天 · 免租 ≤ 9 个月 · 装补 ≤ 600 元/㎡\n"
             "不承诺政府政策、落户、公寓。不收租户或中介回扣。\n"
             "详细条款在 05 号全稿，会上先把简版确认栏勾完。",
             size=16, color=INK)

    # 6 请定
    s = new_slide(prs)
    add_chrome(s, prs, page_no=6, label="拍板",
               title="今天勾这四项就够",
               subtitle="数字写在商务简版确认栏。法务随后出全稿，不改收费结构。")

    asks = [
        ("01", "签法", "建议 B：先 90 天。"),
        ("02", "收费", "包干 8 万 · 首月 100% / 锚定 150%。"),
        ("03", "对接人", "唯一商务对接人 + 红区谁拍板。"),
        ("04", "绿区", "能否直接报价。每单都开会，节奏会断。"),
    ]
    for i, (n, k, v) in enumerate(asks):
        y = Inches(1.25) + Inches(1.25) * i
        add_round(s, Inches(0.5), y, Inches(1.15), Inches(1.05), n,
                  fill=GOLD, color=NAVY, size=20)
        add_text(s, Inches(1.9), y, Inches(2.2), Inches(1.05), k,
                 size=22, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(4.2), y, Inches(8.5), Inches(1.05), v,
                 size=18, color=INK, anchor=MSO_ANCHOR.MIDDLE)

    total = len(SLIDES)
    for sl in SLIDES:
        for shape in sl.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.text and run.text.strip().endswith(" / 0"):
                        run.text = run.text.replace(" / 0", f" / {total}")

    out = Path(__file__).resolve().parent.parent / "docs" / "advisory" / \
        "deck" / "冠松01楼-招商合作-会上版.pptx"
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    print(f"✓ Deck written: {out}  ({total} slides)")
    return out


if __name__ == "__main__":
    build()
