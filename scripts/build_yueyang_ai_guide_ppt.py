# -*- coding: utf-8 -*-
"""生成《岳阳医院针灸科｜VPN及国外AI软件安装使用指南》卡通图解 PPT。

面向现场培训：一页一件事，短句加大图。插画位于 assets/illustrations/。
运行：python3 scripts/build_yueyang_ai_guide_ppt.py
"""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
IMG = ROOT / "assets" / "illustrations"
OUT = ROOT / "output" / "岳阳医院针灸科-VPN及国外AI软件安装使用指南.pptx"

FONT = "Microsoft YaHei"
W, H = 13.333, 7.5
TOTAL = 26

PAPER = "FFF6EA"
INK = "2A3340"
MUTED = "5C6B7A"
GREEN = "1F7A62"
GREEN_DEEP = "145744"
GREEN_SOFT = "E5F4EE"
CORAL = "D4533E"
CORAL_SOFT = "FDECEA"
GOLD = "A86B12"
GOLD_SOFT = "FFF3D6"
WHITE = "FFFFFF"
LINE = "E7D7C4"
TEAL = "2F6F8F"
TEAL_SOFT = "E7F2F8"


def rgb(value: str) -> RGBColor:
    value = value.lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def style_run(run, size, bold, color, underline=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = False
    run.font.underline = underline
    run.font.color.rgb = rgb(color)
    run.font.name = FONT
    r_pr = run._r.get_or_add_rPr()
    r_pr.set("lang", "zh-CN")
    r_pr.set("altLang", "en-US")
    r_pr.set("dirty", "0")
    for tag in ("a:latin", "a:ea", "a:cs"):
        node = r_pr.find(qn(tag))
        if node is None:
            node = r_pr.makeelement(qn(tag), {})
            r_pr.append(node)
        node.set("typeface", FONT)


def add_slide(prs, color=PAPER):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(color)
    slide.notes_slide.notes_text_frame.text = ""
    return slide


def notes(slide, text: str):
    slide.notes_slide.notes_text_frame.text = text.strip()


def rect(slide, x, y, w, h, fill, line=None, radius=0.12, line_pt=1.25):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)
    )
    try:
        shape.adjustments[0] = radius
    except Exception:
        pass
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    if line:
        shape.line.color.rgb = rgb(line)
        shape.line.width = Pt(line_pt)
    else:
        shape.line.fill.background()
    return shape


def oval(slide, x, y, w, h, fill):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    shape.line.fill.background()
    return shape


def textbox(slide, x, y, w, h, paragraphs, align=PP_ALIGN.LEFT, anchor="top"):
    """paragraphs: list of list of run tuples (text, size, bold, color, underline, url)."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.word_wrap = True
    frame.auto_size = None
    frame.margin_left = Inches(0.02)
    frame.margin_right = Inches(0.02)
    frame.margin_top = Inches(0.01)
    frame.margin_bottom = Inches(0.01)
    anchor_map = {"top": "t", "middle": "ctr", "bottom": "b"}
    frame._txBody.bodyPr.set("anchor", anchor_map[anchor])
    for index, runs in enumerate(paragraphs):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.alignment = align
        paragraph.space_before = Pt(0)
        paragraph.space_after = Pt(0)
        paragraph.line_spacing = 1.0
        for item in runs:
            text, size, bold, color = item[:4]
            underline = item[4] if len(item) > 4 else False
            url = item[5] if len(item) > 5 else None
            run = paragraph.add_run()
            run.text = text
            style_run(run, size, bold, color, underline)
            if url:
                run.hyperlink.address = url
    return box


def run(text, size, bold, color, underline=False, url=None):
    return (text, size, bold, color, underline, url)


def picture(slide, name, x, y, w):
    path = IMG / name
    slide.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w))


def footer(slide, page: int):
    rect(slide, 0.42, 7.08, 12.5, 0.015, LINE, radius=0.0)
    textbox(
        slide,
        0.42,
        7.12,
        8.6,
        0.30,
        [[run("岳阳医院针灸科  ·  一看就会的图解指南", 11, False, MUTED)]],
        anchor="middle",
    )
    textbox(
        slide,
        9.4,
        7.12,
        3.5,
        0.30,
        [[run(f"{page}  /  {TOTAL}", 11, False, MUTED)]],
        align=PP_ALIGN.RIGHT,
        anchor="middle",
    )


def header(slide, kicker, title, subtitle=None):
    textbox(slide, 0.46, 0.26, 10.5, 0.32, [[run(kicker, 13, True, GREEN)]])
    textbox(slide, 0.44, 0.52, 12.4, 0.62, [[run(title, 32, True, INK)]])
    if subtitle:
        textbox(slide, 0.46, 1.16, 12.3, 0.36, [[run(subtitle, 15, False, MUTED)]])


def star(slide, x, y, size, fill):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.STAR_5_POINT, Inches(x), Inches(y), Inches(size), Inches(size)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    shape.line.fill.background()
    return shape


def number_badge(slide, x, y, number, fill=GREEN, size=0.48, font=18):
    oval(slide, x, y, size, size, fill)
    textbox(
        slide,
        x,
        y,
        size,
        size,
        [[run(str(number), font, True, WHITE)]],
        align=PP_ALIGN.CENTER,
        anchor="middle",
    )


# ---------------------------------------------------------------------------
# 幻灯片
# ---------------------------------------------------------------------------


def slide_cover(prs):
    slide = add_slide(prs)
    rect(slide, 0, 0, 0.18, H, GREEN, radius=0.0)
    textbox(
        slide,
        0.55,
        0.42,
        6.4,
        0.36,
        [[run("上海中医药大学岳阳医院  ·  针灸科", 15, True, GREEN)]],
    )
    textbox(
        slide,
        0.52,
        1.35,
        6.8,
        2.15,
        [
            [run("VPN及国外AI软件", 40, True, INK)],
            [run("安装使用指南", 44, True, INK)],
        ],
    )
    textbox(
        slide,
        0.55,
        3.62,
        6.2,
        0.42,
        [[run("卡通图解  ·  一看就会", 20, True, CORAL)]],
    )
    textbox(
        slide,
        0.55,
        4.18,
        6.3,
        0.85,
        [
            [run("个人电脑和手机，跟着图做就行。", 16, False, MUTED)],
            [run("国内软件和国外软件，都从官网进入。", 16, False, MUTED)],
        ],
    )
    chips = [
        (0.55, "2026年11月9日"),
        (2.55, "Windows · Mac"),
        (4.55, "iPhone · Android"),
    ]
    for left, label in chips:
        rect(slide, left, 5.25, 1.85, 0.46, GREEN_SOFT, radius=0.5)
        textbox(
            slide,
            left,
            5.25,
            1.85,
            0.46,
            [[run(label, 12, True, GREEN_DEEP)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
    textbox(
        slide,
        0.55,
        6.55,
        6.5,
        0.40,
        [[run("先看官方页面，再决定要不要安装。", 15, True, INK)]],
    )
    picture(slide, "cover-scene.jpg", 6.15, 1.35, 6.85)
    picture(slide, "fudan-seal.png", 12.05, 0.18, 0.86)
    textbox(
        slide,
        10.05,
        0.40,
        1.85,
        0.42,
        [[run("复旦大学", 15, True, INK)]],
        align=PP_ALIGN.RIGHT,
        anchor="middle",
    )
    star(slide, 11.55, 1.12, 0.26, GOLD)
    star(slide, 0.32, 6.55, 0.22, CORAL)
    notes(
        slide,
        "开场只说三件事：今天学的是个人电脑和手机上的操作；先认官方页面；"
        "最后要能做一次科研提问。VPN 连上与否，不作为这节课的通过条件。",
    )
    return slide


def slide_three_rules(prs):
    slide = add_slide(prs)
    header(slide, "先看这里", "今天只记三句话")
    picture(slide, "mascot-hello.jpg", 9.55, 1.55, 3.45)
    rules = [
        ("1", GREEN, "先看网页，再装软件", "网页能完成的事，就先别安装。"),
        ("2", TEAL, "下载前看清三样", "产品名称、开发商、安装包签名。"),
        ("3", CORAL, "病人资料留在医院", "病历、照片和敏感信息不要上传。"),
    ]
    top = 1.58
    for number, color, title, desc in rules:
        rect(slide, 0.46, top, 8.85, 1.48, WHITE, LINE, radius=0.14)
        rect(slide, 0.46, top, 0.12, 1.48, color, radius=0.08)
        oval(slide, 0.82, top + 0.40, 0.68, 0.68, color)
        textbox(
            slide,
            0.82,
            top + 0.40,
            0.68,
            0.68,
            [[run(number, 22, True, WHITE)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        textbox(slide, 1.75, top + 0.28, 7.2, 0.50, [[run(title, 24, True, INK)]])
        textbox(slide, 1.75, top + 0.82, 7.2, 0.40, [[run(desc, 16, False, MUTED)]])
        top += 1.66
    footer(slide, 2)
    notes(
        slide,
        "这三句可以让大家跟着读一遍。后面的每一步，都是这三句的具体做法。",
    )
    return slide


def slide_prepare(prs):
    slide = add_slide(prs)
    header(slide, "动手之前", "准备这四样就够了")
    items = [
        (GREEN, GREEN_SOFT, "1", "自己的电脑或手机", "办公电脑要先问信息科。"),
        (TEAL, TEAL_SOFT, "2", "一个常用浏览器", "Chrome、Edge 或 Safari。"),
        (GOLD, GOLD_SOFT, "3", "自己的邮箱", "注册、登录、收验证码都用它。"),
        (CORAL, CORAL_SOFT, "4", "医院允许的网络", "按信息科的规定访问网络。"),
    ]
    for index, (color, soft, number, title, desc) in enumerate(items):
        col, row = index % 2, index // 2
        x = 0.46 + col * 6.35
        y = 1.58 + row * 2.35
        rect(slide, x, y, 6.10, 2.15, WHITE, LINE, radius=0.16)
        oval(slide, x + 0.28, y + 0.62, 0.78, 0.78, soft)
        textbox(
            slide,
            x + 0.28,
            y + 0.62,
            0.78,
            0.78,
            [[run(number, 24, True, color)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        textbox(slide, x + 1.28, y + 0.48, 4.5, 0.55, [[run(title, 22, True, INK)]])
        textbox(slide, x + 1.28, y + 1.12, 4.5, 0.50, [[run(desc, 16, False, MUTED)]])
    footer(slide, 3)
    notes(
        slide,
        "强调：优先使用个人设备。医院配发、接入内网或存有患者信息的电脑，"
        "没有信息科授权就不要安装第三方网络工具和外来软件。",
    )
    return slide


def slide_route(prs):
    slide = add_slide(prs)
    header(slide, "今天的路线", "一共四步，做完即可")
    steps = [
        ("1", "准备好", "设备和自己的账号。"),
        ("2", "核对来源", "看清是谁做的软件。"),
        ("3", "打开官网", "先试用网页版。"),
        ("4", "做一个提问", "用在针灸科研上。"),
    ]
    for index, (number, title, desc) in enumerate(steps):
        x = 0.46 + index * 3.20
        rect(slide, x, 1.62, 3.02, 3.85, WHITE, LINE, radius=0.16)
        oval(slide, x + 1.05, 1.92, 0.92, 0.92, GREEN if index < 3 else CORAL)
        textbox(
            slide,
            x + 1.05,
            1.92,
            0.92,
            0.92,
            [[run(number, 28, True, WHITE)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        textbox(
            slide,
            x + 0.16,
            3.10,
            2.70,
            0.70,
            [[run(title, 24, True, INK)]],
            align=PP_ALIGN.CENTER,
        )
        textbox(
            slide,
            x + 0.18,
            3.90,
            2.66,
            0.90,
            [[run(desc, 16, False, MUTED)]],
            align=PP_ALIGN.CENTER,
        )
        if index < 3:
            textbox(
                slide,
                x + 2.72,
                2.15,
                0.48,
                0.46,
                [[run("→", 22, True, GOLD)]],
                align=PP_ALIGN.CENTER,
            )
    rect(slide, 0.46, 5.68, 12.42, 1.18, GREEN_SOFT, radius=0.14)
    textbox(
        slide,
        0.72,
        5.84,
        11.9,
        0.86,
        [
            [run("某个软件在当地不能用，就换一个可以合法使用的。", 18, True, GREEN_DEEP)],
            [run("后面照样演示。不要改用说不清来源的安装包。", 16, False, INK)],
        ],
    )
    footer(slide, 4)
    notes(
        slide,
        "把路线指给大家看。验收看的是：能打开一个允许使用的工具，并完成一次科研提问和一次文献核对。"
        "VPN 是否连接成功，不写进通过标准。",
    )
    return slide


def slide_vpn(prs):
    slide = add_slide(prs)
    header(slide, "指定入口", "先核对，再安装")
    picture(slide, "mascot-check.jpg", 0.20, 1.48, 4.55)
    rect(slide, 4.90, 1.52, 7.95, 2.05, WHITE, LINE, radius=0.14)
    textbox(slide, 5.12, 1.64, 7.5, 0.32, [[run("在浏览器里打开这个地址", 14, True, MUTED)]])
    textbox(
        slide,
        5.12,
        2.00,
        7.5,
        0.48,
        [[run("ik11.net/23c586b", 26, True, GREEN_DEEP, True, "https://ik11.net/23c586b")]],
    )
    textbox(
        slide,
        5.12,
        2.58,
        7.5,
        0.78,
        [
            [run("打开后会跳到 lk11.net。", 16, False, INK)],
            [run("公开资料提到 Veee+，安装包尚未核实。", 16, False, INK)],
        ],
    )
    textbox(slide, 4.90, 3.72, 7.5, 0.36, [[run("页面上这三样，要能对上", 16, True, INK)]])
    checks = [("产品名称", GREEN_SOFT, GREEN_DEEP), ("开发商", TEAL_SOFT, TEAL), ("数字签名", GOLD_SOFT, GOLD)]
    for index, (label, fill, color) in enumerate(checks):
        x = 4.90 + index * 2.65
        rect(slide, x, 4.16, 2.50, 0.78, fill, radius=0.14)
        textbox(
            slide,
            x,
            4.16,
            2.50,
            0.78,
            [[run(label, 18, True, color)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
    rect(slide, 4.90, 5.14, 7.95, 1.70, CORAL_SOFT, radius=0.14)
    textbox(
        slide,
        5.12,
        5.30,
        7.55,
        1.38,
        [
            [run("三样对不上，先不要装。", 20, True, CORAL)],
            [run("这个跳转页还不能当作已经确认", 16, False, INK)],
            [run("安全的官方安装源。", 16, False, INK)],
        ],
    )
    footer(slide, 5)
    notes(
        slide,
        "指定入口是 https://ik11.net/23c586b ，现场会跳转到 lk11.net/?ref=23c586b。"
        "公开资料将 lk11.net 与 Veee+ 联系起来，但跳转后的下载页此前无法直接核验，"
        "因此不能把它讲成已经确认安全的官方源。请大家先看产品名称、开发商和安装包数字签名。"
        "对不上就暂停。第三方 VPN 不得部署到医院内网、诊疗系统电脑和存有患者信息的设备。"
        "连接测试须符合院内网络规定及当地法律。",
    )
    return slide


def slide_computer(prs):
    slide = add_slide(prs)
    header(slide, "电脑", "Windows 和 Mac，各五步")
    columns = [
        (
            GREEN,
            GREEN_SOFT,
            "Windows",
            [
                "打开指定链接",
                "看产品名和开发商",
                "有 Windows 版时，核对签名",
                "在允许的电脑上安装",
                "用自己的账号登录",
            ],
        ),
        (
            TEAL,
            TEAL_SOFT,
            "Mac",
            [
                "用 Safari 或 Chrome 打开",
                "确认有 macOS 版和开发商",
                "下载核对过的安装文件",
                "按说明放进“应用程序”",
                "登录，并允许必要权限",
            ],
        ),
    ]
    for col, (color, soft, title, steps) in enumerate(columns):
        x = 0.46 + col * 6.42
        rect(slide, x, 1.50, 6.20, 4.55, WHITE, LINE, radius=0.16)
        rect(slide, x, 1.50, 6.20, 0.62, soft, radius=0.12)
        textbox(
            slide,
            x,
            1.50,
            6.20,
            0.62,
            [[run(title, 20, True, color)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        for index, step in enumerate(steps):
            y = 2.32 + index * 0.70
            oval(slide, x + 0.24, y, 0.42, 0.42, color)
            textbox(
                slide,
                x + 0.24,
                y,
                0.42,
                0.42,
                [[run(str(index + 1), 14, True, WHITE)]],
                align=PP_ALIGN.CENTER,
                anchor="middle",
            )
            textbox(slide, x + 0.80, y - 0.02, 5.15, 0.48, [[run(step, 16, True, INK)]], anchor="middle")
    rect(slide, 0.46, 6.18, 12.42, 0.74, GOLD_SOFT, radius=0.12)
    textbox(
        slide,
        0.68,
        6.18,
        12.05,
        0.74,
        [[run("要购买时，先看价格、续费和退款。连接测试须符合医院规定和当地法律。", 16, True, INK)]],
        anchor="middle",
    )
    footer(slide, 6)
    notes(
        slide,
        "Windows 若需要管理员权限，先确认这台电脑允许安装。Mac 按官方说明放进应用程序文件夹，"
        "系统弹出权限时再授权，不要预先关闭系统安全设置。订阅价格、自动续费和退款规则看清楚再付款。"
        "培训现场不要代学员购买，也不要保存他人支付信息。",
    )
    return slide


def slide_phone_vpn(prs):
    slide = add_slide(prs)
    header(slide, "手机", "只从官方应用商店安装")
    picture(slide, "mascot-phone.jpg", 0.22, 1.42, 4.15)
    blocks = [
        (1.48, GREEN, GREEN_SOFT, "iPhone", ["打开 App Store", "搜索软件的英文名", "核对开发商", "用自己的账号登录"]),
        (3.72, TEAL, TEAL_SOFT, "Android", ["打开当地可用的官方商店", "核对开发商和应用名", "正常安装并登录", "保留系统安全检查"]),
    ]
    for y, color, soft, title, steps in blocks:
        rect(slide, 4.55, y, 8.32, 2.10, WHITE, LINE, radius=0.14)
        rect(slide, 4.55, y, 1.62, 2.10, soft, radius=0.12)
        textbox(
            slide,
            4.60,
            y + 0.62,
            1.52,
            0.80,
            [[run(title, 16, True, color)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        for index, step in enumerate(steps):
            sy = y + 0.14 + index * 0.46
            textbox(slide, 6.38, sy, 6.2, 0.42, [[run(f"{index + 1}   {step}", 16, True, INK)]], anchor="middle")
    rect(slide, 0.46, 6.00, 12.42, 0.90, CORAL_SOFT, radius=0.12)
    textbox(
        slide,
        0.68,
        6.00,
        12.0,
        0.90,
        [[run("不要用来历不明的安装包，也不要关闭手机的安全检查。", 18, True, CORAL)]],
        anchor="middle",
    )
    footer(slide, 7)
    notes(
        slide,
        "手机优先走对应的官方应用商店，核对开发商后再安装。不要用来历不明的 IPA、APK 或描述文件，"
        "也不要为了安装而关闭设备安全检查。商店里没有官方版本时，不要把来源不明的替代包当成正版。",
    )
    return slide


def slide_dont(prs):
    slide = add_slide(prs)
    header(slide, "请停一下", "这三件事不要做")
    picture(slide, "mascot-stop.jpg", 9.65, 1.45, 3.35)
    items = [
        ("1", "不要装进医院内网和诊疗电脑", "包括门诊、病房和科室业务系统使用的电脑。"),
        ("2", "不要装进存有病人信息的设备", "病历、影像、照片和联系方式都算。"),
        ("3", "不要把 VPN 连上当成通过", "连上网络，不等于国外软件可以注册或使用。"),
    ]
    for index, (number, title, desc) in enumerate(items):
        y = 1.52 + index * 1.72
        rect(slide, 0.46, y, 8.95, 1.58, WHITE, LINE, radius=0.14)
        oval(slide, 0.70, y + 0.45, 0.68, 0.68, CORAL)
        textbox(
            slide,
            0.70,
            y + 0.45,
            0.68,
            0.68,
            [[run(number, 22, True, WHITE)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        textbox(slide, 1.60, y + 0.28, 7.5, 0.50, [[run(title, 20, True, INK)]])
        textbox(slide, 1.60, y + 0.86, 7.5, 0.46, [[run(desc, 15, False, MUTED)]])
    footer(slide, 8)
    notes(
        slide,
        "这三条是硬边界。第三方 VPN 不得擅自部署到医院内网、诊疗系统电脑及存有患者信息的设备。"
        "即使个人设备连接成功，也不能向学员承诺 ChatGPT 等平台一定可以注册。",
    )
    return slide


def slide_not_equal(prs):
    slide = add_slide(prs)
    header(slide, "容易误会的一点", "网络通了，软件不一定能用")
    rect(slide, 0.46, 1.55, 3.70, 2.15, GREEN_SOFT, radius=0.16)
    textbox(
        slide,
        0.62,
        1.78,
        3.40,
        1.70,
        [
            [run("网络已经连上", 22, True, GREEN_DEEP)],
            [run("只说明这条路通了", 15, False, INK)],
        ],
        align=PP_ALIGN.CENTER,
    )
    textbox(
        slide,
        4.22,
        2.10,
        1.22,
        0.90,
        [[run("但是", 22, True, CORAL)]],
        align=PP_ALIGN.CENTER,
        anchor="middle",
    )
    rect(slide, 5.50, 1.55, 7.35, 2.15, CORAL_SOFT, radius=0.16)
    textbox(
        slide,
        5.70,
        1.72,
        6.95,
        1.80,
        [
            [run("软件未必能打开", 22, True, CORAL)],
            [run("还要看地区和账号", 15, False, INK)],
        ],
        align=PP_ALIGN.CENTER,
    )
    rect(slide, 0.46, 3.92, 12.42, 1.45, WHITE, LINE, radius=0.14)
    textbox(
        slide,
        0.70,
        4.08,
        11.95,
        1.15,
        [
            [run("ChatGPT 当前的支持地区名单里，没有中国大陆。", 18, True, INK)],
            [run("在名单以外访问，官方说明可能导致账号被限制。", 16, False, MUTED)],
        ],
    )
    rect(slide, 0.46, 5.52, 12.42, 1.35, GREEN_SOFT, radius=0.14)
    textbox(
        slide,
        0.70,
        5.68,
        11.95,
        1.05,
        [
            [run("现场这样做", 14, True, GREEN)],
            [run("打不开国外软件时，就打开 Kimi 或 WorkBuddy。", 18, True, GREEN_DEEP)],
        ],
    )
    footer(slide, 9)
    notes(
        slide,
        "依据 OpenAI 帮助中心《ChatGPT Supported Countries》：支持名单中没有中国大陆；"
        "在名单以外访问或提供访问，可能导致账号被限制。香港、澳门同样不在该名单中，台湾在名单中。"
        "讲这一页时不要演示如何绕过地区限制，不要使用不真实的身份信息。"
        "平台不可用时，直接改用 Kimi 或 WorkBuddy；文献整理可以回到 Consensus、Zotero 和 PubMed。",
    )
    return slide


def slide_bundles(prs):
    slide = add_slide(prs)
    header(slide, "今天下午", "先看这四组", "国内的先演示。国外的打不开，就换下一组。")
    bundles = [
        (GREEN, GREEN_SOFT, "国内先用", "Kimi + WorkBuddy", "网页或电脑上就能开工"),
        (TEAL, TEAL_SOFT, "写作", "ChatGPT + Claude", "课题构思、改论文段落"),
        (CORAL, CORAL_SOFT, "整理", "Gemini + 笔记本", "课件提纲、对照公开论文"),
        (GOLD, GOLD_SOFT, "文献", "Consensus + Zotero", "找证据、管参考文献"),
    ]
    for index, (color, soft, tag, title, desc) in enumerate(bundles):
        col, row = index % 2, index // 2
        x = 0.46 + col * 6.42
        y = 1.70 + row * 2.45
        rect(slide, x, y, 6.20, 2.28, WHITE, LINE, radius=0.16)
        rect(slide, x, y, 6.20, 0.14, color, radius=0.08)
        rect(slide, x + 0.28, y + 0.38, 1.72, 0.42, soft, radius=0.5)
        textbox(
            slide,
            x + 0.28,
            y + 0.38,
            1.72,
            0.42,
            [[run(tag, 14, True, color)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        textbox(slide, x + 0.28, y + 0.98, 5.7, 0.52, [[run(title, 22, True, INK)]])
        textbox(slide, x + 0.28, y + 1.55, 5.7, 0.46, [[run(desc, 16, False, MUTED)]])
    footer(slide, 10)
    notes(
        slide,
        "下午先把 Kimi 或 WorkBuddy 演示成功。Cursor、Muse、Antigravity 和 Typeless 在后面单独一页。"
        "Zotero 不是聊天软件，写论文时用来管文献。",
    )
    return slide


def slide_catalog(prs):
    slide = add_slide(prs)
    header(slide, "认准地址", "只点这些官方页面")
    tools = [
        ("Kimi", "长文、提纲、科研问答", "kimi.com", "https://www.kimi.com/", "国内"),
        ("WorkBuddy", "腾讯的办公助手", "workbuddy.cn", "https://www.workbuddy.cn/", "国内"),
        ("ChatGPT", "问答、写作、整理文档", "chatgpt.com/download", "https://chatgpt.com/download/", "优先"),
        ("Claude", "长文分析、论文润色", "claude.com/download", "https://claude.com/download", "优先"),
        ("Gemini", "问答、资料分析、课件", "gemini.google.com", "https://gemini.google.com", "优先"),
        ("Gemini Notebook", "对照自己上传的资料", "notebook.google.com", "https://notebook.google.com", "优先"),
        ("Consensus", "比较医学研究证据", "consensus.app", "https://consensus.app", "文献"),
        ("Zotero", "管文献，不是聊天软件", "zotero.org/download", "https://www.zotero.org/download/", "文献"),
        ("Cursor", "让 AI 帮你改文件", "cursor.com/download", "https://cursor.com/download", "工作台"),
        ("Muse", "Meta 的个人助手", "ai.meta.com/muse", "https://ai.meta.com/muse/", "工作台"),
        ("Antigravity", "谷歌的智能工作台", "antigravity.google", "https://antigravity.google/download/", "工作台"),
        ("Typeless", "说话，自动变成文字", "typeless.com", "https://www.typeless.com/zh-cn", "小软件"),
    ]
    for index, (name, use, label, url, tag) in enumerate(tools):
        col = 0 if index < 6 else 1
        row = index if index < 6 else index - 6
        x = 0.42 + col * 6.46
        y = 1.40 + row * 0.90
        rect(slide, x, y, 6.28, 0.82, WHITE, LINE, radius=0.14)
        textbox(slide, x + 0.16, y + 0.04, 3.4, 0.36, [[run(name, 15, True, INK)]])
        textbox(slide, x + 0.16, y + 0.40, 3.5, 0.34, [[run(use, 12, False, MUTED)]])
        textbox(
            slide,
            x + 3.55,
            y + 0.10,
            2.55,
            0.40,
            [[run(label, 11, True, GREEN_DEEP, True, url)]],
            align=PP_ALIGN.RIGHT,
        )
        tag_fill = {
            "国内": GREEN_SOFT,
            "优先": TEAL_SOFT,
            "文献": GOLD_SOFT,
            "工作台": CORAL_SOFT,
            "小软件": GREEN_SOFT,
        }[tag]
        tag_color = {
            "国内": GREEN_DEEP,
            "优先": TEAL,
            "文献": GOLD,
            "工作台": CORAL,
            "小软件": GREEN_DEEP,
        }[tag]
        rect(slide, x + 4.62, y + 0.44, 1.48, 0.28, tag_fill, radius=0.5)
        textbox(
            slide,
            x + 4.62,
            y + 0.44,
            1.48,
            0.28,
            [[run(tag, 11, True, tag_color)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
    footer(slide, 11)
    notes(
        slide,
        "Gemini Notebook 是 NotebookLM 自 2026 年 7 月起使用的新名称，仍从 notebook.google.com 进入。"
        "Kimi 官网 kimi.com，WorkBuddy 是腾讯的办公工作台，官网 workbuddy.cn。"
        "Cursor 从 cursor.com/download 安装。Muse 是 Meta 的个人智能助手，介绍页为 ai.meta.com/muse，"
        "目前主要在美国提供。Antigravity 是谷歌的智能工作台，下载页 antigravity.google/download。"
        "Typeless 是语音听写小软件，官网 typeless.com。本次培训不再演示 Perplexity 和 Elicit。"
        "Grok 仍可从 grok.com 选装，这次不单独讲。所有产品先确认账号和地区。不要用不真实资料注册。",
    )
    return slide


def install_slide(prs, page, kicker, title, url_label, url, steps, question, image, image_note):
    slide = add_slide(prs)
    header(slide, kicker, title)
    textbox(
        slide,
        0.46,
        1.28,
        8.2,
        0.40,
        [[run(url_label, 16, True, GREEN_DEEP, True, url)]],
    )
    for index, step in enumerate(steps):
        y = 1.78 + index * 0.72
        oval(slide, 0.50, y, 0.48, 0.48, GREEN)
        textbox(
            slide,
            0.50,
            y,
            0.48,
            0.48,
            [[run(str(index + 1), 16, True, WHITE)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        textbox(slide, 1.15, y - 0.02, 6.3, 0.52, [[run(step, 18, True, INK)]], anchor="middle")
    picture(slide, image, 9.20, 1.48, 3.65)
    rect(slide, 0.46, 5.48, 12.42, 1.40, GOLD_SOFT, radius=0.14)
    textbox(slide, 0.70, 5.58, 12.0, 0.32, [[run(image_note, 13, True, GOLD)]])
    textbox(slide, 0.70, 5.94, 12.0, 0.78, [[run(question, 16, True, INK)]])
    footer(slide, page)
    return slide


def slide_chatgpt(prs):
    slide = install_slide(
        prs,
        14,
        "写作",
        "安装 ChatGPT",
        "chatgpt.com/download",
        "https://chatgpt.com/download/",
        [
            "打开官方下载页",
            "选择 Windows 或 macOS",
            "运行官方安装程序",
            "打开软件，登录自己的账号",
            "在允许使用的地区进入对话",
        ],
        "请解释针灸临床研究中随机对照试验的基本设计，并列出常见的偏倚风险及控制措施。",
        "mascot-laptop.jpg",
        "登录后，试着问这一句",
    )
    notes(
        slide,
        "桌面客户端从 chatgpt.com/download 下载。macOS 需要较新的系统版本，页面会按芯片给出对应安装包。"
        "若出现不支持的国家或地区，就停止这一款，改演示其他可以合法使用的工具。不要更换身份信息反复尝试。",
    )
    return slide


def slide_claude(prs):
    slide = install_slide(
        prs,
        15,
        "写作",
        "安装 Claude",
        "claude.com/download",
        "https://claude.com/download",
        [
            "打开官方下载页",
            "选择 Windows 或 macOS",
            "按页面说明完成安装",
            "打开 Claude",
            "在服务地区允许时登录",
        ],
        "请设计一份针灸治疗慢性疼痛的科研项目申请书大纲，包括研究背景、研究问题、方法学和预期成果。",
        "mascot-papers.jpg",
        "登录后，试着问这一句",
    )
    notes(
        slide,
        "Claude 桌面版从 claude.com/download 获取，免费账号也可以使用客户端，部分功能需要付费。"
        "同样先确认服务地区。培训优先使用免费或试用功能。",
    )
    return slide


def slide_gemini(prs):
    slide = add_slide(prs)
    header(slide, "第二组  ·  整理", "使用 Google Gemini")
    textbox(
        slide,
        0.46,
        1.28,
        8.0,
        0.38,
        [[run("gemini.google.com", 16, True, GREEN_DEEP, True, "https://gemini.google.com")]],
    )
    steps = [
        "打开上面的网页",
        "用符合要求的 Google 账号登录",
        "优先使用网页版",
        "手机再到官方商店搜索 Google Gemini",
        "登录后做一次普通问答",
    ]
    for index, step in enumerate(steps):
        y = 1.78 + index * 0.70
        oval(slide, 0.50, y, 0.46, 0.46, TEAL)
        textbox(
            slide,
            0.50,
            y,
            0.46,
            0.46,
            [[run(str(index + 1), 15, True, WHITE)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        textbox(slide, 1.14, y - 0.02, 6.5, 0.50, [[run(step, 18, True, INK)]], anchor="middle")
    picture(slide, "mascot-laptop.jpg", 9.25, 1.48, 3.60)
    rect(slide, 0.46, 5.45, 12.42, 1.42, TEAL_SOFT, radius=0.14)
    textbox(slide, 0.70, 5.58, 12.0, 0.30, [[run("登录后，试着问这一句", 13, True, TEAL)]])
    textbox(
        slide,
        0.70,
        5.95,
        12.0,
        0.72,
        [[run("请制作一份针灸循证医学研究方法的教学课件提纲，适合研究生课堂使用。", 16, True, INK)]],
    )
    footer(slide, 16)
    notes(
        slide,
        "Gemini 先用网页。手机应用只在当地官方应用商店可以搜到、并且开发商核对无误时再安装。",
    )
    return slide


def slide_notebook(prs):
    slide = add_slide(prs)
    header(slide, "第二组  ·  整理", "Gemini Notebook", "2026 年起，NotebookLM 使用这个新名字。")
    textbox(
        slide,
        0.46,
        1.58,
        7.4,
        0.36,
        [[run("notebook.google.com", 16, True, GREEN_DEEP, True, "https://notebook.google.com")]],
    )
    steps = [
        ("1", "登录受支持的 Google 账号"),
        ("2", "新建一个研究笔记本"),
        ("3", "只上传有权使用的公开资料"),
        ("4", "针对资料提问，并回原文核对"),
    ]
    for index, (number, step) in enumerate(steps):
        y = 2.08 + index * 0.78
        oval(slide, 0.50, y, 0.46, 0.46, TEAL)
        textbox(
            slide,
            0.50,
            y,
            0.46,
            0.46,
            [[run(number, 15, True, WHITE)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        textbox(slide, 1.14, y, 6.6, 0.48, [[run(step, 18, True, INK)]], anchor="middle")
    books = [
        (GREEN, GREEN_SOFT, "笔记本 1", "针灸临床循证"),
        (CORAL, CORAL_SOFT, "笔记本 2", "课题与参考文献"),
        (GOLD, GOLD_SOFT, "笔记本 3", "研究生教学资料"),
    ]
    for index, (color, soft, label, name) in enumerate(books):
        x = 8.05
        y = 1.50 + index * 1.28
        rect(slide, x, y, 4.82, 1.16, soft, radius=0.12)
        rect(slide, x, y, 0.14, 1.16, color, radius=0.08)
        textbox(slide, x + 0.32, y + 0.14, 4.3, 0.32, [[run(label, 12, True, color)]])
        textbox(slide, x + 0.32, y + 0.50, 4.3, 0.46, [[run(name, 18, True, INK)]])
    rect(slide, 0.46, 5.40, 12.42, 1.48, CORAL_SOFT, radius=0.14)
    textbox(
        slide,
        0.70,
        5.58,
        11.95,
        1.15,
        [
            [run("可以上传：已经取得使用权的公开论文和课件。", 16, True, INK)],
            [run("不要上传：未经医院授权的真实病历、患者照片和其他敏感资料。", 16, True, CORAL)],
        ],
    )
    footer(slide, 17)
    notes(
        slide,
        "Google 于 2026 年 7 月 16 日宣布 NotebookLM 更名为 Gemini Notebook，产品仍独立使用，入口仍是 notebook.google.com。"
        "它的回答要回到上传的原文核对。患者资料不进入任何通用 AI 或笔记本。",
    )
    return slide


def slide_kimi(prs):
    slide = add_slide(prs)
    header(slide, "国内先用", "打开 Kimi")
    star(slide, 8.55, 1.28, 0.28, GOLD)
    rect(slide, 8.95, 1.22, 3.85, 0.78, WHITE, LINE, radius=0.28)
    textbox(
        slide,
        9.05,
        1.22,
        3.65,
        0.78,
        [[run("国内网页就能用", 18, True, GREEN_DEEP)]],
        align=PP_ALIGN.CENTER,
        anchor="middle",
    )
    textbox(
        slide,
        0.46,
        1.28,
        7.6,
        0.40,
        [[run("kimi.com", 18, True, GREEN_DEEP, True, "https://www.kimi.com/")]],
    )
    steps = [
        "用浏览器打开官网",
        "用自己的手机号注册",
        "先用网页，不必马上安装",
        "只上传有权使用的公开论文",
        "请它整理成一段提纲",
    ]
    for index, step in enumerate(steps):
        y = 1.82 + index * 0.68
        oval(slide, 0.50, y, 0.46, 0.46, GREEN)
        textbox(
            slide,
            0.50,
            y,
            0.46,
            0.46,
            [[run(str(index + 1), 15, True, WHITE)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        textbox(slide, 1.14, y, 7.2, 0.48, [[run(step, 18, True, INK)]], anchor="middle")
    picture(slide, "mascot-spark.jpg", 9.35, 2.18, 3.15)
    rect(slide, 0.46, 5.40, 12.42, 1.48, GREEN_SOFT, radius=0.16)
    textbox(slide, 0.70, 5.52, 12.0, 0.32, [[run("试着问这一句", 13, True, GREEN)]])
    textbox(
        slide,
        0.70,
        5.88,
        12.0,
        0.80,
        [[run("请用大白话解释：针灸随机对照试验怎样减少偏倚？", 18, True, INK)]],
    )
    footer(slide, 12)
    notes(
        slide,
        "Kimi 是月之暗面的 AI 助手，官网 https://www.kimi.com/ 。优先用网页。"
        "可以上传已经有权使用的公开论文，不要上传病历、患者照片和其他敏感资料。"
        "这是国外软件打不开时的首选演示。",
    )
    return slide


def slide_workbuddy(prs):
    slide = add_slide(prs)
    header(slide, "国内先用", "安装 WorkBuddy")
    textbox(
        slide,
        0.46,
        1.28,
        8.0,
        0.38,
        [[
            run("workbuddy.cn", 18, True, GREEN_DEEP, True, "https://www.workbuddy.cn/"),
            run("    ·    腾讯出品", 16, True, TEAL),
        ]],
    )
    steps = [
        "打开官网，进入下载页",
        "选择 Windows 或 Mac",
        "安装后用自己的账号登录",
        "用一句话把任务交给它",
        "只让它打开公开资料文件夹",
    ]
    for index, step in enumerate(steps):
        y = 1.78 + index * 0.68
        oval(slide, 0.50, y, 0.46, 0.46, TEAL)
        textbox(
            slide,
            0.50,
            y,
            0.46,
            0.46,
            [[run(str(index + 1), 15, True, WHITE)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        textbox(slide, 1.14, y, 7.3, 0.48, [[run(step, 18, True, INK)]], anchor="middle")
    picture(slide, "mascot-laptop.jpg", 9.40, 1.48, 3.30)
    rect(slide, 0.46, 5.32, 12.42, 1.56, GOLD_SOFT, radius=0.16)
    textbox(slide, 0.70, 5.44, 12.0, 0.32, [[run("可以这样交代", 13, True, GOLD)]])
    textbox(
        slide,
        0.70,
        5.80,
        12.0,
        0.90,
        [[run("请把针灸治疗慢性腰痛整理成给研究生看的一页提纲。病历文件夹不要交给它。", 16, True, INK)]],
    )
    footer(slide, 13)
    notes(
        slide,
        "WorkBuddy 是腾讯的 AI 办公工作台，官网 https://www.workbuddy.cn/ ，下载页 https://www.workbuddy.cn/app-download/ 。"
        "它能读取你授权的本地文件夹。现场只指定公开论文或教学资料，不要把病历、影像和患者联系方式所在的文件夹交给它。",
    )
    return slide


def slide_consensus(prs):
    slide = add_slide(prs)
    header(slide, "找证据", "打开 Consensus")
    textbox(
        slide,
        0.46,
        1.28,
        7.8,
        0.38,
        [[run("consensus.app", 18, True, GREEN_DEEP, True, "https://consensus.app")]],
    )
    steps = [
        "打开官网并登录",
        "在搜索框输入英文问题",
        "看它列出的研究",
        "点开原文，再决定要不要引用",
    ]
    for index, step in enumerate(steps):
        y = 1.82 + index * 0.78
        oval(slide, 0.50, y, 0.48, 0.48, CORAL)
        textbox(
            slide,
            0.50,
            y,
            0.48,
            0.48,
            [[run(str(index + 1), 16, True, WHITE)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        textbox(slide, 1.16, y, 7.4, 0.50, [[run(step, 20, True, INK)]], anchor="middle")
    picture(slide, "mascot-papers.jpg", 9.50, 1.48, 3.25)
    rect(slide, 0.46, 5.10, 12.42, 1.78, GOLD_SOFT, radius=0.16)
    textbox(
        slide,
        0.68,
        5.22,
        12.0,
        1.52,
        [
            [run("试这句：与假针刺相比，针灸对膝骨关节炎是否有效？", 16, True, INK)],
            [run("Is acupuncture effective for knee osteoarthritis compared with sham acupuncture?", 13, False, MUTED)],
            [run("它不能代替 PubMed、Embase、Cochrane。引用前打开原文。", 15, True, GOLD)],
        ],
    )
    footer(slide, 18)
    notes(
        slide,
        "英文检索式：Is acupuncture effective for knee osteoarthritis compared with sham acupuncture? "
        "Consensus 用来辅助看医学证据，不能代替 PubMed、Embase、Cochrane Library 的规范检索。"
        "本讲义已去掉 Perplexity 和 Elicit。引用前必须打开原始论文。",
    )
    return slide


def slide_agents(prs):
    slide = add_slide(prs)
    header(slide, "三个工作台", "Cursor、Muse、Antigravity")
    cards = [
        (
            GREEN,
            GREEN_SOFT,
            "Cursor",
            "cursor.com/download",
            "https://cursor.com/download",
            ["下载自己电脑的版本", "用自己的账号登录", "只打开公开资料文件夹"],
        ),
        (
            TEAL,
            TEAL_SOFT,
            "Muse",
            "ai.meta.com/muse",
            "https://ai.meta.com/muse/",
            ["商店里认准开发商 Meta", "重要操作它会先问你", "不要连接医院邮箱"],
        ),
        (
            CORAL,
            CORAL_SOFT,
            "Antigravity",
            "antigravity.google",
            "https://antigravity.google/download/",
            ["下载 Windows 或 Mac", "用自己的 Google 账号", "做完的结果人来看"],
        ),
    ]
    for index, (color, soft, name, label, url, steps) in enumerate(cards):
        x = 0.42 + index * 4.28
        rect(slide, x, 1.48, 4.12, 4.35, WHITE, LINE, radius=0.16)
        rect(slide, x, 1.48, 4.12, 1.35, soft, radius=0.14)
        textbox(slide, x + 0.18, 1.58, 3.76, 0.46, [[run(name, 22, True, INK)]])
        textbox(
            slide,
            x + 0.18,
            2.08,
            3.76,
            0.50,
            [[run(label, 12, True, color, True, url)]],
        )
        for step_index, step in enumerate(steps):
            y = 3.05 + step_index * 0.78
            oval(slide, x + 0.20, y, 0.40, 0.40, color)
            textbox(
                slide,
                x + 0.20,
                y,
                0.40,
                0.40,
                [[run(str(step_index + 1), 13, True, WHITE)]],
                align=PP_ALIGN.CENTER,
                anchor="middle",
            )
            textbox(slide, x + 0.70, y - 0.04, 3.20, 0.50, [[run(step, 14, True, INK)]], anchor="middle")
    rect(slide, 0.42, 6.00, 12.48, 0.90, CORAL_SOFT, radius=0.14)
    textbox(
        slide,
        0.64,
        6.00,
        12.1,
        0.90,
        [[run("三款都可能受地区限制。打不开，就回到 Kimi。不要把病历文件夹交给它们。", 16, True, CORAL)]],
        anchor="middle",
    )
    footer(slide, 20)
    notes(
        slide,
        "Cursor 官网下载页 https://cursor.com/download 。它是和 AI 一起改文件的桌面软件。"
        "Muse 是 Meta 于 2026 年推出的个人智能助手，介绍页 https://ai.meta.com/muse/ ，"
        "目前主要在美国的 iOS、Android 和 muse.ai 提供。手机安装时核对开发商为 Meta。"
        "不要把医院邮箱、患者系统或支付信息连给 Muse。重要操作应先经本人确认。"
        "Antigravity 是谷歌的智能工作台，下载页 https://antigravity.google/download/ ，用 Google 账号登录。"
        "三款都不能保证在中国大陆可以直接使用。",
    )
    return slide


def slide_typeless(prs):
    slide = add_slide(prs)
    header(slide, "小软件", "Typeless：说话就能打字")
    picture(slide, "mascot-mic.jpg", 0.18, 1.42, 3.95)
    star(slide, 4.15, 1.45, 0.28, GOLD)
    rect(slide, 4.70, 1.48, 8.15, 0.85, WHITE, LINE, radius=0.28)
    textbox(
        slide,
        4.90,
        1.48,
        7.8,
        0.85,
        [[run("按住快捷键，把话说完。嗯、啊会被去掉。", 18, True, INK)]],
        align=PP_ALIGN.CENTER,
        anchor="middle",
    )
    steps = [
        ("1", "打开 typeless.com，下载电脑版"),
        ("2", "手机则在官方商店搜索 Typeless"),
        ("3", "允许麦克风，它才能听见你"),
        ("4", "口述一句科研笔记，看它写成文字"),
    ]
    textbox(
        slide,
        5.48,
        2.42,
        6.8,
        0.36,
        [[run("typeless.com", 14, True, GREEN_DEEP, True, "https://www.typeless.com/zh-cn")]],
    )
    for index, (number, step) in enumerate(steps):
        y = 2.82 + index * 0.64
        oval(slide, 4.85, y, 0.46, 0.46, CORAL)
        textbox(
            slide,
            4.85,
            y,
            0.46,
            0.46,
            [[run(number, 15, True, WHITE)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        textbox(slide, 5.48, y, 7.1, 0.48, [[run(step, 16, True, INK)]], anchor="middle")
    rect(slide, 0.46, 5.55, 12.42, 1.32, CORAL_SOFT, radius=0.14)
    textbox(
        slide,
        0.70,
        5.70,
        12.0,
        1.05,
        [
            [run("适合口述提纲和教学句子。", 16, True, INK)],
            [run("不要口述真实病历、患者姓名和联系方式。", 16, True, CORAL)],
        ],
    )
    footer(slide, 21)
    notes(
        slide,
        "Typeless 是语音听写小软件，中文介绍页 https://www.typeless.com/zh-cn ，下载页 https://www.typeless.com/downloads 。"
        "支持 Mac、Windows、Linux、iPhone 和 Android。官方说明语音不留在云端，也不用于训练模型。"
        "即便如此，培训现场不要口述真实病历和患者信息。电脑版需要麦克风权限，按官方说明授权即可，不要关闭系统里其他安全检查。",
    )
    return slide



def slide_zotero(prs):
    slide = add_slide(prs)
    header(slide, "文献管理", "安装 Zotero", "它不是聊天机器人。写论文、标书和综述时用它管文献。")
    picture(slide, "mascot-papers.jpg", 8.85, 1.55, 4.05)
    steps = [
        "打开 zotero.org/download",
        "安装 Windows 或 macOS 版",
        "需要的话，建立自己的同步账号",
        "安装对应浏览器的 Connector",
        "从 PubMed 等数据库导入题录",
        "检查作者、期刊、年份和 DOI",
    ]
    for index, step in enumerate(steps):
        y = 1.70 + index * 0.72
        oval(slide, 0.50, y, 0.46, 0.46, GOLD)
        textbox(
            slide,
            0.50,
            y,
            0.46,
            0.46,
            [[run(str(index + 1), 15, True, WHITE)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        if index == 0:
            textbox(
                slide,
                1.14,
                y,
                7.4,
                0.48,
                [[run("打开 ", 18, True, INK), run("zotero.org/download", 18, True, GREEN_DEEP, True, "https://www.zotero.org/download/")]],
                anchor="middle",
            )
        else:
            textbox(slide, 1.14, y, 7.4, 0.48, [[run(step, 18, True, INK)]], anchor="middle")
    footer(slide, 19)
    notes(
        slide,
        "下载页：https://www.zotero.org/download/ 。导入后当场抽查一条记录的作者、期刊、年份和 DOI，"
        "让大家看到题录会出错，需要人工检查。",
    )
    return slide


def slide_phone_ai(prs):
    slide = add_slide(prs)
    header(slide, "手机上的 AI 软件", "商店里核对开发商，再下载")
    columns = [
        (
            "iPhone",
            ["打开 App Store", "搜索准确的英文名称", "核对开发商和官网", "看系统版本和地区是否支持", "下载后用自己的账号登录"],
        ),
        (
            "Android",
            ["在允许使用的地区打开 Google Play", "搜索目标软件", "核对开发商和应用名称", "正常安装并登录", "商店没有官方版就先用网页"],
        ),
    ]
    for index, (title, steps) in enumerate(columns):
        x = 0.46 + index * 6.42
        rect(slide, x, 1.48, 6.22, 4.15, WHITE, LINE, radius=0.14)
        textbox(slide, x + 0.24, 1.62, 5.7, 0.46, [[run(title, 22, True, INK)]])
        for step_index, step in enumerate(steps):
            y = 2.24 + step_index * 0.62
            textbox(slide, x + 0.28, y, 5.7, 0.50, [[run(f"{step_index + 1}    {step}", 16, True, INK)]])
    rect(slide, 0.46, 5.80, 12.42, 1.08, CORAL_SOFT, radius=0.14)
    textbox(
        slide,
        0.70,
        5.96,
        12.0,
        0.78,
        [
            [run("不要从网盘或聊天群下载。Kimi、Typeless 也只在官方商店搜索。", 16, True, CORAL)],
        ],
        anchor="middle",
    )
    footer(slide, 22)
    notes(
        slide,
        "手机安装与电脑同一原则：只认官方商店和可核对的开发商。地区没有上架时，用网页版或更换工具，不使用第三方安装包。",
    )
    return slide


def slide_scenes(prs):
    slide = add_slide(prs)
    header(slide, "针灸科最值得演示的", "三个场景")
    scenes = [
        (GREEN, GREEN_SOFT, "1", "找文献", "用 Consensus 搜到论文，用 Kimi 读摘要，再回 PubMed 核对。", "当场点开一条来源"),
        (TEAL, TEAL_SOFT, "2", "想设计", "用 Kimi、ChatGPT 或 Claude 起草，再由人修改。", "草稿留给人修改"),
        (CORAL, CORAL_SOFT, "3", "理论文", "用笔记本和 Zotero 整理。想口述时打开 Typeless。", "抽查作者和 DOI"),
    ]
    for index, (color, soft, number, title, desc, action) in enumerate(scenes):
        x = 0.46 + index * 4.26
        rect(slide, x, 1.50, 4.08, 4.35, WHITE, LINE, radius=0.16)
        rect(slide, x, 1.50, 4.08, 1.20, soft, radius=0.14)
        oval(slide, x + 0.24, 1.74, 0.68, 0.68, color)
        textbox(
            slide,
            x + 0.24,
            1.74,
            0.68,
            0.68,
            [[run(number, 20, True, WHITE)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        textbox(slide, x + 1.08, 1.80, 2.75, 0.55, [[run(title, 24, True, INK)]], anchor="middle")
        textbox(slide, x + 0.22, 2.88, 3.64, 1.70, [[run(desc, 16, False, INK)]])
        rect(slide, x + 0.22, 4.72, 3.64, 0.88, soft, radius=0.12)
        textbox(
            slide,
            x + 0.32,
            4.72,
            3.44,
            0.88,
            [
                [run("现场就做", 12, True, color)],
                [run(action, 16, True, INK)],
            ],
            anchor="middle",
        )
    footer(slide, 23)
    notes(
        slide,
        "国外写作工具如果登录不了，场景 2 就改用 Kimi。不要为了演完某一款软件而违规注册。"
        "AI 草稿不能直接当作课题申报文本。文献必须回 PubMed 或期刊官网核对。",
    )
    return slide


def slide_faq(prs):
    slide = add_slide(prs)
    header(slide, "卡住了", "先看这五句")
    faqs = [
        ("软件下完打不开", "看系统是否支持、是不是官方安装包、这台电脑有没有管理限制。"),
        ("网页能开，但不能注册", "看邮箱、验证码、账号资格和地区限制。不要使用他人身份。"),
        ("有 VPN，国外 AI 仍不能用", "改用 Kimi 或 WorkBuddy。VPN 不能解决全部问题。"),
        ("软件提示要付款", "培训先用免费或试用。不要代付，也不要保存他人支付信息。"),
        ("AI 给出了查不到的文献", "回到 PubMed、期刊官网或 DOI 页面。核对不了的，不能写进论文。"),
    ]
    for index, (question, answer) in enumerate(faqs):
        y = 1.42 + index * 1.06
        rect(slide, 0.46, y, 12.42, 0.96, WHITE, LINE, radius=0.12)
        oval(slide, 0.66, y + 0.24, 0.48, 0.48, GREEN if index < 4 else CORAL)
        textbox(
            slide,
            0.66,
            y + 0.24,
            0.48,
            0.48,
            [[run(str(index + 1), 14, True, WHITE)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        textbox(slide, 1.32, y + 0.08, 11.2, 0.40, [[run(question, 16, True, INK)]])
        textbox(slide, 1.32, y + 0.48, 11.2, 0.40, [[run(answer, 14, False, MUTED)]])
    footer(slide, 24)
    notes(
        slide,
        "第 5 条建议现场做一次：让工具给一篇文献，当场用 DOI 或 PMID 检索。检索不到就删掉，不进入讲义和论文。",
    )
    return slide


def slide_check(prs):
    slide = add_slide(prs)
    header(slide, "离开之前", "现场打勾")
    items = [
        "这台设备允许安装",
        "VPN 来源已经核对",
        "打开 Kimi 或 WorkBuddy",
        "完成一次医学科研提问",
        "用 Consensus 搜文献并点开来源",
        "认识 Cursor、Muse、Antigravity",
        "试一次 Typeless 说话打字",
        "安装或了解 Zotero",
        "确认不会上传病人敏感信息",
    ]
    for index, item in enumerate(items):
        col = 0 if index < 5 else 1
        row = index if index < 5 else index - 5
        x = 0.46 + col * 6.42
        y = 1.48 + row * 0.82
        rect(slide, x, y, 6.22, 0.72, WHITE, LINE, radius=0.12)
        rect(slide, x + 0.16, y + 0.18, 0.36, 0.36, PAPER, GREEN, radius=0.15, line_pt=1.75)
        textbox(slide, x + 0.66, y, 5.35, 0.72, [[run(item, 16, True, INK)]], anchor="middle")
    rect(slide, 0.46, 5.70, 12.42, 1.18, GREEN_SOFT, radius=0.14)
    textbox(
        slide,
        0.70,
        5.84,
        12.0,
        0.92,
        [
            [run("这节课做到：会打开官网，会提一个科研问题，知道病人资料不能上传。", 16, True, GREEN_DEEP)],
            [run("VPN 没连上，这节课仍然可以通过。", 16, True, INK)],
        ],
    )
    footer(slide, 25)
    notes(
        slide,
        "打勾时口头确认两条：VPN 安装包若未能核对开发商和签名，这一项记为“未安装”，不判为失败。"
        "国外 AI 若因地区限制无法开通，改为完成 Kimi、WorkBuddy、Consensus 或 Zotero，仍然可以完成本次培训。",
    )
    return slide


def slide_close(prs):
    slide = add_slide(prs)
    rect(slide, 0, 0, 0.18, H, GREEN, radius=0.0)
    picture(slide, "mascot-done.jpg", 0.45, 1.15, 5.55)
    textbox(slide, 6.20, 0.55, 6.6, 0.36, [[run("培训结束时，能说出这三句", 14, True, GREEN)]])
    endings = [
        ("1", "我知道去哪个官网。"),
        ("2", "我会提一个科研问题。"),
        ("3", "我知道病人资料不能传。"),
    ]
    for index, (number, line) in enumerate(endings):
        y = 1.15 + index * 1.15
        rect(slide, 6.20, y, 6.60, 1.00, WHITE, LINE, radius=0.14)
        oval(slide, 6.40, y + 0.24, 0.52, 0.52, GREEN)
        textbox(
            slide,
            6.40,
            y + 0.24,
            0.52,
            0.52,
            [[run(number, 16, True, WHITE)]],
            align=PP_ALIGN.CENTER,
            anchor="middle",
        )
        textbox(slide, 7.10, y, 5.45, 1.00, [[run(line, 20, True, INK)]], anchor="middle")
    textbox(
        slide,
        6.20,
        4.75,
        6.6,
        1.35,
        [
            [run("本指南用于软件安装培训和科研学习。", 14, False, MUTED)],
            [run("医院的设备、网络和临床资料，", 14, False, MUTED)],
            [run("按信息科和医院的规定办理。", 14, False, MUTED)],
        ],
    )
    picture(slide, "fudan-seal.png", 6.20, 6.12, 0.72)
    textbox(slide, 7.12, 6.16, 5.6, 0.40, [[run("岳阳医院针灸科", 20, True, GREEN_DEEP)]])
    textbox(slide, 7.12, 6.54, 5.6, 0.30, [[run("复旦大学", 14, True, INK)]])
    textbox(slide, 7.12, 6.86, 5.6, 0.28, [[run("2026年11月9日", 13, False, MUTED)]])
    notes(
        slide,
        "收尾请大家用自己的话复述三句。不要把未核实的 VPN 安装包推荐到科室群作为必装软件。",
    )
    return slide


def build():
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    prs.core_properties.title = "岳阳医院针灸科｜VPN及国外AI软件安装使用指南"
    prs.core_properties.subject = "个人电脑与手机的卡通图解操作指南"
    prs.core_properties.category = "培训讲义"
    prs.core_properties.keywords = "岳阳医院,针灸科,安装指南,科研"
    prs.core_properties.comments = (
        "指定 VPN 入口跳转后的安装包尚未核实开发商与数字签名。"
        "VPN 连接成功不是本培训的通过条件。"
    )
    builders = [
        slide_cover,
        slide_three_rules,
        slide_prepare,
        slide_route,
        slide_vpn,
        slide_computer,
        slide_phone_vpn,
        slide_dont,
        slide_not_equal,
        slide_bundles,
        slide_catalog,
        slide_kimi,
        slide_workbuddy,
        slide_chatgpt,
        slide_claude,
        slide_gemini,
        slide_notebook,
        slide_consensus,
        slide_zotero,
        slide_agents,
        slide_typeless,
        slide_phone_ai,
        slide_scenes,
        slide_faq,
        slide_check,
        slide_close,
    ]
    if len(builders) != TOTAL:
        raise SystemExit(f"页数不一致：{len(builders)} != {TOTAL}")
    for builder in builders:
        builder(prs)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
