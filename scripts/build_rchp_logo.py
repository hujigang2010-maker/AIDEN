# -*- coding: utf-8 -*-
"""合成复旦大学住房政策研究中心（RCHP）组合标识。

校徽取自复旦大学官方标识系统（校徽原稿中的蓝色圆形校徽）。
中英文名称用 Noto 字体实排，版式对齐中心公开使用的
「校徽 + 红底 RCHP 字标 + 中英文全称」锁式组合。
"""
import os

from PIL import Image, ImageDraw, ImageFont, ImageFilter

BRAND = "/workspace/whitepaper/assets/brand"
ASSETS = "/workspace/whitepaper/assets"
SEAL_SHEET = os.path.join(BRAND, "fudan_seal.png")

FUDAN_RED = (196, 18, 32)
FUDAN_RED_DEEP = (154, 12, 26)
INK = (22, 24, 28)
INK_SOFT = (72, 76, 84)
WHITE = (255, 255, 255)
BLACK = (12, 12, 12)

FONT_SANS_B = "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"
FONT_SERIF_B = "/usr/share/fonts/opentype/noto/NotoSerifCJK-Bold.ttc"
FONT_SERIF_R = "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc"
FONT_LATIN_B = "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"
FONT_LATIN_R = "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf"

CN_NAME = "復旦大學住房政策研究中心"
EN_NAME = "Research Center for Housing Policy, Fudan University"

os.makedirs(BRAND, exist_ok=True)


def extract_blue_seal():
    sheet = Image.open(SEAL_SHEET).convert("RGBA")
    seal = sheet.crop((0, 0, 738, 738))
    # 圆形遮罩，去掉原稿黑底/透明边的锯齿外溢
    mask = Image.new("L", seal.size, 0)
    md = ImageDraw.Draw(mask)
    md.ellipse((2, 2, 735, 735), fill=255)
    mask = mask.filter(ImageFilter.GaussianBlur(0.6))
    out = Image.new("RGBA", seal.size, (0, 0, 0, 0))
    out.paste(seal, (0, 0))
    out.putalpha(mask)
    path = os.path.join(BRAND, "fudan_seal_blue.png")
    out.save(path)
    return out


def _font(path, size):
    return ImageFont.truetype(path, size)


def draw_rchp_box(w, h, letters=BLACK, box=FUDAN_RED, swoosh=None):
    """红底 RCHP 字标：H–P 底部连笔，P 尾部上扬箭形。"""
    if swoosh is None:
        swoosh = letters
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.rectangle((0, 0, w - 1, h - 1), fill=box)

    fnt = _font(FONT_LATIN_B, int(h * 0.62))
    text = "RCHP"
    tw = d.textlength(text, font=fnt)
    x0 = (w - tw) / 2
    y0 = int(h * 0.04)
    d.text((x0, y0), text, font=fnt, fill=letters)

    # H–P 底部连笔（加粗，保证红底上可见）
    hx = x0 + d.textlength("RC", font=fnt) - w * 0.02
    hy = int(h * 0.82)
    d.arc((hx, hy - int(h * 0.28), hx + int(w * 0.48), hy + int(h * 0.22)),
          start=12, end=168, fill=swoosh, width=max(5, h // 16))

    # P 右侧上扬小箭，略伸出色块
    px = int(x0 + tw - w * 0.04)
    py = int(h * 0.58)
    d.polygon(
        [(px, py),
         (px + int(w * 0.08), py - int(h * 0.10)),
         (px + int(w * 0.03), py),
         (px + int(w * 0.10), py + int(h * 0.22))],
        fill=swoosh,
    )
    return im


def _paste(base, im, xy):
    base.alpha_composite(im, xy)


def compose_lockup(mode="print"):
    """mode: print（白纸深色字）/ dark（深底浅色字）。"""
    dark = mode == "dark"
    seal_s, box_w, box_h, gap = 320, 560, 200, 40
    f_cn = _font(FONT_SERIF_B, 86)
    f_en = _font(FONT_LATIN_R, 34)
    # 先量中文全称宽度，使书法行成为视觉主轴
    probe = ImageDraw.Draw(Image.new("RGB", (8, 8)))
    tracking = 10
    cn_w = sum(probe.textlength(ch, font=f_cn) + tracking for ch in CN_NAME) - tracking
    en_w = probe.textlength(EN_NAME, font=f_en)
    top_w = seal_s + gap + box_w
    content_w = int(max(cn_w, en_w, top_w))
    pad = 28
    W = content_w + pad * 2
    H = 28 + seal_s + 36 + 110 + 50 + 28
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))

    box = draw_rchp_box(box_w, box_h, letters=BLACK, box=FUDAN_RED, swoosh=BLACK)
    cn_fill = (245, 245, 247) if dark else INK
    en_fill = (210, 212, 218) if dark else INK_SOFT

    top_x = pad + (content_w - top_w) // 2
    _paste(img, extract_blue_seal().resize((seal_s, seal_s), Image.Resampling.LANCZOS),
           (top_x, 28))
    _paste(img, box, (top_x + seal_s + gap, 28 + (seal_s - box_h) // 2))

    d = ImageDraw.Draw(img)
    cn_x = pad + (content_w - cn_w) / 2
    cn_y = 28 + seal_s + 28
    cursor = cn_x
    for ch in CN_NAME:
        d.text((cursor, cn_y), ch, font=f_cn, fill=cn_fill)
        cursor += d.textlength(ch, font=f_cn) + tracking
    d.text((pad + (content_w - en_w) / 2, cn_y + 108), EN_NAME, font=f_en, fill=en_fill)
    return img


def compose_mark(mode="color"):
    """仅校徽 + RCHP 色块，供页眉使用。"""
    seal = extract_blue_seal().resize((220, 220), Image.Resampling.LANCZOS)
    if mode == "mono":
        box = draw_rchp_box(380, 150, letters=BLACK, box=WHITE, swoosh=BLACK)
        # 白底字标加细红边，避免印在白纸上消失
        d = ImageDraw.Draw(box)
        d.rounded_rectangle((1, 1, 378, 148), radius=4, outline=FUDAN_RED, width=4)
    else:
        box = draw_rchp_box(380, 150, letters=BLACK, box=FUDAN_RED, swoosh=FUDAN_RED_DEEP)
    h = 240
    w = 220 + 28 + 380 + 16
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    _paste(img, seal, (8, (h - 220) // 2))
    _paste(img, box, (8 + 220 + 28, (h - 150) // 2))
    return img


def compose_dark_plate():
    """黑底完整标识，贴近中心对外使用的深色版。"""
    lockup = compose_lockup("dark")
    pad = 48
    canvas = Image.new("RGB", (lockup.width + pad * 2, lockup.height + pad * 2), (0, 0, 0))
    canvas.paste(lockup, (pad, pad), lockup)
    return canvas


def main():
    extract_blue_seal()
    print_logo = compose_lockup("print")
    dark_logo = compose_lockup("dark")
    mark = compose_mark("color")
    plate = compose_dark_plate()

    paths = {
        os.path.join(BRAND, "logo_rchp_print.png"): print_logo,
        os.path.join(BRAND, "logo_rchp_dark.png"): dark_logo,
        os.path.join(BRAND, "logo_rchp_mark.png"): mark,
        os.path.join(BRAND, "logo_rchp_blackbg.png"): plate,
        os.path.join(ASSETS, "logo_fudan_hprc.png"): print_logo,  # 兼容既有脚本路径
    }
    for path, im in paths.items():
        im.save(path)
        print("saved", path, im.size)


if __name__ == "__main__":
    main()
