# -*- coding: utf-8 -*-
"""生成「复旦大学住房政策研究中心」文字标识（logo）。

说明：中心暂无公开可下载的官方标识，此处按复旦视觉规范的主色调
（复旦蓝）绘制一枚规范的中英文组合标识，供白皮书封面与页眉使用。
"""
import os

from PIL import Image, ImageDraw, ImageFont

FUDAN_BLUE = (14, 78, 155)
FUDAN_RED = (200, 16, 46)
WHITE = (255, 255, 255)

FONT = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"

OUT_DIR = "/workspace/whitepaper/assets"
os.makedirs(OUT_DIR, exist_ok=True)

W, H = 1900, 420
img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
d = ImageDraw.Draw(img)

cx, cy, r = 210, H // 2, 170
d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=FUDAN_BLUE)
d.ellipse([cx - r + 14, cy - r + 14, cx + r - 14, cy + r - 14],
          outline=WHITE, width=5)

roof = [(cx - 95, cy + 5), (cx, cy - 95), (cx + 95, cy + 5)]
d.polygon(roof, fill=WHITE)
d.polygon([(cx - 72, cy - 2), (cx, cy - 74), (cx + 72, cy - 2)], fill=FUDAN_BLUE)
bars = [(cx - 62, 42), (cx - 14, 72), (cx + 34, 102)]
for bx, bh in bars:
    d.rounded_rectangle([bx, cy + 108 - bh, bx + 30, cy + 108],
                        radius=6, fill=WHITE)
d.rounded_rectangle([cx - 80, cy + 120, cx + 80, cy + 132], radius=6, fill=FUDAN_RED)

f_cn = ImageFont.truetype(FONT, 96)
f_en = ImageFont.truetype(FONT, 36)
tx = cx + r + 70
cn_text = "复旦大学住房政策研究中心"
d.text((tx, cy - 118), cn_text, font=f_cn, fill=FUDAN_BLUE)
cn_w = d.textlength(cn_text, font=f_cn)
d.line([(tx + 4, cy + 24), (tx + cn_w - 4, cy + 24)], fill=FUDAN_RED, width=4)
d.text((tx, cy + 48), "Housing Policy Research Center, Fudan University",
       font=f_en, fill=(90, 96, 108))

out = os.path.join(OUT_DIR, "logo_fudan_hprc.png")
img.save(out)
print("logo saved:", out)
