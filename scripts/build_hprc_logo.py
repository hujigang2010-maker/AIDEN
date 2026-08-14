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

FONT_CANDIDATES = [
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
]
FONT_PATH = next(p for p in FONT_CANDIDATES if os.path.exists(p))

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

f_cn = ImageFont.truetype(FONT_PATH, 96)
f_en = ImageFont.truetype(FONT_PATH, 36)
tx = cx + r + 70
cn_text = "复旦大学住房政策研究中心"
d.text((tx, cy - 118), cn_text, font=f_cn, fill=FUDAN_BLUE)
cn_w = d.textlength(cn_text, font=f_cn)
d.line([(tx + 4, cy + 22), (tx + cn_w - 4, cy + 22)], fill=FUDAN_RED, width=4)
d.text((tx, cy + 44), "Housing Policy Research Center, Fudan University",
       font=f_en, fill=(90, 96, 108))

out = "/workspace/whitepaper/assets/logo_fudan_hprc.png"
os.makedirs(os.path.dirname(out), exist_ok=True)
img.save(out)
print("logo saved", out)
