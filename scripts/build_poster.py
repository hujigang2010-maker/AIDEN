# -*- coding: utf-8 -*-
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

from content import EVENT

OUT = Path(__file__).resolve().parents[1] / "exports" / "飞书DemoDay4_转发海报.png"
FONT = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
W, H = 1080, 1440
NAVY = (11, 22, 44)
BLUE = (51, 112, 255)
CYAN = (44, 211, 200)
GOLD = (240, 192, 90)
WHITE = (247, 250, 255)
MUTED = (168, 184, 208)


def font(size):
    return ImageFont.truetype(FONT, size)


def wrap(draw, text, fnt, max_w):
    lines, cur = [], ""
    for ch in text:
        trial = cur + ch
        if draw.textlength(trial, font=fnt) <= max_w:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = ch
    if cur:
        lines.append(cur)
    return lines


def build():
    img = Image.new("RGB", (W, H), NAVY)
    d = ImageDraw.Draw(img)
    d.rectangle((0, 0, 16, H), fill=BLUE)
    d.rectangle((0, H - 18, W, H), fill=CYAN)
    d.rounded_rectangle((64, 72, 430, 128), 24, fill=BLUE)
    d.text((88, 86), "AI BUILDER  ·  #4", font=font(28), fill=WHITE)
    d.text((64, 170), "飞书 Demo Day", font=font(36), fill=CYAN)
    y = 230
    for line in wrap(d, "让 AI 接住你的日常小麻烦", font(64), 940):
        d.text((64, y), line, font=font(64), fill=WHITE)
        y += 78
    d.text((64, 400), "张咋啦  ·  向阳乔木  ·  Agent 入会", font=font(30), fill=GOLD)

    cards = [
        ("时间", "2026-07-02 周四  11:00–12:30"),
        ("形式", "线上直播  ·  飞书日历入会"),
        ("提醒", "上一场爆满，先加会议"),
    ]
    top = 500
    for title, body in cards:
        d.rounded_rectangle((64, top, 1016, top + 150), 28, fill=(22, 44, 82))
        d.text((96, top + 28), title, font=font(24), fill=CYAN)
        d.text((96, top + 72), body, font=font(32), fill=WHITE)
        top += 174

    d.text((64, 1040), "预约 / 入会", font=font(24), fill=MUTED)
    link = EVENT["share_url"]
    y = 1088
    for line in wrap(d, link, font(22), 940):
        d.text((64, y), line, font=font(22), fill=WHITE)
        y += 36
    d.text((64, 1280), "日历显示已结束也先点开：可能有回放或加场", font=font(24), fill=MUTED)
    d.text((64, 1336), "AIDEN 观摩材料  ·  非官方议程", font=font(22), fill=(90, 110, 140))
    OUT.parent.mkdir(parents=True, exist_ok=True)
    img.save(OUT, "PNG")
    print(f"wrote {OUT}")
    return OUT


if __name__ == "__main__":
    build()
