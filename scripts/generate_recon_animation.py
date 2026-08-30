#!/usr/bin/env python3
"""按交警监控口述生成示范示意图动画（非写实 3D，不替代监控）。

状态：待用。build_all.py 不调用本脚本；本次不重做视频。
事实：护栏开口先借道右靠、再大弧度左拐、转向灯不亮；后方美团二轮车距不足；碰撞侧翻。
禁止编造：具体钟点、车速、颜色品牌、撞击部位、任何一方全责。
"""

from __future__ import annotations

import math
import subprocess
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
FRAMES_DIR = Path("/tmp/recon_frames")
OUT = ROOT / "deliverables" / "事故3D复原_示范动画.mp4"

W, H = 1280, 720
FPS = 24
# 片段(秒): 建立, 掉头, 后方接近, 碰撞, 侧翻后果, 俯视
SECONDS = [4, 6, 4, 3, 3, 4]
TOTAL_S = sum(SECONDS)
N = TOTAL_S * FPS

FONT_CJK = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
FONT_BOLD = FONT_CJK


def font(size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_CJK, size)


def road(draw: ImageDraw.ImageDraw) -> None:
    draw.rectangle([0, 0, W, H], fill="#7d7566")
    draw.rectangle([0, 160, W, 300], fill="#4a4a4a")
    draw.rectangle([0, 420, W, 560], fill="#4a4a4a")
    for x in range(-40, W + 40, 120):
        draw.line([(x, 230), (x + 70, 230)], fill="#d9d0c2", width=6)
        draw.line([(x, 490), (x + 70, 490)], fill="#d9d0c2", width=6)
    draw.rectangle([0, 300, W, 420], fill="#6b5d49")
    draw.rectangle([40, 250, 460, 420], fill="#a9946c")
    draw.rectangle([40, 250, 460, 272], fill="#7a5b3a")
    draw.text((60, 282), "抚顺路批发市场 卸货区", font=font(26), fill="#241f18")


def tri_state(t: float):
    """掉头进度 t 0..1。先右后大弧度左。"""
    cx, cy = 330, 470
    rx = cx + 230 * t
    y = cy
    if t < 0.35:
        y = cy + 52 * (t / 0.35)
    else:
        y = cy + 52 - 145 * ((t - 0.35) / 0.65)
    ang = 180 * max(0.0, (t - 0.35) / 0.65)
    return rx, y, ang


def draw_vehicle(draw, x, y, ang, kind, size, color):
    a = math.radians(-ang)
    ca, sa = math.cos(a), math.sin(a)
    hw = size[0] / 2
    hh = size[1] / 2
    pts = [(-hw, -hh), (hw, -hh), (hw, hh), (-hw, hh)]
    world = []
    for px, py in pts:
        wx = x + px * ca - py * sa
        wy = y + px * sa + py * ca
        world.append((wx, wy))
    draw.polygon(world, fill=color, outline="#2a241c")
    hx = x + hw * ca * 0.72
    hy = y + hw * sa * 0.72
    draw.ellipse([hx - 5, hy - 5, hx + 5, hy + 5], fill="#f3e6c8")


def draw_hu_top(draw, x, y, r=13):
    draw.ellipse([x - r, y - r, x + r, y + r], fill="#26343a", outline="#0f1417")


def draw_rider_top(draw, x, y, r=13):
    draw.ellipse([x - r, y - r, x + r, y + r], fill="#e6b422", outline="#7a5b10")


def subtitle_bar(draw, text: str, tone: str = "dark") -> None:
    draw.rectangle([0, H - 78, W, H], fill="#241f18")
    f = font(30)
    bbox = draw.textbbox((0, 0), text, font=f)
    tw = bbox[2] - bbox[0]
    draw.text(((W - tw) / 2, H - 60), text, font=f, fill="#f4efe4")


def draw_frame(i: int) -> Image.Image:
    t_total = i / N
    acc = 0.0
    seg_idx = 0
    for s_i, s in enumerate(SECONDS):
        acc += s / TOTAL_S
        if t_total <= acc:
            seg_idx = s_i
            break
    seg_start = sum(SECONDS[:seg_idx]) / TOTAL_S
    seg_len = SECONDS[seg_idx] / TOTAL_S
    t = min(1.0, max(0.0, (t_total - seg_start) / seg_len))

    img = Image.new("RGB", (W, H), "#8d8577")
    d = ImageDraw.Draw(img)

    if seg_idx <= 4:
        road(d)
        f26 = font(26)
        d.text((40, 60), "青岛 · 市北区抚顺路批发市场外 · 市政道路", font=f26, fill="#241f18")
        d.text((40, 96), "示意图（非监控原始视频）", font=f26, fill="#7a3a20")

        tri_t = 0.0
        rider_x = -120
        if seg_idx >= 1:
            tri_t = min(1.0, (i / FPS - 4) / 6) if seg_idx == 1 else 1.0
        rx, ry, rang = tri_state(tri_t)

        if seg_idx >= 2:
            rt = t if seg_idx == 2 else 1.0
            rider_x = -40 + (700 + 40) * rt * 0.55
            if seg_idx >= 3:
                rt2 = t if seg_idx == 3 else 1.0
                rider_x = 700 - 40 - 340 * rt2
        rider_y = 500

        if seg_idx == 3:
            t_slow = t
            rider_x = 700 - 40 - 340 * t_slow
            if rider_x <= rx + 40:
                rider_x = rx + 40
                d.text((rx + 90, ry - 40), "碰撞", font=font(34), fill="#a63d2f")

        if seg_idx >= 4:
            rx, ry, _ = tri_state(1.0)
            d.polygon([(rx - 70, ry - 26), (rx + 46, ry - 60), (rx + 46, ry + 60)], fill="#6f6253", outline="#2a241c")
            d.text((rx - 40, ry + 74), "三轮侧翻", font=font(30), fill="#7a3a20")
            draw_hu_top(d, rx - 44, ry + 40, r=15)
            d.text((rx - 118, ry + 34), "胡某右小腿/右踝受伤，不能站起", font=font(22), fill="#241f18")
            draw_rider_top(d, rider_x - 40, rider_y, r=13)
            d.rectangle([rider_x - 70, rider_y - 16, rider_x + 20, rider_y + 16], fill="#e6b422", outline="#7a5b10")
            d.text((rider_x - 88, rider_y - 48), "骑手停车", font=font(22), fill="#241f18")
        else:
            draw_vehicle(d, rx, ry, rang, "tri", (110, 64), "#6f6253")
            draw_hu_top(d, rx - 8, ry, r=13)
            # 转向灯：始终不亮
            lt = rx + 52 * math.cos(math.radians(-rang)) + 30
            ly = ry + 52 * math.sin(math.radians(-rang))
            d.ellipse([lt - 7, ly - 7, lt + 7, ly + 7], fill="#3a3328", outline="#6b5d49")
            d.text((lt + 12, ly - 14), "转向灯不亮", font=font(18), fill="#241f18")
            d.rectangle([rx + 18, ry - 12, rx + 62, ry + 12], fill="#8b7a5c")
            d.text((rx + 20, ry - 10), "市场牌 2588", font=font(12), fill="#f4efe4")

            if rider_x > -100:
                draw_vehicle(d, rider_x, rider_y, 0, "e2", (92, 34), "#e6b422")
                draw_rider_top(d, rider_x, rider_y, r=13)
                d.rectangle([rider_x - 14, rider_y - 24, rider_x + 14, rider_y - 8], fill="#f2d15c", outline="#7a5b10")
                d.text((rider_x + 20, rider_y - 44), "美团二轮（后方）", font=font(20), fill="#241f18")
                if seg_idx >= 1:
                    d.line([(rider_x, rider_y - 30), (rx, ry - 20)], fill="#a63d2f", width=3)
                    midx = (rider_x + rx) / 2
                    midy = (rider_y - 30 + ry - 20) / 2
                    d.text((midx - 40, midy - 8), "车距不足", font=font(20), fill="#a63d2f")

        d.text((40, 306), "刘孝春（车主/雇主）", font=font(20), fill="#241f18")
        d.ellipse([92, 336, 116, 360], fill="#5a4a38", outline="#241f18")

        subtitles = [
            "根据交警已调监控的口述口径还原，非原始视频",
            "掉头：先向右靠，再大弧度左拐；转向灯全程不亮",
            "后方美团二轮：跟车距离不足",
            "碰撞（慢动作占位，精确部位以监控为准）",
            "三轮侧翻。交警排除任何一方全责：可能同等，或外卖主责、三轮次责",
        ]
        if seg_idx < len(subtitles):
            subtitle_bar(d, subtitles[seg_idx])
    else:
        d.rectangle([0, 0, W, H], fill="#3f4a56")
        d.text((60, 70), "俯视路线示意", font=font(40), fill="#f4efe4")
        d.line([(180, 500), (860, 500)], fill="#d9d0c2", width=70)
        d.line([(300, 560), (430, 596)], fill="#3f7fb0", width=22)
        d.polygon([(448, 606), (400, 606), (430, 570)], fill="#3f7fb0")
        d.line([(460, 584), (650, 420)], fill="#3f7fb0", width=22)
        d.polygon([(668, 408), (620, 438), (676, 442)], fill="#3f7fb0")
        d.text((210, 600), "三轮：先右靠 → 大弧度左拐", font=font(30), fill="#bcd7ea")
        d.line([(700, 620), (620, 468)], fill="#e6b422", width=22)
        d.polygon([(612, 452), (590, 496), (648, 486)], fill="#e6b422")
        d.text((650, 636), "美团二轮（后方）", font=font(30), fill="#f2d15c")
        d.ellipse([592, 432, 616, 456], fill="#a63d2f")
        d.text((632, 436), "碰撞点（占位）", font=font(26), fill="#f4efe4")
        d.text((60, 660), "监控排除任何一方全责。以交警案卷和原视频为准。", font=font(24), fill="#f4efe4")
        subtitle_bar(d, "示意图结束 · 非证据 · 有原视频后按原视频校正")

    return img


def main() -> None:
    FRAMES_DIR.mkdir(parents=True, exist_ok=True)
    for f in FRAMES_DIR.glob("f_*.png"):
        f.unlink()
    for i in range(N):
        draw_frame(i).save(FRAMES_DIR / f"f_{i:05d}.png")
    OUT.parent.mkdir(parents=True, exist_ok=True)
    cmd = [
        "ffmpeg", "-y", "-framerate", str(FPS), "-i", str(FRAMES_DIR / "f_%05d.png"),
        "-c:v", "libx264", "-pix_fmt", "yuv420p", "-crf", "21", str(OUT),
    ]
    subprocess.run(cmd, check=True, capture_output=True)
    print("wrote", OUT, OUT.stat().st_size, "bytes", TOTAL_S, "s")


if __name__ == "__main__":
    main()
