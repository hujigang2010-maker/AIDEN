#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""将 PPTX 栅格化为 PNG，便于版式走查。支持形状填色、文本与嵌入图片。"""

from __future__ import annotations

import io
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_AUTO_SHAPE_TYPE
from pptx.util import Inches

FONT_PATH = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
SCALE = 1920 / 13.333


def emu_px(emu) -> int:
    return int(round(int(emu) / 914400 * SCALE))


def rgb_of(color):
    if color is None:
        return None
    try:
        c = color.rgb
        return (int(c[0]), int(c[1]), int(c[2]))
    except Exception:
        return None


def fill_rgb(shape):
    try:
        fill = shape.fill
        if fill is None or fill.type is None:
            return None
        return rgb_of(fill.fore_color)
    except Exception:
        return None


def font(size_pt: float):
    px = max(10, int(size_pt * SCALE / 72))
    try:
        return ImageFont.truetype(FONT_PATH, px)
    except Exception:
        return ImageFont.load_default()


def wrap(draw, text, fnt, max_w):
    if not text:
        return [""]
    lines = []
    for para in text.split("\n"):
        cur = ""
        for ch in para:
            trial = cur + ch
            if draw.textlength(trial, font=fnt) <= max_w or not cur:
                cur = trial
            else:
                lines.append(cur)
                cur = ch
        lines.append(cur)
    return lines or [""]


def draw_text_frame(img, shape):
    draw = ImageDraw.Draw(img)
    x, y = emu_px(shape.left), emu_px(shape.top)
    w, h = max(1, emu_px(shape.width)), max(1, emu_px(shape.height))
    tf = shape.text_frame
    blocks = []
    total_h = 0
    for p in tf.paragraphs:
        parts = []
        line_h = 18
        if p.runs:
            for r in p.runs:
                size = r.font.size.pt if r.font.size else 14
                fill = (42, 36, 53)
                try:
                    rgb = r.font.color.rgb
                    fill = (int(rgb[0]), int(rgb[1]), int(rgb[2]))
                except Exception:
                    pass
                fnt = font(size)
                parts.append((r.text or "", fnt, fill, size))
                line_h = max(line_h, int(size * SCALE / 72 * 1.18))
        else:
            text = p.text or ""
            if not text:
                continue
            fnt = font(14)
            parts.append((text, fnt, (42, 36, 53), 14))
            line_h = int(14 * SCALE / 72 * 1.18)
        if not parts:
            continue
        text = "".join(t for t, *_ in parts)
        fnt0 = parts[0][1]
        wrapped = wrap(draw, text, fnt0, w - 8)
        align = str(p.alignment) if p.alignment else "LEFT"
        blocks.append((wrapped, parts, align, line_h))
        total_h += line_h * len(wrapped)
    cy = y + 2
    try:
        va = str(tf.vertical_anchor)
        if "MIDDLE" in va:
            cy = y + max(0, (h - total_h) // 2)
        elif "BOTTOM" in va:
            cy = y + max(0, h - total_h)
    except Exception:
        pass
    for wrapped, parts, align, line_h in blocks:
        fill = parts[0][2]
        fnt = parts[0][1]
        for line in wrapped:
            tw = draw.textlength(line, font=fnt)
            if "RIGHT" in align:
                tx = x + w - tw - 4
            elif "CENTER" in align:
                tx = x + (w - tw) / 2
            else:
                tx = x + 4
            draw.text((tx, cy), line, font=fnt, fill=fill)
            cy += line_h


def render_slide(slide, path: Path, width_in=13.333, height_in=7.5):
    W, H = int(width_in * SCALE), int(height_in * SCALE)
    img = Image.new("RGB", (W, H), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    for shape in slide.shapes:
        x, y = emu_px(shape.left), emu_px(shape.top)
        w, h = max(1, emu_px(shape.width)), max(1, emu_px(shape.height))
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            blob = Image.open(io.BytesIO(shape.image.blob)).convert("RGBA")
            blob = blob.resize((w, h), Image.Resampling.LANCZOS)
            img.paste(blob, (x, y), blob)
            continue
        color = fill_rgb(shape)
        if color and w > 0 and h > 0:
            try:
                auto = shape.auto_shape_type
            except Exception:
                auto = None
            if auto in (MSO_AUTO_SHAPE_TYPE.OVAL, MSO_AUTO_SHAPE_TYPE.OVAL):
                draw.ellipse([x, y, x + w, y + h], fill=color)
            elif auto == MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE:
                r = max(6, min(w, h) // 8)
                draw.rounded_rectangle([x, y, x + w, y + h], radius=r, fill=color)
            else:
                draw.rectangle([x, y, x + w, y + h], fill=color)
        if shape.has_text_frame:
            draw_text_frame(img, shape)
    img.save(path, "PNG", optimize=True)


def main():
    pptx = Path(sys.argv[1] if len(sys.argv) > 1 else "deliverables/客户经理新员工培养阶段工作汇报.pptx")
    out = Path(sys.argv[2] if len(sys.argv) > 2 else "deliverables/preview")
    out.mkdir(parents=True, exist_ok=True)
    prs = Presentation(str(pptx))
    w_in = prs.slide_width / Inches(1)
    h_in = prs.slide_height / Inches(1)
    for i, slide in enumerate(prs.slides, 1):
        p = out / f"slide-{i:02d}.png"
        render_slide(slide, p, w_in, h_in)
        print("wrote", p)


if __name__ == "__main__":
    main()
