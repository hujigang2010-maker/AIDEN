#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 PPTX 栅格化为 PNG，便于检查中文版式。"""

from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

FONT_PATH = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
SCALE = 1920 / 13.333333


def emu_px(value) -> int:
    return int(round(int(value) / 914400 * SCALE))


def font(size_pt: float):
    px = max(8, int(size_pt * SCALE / 72))
    return ImageFont.truetype(FONT_PATH, px)


def rgb_of(color):
    try:
        c = color.rgb
        return (c[0], c[1], c[2])
    except Exception:
        return None


def fill_rgb(shape):
    try:
        fill = shape.fill
        if fill.type is None:
            return None
        return rgb_of(fill.fore_color)
    except Exception:
        return None


def wrap(draw, text, fnt, max_w):
    lines = []
    for para in text.split("\n"):
        cur = ""
        for ch in para:
            trial = cur + ch
            if not cur or draw.textlength(trial, font=fnt) <= max_w:
                cur = trial
            else:
                lines.append(cur)
                cur = ch
        lines.append(cur)
    return lines or [""]


def draw_shape(base, shape, draw):
    x = emu_px(shape.left)
    y = emu_px(shape.top)
    w = emu_px(shape.width)
    h = emu_px(shape.height)
    color = fill_rgb(shape)
    if color and w > 0 and h > 0:
        try:
            if shape.shape_type == MSO_SHAPE.ROUNDED_RECTANGLE or "ROUNDED" in str(shape.auto_shape_type):
                radius = max(4, int(min(w, h) * 0.08))
                draw.rounded_rectangle([x, y, x + w, y + h], radius=radius, fill=color)
            else:
                draw.rectangle([x, y, x + w, y + h], fill=color)
        except Exception:
            draw.rectangle([x, y, x + w, y + h], fill=color)
    tf = getattr(shape, "text_frame", None)
    if tf is None or not tf.text.strip():
        return
    try:
        anchor = tf.vertical_anchor
    except Exception:
        anchor = MSO_ANCHOR.TOP
    pad_l = emu_px(tf.margin_left)
    pad_r = emu_px(tf.margin_right)
    pad_t = emu_px(tf.margin_top)
    max_w = max(8, w - pad_l - pad_r)
    blocks = []
    for p in tf.paragraphs:
        size = 14
        color_t = (28, 25, 22)
        align = p.alignment or PP_ALIGN.LEFT
        text = "".join(run.text for run in p.runs) if p.runs else p.text
        if p.runs:
            if p.runs[0].font.size:
                size = p.runs[0].font.size.pt
            got = rgb_of(p.runs[0].font.color)
            if got:
                color_t = got
        fnt = font(size)
        lines = wrap(draw, text, fnt, max_w)
        mult = 1.15
        try:
            if isinstance(p.line_spacing, float):
                mult = p.line_spacing
        except Exception:
            pass
        line_h = int(size * mult * SCALE / 72 * 1.15)
        after = 0
        before = 0
        try:
            if p.space_after:
                after = int(p.space_after.pt * SCALE / 72)
            if p.space_before:
                before = int(p.space_before.pt * SCALE / 72)
        except Exception:
            pass
        blocks.append((lines, fnt, color_t, align, line_h, before, after))
    total_h = sum(before + len(lines) * line_h + after for lines, _, _, _, line_h, before, after in blocks)
    if anchor == MSO_ANCHOR.MIDDLE:
        cursor = y + max(pad_t, (h - total_h) // 2)
    else:
        cursor = y + pad_t
    for lines, fnt, color_t, align, line_h, before, after in blocks:
        cursor += before
        for line in lines:
            tw = draw.textlength(line, font=fnt) if line else 0
            if align == PP_ALIGN.CENTER:
                tx = x + pad_l + max(0, (max_w - tw) / 2)
            elif align == PP_ALIGN.RIGHT:
                tx = x + w - pad_r - tw
            else:
                tx = x + pad_l
            draw.text((tx, cursor), line, font=fnt, fill=color_t)
            cursor += line_h
        cursor += after


def render(pptx: Path, out_dir: Path) -> None:
    prs = Presentation(str(pptx))
    out_dir.mkdir(parents=True, exist_ok=True)
    width = emu_px(prs.slide_width)
    height = emu_px(prs.slide_height)
    for i, slide in enumerate(prs.slides, start=1):
        image = Image.new("RGB", (width, height), (246, 241, 232))
        draw = ImageDraw.Draw(image)
        for shape in slide.shapes:
            draw_shape(image, shape, draw)
        dest = out_dir / f"{i:02d}.png"
        image.save(dest)
        print(dest)


def main() -> None:
    if len(sys.argv) < 3:
        raise SystemExit("用法: render_pptx_preview.py 输入.pptx 输出目录")
    render(Path(sys.argv[1]), Path(sys.argv[2]))


if __name__ == "__main__":
    main()
