# -*- coding: utf-8 -*-
"""将 PPTX 栅格化为 PNG，便于版式走查。"""
from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu

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


def font(size_pt: float, bold: bool = False):
    px = max(10, int(size_pt * SCALE / 72 * (1.05 if bold else 1)))
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
    w = max(1, emu_px(shape.width))
    tf = shape.text_frame
    cy = y + 8
    for p in tf.paragraphs:
        text = "".join(r.text or "" for r in p.runs) if p.runs else (p.text or "")
        if not text:
            cy += 10
            continue
        size = 14
        bold = False
        color = (26, 36, 51)
        if p.runs:
            r0 = p.runs[0]
            if r0.font.size:
                size = r0.font.size.pt
            bold = bool(r0.font.bold)
            try:
                if r0.font.color and r0.font.color.rgb:
                    c = r0.font.color.rgb
                    color = (int(c[0]), int(c[1]), int(c[2]))
            except Exception:
                pass
        fnt = font(size, bold)
        align = str(p.alignment) if p.alignment else ""
        lines = wrap(draw, text, fnt, w - 16)
        for line in lines:
            tw = draw.textlength(line, font=fnt)
            if "CENTER" in align:
                lx = x + max(0, (w - tw) / 2)
            elif "RIGHT" in align:
                lx = x + max(0, w - tw - 8)
            else:
                lx = x + 8
            draw.text((lx, cy), line, font=fnt, fill=color)
            cy += int(size * SCALE / 72 * 1.25)


def render_slide(slide, width, height) -> Image.Image:
    img = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    shapes = sorted(slide.shapes, key=lambda s: (s.left or 0, s.top or 0))
    for shape in shapes:
        x, y = emu_px(shape.left or 0), emu_px(shape.top or 0)
        w, h = max(1, emu_px(shape.width or 1)), max(1, emu_px(shape.height or 1))
        if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            table = shape.table
            rows, cols = len(table.rows), len(table.columns)
            cell_h = h / max(rows, 1)
            # 列宽按 EMU
            col_w = []
            for col in table.columns:
                col_w.append(emu_px(col.width))
            if sum(col_w) <= 0:
                col_w = [w // cols] * cols
            scale = w / max(sum(col_w), 1)
            col_w = [max(1, int(cw * scale)) for cw in col_w]
            yy = y
            for ri, row in enumerate(table.rows):
                xx = x
                for ci, cell in enumerate(row.cells):
                    cw = col_w[ci] if ci < len(col_w) else w // cols
                    bg = (7, 26, 43) if ri == 0 else ((255, 255, 255) if ri % 2 else (238, 243, 246))
                    draw.rectangle([xx, yy, xx + cw, yy + cell_h], fill=bg, outline=(213, 221, 229))
                    fnt = font(9 if ri else 10, bold=(ri == 0))
                    color = (255, 255, 255) if ri == 0 else (26, 36, 51)
                    lines = wrap(draw, cell.text or "", fnt, cw - 10)
                    ty = yy + 4
                    for line in lines[:6]:
                        draw.text((xx + 5, ty), line, font=fnt, fill=color)
                        ty += 16
                    xx += cw
                yy += cell_h
            continue
        rgb = fill_rgb(shape)
        if rgb:
            draw.rectangle([x, y, x + w, y + h], fill=rgb)
        if shape.has_text_frame:
            draw_text_frame(img, shape)
    return img


def main():
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("deliverables") / "东昇聚变_政府事务与上海产业落地_90天工作设想.pptx"
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("preview")
    out.mkdir(parents=True, exist_ok=True)
    prs = Presentation(src)
    w, h = emu_px(prs.slide_width), emu_px(prs.slide_height)
    paths = []
    for i, slide in enumerate(prs.slides, 1):
        img = render_slide(slide, w, h)
        p = out / f"slide_{i:02d}.png"
        img.save(p, "PNG")
        paths.append(p)
        print(f"wrote {p}")
    print(f"共 {len(paths)} 页 -> {out}")


if __name__ == "__main__":
    main()
