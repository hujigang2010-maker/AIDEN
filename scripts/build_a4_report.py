# -*- coding: utf-8 -*-
"""将白皮书排成 A4（210×297mm）印刷 PDF，页眉使用 RCHP 正式组合标。"""
import os
import re
from io import BytesIO

from PIL import Image, ImageDraw, ImageFont, ImageEnhance, ImageFilter

BASE = "/workspace/whitepaper"
MD = os.path.join(BASE, "WAIC2026人工智能产业空间白皮书.md")
PHOTO = os.path.join(BASE, "assets/photos")
CHART = os.path.join(BASE, "assets/charts")
BRAND = os.path.join(BASE, "assets/brand")
PAGE_DIR = os.path.join(BASE, "assets/a4_pages")
OUT_PDF = os.path.join(BASE, "WAIC2026人工智能产业空间白皮书-A4.pdf")
DL = os.path.join(BASE, "下载版本")
COVER_PNG = os.path.join(BRAND, "cover_a4.png")

# 180 dpi A4
DPI = 180
W, H = 1488, 2102  # 210mm × 297mm
MM = DPI / 25.4

NAVY = (10, 42, 82)
BLUE = (14, 78, 155)
RED = (196, 18, 32)
GOLD = (201, 162, 79)
WHITE = (255, 255, 255)
INK = (28, 33, 44)
MUTED = (92, 98, 110)
RULE = (220, 224, 230)
LIGHT = (246, 248, 251)
QUOTE_BG = (241, 245, 250)
TABLE_HEAD = (14, 78, 155)

FONT_SANS_R = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
FONT_SANS_B = "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"
FONT_SERIF_R = "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc"
FONT_SERIF_B = "/usr/share/fonts/opentype/noto/NotoSerifCJK-Bold.ttc"
FONT_LATIN_R = "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf"

MARGIN_L = int(18 * MM)
MARGIN_R = int(18 * MM)
HEADER_H = int(22 * MM)
FOOTER_Y = H - int(14 * MM)
CONTENT_TOP = int(26 * MM)
CONTENT_BOTTOM = H - int(18 * MM)
CONTENT_W = W - MARGIN_L - MARGIN_R

os.makedirs(PAGE_DIR, exist_ok=True)
os.makedirs(DL, exist_ok=True)


def font(path, size):
    return ImageFont.truetype(path, size)


def cover_fill(path, size, focus="center"):
    im = Image.open(path).convert("RGB")
    tw, th = size
    scale = max(tw / im.width, th / im.height)
    im = im.resize((int(im.width * scale), int(im.height * scale)), Image.Resampling.LANCZOS)
    x = (im.width - tw) // 2
    if focus == "top":
        y = 0
    elif focus == "bottom":
        y = im.height - th
    else:
        y = (im.height - th) // 2
    return im.crop((x, y, x + tw, y + th))


def darken(im, factor=0.55):
    return ImageEnhance.Brightness(im).enhance(factor)


def fit_logo(path, max_w, max_h):
    im = Image.open(path).convert("RGBA")
    scale = min(max_w / im.width, max_h / im.height)
    return im.resize((int(im.width * scale), int(im.height * scale)), Image.Resampling.LANCZOS)


def wrap_text(draw, text, fnt, max_w):
    lines = []
    for para in text.split("\n"):
        if not para:
            lines.append("")
            continue
        line = ""
        for ch in para:
            trial = line + ch
            if draw.textlength(trial, font=fnt) <= max_w:
                line = trial
            else:
                if line:
                    lines.append(line)
                line = ch
        if line:
            lines.append(line)
    return lines


def parse_runs(text):
    """将 **bold** / *italic* 拆成 (text, bold) 片段。"""
    runs = []
    for seg in re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", text):
        if not seg:
            continue
        if seg.startswith("**") and seg.endswith("**"):
            runs.append((seg[2:-2], True))
        elif seg.startswith("*") and seg.endswith("*") and len(seg) > 2:
            runs.append((seg[1:-1], False))
        else:
            runs.append((seg, False))
    return runs


def wrap_runs(draw, runs, fnt_r, fnt_b, max_w):
    lines = []
    cur = []
    width = 0
    for text, bold in runs:
        fnt = fnt_b if bold else fnt_r
        buf = ""
        for ch in text:
            tw = draw.textlength(ch, font=fnt)
            if width + tw > max_w and (buf or cur):
                if buf:
                    cur.append((buf, bold))
                lines.append(cur)
                cur = []
                buf = ch
                width = tw
            else:
                buf += ch
                width += tw
        if buf:
            cur.append((buf, bold))
    if cur:
        lines.append(cur)
    return lines


class Report:
    def __init__(self):
        self.pages = []
        self.page = None
        self.draw = None
        self.y = 0
        self.f_body = font(FONT_SERIF_R, 28)
        self.f_body_b = font(FONT_SERIF_B, 28)
        self.f_h1 = font(FONT_SANS_B, 42)
        self.f_h2 = font(FONT_SANS_B, 30)
        self.f_h3 = font(FONT_SANS_B, 24)
        self.f_small = font(FONT_SANS_R, 16)
        self.f_caption = font(FONT_SANS_R, 16)
        self.f_quote = font(FONT_SERIF_R, 24)
        self.f_quote_b = font(FONT_SERIF_B, 24)
        self.f_table = font(FONT_SANS_R, 17)
        self.f_table_b = font(FONT_SANS_B, 17)
        self.f_header = font(FONT_SANS_R, 14)
        self.f_page = font(FONT_SANS_R, 14)
        self.logo_header = fit_logo(os.path.join(BRAND, "logo_rchp_mark.png"), int(58 * MM), int(15 * MM))
        self.logo_print = fit_logo(os.path.join(BRAND, "logo_rchp_print.png"), int(148 * MM), int(58 * MM))
        self.logo_dark = fit_logo(os.path.join(BRAND, "logo_rchp_dark.png"), int(152 * MM), int(60 * MM))
        self.line_h = 46
        self.first_h1_skipped = False

    def new_page(self, header=True):
        self.flush()
        self.page = Image.new("RGB", (W, H), WHITE)
        self.draw = ImageDraw.Draw(self.page)
        self.y = CONTENT_TOP
        if header:
            self._header()
            self.y = CONTENT_TOP + 8
        else:
            self.y = MARGIN_L

    def flush(self):
        if self.page is not None:
            self._footer()
            self.pages.append(self.page)
            self.page = None

    def _header(self):
        self.page.paste(self.logo_header, (MARGIN_L, int(7 * MM)), self.logo_header)
        self.draw.text((W - MARGIN_R, int(12 * MM)),
                       "中心研究文稿 · 第二号    FDU-HPRC-WP-2026-02",
                       font=self.f_header, fill=MUTED, anchor="rt")
        y = int(20 * MM)
        self.draw.line([(MARGIN_L, y), (W - MARGIN_R, y)], fill=RED, width=2)
        self.draw.line([(MARGIN_L, y + 3), (W - MARGIN_R, y + 3)], fill=NAVY, width=1)

    def _footer(self):
        if not self.pages and self.page is not None:
            # 封面无页码条也可画，由调用方决定；正文页统一页脚
            pass
        n = len(self.pages) + 1
        self.draw.line([(MARGIN_L, FOOTER_Y - 10), (W - MARGIN_R, FOOTER_Y - 10)],
                       fill=RULE, width=1)
        self.draw.text((MARGIN_L, FOOTER_Y),
                       "复旦大学住房政策研究中心  ·  WAIC2026 人工智能产业空间白皮书",
                       font=self.f_page, fill=MUTED)
        self.draw.text((W - MARGIN_R, FOOTER_Y), f"{n}",
                       font=self.f_page, fill=NAVY, anchor="rt")

    def ensure(self, h):
        if self.page is None or self.y + h > CONTENT_BOTTOM:
            self.new_page()

    def gap(self, h):
        if self.page is None:
            self.new_page()
        if self.y + h > CONTENT_BOTTOM:
            self.new_page()
        else:
            self.y += h

    def add_h1(self, title):
        self.new_page()
        # 章节条
        self.draw.rectangle([MARGIN_L, self.y, MARGIN_L + 8, self.y + 52], fill=RED)
        lines = wrap_text(self.draw, title, self.f_h1, CONTENT_W - 24)
        yy = self.y
        for ln in lines:
            self.draw.text((MARGIN_L + 20, yy), ln, font=self.f_h1, fill=NAVY)
            yy += 44
        self.y = yy + 10
        self.draw.line([(MARGIN_L, self.y), (W - MARGIN_R, self.y)], fill=GOLD, width=2)
        self.y += 18

    def add_h2(self, title):
        h = 44
        self.ensure(h + 12)
        self.y += 8
        self.draw.text((MARGIN_L, self.y), title, font=self.f_h2, fill=BLUE)
        self.y += 36

    def add_h3(self, title):
        self.ensure(36)
        self.y += 6
        self.draw.text((MARGIN_L, self.y), title, font=self.f_h3, fill=NAVY)
        self.y += 30

    def add_para(self, text, indent=True):
        if self.page is None:
            self.new_page()
        runs = parse_runs(text)
        indent_w = self.draw.textlength("　　", font=self.f_body) if indent else 0
        lines = self._wrap_indent(runs, indent_w, CONTENT_W) if indent else wrap_runs(
            self.draw, runs, self.f_body, self.f_body_b, CONTENT_W)
        for i, line in enumerate(lines):
            self.ensure(self.line_h)
            x = MARGIN_L + (indent_w if i == 0 and indent else 0)
            for frag, bold in line:
                fnt = self.f_body_b if bold else self.f_body
                self.draw.text((x, self.y), frag, font=fnt, fill=INK)
                x += self.draw.textlength(frag, font=fnt)
            self.y += self.line_h
        self.y += 6

    def _wrap_indent(self, runs, indent_w, max_w):
        lines = []
        cur = []
        width = indent_w
        for text, bold in runs:
            fnt = self.f_body_b if bold else self.f_body
            buf = ""
            for ch in text:
                tw = self.draw.textlength(ch, font=fnt)
                if width + tw > max_w and (buf or cur):
                    if buf:
                        cur.append((buf, bold))
                    lines.append(cur)
                    cur = []
                    buf = ch
                    width = tw
                else:
                    buf += ch
                    width += tw
            if buf:
                cur.append((buf, bold))
        if cur:
            lines.append(cur)
        return lines

    def add_quote(self, text):
        runs = parse_runs(text)
        lines = wrap_runs(self.draw, runs, self.f_quote, self.f_quote_b, CONTENT_W - 36)
        box_h = len(lines) * 38 + 28
        self.ensure(box_h + 8)
        self.draw.rectangle([MARGIN_L, self.y, W - MARGIN_R, self.y + box_h], fill=QUOTE_BG)
        self.draw.rectangle([MARGIN_L, self.y, MARGIN_L + 7, self.y + box_h], fill=BLUE)
        yy = self.y + 14
        for line in lines:
            x = MARGIN_L + 22
            for frag, bold in line:
                fnt = self.f_quote_b if bold else self.f_quote
                self.draw.text((x, yy), frag, font=fnt, fill=BLUE)
                x += self.draw.textlength(frag, font=fnt)
            yy += 38
        self.y += box_h + 16

    def add_bullets(self, items, ordered=False):
        for i, item in enumerate(items, 1):
            mark = f"{i}. " if ordered else "• "
            runs = parse_runs(item)
            prefix_w = self.draw.textlength(mark, font=self.f_body_b)
            lines = wrap_runs(self.draw, runs, self.f_body, self.f_body_b, CONTENT_W - prefix_w)
            need = max(1, len(lines)) * self.line_h + 4
            self.ensure(need)
            self.draw.text((MARGIN_L, self.y), mark, font=self.f_body_b,
                           fill=RED if not ordered else NAVY)
            yy = self.y
            for li, line in enumerate(lines):
                x = MARGIN_L + prefix_w
                for frag, bold in line:
                    fnt = self.f_body_b if bold else self.f_body
                    self.draw.text((x, yy), frag, font=fnt, fill=INK)
                    x += self.draw.textlength(frag, font=fnt)
                yy += self.line_h
            self.y = yy + 4
        self.y += 6

    def add_image(self, path, caption=""):
        if not os.path.exists(path):
            self.add_para(f"[缺图：{path}]", indent=False)
            return
        im = Image.open(path).convert("RGBA")
        max_w = CONTENT_W
        max_h = int(118 * MM)
        scale = min(max_w / im.width, max_h / im.height, 1.0)
        im = im.resize((int(im.width * scale), int(im.height * scale)), Image.Resampling.LANCZOS)
        cap_h = 24 if caption else 0
        self.ensure(im.height + cap_h + 20)
        x = MARGIN_L + (CONTENT_W - im.width) // 2
        bg = Image.new("RGB", (im.width, im.height), WHITE)
        bg.paste(im, (0, 0), im)
        self.page.paste(bg, (x, self.y))
        self.y += im.height + 6
        if caption:
            self.draw.text((W // 2, self.y), caption, font=self.f_caption, fill=MUTED, anchor="mt")
            self.y += 22
        self.y += 10

    def add_table(self, rows):
        if not rows:
            return
        ncol = max(len(r) for r in rows)
        rows = [r + [""] * (ncol - len(r)) for r in rows]
        col_w = CONTENT_W / ncol
        cell_f = self.f_table
        # 预计算每行高度
        heights = []
        wrapped = []
        for ri, row in enumerate(rows):
            fnt = self.f_table_b if ri == 0 else cell_f
            row_lines = []
            max_h = 0
            for cell in row:
                cell = re.sub(r"\*\*([^*]+)\*\*", r"\1", cell)
                ls = wrap_text(self.draw, cell, fnt, col_w - 16)
                row_lines.append(ls)
                max_h = max(max_h, max(26, len(ls) * 22 + 14))
            wrapped.append(row_lines)
            heights.append(max_h)
        total_h = sum(heights) + 2
        self.ensure(min(total_h, CONTENT_BOTTOM - CONTENT_TOP))
        # 若表过高则逐行分页
        for ri, row in enumerate(rows):
            rh = heights[ri]
            self.ensure(rh)
            y0 = self.y
            if ri == 0:
                self.draw.rectangle([MARGIN_L, y0, W - MARGIN_R, y0 + rh], fill=BLUE)
            elif ri % 2 == 1:
                self.draw.rectangle([MARGIN_L, y0, W - MARGIN_R, y0 + rh], fill=LIGHT)
            for ci in range(ncol + 1):
                x = MARGIN_L + int(ci * col_w)
                self.draw.line([(x, y0), (x, y0 + rh)], fill=RULE if ri else BLUE, width=1)
            self.draw.line([(MARGIN_L, y0), (W - MARGIN_R, y0)], fill=RULE, width=1)
            self.draw.line([(MARGIN_L, y0 + rh), (W - MARGIN_R, y0 + rh)], fill=RULE, width=1)
            fill = WHITE if ri == 0 else INK
            fnt = self.f_table_b if ri == 0 else cell_f
            for ci, ls in enumerate(wrapped[ri]):
                x = MARGIN_L + int(ci * col_w) + 8
                yy = y0 + 6
                for ln in ls:
                    self.draw.text((x, yy), ln, font=fnt, fill=fill)
                    yy += 22
            self.y += rh
        self.y += 14

    def make_cover(self):
        im = darken(cover_fill(os.path.join(PHOTO, "photo_cover_lujiazui.png"), (W, H)), 0.48)
        overlay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        od = ImageDraw.Draw(overlay)
        for i in range(H):
            a = int(90 + 90 * (i / H))
            od.line([(0, i), (W, i)], fill=(6, 16, 36, min(a, 160)))
        im = im.convert("RGBA")
        im.alpha_composite(overlay)
        im = im.convert("RGB")
        d = ImageDraw.Draw(im)
        d.rectangle([0, 0, int(6 * MM), H], fill=GOLD)
        logo = self.logo_dark
        im.paste(logo, (MARGIN_L, int(22 * MM)), logo)
        d.rectangle([MARGIN_L, int(118 * MM), MARGIN_L + int(22 * MM), int(118 * MM) + 4], fill=GOLD)
        d.text((MARGIN_L, int(128 * MM)), "中心研究文稿 · 第二号", font=font(FONT_SANS_R, 22), fill=WHITE)
        d.text((MARGIN_L, int(148 * MM)), "FDU-HPRC-WP-2026-02", font=font(FONT_SANS_R, 18), fill=(210, 214, 220))
        d.text((MARGIN_L, int(175 * MM)), "WAIC2026", font=font(FONT_SANS_B, 72), fill=WHITE)
        title = "人工智能产业空间白皮书"
        d.text((MARGIN_L, int(200 * MM)), title, font=font(FONT_SANS_B, 48), fill=WHITE)
        d.text((MARGIN_L, int(220 * MM)),
               "AI 与产业空间融合的趋势、格局与新范式",
               font=font(FONT_SERIF_R, 24), fill=(220, 224, 230))
        d.line([(MARGIN_L, int(232 * MM)), (MARGIN_L + int(48 * MM), int(232 * MM))], fill=RED, width=3)
        d.text((MARGIN_L, H - int(28 * MM)),
               "复旦大学住房政策研究中心    二〇二六年八月 · 上海",
               font=font(FONT_SANS_R, 18), fill=(210, 214, 220))
        im.save(COVER_PNG, "PNG")
        self.pages.append(im)

    def make_inside_cover(self):
        self.new_page(header=True)
        self.y = int(40 * MM)
        self.page.paste(self.logo_print,
                        (MARGIN_L + (CONTENT_W - self.logo_print.width) // 2, self.y),
                        self.logo_print)
        self.y += self.logo_print.height + 36
        for txt, fnt, fill, sp in [
            ("中心研究文稿 · 第二号", font(FONT_SANS_B, 22), BLUE, 32),
            ("WAIC2026 人工智能产业空间白皮书", font(FONT_SANS_B, 28), NAVY, 40),
            ("文稿编号  FDU-HPRC-WP-2026-02", font(FONT_SANS_R, 18), MUTED, 30),
            ("编制单位  复旦大学住房政策研究中心", font(FONT_SANS_R, 18), INK, 28),
            ("Housing Policy Research Center, Fudan University", font(FONT_LATIN_R, 16), MUTED, 28),
            ("成稿日期  2026 年 8 月  ·  上海", font(FONT_SANS_R, 18), MUTED, 28),
        ]:
            self.draw.text((W // 2, self.y), txt, font=fnt, fill=fill, anchor="mt")
            self.y += sp
        self.y += 24
        note = ("本文稿仅供学术研究与政策参考。数据来源于 WAIC 2026 全量资源整合总表、"
                "XSCT Bench（xsct.ai）、观猹（watcha.cn）及本中心杨浦园区样本口径。"
                "引用请注明：复旦大学住房政策研究中心《WAIC2026 人工智能产业空间白皮书》"
                "（FDU-HPRC-WP-2026-02）。")
        lines = wrap_text(self.draw, note, self.f_body, CONTENT_W)
        for ln in lines:
            self.ensure(self.line_h)
            self.draw.text((MARGIN_L, self.y), ln, font=self.f_body, fill=MUTED)
            self.y += self.line_h
        # 扉页单独成页，正文从下一页开始
        self.flush()


def build():
    md = open(MD, encoding="utf-8").read().split("\n")
    r = Report()
    r.make_cover()
    r.make_inside_cover()

    i = 0
    first_h1 = False
    while i < len(md):
        line = md[i].rstrip()
        if (not line or line == "---" or line.startswith("<p align")
                or line.startswith("<img")):
            i += 1
            continue
        if re.match(r"^\*\*(编制单位|文稿系列|文稿编号|成稿日期)\*\*", line):
            i += 1
            continue
        if line.startswith("# "):
            title = line[2:].strip()
            if not first_h1:
                first_h1 = True
                i += 1
                continue
            r.add_h1(title)
            i += 1
            continue
        if line.startswith("### "):
            r.add_h3(line[4:].strip())
            i += 1
            continue
        if line.startswith("## "):
            r.add_h2(line[3:].strip())
            i += 1
            continue
        m = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)", line)
        if m:
            alt, rel = m.group(1), m.group(2)
            r.add_image(os.path.join(BASE, rel), alt)
            i += 1
            continue
        if line.startswith("> "):
            buf = []
            while i < len(md) and md[i].startswith(">"):
                buf.append(md[i].lstrip("> ").strip())
                i += 1
            r.add_quote(" ".join(b for b in buf if b))
            continue
        if line.startswith("|"):
            rows = []
            while i < len(md) and md[i].strip().startswith("|"):
                cells = [c.strip() for c in md[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{3,}:?", c) for c in cells):
                    rows.append(cells)
                i += 1
            r.add_table(rows)
            continue
        if line.startswith("- "):
            items = []
            while i < len(md) and md[i].startswith("- "):
                items.append(md[i][2:].strip())
                i += 1
            r.add_bullets(items, ordered=False)
            continue
        m = re.match(r"^(\d+)\.\s+(.*)", line)
        if m:
            items = []
            while i < len(md):
                mm = re.match(r"^(\d+)\.\s+(.*)", md[i])
                if not mm:
                    break
                items.append(mm.group(2).strip())
                i += 1
            r.add_bullets(items, ordered=True)
            continue
        r.add_para(line, indent=True)
        i += 1

    r.flush()

    # 全页 JPEG，再用 img2pdf 封装为真正的 A4 页面尺寸
    paths = []
    for n, im in enumerate(r.pages, 1):
        if im.mode != "RGB":
            im = im.convert("RGB")
        path = os.path.join(PAGE_DIR, f"page_{n:02d}.jpg")
        im.save(path, "JPEG", quality=86, optimize=True)
        paths.append(path)

    import img2pdf
    a4 = img2pdf.mm_to_pt(210), img2pdf.mm_to_pt(297)
    layout = img2pdf.get_layout_fun(a4)
    with open(OUT_PDF, "wb") as f:
        f.write(img2pdf.convert(paths, layout_fun=layout))
    print("pages", len(r.pages), "pdf", OUT_PDF)
    return r.pages


if __name__ == "__main__":
    build()
