# -*- coding: utf-8 -*-
"""将白皮书 Markdown 文稿转换为 Word（docx）。"""
import os
import re

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = "/workspace/whitepaper"
MD = os.path.join(BASE, "跨境电商与中国企业出海白皮书.md")
OUTPUT = os.path.join(BASE, "跨境电商与中国企业出海白皮书.docx")

BLUE = RGBColor(0x0E, 0x4E, 0x9B)
RED = RGBColor(0xC8, 0x10, 0x2E)
GRAY = RGBColor(0x6B, 0x6F, 0x78)


def set_font(run, cn="宋体", en="Times New Roman", size=12, bold=False,
             color=None, italic=False):
    run.font.name = en
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    r = run._element.rPr
    rfonts = r.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = r.makeelement(qn("w:rFonts"), {})
        r.append(rfonts)
    rfonts.set(qn("w:eastAsia"), cn)


def add_rich(par, text, size=12, cn="宋体", color=None, base_bold=False):
    for seg in re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", text):
        if not seg:
            continue
        if seg.startswith("**") and seg.endswith("**"):
            set_font(par.add_run(seg[2:-2]), cn="黑体", size=size, bold=True,
                     color=color)
        elif seg.startswith("*") and seg.endswith("*") and len(seg) > 2:
            set_font(par.add_run(seg[1:-1]), cn="楷体", size=size, italic=True,
                     color=color)
        else:
            set_font(par.add_run(seg), cn=cn, size=size, bold=base_bold,
                     color=color)


def add_header_footer(doc):
    sec = doc.sections[0]
    header = sec.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_font(hp.add_run("复旦大学住房政策研究中心  ·  FDU-HPRC-WP-2026-03"),
             cn="楷体", size=9, color=GRAY)
    footer = sec.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(fp.add_run("跨境电商与中国企业出海：空间重构、规则跃迁与二〇三〇展望"),
             cn="楷体", size=9, color=GRAY)


doc = Document()
sec = doc.sections[0]
sec.top_margin, sec.bottom_margin = Cm(2.6), Cm(2.6)
sec.left_margin, sec.right_margin = Cm(3.0), Cm(3.0)
add_header_footer(doc)

# ---------------- 封面 ----------------
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(24)
p.add_run().add_picture(os.path.join(BASE, "assets/logo_fudan_hprc.png"),
                        width=Cm(13.5))

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(48)
set_font(p.add_run("跨境电商与中国企业出海"), cn="黑体", en="Arial", size=28,
         bold=True, color=BLUE)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(10)
set_font(p.add_run("空间重构、规则跃迁与二〇三〇展望"), cn="黑体", en="Arial",
         size=18, bold=True, color=BLUE)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(22)
set_font(p.add_run("—— 从货物通达到系统扎根的白皮书 ——"),
         cn="楷体", size=13, color=GRAY)

for txt, sz, sp in [("中心研究文稿 · 第三号", 15, 56),
                    ("文稿编号：FDU-HPRC-WP-2026-03", 13, 6),
                    ("复旦大学住房政策研究中心", 16, 28),
                    ("Housing Policy Research Center, Fudan University", 11, 2),
                    ("二〇二六年八月 · 上海", 13, 18)]:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(sp)
    bold = txt.startswith(("中心研究", "复旦大学"))
    set_font(p.add_run(txt), cn="黑体" if bold else "宋体", size=sz, bold=bold,
             color=BLUE if bold else GRAY)
doc.add_page_break()

CHART_INSERTS = {
    "1.1 国内坐标": "chart01_cbec_scale.png",
    "1.2 国际坐标": "chart03_rule_timeline.png",
    "1.4 空间坐标": "chart07_space_four.png",
    "2.3 四维合规": "chart06_compliance.png",
    "3.1 三种出海": "chart02_outbound_layers.png",
    "3.2 市场分层": "chart05_market_layers.png",
    "4.1 货之居所": "chart04_overseas_warehouse.png",
    "5.2 三条情景": "chart08_2030_scenarios.png",
    "6.3 对国家政策": "chart09_policy.png",
}


def maybe_insert_chart(heading_text):
    for key, fname in CHART_INSERTS.items():
        if heading_text.startswith(key):
            path = os.path.join(BASE, "assets/charts", fname)
            if os.path.exists(path):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(8)
                p.add_run().add_picture(path, width=Cm(15.2))
            return


lines = open(MD, encoding="utf-8").read().split("\n")
i = 0
first_h1_skipped = False
chapter_no = 0
while i < len(lines):
    line = lines[i].rstrip()

    if (not line or line == "---" or line.startswith("<p align")
            or line.startswith("<img") or line.startswith("## 空间重构")
            or line.startswith("## 从“货物")):
        i += 1
        continue

    if re.match(r"^\*\*(编制单位|文稿系列|文稿编号|成稿日期)\*\*", line):
        i += 1
        continue

    if line.startswith("# "):
        title = line[2:].strip()
        if not first_h1_skipped:
            first_h1_skipped = True
            i += 1
            continue
        if chapter_no > 0:
            doc.add_page_break()
        chapter_no += 1
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(14)
        set_font(p.add_run(title), cn="黑体", en="Arial", size=18, bold=True,
                 color=BLUE)
        i += 1
        continue

    if line.startswith("### "):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(6)
        set_font(p.add_run(line[4:].strip()), cn="黑体", en="Arial", size=13,
                 bold=True, color=BLUE)
        i += 1
        continue

    if line.startswith("## "):
        heading = line[3:].strip()
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(8)
        set_font(p.add_run(heading), cn="黑体", en="Arial", size=14.5, bold=True)
        maybe_insert_chart(heading)
        i += 1
        continue

    m = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)", line)
    if m:
        path = os.path.join(BASE, m.group(2))
        if os.path.exists(path):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(path, width=Cm(15))
        i += 1
        continue

    if line.startswith("> "):
        buf = []
        while i < len(lines) and lines[i].startswith(">"):
            buf.append(lines[i].lstrip("> ").strip())
            i += 1
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.75)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(8)
        add_rich(p, " ".join(b for b in buf if b), size=11.5, cn="楷体",
                 color=BLUE)
        continue

    if line.startswith("|"):
        rows = []
        while i < len(lines) and lines[i].strip().startswith("|"):
            cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
            if not all(re.fullmatch(r":?-{3,}:?", c) for c in cells):
                rows.append(cells)
            i += 1
        if rows:
            ncol = max(len(r) for r in rows)
            table = doc.add_table(rows=len(rows), cols=ncol)
            table.style = "Table Grid"
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            for ri, row in enumerate(rows):
                for ci in range(ncol):
                    cell = table.cell(ri, ci)
                    txt = row[ci] if ci < len(row) else ""
                    txt = re.sub(r"\*\*([^*]+)\*\*", r"\1", txt)
                    cp = cell.paragraphs[0]
                    set_font(cp.add_run(txt), cn="黑体" if ri == 0 else "宋体",
                             size=9.5, bold=(ri == 0))
                    if ri == 0:
                        shd = cp._element.makeelement(qn("w:shd"), {
                            qn("w:val"): "clear", qn("w:fill"): "DCE6F1"})
                        cell._element.get_or_add_tcPr().append(shd)
            doc.add_paragraph().paragraph_format.space_after = Pt(2)
        continue

    if line.startswith("- "):
        p = doc.add_paragraph(style="List Bullet")
        add_rich(p, line[2:].strip(), size=12)
        i += 1
        continue

    m = re.match(r"^(\d+)\.\s+(.*)", line)
    if m:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.75)
        p.paragraph_format.first_line_indent = Cm(-0.5)
        p.paragraph_format.space_after = Pt(4)
        set_font(p.add_run(f"{m.group(1)}. "), cn="黑体", size=12, bold=True)
        add_rich(p, m.group(2).strip(), size=12)
        i += 1
        continue

    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(24)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.35
    add_rich(p, line, size=12)
    i += 1

os.makedirs("/workspace/下载版本", exist_ok=True)
doc.save(OUTPUT)
doc.save("/workspace/下载版本/跨境电商与中国企业出海白皮书.docx")
print("saved:", OUTPUT)
