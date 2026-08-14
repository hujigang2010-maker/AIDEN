# -*- coding: utf-8 -*-
"""将白皮书 Markdown 转换为 Word。"""
import os
import re
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = "/workspace/whitepaper"
MD = os.path.join(BASE, "医疗养老与提效服务白皮书.md")
OUTPUT = os.path.join(BASE, "医疗养老与提效服务白皮书.docx")
BLUE = RGBColor(0x0E, 0x4E, 0x9B)
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
            set_font(par.add_run(seg[2:-2]), cn="黑体", size=size, bold=True, color=color)
        elif seg.startswith("*") and seg.endswith("*") and len(seg) > 2:
            set_font(par.add_run(seg[1:-1]), cn="楷体", size=size, italic=True, color=color)
        else:
            set_font(par.add_run(seg), cn=cn, size=size, bold=base_bold, color=color)


doc = Document()
sec = doc.sections[0]
sec.top_margin, sec.bottom_margin = Cm(2.6), Cm(2.6)
sec.left_margin, sec.right_margin = Cm(3.0), Cm(3.0)
hp = sec.header.paragraphs[0]
hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
set_font(hp.add_run("复旦大学住房政策研究中心  ·  FDU-HPRC-WP-2026-04"),
         cn="楷体", size=9, color=GRAY)
fp = sec.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(fp.add_run("医疗、养老与提效服务：照护空间、劳动释放与二〇三〇展望"),
         cn="楷体", size=9, color=GRAY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(24)
p.add_run().add_picture(os.path.join(BASE, "assets/logo_fudan_hprc.png"), width=Cm(13.5))

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(44)
set_font(p.add_run("医疗、养老与提效服务"), cn="黑体", en="Arial", size=28, bold=True, color=BLUE)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(10)
set_font(p.add_run("照护空间、劳动释放与二〇三〇展望"), cn="黑体", en="Arial", size=18, bold=True, color=BLUE)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(20)
set_font(p.add_run("—— 把家庭隐形税改写成可及的社会化服务 ——"), cn="楷体", size=13, color=GRAY)

for txt, sz, sp in [("中心研究文稿 · 第四号", 15, 52),
                    ("文稿编号：FDU-HPRC-WP-2026-04", 13, 6),
                    ("复旦大学住房政策研究中心", 16, 26),
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
    "1.1 人口账": "chart01_aging.png",
    "1.2 政策账": "chart07_timeline.png",
    "1.3 效率账": "chart05_three_services.png",
    "2.1 资源总量": "chart03_health.png",
    "3.1 供给结构": "chart02_beds.png",
    "5.1 住宅": "chart06_space.png",
    "6.1 参照系": "chart04_15th_plan.png",
    "6.2 三条情景": "chart08_scenarios.png",
    "7.4 对国家制度": "chart09_policy.png",
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
            or line.startswith("<img") or line.startswith("## 照护空间")):
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
        set_font(p.add_run(title), cn="黑体", en="Arial", size=18, bold=True, color=BLUE)
        i += 1
        continue
    if line.startswith("### "):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(6)
        set_font(p.add_run(line[4:].strip()), cn="黑体", en="Arial", size=13, bold=True, color=BLUE)
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
        add_rich(p, " ".join(b for b in buf if b), size=11.5, cn="楷体", color=BLUE)
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
doc.save("/workspace/下载版本/医疗养老与提效服务白皮书.docx")
print("saved:", OUTPUT)
