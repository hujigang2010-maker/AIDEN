#!/usr/bin/env python3
"""整理《江苏省高级工程师申报表》的字号、段间距和版式。

只调整版式，不改申报内容。工作总结里重复的“三、”会按出现顺序改为连续序号。
"""

from __future__ import annotations

import copy
import re
import sys
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
NS = {"w": W}

# A4 11906 x 16838 twips。页边距上下 1.5cm、左右 2.0cm 后的版心宽度。
PAGE_W = 11906
MARGIN_TB = 850
MARGIN_LR = 1134
CONTENT_W = PAGE_W - MARGIN_LR * 2  # 9638

SONG = "宋体"
HEI = "黑体"
KAI = "楷体"
ASCII = "Times New Roman"

H1_NUMS = "一二三四五六七八九十"
FIELD_KEYS = ("项目属性", "角色", "技术突破", "成果")
SUMMARY_START = "本人任现职以来工作总结"

# 简单数据表：首行是表头，套用表头字重和浅底。
HEADER_TABLES = {3, 4, 5, 6, 7, 8, 9, 10}

# 基本情况表里的栏目标题。比较时先去掉空白。
FORM_LABELS = {
    "姓名",
    "性别",
    "民族",
    "公民身份号码（社会保障号）",
    "出生年月",
    "参加工作年月",
    "取得本行业(专业)准入资格、时间及证书号码",
    "执业类别",
    "现专业技术资格及取得时间、批准单位",
    "现任专业技术职务及聘任时间",
    "申报类型",
    "现从事专业技术工作及年限",
    "拟评审专业技术资格",
    "学校",
    "学习专业",
    "学历",
    "学位",
    "毕(肄、结)业年月",
    "荣誉称号、表彰奖励名称",
    "获奖时间",
    "授奖部门",
    "处分",
}


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def paragraph_lines(p: etree._Element) -> list[str]:
    lines = [""]
    for child in p:
        if child.tag != q("r"):
            continue
        for el in child:
            if el.tag == q("t"):
                lines[-1] += el.text or ""
            elif el.tag in (q("br"), q("cr")):
                lines.append("")
    return lines


def paragraph_text(p: etree._Element) -> str:
    return "".join(paragraph_lines(p))


def has_drawing(el: etree._Element) -> bool:
    return (
        el.find(".//w:drawing", NS) is not None
        or el.find(".//w:pict", NS) is not None
        or el.find(".//{urn:schemas-microsoft-com:vml}shape") is not None
    )


def ensure_ppr(p: etree._Element) -> etree._Element:
    ppr = p.find(q("pPr"))
    if ppr is None:
        ppr = etree.Element(q("pPr"))
        p.insert(0, ppr)
    return ppr


def upsert(parent: etree._Element, tag: str, attrib: dict[str, str] | None = None) -> etree._Element:
    el = parent.find(q(tag))
    if el is None:
        el = etree.SubElement(parent, q(tag))
    if attrib is not None:
        el.attrib.clear()
        for key, value in attrib.items():
            el.set(q(key), value)
    return el


def remove_child(parent: etree._Element, tag: str) -> None:
    el = parent.find(q(tag))
    if el is not None:
        parent.remove(el)


def set_run_font(r: etree._Element, size_pt: float, east: str, bold: bool | None) -> None:
    rpr = r.find(q("rPr"))
    if rpr is None:
        rpr = etree.Element(q("rPr"))
        r.insert(0, rpr)
    half = str(int(round(size_pt * 2)))
    upsert(
        rpr,
        "rFonts",
        {"ascii": ASCII, "hAnsi": ASCII, "eastAsia": east, "cs": ASCII},
    )
    upsert(rpr, "sz", {"val": half})
    upsert(rpr, "szCs", {"val": half})
    if bold is None:
        return
    if bold:
        upsert(rpr, "b", {})
        upsert(rpr, "bCs", {})
    else:
        upsert(rpr, "b", {"val": "0"})
        upsert(rpr, "bCs", {"val": "0"})


def set_mark_font(ppr: etree._Element, size_pt: float, east: str, bold: bool) -> None:
    rpr = upsert(ppr, "rPr")
    half = str(int(round(size_pt * 2)))
    upsert(
        rpr,
        "rFonts",
        {"ascii": ASCII, "hAnsi": ASCII, "eastAsia": east, "cs": ASCII},
    )
    upsert(rpr, "sz", {"val": half})
    upsert(rpr, "szCs", {"val": half})
    if bold:
        upsert(rpr, "b", {})
        upsert(rpr, "bCs", {})
    else:
        upsert(rpr, "b", {"val": "0"})
        upsert(rpr, "bCs", {"val": "0"})


def clear_runs(p: etree._Element) -> None:
    for child in list(p):
        if child.tag != q("pPr"):
            p.remove(child)


def add_text_run(p: etree._Element, text: str, size_pt: float, east: str, bold: bool) -> None:
    r = etree.SubElement(p, q("r"))
    set_run_font(r, size_pt, east, bold)
    t = etree.SubElement(r, q("t"))
    if text[:1].isspace() or text[-1:].isspace():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text


def write_text(
    p: etree._Element,
    text: str,
    size_pt: float,
    east: str,
    *,
    bold: bool = False,
    label: bool = False,
) -> None:
    clear_runs(p)
    if label:
        for key in FIELD_KEYS:
            if text.startswith(key):
                rest = text[len(key) :].lstrip("：: \t")
                add_text_run(p, f"{key}：", size_pt, east, True)
                if rest:
                    add_text_run(p, rest, size_pt, east, False)
                return
        idx = text.find("：")
        if 0 < idx <= 18:
            add_text_run(p, text[: idx + 1], size_pt, east, True)
            rest = text[idx + 1 :]
            if rest:
                add_text_run(p, rest, size_pt, east, False)
            return
    add_text_run(p, text, size_pt, east, bold)


def format_paragraph(
    p: etree._Element,
    *,
    size_pt: float,
    east: str,
    before: int = 0,
    after: int = 0,
    line: int = 240,
    line_rule: str = "auto",
    align: str | None = None,
    first_chars: int = 0,
    left_chars: int = 0,
    bold: bool | None = False,
    keep_next: bool = False,
    rewrite: str | None = None,
    label: bool = False,
) -> None:
    """before/after/line 单位为 twips。first_chars、left_chars 为字符数百分之一（200 = 2 字）。"""
    ppr = ensure_ppr(p)
    upsert(ppr, "pStyle", {"val": "1"})
    upsert(
        ppr,
        "spacing",
        {
            "before": str(before),
            "after": str(after),
            "line": str(line),
            "lineRule": line_rule,
        },
    )
    if align:
        upsert(ppr, "jc", {"val": align})
    else:
        remove_child(ppr, "jc")

    if first_chars or left_chars:
        # 2 字符缩进按当前字号换算，避免只写 chars 时部分软件不生效。
        font_twips = int(round(size_pt * 20))
        attrib: dict[str, str] = {}
        if left_chars:
            attrib["leftChars"] = str(left_chars)
            attrib["left"] = str(int(round(font_twips * left_chars / 100)))
        if first_chars:
            attrib["firstLineChars"] = str(first_chars)
            attrib["firstLine"] = str(int(round(font_twips * first_chars / 100)))
        upsert(ppr, "ind", attrib)
    else:
        remove_child(ppr, "ind")

    if keep_next:
        upsert(ppr, "keepNext", {})
        upsert(ppr, "keepLines", {})
    else:
        remove_child(ppr, "keepNext")
        remove_child(ppr, "keepLines")
    upsert(ppr, "widowControl", {})

    set_mark_font(ppr, size_pt, east, bool(bold))
    if rewrite is not None:
        write_text(p, rewrite, size_pt, east, bold=bool(bold), label=label)
    else:
        for r in p.findall(q("r")):
            if has_drawing(r):
                continue
            set_run_font(r, size_pt, east, bold)


def is_field_line(text: str) -> bool:
    stripped = text.strip()
    return any(stripped.startswith(key) for key in FIELD_KEYS)


def explode_field_paragraph(p: etree._Element) -> None:
    if has_drawing(p):
        return
    lines = [line.strip() for line in paragraph_lines(p) if line.strip()]
    if len(lines) < 2 or not all(is_field_line(line) for line in lines):
        return
    ppr = p.find(q("pPr"))
    write_text(p, lines[0], 12, SONG, label=True)
    anchor = p
    for line in lines[1:]:
        np = etree.Element(q("p"))
        if ppr is not None:
            np.append(copy.deepcopy(ppr))
        np.attrib.pop(f"{{{W14}}}paraId", None)
        np.attrib.pop(f"{{{W14}}}textId", None)
        write_text(np, line, 12, SONG, label=True)
        anchor.addnext(np)
        anchor = np


def join_soft_wraps(p: etree._Element) -> None:
    lines = paragraph_lines(p)
    if len(lines) < 2:
        return
    if all(is_field_line(line) for line in lines if line.strip()):
        return
    text = "".join(lines)
    size = 14
    write_text(p, text, size, SONG, bold=False)


def summary_bounds(body: etree._Element) -> tuple[int, int]:
    children = list(body)
    start = None
    for i, child in enumerate(children):
        if child.tag == q("p") and paragraph_text(child).strip() == SUMMARY_START:
            start = i
            break
    if start is None:
        raise RuntimeError("未找到工作总结标题")
    end = len(children)
    for i in range(start + 1, len(children)):
        child = children[i]
        if child.tag == q("p") and child.find("w:pPr/w:sectPr", NS) is not None:
            end = i
            break
    return start, end


def renumber_summary_headings(paras: list[etree._Element]) -> None:
    headings = []
    for p in paras:
        text = paragraph_text(p).strip()
        if re.match(rf"^[{H1_NUMS}]+、", text):
            headings.append((p, text))
    for index, (p, text) in enumerate(headings):
        body = text.split("、", 1)[1]
        new = f"{H1_NUMS[index]}、{body}"
        if new != text:
            write_text(p, new, 14, HEI, bold=True)


def format_summary_paragraph(p: etree._Element) -> None:
    if has_drawing(p):
        return
    text = paragraph_text(p).strip()
    if not text:
        return

    if text == SUMMARY_START:
        format_paragraph(
            p, size_pt=18, east=HEI, before=0, after=40, align="center", bold=True, keep_next=True
        )
        return
    if text.startswith("（包括思想品德") or text.startswith("(包括思想品德"):
        format_paragraph(
            p, size_pt=10.5, east=KAI, before=0, after=160, align="center", bold=False, keep_next=True
        )
        return
    if text == "个人专业技术工作总结":
        format_paragraph(
            p, size_pt=14, east=HEI, before=80, after=200, align="center", bold=True, keep_next=True
        )
        return
    if re.match(rf"^[{H1_NUMS}]+、", text):
        format_paragraph(
            p,
            size_pt=14,
            east=HEI,
            before=280,
            after=120,
            align="left",
            bold=True,
            keep_next=True,
            rewrite=text,
        )
        return
    if re.match(r"^\d+\.\s*\S", text) and "：" not in text[:24]:
        format_paragraph(
            p,
            size_pt=12,
            east=HEI,
            before=240,
            after=80,
            align="left",
            bold=True,
            keep_next=True,
            rewrite=text,
        )
        return
    if is_field_line(text):
        format_paragraph(
            p,
            size_pt=12,
            east=SONG,
            before=20,
            after=80 if text.startswith("成果") else 20,
            line=360,
            line_rule="auto",
            align="both",
            left_chars=200,
            bold=False,
            keep_next=not text.startswith("成果"),
            rewrite=text,
            label=True,
        )
        return
    if text.startswith("本人签名"):
        format_paragraph(
            p, size_pt=12, east=SONG, before=360, after=0, align="right", bold=False, rewrite=text
        )
        return
    if re.fullmatch(r"年[\s\u3000]*月[\s\u3000]*日", text):
        format_paragraph(
            p, size_pt=12, east=SONG, before=80, after=0, align="right", bold=False, rewrite=text
        )
        return

    # 带冒号的编号短论（专业能力总结）与普通正文。
    lead = bool(re.match(r"^（?\d+", text) and "：" in text[:24])
    format_paragraph(
        p,
        size_pt=12,
        east=SONG,
        before=60 if lead or text.startswith("（") else 0,
        after=80,
        line=360,
        line_rule="auto",
        align="both",
        first_chars=200,
        bold=False,
        rewrite=text,
        label=lead or text.startswith("主要学术"),
    )


def format_front_paragraph(p: etree._Element) -> None:
    if has_drawing(p):
        return
    raw = paragraph_text(p)
    text = raw.strip()
    if not text:
        format_paragraph(p, size_pt=12, east=SONG, before=0, after=0, bold=False)
        return

    if text.startswith("专") and "资" in text and "格" in text:
        format_paragraph(p, size_pt=36, east=SONG, before=360, after=0, align="center", bold=False)
        return
    if text.startswith("评") and "申" in text and "报" in text:
        format_paragraph(p, size_pt=36, east=SONG, before=60, after=280, align="center", bold=False)
        return
    if text.startswith("填表时间"):
        format_paragraph(
            p, size_pt=14, east=KAI, before=280, after=80, align="center", bold=False, rewrite=text
        )
        return
    if "领导小组办公室制" in text:
        format_paragraph(p, size_pt=14, east=SONG, before=200, after=0, align="center", bold=False)
        return
    if text.replace(" ", "") in ("个人承诺书",):
        format_paragraph(
            p, size_pt=22, east=HEI, before=200, after=280, align="center", bold=True, keep_next=True
        )
        return
    if text.startswith("本人申报"):
        join_soft_wraps(p)
        format_paragraph(
            p,
            size_pt=14,
            east=SONG,
            before=80,
            after=80,
            line=560,
            line_rule="exact",
            align="both",
            first_chars=200,
            bold=False,
            rewrite="".join(paragraph_lines(p)),
        )
        return
    if text.startswith("承诺人"):
        format_paragraph(p, size_pt=14, east=SONG, before=400, after=0, align="right", bold=False)
        return
    if text.startswith("日") and "期" in text:
        format_paragraph(p, size_pt=14, east=SONG, before=160, after=200, align="right", bold=False)
        return
    if text.replace(" ", "") == "填表说明":
        format_paragraph(
            p, size_pt=18, east=HEI, before=280, after=160, align="center", bold=True, keep_next=True
        )
        return
    if re.match(r"^[123]、", text):
        join_soft_wraps(p)
        format_paragraph(
            p,
            size_pt=14,
            east=SONG,
            before=60,
            after=60,
            line=560,
            line_rule="exact",
            align="both",
            first_chars=200,
            bold=False,
            rewrite="".join(paragraph_lines(p)),
        )
        return

    section_18 = {
        "基  本  情  况",
        "继 续 教 育 情 况",
        "工  作  经  历",
        "任现职以来考核情况",
        "组  织  意  见",
        "评审、登记备案情况",
    }
    if text in section_18 or text.replace(" ", "") in {"基本情况", "继续教育情况", "工作经历", "组织意见"}:
        format_paragraph(
            p, size_pt=18, east=HEI, before=40, after=120, align="center", bold=True, keep_next=True
        )
        return
    if text in {"专业考试成绩情况"}:
        format_paragraph(
            p, size_pt=16, east=HEI, before=280, after=120, align="center", bold=True, keep_next=True
        )
        return
    if text in {"任现职后主要专业技术工作业绩登记", "任现职后著作、论文及重要技术报告登记"}:
        format_paragraph(
            p, size_pt=16, east=HEI, before=40, after=120, align="center", bold=True, keep_next=True
        )
        return
    if text.startswith("工作单位或行业协会"):
        format_paragraph(
            p, size_pt=14, east=HEI, before=40, after=120, align="center", bold=True, keep_next=True
        )
        return
    if text.startswith("（") and text.endswith("）"):
        format_paragraph(
            p, size_pt=12, east=KAI, before=0, after=140, align="center", bold=False, keep_next=True
        )
        return
    if text.startswith("说明"):
        format_paragraph(
            p, size_pt=10.5, east=SONG, before=80, after=40, align="left", bold=False, rewrite=text
        )
        return

    format_paragraph(p, size_pt=12, east=SONG, before=0, after=60, align="left", bold=False)


def scale_table(tbl: etree._Element, target: int) -> None:
    grid = tbl.findall("w:tblGrid/w:gridCol", NS)
    total = sum(int(col.get(q("w")) or 0) for col in grid)
    if total <= 0:
        return
    factor = target / total
    widths = [max(120, int(round(int(col.get(q("w")) or 0) * factor))) for col in grid]
    widths[-1] += target - sum(widths)
    if widths[-1] < 120:
        widths[-1] = 120
    for col, width in zip(grid, widths):
        col.set(q("w"), str(width))
    tbl_pr = tbl.find(q("tblPr"))
    if tbl_pr is None:
        tbl_pr = etree.Element(q("tblPr"))
        tbl.insert(0, tbl_pr)
    upsert(tbl_pr, "tblW", {"w": str(target), "type": "dxa"})
    upsert(tbl_pr, "jc", {"val": "center"})
    upsert(tbl_pr, "tblLayout", {"type": "fixed"})
    for tc_w in tbl.findall(".//w:tcW", NS):
        if tc_w.get(q("type")) != "dxa":
            continue
        old = int(tc_w.get(q("w")) or 0)
        tc_w.set(q("w"), str(max(120, int(round(old * factor)))))


def set_cover_table_widths(tbl: etree._Element) -> None:
    """封面信息表保持居中，并加宽内容列，避免单位名称挤成两行。"""
    widths = [1100, 1900, 6400]
    grid = tbl.findall("w:tblGrid/w:gridCol", NS)
    if len(grid) != 3:
        scale_table(tbl, sum(widths))
        return
    for col, width in zip(grid, widths):
        col.set(q("w"), str(width))
    tbl_pr = tbl.find(q("tblPr"))
    if tbl_pr is None:
        return
    upsert(tbl_pr, "tblW", {"w": str(sum(widths)), "type": "dxa"})
    upsert(tbl_pr, "jc", {"val": "center"})
    upsert(tbl_pr, "tblLayout", {"type": "fixed"})
    for tr in tbl.findall(q("tr")):
        cells = tr.findall(q("tc"))
        if len(cells) != 3:
            continue
        for tc, width in zip(cells, widths):
            tc_w = tc.find(f"{q('tcPr')}/{q('tcW')}")
            if tc_w is None:
                tc_pr = tc.find(q("tcPr"))
                if tc_pr is None:
                    tc_pr = etree.Element(q("tcPr"))
                    tc.insert(0, tc_pr)
                tc_w = etree.SubElement(tc_pr, q("tcW"))
            tc_w.set(q("w"), str(width))
            tc_w.set(q("type"), "dxa")


def apply_column_widths(tbl: etree._Element, widths: list[int]) -> None:
    """按列写入栏宽，合并单元格取所跨列之和，避免文字被挤成单字换行。"""
    grid = tbl.findall("w:tblGrid/w:gridCol", NS)
    if len(grid) != len(widths):
        return
    for col, width in zip(grid, widths):
        col.set(q("w"), str(width))
    total = sum(widths)
    tbl_pr = tbl.find(q("tblPr"))
    if tbl_pr is None:
        return
    upsert(tbl_pr, "tblW", {"w": str(total), "type": "dxa"})
    upsert(tbl_pr, "jc", {"val": "center"})
    upsert(tbl_pr, "tblLayout", {"type": "fixed"})
    for tr in tbl.findall(q("tr")):
        cursor = 0
        for tc in tr.findall(q("tc")):
            tc_pr = tc.find(q("tcPr"))
            span = 1
            if tc_pr is not None:
                grid_span = tc_pr.find(q("gridSpan"))
                if grid_span is not None:
                    span = int(grid_span.get(q("val")) or 1)
            if cursor + span > len(widths):
                break
            width = sum(widths[cursor : cursor + span])
            cursor += span
            if tc_pr is None:
                tc_pr = etree.Element(q("tcPr"))
                tc.insert(0, tc_pr)
            tc_w = tc_pr.find(q("tcW"))
            if tc_w is None:
                tc_w = etree.SubElement(tc_pr, q("tcW"))
            tc_w.set(q("w"), str(width))
            tc_w.set(q("type"), "dxa")


def set_cell_margins(tbl: etree._Element, top: int, side: int) -> None:
    tbl_pr = tbl.find(q("tblPr"))
    if tbl_pr is None:
        return
    mar = upsert(tbl_pr, "tblCellMar")
    for name, value in (("top", top), ("left", side), ("bottom", top), ("right", side)):
        upsert(mar, name, {"w": str(value), "type": "dxa"})


def shade_cell(tc: etree._Element, fill: str) -> None:
    tc_pr = tc.find(q("tcPr"))
    if tc_pr is None:
        tc_pr = etree.Element(q("tcPr"))
        tc.insert(0, tc_pr)
    upsert(tc_pr, "shd", {"val": "clear", "color": "auto", "fill": fill})


def set_valign(tc: etree._Element, val: str) -> None:
    tc_pr = tc.find(q("tcPr"))
    if tc_pr is None:
        tc_pr = etree.Element(q("tcPr"))
        tc.insert(0, tc_pr)
    upsert(tc_pr, "vAlign", {"val": val})


def cell_text(tc: etree._Element) -> str:
    parts = []
    for p in tc.findall(q("p")):
        parts.append(paragraph_text(p).strip())
    return "\n".join(part for part in parts if part)


def is_vertical_stack(tc: etree._Element) -> bool:
    paras = [paragraph_text(p).strip() for p in tc.findall(q("p"))]
    meaningful = [part for part in paras if part]
    return len(meaningful) >= 3 and all(len(part) <= 1 for part in meaningful)


def format_cell_paragraphs(
    tc: etree._Element,
    *,
    size_pt: float,
    east: str,
    bold: bool,
    align: str,
    before: int,
    after: int,
) -> None:
    for p in tc.findall(q("p")):
        if has_drawing(p):
            continue
        if not paragraph_text(p).strip() and not is_vertical_stack(tc):
            format_paragraph(p, size_pt=size_pt, east=east, before=0, after=0, align=align, bold=False)
            continue
        format_paragraph(
            p,
            size_pt=size_pt,
            east=east,
            before=before,
            after=after,
            align=align,
            bold=bold,
        )


def format_table(tbl: etree._Element, index: int) -> None:
    if index == 1:
        set_cover_table_widths(tbl)
        set_cell_margins(tbl, 60, 80)
    else:
        scale_table(tbl, CONTENT_W)
        set_cell_margins(tbl, 30, 46)
    custom_grids = {
        # 基本情况：加宽学历、学位、出生年月，避免“硕士研究生”“工商管理硕士”拆字。
        2: [1000, 900, 560, 960, 600, 480, 1140, 1420, 1340, 1238],
        # 工作经历：日期和职务留出余量，避免“至今”“工程师”落到下一行。
        5: [2360, 2318, 800, 840, 2040, 1280],
    }
    if index in custom_grids:
        apply_column_widths(tbl, custom_grids[index])
    if index == 2:
        set_cell_margins(tbl, 24, 28)

    rows = tbl.findall(q("tr"))
    for ri, tr in enumerate(rows):
        tr_pr = tr.find(q("trPr"))
        if tr_pr is None:
            tr_pr = etree.Element(q("trPr"))
            tr.insert(0, tr_pr)
        if index in HEADER_TABLES and ri == 0:
            upsert(tr_pr, "tblHeader")
        upsert(tr_pr, "cantSplit")
        height = tr_pr.find(q("trHeight"))
        if height is not None:
            val = int(height.get(q("val")) or 0)
            if val < 260:
                height.set(q("val"), "280")
                height.set(q("hRule"), "atLeast")

        header = index in HEADER_TABLES and ri == 0
        for tc in tr.findall(q("tc")):
            if has_drawing(tc):
                continue
            text = cell_text(tc)
            stack = is_vertical_stack(tc)
            if index == 1:
                tcs = tr.findall(q("tc"))
                col = tcs.index(tc)
                if col == 1:
                    format_cell_paragraphs(
                        tc, size_pt=14, east=KAI, bold=False, align="center", before=40, after=40
                    )
                elif col == 2 and text == "胡继刚":
                    format_cell_paragraphs(
                        tc, size_pt=18, east=HEI, bold=True, align="center", before=40, after=40
                    )
                else:
                    format_cell_paragraphs(
                        tc, size_pt=16, east=SONG, bold=False, align="center", before=40, after=40
                    )
                set_valign(tc, "center")
                continue
            if index == 0:
                align = "left" if tc is tr.findall(q("tc"))[0] else "right"
                format_cell_paragraphs(
                    tc, size_pt=12, east=SONG, bold=False, align=align, before=0, after=0
                )
                continue
            if stack:
                format_cell_paragraphs(
                    tc, size_pt=11, east=SONG, bold=True, align="center", before=0, after=0
                )
                set_valign(tc, "center")
                continue
            if index == 2 and re.sub(r"\s+", "", text) in FORM_LABELS:
                format_cell_paragraphs(
                    tc, size_pt=10.5, east=SONG, bold=True, align="center", before=30, after=30
                )
                set_valign(tc, "center")
                continue
            compact = re.sub(r"\s+", "", text)
            id_like = bool(re.fullmatch(r"\d{15,18}", compact))
            long_text = not id_like and (
                len(text.replace("\n", "")) >= 16
                or any(mark in text for mark in ("。", "：", "\n"))
            )
            if header:
                format_cell_paragraphs(
                    tc, size_pt=10.5, east=SONG, bold=True, align="center", before=40, after=40
                )
                shade_cell(tc, "E8EEF4")
                set_valign(tc, "center")
            elif long_text:
                format_cell_paragraphs(
                    tc, size_pt=10.5, east=SONG, bold=False, align="left", before=30, after=30
                )
                set_valign(tc, "center")
            else:
                format_cell_paragraphs(
                    tc, size_pt=10.5, east=SONG, bold=False, align="center", before=30, after=30
                )
                set_valign(tc, "center")


def set_section_margins(body: etree._Element) -> None:
    for sect in body.findall(".//w:sectPr", NS):
        pg_mar = sect.find(q("pgMar"))
        if pg_mar is None:
            continue
        pg_mar.set(q("top"), str(MARGIN_TB))
        pg_mar.set(q("bottom"), str(MARGIN_TB))
        pg_mar.set(q("left"), str(MARGIN_LR))
        pg_mar.set(q("right"), str(MARGIN_LR))
        pg_mar.set(q("header"), "340")
        pg_mar.set(q("footer"), "340")
        pg_mar.set(q("gutter"), "0")


def drop_empty_summary_paragraphs(body: etree._Element, start: int, end: int) -> None:
    children = list(body)
    for child in children[start:end]:
        if child.tag != q("p"):
            continue
        if child.find("w:pPr/w:sectPr", NS) is not None:
            continue
        if has_drawing(child):
            continue
        if paragraph_text(child).strip():
            continue
        body.remove(child)


def format_document(root: etree._Element) -> None:
    body = root.find(q("body"))
    if body is None:
        raise RuntimeError("文档缺少正文")

    set_section_margins(body)

    for p in list(body.findall(q("p"))):
        text = paragraph_text(p).strip()
        if text.startswith("本人申报") or re.match(r"^[23]、", text):
            join_soft_wraps(p)
        elif any(is_field_line(line) for line in paragraph_lines(p)):
            explode_field_paragraph(p)

    start, end = summary_bounds(body)
    summary_paras = [child for child in list(body)[start:end] if child.tag == q("p")]
    renumber_summary_headings(summary_paras)
    drop_empty_summary_paragraphs(body, start, end)

    start, end = summary_bounds(body)
    children = list(body)
    summary_ids = {id(child) for child in children[start:end]}
    for child in children:
        if child.tag != q("p"):
            continue
        if child.find("w:pPr/w:sectPr", NS) is not None:
            continue
        if id(child) in summary_ids:
            format_summary_paragraph(child)
        else:
            format_front_paragraph(child)

    for index, tbl in enumerate(body.findall(q("tbl"))):
        format_table(tbl, index)

def main() -> None:
    from docx import Document

    repo = Path(__file__).resolve().parents[1]
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else repo / "sources" / "11月12日1-江苏省高级工程师申报表_文字可编辑版.docx"
    dst = Path(sys.argv[2]) if len(sys.argv) > 2 else repo / "output" / "11月12日1-江苏省高级工程师申报表_排版.docx"
    doc = Document(str(src))
    format_document(doc.element)
    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(dst))
    print(f"已写出 {dst}")


if __name__ == "__main__":
    main()
