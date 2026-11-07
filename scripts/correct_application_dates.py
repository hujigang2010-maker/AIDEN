#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按《工作经历一览表》校正申报表中的任职时间，并标出无法归属的奖项与项目。

原则：
- 一览表是任职时间与单位的唯一依据，不补写空档月份，不编造日。
- 同一单位只改正文中的起止与职务，不把无证据的奖项或项目改挂到别的年份。
- 无法归属的奖励、项目保留原文，并标明“待核查”。
- 水印、表格栏线和页数保持不变。
"""

from __future__ import annotations

import re
import shutil
from pathlib import Path

import openpyxl
import pymupdf
from fontTools.subset import Options, Subsetter
from fontTools.ttLib import TTFont
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path("/workspace")
SRC_PDF = Path("/home/ubuntu/.cursor/projects/workspace/uploads/PDF-2_91b1.pdf")
SRC_OTF = Path("/tmp/NotoSerifSC-Regular.otf")
SRC_TTF = Path("/tmp/NotoSerifSC-Regular.ttf")
OUT_DIR = ROOT / "交付"
FONT_SUBSET = Path("/tmp/NotoSerifSC-subset.ttf")
OUT_PDF = OUT_DIR / "江苏省高级工程师专业技术资格评审申报表-任职时间校正.pdf"
OUT_XLSX = OUT_DIR / "任职时间修改对照表.xlsx"
ART_DIR = Path("/opt/cursor/artifacts")

# 一览表 18 段，新到旧。截止“今”不补结束月。岗位栏一览表没有：
# 管理职务沿用原表“管理岗”；资料员、施工员不记成管理岗，记为技术岗。
JOBS = [
    ("2026年05月", "今", "上海贝迪创建科技有限公司", "工程部", "管理岗", "工程技术及项目管理", "高级工程师"),
    ("2026年01月", "2026年04月", "上海宝龙商办企业管理有限公司", "工程部", "管理岗", "工程技术及项目管理", "副总经理"),
    ("2025年03月", "2025年12月", "江苏省建筑装饰设计研究院有限公司", "工程部", "管理岗", "工程技术及项目管理", "副总经理"),
    ("2024年08月", "2025年02月", "上海惟嘉建筑工程有限公司", "工程部", "管理岗", "工程施工技术管理", "副总经理"),
    ("2023年10月", "2024年07月", "上海徐汇绿地商业管理有限公司", "工程部", "管理岗", "工程施工技术管理", "副总经理"),
    ("2023年05月", "2023年08月", "上海容钲建筑科技有限公司", "工程部", "管理岗", "工程施工技术管理", "副总经理"),
    ("2022年08月", "2023年04月", "上海苏科企业管理有限公司", "工程部", "管理岗", "钢结构施工技术管理", "副总经理"),
    ("2021年09月", "2022年07月", "上海社通实业有限公司", "工程部", "管理岗", "工程施工技术管理", "总监"),
    ("2021年03月", "2021年08月", "上海融之旅文化旅游发展有限公司", "工程部", "管理岗", "工程施工技术管理", "总监"),
    ("2021年01月", "2021年02月", "上海弘阳汇商业管理有限公司", "工程部", "管理岗", "工程施工技术管理", "总监"),
    ("2020年11月", "2020年12月", "上海欣驭文化传播有限公司", "工程部", "管理岗", "工程施工技术管理", "总监"),
    ("2020年06月", "2020年10月", "上海金科文化旅游发展集团有限公司", "工程部", "管理岗", "工程施工技术管理", "总监"),
    ("2020年03月", "2020年05月", "金科（上海）建筑设计有限公司", "工程部", "管理岗", "工程施工技术管理", "总监"),
    ("2019年04月", "2020年02月", "上海万居企业管理有限公司", "工程部", "管理岗", "工程施工技术管理", "总监"),
    ("2018年01月", "2019年03月", "万科（上海）实业发展有限公司", "工程部", "管理岗", "工程施工技术管理", "总监"),
    ("2017年03月", "2017年12月", "上海万科长宁置业有限公司", "工程部", "管理岗", "工程施工技术管理", "总监"),
    ("2015年12月", "2017年02月", "新城控股集团股份有限公司上海第一分公司", "工程部", "管理岗", "工程管理级施工协调", "高级经理"),
    ("2015年09月", "2015年10月", "上海睿泽股权投资管理有限公司", "工程部", "技术岗", "施工资料及现场管理", "资料员、施工员"),
]

PAGE5_NOTE = "任职年月按工作经历一览表填写，不补具体日。空档2015-11与2023-09未补。"
PAGE6_NOTE = "本页三项起止与任职时段或奖项日期不一致，只标待核查，不改项目归属。"
PAGE8_NOTE = "空档2015-11、2023-09未补。一览表没有的单位已从本栏删除。"

# 原表第5页 8 行，仅用于对照，不直接写回。
OLD_PAGE5 = [
    "2025年03月01日至2025年07月10日，江苏省建筑装饰设计研究院有限公司，副总经理",
    "2023年06月07日至2025年02月28日，上海徐汇绿地商业管理有限公司，副总经理",
    "2022年04月11日至2023年05月13日，上海融之旅文化旅游发展有限公司（融创文旅华东区域），总监",
    "2018年12月10日至2022年03月18日，上海万科长宁置业有限公司，总监",
    "2017年11月06日至2018年11月10日，上海万限实业有限公司【万科（上海）实业发展有限公司】，总监",
    "2017年06月01日至2017年10月26日，上海金科文化旅游发展集团有限公司，资深经理",
    "2017年03月20日至2017年05月29日，金科（上海）建筑设计有限公司，资深经理",
    "2015年06月08日至2017年02月24日，新城控股集团股份有限公司上海第一分公司，高级经理",
]

LINE_REPLACEMENTS = [
    (
        "1. 中南海湾新城·金石国际广场（2011.07-2015.08）",
        "1. 中南海湾新城·金石国际广场（2011.07-2015.08，待核查）",
    ),
    (
        "3、合肥金科博翠天辰2017.03-2017.10",
        "3、合肥金科博翠天辰2017.03-2017.10（待核查）",
    ),
    (
        "角色：项目技术负责人（隶属金科集团华东区域公司）",
        "角色：项目技术负责人（隶属金科集团华东区域公司，待核查）",
    ),
    (
        "4.上海万科天空之城（2017.11-2022.02）",
        "4.上海万科天空之城（2017.11-2022.02，待核查）",
    ),
    (
        "5.杭州湾融创文旅城水路乐园项目（海盐）（2022.04-2023.05）",
        "5.杭州湾融创文旅城水路乐园项目（海盐）（2022.04-2023.05，待核查）",
    ),
    (
        "6.苏州绿地江南映月花园（2023.06-2025.03）",
        "6.苏州绿地江南映月花园（2023.06-2025.03，待核查）",
    ),
    (
        "6、项目名称：东台市养老服务中心地下车库装饰工程（2025.3-至今）",
        "6、项目名称：东台市养老服务中心地下车库装饰工程（2025.03起，待核查）",
    ),
]
PAGE9_LINES = LINE_REPLACEMENTS[:3]
PAGE10_LINES = LINE_REPLACEMENTS[3:]


def collect_chars() -> str:
    parts = [PAGE5_NOTE, PAGE6_NOTE, PAGE8_NOTE, "待核查", "至", "今"]
    for job in JOBS:
        parts.extend(job)
        parts.append(f"{job[0]}-{job[1]}".replace("年", ".").replace("月", ""))
    for _old, new in LINE_REPLACEMENTS:
        parts.append(new)
    parts.append("专业技术工作年限：9年")
    parts.append("2015年09月")
    # 页8用点号年月
    for start, end, company, _dept, _post, _work, title in JOBS:
        a = start.replace("年", ".").replace("月", "")
        b = end if end == "今" else end.replace("年", ".").replace("月", "")
        parts.append(f"{a}-{b}{company}{title}")
    return "".join(parts)


def ensure_ttf() -> None:
    if SRC_TTF.exists():
        return
    if not SRC_OTF.exists():
        raise SystemExit(f"缺少字体：{SRC_OTF}")
    from otf2ttf import main as otf2ttf_main

    otf2ttf_main([str(SRC_OTF), "-o", str(SRC_TTF), "--overwrite"])


def is_compat(code: int) -> bool:
    return (
        0xF900 <= code <= 0xFAFF
        or 0x2F00 <= code <= 0x2FDF
        or code in (0xAD, 0x2010, 0x2011)
    )


def clean_cmap(font: TTFont, used: set[int]) -> None:
    """同一字形只保留正文要用的码位，避免写出兼容汉字。"""
    for table in font["cmap"].tables:
        if not hasattr(table, "cmap"):
            continue
        reverse: dict[str, list[int]] = {}
        for code, name in table.cmap.items():
            reverse.setdefault(name, []).append(code)
        cleaned: dict[int, str] = {}
        for name, codes in reverse.items():
            keep = [code for code in codes if code in used]
            if not keep:
                keep = [code for code in codes if not is_compat(code)] or codes
                keep = [min(keep)]
            for code in keep:
                cleaned[code] = name
        table.cmap = cleaned


def build_subset(chars: str) -> None:
    ensure_ttf()
    font = TTFont(SRC_TTF)
    clean_cmap(font, set(map(ord, chars)))
    options = Options()
    options.layout_features = ["*"]
    options.name_IDs = ["*"]
    options.notdef_outline = True
    options.hinting = False
    subsetter = Subsetter(options=options)
    subsetter.populate(text=chars)
    subsetter.subset(font)
    font.save(FONT_SUBSET)


def assert_font_readable(sample: str) -> None:
    probe = pymupdf.open()
    page = probe.new_page()
    page.insert_font(fontname="f", fontfile=str(FONT_SUBSET))
    page.insert_text((36, 80), sample, fontname="f", fontsize=14, color=(0, 0, 0))
    got = "".join(page.get_text().split())
    want = "".join(sample.split())
    if got != want:
        raise SystemExit(f"字体编码不正确：{got}")
    pix = page.get_pixmap(matrix=pymupdf.Matrix(2, 2), clip=pymupdf.Rect(30, 45, 560, 110), alpha=False)
    dark = sum(1 for y in range(pix.height) for x in range(pix.width) if pix.pixel(x, y)[0] < 40)
    if dark < 800:
        raise SystemExit(f"字体没有画出字形 dark={dark}")


def wrap_text(font: pymupdf.Font, text: str, size: float, max_w: float) -> list[str]:
    if font.text_length(text, size) <= max_w:
        return [text]
    lines: list[str] = []
    cur = ""
    for ch in text:
        trial = cur + ch
        if font.text_length(trial, size) <= max_w:
            cur = trial
        else:
            if not cur:
                raise SystemExit(f"单字超宽：{ch}")
            lines.append(cur)
            cur = ch
    if cur:
        lines.append(cur)
    return lines


WATERMARK_KEY = (
    "54 823.14 m\n540.75 823.14 l\n540.75 19.14 l\n54 19.14 l\n"
    "54 823.14 l\nh\nW\nn\n/GS"
)


def capture_extgstate(doc: pymupdf.Document, page: pymupdf.Page) -> str | None:
    raw = doc.xref_object(page.xref)
    match = re.search(r"/ExtGState\s*<<[^>]*>>", raw)
    return match.group(0) if match else None


def detach_watermark(doc: pymupdf.Document, page: pymupdf.Page):
    xrefs = page.get_contents()
    if len(xrefs) != 1:
        raise SystemExit(f"页内容流数量不是 1：{xrefs}")
    xref = xrefs[0]
    text = doc.xref_stream(xref).decode("latin1")
    idx = text.find(WATERMARK_KEY)
    if idx < 0 or text.find(WATERMARK_KEY, idx + 1) >= 0:
        raise SystemExit("未能唯一拆出水印")
    gs = capture_extgstate(doc, page)
    doc.update_stream(xref, text[:idx].encode("latin1"))
    page = doc.reload_page(page)
    return page, text[idx:].encode("latin1"), gs


def attach_watermark(doc: pymupdf.Document, page: pymupdf.Page, water: bytes, gs: str | None):
    xrefs = page.get_contents()
    xref = xrefs[-1]
    current = doc.xref_stream(xref)
    doc.update_stream(xref, current + b"\n" + water)
    page = doc.reload_page(page)
    if gs:
        name = re.search(r"/(GS\d+)", gs).group(1)
        ref = re.search(r"(\d+\s+0\s+R)", gs).group(1)
        doc.xref_set_key(page.xref, f"Resources/ExtGState/{name}", ref)
        raw = doc.xref_object(page.xref)
        if f"/{name}" not in raw:
            raise SystemExit(f"水印透明度资源未写回：{name}\n{raw[:500]}")
    return page


FONTS_READY: set[int] = set()


def ensure_font(page: pymupdf.Page) -> None:
    if page.xref in FONTS_READY:
        return
    page.insert_font(fontname="nsong", fontfile=str(FONT_SUBSET))
    FONTS_READY.add(page.xref)


def put(page: pymupdf.Page, point, text: str, size: float) -> None:
    ensure_font(page)
    page.insert_text(point, text, fontname="nsong", fontsize=size, color=(0, 0, 0))


def prepare_page(doc: pymupdf.Document, index: int):
    page, water, gs = detach_watermark(doc, doc[index])
    return page, water, gs


def finish_page(doc, page, water, gs):
    return attach_watermark(doc, page, water, gs)


def redact_rects(page: pymupdf.Page, rects: list[pymupdf.Rect]) -> None:
    for rect in rects:
        page.add_redact_annot(rect, fill=None)
    if rects:
        page.apply_redactions(images=0, graphics=0, text=0)


def find_line(page: pymupdf.Page, text: str):
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            joined = "".join(span["text"] for span in line["spans"])
            if joined == text:
                return line
    raise SystemExit(f"找不到原文行：{text}")


def draw_lines(page, font, rect, lines, size, align):
    leading = size + 2.3
    block_h = size + (len(lines) - 1) * leading
    top = rect.y0 + max(0.8, (rect.height - block_h) / 2)
    baseline = top + size * 0.80
    for index, line in enumerate(lines):
        width = font.text_length(line, size)
        if align == "center":
            x = rect.x0 + (rect.width - width) / 2
        else:
            x = rect.x0 + 2.2
        if x < rect.x0 + 0.4:
            x = rect.x0 + 0.4
        if x + width > rect.x1 - 0.4:
            raise SystemExit(f"文字超出单元格：{line} {x + width:.1f}>{rect.x1 - 0.4:.1f}")
        y = baseline + index * leading
        if y > rect.y1 - 0.6:
            raise SystemExit(f"文字超出单元格底部：{line} {y:.1f}>{rect.y1:.1f}")
        put(page, (x, y), line, size)


def edit_page3(page, font):
    # 参加工作年月、专业技术工作年限。单元格右边界紧，只缩短不增字。
    targets = {
        "2011年06月27日": "2015年09月",
        "专业技术工作年限：14年": "专业技术工作年限：9年",
    }
    saved = []
    rects = []
    for old, new in targets.items():
        line = find_line(page, old)
        span = line["spans"][0]
        size = span["size"]
        if font.text_length(new, size) > (span["bbox"][2] - span["bbox"][0]) + 8:
            raise SystemExit(f"替换后超宽：{new}")
        saved.append((span["origin"], new, size))
        rects.append(pymupdf.Rect(line["bbox"]) + (-0.4, -0.4, 0.4, 0.4))
    redact_rects(page, rects)
    for origin, new, size in saved:
        put(page, origin, new, size)


def edit_page5(page, font):
    # 去掉原 8 行正文，保留栏线与表头。下方空白按同样栏宽续 10 行。
    redact_rects(page, [pymupdf.Rect(54.4, 105.0, 525.4, 412.6)])
    bounds = [
        104.25, 143.25, 181.50, 220.50, 258.75, 297.00, 336.00, 374.25, 413.25,
        451.50, 490.50, 528.75, 567.75, 606.00, 645.00, 683.25, 722.25, 760.50, 799.50,
    ]
    if len(bounds) != len(JOBS) + 1:
        raise SystemExit("行数与任职段数不一致")
    xs = [54.0, 134.25, 236.25, 304.5, 372.75, 457.5, 525.75]
    bottom = bounds[-1]
    for x in xs:
        page.draw_line((x, 413.25), (x, bottom), color=(0, 0, 0), width=0.75)
    for y in bounds[9:]:
        page.draw_line((xs[0], y), (xs[-1], y), color=(0, 0, 0), width=0.75)
    for row, job in enumerate(JOBS):
        start, end, company, dept, post, work, title = job
        y0, y1 = bounds[row], bounds[row + 1]
        rects = [pymupdf.Rect(xs[i], y0, xs[i + 1], y1) for i in range(6)]
        draw_lines(page, font, rects[0], [f"{start}至", end], 9, "left")
        company_lines = wrap_text(font, company, 9, rects[1].width - 3.2)
        if len(company_lines) > 2:
            raise SystemExit(f"单位超过两行：{company}")
        draw_lines(page, font, rects[1], company_lines, 9, "center")
        draw_lines(page, font, rects[2], [dept], 9, "center")
        draw_lines(page, font, rects[3], [post], 9, "center")
        draw_lines(page, font, rects[4], [work], 9, "center")
        draw_lines(page, font, rects[5], [title], 9, "center")
    put(page, (56, 816), PAGE5_NOTE, 8)


def edit_page6(page, font):
    # 不改项目名称、奖项句子和起止数字，只在日期栏下标注。
    stamps = [(81.4, 186), (81.4, 278), (81.4, 404)]
    for x, y in stamps:
        text = "待核查"
        width = font.text_length(text, 8)
        put(page, (x - width / 2, y), text, 8)
    put(page, (56, 476), PAGE6_NOTE, 9)


def edit_page8(page, font):
    redact_rects(page, [pymupdf.Rect(58, 416, 524, 800)])
    # 盖住原来 7 条行首短竖线，避免和 18 行新文字错位。
    page.draw_rect(
        pymupdf.Rect(59.3, 416, 62.6, 780),
        color=(1, 1, 1),
        fill=(1, 1, 1),
        width=0,
    )
    for index, job in enumerate(JOBS):
        start, end, company, _dept, _post, _work, title = job
        a = start.replace("年", ".").replace("月", "")
        b = "今" if end == "今" else end.replace("年", ".").replace("月", "")
        y = 432 + index * 19.2
        put(page, (66, y), f"{a}-{b}", 9)
        put(page, (176, y), company, 9)
        put(page, (400, y), title, 9)
    put(page, (66, 792), PAGE8_NOTE, 8)


def edit_summary_lines(page, font, pairs):
    saved = []
    rects = []
    for old, new in pairs:
        line = find_line(page, old)
        span = line["spans"][0]
        size = span["size"]
        # 右边界留在版心内
        if span["origin"][0] + font.text_length(new, size) > 524:
            raise SystemExit(f"总结行超出版心：{new}")
        saved.append((span["origin"], new, size))
        rects.append(pymupdf.Rect(line["bbox"]) + (-0.3, -0.3, 0.6, 0.3))
    redact_rects(page, rects)
    for origin, new, size in saved:
        put(page, origin, new, size)


def job_label(job) -> str:
    start, end, company, _dept, _post, work, title = job
    return f"{start}至{end}，{company}，{title}，{work}"


def build_rows() -> list[dict]:
    rows: list[dict] = []

    def add(page, item, old, new, action, reason):
        rows.append(
            {
                "页": page,
                "栏目": item,
                "原文": old,
                "修改后": new,
                "处理": action,
                "依据与说明": reason,
            }
        )

    add(
        "1",
        "封面单位、填表时间",
        "单位：江苏省建筑装饰设计研究院有限公司；填表时间2025年08月12日",
        "未改",
        "核对后未改",
        "填表日处于一览表第16段（2025-03至2025-12，该单位，工程部副总经理）之内。",
    )
    add(
        "3",
        "参加工作年月",
        "2011年06月27日",
        "2015年09月",
        "已改正文",
        "一览表首段为2015-09上海睿泽股权投资管理有限公司。一览表只到月，不把27日挪到2015年，也不补日。",
    )
    add(
        "3",
        "专业技术工作年限",
        "14年",
        "9年",
        "已改正文",
        "原14年对应2011-06。按参加工作2015-09至填表日2025-08，实足9年11个月，取已满年份9年。单元格右边界紧，不能写成“9年11个月”。",
    )
    add(
        "3",
        "现资格及聘用时间",
        "工程师，批准20181225，聘用2018-12-25",
        "未改",
        "核对后未改",
        "该日处于万科（上海）实业发展有限公司任职（2018-01至2019-03）内。这是资格日期，不是任职起止。",
    )
    add(
        "3",
        "本科、硕士毕结业时间",
        "海洋大学2011-06-30；复旦大学2021-01-13",
        "未改",
        "核对后未改",
        "学历时间，不是任职。2021-01与弘阳汇任职重叠，不构成冲突。",
    )
    add(
        "3",
        "高中毕结业时间",
        "青岛市城阳一中，于2016年07月04日毕业",
        "未改，待核查",
        "保留待核查",
        "晚于2011年本科毕业，也落在新城控股任职期内，时间顺序异常。一览表不提供学历日期，不臆改，正文不添加字。",
    )
    add(
        "3",
        "住房政策研究创新贡献奖",
        "2024-12-18，复旦大学住房政策中心专业委员会",
        "未改",
        "核对后未改",
        "获奖月任职上海惟嘉建筑工程有限公司。授奖单位不是任职单位，摘要里的秘书长职务没有起止。不把奖项改挂到惟嘉，也不改日期。",
    )
    add(
        "3",
        "城市更新卓越实践奖",
        "2024-04-26，绿地控股集团技术委员会",
        "未改",
        "核对后未改",
        "获奖月任职上海徐汇绿地商业管理有限公司（2023-10至2024-07），与“绿地”相符。第6页把该奖挂到另一段项目起止上，项目行另标待核查，本行不改。",
    )
    add(
        "3",
        "可持续社区营造标杆项目奖",
        "2019-02-24，上海万科产品创新研究院",
        "未改",
        "核对后未改",
        "获奖月任职万科（上海）实业发展有限公司（2018-01至2019-03），与“万科”相符。第6页项目起始写2019-11-15，晚于获奖日，项目行另标待核查，本行不改。",
    )
    add("3", "处分", "无", "未改", "核对后未改", "无时间信息。")
    add(
        "4",
        "继续教育三行",
        "2025-05-03至2025-06-21，南京市人社局",
        "未改",
        "核对后未改",
        "培训区间落在江苏省建筑装饰设计研究院有限公司任职（2025-03至2025-12，南京）内。",
    )

    for index, old in enumerate(OLD_PAGE5, start=1):
        add(
            "5",
            f"原工作经历第{index}行",
            old,
            "整表改为下方18段，本行不再保留原起止",
            "已改正文",
            "原8行把多家单位并进错误年份，且顺序与一览表不符。不在错误年份上只改一两个数字。",
        )
    for index, job in enumerate(JOBS, start=1):
        add(
            "5",
            f"工作经历新第{index}行（新到旧）",
            "（原表无此完整一行，或单位/起止不同）",
            job_label(job),
            "已改正文",
            "来自一览表。部门均为工程部。岗位栏一览表没有：管理类职务写管理岗；睿泽的资料员、施工员写技术岗。主要工作内容写入原为空的“从事的主要专业技术工作”栏。",
        )
    add(
        "5",
        "表下附注",
        "无",
        PAGE5_NOTE,
        "已改正文",
        "说明不补日、不补两处空档，避免把空档理解成漏填。",
    )
    add(
        "5",
        "未写入履历的原单位",
        "上海万限实业有限公司；上海万科企业有限公司；金科地产集团有限公司华东区域公司；江苏中南建设集团股份有限公司；括注“融创文旅华东区域”",
        "从工作经历和总结履历中删除",
        "已改正文",
        "一览表没有这些单位。万限仅是原表对万科（上海）实业发展有限公司的括注，改用一览表单位全称。不另造任职段。",
    )

    add(
        "6",
        "杭州湾融创文旅城项目",
        "2018年02月03日至2023年05月11日；主持；奖项无",
        "起止数字未改，日期栏下加“待核查”",
        "保留待核查",
        "一览表中名称最接近的是上海融之旅，但任职只有2021-03至2021-08，与本行起止不重合。不把项目改写到2021年。同项目在总结写2022.04-2023.05，两处也不一致。",
    )
    add(
        "6",
        "万科青浦天空之城项目",
        "2019年11月15日至2022年03月25日；奖项写可持续社区营造标杆项目奖；专利ZL202020536480.4、ZL202020557884.1",
        "起止、奖项句子和专利号未改，日期栏下加“待核查”",
        "保留待核查",
        "万科任职只到2019-03（实业公司）。奖惩表获奖日是2019-02-24，早于本行起始2019-11-15。专利号年份为2020，2020年任职已不是万科。不把专利改挂到万居、金科或欣驭。",
    )
    add(
        "6",
        "绿地苏州东太湖项目",
        "2020年10月15日至2024年01月25日；奖项写城市更新卓越实践奖",
        "起止和奖项句子未改，日期栏下加“待核查”",
        "保留待核查",
        "徐汇绿地任职是2023-10至2024-07。本行起始早于该任职，截止2024-01-25又早于奖惩表获奖日2024-04-26。总结另写2023.06-2025.03。三处不一致，不裁剪到交集，不改挂奖项。",
    )
    add("6", "表下附注", "无", PAGE6_NOTE, "已改正文", "说明只标注、不改归属。")

    add(
        "7",
        "论文三篇",
        "2021-07-08、2021-08-05、2024-07-10",
        "未改",
        "核对后未改",
        "均在工程师聘用日2018-12-25之后。论文未写任职单位。与第11页标题用字不同，但日期一致，不改标题。",
    )

    add(
        "8",
        "总结中的职业履历",
        "7行，含2025.03-至今、万科企业、金科地产华东、中南建设2011.07-2015.05",
        "改为与第5页相同的18段；装饰院止于2025.12，不再写至今",
        "已改正文",
        "原表把多家单位合并，且“至今”与一览表不符。一览表没有的单位从本栏删除。",
    )
    add("8", "履历附注", "无", PAGE8_NOTE, "已改正文", "与第5页同一口径。")
    add(
        "8",
        "教育背景两句",
        "2011年6月海洋大学毕业；2021年1月复旦大学硕士",
        "未改",
        "核对后未改",
        "与第3页学历栏一致，不是任职。",
    )
    add(
        "9",
        "社会职务",
        "复旦住房政策研究中心秘书长；杨浦区科技企业联合会执行会长",
        "未改",
        "核对后未改",
        "没有起止时间，不能据任职表改写。",
    )
    add(
        "9",
        "中南海湾新城·金石国际广场",
        "2011.07-2015.08；成果含江苏省优质工程奖",
        "标题日期后加“待核查”；事迹文字未改",
        "保留待核查",
        "一览表无2015-09之前的任职，不能把本项目改挂到睿泽（2015-09至2015-10，资料员、施工员）。江苏省优质工程奖不在奖惩表，不新造奖项行。",
    )
    add(
        "9",
        "新城香溢紫郡",
        "2015.12-2017.02，工程高级经理",
        "未改",
        "核对后未改",
        "起止与新城控股集团股份有限公司上海第一分公司（2015-12至2017-02，高级经理）一致，项目名含“新城”。不改写。",
    )
    add(
        "9",
        "合肥金科博翠天辰",
        "2017.03-2017.10，隶属金科集团华东区域公司",
        "标题和隶属句加“待核查”；年份未改",
        "保留待核查",
        "2017-03至2017-12的任职是上海万科长宁置业有限公司。金科两家单位在2020-03至2020-05、2020-06至2020-10，且均为上海单位。不把合肥项目改挂到万科，也不改到2020年。",
    )
    add(
        "10",
        "上海万科天空之城",
        "2017.11-2022.02，工程总监",
        "标题日期后加“待核查”",
        "保留待核查",
        "与第6页2019-11-15至2022-03-25不一致，且长于万科任职（止于2019-03）。不取其中一段当作项目期。",
    )
    add(
        "10",
        "杭州湾融创文旅城水路乐园",
        "2022.04-2023.05，工程总监",
        "标题日期后加“待核查”",
        "保留待核查",
        "与第6页2018-02-03至2023-05-11不一致。融之旅任职2021-03至2021-08，与这两段都不重合。",
    )
    add(
        "10",
        "苏州绿地江南映月花园",
        "2023.06-2025.03，工程部副总经理",
        "标题日期后加“待核查”",
        "保留待核查",
        "徐汇绿地任职是2023-10至2024-07。本行起早、止晚，且与第6页截止2024-01-25不一致。不裁成交集。",
    )
    add(
        "10",
        "东台市养老服务中心地下车库装饰工程",
        "2025.3-至今，施工技术负责人",
        "改为“2025.03起，待核查”",
        "保留待核查",
        "一览表无项目名称。装饰院任职止于2025-12，其后是宝龙、贝迪，不能再写至今。不把截止写成2025-12，避免把项目说成恰好做到离职月。",
    )
    add(
        "11",
        "两项专利",
        "ZL202020536480.4、ZL202020557884.1",
        "未改",
        "核对后未改",
        "此处未写任职单位。2020年的任职单位不是万科。不改挂，第6页的项目归属保持待核查。",
    )
    add(
        "11",
        "学术交流",
        "2022.07上海市建筑施工技术交流会；2022.11江苏省建筑技术研讨会",
        "未改",
        "核对后未改",
        "2022-07在社通任职末月，2022-11在苏科任职内。发言未写由哪家单位派出，不补单位。",
    )
    add(
        "11",
        "总结内论文表",
        "2021.07、2021.08、2024.07",
        "未改",
        "核对后未改",
        "日期与第7页一致。标题用字与第7页不同，不属于时间冲突，不改。",
    )
    add(
        "12",
        "考核情况",
        "无",
        "未改",
        "核对后未改",
        "没有可对照的考核年份。",
    )
    add(
        "13",
        "核实栏日期",
        "2025年08月12日",
        "未改",
        "核对后未改",
        "流程日期，与填表日相同，当时任职装饰院。",
    )
    add(
        "14",
        "评议、评审、备案日期及备注单位",
        "2025-09-21、2025-10-19、2025-12-09；工作单位为江苏省建筑装饰设计研究院有限公司",
        "未改",
        "核对后未改",
        "三个日期都在装饰院任职截止月2025-12之内。2026年宝龙、贝迪两段发生在备案之后，只写入工作经历，不回改评审日期。",
    )
    return rows


def style_xlsx(rows: list[dict]) -> None:
    wb = openpyxl.Workbook()
    intro = wb.active
    intro.title = "填写说明"
    intro["A1"] = "申报表任职时间校正说明"
    intro["A1"].font = Font(name="宋体", size=16, bold=True)
    notes = [
        "依据：本轮上传的《工作经历一览表.xlsx》是任职时间与单位的唯一依据。申报表原文只作对照，不反向覆盖一览表。",
        "一览表共18段。原文写明未覆盖、不作补写的空档是2015-11、2023-09。校正时不补这两个月，也不把年月补成具体日。",
        "工作经历、总结履历中的单位和起止，按一览表改正。原表有而一览表没有的单位，从履历删除，不另造一段。",
        "奖项和项目：只有起止、角色与一览表中的同一单位能够对应时才保留原句。对不上的保留原文，正文标“待核查”，不为了对齐年份改挂到别的单位。",
        "奖惩表上的三个获奖日期本身都落在某段任职内，因此奖惩表三行不改。把这些奖写进某一项目、且项目起止对不上的，只在项目行标待核查。",
        "封面单位、填表时间和第14页评审备案日期不改。2025-08-12至2025-12-09期间，一览表单位是江苏省建筑装饰设计研究院有限公司。",
        "2026-01至2026-04上海宝龙、2026-05至今上海贝迪，按一览表写入工作经历。这两段晚于备案日2025-12-09，评审日期不回改。",
        "专业技术工作年限改为9年：2015年09月至填表日2025年08月，实足9年11个月，取已满年份。",
        "第5页原表只有8行，装不下18段。原有8行栏线未改，在表下空白处按同一栏宽续了10行。水印、页数和其他页的表格未改。",
        "岗位栏一览表没有。总监、副总经理、高级经理、高级工程师沿用原表“管理岗”；资料员、施工员写作“技术岗”，避免记成管理岗。",
        "处理分三类：已改正文、保留待核查、核对后未改。待核查项必须仍能在正文里看到，不能被删掉或改挂。",
    ]
    for index, line in enumerate(notes, start=3):
        intro.cell(index, 1, line).font = Font(name="宋体", size=11)
        intro.cell(index, 1).alignment = Alignment(wrap_text=True, vertical="center")
        intro.row_dimensions[index].height = 36
    intro.column_dimensions["A"].width = 108
    intro.row_dimensions[1].height = 24
    intro.page_setup.orientation = "landscape"
    intro.page_setup.fitToPage = True
    intro.page_setup.fitToWidth = 1
    intro.page_setup.fitToHeight = 1
    intro.sheet_properties.pageSetUpPr.fitToPage = True
    intro.page_setup.paperSize = intro.PAPERSIZE_A4
    intro.print_title_rows = "1:1"
    intro.page_setup.horizontalCentered = True
    intro.sheet_view.showGridLines = False
    intro.oddHeader.left.text = "任职时间校正说明"
    intro.oddFooter.center.text = "第 &P 页"

    ws = wb.create_sheet("逐项对照", 0)
    headers = ["序号", "页", "栏目", "原文", "修改后", "处理", "依据与说明"]
    ws.append(headers)
    fills = {
        "已改正文": PatternFill("solid", fgColor="E2EFDA"),
        "保留待核查": PatternFill("solid", fgColor="FFF2CC"),
        "核对后未改": PatternFill("solid", fgColor="F2F2F2"),
    }
    thin = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )
    for index, row in enumerate(rows, start=1):
        ws.append(
            [index, row["页"], row["栏目"], row["原文"], row["修改后"], row["处理"], row["依据与说明"]]
        )
        for col in range(1, 8):
            cell = ws.cell(index + 1, col)
            cell.font = Font(name="宋体", size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            cell.border = thin
            if col == 6:
                cell.fill = fills.get(row["处理"], PatternFill())
        ws.row_dimensions[index + 1].height = 48
    for col, width in enumerate([6, 8, 28, 46, 46, 14, 62], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    for col in range(1, 8):
        cell = ws.cell(1, col)
        cell.font = Font(name="宋体", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E79")
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin
    ws.row_dimensions[1].height = 22
    ws.auto_filter.ref = f"A1:G{len(rows) + 1}"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{len(rows) + 1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.horizontalCentered = True
    ws.print_title_rows = "1:1"
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.6
    ws.page_margins.bottom = 0.5
    ws.oddHeader.left.text = "申报表任职时间逐项对照"
    ws.oddFooter.center.text = "第 &P 页 / 共 &N 页"
    ws.sheet_view.zoomScale = 110
    # 待核查单独一页，方便核对没有被改挂
    pending = wb.create_sheet("待核查清单", 1)
    pending.append(["序号", "页", "栏目", "原文如何保留", "为何不能改挂"])
    n = 1
    for row in rows:
        if row["处理"] != "保留待核查":
            continue
        pending.append([n, row["页"], row["栏目"], row["修改后"], row["依据与说明"]])
        n += 1
    for col, width in enumerate([6, 8, 32, 55, 70], start=1):
        pending.column_dimensions[get_column_letter(col)].width = width
    for col in range(1, 6):
        cell = pending.cell(1, col)
        cell.font = Font(name="宋体", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="BF8F00")
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    for r in range(2, pending.max_row + 1):
        for col in range(1, 6):
            cell = pending.cell(r, col)
            cell.font = Font(name="宋体", size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            cell.fill = PatternFill("solid", fgColor="FFF2CC")
        pending.row_dimensions[r].height = 48
    pending.freeze_panes = "A2"
    pending.auto_filter.ref = f"A1:E{pending.max_row}"
    pending.page_setup.orientation = "landscape"
    pending.page_setup.fitToPage = True
    pending.page_setup.fitToWidth = 1
    pending.page_setup.fitToHeight = 1
    pending.sheet_properties.pageSetUpPr.fitToPage = True
    pending.page_setup.paperSize = pending.PAPERSIZE_A4
    pending.print_title_rows = "1:1"
    pending.oddHeader.left.text = "待核查清单（未改挂）"
    pending.oddFooter.center.text = "第 &P 页"
    pending.page_margins.left = 0.4
    pending.page_margins.right = 0.4

    # 18段一览，便于和正文逐行对
    roster = wb.create_sheet("一览表18段", 2)
    roster.append(["序号", "起止", "单位", "职务", "主要工作内容", "写入申报表的位置"])
    for index, job in enumerate(JOBS, start=1):
        start, end, company, _dept, _post, work, title = job
        roster.append([index, f"{start}至{end}", company, title, work, "第5页工作经历；第8页总结履历"])
    for col, width in enumerate([6, 28, 46, 18, 24, 36], start=1):
        roster.column_dimensions[get_column_letter(col)].width = width
    for col in range(1, 7):
        cell = roster.cell(1, col)
        cell.font = Font(name="宋体", bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E79")
        cell.alignment = Alignment(horizontal="center")
    for r in range(2, 20):
        for col in range(1, 7):
            roster.cell(r, col).font = Font(name="宋体", size=10)
            roster.cell(r, col).alignment = Alignment(wrap_text=True, vertical="center")
        roster.row_dimensions[r].height = 20
    roster.page_setup.orientation = "landscape"
    roster.page_setup.fitToPage = True
    roster.page_setup.fitToWidth = 1
    roster.page_setup.fitToHeight = 1
    roster.sheet_properties.pageSetUpPr.fitToPage = True
    roster.page_setup.paperSize = roster.PAPERSIZE_A4
    roster.print_title_rows = "1:1"
    roster.oddFooter.center.text = "第 &P 页"
    roster.freeze_panes = "A2"

    wb.save(OUT_XLSX)


def assert_result(doc: pymupdf.Document) -> None:
    if doc.page_count != 14:
        raise SystemExit(f"页数变为 {doc.page_count}")
    texts = [doc[i].get_text() for i in range(14)]
    for index, text in enumerate(texts, start=1):
        if "一体化" not in text:
            raise SystemExit(f"第{index}页水印丢失")
    page3, page5, page6, page8, page9, page10 = texts[2], texts[4], texts[5], texts[7], texts[8], texts[9]
    flat5, flat8 = page5.replace("\n", ""), page8.replace("\n", "")
    if "2011年06月27日" in page3 or "专业技术工作年限：9年" not in page3:
        raise SystemExit("第3页参加工作或年限未改对")
    if "2015年09月" not in page3:
        raise SystemExit("第3页缺2015年09月")
    for job in JOBS:
        if job[2] not in flat5 or job[2] not in flat8:
            raise SystemExit(f"履历缺少单位：{job[2]}")
    for bad in ["万限", "中南建设", "万科企业有限公司", "金科地产", "资深经理", "融创文旅", "2025年07月10日", "2018年12月10日"]:
        if bad in page5:
            raise SystemExit(f"第5页仍有旧表述：{bad}")
    if "至今" in page8 or "中南" in page8 or "万科企业" in page8:
        raise SystemExit("第8页仍有旧履历")
    if page6.count("待核查") < 3 or PAGE6_NOTE not in page6:
        raise SystemExit("第6页待核查标注不完整")
    if "2019年11月" not in page6 or "ZL 2020 2 0536480.4" not in page6:
        raise SystemExit("第6页项目原文被删")
    if "新城香溢紫郡（2015.12-2017.02）" not in page9:
        raise SystemExit("新城香溢紫郡被误改")
    if "2011.07-2015.08，待核查" not in page9 or "金科集团华东区域公司，待核查" not in page9:
        raise SystemExit("第9页待核查未写入")
    for snippet in ["2017.11-2022.02，待核查", "2022.04-2023.05，待核查", "2023.06-2025.03，待核查", "2025.03起，待核查"]:
        if snippet not in page10:
            raise SystemExit(f"第10页缺少：{snippet}")
    if "至今" in page10:
        raise SystemExit("第10页仍有至今")
    if "2024-12-18" not in page3 or "2019-02-24" not in page3 or "2024-04-26" not in page3:
        raise SystemExit("奖惩日期被改动")
    if "2021-07-08" not in texts[6] or "2025 年12 月09 日" not in texts[13]:
        raise SystemExit("论文或备案日期异常")
    # 表格右侧空白处的水印应仍是浅灰。透明度丢失时会变成接近黑色。
    pix = doc[4].get_pixmap(matrix=pymupdf.Matrix(1, 1), alpha=False)
    dark = 0
    light = 0
    for y in range(136, 146):
        for x in range(530, 542):
            red = pix.pixel(x, y)[0]
            if red < 180:
                dark += 1
            elif red < 252:
                light += 1
    if dark or light < 8:
        raise SystemExit(f"第5页水印深浅异常 dark={dark} light={light}")


def render_checks(doc: pymupdf.Document) -> None:
    ART_DIR.mkdir(parents=True, exist_ok=True)
    # 关键页：基本情况、工作经历、业绩、总结履历、项目上、项目下
    for page_no in (3, 5, 6, 8, 9, 10):
        pix = doc[page_no - 1].get_pixmap(matrix=pymupdf.Matrix(1.7, 1.7), alpha=False)
        pix.save(ART_DIR / f"page-{page_no:02d}.png")
    # 工作经历下半，便于看新增行和附注
    clip = pymupdf.Rect(40, 400, 560, 830)
    pix = doc[4].get_pixmap(matrix=pymupdf.Matrix(1.8, 1.8), clip=clip, alpha=False)
    pix.save(ART_DIR / "page-05-lower.png")
    pix = doc[7].get_pixmap(matrix=pymupdf.Matrix(1.8, 1.8), clip=pymupdf.Rect(40, 360, 560, 820), alpha=False)
    pix.save(ART_DIR / "page-08-career.png")


def main() -> None:
    chars = collect_chars()
    build_subset(chars)
    assert_font_readable("2015年09月至上海贝迪创建科技有限公司资料员、施工员待核查空档")
    font = pymupdf.Font(fontfile=str(FONT_SUBSET))
    missing = sorted({ch for ch in chars if font.has_glyph(ord(ch)) == 0})
    if missing:
        raise SystemExit(f"子集缺字：{''.join(missing)}")

    doc = pymupdf.open(SRC_PDF)
    plan = [
        (2, edit_page3),
        (4, edit_page5),
        (5, edit_page6),
        (7, edit_page8),
        (8, lambda page, font: edit_summary_lines(page, font, PAGE9_LINES)),
        (9, lambda page, font: edit_summary_lines(page, font, PAGE10_LINES)),
    ]
    for index, editor in plan:
        page, water, gs = prepare_page(doc, index)
        editor(page, font)
        finish_page(doc, page, water, gs)

    assert_result(doc)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc.save(OUT_PDF, garbage=4, deflate=True)
    saved = pymupdf.open(OUT_PDF)
    assert_result(saved)
    render_checks(saved)
    rows = build_rows()
    style_xlsx(rows)
    print(f"PDF {OUT_PDF} bytes {OUT_PDF.stat().st_size}")
    print(f"XLSX {OUT_XLSX}")
    print(f"对照 {len(rows)} 条，待核查 {sum(r['处理']=='保留待核查' for r in rows)} 条")


if __name__ == "__main__":
    main()
