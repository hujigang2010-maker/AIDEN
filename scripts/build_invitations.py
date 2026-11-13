#!/usr/bin/env python3
"""按最优议程为每位拟邀嘉宾生成邀请函。所附议程不出现收信人本人。"""

import importlib.util
import shutil
import subprocess
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "deliverables"
LETTER_DIR = OUT / "邀请函"
FONT = "微软雅黑"
PINE = "0F5C58"
GOLD = "8C6A32"
INK = "1B3A4B"
MUTED = "5E6A66"
PAPER = "F7F5F0"
ALT = "F4F8F7"
LINE = "D5DDD8"
PINE_RGB = (0x0F, 0x5C, 0x58)
GOLD_RGB = (0x8C, 0x6A, 0x32)
INK_RGB = (0x1B, 0x3A, 0x4B)
WHITE = (255, 255, 255)
CONTENT_W = 17.6

DATE_LINE = "2026年11月28日（周六）12:30—17:30"
PLACE = "上海展览中心中央大厅主论坛区（一楼）（上海市静安区延安中路1000号）"
SIGN_DATE = "二〇二六年十一月十三日"
MARKER = "附：活动议程"


def load_agenda():
    path = ROOT / "scripts" / "build_optimal_agenda.py"
    spec = importlib.util.spec_from_file_location("optimal_agenda", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


AGENDA = load_agenda()

LABELS = {
    "彭希哲": "彭希哲（复旦大学老龄研究院院长、文科资深教授）",
    "上海辰山国家植物园代表": "上海辰山国家植物园",
    "朱震": "朱震（复旦大学国家科技园）",
    "上海视觉艺术学院代表": "上海视觉艺术学院",
    "顾捷": "顾捷（傅利叶智能创始人、CEO）",
    "王兴兴": "王兴兴（宇树科技创始人、董事长）",
    "邓泰华": "邓泰华（智元机器人创始人、董事长兼CEO）",
    "邱锡鹏": "邱锡鹏（复旦大学教授）",
    "徐立": "徐立（商汤科技董事长兼CEO）",
    "韩璧丞": "韩璧丞（强脑科技创始人、CEO）",
    "赵博誉": "赵博誉（主持）",
    "胡继刚": "胡继刚（主持）",
    "朱翊元": "朱翊元（示范区新发展公司，主持）",
    "吴玉韶": "吴玉韶（复旦大学老龄研究院，主持）",
    "邱建伟": "邱建伟（泰康之家，主持）",
    "封进": "封进（复旦大学，主持）",
    "彭志辉": "彭志辉（智元机器人）",
    "朱秋国": "朱秋国（云深处科技）",
    "易港": "易港（钛虎机器人）",
    "钱东奇": "钱东奇（科沃斯集团）",
    "钱程": "钱程（科沃斯集团）",
    "傅利叶康复业务拟邀负责人": "傅利叶智能康复业务",
    "擎朗智能拟邀业务负责人": "擎朗智能",
    "胡湛": "胡湛（复旦大学老龄研究院）",
    "肖仰华": "肖仰华（复旦大学）",
    "王学集": "王学集（涂鸦智能）",
    "张强": "张强（联影医疗）",
    "沈定刚": "沈定刚（联影智能）",
    "钱鹏鹤": "钱鹏鹤（移远通信）",
    "张瑞安": "张瑞安（乐鑫科技）",
    "赵一洪": "赵一洪（福寿康）",
    "康剑": "康剑（福寿康）",
    "瞿文": "瞿文（福苑康养）",
    "陈楠": "陈楠（博斯腾科技）",
    "九如城拟邀负责人": "九如城",
    "第一社会福利院拟邀负责人": "上海市第一社会福利院",
    "康复医院拟邀负责人": "康复医院",
    "顾建忠": "顾建忠（上海银行董事长）",
    "工商银行上海市分行拟邀负责人": "工商银行上海市分行",
    "优势资本拟邀投资合伙人": "优势资本",
    "启明创投拟邀负责人": "启明创投",
    "上海国投拟邀负责人": "上海国投",
    "张江科投拟邀负责人": "张江科投",
    "专项基金拟邀负责人": "机器人及具身智能专项产业基金",
    "尚海龙": "尚海龙（商汤医疗）",
    "萤石网络拟邀负责人": "萤石网络",
    "海康威视拟邀负责人": "海康威视",
    "大华股份拟邀负责人": "大华股份",
    "复星康养拟邀负责人": "复星康养",
    "申园拟邀负责人": "泰康之家·申园",
    "交通银行拟邀负责人": "交通银行",
    "浦发银行拟邀负责人": "浦发银行",
    "杭州银行拟邀负责人": "杭州银行上海分行",
    "泰隆银行拟邀负责人": "泰隆银行",
    "养老险机构拟邀负责人": "太平养老、平安养老",
    "浦东创投拟邀负责人": "浦东创投",
    "临港科创投拟邀负责人": "临港科创投",
    "复星投资拟邀负责人": "复星投资",
    "长三角基金拟邀负责人": "长三角医疗科技及智能制造产业基金",
    "徐洁平": "徐洁平",
}

MARKS = {
    "上海辰山国家植物园代表": ("辰山", "上海辰山国家植物园代表"),
    "上海视觉艺术学院代表": ("视觉艺术学院", "上海视觉艺术学院代表"),
    "傅利叶康复业务拟邀负责人": ("傅利叶智能康复业务", "傅利叶康复业务拟邀负责人"),
    "擎朗智能拟邀业务负责人": ("擎朗智能", "擎朗智能拟邀业务负责人"),
    "九如城拟邀负责人": ("九如城",),
    "第一社会福利院拟邀负责人": ("第一社会福利院",),
    "康复医院拟邀负责人": ("康复医院",),
    "工商银行上海市分行拟邀负责人": ("工商银行上海市分行",),
    "优势资本拟邀投资合伙人": ("优势资本",),
    "启明创投拟邀负责人": ("启明创投",),
    "上海国投拟邀负责人": ("上海国投",),
    "张江科投拟邀负责人": ("张江科投",),
    "专项基金拟邀负责人": ("专项产业基金", "专项基金拟邀负责人"),
    "萤石网络拟邀负责人": ("萤石网络",),
    "海康威视拟邀负责人": ("海康威视",),
    "大华股份拟邀负责人": ("大华股份",),
    "复星康养拟邀负责人": ("复星康养",),
    "申园拟邀负责人": ("泰康之家·申园", "申园拟邀负责人"),
    "交通银行拟邀负责人": ("交通银行",),
    "浦发银行拟邀负责人": ("浦发银行",),
    "杭州银行拟邀负责人": ("杭州银行",),
    "泰隆银行拟邀负责人": ("泰隆银行",),
    "养老险机构拟邀负责人": ("太平养老", "平安养老"),
    "浦东创投拟邀负责人": ("浦东创投",),
    "临港科创投拟邀负责人": ("临港科创投",),
    "复星投资拟邀负责人": ("复星投资",),
    "长三角基金拟邀负责人": ("长三角医疗科技及智能制造产业基金",),
}

NAMED_KEYNOTES = ("彭希哲", "朱震", "顾捷", "王兴兴", "邓泰华", "邱锡鹏", "徐立", "韩璧丞")


def is_unnamed(name: str) -> bool:
    return "拟邀" in name or name.endswith("代表") or name.endswith("负责人")


@dataclass
class Slot:
    time: str
    place: str
    action: str
    topic: str


@dataclass
class Invitee:
    name: str
    address: str
    is_person: bool
    duty: str
    ask: str
    label: str
    marks: tuple[str, ...]
    card: list[tuple[str, str]] = field(default_factory=list)
    filename: str = ""


def panel_place(index: int) -> str:
    title = AGENDA.PANELS[index]["title"].replace("　", "「", 1) + "」"
    return title


def host_action(topic: str) -> str:
    if topic.startswith("开题"):
        return "开题"
    if topic.startswith("收束"):
        return "收束"
    return "主持"


def build_ask(person: bool, slots: list[Slot]) -> str:
    who = "您" if person else "贵单位代表"
    invite = "诚邀您" if person else "诚邀贵单位委派代表"
    kinds = {slot.action for slot in slots}
    if kinds <= {"开幕"}:
        return f"{invite}担任开幕主持，负责开场、串场与控时。开幕时段为{slots[0].time}。"
    if kinds <= {"主旨"}:
        topic = slots[0].topic
        shown = topic if topic.startswith("《") else f"《{topic}》"
        return f"{invite}作主旨分享。拟定时段为{slots[0].time}，题目为{shown}。"
    if kinds <= {"交流"}:
        return f"{invite}出席峰会，并参加{slots[0].time}茶歇分组交流。交流主题为{slots[0].topic}。本环节不逐一上台。"
    if kinds <= {"发言"}:
        slot = slots[0]
        return f"{invite}出席{slot.place}并发言。拟定时段为{slot.time}，讨论要点为{slot.topic}。"
    parts = []
    for slot in slots:
        parts.append(f"于{slot.time}为{slot.place}{slot.action}")
    joined = "，并于".join(parts) if len(parts) == 2 else "；".join(parts)
    return f"诚邀{who}担任圆桌主持。请{joined}。"


def build_card(slots: list[Slot]) -> list[tuple[str, str]]:
    actions = {slot.action for slot in slots}
    if actions <= {"开幕"}:
        return [("拟请担任", "开幕主持"), ("主持时段", slots[0].time), ("主持任务", "开场、串场与控时")]
    if actions <= {"主旨"}:
        topic = slots[0].topic
        shown = topic if topic.startswith("《") else f"《{topic}》"
        return [("拟请担任", "主旨分享"), ("发言时段", slots[0].time), ("拟定题目", shown)]
    if actions <= {"交流"}:
        return [("拟请担任", "茶歇分组交流"), ("交流时段", slots[0].time), ("交流主题", slots[0].topic)]
    if actions <= {"发言"}:
        return [
            ("拟请担任", slots[0].place),
            ("发言时段", slots[0].time),
            ("讨论要点", slots[0].topic),
        ]
    rows = [("拟请担任", "圆桌主持")]
    for slot in slots:
        short, _, theme = slot.place.partition("「")
        rows.append((f"{short}{slot.action}", f"{slot.time}　{theme.rstrip('」')}"))
    return rows


def duty_of(slots: list[Slot]) -> str:
    actions = {slot.action for slot in slots}
    if actions <= {"开幕"}:
        return "开幕主持"
    if actions <= {"主旨"}:
        return "主旨分享"
    if actions <= {"交流"}:
        return "茶歇交流"
    if actions <= {"发言"}:
        return slots[0].place.split("「")[0] + "嘉宾"
    places = []
    for slot in slots:
        short = slot.place.split("「")[0]
        if short not in places:
            places.append(short)
    return "、".join(places) + "主持"


def collect() -> list[Invitee]:
    grouped: dict[str, list[Slot]] = {}
    order: list[str] = []

    def add(name: str, slot: Slot) -> None:
        if name not in grouped:
            grouped[name] = []
            order.append(name)
        grouped[name].append(slot)

    add("徐洁平", Slot("13:00—13:06", "开幕", "开幕", "开场、串场与控时"))
    for time, name, _org, topic in AGENDA.KEYNOTES:
        add(name, Slot(time, "主旨分享", "主旨", topic))
    for index, panel in enumerate(AGENDA.PANELS):
        place = panel_place(index)
        for time, name, org, topic in panel["rows"]:
            if "主持" in org:
                add(name, Slot(time, place, host_action(topic), topic))
            else:
                add(name, Slot(time, place, "发言", topic))
    for name, _org, topic in AGENDA.EXCHANGE:
        add(name, Slot("14:34—14:42", "茶歇分组交流", "交流", topic))

    invitees = []
    for name in order:
        slots = grouped[name]
        person = not is_unnamed(name)
        label = LABELS[name]
        address = name if person else label
        marks = (name,) if person else MARKS[name]
        item = Invitee(
            name=name,
            address=address,
            is_person=person,
            duty=duty_of(slots),
            ask=build_ask(person, slots),
            label=label,
            marks=marks,
            card=build_card(slots),
            filename=f"邀请函_{address}.docx",
        )
        invitees.append(item)
    return invitees


def atoms() -> dict[str, list[tuple[str, str]]]:
    """各环节可写入议程的嘉宾。元组为（名册键，对外称呼）。"""
    opening = [("徐洁平", LABELS["徐洁平"])]
    keynotes = [(name, LABELS[name]) for _time, name, _org, _topic in AGENDA.KEYNOTES]
    panels = []
    for panel in AGENDA.PANELS:
        row_atoms = []
        for _time, name, org, _topic in panel["rows"]:
            label = LABELS[name]
            if "主持" in org and "主持" not in label:
                label += "（主持）"
            row_atoms.append((name, label))
        panels.append(row_atoms)
    exchange = [(name, LABELS[name]) for name, _org, _topic in AGENDA.EXCHANGE]
    return {"opening": opening, "keynotes": keynotes, "panels": panels, "exchange": exchange}


ATOMS = atoms()


def visible(pairs: list[tuple[str, str]], invitee: Invitee) -> str:
    kept = [label for key, label in pairs if key != invitee.name]
    return "、".join(kept)


def agenda_rows(invitee: Invitee) -> list[tuple[str, str, str]]:
    opening = "专业主持人" if invitee.name == "徐洁平" else "徐洁平"
    panels = [visible(pairs, invitee) for pairs in ATOMS["panels"]]
    return [
        ("12:30—13:00", "签到入场", "嘉宾与观众入场"),
        ("13:00—13:06", "开幕", opening),
        ("13:06—13:14", "致辞", "指导单位代表"),
        ("13:14—14:34", "主旨分享", visible(ATOMS["keynotes"], invitee)),
        ("14:34—14:42", "茶歇分组交流", visible(ATOMS["exchange"], invitee)),
        ("14:42—15:22", "圆桌一\n具身智能进院与进家", panels[0]),
        ("15:22—16:02", "圆桌二\n数智康养、医疗AI与智慧物联", panels[1]),
        ("16:02—16:42", "圆桌三\n机构采购、试点与运营", panels[2]),
        ("16:42—17:22", "圆桌四\n产业资本、银行与支付", panels[3]),
        ("17:22—17:30", "总结合影", "主办、承办方代表"),
    ]


def check_roster(invitees: list[Invitee]) -> None:
    labels = [item.label for item in invitees]
    assert len(labels) == len(set(labels)), "对外称呼有重复"
    assert len(invitees) == 60, len(invitees)
    for item in invitees:
        for other in invitees:
            if other.name == item.name:
                continue
            for mark in item.marks:
                assert mark not in other.label, (item.address, mark, other.label)
        rows = agenda_rows(item)
        agenda = "\n".join(cell for row in rows for cell in row)
        for mark in item.marks:
            assert mark not in agenda, (item.address, mark)
        if item.name in NAMED_KEYNOTES:
            others = [name for name in NAMED_KEYNOTES if name != item.name]
            missing = [name for name in others if name not in agenda]
            assert not missing, (item.address, missing)
        if item.name == "王兴兴":
            assert "王兴兴" not in agenda
            assert "宇树" not in agenda
            for name in ("彭希哲", "顾捷", "邓泰华", "邱锡鹏", "徐立", "韩璧丞", "朱震"):
                assert name in agenda
        if item.name == "彭志辉":
            assert "彭志辉" not in agenda
            assert "王兴兴" in agenda and "朱秋国" in agenda
        if item.name == "赵博誉":
            assert "赵博誉" not in agenda
            assert "朱翊元" in agenda and "邱建伟" in agenda
        if item.name == "上海辰山国家植物园代表":
            assert "辰山" not in agenda
            assert "王兴兴" in agenda and "彭希哲" in agenda


def set_run_font(run, size: float, bold: bool = False, color: tuple | None = None, spacing: int | None = None) -> None:
    run.font.name = FONT
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        r_fonts.set(qn(attr), FONT)
    if spacing is not None:
        node = OxmlElement("w:spacing")
        node.set(qn("w:val"), str(spacing))
        r_pr.append(node)


def shade(cell, color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    node = OxmlElement("w:shd")
    node.set(qn("w:val"), "clear")
    node.set(qn("w:color"), "auto")
    node.set(qn("w:fill"), color)
    tc_pr.append(node)


def set_cell_margins(cell, top="40", left="80", bottom="40", right="80") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = OxmlElement("w:tcMar")
    for edge, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:w"), value)
        node.set(qn("w:type"), "dxa")
        tc_mar.append(node)
    tc_pr.append(tc_mar)
    valign = OxmlElement("w:vAlign")
    valign.set(qn("w:val"), "center")
    tc_pr.append(valign)


def set_table_fixed(table, widths: list[float]) -> None:
    table.autofit = False
    table.allow_autofit = False
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    total = str(int(sum(widths) * 567))
    tbl_w.set(qn("w:w"), total)
    tbl_w.set(qn("w:type"), "dxa")
    indent = tbl_pr.find(qn("w:tblInd"))
    if indent is None:
        indent = OxmlElement("w:tblInd")
        tbl_pr.append(indent)
    indent.set(qn("w:w"), "0")
    indent.set(qn("w:type"), "dxa")
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tbl_pr.append(layout)
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tbl_pr.addnext(grid)
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(int(width * 567)))
        grid.append(col)
    for row in table.rows:
        tr_pr = row._tr.get_or_add_trPr()
        tr_pr.append(OxmlElement("w:cantSplit"))
        for index, cell in enumerate(row.cells):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(int(widths[index] * 567)))
            tc_w.set(qn("w:type"), "dxa")


def set_borders(table, color: str = LINE, sz: str = "4", val: str = "single") -> None:
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), val)
        node.set(qn("w:sz"), sz)
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), color)
        borders.append(node)
    table._tbl.tblPr.append(borders)


def set_cell_borders(cell, edges: dict[str, tuple[str, str, str]]) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        node = OxmlElement(f"w:{edge}")
        if edge in edges:
            val, sz, color = edges[edge]
            node.set(qn("w:val"), val)
            node.set(qn("w:sz"), sz)
            node.set(qn("w:space"), "0")
            node.set(qn("w:color"), color)
        else:
            node.set(qn("w:val"), "nil")
        borders.append(node)
    tc_pr.append(borders)


def keep_token(text: str) -> str:
    """人名和职衔整体换行，不断在字中间。"""
    return "\u2060".join(text)


def write_cell(cell, text: str, size: float = 9, bold: bool = False, color: tuple | None = None,
               fill: str | None = None, align: str = "left", keep_names: bool = False) -> None:
    cell.text = ""
    alignment = {
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    lines = text.split("\n")
    for line_index, line in enumerate(lines):
        paragraph = cell.paragraphs[0] if line_index == 0 else cell.add_paragraph()
        paragraph.alignment = alignment
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.08
        if keep_names and "、" in line:
            parts = line.split("、")
            for part_index, part in enumerate(parts):
                if part_index:
                    set_run_font(paragraph.add_run("、"), size, bold, color)
                set_run_font(paragraph.add_run(keep_token(part)), size, bold, color)
        else:
            set_run_font(paragraph.add_run(line), size, bold, color)
    if fill:
        shade(cell, fill)
    set_cell_margins(cell)


def add_line(doc, text: str, size: float = 12, bold: bool = False, color: tuple | None = None,
             align: str = "left", before: float = 0, after: float = 3, indent: bool = False,
             spacing: int | None = None):
    paragraph = doc.add_paragraph()
    paragraph.alignment = {
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = 1.35
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    if indent:
        paragraph.paragraph_format.first_line_indent = Cm(0.74)
    set_run_font(paragraph.add_run(text), size, bold, color, spacing)
    return paragraph


def gold_rule(paragraph) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), GOLD)
    borders.append(bottom)
    p_pr.append(borders)


def prepare(doc: Document, footer: str) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(10.5)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    r_pr = normal.element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        r_fonts.set(qn(attr), FONT)
    if doc.paragraphs:
        empty = doc.paragraphs[0]._element
        empty.getparent().remove(empty)

    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.55)
    section.bottom_margin = Cm(1.45)
    section.left_margin = Cm(1.7)
    section.right_margin = Cm(1.7)
    section.header_distance = Cm(0.45)
    section.footer_distance = Cm(0.35)
    footer_p = section.footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(footer_p.add_run(footer), 8, color=(0x6E, 0x78, 0x74))


def add_banner(doc: Document) -> None:
    table = doc.add_table(rows=2, cols=1)
    set_table_fixed(table, [CONTENT_W])
    set_borders(table, PINE, "2", "single")
    top = table.cell(0, 0)
    shade(top, PINE)
    set_cell_margins(top, "90", "120", "80", "120")
    top.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    top.paragraphs[0].paragraph_format.space_after = Pt(0)
    set_run_font(top.paragraphs[0].add_run("2026上海国际智慧康养产业博览会"), 11, color=WHITE, spacing=60)
    second = top.add_paragraph()
    second.alignment = WD_ALIGN_PARAGRAPH.CENTER
    second.paragraph_format.space_before = Pt(2)
    second.paragraph_format.space_after = Pt(0)
    set_run_font(second.add_run("AI重塑养老未来"), 16, True, (0xE6, 0xD3, 0xA3), spacing=180)
    bar = table.cell(1, 0)
    shade(bar, GOLD)
    set_cell_margins(bar, "0", "0", "0", "0")
    bar.paragraphs[0].paragraph_format.space_before = Pt(0)
    bar.paragraphs[0].paragraph_format.space_after = Pt(0)
    tr_pr = table.rows[1]._tr.get_or_add_trPr()
    height = OxmlElement("w:trHeight")
    height.set(qn("w:val"), "90")
    height.set(qn("w:hRule"), "exact")
    tr_pr.append(height)


def add_role_card(doc: Document, invitee: Invitee) -> None:
    table = doc.add_table(rows=1, cols=1)
    set_table_fixed(table, [CONTENT_W])
    set_borders(table, "E4D7C3", "4", "single")
    cell = table.cell(0, 0)
    shade(cell, PAPER)
    set_cell_borders(cell, {"left": ("single", "24", PINE), "top": ("single", "4", "E4D7C3"),
                            "bottom": ("single", "4", "E4D7C3"), "right": ("single", "4", "E4D7C3")})
    set_cell_margins(cell, "80", "140", "80", "140")
    cell.paragraphs[0].paragraph_format.space_after = Pt(0)
    first = True
    for label, value in invitee.card:
        paragraph = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        paragraph.paragraph_format.space_before = Pt(1)
        paragraph.paragraph_format.space_after = Pt(1)
        paragraph.paragraph_format.line_spacing = 1.2
        set_run_font(paragraph.add_run(label), 10.5, True, PINE_RGB)
        set_run_font(paragraph.add_run("　　" + value), 10.5, color=INK_RGB)


def render_letter(doc: Document, invitee: Invitee) -> None:
    add_banner(doc)
    title = add_line(doc, "邀　请　函", 26, True, PINE_RGB, "center", before=12, after=2, spacing=240)
    gold_rule(title)
    add_line(doc, "银发科技与具身智能产业创新峰会", 12, True, INK_RGB, "center", before=6, after=1)
    add_line(doc, DATE_LINE, 10, color=(0x5E, 0x6A, 0x66), align="center", before=2, after=0)
    add_line(doc, PLACE, 10, color=(0x5E, 0x6A, 0x66), align="center", before=0, after=8)

    add_line(doc, f"尊敬的{invitee.address}：", 14, True, INK_RGB, before=2, after=6)
    you = "您" if invitee.is_person else "贵单位"
    add_line(
        doc,
        f"2026上海国际智慧康养产业博览会将于{DATE_LINE}，在{PLACE}举办“2026银发科技与具身智能产业创新峰会”。",
        11, color=INK_RGB, indent=True, after=4,
    )
    add_line(
        doc,
        "本次峰会以“AI重塑养老未来”为主题。指导单位为上海现代服务业联合会、上海市养老服务和老龄产业协会、国家技术转移东部中心；主办单位为上海星智耀科技有限公司；承办单位为上海市杨浦区科技企业联合会、复旦大学住房政策研究中心；协办单位为长三角一体化示范区新发展建设有限公司；支持单位为上海辰山国家植物园、璞跃中国、上海视觉艺术学院；学术支持单位为复旦大学老龄研究院。",
        11, color=INK_RGB, indent=True, after=4,
    )
    add_line(
        doc,
        "当日安排签到入场、开幕、致辞、主旨分享、茶歇分组交流、四场圆桌及总结合影。四场圆桌依次讨论具身智能进院与进家，数智康养、医疗AI与智慧物联，机构采购、试点与运营，产业资本、银行与支付。",
        11, color=INK_RGB, indent=True, after=8,
    )
    add_role_card(doc, invitee)
    omit = f"不含{you}本人" if invitee.is_person else "不含贵单位在本函所请环节的代表"
    add_line(
        doc,
        f"附后为活动议程，只列同期出席的其他嘉宾，{omit}。大会最终议程以现场为准。",
        11, color=INK_RGB, indent=True, before=8, after=2,
    )
    add_line(doc, "敬请拨冗莅临。", 11, color=INK_RGB, indent=True, before=0, after=8)
    add_line(doc, "上海国际智慧康养产业博览会组委会", 11, True, INK_RGB, "right", before=6, after=0)
    add_line(doc, SIGN_DATE, 11, color=INK_RGB, align="right", before=0, after=0)

    doc.add_page_break()
    add_line(doc, "2026上海国际智慧康养产业博览会", 10, color=GOLD_RGB, align="center", before=2, after=0)
    agenda_title = add_line(doc, "银发科技与具身智能产业创新峰会　活动议程", 16, True, PINE_RGB, "center", before=2, after=2)
    gold_rule(agenda_title)
    add_line(doc, MARKER, 11, True, INK_RGB, align="center", before=6, after=1)
    add_line(
        doc,
        f"本页只列同期出席的其他嘉宾，{omit}。",
        9, color=(0x5E, 0x6A, 0x66), align="center", before=0, after=6,
    )
    rows = agenda_rows(invitee)
    table = doc.add_table(rows=1 + len(rows), cols=3)
    set_borders(table)
    set_table_fixed(table, [3.2, 4.6, 9.8])
    for index, header in enumerate(("时间", "环节", "同期出席嘉宾")):
        write_cell(table.rows[0].cells[index], header, size=9, bold=True, color=WHITE, fill=PINE, align="center")
    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    mark = OxmlElement("w:tblHeader")
    mark.set(qn("w:val"), "true")
    tr_pr.append(mark)
    for row_index, row in enumerate(rows):
        fill = ALT if row_index % 2 == 0 else "FFFFFF"
        for col_index, value in enumerate(row):
            write_cell(
                table.rows[row_index + 1].cells[col_index],
                value,
                size=8.5,
                bold=col_index == 1,
                color=PINE_RGB if col_index == 0 else INK_RGB,
                fill=fill,
                align="center" if col_index == 0 else "left",
                keep_names=col_index == 2,
            )
    add_line(doc, "大会最终议程以现场实际为准。", 9, color=(0x5E, 0x6A, 0x66), align="center", before=8, after=0)


def render_cover(doc: Document, invitees: list[Invitee]) -> None:
    add_banner(doc)
    add_line(doc, "嘉宾邀请函", 26, True, PINE_RGB, "center", before=14, after=2, spacing=160)
    add_line(doc, "组委会发函底稿", 14, True, GOLD_RGB, "center", before=2, after=8)
    add_line(
        doc,
        f"共 {len(invitees)} 封。每封独立成页，所附议程已隐去收件人本人，只保留其他嘉宾。对外发送请使用“邀请函”文件夹中的单封文件，不要把本册整本转给任何一位嘉宾。",
        11, color=INK_RGB, indent=True, after=4,
    )
    add_line(
        doc,
        f"{DATE_LINE}。{PLACE}。",
        10.5, color=INK_RGB, indent=True, after=8,
    )
    table = doc.add_table(rows=1 + len(invitees), cols=4)
    set_borders(table)
    set_table_fixed(table, [1.3, 5.5, 4.6, 6.2])
    for index, header in enumerate(("序号", "收件人", "拟请担任", "单封文件")):
        write_cell(table.rows[0].cells[index], header, size=9, bold=True, color=WHITE, fill=PINE, align="center")
    head_pr = table.rows[0]._tr.get_or_add_trPr()
    head_mark = OxmlElement("w:tblHeader")
    head_mark.set(qn("w:val"), "true")
    head_pr.append(head_mark)
    for index, item in enumerate(invitees, start=1):
        fill = ALT if index % 2 == 0 else "FFFFFF"
        values = (str(index), item.address, item.duty, item.filename)
        for col_index, value in enumerate(values):
            write_cell(table.rows[index].cells[col_index], value, size=8, fill=fill,
                       align="center" if col_index == 0 else "left")
    doc.add_page_break()


def new_document(footer: str) -> Document:
    doc = Document()
    prepare(doc, footer)
    return doc


def paragraph_text(child) -> str:
    return "".join(node.text or "" for node in child.iter(qn("w:t"))).replace("\u2060", "")


def has_page_break(child) -> bool:
    return any(br.get(qn("w:type")) == "page" for br in child.iter(qn("w:br")))


def page_chunks(doc: Document) -> list[str]:
    chunks = []
    current: list[str] = []
    for child in doc.element.body:
        if child.tag == qn("w:p") and has_page_break(child):
            chunks.append("\n".join(current))
            current = []
            text = paragraph_text(child)
            if text.strip():
                current.append(text)
            continue
        if child.tag == qn("w:p"):
            current.append(paragraph_text(child))
        elif child.tag == qn("w:tbl"):
            current.append(paragraph_text(child))
    if current:
        chunks.append("\n".join(current))
    return chunks


def block_text(doc: Document) -> list[tuple[str, str]]:
    blocks = []
    for child in doc.element.body:
        if child.tag == qn("w:p"):
            blocks.append(("p", paragraph_text(child)))
        elif child.tag == qn("w:tbl"):
            blocks.append(("t", paragraph_text(child)))
    return blocks


def agenda_slices(doc: Document) -> list[str]:
    blocks = block_text(doc)
    slices = []
    current = []
    capturing = False
    for _kind, text in blocks:
        if text == MARKER:
            capturing = True
            current = []
            continue
        if capturing and text.startswith("尊敬的"):
            slices.append("\n".join(current))
            capturing = False
            current = []
            continue
        if capturing:
            current.append(text)
    if capturing:
        slices.append("\n".join(current))
    return slices


def verify_documents(invitees: list[Invitee], combined: Path) -> None:
    combined_doc = Document(combined)
    slices = agenda_slices(combined_doc)
    assert len(slices) == len(invitees), (len(slices), len(invitees))
    pages = page_chunks(combined_doc)
    for page in pages:
        if "大会最终议程以现场实际为准" in page:
            assert "邀　请　函" not in page
            assert "尊敬的" not in page
    for item, agenda in zip(invitees, slices):
        for mark in item.marks:
            assert mark not in agenda, (item.address, mark)
        single = Document(LETTER_DIR / item.filename)
        single_pages = page_chunks(single)
        assert len(single_pages) == 2, (item.filename, len(single_pages))
        assert "尊敬的" + item.address in single_pages[0]
        assert MARKER in single_pages[1]
        assert "邀　请　函" not in single_pages[1]
        single_slices = agenda_slices(single)
        assert len(single_slices) == 1
        for mark in item.marks:
            assert mark not in single_slices[0], (item.filename, mark)
        head = []
        seen = False
        for _kind, text in block_text(single):
            if text == MARKER:
                seen = True
                break
            head.append(text)
        assert seen
        head_text = "\n".join(head)
        assert item.address in head_text
        assert "12:30—17:30" in head_text
        assert "8:30" not in head_text
        assert "复旦大学老龄研究院" in head_text
        for _label, value in item.card:
            assert value in head_text, (item.address, value)


def export_pdf(docx_path: Path, outdir: Path) -> Path | None:
    soffice = shutil.which("soffice")
    if not soffice:
        print("未找到 LibreOffice，已跳过 PDF 导出。")
        return None
    outdir.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(docx_path)],
        check=True,
    )
    return outdir / (docx_path.stem + ".pdf")


def build() -> list[Invitee]:
    invitees = collect()
    check_roster(invitees)
    if LETTER_DIR.exists():
        for old in LETTER_DIR.glob("*.docx"):
            old.unlink()
    LETTER_DIR.mkdir(parents=True, exist_ok=True)
    single_footer = "专函　·　所附议程不含收件人本人　·　最终议程以现场为准"
    for item in invitees:
        doc = new_document(single_footer)
        render_letter(doc, item)
        doc.core_properties.title = f"邀请函　{item.address}"
        doc.core_properties.subject = f"{item.duty}　所附议程不含收件人本人"
        doc.save(LETTER_DIR / item.filename)

    combined = new_document("组委会底稿　·　勿整册外发")
    render_cover(combined, invitees)
    for index, item in enumerate(invitees):
        if index:
            combined.add_page_break()
        render_letter(combined, item)
    combined.core_properties.title = "2026银发科技与具身智能产业创新峰会嘉宾邀请函"
    combined.core_properties.subject = f"{len(invitees)}封，所附议程均不含收件人本人"
    combined_path = OUT / "2026银发科技与具身智能产业创新峰会_嘉宾邀请函.docx"
    combined.save(combined_path)
    verify_documents(invitees, combined_path)
    export_pdf(combined_path, OUT)
    print(f"{len(invitees)} 封")
    print(combined_path)
    return invitees


if __name__ == "__main__":
    build()
