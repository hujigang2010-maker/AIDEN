#!/usr/bin/env python3
"""嘉宾均可由复旦大学请到时，生成 8:45—12:30 的主旨与四场圆桌议程。"""

import shutil
import subprocess
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "deliverables"
FONT = "微软雅黑"
PINE = "0F5C58"
PINE_RGB = (0x0F, 0x5C, 0x58)
INK_RGB = (0x1B, 0x3A, 0x4B)
GOLD_RGB = (0x8C, 0x6A, 0x32)
ALT = "F4F8F7"
LINE = "D5DDD8"

ORGANIZERS = [
    ("指导单位", "上海现代服务业联合会、上海市养老服务和老龄产业协会、国家技术转移东部中心"),
    ("主办", "上海星智耀科技有限公司"),
    ("承办", "上海市杨浦区科技企业联合会、复旦大学住房政策研究中心"),
    ("协办", "长三角一体化示范区新发展建设有限公司"),
    ("支持单位", "上海辰山国家植物园、璞跃中国、上海视觉艺术学院"),
    ("学术支持", "复旦大学老龄研究院"),
]

OVERVIEW = [
    ("8:45—9:00", "签到入场", "嘉宾与观众入场"),
    ("9:00—9:06", "开幕", "徐洁平（璞跃中国）"),
    ("9:06—9:14", "致辞", "指导单位代表"),
    ("9:14—10:14", "主旨分享", "10人，每人6分钟"),
    ("10:14—10:20", "茶歇分组交流", "未上台的拟邀单位同一时段交流"),
    ("10:20—10:50", "圆桌一", "具身智能进院与进家；赵博誉开题，朱翊元收束"),
    ("10:50—11:20", "圆桌二", "数智康养、医疗AI与智慧物联；胡继刚开题，吴玉韶收束"),
    ("11:20—11:50", "圆桌三", "机构采购、试点与运营；邱建伟开题，赵博誉收束"),
    ("11:50—12:20", "圆桌四", "产业资本、银行与支付；胡继刚开题，封进收束"),
    ("12:20—12:30", "总结合影", "主办、承办方代表"),
]

KEYNOTES = [
    ("9:14—9:20", "彭希哲", "复旦大学老龄研究院院长、文科资深教授", "《“十五五”银发经济与长寿社会发展的新机遇》"),
    ("9:20—9:26", "上海辰山国家植物园代表", "上海辰山国家植物园", "园艺康养与银发友好的绿色公共空间"),
    ("9:26—9:32", "朱震", "复旦大学国家科技园", "科技成果转化与康养科技孵化"),
    ("9:32—9:38", "上海视觉艺术学院代表", "上海视觉艺术学院", "康养空间与适老设计"),
    ("9:38—9:44", "顾捷", "傅利叶智能创始人、CEO", "《从康复机器人到具身智能：AI如何改变未来养老》"),
    ("9:44—9:50", "王兴兴", "宇树科技创始人、董事长", "《机器人走入千家万户：具身智能的未来十年》"),
    ("9:50—9:56", "邓泰华", "智元机器人创始人、董事长兼CEO", "《具身智能产业化：人形机器人走向现实场景》"),
    ("9:56—10:02", "邱锡鹏", "复旦大学教授、上海创智学院全时导师", "《从大模型到智能体：AI康养的下一代交互技术》"),
    ("10:02—10:08", "徐立", "商汤科技董事长兼CEO", "《大模型与具身智能：AI赋能银发科技新场景》"),
    ("10:08—10:14", "韩璧丞", "强脑科技创始人、CEO", "《脑机接口与智能康复：AI赋能主动健康》"),
]

PANELS = [
    {
        "title": "圆桌一　具身智能进院与进家",
        "time": "10:20—10:50",
        "question": "机器人进入养老院和家庭，哪一段已经能交付？主持：赵博誉、朱翊元",
        "rows": [
            ("10:20—10:24", "赵博誉", "上海星智耀科技有限公司，圆桌主持", "开题：先回答能交付的任务，再回答谁付费"),
            ("10:24—10:27", "彭志辉", "智元机器人联合创始人、总裁、CTO", "人形机器人进入家庭：技术突破之后，哪一步才能规模化"),
            ("10:27—10:30", "朱秋国", "云深处科技创始人", "智能机器人产业化与可进入的场景"),
            ("10:30—10:33", "易港", "钛虎机器人创始人、董事长", "核心硬件、灵巧操作和产业链是否撑得住养老场景"),
            ("10:33—10:36", "钱东奇", "科沃斯集团创始人、董事长", "服务机器人进入家庭，与长寿经济怎样衔接"),
            ("10:36—10:39", "钱程", "科沃斯集团副董事长、科沃斯品牌CEO", "家庭服务机器人与智能居家的现有用户"),
            ("10:39—10:42", "傅利叶康复业务拟邀负责人", "傅利叶智能康复业务团队", "康复设备进机构后的人员、流程和人机交互"),
            ("10:42—10:45", "擎朗智能拟邀业务负责人", "擎朗智能", "服务机器人在养老机构里的日常运营"),
            ("10:45—10:50", "朱翊元", "长三角一体化示范区新发展建设有限公司，圆桌主持", "收束：示范区场景里，今天能进场的任务和还不能承诺的任务"),
        ],
    },
    {
        "title": "圆桌二　数智康养、医疗AI与智慧物联",
        "time": "10:50—11:20",
        "question": "智能体、医疗AI和家庭感知，分别交到谁手里？主持：胡继刚、吴玉韶",
        "rows": [
            ("10:50—10:54", "胡继刚", "复旦大学住房政策研究中心秘书长、副教授级高级工程师、上海市杨浦区科技企业联合会执行会长，圆桌主持", "开题：不谈参数，只谈服务交到哪一级"),
            ("10:54—10:57", "胡湛", "复旦大学老龄研究院副院长、银龄未来实验室主任", "数智养老与AI向善，哪些环节必须留下人"),
            ("10:57—11:00", "肖仰华", "复旦大学教授、上海市数据科学重点实验室主任", "AI智能体、知识图谱如何支撑智慧健康服务"),
            ("11:00—11:03", "王学集", "涂鸦智能创始人兼CEO", "AI Home怎样成为未来智慧养老家庭的系统"),
            ("11:03—11:06", "张强", "联影医疗董事长、联席CEO", "高端医疗设备创新与主动健康"),
            ("11:06—11:09", "沈定刚", "联影智能联合创始人、联席CEO", "医疗AI与智能化健康管理的责任边界"),
            ("11:09—11:12", "钱鹏鹤", "移远通信董事长兼CEO", "AIoT、智能感知与智慧养老物联网络"),
            ("11:12—11:15", "张瑞安", "乐鑫科技创始人、CEO", "AIoT芯片与低成本智能健康设备"),
            ("11:15—11:20", "吴玉韶", "复旦大学老龄研究院副院长、教授，圆桌主持", "收束：机构、社区、居家分别接住哪一类技术"),
        ],
    },
    {
        "title": "圆桌三　机构采购、试点与运营",
        "time": "11:20—11:50",
        "question": "谁来为养老机器人买单，试点要满足什么条件？主持：邱建伟、赵博誉",
        "rows": [
            ("11:20—11:24", "邱建伟", "泰康之家首席执行官，圆桌主持", "开题：买过什么、退过什么、怎样才愿意再试"),
            ("11:24—11:27", "赵一洪", "福寿康联合创始人、副总裁", "AI+居家养老及长期照护，家属是否付费"),
            ("11:27—11:30", "康剑", "福寿康集团产品及供给负责人、上海分公司总经理", "智慧护理技术的采购标准与落地"),
            ("11:30—11:33", "瞿文", "上海福苑康养健康管理有限公司联合创始人、董事长", "智慧健康、康复和养老运营怎样协同"),
            ("11:33—11:36", "陈楠", "博斯腾科技联合创始人、CEO", "AI认知健康与老年认知障碍筛查能否进入机构"),
            ("11:36—11:39", "九如城拟邀负责人", "九如城养老集团", "连锁养老机构的智慧化改造，先改哪一层"),
            ("11:39—11:42", "第一社会福利院拟邀负责人", "上海市第一社会福利院", "公办机构的产品试点与采购需求"),
            ("11:42—11:45", "康复医院拟邀负责人", "上海相关康复医院", "康复机器人的临床应用与验证条件"),
            ("11:45—11:50", "赵博誉", "上海星智耀科技有限公司，圆桌主持", "收束：愿意付费的类别、试点条件、明确不买的类别"),
        ],
    },
    {
        "title": "圆桌四　产业资本、银行与支付",
        "time": "11:50—12:20",
        "question": "设备租赁、采购支付和股权投资，哪一条现在就能做？主持：胡继刚、封进",
        "rows": [
            ("11:50—11:54", "胡继刚", "复旦大学住房政策研究中心秘书长、副教授级高级工程师、上海市杨浦区科技企业联合会执行会长，圆桌主持", "开题：只谈有项目权限的资金，不谈品牌宣传"),
            ("11:54—11:57", "顾建忠", "上海银行董事长", "银发经济及科技金融协同"),
            ("11:57—12:00", "工商银行上海市分行拟邀负责人", "养老金融或科技金融业务", "银发科技企业的融资与金融服务"),
            ("12:00—12:03", "优势资本拟邀投资合伙人", "优势资本", "有项目投资权限的产业判断"),
            ("12:03—12:06", "启明创投拟邀负责人", "医疗科技或硬科技投资团队", "医疗科技、硬科技项目的投资标准"),
            ("12:06—12:09", "上海国投拟邀负责人", "相关产业基金", "市级产业基金如何承接康养科技项目"),
            ("12:09—12:12", "张江科投拟邀负责人", "张江科投", "张江科技产业投资与场景对接"),
            ("12:12—12:15", "专项基金拟邀负责人", "机器人及具身智能专项产业基金", "机器人项目的股权、租赁和回报"),
            ("12:15—12:20", "封进", "复旦大学特聘教授，圆桌主持", "收束：租赁、采购支付和长期护理保险，哪一条现在就能做"),
        ],
    },
]

EXCHANGE = [
    ("尚海龙", "商汤医疗联合创始人兼首席战略官", "医疗大模型与智慧健康应用"),
    ("萤石网络拟邀负责人", "萤石网络", "家庭看护、智能居家与物联网终端"),
    ("海康威视拟邀负责人", "海康威视智慧养老相关业务", "智能视觉、安全监测、居家感知"),
    ("大华股份拟邀负责人", "大华股份智慧物联相关业务", "智慧社区、养老安全及监测设备"),
    ("复星康养拟邀负责人", "复星康养、星堡养老", "高品质养老社区的科技应用"),
    ("申园拟邀负责人", "泰康之家·申园", "上海养老机构的真实场景"),
    ("交通银行拟邀负责人", "交通银行养老金融业务", "养老金金融及银发产业"),
    ("浦发银行拟邀负责人", "浦发银行科技金融业务", "机器人及医疗科技企业融资"),
    ("杭州银行拟邀负责人", "杭州银行上海分行科技金融业务", "长三角科创企业金融支持"),
    ("泰隆银行拟邀负责人", "泰隆银行上海相关分支机构", "科技型中小企业金融服务"),
    ("太平养老拟邀负责人", "太平养老", "长护险、养老服务支付与保险协同"),
    ("平安养老拟邀负责人", "平安养老", "长护险、养老服务支付与保险协同"),
    ("浦东创投拟邀负责人", "浦东创投", "产业投资交流"),
    ("临港科创投拟邀负责人", "临港科创投", "产业投资交流"),
    ("复星投资拟邀负责人", "复星相关投资团队", "产业投资交流"),
    ("长三角基金拟邀负责人", "长三角医疗科技及智能制造产业基金", "产业投资交流"),
]


def validate() -> None:
    names = []
    for row in KEYNOTES:
        names.append(row[1])
    hosts = {"赵博誉", "胡继刚", "朱翊元", "吴玉韶", "邱建伟", "封进"}
    for panel in PANELS:
        for row in panel["rows"]:
            if row[1] not in hosts:
                names.append(row[1])
    assert len(names) == len(set(names)), [n for n in names if names.count(n) > 1]
    assert len(names) == 38, len(names)
    assert "朱震" in names
    for host in ("朱翊元", "吴玉韶", "邱建伟", "封进"):
        assert any(row[1] == host and "主持" in row[2] for panel in PANELS for row in panel["rows"])
    assert "上海辰山国家植物园代表" in names and "上海视觉艺术学院代表" in names
    assert OVERVIEW[0][0].startswith("8:45")
    assert OVERVIEW[-1][0].endswith("12:30")
    assert [len(panel["rows"]) for panel in PANELS] == [9, 9, 9, 9]


def set_run_font(run, size: float, bold: bool = False, color: tuple | None = None) -> None:
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


def shade(cell, color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    node = OxmlElement("w:shd")
    node.set(qn("w:val"), "clear")
    node.set(qn("w:color"), "auto")
    node.set(qn("w:fill"), color)
    tc_pr.append(node)


def set_cell_margins(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = OxmlElement("w:tcMar")
    for edge, value in (("top", "36"), ("left", "60"), ("bottom", "36"), ("right", "60")):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:w"), value)
        node.set(qn("w:type"), "dxa")
        tc_mar.append(node)
    tc_pr.append(tc_mar)
    valign = OxmlElement("w:vAlign")
    valign.set(qn("w:val"), "center")
    tc_pr.append(valign)


def set_table_widths(table, widths: list[float]) -> None:
    table.autofit = False
    table.allow_autofit = False
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(int(sum(widths) * 567)))
    tbl_w.set(qn("w:type"), "dxa")
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


def set_table_borders(table) -> None:
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "4")
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), LINE)
        borders.append(node)
    table._tbl.tblPr.append(borders)


def write_cell(cell, text: str, size: float = 9, bold: bool = False, color: tuple | None = None,
               fill: str | None = None, align: str = "left") -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if align == "center" else WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.05
    set_run_font(paragraph.add_run(text), size, bold, color)
    if fill:
        shade(cell, fill)
    set_cell_margins(cell)


def add_table(doc, headers: list[str], rows: list[list[str]], widths: list[float], size: float = 9) -> None:
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    set_table_borders(table)
    set_table_widths(table, widths)
    for index, header in enumerate(headers):
        write_cell(table.rows[0].cells[index], header, size=9, bold=True, color=(255, 255, 255), fill=PINE, align="center")
    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    mark = OxmlElement("w:tblHeader")
    mark.set(qn("w:val"), "true")
    tr_pr.append(mark)
    for row_index, row in enumerate(rows):
        fill = ALT if row_index % 2 == 0 else "FFFFFF"
        for col_index, value in enumerate(row):
            write_cell(table.rows[row_index + 1].cells[col_index], value, size=size, fill=fill,
                       align="center" if col_index == 0 else "left")


def add_line(doc, text: str, size: float = 12, bold: bool = False, color: tuple | None = None,
             align: str = "left", before: float = 0, after: float = 3, keep: bool = False):
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if align == "center" else WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = 1.08
    paragraph.paragraph_format.keep_with_next = keep
    set_run_font(paragraph.add_run(text), size, bold, color)
    return paragraph


def add_page_field(paragraph) -> None:
    run = paragraph.add_run()
    set_run_font(run, 8, color=(0x66, 0x66, 0x66))
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)


def build(path: Path) -> None:
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(10.5)
    r_pr = normal.element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        r_fonts.set(qn(attr), FONT)

    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.4)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.header_distance = Cm(0.5)
    section.footer_distance = Cm(0.4)
    header = section.header.paragraphs[0]
    set_run_font(header.add_run("银发科技与具身智能产业创新峰会　·　主旨与圆桌议程"), 8, color=PINE_RGB)
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(footer.add_run("2026年11月28日　8:45—12:30　·　第 "), 8, color=(0x66, 0x66, 0x66))
    add_page_field(footer)
    set_run_font(footer.add_run(" 页"), 8, color=(0x66, 0x66, 0x66))

    add_line(doc, "2026上海国际智慧康养产业博览会", 11, color=GOLD_RGB, align="center", after=1)
    add_line(doc, "AI重塑养老未来", 18, True, PINE_RGB, "center", after=0)
    add_line(doc, "2026银发科技与具身智能产业创新峰会", 13, True, INK_RGB, "center", after=1)
    add_line(doc, "2026年11月28日（周六）8:45—12:30　上海展览中心中央大厅主论坛区（一楼）", 10.5, align="center", after=2)
    add_line(doc, "嘉宾由复旦大学出面邀请。", 10.5, align="center", after=6)

    add_line(doc, "主办方", 15, True, PINE_RGB, before=2, after=2)
    for role, names in ORGANIZERS:
        add_line(doc, f"{role}：{names}", 10.5, after=1)

    add_line(doc, "议程总览", 15, True, PINE_RGB, before=8, after=2, keep=True)
    add_table(doc, ["时间", "环节", "安排"], [list(row) for row in OVERVIEW], [3.2, 3.6, 11.2])

    keynote_span = f"{KEYNOTES[0][0].split('—')[0]}—{KEYNOTES[-1][0].split('—')[1]}"
    add_line(doc, f"主旨分享　{keynote_span}", 15, True, PINE_RGB, before=10, after=2, keep=True)
    add_table(
        doc,
        ["时间", "嘉宾", "职务", "题目"],
        [list(row) for row in KEYNOTES],
        [3.0, 2.3, 5.5, 7.2],
        size=8.5,
    )

    for panel in PANELS:
        add_line(doc, f"{panel['title']}　{panel['time']}", 14, True, PINE_RGB, before=10, after=1, keep=True)
        add_line(doc, panel["question"], 11, True, INK_RGB, after=2, keep=True)
        add_table(
            doc,
            ["时间", "嘉宾", "单位及身份", "讨论内容"],
            [list(row) for row in panel["rows"]],
            [2.8, 3.4, 5.6, 6.2],
            size=8.5,
        )

    tea_time = next(row[0] for row in OVERVIEW if "茶歇" in row[1])
    add_line(doc, f"茶歇分组交流　{tea_time}", 14, True, PINE_RGB, before=10, after=1, keep=True)
    add_line(doc, "同一时段分组交流，不逐一上台。", 10.5, after=2, keep=True)
    add_table(
        doc,
        ["嘉宾", "单位及身份", "交流内容"],
        [list(row) for row in EXCHANGE],
        [4.2, 6.4, 7.4],
        size=8.5,
    )

    closing = next(row for row in OVERVIEW if row[1] == "总结合影")
    add_line(doc, f"总结及合影　{closing[0]}　{closing[2]}收束后合影。", 13, True, PINE_RGB, before=8, after=2)

    core = doc.core_properties
    core.title = "2026银发科技与具身智能产业创新峰会主旨与圆桌议程"
    core.subject = "8:45—12:30，主旨10人；圆桌主持为赵博誉、胡继刚、朱翊元、吴玉韶、邱建伟、封进"
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


def export_pdf(docx_path: Path) -> None:
    soffice = shutil.which("soffice")
    if not soffice:
        return
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(docx_path.parent), str(docx_path)],
        check=True,
    )


def main() -> None:
    validate()
    docx_path = OUT / "2026上海国际智慧康养产业博览会_主旨与圆桌最优议程.docx"
    build(docx_path)
    export_pdf(docx_path)
    print(docx_path)


if __name__ == "__main__":
    main()
