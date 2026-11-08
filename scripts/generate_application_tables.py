#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按原申报表的表格版式，从上到下排成一份表格。

参加工作年月改为 2011年06月27日。
工作经历和总结中的职业履历去掉上海贝迪创建科技有限公司、上海宝龙商办企业管理有限公司。
其余栏目按原表抄录。水印不录入。
"""

from __future__ import annotations

import html
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.worksheet.page import PageMargins

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "output"
XLSX_NAME = "专业技术资格评审申报表.xlsx"
HTML_NAME = "专业技术资格评审申报表.html"

THIN = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)
FONT = Font(name="宋体", size=11, color="000000")
LABEL = Font(name="宋体", size=11, bold=True, color="000000")
TITLE = Font(name="宋体", size=18, bold=True, color="000000")
SMALL = Font(name="宋体", size=10, color="000000")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

# 工作经历按原表由近及远。已去掉贝迪、宝龙两段。
JOBS: list[tuple[str, str, str, str, str, str, str]] = [
    ("2025年03月至\n2025年12月", "江苏省建筑装饰设计研究院有限公司", "工程部", "管理岗", "工程技术及项目管理", "副总经理", "2025.03-2025.12"),
    ("2024年08月至\n2025年02月", "上海惟嘉建筑工程有限公司", "工程部", "管理岗", "工程施工技术管理", "副总经理", "2024.08-2025.02"),
    ("2023年10月至\n2024年07月", "上海徐汇绿地商业管理有限公司", "工程部", "管理岗", "工程施工技术管理", "副总经理", "2023.10-2024.07"),
    ("2023年05月至\n2023年08月", "上海容钲建筑科技有限公司", "工程部", "管理岗", "工程施工技术管理", "副总经理", "2023.05-2023.08"),
    ("2022年08月至\n2023年04月", "上海苏科企业管理有限公司", "工程部", "管理岗", "钢结构施工技术管理", "副总经理", "2022.08-2023.04"),
    ("2021年09月至\n2022年07月", "上海社通实业有限公司", "工程部", "管理岗", "工程施工技术管理", "总监", "2021.09-2022.07"),
    ("2021年03月至\n2021年08月", "上海融之旅文化旅游发展有限公司", "工程部", "管理岗", "工程施工技术管理", "总监", "2021.03-2021.08"),
    ("2021年01月至\n2021年02月", "上海弘阳汇商业管理有限公司", "工程部", "管理岗", "工程施工技术管理", "总监", "2021.01-2021.02"),
    ("2020年11月至\n2020年12月", "上海欣驭文化传播有限公司", "工程部", "管理岗", "工程施工技术管理", "总监", "2020.11-2020.12"),
    ("2020年06月至\n2020年10月", "上海金科文化旅游发展集团有限公司", "工程部", "管理岗", "工程施工技术管理", "总监", "2020.06-2020.10"),
    ("2020年03月至\n2020年05月", "金科（上海）建筑设计有限公司", "工程部", "管理岗", "工程施工技术管理", "总监", "2020.03-2020.05"),
    ("2019年04月至\n2020年02月", "上海万居企业管理有限公司", "工程部", "管理岗", "工程施工技术管理", "总监", "2019.04-2020.02"),
    ("2018年01月至\n2019年03月", "万科（上海）实业发展有限公司", "工程部", "管理岗", "工程施工技术管理", "总监", "2018.01-2019.03"),
    ("2017年03月至\n2017年12月", "上海万科长宁置业有限公司", "工程部", "管理岗", "工程施工技术管理", "总监", "2017.03-2017.12"),
    ("2015年12月至\n2017年02月", "新城控股集团股份有限公司上海第一分公司", "工程部", "管理岗", "工程管理级施工协调", "高级经理", "2015.12-2017.02"),
    ("2015年09月至\n2015年10月", "上海睿泽股权投资管理有限公司", "工程部", "技术岗", "施工资料及现场管理", "资料员、施工员", "2015.09-2015.10"),
]

TRAINING = [
    ("2025年06月09日至\n2025年06月21日", "职称继续教育专业课（工程系列）", "南京市人力资源和社会保障局专业技术人员管理处", "学时：240；\n学分：240", "无"),
    ("2025年05月05日至\n2025年06月09日", "公需课有效学时累计", "南京市人力资源和社会保障局专业技术人员管理处", "学时：152；\n学分：152", "无"),
    ("2025年05月03日至\n2025年06月18日", "知识更新工程高级研修项目", "南京市人力资源和社会保障局专业技术人员管理处", "学时：17；\n学分：16.5", "无"),
]

AWARDS = [
    ("住房政策研究创新贡献奖", "2024-12-18", "复旦大学住房政策中心专业委员会"),
    ("城市更新卓越实践奖", "2024-04-26", "绿地控股集团技术委员会"),
    ("可持续社区营造标杆项目奖", "2019-02-24", "上海万科产品创新研究院"),
]

EDU = [
    ("复旦大学", "工商管理", "硕士研究生", "", "于2021年01月13日"),
    ("中国海洋大学青岛学院", "土木工程", "大学本科", "", "于2011年06月30日"),
    ("青岛市城阳一中", "理科", "高中", "其他", "于2016年07月04日毕业"),
]

PAPERS = [
    ("2021-07-08", "铝膜板在高层住宅建筑工程中的应用", "中国民营科技促进会，《中国新技术新产品》", "中国民营科技促进会", "合著"),
    ("2021-08-05", "房屋建筑施工中人工智能技术的应用", "深圳住宅与房地产杂志社;中国房地产及住宅研究会，《住宅与房地产》", "深圳住宅与房地产杂志社;中国房地产及住宅研究会", "合著"),
    ("2024-07-10", "基于BIM的装配式钢结构建筑施工新技术研究", "中国建筑金属结构协会，《中国建筑金属结构》", "中国建筑金属结构协会", "独著"),
]

PERFORMANCE = [
    (
        "2018年02月03日至\n2023年05月11日\n待核查",
        "专业技术工作项目：杭州湾融创文旅城一水、陆乐园项目\n课题名称：杭州湾融创文旅城",
        "获奖情况：无\n专利情况：无\n效益情况：无",
        "主持",
    ),
    (
        "2019年11月15日至\n2022年03月25日\n待核查",
        "专业技术工作项目：万科青浦区徐泾镇徐盈路西侧地块（01-06F2-04 地块一、二期）\n课题名称：万科天空之城",
        "获奖情况：荣获上海万科产品创新研究院颁发的“可持续社区营造标杆项目奖”\n专利情况：专利号：ZL 2020 2 0536480.4；专利号：ZL 2020 2 0557884.1\n效益情况：在上海万科天空之城项目中，主导完成以下创新实践：1. 构建社区级“零碳”能源系统，光伏覆盖率达32%；2. 研发模块化雨水回收装置，节水效率提升45%；3. 搭建社区参与式设计平台，客户满意度达92%。",
        "主持",
    ),
    (
        "2020年10月15日至\n2024年01月25日\n待核查",
        "专业技术工作项目：绿地苏州东太湖旅游度假区WJ-J-2020-013 地块商住\n课题名称：绿地苏州东太湖旅游度假区",
        "获奖情况：荣获绿地控股集团技术委员会颁发的“城市更新卓越实践奖”\n专利情况：无\n效益情况：1. 技术体系创新：构建全流程技术标准体系，集成绿色建材与智能建造技术，单位面积改造成本降低18%，获集团技术创新标杆认证。2. 智慧工地研发：开发研发智慧工地管理系统思路，通过物联网与AI算法优化施工调度，降低人力成本15%、材料损耗率23%，并推动BIM技术应用达LOD500标准。3. 经济效益：累计节约项目成本2,800万元，投资回报率提升32%。",
        "主持",
    ),
]


def _style_range(ws, r1: int, c1: int, r2: int, c2: int, align: Alignment | None = None) -> None:
    for row in range(r1, r2 + 1):
        for col in range(c1, c2 + 1):
            cell = ws.cell(row, col)
            cell.border = THIN
            cell.font = FONT
            cell.alignment = align or CENTER


def _put(ws, row: int, col: int, value: str, rows: int = 1, cols: int = 1, font: Font | None = None, align: Alignment | None = None, height: float | None = None) -> None:
    if rows > 1 or cols > 1:
        ws.merge_cells(start_row=row, start_column=col, end_row=row + rows - 1, end_column=col + cols - 1)
    cell = ws.cell(row, col, value)
    cell.font = font or FONT
    cell.alignment = align or CENTER
    _style_range(ws, row, col, row + rows - 1, col + cols - 1, align or CENTER)
    cell.font = font or FONT
    cell.alignment = align or CENTER
    if height:
        ws.row_dimensions[row].height = height


def _title(ws, row: int, text: str) -> None:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    cell = ws.cell(row, 1, text)
    cell.font = TITLE
    cell.alignment = CENTER
    ws.row_dimensions[row].height = 28


def _note(ws, row: int, text: str) -> None:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    cell = ws.cell(row, 1, text)
    cell.font = SMALL
    cell.alignment = LEFT
    ws.row_dimensions[row].height = 32


def write_xlsx(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "申报表"
    for index, width in enumerate([12, 12, 14, 14, 12, 12, 14, 12, 12, 14, 14, 16], start=1):
        ws.column_dimensions[chr(64 + index)].width = width
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.45, right=0.45, top=0.5, bottom=0.45)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.oddFooter.center.text = "第 &P 页"
    ws.print_options.horizontalCentered = True

    row = 1
    _put(ws, row, 1, "地区代码：320199", cols=4, align=LEFT)
    _put(ws, row, 7, "编号：2506170651940505", cols=6, align=LEFT)
    # 封面这行不要网格，改回无边框。
    for col in range(1, 13):
        ws.cell(row, col).border = Border()
    row += 2
    _title(ws, row, "专业技术资格评审申报表")
    row += 2
    cover = [
        ("单    位", "江苏省建筑装饰设计研究院有限公司"),
        ("姓    名", "胡继刚"),
        ("评审系列", "建设工程"),
        ("评审专业", "科技管理·工程管理"),
        ("技术资格", "高级工程师"),
        ("填表时间", "2025年08月12日"),
    ]
    for label, value in cover:
        _put(ws, row, 3, label, cols=2, font=LABEL, align=CENTER)
        _put(ws, row, 5, value, cols=5)
        for col in range(1, 13):
            ws.cell(row, col).border = Border()
        for col in range(5, 10):
            ws.cell(row, col).border = Border(bottom=Side(style="thin", color="000000"))
        ws.row_dimensions[row].height = 22
        row += 1
    row += 1
    _note(ws, row, "江苏省人力资源和社会保障厅制")
    ws.cell(row, 1).alignment = CENTER
    row += 2

    _title(ws, row, "基  本  情  况")
    row += 1
    basic_start = row
    _put(ws, row, 1, "姓  名", cols=2, font=LABEL)
    _put(ws, row, 3, "胡继刚", cols=2)
    _put(ws, row, 5, "性别", font=LABEL)
    _put(ws, row, 6, "男")
    _put(ws, row, 7, "民族", font=LABEL)
    _put(ws, row, 8, "汉族", cols=2)
    _put(ws, row, 10, "照片", rows=3, cols=3, font=LABEL)
    row += 1
    _put(ws, row, 1, "公民身份号码\n（社会保障号）", cols=2, font=LABEL, height=32)
    _put(ws, row, 3, "2 2 0 2 8 3 1 9 8 7 0 4 2 5 0 6 1 4", cols=7)
    row += 1
    _put(ws, row, 1, "出生年月", cols=2, font=LABEL, height=24)
    _put(ws, row, 3, "1987年04月25日", cols=2)
    _put(ws, row, 5, "参加工作年月", cols=2, font=LABEL)
    _put(ws, row, 7, "2011年06月27日", cols=3)
    row += 1
    _put(ws, row, 1, "取得本行业(专业)准入\n资格、时间及证书号码", cols=3, font=LABEL, height=36)
    _put(ws, row, 4, "", cols=5)
    _put(ws, row, 9, "执业\n类别", font=LABEL)
    _put(ws, row, 10, "工程管理", cols=3)
    row += 1
    _put(ws, row, 1, "现专业技术资格及取得\n时间、批准单位", cols=3, font=LABEL, height=36)
    _put(ws, row, 4, "上海市住房和城乡建设管理委员会批准取得工程师(20181225)", cols=9, align=LEFT)
    row += 1
    _put(ws, row, 1, "现任专业技术职务\n及聘任时间", cols=3, font=LABEL, height=36)
    _put(ws, row, 4, "职务：工程-建设工程-工程师\n聘用时间：2018-12-25", cols=5, align=LEFT)
    _put(ws, row, 9, "申报\n类型", font=LABEL)
    _put(ws, row, 10, "正常申报", cols=3)
    row += 1
    _put(ws, row, 1, "现从事专业技术工\n作及年限", cols=3, font=LABEL, height=36)
    _put(ws, row, 4, "现工作：\n专业技术工作年限：9年", cols=5, align=LEFT)
    _put(ws, row, 9, "拟评审专业技术\n资    格", font=LABEL)
    _put(ws, row, 10, "高级工程师", cols=3)
    row += 1
    _put(ws, row, 1, "学\n历\n情\n况", rows=4, font=LABEL)
    _put(ws, row, 2, "学    校", cols=3, font=LABEL)
    _put(ws, row, 5, "学习专业", cols=2, font=LABEL)
    _put(ws, row, 7, "学历", cols=2, font=LABEL)
    _put(ws, row, 9, "学位", font=LABEL)
    _put(ws, row, 10, "毕(肄、结)业年月", cols=3, font=LABEL)
    for school, major, level, degree, when in EDU:
        row += 1
        _put(ws, row, 2, school, cols=3, height=22)
        _put(ws, row, 5, major, cols=2)
        _put(ws, row, 7, level, cols=2)
        _put(ws, row, 9, degree)
        _put(ws, row, 10, when, cols=3)
    row += 1
    _put(ws, row, 1, "奖\n惩\n情\n况", rows=5, font=LABEL)
    _put(ws, row, 2, "荣誉称号、表彰奖励名称", cols=5, font=LABEL)
    _put(ws, row, 7, "获奖时间", cols=2, font=LABEL)
    _put(ws, row, 9, "授奖部门", cols=4, font=LABEL)
    for name, when, org in AWARDS:
        row += 1
        _put(ws, row, 2, name, cols=5, height=24)
        _put(ws, row, 7, when, cols=2)
        _put(ws, row, 9, org, cols=4)
    row += 1
    _put(ws, row, 2, "处分", font=LABEL, height=22)
    _put(ws, row, 3, "无", cols=10)
    row += 1
    _note(ws, row, "说明：毕(肄、结)业情况应在栏目中注明。")
    row += 2

    _title(ws, row, "继 续 教 育 情 况")
    row += 1
    _note(ws, row, "（国内外培训、进修或考察情况）")
    ws.cell(row, 1).alignment = CENTER
    row += 1
    headers = [("起止时间", 2), ("专业或主要内容", 3), ("组织单位", 3), ("学习成绩", 2), ("备注", 2)]
    col = 1
    for text, span in headers:
        _put(ws, row, col, text, cols=span, font=LABEL, height=22)
        col += span
    for item in TRAINING:
        row += 1
        col = 1
        for value, span in zip(item, [2, 3, 3, 2, 2]):
            _put(ws, row, col, value, cols=span, height=36)
            col += span
    row += 2
    _title(ws, row, "专业考试成绩情况")
    row += 1
    exam_headers = [("时  间", 2), ("组 织 单 位", 3), ("考试名称", 3), ("成  绩", 2), ("合格标志", 2)]
    col = 1
    for text, span in exam_headers:
        _put(ws, row, col, text, cols=span, font=LABEL)
        col += span
    row += 1
    _put(ws, row, 1, "无", cols=12, height=28)
    row += 2

    _title(ws, row, "工  作  经  历")
    row += 1
    _note(ws, row, "（专业技术工作经历）")
    ws.cell(row, 1).alignment = CENTER
    row += 1
    job_headers = [("起止时间", 2), ("工作单位", 3), ("部门", 1), ("岗位", 1), ("从事的主要\n专业技术工作", 3), ("任何职", 2)]
    col = 1
    for text, span in job_headers:
        _put(ws, row, col, text, cols=span, font=LABEL, height=32)
        col += span
    for period, org, dept, post, work, title, _short in JOBS:
        row += 1
        values = [period, org, dept, post, work, title]
        spans = [2, 3, 1, 1, 3, 2]
        col = 1
        for value, span in zip(values, spans):
            _put(ws, row, col, value, cols=span, height=36)
            col += span
    row += 1
    _note(ws, row, "任职年月按工作经历一览表填写，不补具体日。空档2015-11与2023-09未补。")
    row += 2

    _title(ws, row, "任现职后主要专业技术工作业绩登记")
    row += 1
    perf_headers = [("起止时间", 2), ("专业技术工作\n项目、课题名称", 4), ("获奖励、专利\n及效益情况", 4), ("本人起\n何作用", 2)]
    col = 1
    for text, span in perf_headers:
        _put(ws, row, col, text, cols=span, font=LABEL, height=32)
        col += span
    for period, project, benefit, role in PERFORMANCE:
        row += 1
        _put(ws, row, 1, period, cols=2, height=78)
        _put(ws, row, 3, project, cols=4, align=LEFT)
        _put(ws, row, 7, benefit, cols=4, align=LEFT)
        _put(ws, row, 11, role, cols=2)
    row += 1
    _note(ws, row, "说明：“本人起何作用”分为主持、参加、独立承担。本页三项起止与任职时段或奖项日期不一致，只标待核查，不改项目归属。")
    row += 2

    _title(ws, row, "任现职后著作、论文及重要技术报告登记")
    row += 1
    paper_headers = [("日  期", 2), ("论著(报告)标题", 3), ("刊物、出版单位", 3), ("主办单位\n（学术会议）名称", 2), ("本人承担情况", 2)]
    col = 1
    for text, span in paper_headers:
        _put(ws, row, col, text, cols=span, font=LABEL, height=32)
        col += span
    for item in PAPERS:
        row += 1
        col = 1
        for value, span in zip(item, [2, 3, 3, 2, 2]):
            _put(ws, row, col, value, cols=span, height=36, align=CENTER if span <= 2 else LEFT)
            col += span
    row += 1
    _note(ws, row, "说明：本人承担情况按独著(译)、合著(译)填写。合著(译)应注明本人实际承担的部分。")
    row += 2

    _title(ws, row, "本人任现职以来工作总结")
    row += 1
    _note(ws, row, "职业履历。时间、公司名称、职务三列与原表一致。")
    row += 1
    for text, span in [("时间", 3), ("公司名称", 6), ("职务", 3)]:
        _put(ws, row, 1 if text == "时间" else 4 if text == "公司名称" else 10, text, cols=span, font=LABEL)
    for _period, org, _dept, _post, _work, title, short in JOBS:
        row += 1
        _put(ws, row, 1, short, cols=3, height=22)
        _put(ws, row, 4, org, cols=6)
        _put(ws, row, 10, title, cols=3)
    row += 1
    _note(ws, row, "空档2015-11、2023-09未补。一览表没有的单位已从本栏删除。")
    row += 2

    _title(ws, row, "任现职以来考核情况")
    row += 1
    _put(ws, row, 1, "年度\n（任期）\n考核\n情况", rows=4, font=LABEL)
    for text, col, span in [("年  份", 2, 2), ("受聘专业技术职务(岗位)", 4, 3), ("考核等次", 7, 2), ("考核单位", 9, 2), ("备注", 11, 2)]:
        _put(ws, row, col, text, cols=span, font=LABEL, height=22)
    _put(ws, row + 1, 2, "", cols=2, height=22)
    _put(ws, row + 1, 4, "无", cols=3)
    _put(ws, row + 1, 7, "", cols=2)
    _put(ws, row + 1, 9, "", cols=2)
    _put(ws, row + 1, 11, "", cols=2)
    for offset in (2, 3):
        _put(ws, row + offset, 2, "", cols=2, height=22)
        _put(ws, row + offset, 4, "", cols=3)
        _put(ws, row + offset, 7, "", cols=2)
        _put(ws, row + offset, 9, "", cols=2)
        _put(ws, row + offset, 11, "", cols=2)
    row += 5

    _title(ws, row, "评审、登记备案情况")
    row += 1
    _put(ws, row, 1, "专业\n（学科）\n评议组\n意见", rows=2, font=LABEL)
    _put(ws, row, 2, "应到人数", font=LABEL)
    _put(ws, row, 3, "3")
    _put(ws, row, 4, "实到人数", font=LABEL)
    _put(ws, row, 5, "3")
    _put(ws, row, 6, "赞成票", font=LABEL)
    _put(ws, row, 7, "3")
    _put(ws, row, 8, "反对票", font=LABEL)
    _put(ws, row, 9, "0")
    _put(ws, row, 10, "弃权票", font=LABEL)
    _put(ws, row, 11, "0", cols=2)
    row += 1
    _put(ws, row, 2, "评议组组长(签字)                    2025 年 09 月 21 日", cols=11, align=LEFT, height=28)
    row += 1
    _put(ws, row, 1, "评\n审\n委\n员\n会\n意\n见", rows=2, font=LABEL)
    _put(ws, row, 2, "应到人数", font=LABEL)
    _put(ws, row, 3, "27")
    _put(ws, row, 4, "实到人数", font=LABEL)
    _put(ws, row, 5, "27")
    _put(ws, row, 6, "赞成票", font=LABEL)
    _put(ws, row, 7, "27")
    _put(ws, row, 8, "反对票", font=LABEL)
    _put(ws, row, 9, "0")
    _put(ws, row, 10, "弃权票", font=LABEL)
    _put(ws, row, 11, "0", cols=2)
    row += 1
    _put(
        ws,
        row,
        2,
        "经评审，胡继刚（220283198704250614）已具备建设工程·科技管理·工程管理专业（学科）高级工程师资格。\n主任委员(签字)          公 章          2025 年 10 月 19 日",
        cols=11,
        align=LEFT,
        height=48,
    )
    row += 1
    _put(ws, row, 1, "省、市、县或相应职称主管部门登记备案情况", font=LABEL, height=48)
    _put(ws, row, 2, "公 章\n2025 年 12 月 09 日", cols=11)
    row += 1
    _put(ws, row, 1, "备\n注", font=LABEL, height=48)
    _put(
        ws,
        row,
        2,
        "人员姓名：胡继刚          证件号码：220283198704250614\n联系电话：18678408669      流程编号：2506170651940505\n工作单位：江苏省建筑装饰设计研究院有限公司",
        cols=11,
        align=LEFT,
    )
    row += 1
    _note(ws, row, "说明：“评审委员会评审意见”栏，应在“具备”前写明“已”或“不”。")

    ws.row_dimensions[basic_start].height = 24
    ws.page_setup.scale = 70
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _td(text: str, **attrs: str) -> str:
    extra = "".join(f' {key}="{value}"' for key, value in attrs.items())
    return f"<td{extra}>{html.escape(text)}</td>"


def write_html(path: Path) -> None:
    job_rows = []
    for period, org, dept, post, work, title, _short in JOBS:
        cells = "".join(f"<td>{html.escape(value).replace(chr(10), '<br>')}</td>" for value in (period, org, dept, post, work, title))
        job_rows.append(f"<tr>{cells}</tr>")
    summary_rows = []
    for _period, org, _dept, _post, _work, title, short in JOBS:
        summary_rows.append(
            "<tr>"
            + _td(short)
            + _td(org)
            + _td(title)
            + "</tr>"
        )
    train_rows = []
    for item in TRAINING:
        train_rows.append("<tr>" + "".join(f"<td>{html.escape(value).replace(chr(10), '<br>')}</td>" for value in item) + "</tr>")
    paper_rows = []
    for item in PAPERS:
        paper_rows.append("<tr>" + "".join(_td(value) for value in item) + "</tr>")
    perf_rows = []
    for period, project, benefit, role in PERFORMANCE:
        perf_rows.append(
            "<tr>"
            + f"<td>{html.escape(period).replace(chr(10), '<br>')}</td>"
            + f"<td class=\"left\">{html.escape(project).replace(chr(10), '<br>')}</td>"
            + f"<td class=\"left\">{html.escape(benefit).replace(chr(10), '<br>')}</td>"
            + _td(role)
            + "</tr>"
        )

    page = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>专业技术资格评审申报表</title>
<style>
body {{ margin: 0; background: #f3f3f3; color: #000; font-family: SimSun, "Songti SC", "Noto Serif SC", serif; }}
main {{ max-width: 980px; margin: 0 auto; padding: 28px 16px 72px; }}
section {{ background: #fff; border: 1px solid #111; padding: 22px 18px 26px; margin: 0 0 28px; }}
h1 {{ margin: 28px 0 8px; text-align: center; font-size: 32px; letter-spacing: .35em; font-weight: 700; }}
h2 {{ margin: 0 0 14px; text-align: center; font-size: 24px; letter-spacing: .4em; font-weight: 700; }}
.sub {{ margin: -8px 0 12px; text-align: center; font-size: 14px; }}
.meta {{ display: flex; justify-content: space-between; font-size: 15px; }}
.cover {{ width: 78%; margin: 18px auto 8px; border-collapse: collapse; }}
.cover td {{ border: none; padding: 7px 8px; font-size: 16px; }}
.cover .label {{ width: 28%; text-align: right; letter-spacing: .2em; }}
.cover .value {{ border-bottom: 1px solid #000; text-align: center; }}
.maker {{ text-align: center; margin: 22px 0 0; font-size: 15px; }}
table.grid {{ width: 100%; border-collapse: collapse; table-layout: fixed; }}
table.grid th, table.grid td {{ border: 1px solid #000; padding: 6px 6px; text-align: center; vertical-align: middle; font-size: 14px; line-height: 1.45; font-weight: 400; }}
table.grid th {{ font-weight: 700; }}
td.left {{ text-align: left; }}
td.label {{ font-weight: 700; }}
td.photo {{ height: 118px; }}
.note {{ margin: 8px 0 0; font-size: 13px; line-height: 1.5; }}
</style>
</head>
<body>
<main>
<section id="cover">
<p class="meta"><span>地区代码：320199</span><span>编号：2506170651940505</span></p>
<h1>专业技术资格评审申报表</h1>
<table class="cover">
<tr><td class="label">单位</td><td class="value">江苏省建筑装饰设计研究院有限公司</td></tr>
<tr><td class="label">姓名</td><td class="value">胡继刚</td></tr>
<tr><td class="label">评审系列</td><td class="value">建设工程</td></tr>
<tr><td class="label">评审专业</td><td class="value">科技管理·工程管理</td></tr>
<tr><td class="label">技术资格</td><td class="value">高级工程师</td></tr>
<tr><td class="label">填表时间</td><td class="value">2025年08月12日</td></tr>
</table>
<p class="maker">江苏省人力资源和社会保障厅制</p>
</section>

<section id="basic">
<h2>基本情况</h2>
<table class="grid">
<tr>
<td class="label" style="width:16%">姓 名</td><td>胡继刚</td>
<td class="label" style="width:10%">性别</td><td style="width:8%">男</td>
<td class="label" style="width:10%">民族</td><td style="width:10%">汉族</td>
<td class="photo label" rowspan="3" style="width:16%">照片</td>
</tr>
<tr>
<td class="label">公民身份号码<br>（社会保障号）</td>
<td colspan="5">2 2 0 2 8 3 1 9 8 7 0 4 2 5 0 6 1 4</td>
</tr>
<tr>
<td class="label">出生年月</td><td>1987年04月25日</td>
<td class="label" colspan="2">参加工作年月</td>
<td colspan="2">2011年06月27日</td>
</tr>
<tr>
<td class="label" colspan="2">取得本行业(专业)准入资格、时间及证书号码</td>
<td colspan="3"></td>
<td class="label">执业类别</td><td>工程管理</td>
</tr>
<tr>
<td class="label" colspan="2">现专业技术资格及取得时间、批准单位</td>
<td class="left" colspan="5">上海市住房和城乡建设管理委员会批准取得工程师(20181225)</td>
</tr>
<tr>
<td class="label" colspan="2">现任专业技术职务及聘任时间</td>
<td class="left" colspan="3">职务：工程-建设工程-工程师<br>聘用时间：2018-12-25</td>
<td class="label">申报类型</td><td>正常申报</td>
</tr>
<tr>
<td class="label" colspan="2">现从事专业技术工作及年限</td>
<td class="left" colspan="3">现工作：<br>专业技术工作年限：9年</td>
<td class="label">拟评审专业技术资格</td><td>高级工程师</td>
</tr>
<tr>
<td class="label" rowspan="4">学<br>历<br>情<br>况</td>
<th>学 校</th><th>学习专业</th><th>学历</th><th>学位</th><th colspan="2">毕(肄、结)业年月</th>
</tr>
<tr><td>复旦大学</td><td>工商管理</td><td>硕士研究生</td><td></td><td colspan="2">于2021年01月13日</td></tr>
<tr><td>中国海洋大学青岛学院</td><td>土木工程</td><td>大学本科</td><td></td><td colspan="2">于2011年06月30日</td></tr>
<tr><td>青岛市城阳一中</td><td>理科</td><td>高中</td><td>其他</td><td colspan="2">于2016年07月04日毕业</td></tr>
<tr>
<td class="label" rowspan="5">奖<br>惩<br>情<br>况</td>
<th colspan="3">荣誉称号、表彰奖励名称</th><th>获奖时间</th><th colspan="2">授奖部门</th>
</tr>
{''.join(f'<tr><td colspan="3">{html.escape(name)}</td><td>{html.escape(when)}</td><td colspan="2">{html.escape(org)}</td></tr>' for name, when, org in AWARDS)}
<tr><td class="label">处分</td><td colspan="5">无</td></tr>
</table>
<p class="note">说明：毕(肄、结)业情况应在栏目中注明。</p>
</section>

<section id="train">
<h2>继续教育情况</h2>
<p class="sub">（国内外培训、进修或考察情况）</p>
<table class="grid">
<tr><th>起止时间</th><th>专业或主要内容</th><th>组织单位</th><th>学习成绩</th><th>备注</th></tr>
{''.join(train_rows)}
</table>
<h2 style="margin-top:28px">专业考试成绩情况</h2>
<table class="grid">
<tr><th>时 间</th><th>组 织 单 位</th><th>考试名称</th><th>成 绩</th><th>合格标志</th></tr>
<tr><td colspan="5">无</td></tr>
</table>
</section>

<section id="jobs">
<h2>工作经历</h2>
<p class="sub">（专业技术工作经历）</p>
<table class="grid">
<tr><th>起止时间</th><th>工作单位</th><th>部门</th><th>岗位</th><th>从事的主要专业技术工作</th><th>任何职</th></tr>
{''.join(job_rows)}
</table>
<p class="note">任职年月按工作经历一览表填写，不补具体日。空档2015-11与2023-09未补。</p>
</section>

<section id="performance">
<h2>任现职后主要专业技术工作业绩登记</h2>
<table class="grid">
<tr><th style="width:18%">起止时间</th><th>专业技术工作项目、课题名称</th><th>获奖励、专利及效益情况</th><th style="width:10%">本人起何作用</th></tr>
{''.join(perf_rows)}
</table>
<p class="note">说明：“本人起何作用”分为主持、参加、独立承担。本页三项起止与任职时段或奖项日期不一致，只标待核查，不改项目归属。</p>
</section>

<section id="papers">
<h2>任现职后著作、论文及重要技术报告登记</h2>
<table class="grid">
<tr><th>日 期</th><th>论著(报告)标题</th><th>刊物、出版单位</th><th>主办单位（学术会议）名称</th><th>本人承担情况</th></tr>
{''.join(paper_rows)}
</table>
<p class="note">说明：本人承担情况按独著(译)、合著(译)填写。合著(译)应注明本人实际承担的部分。</p>
</section>

<section id="summary">
<h2>本人任现职以来工作总结</h2>
<p class="sub">职业履历</p>
<table class="grid">
<tr><th style="width:24%">时间</th><th>公司名称</th><th style="width:22%">职务</th></tr>
{''.join(summary_rows)}
</table>
<p class="note">空档2015-11、2023-09未补。一览表没有的单位已从本栏删除。</p>
</section>

<section id="review">
<h2>任现职以来考核情况</h2>
<table class="grid">
<tr>
<td class="label" rowspan="4" style="width:12%">年度<br>（任期）<br>考核<br>情况</td>
<th>年 份</th><th>受聘专业技术职务(岗位)</th><th>考核等次</th><th>考核单位</th><th>备注</th>
</tr>
<tr><td></td><td>无</td><td></td><td></td><td></td></tr>
<tr><td></td><td></td><td></td><td></td><td></td></tr>
<tr><td></td><td></td><td></td><td></td><td></td></tr>
</table>
<h2 style="margin-top:28px">评审、登记备案情况</h2>
<table class="grid">
<tr>
<td class="label" rowspan="2" style="width:12%">专业（学科）评议组意见</td>
<td class="label">应到人数</td><td>3</td><td class="label">实到人数</td><td>3</td>
<td class="label">赞成票</td><td>3</td><td class="label">反对票</td><td>0</td><td class="label">弃权票</td><td>0</td>
</tr>
<tr><td class="left" colspan="10">评议组组长(签字)　　2025 年 09 月 21 日</td></tr>
<tr>
<td class="label" rowspan="2">评审委员会意见</td>
<td class="label">应到人数</td><td>27</td><td class="label">实到人数</td><td>27</td>
<td class="label">赞成票</td><td>27</td><td class="label">反对票</td><td>0</td><td class="label">弃权票</td><td>0</td>
</tr>
<tr><td class="left" colspan="10">经评审，胡继刚（220283198704250614）已具备建设工程·科技管理·工程管理专业（学科）高级工程师资格。<br>主任委员(签字)　　公 章　　2025 年 10 月 19 日</td></tr>
<tr>
<td class="label">省、市、县或相应职称主管部门登记备案情况</td>
<td colspan="10">公 章<br>2025 年 12 月 09 日</td>
</tr>
<tr>
<td class="label">备注</td>
<td class="left" colspan="10">人员姓名：胡继刚　　证件号码：220283198704250614<br>联系电话：18678408669　　流程编号：2506170651940505<br>工作单位：江苏省建筑装饰设计研究院有限公司</td>
</tr>
</table>
<p class="note">说明：“评审委员会评审意见”栏，应在“具备”前写明“已”或“不”。</p>
</section>
</main>
</body>
</html>
"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(page, encoding="utf-8")


def main() -> None:
    write_xlsx(OUT_DIR / XLSX_NAME)
    write_html(OUT_DIR / HTML_NAME)
    print(OUT_DIR / XLSX_NAME)
    print(OUT_DIR / HTML_NAME)


if __name__ == "__main__":
    main()
