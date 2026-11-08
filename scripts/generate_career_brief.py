#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""整理胡继刚职业履历与核心项目业绩。

正文来源：
- 《专业技术资格评审申报表》第5页工作经历、第6页业绩登记、第8–11页个人专业技术工作总结。
- 15份《项目证明》只读到文件名中的单位，未读到.doc正文。
- 申报表和总结都没有单列的单位，起止与职务沿用此前《工作经历一览表》原文，并标明来源。
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "output"
DOCX_NAME = "胡继刚-职业履历与核心项目.docx"
XLSX_NAME = "胡继刚-职业履历与核心项目.xlsx"
HTML_NAME = "胡继刚-职业履历与核心项目.html"

NAVY = "1F4E79"
INK = "1F2933"
MUTED = "334155"
LINE = "1F2933"

FORBIDDEN = (
    "220283198704250614",
    "18678408669",
    "2506170651940505",
    "上海贝迪创建科技有限公司",
    "上海宝龙商办企业管理有限公司",
)

INTRO = (
    "职业履历按两套原文并列，不把项目改挂到另一段年份。"
    "第一套是申报表所附《个人专业技术工作总结》的7段，与后面的项目时间线同一口径。"
    "第二套是申报表第5页工作经历栏的8段，保留到日的起止和工商登记名称。"
    "15份项目证明的.doc正文本次没有读到，单位名称取自文件名。"
    "其中7家单位在申报表和工作总结里都没有单独一行，起止与职务改用此前《工作经历一览表》已经录入的原文。"
)

SUMMARY_HEADERS = ["序号", "起止时间", "公司名称", "职务"]
SUMMARY_ROWS = [
    ["1", "2011.07–2015.05", "江苏中南建设集团股份有限公司", "施工员/资料员"],
    ["2", "2015.06–2017.02", "新城控股集团上海第一分公司", "工程部高级经理"],
    ["3", "2017.03–2017.10", "金科地产集团有限公司华东区域公司", "工程总监"],
    ["4", "2017.11–2022.03", "上海万科企业有限公司", "工程部总监"],
    ["5", "2022.04–2023.05", "上海融之旅文化旅游发展有限公司（融创文旅华东区域）", "工程总监"],
    ["6", "2023.06–2025.02", "上海徐汇绿地商业管理有限公司", "工程部副总经理"],
    ["7", "2025.03起", "江苏省建筑装饰设计研究院有限公司", "工程管理副总经理"],
]
SUMMARY_NOTE = (
    "上表按工作总结原文转录。总结对第7段写作“2025.03-至今”。"
    "申报表第5页同一单位写至2025年07月10日。本表不把“至今”延写到2026年。"
    "申报表基本情况栏的参加工作年月是2011年06月27日，与第1段起始月相邻。"
    "第4段“上海万科企业有限公司”、第3段“金科地产集团有限公司华东区域公司”不是本批项目证明上的单位全称。"
)

FORM_HEADERS = ["序号", "起止时间", "公司名称", "职务"]
FORM_ROWS = [
    ["1", "2015.06.08–2017.02.24", "新城控股集团股份有限公司上海第一分公司", "工程部高级经理"],
    ["2", "2017.03.20–2017.05.29", "金科（上海）建筑设计有限公司", "工程部资深经理"],
    ["3", "2017.06.01–2017.10.26", "上海金科文化旅游发展集团有限公司", "工程部资深经理"],
    ["4", "2017.11.06–2018.11.10", "万科（上海）实业发展有限公司", "工程部总监"],
    ["5", "2018.12.10–2022.03.18", "上海万科长宁置业有限公司", "工程部总监"],
    ["6", "2022.04.11–2023.05.13", "上海融之旅文化旅游发展有限公司", "工程部总监"],
    ["7", "2023.06.07–2025.02.28", "上海徐汇绿地商业管理有限公司", "工程部副总经理"],
    ["8", "2025.03.01–2025.07.10", "江苏省建筑装饰设计研究院有限公司", "工程部副总经理"],
]
FORM_NOTE = (
    "上表按申报表第5页转录，部门“工程部”与“任何职”合成职务。"
    "第4段申报表原文写“上海万限实业有限公司”，括注“万科（上海）实业发展有限公司”；本表用括注中的单位全称，与项目证明文件名一致。"
    "第6段括注“融创文旅华东区域”。"
    "第5页没有江苏中南建设集团股份有限公司。工作总结有2011.07–2015.05、施工员/资料员。"
)

PROOF_HEADERS = ["证明文件", "公司名称", "申报表或工作总结", "一览表对照"]
PROOF_ROWS = [
    [
        "14-5-15",
        "江苏中南建设集团股份有限公司",
        "工作总结：2011.07–2015.05，施工员/资料员。申报表第5页未列。",
        "一览表无此单位。",
    ],
    [
        "14-5-14",
        "上海睿泽股权投资管理有限公司",
        "申报表与工作总结均未单列。",
        "2015.09–2015.10，工程部资料员、施工员。来源是一览表，不是本轮.doc正文。",
    ],
    [
        "14-5-13",
        "新城控股集团股份有限公司上海第一分公司",
        "第5页：2015.06.08–2017.02.24，工程部高级经理。总结：2015.06–2017.02，工程部高级经理。",
        "2015.12–2017.02，工程部高级经理。与香溢紫郡项目起始月相同，与第5页起始月不同。",
    ],
    [
        "14-5-12",
        "上海万科长宁置业有限公司",
        "第5页：2018.12.10–2022.03.18，工程部总监。总结未单列，并入上海万科企业有限公司2017.11–2022.03。",
        "2017.03–2017.12，工程部总监。",
    ],
    [
        "14-5-11",
        "万科（上海）实业发展有限公司",
        "第5页写作上海万限实业有限公司，括注本单位：2017.11.06–2018.11.10，工程部总监。总结并入上海万科企业有限公司。",
        "2018.01–2019.03，工程部总监。",
    ],
    [
        "14-5-10",
        "上海万居企业管理有限公司",
        "申报表与工作总结均未单列。",
        "2019.04–2020.02，工程部总监。来源是一览表，不是本轮.doc正文。",
    ],
    [
        "14-5-9",
        "金科（上海）建筑设计有限公司",
        "第5页：2017.03.20–2017.05.29，工程部资深经理。总结并入金科地产集团有限公司华东区域公司2017.03–2017.10、工程总监。",
        "2020.03–2020.05，工程部总监。",
    ],
    [
        "14-5-8",
        "上海金科文化旅游发展集团有限公司",
        "第5页：2017.06.01–2017.10.26，工程部资深经理。总结并入上述金科华东区域公司。",
        "2020.06–2020.10，工程部总监。",
    ],
    [
        "14-5-7",
        "上海欣驭文化传播有限公司",
        "申报表与工作总结均未单列。",
        "2020.11–2020.12，工程部总监。来源是一览表，不是本轮.doc正文。",
    ],
    [
        "14-5-6",
        "上海弘阳汇商业管理有限公司",
        "申报表与工作总结均未单列。",
        "2021.01–2021.02，工程部总监。来源是一览表，不是本轮.doc正文。",
    ],
    [
        "14-5-5",
        "上海融之旅文化旅游发展有限公司",
        "第5页：2022.04.11–2023.05.13，工程部总监，括注融创文旅华东区域。总结同一区间，职务为工程总监。",
        "2021.03–2021.08，工程部总监。",
    ],
    [
        "14-5-4",
        "上海社通实业有限公司",
        "申报表与工作总结均未单列。",
        "2021.09–2022.07，工程部总监。来源是一览表，不是本轮.doc正文。",
    ],
    [
        "14-5-3",
        "上海苏科企业管理有限公司",
        "申报表与工作总结均未单列。",
        "2022.08–2023.04，工程部副总经理。一览表主要工作内容为钢结构施工技术管理。来源是一览表，不是本轮.doc正文。",
    ],
    [
        "14-5-2",
        "上海容钲建筑科技有限公司",
        "申报表与工作总结均未单列。",
        "2023.05–2023.08，工程部副总经理。来源是一览表，不是本轮.doc正文。",
    ],
    [
        "14-5-1",
        "上海徐汇绿地商业管理有限公司",
        "第5页：2023.06.07–2025.02.28，工程部副总经理。总结：2023.06–2025.02，工程部副总经理。",
        "2023.10–2024.07，工程部副总经理。",
    ],
]

PROJECTS: list[dict[str, object]] = [
    {
        "title": "中南海湾新城·金石国际广场（2011.07–2015.08）",
        "attr": "超高层五星级酒店公寓（26万㎡，120m）",
        "role": "施工技术管理",
        "breaks": [
            "攻克核心筒框架剪力墙施工难题，通过三维建模预演优化节点处理方案",
            "建立实时沉降监测系统，控制结构偏移误差＜3mm",
        ],
        "result": "节约工期45天，获江苏省优质工程奖",
        "note": (
            "来自工作总结。对应任职是江苏中南建设集团股份有限公司2011.07–2015.05、施工员/资料员，"
            "项目截止比任职截止晚3个月。江苏省优质工程奖只写在本项目，申报表奖惩栏没有这一项。"
        ),
    },
    {
        "title": "新城香溢紫郡（2015.12–2017.02）",
        "attr": "高端住宅（23万㎡）",
        "role": "工程高级经理",
        "breaks": [
            "创新桩基切顶高程控制法，降低地基处理成本23%",
            "研发混凝土裂缝预警系统，楼板开裂率下降至0.8‰",
        ],
        "result": "项目提前3个月交付，利润率提升12%",
        "note": (
            "来自工作总结。项目起止与一览表中新城控股集团股份有限公司上海第一分公司的任职一致。"
            "申报表第5页该单位自2015年06月08日起，早于项目起始月。职务与第5页“高级经理”相符。"
        ),
    },
    {
        "title": "合肥金科博翠天辰（2017.03–2017.10）",
        "attr": "高端住宅综合体（总建面约34万㎡，容积率2.5）",
        "role": "项目技术负责人（隶属金科集团华东区域公司）",
        "breaks": [
            "超高层住宅剪力墙结构抗震优化：主导结构模型风洞试验，优化型钢排布方案，使结构刚度分布均匀化，减少扭转效应18%",
            "采用BIM预拼装技术解决型钢与钢筋密集碰撞问题，缩短关键路径工期23天",
            "高端住宅绿色科技系统集成：优化管网穿剪力墙节点密封工艺，解决冷桥结露问题，系统能效比提升至4.8（国标≥3.8）",
            "嵌入金科“Life+云平台”，实现户内能耗智能调节",
        ],
        "result": "项目获合肥市绿色建筑示范工程认证",
        "note": (
            "来自工作总结，跨页处“示范工程”与“认证”已接回一句。"
            "同一时段申报表第5页的单位是金科（上海）建筑设计有限公司（2017.03.20–2017.05.29，资深经理）"
            "和上海金科文化旅游发展集团有限公司（2017.06.01–2017.10.26，资深经理）。"
            "一览表把这两家公司写在2020年，职务为工程部总监。本项目不改挂到2020年。"
        ),
    },
    {
        "title": "上海万科天空之城（2017.11–2022.02）",
        "attr": "商业综合体+住宅（42万㎡）",
        "role": "工程总监",
        "breaks": [
            "首创商业-住宅协同施工模型，解决交叉作业冲突",
            "应用自愈合混凝土技术于剪力墙节点",
            "构建社区级“零碳”能源系统，光伏覆盖率达32%",
            "研发模块化雨水回收装置，节水效率提升45%",
            "搭建社区参与式设计平台，客户满意度达92%",
        ],
        "result": "获LEED金级认证，商业区开业首年客流超预期40%；荣获上海万科产品创新研究院“可持续社区营造标杆项目奖”",
        "note": (
            "前两项技术突破，以及LEED金级认证、客流，来自工作总结。"
            "后三项技术突破、奖项名称来自申报表第6页业绩登记。"
            "业绩登记项目全称是“万科青浦区徐泾镇徐盈路西侧地块（01-06F2-04地块一、二期）”，"
            "起止2019年11月15日至2022年03月25日，作用为主持。"
            "同页专利号为ZL202020536480.4、ZL202020557884.1；总结另写发明专利《建筑工程管理用功能脚架》"
            "与实用新型专利《建筑工程防护围栏》，号码相同。"
            "奖惩栏该奖的获奖日是2019-02-24，早于业绩登记的起始日。"
            "总结把任职写成上海万科企业有限公司2017.11–2022.03。"
            "第5页对应的是万科（上海）实业发展有限公司2017.11.06–2018.11.10，以及上海万科长宁置业有限公司2018.12.10–2022.03.18。"
            "项目起止长于其中任何一段，本项不改挂。"
        ),
    },
    {
        "title": "杭州湾融创文旅城水路乐园项目（海盐）（2022.04–2023.05）",
        "attr": "高端抗震住宅",
        "role": "工程总监",
        "breaks": [
            "优化框架剪力墙抗震参数，地震模拟性能提升35%",
            "实施无人机巡检系统，质量验收效率提升60%",
        ],
        "result": "工作总结没有成果句。申报表业绩登记写获奖、专利、效益均为无。",
        "note": (
            "项目属性按工作总结。业绩登记的项目名是“杭州湾融创文旅城一水、陆乐园项目”，"
            "起止2018年02月03日至2023年05月11日，作用为主持。"
            "两处名称、属性、起始时间不同，不合并改写。"
            "第5页上海融之旅文化旅游发展有限公司为2022.04.11–2023.05.13、工程部总监，与总结的项目区间一致。"
            "一览表该单位为2021.03–2021.08。业绩登记的2018年起始与这两段都不重合。"
        ),
    },
    {
        "title": "苏州绿地江南映月花园（2023.06–2025.03）",
        "attr": "现状高端住宅项目",
        "role": "工程部副总经理",
        "breaks": [
            "应用BIM+3D打印技术实现异形钢结构精准安装",
            "开发碳足迹监测平台，施工碳排放降低28%",
        ],
        "result": "工作总结没有另写成果句。",
        "note": (
            "来自工作总结。原文编号也是“6”，与下一条业绩登记项目重复编号，此处按先后单列。"
            "上海徐汇绿地商业管理有限公司在第5页为2023.06.07–2025.02.28，在一览表为2023.10–2024.07。"
            "本项目起止比这两段都长。申报表第6页另有东太湖地块，名称不同，下一条单列。"
        ),
    },
    {
        "title": "绿地苏州东太湖旅游度假区（2020.10.15–2024.01.25）",
        "attr": "WJ-J-2020-013地块商住。原文未写建筑面积。",
        "role": "主持",
        "breaks": [
            "构建全流程技术标准体系，集成绿色建材与智能建造技术，单位面积改造成本降低18%，获集团技术创新标杆认证",
            "形成智慧工地管理系统思路，通过物联网与AI算法优化施工调度，降低人力成本15%、材料损耗率23%，并推动BIM技术应用达LOD500标准",
        ],
        "result": "累计节约项目成本2,800万元，投资回报率提升32%；荣获绿地控股集团技术委员会“城市更新卓越实践奖”",
        "note": (
            "来自申报表第6页业绩登记，课题名称为绿地苏州东太湖旅游度假区。"
            "与江南映月花园名称不同，不并成一个项目。"
            "奖惩栏该奖获奖日是2024-04-26，晚于本段截止日。获奖月落在第5页徐汇绿地任职区间内。"
        ),
    },
    {
        "title": "东台市养老服务中心地下车库装饰工程（2025.03起）",
        "attr": "适老化公共空间改造",
        "role": "施工技术负责人",
        "breaks": [
            "适老化地下空间无障碍导视：原文指出老年群体空间辨识能力弱，传统车库标识系统易引发迷失焦虑；列出地坪发光标线与AI语音联动、BIM模拟老人动线。两处冒号后没有继续写做法",
            "建立无人机实景扫描与BIM模型自动比对系统，偏差超5%自动预警",
        ],
        "result": "减少返工损失150万元（占预算8.3%），关键节点“零延误”",
        "note": (
            "来自工作总结，原文写“2025.3-至今”。"
            "申报表第5页江苏省建筑装饰设计研究院有限公司写至2025年07月10日，总结对该单位写“至今”。"
            "本批15份项目证明没有装饰院。不把截止日改成2025年07月10日或2025年12月。"
        ),
    },
]


def assert_public_only(text: str) -> None:
    for token in FORBIDDEN:
        if token in text:
            raise SystemExit(f"输出含不应出现的内容：{token}")


def html_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def render_table(headers: list[str], rows: list[list[str]], css_class: str = "") -> str:
    head = "".join(f"<th>{html_escape(cell)}</th>" for cell in headers)
    body = []
    for row in rows:
        cells = "".join(f"<td>{html_escape(cell)}</td>" for cell in row)
        body.append(f"<tr>{cells}</tr>")
    class_attr = f" class='{css_class}'" if css_class else ""
    return f"<table{class_attr}><thead><tr>{head}</tr></thead><tbody>{''.join(body)}</tbody></table>"


def render_projects() -> str:
    blocks = []
    for index, project in enumerate(PROJECTS, start=1):
        breaks = "".join(f"<li>{html_escape(str(item))}</li>" for item in project["breaks"])
        blocks.append(
            "<section class='project'>"
            f"<h3>{index}. {html_escape(str(project['title']))}</h3>"
            f"<p><strong>项目属性：</strong>{html_escape(str(project['attr']))}</p>"
            f"<p><strong>角色：</strong>{html_escape(str(project['role']))}</p>"
            f"<p><strong>技术突破：</strong></p><ul>{breaks}</ul>"
            f"<p><strong>成果：</strong>{html_escape(str(project['result']))}</p>"
            f"<p class='note'><strong>核对：</strong>{html_escape(str(project['note']))}</p>"
            "</section>"
        )
    return "\n".join(blocks)


def build_html() -> str:
    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8" />
<title>胡继刚职业履历与核心项目业绩</title>
<style>
  body {{
    margin: 32px auto;
    max-width: 1080px;
    color: #{INK};
    background: #fff;
    font-family: "WenQuanYi Micro Hei", "Noto Sans CJK SC", "Microsoft YaHei", sans-serif;
    line-height: 1.65;
  }}
  h1 {{ text-align: center; color: #{NAVY}; font-size: 28px; margin: 0 0 8px; }}
  h2 {{ color: #{NAVY}; font-size: 20px; margin: 28px 0 8px; border-bottom: 2px solid #{NAVY}; padding-bottom: 4px; }}
  h3 {{ color: #{NAVY}; font-size: 17px; margin: 0 0 8px; }}
  p {{ margin: 8px 0; }}
  .lead, .note {{ color: #{MUTED}; font-size: 14px; }}
  table {{ width: 100%; border-collapse: collapse; margin: 10px 0 6px; }}
  th, td {{ border: 1px solid #{LINE}; padding: 8px 6px; vertical-align: top; font-size: 13px; }}
  th {{ background: #{NAVY}; color: #fff; text-align: center; }}
  tbody tr:nth-child(even) td {{ background: #f4f7fb; }}
  ul {{ margin: 4px 0 8px 1.2em; padding: 0; }}
  .project {{ margin: 16px 0 20px; padding-top: 8px; border-top: 1px solid #d6dee8; }}
  table.proof td:first-child, table.proof th:first-child {{ white-space: nowrap; width: 7%; }}
</style>
</head>
<body>
  <h1>胡继刚职业履历与核心项目业绩</h1>
  <p class="lead">{html_escape(INTRO)}</p>
  <h2>一、职业履历</h2>
  <h3>1. 工作总结口径</h3>
  <p>时间、公司名称、职务。这一套与第二节项目时间线对应。</p>
  {render_table(SUMMARY_HEADERS, SUMMARY_ROWS)}
  <p class="note">{html_escape(SUMMARY_NOTE)}</p>
  <h3>2. 申报表第5页工作经历</h3>
  <p>同一人在申报表工作经历栏里的到日记录。中南建设不在本页。</p>
  {render_table(FORM_HEADERS, FORM_ROWS)}
  <p class="note">{html_escape(FORM_NOTE)}</p>
  <h3>3. 十五份项目证明对照</h3>
  <p>单位名称来自文件名。未单列单位的起止和职务来自此前工作经历一览表，不来自项目证明正文。</p>
  {render_table(PROOF_HEADERS, PROOF_ROWS, "proof")}
  <h2>二、核心项目业绩与技术突破</h2>
  <p class="lead">技术突破和成果只保留申报表、工作总结里的原句，不补写证明正文里尚未读到的内容。</p>
  {render_projects()}
</body>
</html>
"""
    assert_public_only(html)
    return html


def set_run_font(run, name: str, size: float, bold: bool = False, color: str | None = None) -> None:
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = name
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for key in ("w:ascii", "w:hAnsi", "w:eastAsia"):
        r_fonts.set(qn(key), name)


def shade_cell(cell, fill: str) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")


def set_cell_border(cell) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "4")
        node.set(qn("w:color"), LINE)


def add_paragraph(doc: Document, text: str, size: float = 12, bold: bool = False, color: str = INK, center: bool = False, space_after: float = 6):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(space_after)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    if center:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(text)
    set_run_font(run, "宋体", size, bold, color)
    return paragraph


def add_table(doc: Document, headers: list[str], rows: list[list[str]]) -> None:
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for col, header in enumerate(headers):
        cell = table.rows[0].cells[col]
        cell.text = ""
        run = cell.paragraphs[0].add_run(header)
        set_run_font(run, "宋体", 10.5, True, "FFFFFF")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        shade_cell(cell, NAVY)
        set_cell_border(cell)
    for r_index, row in enumerate(rows):
        for c_index, value in enumerate(row):
            cell = table.rows[r_index + 1].cells[c_index]
            cell.text = ""
            run = cell.paragraphs[0].add_run(value)
            set_run_font(run, "宋体", 10.5, False, INK)
            shade_cell(cell, "F4F7FB" if r_index % 2 else "FFFFFF")
            set_cell_border(cell)
    doc.add_paragraph()


def build_docx(path: Path) -> None:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.6)
    section.right_margin = Cm(1.6)
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.6)
    add_paragraph(doc, "胡继刚职业履历与核心项目业绩", 18, True, NAVY, True, 8)
    add_paragraph(doc, INTRO, 10.5, False, MUTED, False, 8)
    add_paragraph(doc, "一、职业履历", 16, True, NAVY, False, 6)
    add_paragraph(doc, "1. 工作总结口径", 14, True, NAVY, False, 4)
    add_paragraph(doc, "时间、公司名称、职务。这一套与第二节项目时间线对应。", 10.5, False, INK, False, 4)
    add_table(doc, SUMMARY_HEADERS, SUMMARY_ROWS)
    add_paragraph(doc, SUMMARY_NOTE, 10.5, False, MUTED, False, 8)
    add_paragraph(doc, "2. 申报表第5页工作经历", 14, True, NAVY, False, 4)
    add_paragraph(doc, "同一人在申报表工作经历栏里的到日记录。中南建设不在本页。", 10.5, False, INK, False, 4)
    add_table(doc, FORM_HEADERS, FORM_ROWS)
    add_paragraph(doc, FORM_NOTE, 10.5, False, MUTED, False, 8)
    add_paragraph(doc, "3. 十五份项目证明对照", 14, True, NAVY, False, 4)
    add_paragraph(doc, "单位名称来自文件名。未单列单位的起止和职务来自此前工作经历一览表，不来自项目证明正文。", 10.5, False, INK, False, 4)
    add_table(doc, PROOF_HEADERS, PROOF_ROWS)
    add_paragraph(doc, "二、核心项目业绩与技术突破", 16, True, NAVY, False, 6)
    add_paragraph(doc, "技术突破和成果只保留申报表、工作总结里的原句，不补写证明正文里尚未读到的内容。", 10.5, False, MUTED, False, 6)
    for index, project in enumerate(PROJECTS, start=1):
        add_paragraph(doc, f"{index}. {project['title']}", 13, True, NAVY, False, 2)
        add_paragraph(doc, f"项目属性：{project['attr']}", 12, True, INK, False, 2)
        add_paragraph(doc, f"角色：{project['role']}", 12, True, INK, False, 2)
        add_paragraph(doc, "技术突破：", 12, True, INK, False, 2)
        for item in project["breaks"]:
            add_paragraph(doc, f"▶ {item}", 12, False, INK, False, 1)
        add_paragraph(doc, f"成果：{project['result']}", 12, True, INK, False, 2)
        add_paragraph(doc, f"核对：{project['note']}", 10.5, False, MUTED, False, 8)
    text = "\n".join(p.text for p in doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            text += "\n" + "\n".join(cell.text for cell in row.cells)
    assert_public_only(text)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


def write_sheet(
    ws,
    title: str,
    headers: list[str],
    rows: list[list[str]],
    widths: list[int],
    row_height: int = 36,
) -> None:
    ws["A1"] = title
    ws["A1"].font = Font(name="宋体", size=16, bold=True, color=NAVY)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.row_dimensions[1].height = 28
    header_fill = PatternFill("solid", fgColor=NAVY)
    alt = PatternFill("solid", fgColor="F4F7FB")
    border = Border(
        left=Side(style="thin", color=LINE),
        right=Side(style="thin", color=LINE),
        top=Side(style="thin", color=LINE),
        bottom=Side(style="thin", color=LINE),
    )
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(3, col, header)
        cell.font = Font(name="宋体", size=11, bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    for r_index, row in enumerate(rows):
        excel_row = 4 + r_index
        for c_index, value in enumerate(row, start=1):
            cell = ws.cell(excel_row, c_index, value)
            cell.font = Font(name="宋体", size=10, color=INK)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border = border
            if r_index % 2:
                cell.fill = alt
        ws.row_dimensions[excel_row].height = row_height
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.4)
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{3 + len(rows)}"
    ws.print_title_rows = "1:3"
    ws.page_setup.horizontalCentered = True


def build_xlsx(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "工作总结履历"
    write_sheet(ws, "职业履历（工作总结口径）", SUMMARY_HEADERS, SUMMARY_ROWS, [8, 22, 46, 28])
    form = wb.create_sheet("申报表工作经历")
    write_sheet(form, "职业履历（申报表第5页）", FORM_HEADERS, FORM_ROWS, [8, 26, 46, 24])
    proof = wb.create_sheet("项目证明对照")
    write_sheet(proof, "十五份项目证明对照", PROOF_HEADERS, PROOF_ROWS, [14, 36, 55, 42], row_height=72)
    projects = wb.create_sheet("核心项目")
    project_rows = []
    for index, project in enumerate(PROJECTS, start=1):
        project_rows.append(
            [
                str(index),
                str(project["title"]),
                str(project["attr"]),
                str(project["role"]),
                "\n".join(f"▶ {item}" for item in project["breaks"]),
                str(project["result"]),
                str(project["note"]),
            ]
        )
    write_sheet(
        projects,
        "核心项目业绩与技术突破",
        ["序号", "项目", "项目属性", "角色", "技术突破", "成果", "核对"],
        project_rows,
        [8, 36, 28, 24, 55, 36, 42],
        row_height=96,
    )
    note = wb.create_sheet("整理说明", 0)
    note["A1"] = "整理说明"
    note["A1"].font = Font(name="宋体", size=16, bold=True, color=NAVY)
    note["A2"] = INTRO
    note["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    note["A2"].font = Font(name="宋体", size=11, color=INK)
    note.column_dimensions["A"].width = 120
    note.row_dimensions[2].height = 72
    blob = INTRO + SUMMARY_NOTE + FORM_NOTE
    for row in SUMMARY_ROWS + FORM_ROWS + PROOF_ROWS + project_rows:
        blob += "".join(row)
    assert_public_only(blob)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    html_path = OUT_DIR / HTML_NAME
    html_path.write_text(build_html(), encoding="utf-8")
    build_docx(OUT_DIR / DOCX_NAME)
    build_xlsx(OUT_DIR / XLSX_NAME)
    print(html_path)
    print(OUT_DIR / DOCX_NAME)
    print(OUT_DIR / XLSX_NAME)


if __name__ == "__main__":
    main()
