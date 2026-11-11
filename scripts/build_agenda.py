#!/usr/bin/env python3
"""根据《2026上海国际智慧康养产业博览会》拟邀方案，生成分论坛详细议程。

正式时段、题目、主办单位和拟邀人选均来自源方案。
分钟脚本、圆桌追问、控时和缺席顺延，是在该框架内做的执行设计。
斜线两侧是替代人选，每个发言席只保留一名最终主讲人。
"""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "deliverables"
FONT = "微软雅黑"

PINE = "0F5C58"
PINE_RGB = (0x0F, 0x5C, 0x58)
INK = "1B3A4B"
INK_RGB = (0x1B, 0x3A, 0x4B)
GOLD = "8C6A32"
GOLD_RGB = (0x8C, 0x6A, 0x32)
CREAM = "FBF6EE"
MINT = "E7F2F1"
ALT = "F4F8F7"
RED = "8C3A32"
RED_RGB = (0x8C, 0x3A, 0x32)
LINE = "D5DDD8"
WHITE = "FFFFFF"


def to_min(hm: str) -> int:
    hour, minute = hm.split(":")
    return int(hour) * 60 + int(minute)


def duration(start: str, end: str) -> int:
    return to_min(end) - to_min(start)


# 正式议程。beats 必须首尾相接，并铺满该环节。
SESSIONS = [
    {
        "start": "13:00",
        "end": "13:08",
        "name": "开幕",
        "title": "AI重塑养老未来",
        "kind": "开幕",
        "goal": "用八分钟把场次名称、讨论边界和全场节奏说完，不提前展开任何一家企业的产品。",
        "primary": "徐洁平",
        "backup": "专业主持人。源方案未提供徐洁平的职务，确认函以组委会最终简介为准。",
        "outline": [
            "主持人只负责开场、串场和控时，不发表主题演讲。",
            "点名范围限于已经入座的主旨嘉宾和圆桌嘉宾。未确认、未到场的拟邀名单不在台上宣读。",
        ],
        "beats": [
            ("13:00", "13:01", "主持", "音乐收住，上台报出全称：2026上海国际智慧康养产业博览会，本场为“AI重塑养老未来——2026银发科技与具身智能产业创新峰会”。"),
            ("13:01", "13:03", "主持", "限定讨论边界：“今天下午不讨论抽象的养老未来。我们只讨论机器人、大模型和智慧家庭怎样进入养老院和家庭，以及谁愿意为此付费。”请观众将手机调至静音。"),
            ("13:03", "13:06", "主持", "用五个组别介绍在场结构：老龄研究、具身智能、人工智能与物联网、养老机构、产业资本与银行。只点名已入座的主旨和圆桌嘉宾。"),
            ("13:06", "13:08", "主持", "预告节奏：四场主旨、一场买单方实践、十分钟看设备、四十五分钟圆桌、两到三家成果秀、意向发布。声明16:30准点合影，超时由主持收束。"),
        ],
    },
    {
        "start": "13:08",
        "end": "13:18",
        "name": "开幕致辞",
        "title": "共建长三角银发科技产业生态",
        "kind": "致辞",
        "goal": "用主办方和行业协会的口径，说明这场分论坛放在博览会中的原因。致辞不承担产品介绍。",
        "primary": "第一位：上海市养老服务和老龄产业协会代表",
        "backup": "第二位在上海现代服务业联合会、中国工商银行上海市分行之间二选一。上海丰之米会展有限公司负责会务，不单独致辞。若只确认一位，13:13直接进入主旨一。",
        "outline": [
            "协会代表四分钟，只讲三句：本场为何放在博览会里；希望企业带来可试点的场景；协会可协助对接真实机构。",
            "第二位若是联合会，讲产业协同；若是工商银行上海市分行，讲科技与养老金融如何配套。不发布具体信贷产品。",
            "分论坛主办方上海星智耀科技有限公司、联合承办方不占用开幕致辞，代表放在总结和前排。",
        ],
        "beats": [
            ("13:08", "13:09", "主持", "介绍第一位致辞人的单位和姓名，不念长简介。优先为上海市养老服务和老龄产业协会代表。"),
            ("13:09", "13:13", "致辞人", "完成四分钟致辞。计时员在第三分钟举黄牌，到点举红牌。主持上前半步致谢，不让致辞进入第五分钟。"),
            ("13:13", "13:17", "致辞人", "第二位致辞，同样四分钟。未确认第二位时，主持宣布“时间留给第一场主旨”，本段取消，主旨一提前开始。"),
            ("13:17", "13:18", "主持", "致谢，并说明第一场主旨负责把需求边界立住，再把话筒交给主讲人。"),
        ],
    },
    {
        "start": "13:18",
        "end": "13:38",
        "name": "主旨一",
        "title": "“十五五”银发经济与AI时代的长寿社会",
        "kind": "主旨",
        "goal": "先把老龄社会的真实需求和支付边界讲清楚，给后面所有技术演讲一个共同前提。",
        "primary": "彭希哲（复旦大学老龄研究院院长、文科资深教授），题目《“十五五”银发经济与长寿社会发展的新机遇》",
        "backup": "吴玉韶（复旦大学老龄研究院副院长、教授），题目改为《中国养老服务变革与银发科技产业新趋势》。再备选胡湛（副院长、银龄未来实验室主任），题目改为《数智养老与AI向善：面向未来的中国式养老》。只上一人。",
        "outline": [
            "用两分钟说明：长寿社会已经同时摆在家庭、机构和支付制度面前，不做成人口数据展览。",
            "把能形成产业的需求收成三类：主动健康管理、居家安全与辅助、认知和康复。明确哪些老人不是人形机器人的第一批用户。",
            "给出三条边界：向善，不替代必要的人际照护；可及，区县机构和普通家庭用得起；可支付，存在明确采购主体或支付制度。",
            "最后四分钟向后几位技术嘉宾提问：哪一项服务有可能先成为标准配置。",
            "若由吴玉韶主讲，中间改为养老服务体系变革，以及机构、社区、居家对科技的不同接受度。",
            "若由胡湛主讲，减少宏观人口，重点讲数智养老、AI向善，以及老年人使用技术时已经看到的障碍。",
        ],
        "beats": [
            ("13:18", "13:19", "主持", "介绍姓名、一个职务和题目，不超过四十秒。开场前已与嘉宾确认十八分钟时限。学术嘉宾不播放产品片。"),
            ("13:19", "13:37", "主讲", "完成十八分钟主旨。剩两分钟时黄牌，到点红牌。主持不在中途提问，也不插入企业案例。"),
            ("13:37", "13:38", "主持", "用一句话收束：“后面的技术演讲，请回应刚才三个词：可及、可支付、向善。”不开放现场提问。"),
        ],
    },
    {
        "start": "13:38",
        "end": "14:00",
        "name": "主旨二",
        "title": "具身智能走进养老：从康复机器人到家庭陪护",
        "kind": "主旨",
        "goal": "这是下午唯一预留演示的主旨。观众要看懂一个养老动作，以及护理员如何接手。",
        "primary": "顾捷（傅利叶智能创始人、CEO），题目《从康复机器人到具身智能：AI如何改变未来养老》",
        "backup": "按源方案顺序只替补一人：邓泰华（智元机器人创始人、董事长兼CEO），题目《具身智能产业化：人形机器人走向现实场景》；彭志辉（智元机器人联合创始人、总裁、CTO），题目《人形机器人进入家庭：从技术突破到规模化应用》；或傅利叶智能相关高管，沿用本场题目讲落地。王兴兴在嘉宾池中，题目为《机器人走入千家万户：具身智能的未来十年》，不在这条已指定的替补顺序里；若四位均未确认，由组委会另行决定，不新增第五场主旨。",
        "outline": [
            "十三分钟回答四件事：康复机器人已经用在什么场景；走到具身智能后真正多出的能力；养老院现在能接的任务；家庭场景还有哪些不能承诺。",
            "安全、责任划分和护理员培训放在商业化里讲，不放到售后补充。",
            "七分钟只演示一个养老相关动作。不演示与养老无关的整机能力。",
            "无法带设备时，改为九十秒实景视频，其余时间讲已部署流程和一次失败案例，不改成品牌宣传片。",
            "邓泰华出场仍按“已落地、可试点、还不该进院”三段讲。彭志辉出场必须单列家庭场景尚未成熟的部分。",
        ],
        "beats": [
            ("13:38", "13:39", "主持", "点明这是下午唯一预留演示的主旨：“前面十三分钟讲边界，后面七分钟只看一个养老场景里的动作。”"),
            ("13:39", "13:52", "主讲", "完成十三分钟演讲。计时从演讲开始单独计算，不把上台调试算进演示。"),
            ("13:52", "13:59", "主讲", "演示七分钟。演示前用一句话说明急停方式。动作失控时主持有权叫停，损失的时间不顺延。"),
            ("13:59", "14:00", "主持", "收束：“谁来买单放到十四点三十六分，这一场不插问。”设备退至展示位，不在讲台停留。"),
        ],
    },
    {
        "start": "14:00",
        "end": "14:18",
        "name": "主旨三",
        "title": "大模型与智能体：重构未来健康管理与养老服务",
        "kind": "主旨",
        "goal": "把智能体和机器人分开：机器人执行动作，智能体管理任务、记录和家属协同，并讲清责任边界。",
        "primary": "邱锡鹏（复旦大学教授、上海创智学院全时导师），题目《从大模型到智能体：AI康养的下一代交互技术》",
        "backup": "肖仰华（复旦大学教授、上海市数据科学重点实验室主任），议题为AI智能体、知识图谱与智慧健康服务；徐立（商汤科技董事长兼CEO），题目《大模型与具身智能：AI赋能银发科技新场景》。两人均未确认时，可由相关业务代表上场；若需要医疗健康落地，尚海龙（商汤医疗联合创始人兼首席战略官）可讲医疗大模型如何支持健康管理，不做产品发布。",
        "outline": [
            "说明连续照护不能停在问答机器人。",
            "适合承担的四项：评估记录、用药和复诊提醒、照护计划拆解、家属同步。不适合承担的两项：替代诊疗、替代夜间值守责任。",
            "讲清数据来源、家属是否可见、错误建议由谁负责。",
            "收在一个可以试点的流程上，不收在模型参数上。",
            "肖仰华出场时，以知识图谱和健康服务流程为主线。徐立出场时仍要回答责任边界。",
        ],
        "beats": [
            ("14:00", "14:01", "主持", "介绍主讲人，并预告：“请把机器人已经讲过的动作放下，这一场只讲判断、记录和协同。”"),
            ("14:01", "14:17", "主讲", "完成十六分钟演讲。不播放与健康管理无关的通用大模型演示。"),
            ("14:17", "14:18", "主持", "留下圆桌要用的问题：“智能体给出的建议，谁负责？”不开放提问。"),
        ],
    },
    {
        "start": "14:18",
        "end": "14:36",
        "name": "主旨四",
        "title": "万物互联的养老新时代：AIoT与智慧家庭",
        "kind": "主旨",
        "goal": "讲清家庭和机构里要感知什么、报警交给谁，以及低成本是否把误报后的人工也算进去。",
        "primary": "钱鹏鹤（移远通信董事长兼CEO），议题为AIoT、智能感知与智慧养老物联网络",
        "backup": "王学集（涂鸦智能创始人兼CEO），议题为AI Home与未来智慧养老家庭；张瑞安（乐鑫科技创始人、CEO），议题为AIoT芯片与低成本智能健康设备。三位均未确认时，才由萤石网络、海康威视智慧养老或大华股份智慧物联的拟邀负责人上场，且必须讲已有的养老或家庭看护场景。",
        "outline": [
            "开场就说明感知对象：跌倒、离床、长时间不动、用药、烟火和出入。不从芯片参数讲起。",
            "家庭、社区、机构三级数据最终交到谁手里，误报之后谁上门。",
            "低成本同时计算设备、安装和误报造成的人工。",
            "产品原则：不给护理员增加第二块需要时刻盯着的屏幕。",
            "说明感知是机器人进入家庭之前要先解决的条件。",
            "张瑞安出场时可以讲到终端价格，但至少一半时间留给养老设备形态。",
        ],
        "beats": [
            ("14:18", "14:19", "主持", "介绍主讲人，并限制开场：“请从养老场景讲起，不要从参数讲起。”"),
            ("14:19", "14:35", "主讲", "完成十六分钟演讲。如需展示终端，只展示与刚才点名的感知场景直接相关的设备。"),
            ("14:35", "14:36", "主持", "收束并交给场景方：“设备可以很便宜。下面请机构回答，便宜是不是就会买。”"),
        ],
    },
    {
        "start": "14:36",
        "end": "14:50",
        "name": "产业实践",
        "title": "谁来为养老机器人买单",
        "kind": "实践",
        "goal": "这十四分钟不讲技术趋势，只留下三句能带进圆桌的话：愿为什么付费、试点条件、明确不买什么。",
        "primary": "邱建伟（泰康之家首席执行官），角色为养老集团战略及智慧康养实践",
        "backup": "赵一洪（福寿康联合创始人、副总裁），侧重AI与居家养老、长期照护；康剑（福寿康集团产品及供给负责人、上海分公司总经理），侧重智慧护理技术的采购与应用。其后为已确认且拥有采购或试点决策权的机构负责人：瞿文，复星康养、星堡养老，泰康之家·申园，上海市第一社会福利院，九如城养老集团。陈楠的议题是AI认知健康与认知障碍筛查，不作为本场第一替补。",
        "outline": [
            "已经引进并留下的技术是什么，撤掉的技术是什么。",
            "采购时谁有否决权：运营、护理、家属、财务、监管。",
            "能接受的方式：买断、租赁、按效果付费或联合试点。可以说量级，不要求公开报价。",
            "向在场机器人企业提一个今天可以回应的请求。",
        ],
        "beats": [
            ("14:36", "14:37", "主持", "定调：“这一段不讲趋势。请直接讲买过什么、退过什么、怎样才愿意试点。”"),
            ("14:37", "14:48", "场景方", "完成十一分钟发言。主持在开场前把三句收束结构告诉发言人，避免讲成企业介绍。"),
            ("14:48", "14:50", "主持", "对着题板复述三句：愿意付费的类别、试点条件、明确不买的类别。这三句原样带进圆桌。"),
        ],
    },
    {
        "start": "14:50",
        "end": "15:00",
        "name": "茶歇与展示",
        "title": "机器人展示交流",
        "kind": "茶歇",
        "goal": "让观众看见设备，同时保证15:00圆桌六人已经入座。茶歇不是第二场发布会。",
        "primary": "与创新成果秀相同的已确认企业，最多三家",
        "backup": "设备未进场时取消展示，保留十分钟离座交流。圆桌嘉宾仍须在14:58入座。不临时增加展示企业。",
        "outline": [
            "展示位不超过三个，且与15:45成果秀为同一批企业，避免茶歇变成另一份名单。",
            "只看、只许问一句，企业人员不在茶歇区做产品发布。",
            "话筒和席卡在12:40已经完成检查，茶歇期间不再试音。",
        ],
        "beats": [
            ("14:50", "14:52", "主持", "宣布规则：“展示区只看、只问，不做发布。三点整开始圆桌，六位嘉宾先入座，观众三点前回到座位。”"),
            ("14:52", "14:58", "会务", "观众离座观看。每家设备旁保留一名企业人员和一名现场安全人员。机器人低速、有人看护，不进入观众座位区。"),
            ("14:58", "15:00", "会务", "清场回座。圆桌六人按席卡入座，主持就位。到点立即开始，不等仍在展示区的观众。"),
        ],
    },
    {
        "start": "15:00",
        "end": "15:45",
        "name": "高端圆桌",
        "title": "机器人进入养老院：技术成熟了，商业模式成熟了吗",
        "kind": "圆桌",
        "goal": "六个人用四十五分钟判断商业模式是否成熟，并留下可以进场的任务、试点条件和会后只促成的那一类合作。",
        "primary": "主持仍为徐洁平或专业主持人。胡继刚为第一席，不兼任主持。",
        "backup": "任一席位未确认则空出，不临时改成产品介绍。第四席和第六席至少各有一人；两席同时空缺时，圆桌改为“场景问诊”，由胡继刚与在场的技术方、学者讨论，时长仍止于15:45。",
        "outline": [
            "第一席胡继刚：高校成果转化、智慧居住、产业场景与资本协同。职务以组委会确认简介为准，源方案未单列头衔。",
            "第二席顾捷，或傅利叶智能业务负责人。若顾捷已讲主旨二，改由尚未发言的技术负责人，顺序为傅利叶康复业务、智元具身智能业务、云深处朱秋国、钛虎易港、擎朗智能业务负责人。科沃斯钱东奇或钱程仅在需要家庭服务机器人视角、且前面人选未到时使用。",
            "第三席胡湛，或复旦大学老龄研究专家。胡湛若已讲主旨一，改吴玉韶；两位都已发言时，由肖仰华或封进补席。封进更适合长期护理保险，也可改作第六席的学术补充。",
            "第四席邱建伟、赵一洪、康剑，或养老集团负责人。与14:36产业实践不使用同一人。",
            "第五席王学集、钱鹏鹤，或AIoT企业负责人。已讲主旨四的人退出，改张瑞安、萤石、海康威视智慧养老或大华智慧物联中尚未发言者。",
            "第六席为产业投资基金或银行科技金融负责人，须是有项目权限的合伙人、投资总监或业务负责人。不安排纯品牌推广人员。",
        ],
        "beats": [
            ("15:00", "15:03", "主持", "定题并宣布规则：“每人先讲四分钟，互不打断。之后只有三轮追问。本场不开放观众提问。”把题板上的三句买单结论让六人看见。"),
            ("15:03", "15:07", "第一席", "胡继刚从场景和成果转化开场：哪些居住空间真正需要机器人，高校和企业怎样签试点，资本要看到的是场景方的下一单，而不是一段演示视频。"),
            ("15:07", "15:11", "第二席", "技术方把能力分成三级：今天可以卖、可以做九十天试点、还不该进养老院。必须举一个已经发生的失败或撤回，不能只重复主旨二的演示。"),
            ("15:11", "15:15", "第三席", "学者回答老人、家属、护理员谁是用户，哪些环节必须留下人。指出一个今天听到的伪需求。"),
            ("15:15", "15:19", "第四席", "机构方回答买或不买，并给出试点门槛：培训天数、责任保险、故障响应时间和最短试用周期。不讲企业文化。"),
            ("15:19", "15:23", "第五席", "AIoT方说明机器人进院之前要先布什么。明确传感器和机器人不能各做一套报警。"),
            ("15:23", "15:27", "第六席", "资本或银行说明什么证据值得投或值得贷，并区分股权、设备租赁、采购分期和保险支付。缺数据就说明缺哪三项。"),
            ("15:27", "15:32", "主持", "追问第二席和第四席：“如果按月租赁并且包含驻场维护，是否愿意用一个楼层做九十天试点？”机构方先回答愿意或不愿意，再讲条件。"),
            ("15:32", "15:37", "主持", "追问第三席和第五席：“家属会先为摔倒预警付费，还是先为陪伴付费？”感知企业按这个答案说明要改的一项功能。"),
            ("15:37", "15:42", "主持", "追问第六席和第一席：“如果会后只促成一单，这一单应该是股权、试点合同，还是金融方案？”允许两人答案不同，主持不把分歧收成空话。"),
            ("15:42", "15:45", "主持", "面对观众收三句：今天可以进场的任务、机构提出的试点条件、四点十五分只接收对得上这两句的合作。"),
        ],
    },
    {
        "start": "15:45",
        "end": "16:15",
        "name": "创新成果秀",
        "title": "未来养老的AI新产品与应用场景",
        "kind": "成果秀",
        "goal": "只保留已经确认的两到三家。每家八分钟，留下一个场景、一个动作、一种卖法和一条对接需求。",
        "primary": "第一家傅利叶智能；第二家智元机器人；第三家强脑科技",
        "backup": "强脑未确认时，由一家尚未做主旨四的AIoT企业递补。确认不足三家就减场，不加未确认项目，也不把空出的时间分给第四家。源方案写明的范围是傅利叶、智元、强脑及AIoT创新企业中已确认的2—3家。",
        "outline": [
            "每家八分钟的内部分配为：主持引入约四十秒，场景四十秒，演示约两分二十秒，说明卖给谁约一分四十秒，最后一分二十秒只提一个对接需求。到点时只保留需求的最后一句。",
            "傅利叶若已在主旨二演示过同一动作，成果秀改由康复业务团队讲机构部署：人员、设备和收费方式，不再重复该动作。",
            "智元由未在主旨二发言的业务或产品负责人出场，不把董事长的主旨再讲一遍。",
            "强脑必须说明适用人群和医疗边界。韩璧丞可以本人出场，也可以由业务负责人出场；不因此加开一场主旨。",
            "16:10之后的五分钟只用于消化超时和口头汇总需求。三家都准点时，主持用三分钟复述需求，剩余两分钟入座。",
        ],
        "beats": [
            ("15:45", "15:46", "主持", "开场：“只留已经确认的两到三家。每家八分钟，请观众只看护理员能否接手。”宣布第一家名称和演示任务。"),
            ("15:46", "15:54", "第一家", "傅利叶智能按固定结构完成。演示前重复急停方式。到点未讲完需求时，主持只留最后一句。"),
            ("15:54", "16:02", "第二家", "智元机器人完成一个养老或服务动作。与主旨二重复的整机介绍直接跳过。"),
            ("16:02", "16:10", "第三家", "强脑科技或递补的AIoT企业完成。第三家未确认时，本段取消，16:02进入汇总，不另找企业填满。"),
            ("16:10", "16:15", "主持", "把各家留下的一条需求念完。超时已从这五分钟扣除。不允许加演。需求写进工作人员的登记板，供下一段使用。"),
        ],
    },
    {
        "start": "16:15",
        "end": "16:25",
        "name": "供需对接",
        "title": "科技与康养产业供需对接及合作意向发布",
        "kind": "对接",
        "goal": "十分钟只发布已经对得上圆桌结论的意向。交换名片留到16:30之后的离场区，不写入正式议程。",
        "primary": "会前选定的三个发布席：一个场景需求、一个临床或康复验证需求、一个科技企业的试点请求",
        "backup": "发布席不足三个时按实际数量结束，不现场征集。工商银行或在场基金只有在确认设咨询席时，才用一分钟告知位置。",
        "outline": [
            "发布席从已报名单位中选定，条件是有明确场景、采购、试点或联合验证。纯宣传材料不占用这十分钟。",
            "场景席优先使用尚未在14:36和圆桌发言的机构：上海市第一社会福利院、泰康之家·申园、九如城、福寿康、复星康养或星堡中已确认者。",
            "临床席使用上海相关康复医院的拟邀负责人。科技企业的请求必须对应圆桌上已经说出的试点条件。",
            "每家只说三句：我要什么，我能提供什么，会后找谁。",
        ],
        "beats": [
            ("16:15", "16:16", "主持", "说明规则：“每家两分钟，只说我要什么、我能提供什么、会后找谁。不对着名片做第二轮演讲。”"),
            ("16:16", "16:18", "发布席一", "养老机构或社区发布场景需求。工作人员同步写上登记板。"),
            ("16:18", "16:20", "发布席二", "康复医院或临床单位发布验证需求。未确认则跳过，时间不转给下一家做广告。"),
            ("16:20", "16:22", "发布席三", "科技企业发布一条试点请求，主持若听到与圆桌条件明显不符，直接请其收在一句“会后再约”。"),
            ("16:22", "16:24", "主持", "若工商银行上海市分行或在场基金已设咨询席，用一分钟告知位置，并说明今天不收商业计划书，只登记试点和融资需求。其余时间发放对接卡。未设咨询席则整段用于发登记卡。"),
            ("16:24", "16:25", "主持", "宣布离场区登记保留到16:45。会后由分论坛主办单位按登记表回访，回访时限会前由主办方写进确认函，不在现场临时承诺。"),
        ],
    },
    {
        "start": "16:25",
        "end": "16:30",
        "name": "总结及合影",
        "title": "论坛总结及合影",
        "kind": "总结",
        "goal": "只重复圆桌已经形成的三句话，然后完成合影。不开启新议题，不念长篇致谢。",
        "primary": "主承办方代表",
        "backup": "主办方代表未能留台时，由主持人按题板原文复述，不另请未参会的嘉宾代总结。",
        "outline": [
            "总结不超过两分半，只确认可进场的任务、试点条件和会后对接范围。",
            "合影控制在十五人以内：主持、已发言的致辞代表、主旨嘉宾、圆桌六席、成果秀企业各一人、主承办方代表。每家企业只上一人。",
            "观众不登台。合影背景使用本场主题，不临时更换成某一家企业的主视觉。",
        ],
        "beats": [
            ("16:25", "16:28", "主办代表", "完成总结。主持在两分半时准备接话。致谢不超过三句，点到协会、大会主办和联合承办即可。"),
            ("16:28", "16:30", "会务", "按名单登台合影。摄影只在此段拍摄集体照，不在成果秀演示区穿行。16:30散台。"),
        ],
    },
]


PRE_EVENT = [
    ("12:20", "12:30", "会务", "工作人员就位，签到台开启，席卡按最终确认名单摆放。未确认嘉宾不提前印上席卡。"),
    ("12:30", "12:40", "会务", "嘉宾签到、领取资料和对接卡。资料只含对外议程，不含仍在变动的替补顺序。"),
    ("12:40", "12:50", "会务", "展示设备完成安全检查。圆桌话筒、题板和计时牌试音、试位。此段之后不再试音。"),
    ("12:50", "12:55", "主持", "已确认的发言嘉宾在休息室对表，确认本人出场顺序、时限和是否携带演示。"),
    ("12:55", "13:00", "会务", "观众入座。12:58主持与计时员最后核对出场顺序。13:00准点，不等迟到嘉宾。"),
]

POST_EVENT = [
    ("16:30", "16:45", "会务", "离场区接受对接卡登记。机器人先断电，再撤场。16:45关闭登记，未登记事项不进入主办方的正式回访。"),
]


GUESTS = [
    ("主旨一", "1", "彭希哲", "复旦大学老龄研究院院长、文科资深教授", "《“十五五”银发经济与长寿社会发展的新机遇》", "优先。确认后取消同场其他主讲。"),
    ("主旨一", "2", "吴玉韶", "复旦大学老龄研究院副院长、教授", "《中国养老服务变革与银发科技产业新趋势》", "第一备选。"),
    ("主旨一", "3", "胡湛", "复旦大学老龄研究院副院长、银龄未来实验室主任", "《数智养老与AI向善：面向未来的中国式养老》", "第二备选。若未讲主旨，可保留圆桌第三席。"),
    ("主旨二", "1", "顾捷", "傅利叶智能创始人、CEO", "《从康复机器人到具身智能：AI如何改变未来养老》", "优先。若已讲主旨，圆桌第二席改由本公司业务负责人。"),
    ("主旨二", "2", "邓泰华", "智元机器人创始人、董事长兼CEO", "《具身智能产业化：人形机器人走向现实场景》", "备选。不与彭志辉同时上台。"),
    ("主旨二", "3", "彭志辉", "智元机器人联合创始人、总裁、CTO", "《人形机器人进入家庭：从技术突破到规模化应用》", "备选。家庭场景必须单列尚未成熟的部分。"),
    ("主旨二", "4", "傅利叶相关高管", "傅利叶智能", "沿用主旨二题目，讲落地", "邓泰华、彭志辉均未确认时使用。"),
    ("主旨二", "—", "王兴兴", "宇树科技创始人、董事长", "《机器人走入千家万户：具身智能的未来十年》", "在嘉宾池中，不在源方案已写明的主旨二替补顺序内。"),
    ("主旨三", "1", "邱锡鹏", "复旦大学教授、上海创智学院全时导师", "《从大模型到智能体：AI康养的下一代交互技术》", "优先。"),
    ("主旨三", "2", "肖仰华", "复旦大学教授、上海市数据科学重点实验室主任", "AI智能体、知识图谱与智慧健康服务", "备选。未讲主旨时，也可补圆桌第三席。"),
    ("主旨三", "3", "徐立", "商汤科技董事长兼CEO", "《大模型与具身智能：AI赋能银发科技新场景》", "备选。"),
    ("主旨三", "4", "尚海龙", "商汤医疗联合创始人兼首席战略官", "医疗大模型与智慧健康应用", "仅当前三位未确认、且需要医疗健康口径时，作为业务代表。"),
    ("主旨四", "1", "钱鹏鹤", "移远通信董事长兼CEO", "AIoT、智能感知与智慧养老物联网络", "优先。已讲主旨则退出圆桌第五席。"),
    ("主旨四", "2", "王学集", "涂鸦智能创始人兼CEO", "AI Home与未来智慧养老家庭", "备选。"),
    ("主旨四", "3", "张瑞安", "乐鑫科技创始人、CEO", "AIoT芯片与低成本智能健康设备", "备选。至少一半时间讲养老设备形态。"),
    ("主旨四", "4", "萤石 / 海康威视 / 大华", "家庭看护、智慧养老或智慧物联业务负责人", "家庭看护、安全监测、智慧社区", "仅当前三位均未确认时上场，只上一人。"),
    ("产业实践", "1", "邱建伟", "泰康之家首席执行官", "养老集团战略、采购与支付", "优先。已讲本场则不占圆桌第四席，也不占金融席。"),
    ("产业实践", "2", "赵一洪", "福寿康联合创始人、副总裁", "AI+居家养老及长期照护", "备选。"),
    ("产业实践", "3", "康剑", "福寿康集团产品及供给负责人、上海分公司总经理", "智慧护理技术采购与应用", "备选。"),
    ("产业实践", "4", "瞿文等机构负责人", "福苑康养，复星康养、星堡，申园，第一社会福利院，九如城", "有采购或试点决策权的机构实践", "按确认顺序只上一人。未上台者可竞争意向发布的场景席。"),
    ("圆桌第一席", "1", "胡继刚", "职务以组委会确认简介为准", "高校成果转化、智慧居住、产业场景及资本协同", "固定第一席，不兼任主持。"),
    ("圆桌第二席", "2", "顾捷或机器人业务负责人", "傅利叶优先，其后为智元业务、朱秋国、易港、擎朗", "技术成熟度：能卖、能试点、不该进院", "与主旨二错开人选。钱东奇、钱程仅作家庭服务机器人的后位备选。"),
    ("圆桌第三席", "3", "胡湛或老龄研究专家", "复旦大学老龄研究院", "真实需求、主动健康、AI向善", "与主旨一错开人选。"),
    ("圆桌第四席", "4", "养老集团负责人", "泰康之家、福寿康或其他已确认机构", "是否采购、试点和运营", "不与产业实践同一人。"),
    ("圆桌第五席", "5", "AIoT企业负责人", "涂鸦、移远、乐鑫或安防物联企业", "感知、网络和报警如何进入养老", "不与主旨四同一人。"),
    ("圆桌第六席", "6", "产业基金或银行科技金融负责人", "见人选池。须有项目投资或业务权限", "融资、租赁、采购支付和投资回报", "不接受纯品牌推广人员。保险机构优先进入会后支付对接，不增加第七席。"),
    ("成果秀", "1", "傅利叶智能", "创始人或康复业务团队，视主旨二是否已演示而定", "一个康复或照护动作，加机构部署方式", "优先安排，出席仍须确认。"),
    ("成果秀", "2", "智元机器人", "未在主旨二发言的业务或产品负责人", "一个养老或服务动作", "优先安排，出席仍须确认。"),
    ("成果秀", "3", "强脑科技", "韩璧丞或业务负责人", "脑机接口与智能康复，说明人群和医疗边界", "未确认则由一家AIoT企业递补，或直接减场。"),
    ("意向发布", "1", "场景机构", "尚未发言的养老机构或社区", "我要什么、我能提供什么、会后找谁", "会前选定，不现场征集。"),
    ("意向发布", "2", "康复医院", "上海相关康复医院拟邀负责人", "临床验证需求", "未确认则空出。"),
    ("意向发布", "3", "科技企业", "已参加成果秀或主旨、且需求对得上圆桌条件的企业", "一条试点请求", "不接受新的产品介绍。"),
    ("资本池", "—", "产业投资基金", "优势资本、启明创投医疗科技或硬科技团队、上海国投、张江科投、浦东创投、临港科创投、复星相关投资团队、机器人及具身智能专项基金、长三角医疗科技及智能制造产业基金", "圆桌第六席", "只邀请合伙人、投资总监或业务负责人。"),
    ("金融池", "—", "银行与保险", "工商银行上海市分行养老金融或科技金融、交通银行养老金融、浦发银行科技金融、杭州银行上海分行科技金融、泰隆银行上海相关分支、上海银行、太平养老、平安养老；封进可从长期护理保险作学术补充", "圆桌第六席或咨询席", "顾建忠若确认，优先于分支机构负责人。保险机构不另加椅子。"),
    ("在场但不单设发言", "—", "璞跃中国、国家技术转移东部中心", "拟邀产业创新合作单位", "成果转化与创新生态", "建议前排就座。东部中心可与第一席会后对接，不加圆桌席。"),
    ("主持", "—", "徐洁平或专业主持人", "源方案未列职务", "开场、串场、控时、圆桌追问", "不发表主题演讲。"),
]


CONTINGENCY = [
    ("主旨超时", "剩两分钟黄牌，到点红牌。主持上前致谢并引入下一位，不在句中切断，但不再允许展开新的段落。演示失控按安全规则立即停止。"),
    ("致辞只到一位", "13:13直接进入主旨一。省下的四分钟不平均分给后面的企业，主旨二的演示时限仍是七分钟。"),
    ("主讲人缺席", "按该场备选顺序现场替补。替补的人使用自己名下的题目。没有替补时缩短该场，主持用两分钟复述该场必须留下的问题，然后进入下一场。不把话筒交给无关企业。"),
    ("同一人重复登台", "主旨和圆桌原则上错开。成果秀允许同一企业更换业务负责人，但不重复同一段演示。"),
    ("圆桌缺席", "可以空席。第四席和第六席同时空缺时改为场景问诊，15:45仍然结束。不把空席改成临时路演。"),
    ("成果秀只有两家", "第三家时段取消，提前进入需求汇总。多出的时间不回加给圆桌，圆桌已在15:45结束。"),
    ("成果秀只有一家或没有", "保留已确认的一家。完全没有时，15:45由主持把茶歇中观众提出的共性问题归纳五分钟，然后休息至16:15，不临时点名观众席企业。"),
    ("全场已经超时", "压缩顺序固定：先取消意向发布中的咨询席说明，再把每家成果秀压到八分钟整，最后把总结压到题板原文。16:30必须下台合影。"),
    ("机器人不能动作", "改为九十秒视频加失败案例。茶歇展示同步取消。不播放与养老无关的宣传片。"),
    ("设备安全", "演示区和第一排之间留出通道。每台设备有急停说明和一名看护。主持可以中止演示。茶歇期间机器人不进入座位区。合影前先断电。"),
    ("场地变化", "地点现为上海展览中心中央大厅会议室C，以组委会最终通知为准。换厅后席卡、展示动线和合影人数规则不变。"),
    ("经费边界", "源方案写明场地免费，主办方不提供现金经费。执行使用现有会场、企业自带可安全运行的设备和主办方已有人员，不单列礼仪、演艺和复杂灯光。"),
]


def validate() -> None:
    assert SESSIONS[0]["start"] == "13:00"
    assert SESSIONS[-1]["end"] == "16:30"
    total = 0
    for index, session in enumerate(SESSIONS):
        span = duration(session["start"], session["end"])
        assert span > 0, session["name"]
        total += span
        if index:
            assert SESSIONS[index - 1]["end"] == session["start"], session["name"]
        cursor = session["start"]
        covered = 0
        for start, end, owner, action in session["beats"]:
            assert start == cursor, (session["name"], start, cursor)
            assert owner and action
            covered += duration(start, end)
            cursor = end
        assert cursor == session["end"], session["name"]
        assert covered == span, session["name"]
    assert total == 210, total
    pre_cursor = PRE_EVENT[0][0]
    for start, end, _owner, _action in PRE_EVENT:
        assert start == pre_cursor
        pre_cursor = end
    assert pre_cursor == "13:00"
    assert POST_EVENT[0][0] == "16:30"


def shade(cell, color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    node = OxmlElement("w:shd")
    node.set(qn("w:val"), "clear")
    node.set(qn("w:color"), "auto")
    node.set(qn("w:fill"), color)
    tc_pr.append(node)


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


def set_cell_margins(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = OxmlElement("w:tcMar")
    for edge, value in (("top", "50"), ("left", "70"), ("bottom", "50"), ("right", "70")):
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
    total = str(int(sum(widths) * 567))
    tbl_w.set(qn("w:w"), total)
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
        for index, cell in enumerate(row.cells):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(int(widths[index] * 567)))
            tc_w.set(qn("w:type"), "dxa")


def set_table_borders(table, color: str = LINE) -> None:
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "4")
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), color)
        borders.append(node)
    tbl_pr.append(borders)


def keep_row(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    node = OxmlElement("w:cantSplit")
    tr_pr.append(node)


def repeat_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    node = OxmlElement("w:tblHeader")
    node.set(qn("w:val"), "true")
    tr_pr.append(node)


def write_cell(cell, text: str, size: float = 9, bold: bool = False, color: tuple | None = None,
               fill: str | None = None, align: str = "left") -> None:
    cell.text = ""
    lines = str(text).split("\n")
    for index, line in enumerate(lines):
        paragraph = cell.paragraphs[0] if index == 0 else cell.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if align == "center" else WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.08
        set_run_font(paragraph.add_run(line), size, bold, color)
    if fill:
        shade(cell, fill)
    set_cell_margins(cell)


def add_table(doc, headers: list[str], rows: list[list[str]], widths: list[float],
              font_size: float = 9, header_fill: str = PINE) -> None:
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    set_table_borders(table)
    set_table_widths(table, widths)
    for index, header in enumerate(headers):
        write_cell(table.rows[0].cells[index], header, size=9, bold=True, color=(255, 255, 255),
                   fill=header_fill, align="center")
    repeat_header(table.rows[0])
    keep_row(table.rows[0])
    for row_index, row in enumerate(rows):
        fill = ALT if row_index % 2 == 0 else WHITE
        for col_index, value in enumerate(row):
            write_cell(table.rows[row_index + 1].cells[col_index], value, size=font_size, fill=fill,
                       align="center" if col_index == 0 else "left")
        keep_row(table.rows[row_index + 1])
    doc.add_paragraph()


def add_body(doc, text: str, size: float = 10.5, bold: bool = False, color: tuple | None = None,
             align: str = "left", space_after: float = 6, space_before: float = 0):
    paragraph = doc.add_paragraph()
    paragraph.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }[align]
    paragraph.paragraph_format.space_before = Pt(space_before)
    paragraph.paragraph_format.space_after = Pt(space_after)
    paragraph.paragraph_format.line_spacing = 1.15
    set_run_font(paragraph.add_run(text), size, bold, color)
    return paragraph


def add_bullets(doc, items: list[str]) -> None:
    for item in items:
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.left_indent = Cm(0.75)
        paragraph.paragraph_format.first_line_indent = Cm(-0.4)
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(2)
        paragraph.paragraph_format.line_spacing = 1.12
        set_run_font(paragraph.add_run("·  " + item), 10.5, color=INK_RGB)


def add_heading_cn(doc, text: str, level: int):
    heading = doc.add_heading(text, level=level)
    heading.paragraph_format.keep_with_next = True
    heading.paragraph_format.space_before = Pt(14 if level == 1 else 10)
    heading.paragraph_format.space_after = Pt(4)
    color = PINE_RGB if level == 1 else INK_RGB
    size = 16 if level == 1 else 13
    for run in heading.runs:
        set_run_font(run, size, True, color)
    return heading


def set_styles(doc) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor(*INK_RGB)
    r_pr = normal.element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        r_fonts.set(qn(attr), FONT)
    lang = r_pr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        r_pr.append(lang)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:eastAsia"), "zh-CN")


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


def configure_section(section) -> None:
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.7)
    section.bottom_margin = Cm(1.6)
    section.left_margin = Cm(1.7)
    section.right_margin = Cm(1.7)
    section.header_distance = Cm(0.6)
    section.footer_distance = Cm(0.5)
    section.different_first_page_header_footer = True

    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.LEFT
    header.paragraph_format.space_after = Pt(2)
    set_run_font(header.add_run("银发科技与具身智能产业创新峰会  ·  详细议程（拟邀稿）"), 8, color=PINE_RGB)
    p_pr = header._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), PINE)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(footer.add_run("2026年11月28日  13:00—16:30  ·  上海展览中心  ·  第 "), 8, color=(0x66, 0x66, 0x66))
    add_page_field(footer)
    set_run_font(footer.add_run(" 页"), 8, color=(0x66, 0x66, 0x66))


def build_docx(path: Path) -> None:
    doc = Document()
    set_styles(doc)
    configure_section(doc.sections[0])
    core = doc.core_properties
    core.title = "2026上海国际智慧康养产业博览会分论坛详细议程"
    core.subject = "AI重塑养老未来——2026银发科技与具身智能产业创新峰会"
    core.category = "议程"
    core.comments = "依据拟邀方案整理。嘉宾均为拟邀，斜线为替代人选。"

    add_body(doc, "2026 上海国际智慧康养产业博览会", size=12, color=GOLD_RGB, align="center", space_after=4)
    add_body(doc, "分论坛详细议程", size=26, bold=True, color=PINE_RGB, align="center", space_after=2)
    add_body(doc, "AI重塑养老未来", size=18, bold=True, color=INK_RGB, align="center", space_after=2)
    add_body(doc, "2026银发科技与具身智能产业创新峰会", size=14, color=INK_RGB, align="center", space_after=8)
    add_body(doc, "2026年11月28日（周六）  13:00—16:30", size=12, bold=True, align="center", space_after=2)
    add_body(doc, "上海展览中心中央大厅会议室C", size=11, align="center", space_after=2)
    add_body(doc, "地点以组委会最终确认为准", size=9, color=(0x88, 0x88, 0x88), align="center", space_after=10)
    add_body(
        doc,
        "副标题：“十五五”银发经济 · 人形机器人 · AI健康管理 · 智慧物联网 · 产业资本",
        size=10.5,
        align="center",
        space_after=12,
    )
    add_table(
        doc,
        ["项目", "内容"],
        [
            ["大会主办", "上海现代服务业联合会；上海市养老服务和老龄产业协会；上海丰之米会展有限公司；\n中国工商银行股份有限公司上海市分行"],
            ["分论坛主办", "上海星智耀科技有限公司"],
            ["联合承办", "上海市杨浦区科技企业联合会；复旦大学住房政策研究中心"],
            ["拟邀学术支持", "复旦大学老龄研究院"],
            ["规模与形式", "100—120人。主旨演讲、科技展示、产业圆桌、创新成果秀、供需对接"],
            ["文件性质", "执行议程拟邀稿 V1.0。正式发言嘉宾以确认函为准"],
        ],
        [3.4, 14.2],
        font_size=9,
    )
    add_body(doc, "场地免费，主办方不提供现金经费。议程按现有会场条件编排。", size=9, color=GOLD_RGB, align="center")

    add_heading_cn(doc, "一、这场议程怎么用", 1)
    add_body(
        doc,
        "源方案已经确定这是博览会中的一场分论坛，正式时间就是11月28日13:00至16:30，共210分钟。本文不另编全馆展期，只把这一场写成可以照着执行的议程。观众手册使用第二章的付印表；内部催确认和现场执行分别使用第二章底稿与第四章分钟脚本。",
    )
    add_table(
        doc,
        ["规则", "执行"],
        [
            ["一名主讲", "每个主旨、致辞席和圆桌席只保留一名最终确认的主讲人。"],
            ["斜线是替代", "源方案单元格中的斜线表示替代人选，不是要求这些人一起上台。"],
            ["拟邀不是确认", "姓名进入本文，只说明在拟邀范围内。对外手册在收到确认前不印具体姓名。"],
            ["错开登台", "同一人原则上不既讲主旨又坐圆桌。成果秀可以换同企业的业务负责人，但不重复同一段演示。"],
            ["投资席的资格", "优先邀请有项目投资权限的合伙人、投资总监或业务负责人。不安排纯品牌推广人员代替产业投资嘉宾。"],
            ["时间红线", "16:30必须进入合影。压缩时先动咨询席说明和成果秀超时，不动圆桌的结束时间。"],
        ],
        [3.4, 14.2],
    )

    add_heading_cn(doc, "二、对外议程", 1)
    add_body(doc, "手册付印使用下表。姓名在确认出席前不印，只印席位方向。带替补顺序的底稿放在第二章后面，供内部催确认，不放进观众资料。")
    add_table(
        doc,
        ["时间", "环节及题目", "席位方向"],
        [
            ["13:00—13:08", "开幕：AI重塑养老未来", "主持人"],
            ["13:08—13:18", "开幕致辞：共建长三角银发科技产业生态", "行业协会、大会主办方"],
            ["13:18—13:38", "主旨一：“十五五”银发经济与AI时代的长寿社会", "老龄研究学者"],
            ["13:38—14:00", "主旨二：具身智能走进养老", "康复机器人或具身智能企业"],
            ["14:00—14:18", "主旨三：大模型与智能体", "人工智能学者或企业"],
            ["14:18—14:36", "主旨四：AIoT与智慧家庭", "物联网企业"],
            ["14:36—14:50", "产业实践：谁来为养老机器人买单", "养老机构负责人"],
            ["14:50—15:00", "茶歇与机器人展示交流", "参会嘉宾"],
            ["15:00—15:45", "高端圆桌：技术成熟了，商业模式成熟了吗", "场景、技术、学术、机构、物联、资本"],
            ["15:45—16:15", "创新成果秀：未来养老的AI新产品与应用场景", "已确认的2—3家"],
            ["16:15—16:25", "供需对接及合作意向发布", "有明确场景或验证需求的单位"],
            ["16:25—16:30", "论坛总结及合影", "主承办方"],
        ],
        [3.2, 8.8, 5.6],
        font_size=9,
    )
    add_body(doc, "内部确认底稿。付印前不放入观众手册。", size=10.5, bold=True, color=INK_RGB, space_after=2)
    public_rows = []
    for session in SESSIONS:
        public_rows.append([
            f"{session['start']}—{session['end']}",
            f"{session['name']}：{session['title']}",
            session["primary"],
            session["backup"],
        ])
    add_table(
        doc,
        ["时间", "环节及题目", "拟邀主讲", "备选（付印时删除）"],
        public_rows,
        [2.8, 4.6, 5.2, 5.0],
        font_size=8,
    )

    add_heading_cn(doc, "三、时间预算", 1)
    add_body(
        doc,
        "上半场用一百一十分钟把问题讲清：老龄社会需要什么，机器人、智能体和物联网各自能做什么，以及机构是否买单。茶歇十分钟只看设备。下半场用九十分钟做判断：圆桌决定商业模式是否成熟，成果秀留下证据和需求，意向发布只接收对得上的合作。",
    )
    budget_rows = []
    for session in SESSIONS:
        span = duration(session["start"], session["end"])
        budget_rows.append([
            f"{session['start']}—{session['end']}",
            str(span),
            session["name"],
            session["goal"],
        ])
    budget_rows.append(["合计", "210", "正式议程", "13:00开始，16:30结束"])
    add_table(doc, ["时间", "分钟", "环节", "这一段只抓一件事"], budget_rows, [3.2, 1.6, 3.0, 9.8], font_size=8)

    add_heading_cn(doc, "四、会前准备与逐分钟脚本", 1)
    add_heading_cn(doc, "4.1  会前，不计入210分钟", 2)
    add_body(doc, "12:20至13:00是会务时间。迟到嘉宾不改变13:00的开场。")
    add_table(
        doc,
        ["时间", "负责人", "动作"],
        [[a, b, c] for a, _end, b, c in [(f"{s}—{e}", e, o, act) for s, e, o, act in PRE_EVENT]],
        [3.0, 2.4, 12.2],
        font_size=9,
    )
    add_body(
        doc,
        "会务建议：观众席按100—120人准备，工作人员另计，总数以场地核载为准。发言使用讲台，圆桌六席在14:58入座。前排右侧预留主旨和圆桌座位，茶歇后直接上台。展示放在厅侧，留出疏散通道。岗位为主持、计时、签到登记两人、设备看护、摄影、会务总控，约七人，由大会主办和分论坛承办分摊。",
        size=10.5,
    )

    for session in SESSIONS:
        span = duration(session["start"], session["end"])
        add_heading_cn(doc, f"{session['start']}—{session['end']}  {session['name']}（{span}分钟）", 2)
        add_body(doc, session["title"], size=12, bold=True, color=PINE_RGB, space_after=2)
        add_body(doc, session["goal"])
        add_table(
            doc,
            ["项目", "安排"],
            [
                ["优先出场", session["primary"]],
                ["备选与限制", session["backup"]],
            ],
            [3.2, 14.4],
            font_size=9,
        )
        add_body(doc, "内容要求", size=10.5, bold=True, color=INK_RGB, space_after=2)
        add_bullets(doc, session["outline"])
        add_body(doc, "分钟脚本", size=10.5, bold=True, color=INK_RGB, space_before=6, space_after=2)
        add_table(
            doc,
            ["时间", "负责人", "动作"],
            [[f"{start}—{end}", owner, action] for start, end, owner, action in session["beats"]],
            [3.0, 2.4, 12.2],
            font_size=8.5,
        )

    add_heading_cn(doc, "4.2  会后十五分钟，不计入正式议程", 2)
    add_table(
        doc,
        ["时间", "负责人", "动作"],
        [[f"{start}—{end}", owner, action] for start, end, owner, action in POST_EVENT],
        [3.0, 2.4, 12.2],
    )

    add_heading_cn(doc, "五、控时、缺席与现场安全", 1)
    add_body(doc, "计时牌只给主持和发言人看。黄牌表示剩两分钟，红牌表示到点。红牌出现后，主持接过下一句，话筒不必突然关闭。")
    add_table(
        doc,
        ["情形", "处理"],
        CONTINGENCY,
        [3.6, 14.0],
        font_size=9,
    )

    add_heading_cn(doc, "六、人选确认表", 1)
    add_body(doc, "下表把源方案中的拟邀名单收束到各个席位。状态统一为待确认。收到回执后只改这一张表，再回头修改第二章的对外议程。")
    add_table(
        doc,
        ["席位", "顺位", "姓名", "单位及身份", "题目或任务", "使用规则"],
        [list(row) for row in GUESTS],
        [2.3, 1.2, 2.4, 3.8, 3.9, 4.0],
        font_size=7.5,
    )
    add_body(
        doc,
        "产业创新合作单位中，傅利叶智能、智元机器人、宇树科技、强脑科技、商汤科技已进入发言或成果秀池。璞跃中国、国家技术转移东部中心建议到场，不单设发言。联合承办的上海市杨浦区科技企业联合会、复旦大学住房政策研究中心出席总结和前排，不另开致辞。",
        size=10.5,
    )

    add_heading_cn(doc, "七、离场时应留下的记录", 1)
    add_bullets(doc, [
        "题板三句：机构愿意付费的类别、试点条件、明确不买的类别。",
        "圆桌收束三句：可以进场的任务、试点条件、会后只促成的合作类型。",
        "成果秀需求：每家一条，最多三条。",
        "对接卡：单位、需求类型（试点、采购、临床、融资、渠道）、对接人、是否允许会后公开。",
        "合影：按16:28的人数规则完成，不在散会后补拍企业单独广告画面。",
    ])
    add_body(
        doc,
        "本稿依据组委会拟邀方案编制。正式嘉宾、第二位致辞人和会议室编号，以书面确认为准。",
        size=9,
        color=(0x66, 0x66, 0x66),
        space_before=6,
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


def fill_header(ws, headers: list[str], fill: str) -> None:
    thin = Border(
        left=Side(style="thin", color="E4EAE8"),
        right=Side(style="thin", color="E4EAE8"),
        top=Side(style="thin", color="E4EAE8"),
        bottom=Side(style="thin", color="E4EAE8"),
    )
    for col, header in enumerate(headers, 1):
        cell = ws.cell(1, col, header)
        cell.font = Font(name=FONT, bold=True, color=WHITE, size=10)
        cell.fill = PatternFill("solid", fgColor=fill)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 22
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.6, bottom=0.5, header=0.2, footer=0.2)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddHeader.left.text = "2026银发科技与具身智能产业创新峰会"
    ws.oddFooter.center.text = "拟邀稿  ·  第 &P 页"
    ws.oddFooter.center.font = FONT
    ws.print_title_rows = "1:1"


def write_rows(ws, rows: list[list], fills: list[str] | None = None, widths: list[float] | None = None) -> None:
    thin = Border(
        left=Side(style="thin", color="E4EAE8"),
        right=Side(style="thin", color="E4EAE8"),
        top=Side(style="thin", color="E4EAE8"),
        bottom=Side(style="thin", color="E4EAE8"),
    )
    for row_index, row in enumerate(rows, 2):
        fill = (fills[row_index - 2] if fills else (ALT if row_index % 2 == 0 else WHITE))
        lines = 1
        for col_index, value in enumerate(row, 1):
            cell = ws.cell(row_index, col_index, value)
            cell.font = Font(name=FONT, size=9, color=INK)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.border = thin
            width = widths[col_index - 1] if widths else 20
            # 中文在默认列宽下大约每 1.7 个宽度单位容纳一个字。
            capacity = max(8, int(width / 1.8))
            lines = max(lines, (len(str(value)) + capacity - 1) // capacity)
        ws.row_dimensions[row_index].height = min(96, max(22, lines * 15 + 8))


def set_widths(ws, widths: list[float]) -> None:
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(index)].width = width


KIND_FILL = {
    "开幕": "F8F1E3",
    "致辞": "F8F1E3",
    "主旨": "E7F2F1",
    "实践": "FBF6EE",
    "茶歇": "F3F4F4",
    "圆桌": "DCECEA",
    "成果秀": "F8F1E3",
    "对接": "FBF6EE",
    "总结": "E7EEF2",
}


def build_xlsx(path: Path) -> None:
    wb = Workbook()

    ws = wb.active
    ws.title = "对外议程"
    headers = ["时间", "分钟", "环节", "题目", "拟邀主讲", "备选（付印前删除）"]
    fill_header(ws, headers, PINE)
    rows = []
    fills = []
    for session in SESSIONS:
        rows.append([
            f"{session['start']}—{session['end']}",
            duration(session["start"], session["end"]),
            session["name"],
            session["title"],
            session["primary"],
            session["backup"],
        ])
        fills.append(KIND_FILL[session["kind"]])
    public_widths = [16, 8, 14, 42, 46, 55]
    write_rows(ws, rows, fills, public_widths)
    set_widths(ws, public_widths)
    for row in ws.iter_rows(min_row=2, max_col=2):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    cue = wb.create_sheet("分钟执行")
    fill_header(cue, ["开始", "结束", "分钟", "环节", "负责人", "动作"], "1B3A4B")
    cue_rows = []
    cue_fills = []
    for start, end, owner, action in PRE_EVENT:
        cue_rows.append([start, end, duration(start, end), "会前（不计正式议程）", owner, action])
        cue_fills.append("F3F4F4")
    for session in SESSIONS:
        for start, end, owner, action in session["beats"]:
            cue_rows.append([start, end, duration(start, end), f"{session['name']}｜{session['title']}", owner, action])
            cue_fills.append(KIND_FILL[session["kind"]])
    for start, end, owner, action in POST_EVENT:
        cue_rows.append([start, end, duration(start, end), "会后（不计正式议程）", owner, action])
        cue_fills.append("F3F4F4")
    cue_widths = [10, 10, 8, 38, 14, 88]
    write_rows(cue, cue_rows, cue_fills, cue_widths)
    set_widths(cue, cue_widths)
    for row in cue.iter_rows(min_row=2, max_col=3):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    panel = wb.create_sheet("圆桌手卡")
    fill_header(panel, ["时间", "对象", "必须完成的判断"], "0F5C58")
    panel_session = next(item for item in SESSIONS if item["kind"] == "圆桌")
    panel_rows = [[start + "—" + end, owner, action] for start, end, owner, action in panel_session["beats"]]
    panel_widths = [16, 14, 100]
    write_rows(panel, panel_rows, ["DCECEA"] * len(panel_rows), panel_widths)
    set_widths(panel, panel_widths)

    people = wb.create_sheet("人选确认")
    fill_header(people, ["席位", "顺位", "姓名", "单位及身份", "题目或任务", "使用规则", "确认状态"], GOLD)
    people_rows = [list(row) + ["待确认"] for row in GUESTS]
    people_widths = [16, 8, 22, 46, 42, 46, 12]
    write_rows(people, people_rows, widths=people_widths)
    set_widths(people, people_widths)

    risk = wb.create_sheet("控时与预案")
    fill_header(risk, ["情形", "处理"], RED)
    risk_widths = [24, 110]
    write_rows(risk, [list(item) for item in CONTINGENCY], ["FBF6EE", WHITE] * 6, risk_widths)
    set_widths(risk, risk_widths)

    note = wb.create_sheet("读表")
    fill_header(note, ["事项", "说明"], PINE)
    notes = [
        ["文件来源", "2026上海国际智慧康养产业博览会拟邀方案中的分论坛。本稿只展开11月28日13:00—16:30这一场。"],
        ["斜线", "替代人选，不是联合上台。每个发言席只留一名最终主讲人。"],
        ["状态", "人选确认表中的“待确认”需要改成确认出席、无法出席或改期。未确认的姓名不进入对外付印稿。"],
        ["投资嘉宾", "只邀请有实际项目权限的合伙人、投资总监或业务负责人，不接受纯品牌推广人员代为出席圆桌。"],
        ["经费", "场地免费，主办方不提供现金经费。"],
        ["正式时长", "210分钟。会前12:20—13:00和会后16:30—16:45不计入。"],
    ]
    note_widths = [18, 100]
    write_rows(note, notes, widths=note_widths)
    set_widths(note, note_widths)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def export_pdf(docx_path: Path) -> None:
    soffice = shutil.which("soffice")
    if not soffice:
        print("未检测到 LibreOffice，已跳过 PDF。")
        return
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(docx_path.parent), str(docx_path)],
        check=True,
    )


def main() -> None:
    validate()
    docx_path = OUT / "2026上海国际智慧康养产业博览会_分论坛详细议程.docx"
    xlsx_path = OUT / "2026上海国际智慧康养产业博览会_分论坛执行议程表.xlsx"
    build_docx(docx_path)
    build_xlsx(xlsx_path)
    export_pdf(docx_path)
    print(f"正式议程 {sum(duration(s['start'], s['end']) for s in SESSIONS)} 分钟，{len(SESSIONS)} 个环节")
    print(docx_path)
    print(xlsx_path)


if __name__ == "__main__":
    main()
