# -*- coding: utf-8 -*-
"""创智汇 · 30场活动专项交付（修订版）
仅针对活动部分单独交付；不做招商制度保障。
落实修改：
1) 压低人数目标
2) 出海场次慎重（不强制外企/领事到场）
3) 复盘与客户跟踪由同浦汇主责，园区重心放后续活动
4) 回访纪要交园区为知会摘要，弱化园区责任
5) 租金基准配套建议免租期
6) 付款节点给建议方案供线下讨论
协同着眼点：活动结束建群后，园区可在群内逼单促成交。
"""
import os, shutil
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
ART = "/opt/cursor/artifacts"
PHOTO = os.path.join(HERE, "cases", "photos")
FONT = "Microsoft YaHei"

INK = RGBColor(0x1A, 0x1A, 0x2E)
MUT = RGBColor(0x4A, 0x4A, 0x5A)
SOFT = RGBColor(0x8A, 0x8A, 0x9A)
BLUE = RGBColor(0x1E, 0x4D, 0x8C)
GOLD = RGBColor(0xC4, 0x9A, 0x3C)
LINE = RGBColor(0xD8, 0xDE, 0xE8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG = RGBColor(0xF7, 0xF8, 0xFA)
ACC = RGBColor(0xE8, 0xF0, 0xFA)
GREEN = RGBColor(0x1F, 0x7A, 0x5C)
ROSE = RGBColor(0xA6, 0x3D, 0x40)

SW, SH = Inches(13.333), Inches(7.5)
ML, CW = Inches(0.55), Inches(12.23)

# ---------- 30场（压低人数；出海慎重） ----------
EVENTS = [
    dict(code="A1", cat="A", name="智能伙伴·创智汇年度启动沙龙", month="2026年8月上旬", fmt="主题沙龙",
         people="40人", dm="80%", assoc="100%", floor="5F沙龙", rhythm="T-14嘉宾物料→T-7邀约→T-1彩排→D日→T+7回访建群",
         lease="集中邀约潜在客户，为全年定调",
         agenda="签到参观→主题解读→园区定位→圆桌→看场意向",
         guests="联合会、AI/OPC企业负责人、服中心", value="开年定调；待租单元曝光",
         deliver="方案+议程+关联度核验表+签到+意向表+建群+回访摘要+复盘"),
    dict(code="E1", cat="E", name="WAIC成果承接·创智汇开放日", month="2026年8月中旬", fmt="开放日",
         people="120–150人", dm="35%", assoc="100%", floor="3F+5F", rhythm="同上（大场提前T-21启动）",
         lease="大会资源导入园区看场",
         agenda="致辞→主题速递→开放参观→双展台→政策礼包→意向洽谈",
         guests="WAIC回流嘉宾、媒体、AI/内容企业", value="流量导入；集中看场",
         deliver="同上+动线图+通稿"),
    dict(code="E2", cat="E", name="AI Builders Night·建造者之夜", month="2026年8月中旬", fmt="夜场社交",
         people="60–80人", dm="50%", assoc="100%", floor="5F主厅", rhythm="同上",
         lease="建造者当场看3F单元",
         agenda="签到→致辞→闪电演讲→开放麦→AMA→看场",
         guests="创业者、开发者、投资人", value="Builders看单元",
         deliver="同上+对接表"),
    dict(code="E3", cat="E", name="通往AGI之夜", month="2026年8月下旬", fmt="学术夜场",
         people="40–50人", dm="60%", assoc="100%", floor="5F圆桌", rhythm="同上",
         lease="招引研究型/模型团队",
         agenda="致辞→圆桌→围炉→投票→精品看场",
         guests="学者、模型负责人、Agent创始人", value="研究型客户",
         deliver="同上"),
    dict(code="C5", cat="C", name="东莞潮玩品牌入沪推介", month="2026年8月下旬", fmt="推介会",
         people="40人", dm="80%", assoc="100%", floor="5F潮玩区", rhythm="同上",
         lease="品牌展位与快闪",
         agenda="渠道介绍→落位→联名→看场→意向",
         guests="潮玩品牌、渠道商", value="展位/快闪去化",
         deliver="同上"),
    dict(code="A2", cat="A", name="Agent智能体搭建一日营", month="2026年9月上旬", fmt="实训营",
         people="30–40人", dm="60%", assoc="100%", floor="3F OPC", rhythm="同上",
         lease="Agent初创落位",
         agenda="签到→工作流→实操→Demo→看场",
         guests="Agent工程师、开源Maintainer", value="入孵转化",
         deliver="同上"),
    dict(code="C6", cat="C", name="IP授权×AI衍生交易撮合会", month="2026年9月中旬", fmt="撮合会",
         people="40–50人", dm="80%", assoc="100%", floor="5F展示中心", rhythm="同上",
         lease="IP/衍生品团队入驻",
         agenda="IP路演→专题桌→一对一→成交看板→落户",
         guests="IP方、被授权商、渠道", value="贸易+入驻",
         deliver="同上"),
    # 出海慎重：不办国际买家到场，改为国内出海企业路径沙龙
    dict(code="F1", cat="F", name="AI出海路径与合规沙龙（国内企业专场）", month="2026年9月下旬", fmt="政策/路径沙龙",
         people="30–40人", dm="85%", assoc="100%", floor="3F/5F", rhythm="同上",
         lease="服务有出海意向的国内企业",
         agenda="出海路径→合规要点→案例→一对一咨询→看场",
         guests="出海顾问、已出海国内企业、合规顾问（不邀请境外企业到场）",
         value="低成本验证出海需求；外企/买家团另案评估",
         deliver="同上；注明：不含境外企业接待成本"),
    dict(code="A3", cat="A", name="多模态智能体工作坊", month="2026年10月上旬", fmt="工作坊",
         people="30–40人", dm="60%", assoc="100%", floor="3F培训", rhythm="同上",
         lease="多模态团队入驻",
         agenda="案例→实操→共创→评审→看场",
         guests="多模态研究者、应用厂商", value="技术团队入驻",
         deliver="同上"),
    dict(code="E4", cat="E", name="创智汇秋季潮玩×AI主题日", month="2026年10月中旬", fmt="主题日（1天）",
         people="150–200人", dm="30%", assoc="100%", floor="5F+公区", rhythm="大场T-21",
         lease="集中签约展位/办公",
         agenda="开幕→市集展洽→互动→洽谈签约",
         guests="潮玩品牌、达人、AI创作者", value="集中签约（规模可控）",
         deliver="同上"),
    dict(code="B2", cat="B", name="AI治理与可信智能体沙龙", month="2026年10月下旬", fmt="沙龙",
         people="30–40人", dm="85%", assoc="100%", floor="3F", rhythm="同上",
         lease="合规型企业信任",
         agenda="治理框架→合规清单→案例→问答→对接",
         guests="治理学者、合规顾问、企业法务", value="信任入驻",
         deliver="同上"),
    dict(code="A4", cat="A", name="火山引擎×算力Infra实务营", month="2026年11月上旬", fmt="厂商联训",
         people="30–40人", dm="70%", assoc="100%", floor="3F+云创", rhythm="同上",
         lease="高算力企业定向",
         agenda="政策包→代金券→案例→诊断→捆绑入驻",
         guests="火山引擎、云创基地、企业CTO", value="高算力招商",
         deliver="同上"),
    dict(code="B4", cat="B", name="创新券·算力券·模型券实务沙龙", month="2026年11月中旬", fmt="实务沙龙",
         people="30–40人", dm="85%", assoc="100%", floor="3F", rhythm="同上",
         lease="用券企业聚集",
         agenda="三券规则→核销→案例→开户→入驻",
         guests="券平台、云厂商、企业负责人", value="券务转化",
         deliver="同上"),
    # 原领事专题改为慎重可选/国内涉外服务专场
    dict(code="F2", cat="F", name="涉外服务与出海准备专场（慎重场）", month="2026年11月下旬", fmt="闭门沙龙",
         people="20–30人", dm="90%", assoc="100%", floor="会客厅/5F", rhythm="同上；外事级另案",
         lease="服务国内涉外/出海准备企业",
         agenda="涉外服务路径→合规→案例→一对一→看场",
         guests="涉外服务机构、国内出海企业（默认不安排领事/外企到场；若需另案评估费用）",
         value="控制成本；领事/外企场次单独立项",
         deliver="同上；费用边界书面确认"),
    dict(code="A5", cat="A", name="具身智能空间交互体验日", month="2026年12月上旬", fmt="体验日",
         people="40–50人", dm="50%", assoc="100%", floor="5F+3F", rhythm="同上",
         lease="机器人团队看场",
         agenda="演示→讲解→场景清单→踩点→洽谈",
         guests="具身团队、高校实验室", value="具身看场",
         deliver="同上"),
    dict(code="F4", cat="F", name="创智汇AI年度Demo Day", month="2026年12月中旬", fmt="路演日",
         people="80–120人", dm="65%", assoc="100%", floor="5F主展", rhythm="大场T-21",
         lease="集中签约与媒体背书",
         agenda="开幕→十强路演→颁奖→洽谈签约",
         guests="投资人、链主、媒体、十强项目", value="集中签约",
         deliver="同上+签约台账"),
    dict(code="D1", cat="D", name="AIGC微短剧制片特训", month="2027年1月上旬", fmt="特训营",
         people="30–40人", dm="60%", assoc="100%", floor="3F+5F", rhythm="同上",
         lease="厂牌/工作室入驻",
         agenda="政策→制片→脚本→路演→看场",
         guests="导演、厂牌、平台方", value="内容入驻",
         deliver="同上"),
    dict(code="A6", cat="A", name="AI营销Agent实战营", month="2027年1月中旬", fmt="实战营",
         people="30–40人", dm="70%", assoc="100%", floor="3F", rhythm="同上",
         lease="营销科技公司",
         agenda="策略→Agent→素材→复盘→转化",
         guests="投放操盘手、产品经理", value="营销科技入驻",
         deliver="同上"),
    dict(code="B1", cat="B", name="YOUNG立方×智能伙伴政策沙龙", month="2027年2月中旬", fmt="政策沙龙",
         people="40人", dm="85%", assoc="100%", floor="5F沙龙", rhythm="同上",
         lease="内容/AI企业导入",
         agenda="政策→礼包→画像→诊断→看场",
         guests="政策宣讲、服中心、企业负责人", value="政策获客",
         deliver="同上"),
    dict(code="B3", cat="B", name="高企认定冲刺（AI企业专场）", month="2027年2月下旬", fmt="辅导会",
         people="30–40人", dm="90%", assoc="100%", floor="3F", rhythm="同上",
         lease="外区待认定企业带政策入驻",
         agenda="条件→材料→初筛→套餐→激励",
         guests="辅导顾问、待认定企业负责人", value="带政策入驻",
         deliver="同上"),
    dict(code="A7", cat="A", name="OPC超级个体黑客松（春）", month="2027年3月上旬", fmt="黑客松（1.5天）",
         people="40–50人", dm="45%", assoc="100%", floor="3F", rhythm="同上",
         lease="获奖团队优先谈单元",
         agenda="开题→开发→路演→礼包→看场",
         guests="评委、投资人、Builders", value="获奖谈单元",
         deliver="同上"),
    dict(code="C2", cat="C", name="高校成果转化·AI for Science日", month="2027年3月下旬", fmt="对接日",
         people="40–50人", dm="80%", assoc="100%", floor="3F+沙龙", rhythm="同上",
         lease="成果公司/实验室落户",
         agenda="成果路演→场景→承接→看单元→洽谈",
         guests="复旦/同济技转、教授团队", value="成果落户",
         deliver="同上"),
    dict(code="F3", cat="F", name="通往AGI季度圆桌", month="2027年3月下旬", fmt="闭门圆桌",
         people="20–30人", dm="90%", assoc="100%", floor="5F沙龙", rhythm="同上",
         lease="研究型/模型团队",
         agenda="议题→圆桌→围炉→问答→看场",
         guests="学者、模型企业、投资人", value="研究型客户",
         deliver="同上"),
    dict(code="D2", cat="D", name="AI创作者内容首发①", month="2027年4月中旬", fmt="发布会",
         people="60–80人", dm="40%", assoc="100%", floor="5F", rhythm="同上",
         lease="关联品牌问询",
         agenda="揭幕→发布→快闪→专访→招商通道",
         guests="创作者、媒体、渠道", value="品牌问询",
         deliver="同上"),
    dict(code="C3", cat="C", name="汕头玩具×AIGC供应链对接会", month="2027年4月下旬", fmt="对接会",
         people="40人", dm="80%", assoc="100%", floor="5F玩具区", rhythm="同上",
         lease="玩具展位落位",
         agenda="集群→AIGC提效→对接→参观→报价",
         guests="商协会、采购、玩具企业", value="展位去化",
         deliver="同上"),
    dict(code="C1", cat="C", name="专精特新·AI应用培育路演", month="2027年5月中旬", fmt="路演",
         people="40–50人", dm="75%", assoc="100%", floor="5F", rhythm="同上",
         lease="成长型AI补位",
         agenda="政策→路演→点评→金融→看场",
         guests="顾问、银行、基金、企业", value="成长型企业",
         deliver="同上"),
    dict(code="D3", cat="D", name="潮玩×AIGC主题市集", month="2027年5月下旬", fmt="市集（1–2天）",
         people="150–200人", dm="25%", assoc="100%", floor="5F+公区", rhythm="大场",
         lease="摊主升级固定展位",
         agenda="布摊→开市→打卡→转正谈",
         guests="摊主、达人、潮玩品牌", value="摊主转正",
         deliver="同上"),
    dict(code="C4", cat="C", name="扬州毛绒×数字人联名对接", month="2027年6月中旬", fmt="对接会",
         people="30–40人", dm="85%", assoc="100%", floor="5F毛绒区", rhythm="同上",
         lease="700㎡量级客户线索",
         agenda="联名→测算→看场→MOU→跟进",
         guests="扬州集群、数字人厂商", value="大客户线索",
         deliver="同上"),
    dict(code="D4", cat="D", name="沉浸式AI+IP联展", month="2027年6月下旬", fmt="联展（5–7天）",
         people="日均80–120人 / 累计约600人次", dm="20%", assoc="100%", floor="5F", rhythm="展期日报",
         lease="闭幕酒会集中转化",
         agenda="布展→开幕→观展→教育场→闭幕酒会",
         guests="联合IP、渠道、品牌", value="闭幕转化",
         deliver="同上+展期日报"),
    dict(code="D5", cat="D", name="AI创作者内容首发②（WAIC2027预热）", month="2027年7月中旬", fmt="发布会",
         people="60–80人", dm="40%", assoc="100%", floor="5F", rhythm="同上",
         lease="旺季补位，衔接大会",
         agenda="发布→渠道→看场→年框→预热",
         guests="IP方、渠道、媒体", value="衔接WAIC2027",
         deliver="同上"),
]

assert len(EVENTS) == 30

CASES = [
    {
        "name": "杨「数」浦数字沙龙：AI如何重塑企业DNA",
        "meta": "2025-06-30 · 美团上海综合指挥中心 · 沙龙类",
        "orgs": "杨浦区委网信办；赛博院；人工智能与社会发展研究会",
        "companies": "美团；大众点评大模型团队；赛博院研究员等",
        "agenda": ["参观数字化场景", "主题分享：AI重塑企业DNA", "案例：大众点评AI落地", "政策解读", "互动答疑"],
        "photos": ["case_c987bd91.jpg", "case_59157ea5.jpg", "case_6e51af9e.jpg"],
        "map": "对应：B类政策治理沙龙、A类训练营",
    },
    {
        "name": "「融见科创」人工智能专场路演暨投融资对接会",
        "meta": "2025-10 · 杨浦 · 路演对接类",
        "orgs": "杨浦科创促进会、邮储银行；科辰创投、益华资本等",
        "companies": "复楚智能、卡房信息、一造科技、中科趋势、万笔千墨等",
        "agenda": ["开场致辞", "主办方致辞", "AI趋势分享", "6+6+6+1路演对接", "投促点评与跟进"],
        "photos": ["case_db1d5cac.jpg", "case_65931516.jpg", "case_cb0a3812.jpg"],
        "map": "对应：F4 Demo Day、C1路演、E1开放日",
    },
]


def rect(s, x, y, w, h, fill=None, line=None, lw=0.8):
    b = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    b.shadow.inherit = False
    if fill is None: b.fill.background()
    else: b.fill.solid(); b.fill.fore_color.rgb = fill
    if line is None: b.line.fill.background()
    else: b.line.color.rgb = line; b.line.width = Pt(lw)
    return b


def text(s, x, y, w, h, content, size=14, color=INK, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = Pt(2); tf.margin_right = Pt(2); tf.margin_top = Pt(1); tf.margin_bottom = Pt(1)
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = content; r.font.size = Pt(size); r.font.bold = bold
    r.font.color.rgb = color; r.font.name = FONT
    return tb


def bullets(s, x, y, w, h, items, size=12, color=MUT, mark=BLUE):
    tb = s.shapes.add_textbox(x, y, w, h); tf = tb.text_frame; tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(4); p.line_spacing = 1.15
        r0 = p.add_run(); r0.text = "▪ "; r0.font.size = Pt(size); r0.font.color.rgb = mark; r0.font.name = FONT
        r = p.add_run(); r.text = it; r.font.size = Pt(size); r.font.color.rgb = color; r.font.name = FONT
    return tb


def slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, SW, SH, fill=WHITE)
    rect(s, 0, 0, SW, Inches(0.06), fill=BLUE)
    return s


def header(s, sec, title, sub=None):
    text(s, ML, Inches(0.22), Inches(12), Inches(0.28), sec, size=12, color=BLUE, bold=True)
    text(s, ML, Inches(0.48), Inches(12), Inches(0.42), title, size=22, color=INK, bold=True)
    if sub:
        text(s, ML, Inches(0.92), Inches(12), Inches(0.28), sub, size=12, color=MUT)


def footer(s, page, total=22):
    text(s, ML, Inches(7.15), Inches(9), Inches(0.25),
         "上海创智汇 × 同浦汇 · 30场活动专项交付（修订版）· 仅活动部分", size=10, color=SOFT)
    text(s, Inches(11.3), Inches(7.15), Inches(1.5), Inches(0.25), f"{page}/{total}", size=10, color=SOFT, align=PP_ALIGN.RIGHT)


def card(s, x, y, w, h, title, items, accent=BLUE):
    rect(s, x, y, w, h, fill=BG, line=LINE)
    rect(s, x, y, Inches(0.08), h, fill=accent)
    text(s, Emu(x + Inches(0.22)), Emu(y + Inches(0.12)), Emu(w - Inches(0.35)), Inches(0.32), title, size=14, color=INK, bold=True)
    bullets(s, Emu(x + Inches(0.22)), Emu(y + Inches(0.48)), Emu(w - Inches(0.4)), Emu(h - Inches(0.55)), items, size=12, mark=accent)


def table(s, y0, rows, col_w, sizes=None, rh=0.32):
    yy = y0
    if sizes is None: sizes = [10] * len(rows[0])
    for ri, row in enumerate(rows):
        xx = ML; h = Inches(0.38 if ri == 0 else rh)
        for ci, val in enumerate(row):
            cw = Emu(int(CW * col_w[ci]))
            bg = BLUE if ri == 0 else (ACC if ri % 2 == 0 else WHITE)
            fg = WHITE if ri == 0 else INK
            rect(s, xx, yy, cw, h, fill=bg, line=LINE, lw=0.5)
            text(s, Emu(xx + Inches(0.04)), Emu(yy + Inches(0.04)), Emu(cw - Inches(0.06)), Emu(h - Inches(0.06)),
                 str(val), size=sizes[ci], color=fg, bold=(ri == 0), anchor=MSO_ANCHOR.MIDDLE)
            xx = Emu(xx + cw)
        yy = Emu(yy + h)
    return yy


def add_pic(s, path, x, y, w, h):
    if os.path.exists(path):
        s.shapes.add_picture(path, x, y, width=w, height=h)
    else:
        rect(s, x, y, w, h, fill=ACC, line=LINE)
        text(s, x, y, w, h, "照片", size=12, color=SOFT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def build():
    prs = Presentation(); prs.slide_width = SW; prs.slide_height = SH
    TOTAL = 22

    # 1 cover
    s = slide(prs)
    text(s, ML, Inches(1.8), Inches(12), Inches(0.35), "上海创智汇 × 同浦汇｜活动专项交付", size=14, color=BLUE, bold=True)
    text(s, ML, Inches(2.3), Inches(12), Inches(0.8), "创智汇 · 30场活动怎么做", size=36, color=INK, bold=True)
    text(s, ML, Inches(3.2), Inches(12), Inches(0.4), "2026.08—2027.07　｜　表格说清楚：概况 · 价值 · 交付 · 分工 · 收费 · 案例", size=15, color=GOLD, bold=True)
    text(s, ML, Inches(3.75), Inches(12), Inches(0.35), "本册仅为活动运营专项交付，不含招商制度保障条款", size=13, color=ROSE, bold=True)
    card(s, ML, Inches(4.4), Inches(3.9), Inches(2.4), "修订要点", [
        "人数目标下调更可执行", "出海场次慎重、控成本", "同浦汇主责客户跟踪", "园区重心=后续活动+群内逼单",
    ], BLUE)
    card(s, Inches(4.7), Inches(4.4), Inches(3.9), Inches(2.4), "协同着眼点", [
        "活动结束立即建群", "同浦汇完成回访与台账", "园区在群内促进成交", "不把回访责任压给园区",
    ], GREEN)
    card(s, Inches(8.85), Inches(4.4), Inches(3.9), Inches(2.4), "商业建议（线下谈）", [
        "活动服务费打包/按场", "建议配套免租期", "付款节点建议方案", "佣金仅作参考非本册重点",
    ], GOLD)
    footer(s, 1, TOTAL)

    # 2 toc
    s = slide(prs); header(s, "目录", "本册结构（活动专项）")
    rows = [
        ["章节", "内容", "页"],
        ["01", "总体口径：关联度100%、人数修订、出海慎重", "03"],
        ["02", "30场概况总表（主题/时间/形式/人数/负责人占比）", "04–05"],
        ["03", "活动价值 · 具体交付 · 执行分工 · 建群逼单协同", "06–09"],
        ["04", "收费模式 · 免租期建议 · 付款节点建议", "10"],
        ["05", "单场执行节奏与转化（同浦汇主责跟踪）", "11"],
        ["06", "往期案例（企业+议程+现场照片）", "12–13"],
        ["07", "逐场执行要素表（行程/嘉宾/交付）", "14–17"],
        ["08", "一页纸结论", "18"],
    ]
    table(s, Inches(1.4), rows, [0.12, 0.72, 0.16], sizes=[12, 14, 12], rh=0.48)
    footer(s, 2, TOTAL)

    # 3 口径
    s = slide(prs); header(s, "01 · 总体口径", "三个关键口径先说清楚")
    card(s, ML, Inches(1.35), Inches(4.0), Inches(5.4), "① 孵化器关联度 = 100%", [
        "全部场次按园区产业主题定向邀约",
        "不设随机散客 / 公开路人渠道",
        "每场交付《关联度核验表》",
        "负责人占比按场次目标列示",
        "政府/高校/投资人/媒体为背书席位，单独标注",
    ], BLUE)
    card(s, Inches(4.75), Inches(1.35), Inches(4.0), Inches(5.4), "② 人数已下调", [
        "开放日：300→120–150人",
        "嘉年华：500→150–200人（改为主题日）",
        "联展：2000人次→累计约600人次",
        "Demo Day：200→80–120人",
        "全年触达目标同步下调为约600人次+",
        "以可执行、可核验为准",
    ], GOLD)
    card(s, Inches(8.95), Inches(1.35), Inches(3.8), Inches(5.4), "③ 出海慎重", [
        "不默认邀请境外企业到场",
        "F1改为国内出海路径/合规沙龙",
        "F2改为涉外服务准备专场",
        "领事/外企接待另案评估费用",
        "控制外事与接待成本",
    ], ROSE)
    footer(s, 3, TOTAL)

    # 4-5 总表
    for start, end, title, pg in ((0, 15, "30场概况总表（上）2026.08—12", 4), (16, 30, "30场概况总表（下）2027.01—07", 5)):
        s = slide(prs); header(s, "02 · 活动概况", title, "主题 · 时间 · 形式 · 执行节奏 · 人数 · 负责人占比 · 关联度")
        rows = [["编号", "主题/活动", "时间", "形式", "人数", "负责人%", "关联度", "价值落点"]]
        subset = EVENTS[:16] if start == 0 else EVENTS[16:]
        for e in subset:
            rows.append([e["code"], e["name"][:14], e["month"].replace("2026年", "").replace("2027年", "次"),
                         e["fmt"][:8], e["people"], e["dm"], e["assoc"], e["lease"][:12]])
        table(s, Inches(1.25), rows, [0.07, 0.24, 0.14, 0.12, 0.12, 0.09, 0.08, 0.14], sizes=[9, 10, 9, 9, 9, 9, 9, 9], rh=0.3)
        footer(s, pg, TOTAL)

    # 6 价值
    s = slide(prs); header(s, "03 · 活动价值", "活动到底带来什么（可衡量）")
    card(s, ML, Inches(1.35), Inches(4.0), Inches(5.4), "给园区", [
        "议题变日常，每月有主题不断档",
        "3F智能体/算力、5F内容IP标签坐实",
        "载体考核与活动补贴台账可归档",
        "待租单元获得持续看场曝光",
        "活动资产可衔接 WAIC 2027",
    ], BLUE)
    card(s, Inches(4.75), Inches(1.35), Inches(4.0), Inches(5.4), "给转化", [
        "客群100%主题定向，更准",
        "每场固定看场≥20分钟",
        "留资率目标≥60%，看场率≥25%",
        "活动结束建群，便于持续触达",
        "群内逼单促进成交（园区协同点）",
    ], GREEN)
    card(s, Inches(8.95), Inches(1.35), Inches(3.8), Inches(5.4), "不做的事", [
        "本册不做招商制度保障",
        "不对出租率/去化对赌",
        "不默认高成本外企接待",
        "不把回访责任压给园区",
    ], ROSE)
    footer(s, 6, TOTAL)

    # 7 交付
    s = slide(prs); header(s, "03 · 具体交付", "我们交给园区的，是这些具体的东西")
    rows = [
        ["周期", "交付物", "说明"],
        ["签约后2周", "年度活动方案细化版 + 排期日历", "活动专项"],
        ["每场结束后7天", "执行卡、议程、嘉宾确认、报名页、签到表、意向表、动线图、现场照片/通稿", "同浦汇归档"],
        ["每场结束后7天", "回访纪要（摘要知会版）", "交园区知会，不要求园区承担回访"],
        ["每场结束后7天", "活动社群（企微/微信群）建群清单", "便于园区在群内逼单"],
        ["每月", "活动数据月报 + 下月排期确认单", "联席会前"],
        ["每季/年末", "季度复盘、全年台账、WAIC衔接建议", "活动资产沉淀"],
    ]
    table(s, Inches(1.35), rows, [0.18, 0.5, 0.32], sizes=[12, 12, 12], rh=0.55)
    text(s, ML, Inches(6.0), CW, Inches(0.8),
         "说明：完整回访记录、报价跟进由同浦汇内部主责完成；交园区的回访纪要为摘要知会，减轻园区事务负担，把精力留给后续活动与群内促成。",
         size=12, color=MUT)
    footer(s, 7, TOTAL)

    # 8 分工
    s = slide(prs); header(s, "03 · 执行分工", "哪些同浦汇做，哪些请园区支持（写清楚）")
    rows = [
        ["事项", "同浦汇（主责）", "园区（支持）", "服中心/高校"],
        ["策划与议程", "全案策划、逐场执行卡", "方案审定、档期确认", "—"],
        ["邀约", "定向邀约、报名、关联度核验", "共享在谈名单（避免重复）", "推荐学者/成果团队"],
        ["场地保障", "提前7天提报需求、物料", "报备、导视停车、物业安保", "—"],
        ["现场执行", "主持控场、看场带队、建档", "开放样板间与看场动线", "政策宣讲位（如需）"],
        ["客户跟踪/回访", "主责到底：回访、报价、台账", "不承担逐户回访；可在群内逼单", "入驻手续（成交后）"],
        ["复盘", "单场复盘+月报", "联席确认下月活动安排", "—"],
        ["传播", "通稿、影像", "官方渠道转发（可选）", "联合会联动（可选）"],
    ]
    table(s, Inches(1.3), rows, [0.14, 0.32, 0.32, 0.22], sizes=[11, 11, 11, 11], rh=0.48)
    footer(s, 8, TOTAL)

    # 9 建群逼单协同
    s = slide(prs); header(s, "03 · 协同着眼点", "做完活动、建完群 → 园区在群里逼单促成交")
    card(s, ML, Inches(1.35), Inches(5.9), Inches(5.4), "标准动作", [
        "D日结束：同浦汇完成签到分层，当日建群",
        "T+1：同浦汇发回访与看场预约链接入群",
        "T+1~7：同浦汇主责逐户跟踪与报价",
        "园区角色：在群内发布房源/政策/限时条件，促进意向成交",
        "园区重心：排好下一场活动，保持节奏，不陷入逐户回访",
    ], GREEN)
    card(s, Inches(7.2), Inches(1.35), Inches(5.55), Inches(5.4), "责任边界", [
        "回访纪要：同浦汇做，交园区摘要知会",
        "逼单成交：园区可基于群与名单促进",
        "不把“客户跟踪制度”写成园区刚性义务",
        "本册聚焦活动交付，不做招商制度保障",
        "成交后的合同交割仍归园区流程",
    ], BLUE)
    footer(s, 9, TOTAL)

    # 10 收费+免租+付款
    s = slide(prs); header(s, "04 · 收费与付款建议", "活动专项收费 · 免租期建议 · 付款节点（线下讨论）")
    rows = [
        ["项目", "建议标准", "说明"],
        ["活动策划执行服务费", "年度打包或按场结算（金额线下确认）", "含策划执行、基础物料、主持协调、建群与回访"],
        ["办公租金基准（参考）", "3.3 元/㎡/天", "活动转化客户按园区口径执行"],
        ["建议免租期", "建议给予 1–3 个月免租（或阶梯免租）", "降低决策门槛，否则招商阻力大；具体面议"],
        ["物业费（参考）", "13.8 元/㎡/月", "园区统一标准"],
        ["待租房源", "8 间优先纳入看场动线", "有意向当场可看"],
        ["佣金（参考，非本册制度）", "若发生成交，可参考 2 个月租金", "本册不做招商制度保障，仅供线下参考"],
    ]
    table(s, Inches(1.25), rows, [0.22, 0.4, 0.38], sizes=[11, 12, 11], rh=0.42)
    text(s, ML, Inches(5.0), CW, Inches(0.3), "付款节点建议方案（供线下讨论，非最终条款）", size=13, color=BLUE, bold=True)
    rows2 = [
        ["节点", "建议比例", "对应交付"],
        ["合同签署后 7 个工作日", "30%", "启动策划、年度日历锁定"],
        ["每季度首月", "20% × 3 季（共60%）", "按季度场次执行与月报验收"],
        ["年度收官（Demo Day后）", "10%", "年报、台账、影像资产移交"],
    ]
    table(s, Inches(5.35), rows2, [0.34, 0.28, 0.38], sizes=[11, 11, 11], rh=0.36)
    footer(s, 10, TOTAL)

    # 11 执行节奏
    s = slide(prs); header(s, "05 · 单场执行节奏", "同一套打法；客户跟踪由同浦汇主责")
    rows = [
        ["节点", "动作", "责任", "交付"],
        ["T-14", "锁定嘉宾议程物料报名", "同浦汇", "议程/嘉宾函/报名页"],
        ["T-7", "邀约确认+场地报备", "同浦汇主责；园区场地确认", "名单/场地单"],
        ["T-1", "彩排+看场动线", "同浦汇；园区样板间就绪", "彩排记录"],
        ["D日", "签到→主区→看场≥20分钟→建群", "同浦汇执行；园区现场支持", "签到/意向/群"],
        ["T+1~7", "逐户回访、报价、三轮跟进", "同浦汇主责到底", "回访台账；摘要知会园区"],
        ["群内协同", "发布政策/房源/限时条件促成交", "园区可在群内逼单", "成交线索反馈"],
    ]
    table(s, Inches(1.35), rows, [0.12, 0.34, 0.3, 0.24], sizes=[11, 12, 11, 11], rh=0.55)
    footer(s, 11, TOTAL)

    # 12-13 cases
    for i, c in enumerate(CASES):
        s = slide(prs); header(s, f"06 · 往期案例（{i+1}）", c["name"], c["meta"])
        for j, fn in enumerate(c["photos"][:3]):
            add_pic(s, os.path.join(PHOTO, fn), Emu(ML + j * Inches(4.1)), Inches(1.35), Inches(3.95), Inches(2.3))
        rect(s, ML, Inches(3.85), Inches(7.5), Inches(2.95), fill=BG, line=LINE)
        text(s, Emu(ML + Inches(0.2)), Inches(3.95), Inches(7.1), Inches(0.3), "当天议程", size=13, color=BLUE, bold=True)
        bullets(s, Emu(ML + Inches(0.2)), Inches(4.35), Inches(7.1), Inches(2.2), c["agenda"], size=13)
        rect(s, Inches(8.3), Inches(3.85), Inches(4.5), Inches(2.95), fill=ACC, line=LINE)
        text(s, Inches(8.45), Inches(3.95), Inches(4.2), Inches(0.3), "可见机构/企业", size=13, color=BLUE, bold=True)
        text(s, Inches(8.45), Inches(4.35), Inches(4.2), Inches(2.2),
             f"{c['orgs']}\n\n企业侧：{c['companies']}\n\n映射：{c['map']}", size=12, color=INK)
        footer(s, 12 + i, TOTAL)

    # 14-17 逐场要素
    chunks = [(EVENTS[:8], 14), (EVENTS[8:16], 15), (EVENTS[16:23], 16), (EVENTS[23:], 17)]
    for subset, pg in chunks:
        s = slide(prs); header(s, "07 · 逐场执行要素", f"行程 / 嘉宾 / 交付（{subset[0]['code']}–{subset[-1]['code']}）")
        rows = [["编号", "活动", "行程", "嘉宾", "交付"]]
        for e in subset:
            rows.append([e["code"], e["name"][:12], e["agenda"][:26], e["guests"][:16], "标准包+建群"])
        table(s, Inches(1.3), rows, [0.07, 0.2, 0.32, 0.26, 0.15], sizes=[10, 10, 10, 10, 10], rh=0.48)
        footer(s, pg, TOTAL)

    # 18 conclusion
    s = slide(prs); header(s, "08 · 一页纸结论", "可直接回复对方的修订要点")
    card(s, ML, Inches(1.35), Inches(4.0), Inches(5.4), "我们交付", [
        "30场可执行排期（人数已下调）",
        "关联度100%定向邀约+核验表",
        "逐场行程/嘉宾/交付要素",
        "同浦汇主责回访跟踪与台账",
        "活动建群，便于园区群内逼单",
        "往期案例含企业、议程、现场照",
    ], BLUE)
    card(s, Inches(4.75), Inches(1.35), Inches(4.0), Inches(5.4), "请园区协同", [
        "场地物业与样板间看场",
        "共享在谈名单、确认档期",
        "大场现场支持（按档配置）",
        "活动后在群内发布条件逼单",
        "重心放在后续活动节奏",
        "不承担逐户回访事务",
    ], GREEN)
    card(s, Inches(8.95), Inches(1.35), Inches(3.8), Inches(5.4), "线下再谈", [
        "活动服务费金额",
        "免租期 1–3 个月建议",
        "付款节点建议方案",
        "出海慎重场是否另案",
        "佣金仅参考，非本册制度",
    ], GOLD)
    footer(s, 18, TOTAL)

    # keep remaining pages as spare detail? User asked for one PPT - 18 is enough but TOTAL was 22.
    # Add 4 more useful pages: 六档标准、出海说明、付款详细、封面后说明

    # 19 六档
    s = slide(prs); header(s, "附录 · 六大主题", "场次结构（与运营方案一致，人数已修订）")
    rows = [
        ["档", "主题", "场次", "人数口径", "关联度"],
        ["A", "智能体与算力训练", "7", "30–50人为主", "100%"],
        ["B", "政策与治理沙龙", "4", "30–40人", "100%"],
        ["C", "产业与科学对接", "6", "30–50人", "100%"],
        ["D", "AI内容与IP", "5", "60–200人（控规模）", "100%"],
        ["E", "WAIC成果承接", "4", "开放日≤150；主题日≤200", "100%"],
        ["F", "出海慎重场/AGI/收官", "4", "出海场20–40人；Demo≤120", "100%"],
    ]
    table(s, Inches(1.4), rows, [0.1, 0.28, 0.12, 0.35, 0.15], sizes=[12, 13, 12, 12, 12], rh=0.55)
    footer(s, 19, TOTAL)

    # 20 出海专项说明
    s = slide(prs); header(s, "附录 · 出海场次说明", "涉及外国企业/领事到场：慎重立项，费用另案")
    card(s, ML, Inches(1.35), Inches(6.0), Inches(5.4), "本册默认做法", [
        "F1：国内出海路径与合规沙龙",
        "F2：涉外服务准备专场（闭门）",
        "默认不邀请境外企业到场",
        "默认不安排领事出席",
        "先用低成本场次验证需求",
    ], BLUE)
    card(s, Inches(7.25), Inches(1.35), Inches(5.5), Inches(5.4), "若确需升级", [
        "单独立项：国际买家/领事专题",
        "单独预算：接待、翻译、外事流程",
        "单独排期：与常规30场可替换/追加",
        "书面确认费用与责任边界后再执行",
        "避免把高成本默认写进年度包",
    ], ROSE)
    footer(s, 20, TOTAL)

    # 21 付款建议细化
    s = slide(prs); header(s, "附录 · 付款节点建议", "供线下讨论的两种备选（非正式条款）")
    card(s, ML, Inches(1.35), Inches(5.9), Inches(5.4), "方案A · 季度验收型（推荐）", [
        "签约后7日内：30%启动款",
        "Q2/Q3/Q4季初：各20%（合计60%）",
        "年终收官后：10%尾款",
        "与季度场次完成度、月报验收挂钩",
        "适合全年持续执行",
    ], GREEN)
    card(s, Inches(7.2), Inches(1.35), Inches(5.55), Inches(5.4), "方案B · 半年结算型", [
        "签约后：40%启动",
        "半年复盘后：40%",
        "年终移交资产后：20%",
        "节点更少，适合简化财务流程",
        "具体比例双方线下确认",
    ], BLUE)
    text(s, ML, Inches(6.9), CW, Inches(0.3), "以上均为建议，最终以商务合同为准。", size=12, color=SOFT)
    footer(s, 21, TOTAL)

    # 22 end
    s = slide(prs)
    text(s, ML, Inches(2.4), Inches(12), Inches(0.5), "谢谢审阅", size=36, color=INK, bold=True, align=PP_ALIGN.CENTER)
    text(s, ML, Inches(3.2), Inches(12), Inches(0.4), "创智汇 · 30场活动专项交付（修订版）", size=18, color=BLUE, bold=True, align=PP_ALIGN.CENTER)
    text(s, ML, Inches(3.8), Inches(12), Inches(0.4), "人数更可执行 · 出海更慎重 · 回访同浦汇主责 · 园区群内逼单协同", size=14, color=MUT, align=PP_ALIGN.CENTER)
    text(s, ML, Inches(4.5), Inches(12), Inches(0.35), "编制：同浦汇　·　仅活动部分　·　付款与免租期建议供线下讨论", size=13, color=SOFT, align=PP_ALIGN.CENTER)
    footer(s, 22, TOTAL)

    os.makedirs(ART, exist_ok=True)
    outs = [
        "创智汇30场活动专项交付-修订版.pptx",
        "chuangzhihui-30-events-activity-only-rev.pptx",
        "创智汇30场活动具体方案.pptx",
        "chuangzhihui-30-events-plan.pptx",
        "创智汇年度活动运营方案-融合版.pptx",
        "chuangzhihui-annual-ops-plan-fused.pptx",
    ]
    tmp = os.path.join(HERE, "_tmp_activity_rev.pptx")
    prs.save(tmp)
    for fn in outs:
        p = os.path.join(HERE, fn)
        shutil.copy2(tmp, p)
        shutil.copy2(p, os.path.join(ART, fn))
    os.remove(tmp)
    print("PPT slides:", len(prs.slides))

    # Excel
    wb = Workbook()
    thin = Border(left=Side(style="thin", color="D0D0D0"), right=Side(style="thin", color="D0D0D0"),
                  top=Side(style="thin", color="D0D0D0"), bottom=Side(style="thin", color="D0D0D0"))
    head = PatternFill("solid", fgColor="1E4D8C")
    alt = PatternFill("solid", fgColor="EEF2F8")

    def hdr(ws, headers):
        for j, h in enumerate(headers, 1):
            c = ws.cell(1, j, h); c.font = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)
            c.fill = head; c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = thin

    def write(ws, data, widths, h=46):
        for i, row in enumerate(data, 2):
            for j, v in enumerate(row, 1):
                c = ws.cell(i, j, v); c.font = Font(name="微软雅黑", size=9)
                c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = thin
                if i % 2 == 0: c.fill = alt
            ws.row_dimensions[i].height = h
        ws.row_dimensions[1].height = 28
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    ws = wb.active; ws.title = "01-活动概况"
    hdr(ws, ["编号", "主题", "时间", "形式", "执行节奏", "人数", "负责人占比", "关联度", "场地", "价值落点"])
    write(ws, [[e["code"], e["name"], e["month"], e["fmt"], e["rhythm"], e["people"], e["dm"], e["assoc"], e["floor"], e["lease"]] for e in EVENTS],
          [8, 28, 16, 14, 36, 16, 10, 8, 14, 22])

    ws2 = wb.create_sheet("02-价值交付分工收费")
    hdr(ws2, ["模块", "内容"])
    write(ws2, [
        ["活动价值", "主题日常化；看场转化；载体台账；建群持续触达；园区群内逼单"],
        ["具体交付", "方案议程嘉宾函报名签到意向动线照片通稿；回访摘要知会版；建群清单；月季年报"],
        ["同浦汇主责", "策划邀约执行、客户跟踪回访报价台账、建群"],
        ["园区支持", "场地物业样板间、名单档期、大场支持；群内逼单；重心放后续活动"],
        ["关联度", "定向邀约100%，不设散客"],
        ["出海慎重", "默认不邀请外企/领事到场；F1/F2为国内专场；升级另案"],
        ["租金参考", "3.3元/㎡/天；建议免租期1–3个月"],
        ["物业参考", "13.8元/㎡/月"],
        ["待租", "8间优先看场"],
        ["活动服务费", "打包或按场，金额线下确认"],
        ["付款建议A", "签约30% + 每季20%×3 + 年终10%"],
        ["付款建议B", "签约40% + 半年40% + 年终20%"],
        ["范围声明", "本册仅为活动专项，不做招商制度保障"],
    ], [16, 70], 32)

    ws3 = wb.create_sheet("03-逐场行程嘉宾")
    hdr(ws3, ["编号", "活动", "行程", "嘉宾", "交付"])
    write(ws3, [[e["code"], e["name"], e["agenda"], e["guests"], e["deliver"]] for e in EVENTS],
          [8, 28, 42, 36, 36], 50)

    ws4 = wb.create_sheet("04-往期案例")
    hdr(ws4, ["案例", "时间地点", "机构", "企业", "议程", "照片", "映射"])
    write(ws4, [[c["name"], c["meta"], c["orgs"], c["companies"], " | ".join(c["agenda"]), "；".join(c["photos"]), c["map"]] for c in CASES],
          [32, 28, 36, 36, 40, 30, 24], 60)

    xouts = [
        "创智汇30场活动专项交付-修订版.xlsx",
        "chuangzhihui-30-events-activity-only-rev.xlsx",
        "创智汇30场活动方案排期表.xlsx",
        "chuangzhihui-30-events-schedule.xlsx",
        "创智汇年度活动运营方案-融合版.xlsx",
        "chuangzhihui-annual-ops-plan-fused.xlsx",
    ]
    for fn in xouts:
        p = os.path.join(HERE, fn); wb.save(p); shutil.copy2(p, os.path.join(ART, fn))
    print("Excel OK")


if __name__ == "__main__":
    build()
