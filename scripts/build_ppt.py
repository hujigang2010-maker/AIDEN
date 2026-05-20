"""生成 5/22 峰会会务汇报 PPT。"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

OUT = "/workspace/5月22日峰会/导出文件/2026峰会会务汇报.pptx"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Colors
NAVY = RGBColor(0x1F, 0x4E, 0x79)
BLUE = RGBColor(0x30, 0x54, 0x96)
LIGHT_BLUE = RGBColor(0xD9, 0xE1, 0xF2)
DARK_TEXT = RGBColor(0x33, 0x33, 0x33)
GOLD = RGBColor(0xBF, 0x9D, 0x37)
RED = RGBColor(0xC0, 0x39, 0x2B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0x7F, 0x7F, 0x7F)

FONT_CN = "WenQuanYi Micro Hei"


def set_run(run, text, size=14, bold=False, color=DARK_TEXT, font=FONT_CN):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def add_textbox(slide, x, y, w, h, text, size=14, bold=False, color=DARK_TEXT,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    set_run(p.add_run(), text, size=size, bold=bold, color=color)
    return tb


def add_rect(slide, x, y, w, h, fill, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background() if line is None else None
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
    shape.shadow.inherit = False
    return shape


def add_header(slide, title, sub=None):
    add_rect(slide, 0, 0, 13.333, 0.9, NAVY)
    add_textbox(slide, 0.4, 0.1, 12, 0.7, title, size=24, bold=True,
                color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    if sub:
        add_textbox(slide, 10, 0.25, 3.2, 0.5, sub, size=12, color=LIGHT_BLUE,
                    align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(slide, 0, 7.15, 13.333, 0.35, BLUE)
    add_textbox(slide, 0.4, 7.18, 12, 0.3, "2026 AI 商业化落地与硬核投资破局峰会 · 5/22 一滴水",
                size=10, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)


def add_table(slide, x, y, w, h, headers, rows, header_color=BLUE,
              header_font_size=12, body_font_size=10, col_widths=None):
    n_cols = len(headers)
    n_rows = 1 + len(rows)
    tbl_shape = slide.shapes.add_table(n_rows, n_cols, Inches(x), Inches(y), Inches(w), Inches(h))
    tbl = tbl_shape.table
    if col_widths:
        total = sum(col_widths)
        for i, cw in enumerate(col_widths):
            tbl.columns[i].width = Inches(w * cw / total)
    # header
    for i, h_text in enumerate(headers):
        cell = tbl.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_color
        cell.text = ""
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        set_run(p.add_run(), h_text, size=header_font_size, bold=True, color=WHITE)
    # rows
    for ri, row in enumerate(rows, 1):
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if ri % 2 else LIGHT_BLUE
            cell.text = ""
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            set_run(p.add_run(), str(val) if val is not None else "", size=body_font_size, color=DARK_TEXT)
    return tbl


def blank_slide():
    return prs.slides.add_slide(prs.slide_layouts[6])


# ============ 封面 ============
s = blank_slide()
add_rect(s, 0, 0, 13.333, 7.5, NAVY)
add_rect(s, 0, 5.5, 13.333, 0.05, GOLD)
add_textbox(s, 0.5, 1.8, 12.3, 1.2,
            "2026 人工智能商业化落地与硬核投资破局峰会",
            size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_textbox(s, 0.5, 3.0, 12.3, 0.7,
            "重构与突围 · 寻找 AI 时代的超级个体、新质资产与资本新风口",
            size=22, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)
add_textbox(s, 0.5, 4.0, 12.3, 0.6,
            "5 月 22 日会务执行汇报",
            size=24, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
add_textbox(s, 0.5, 5.8, 12.3, 0.5,
            "时间：2026 年 5 月 22 日 13:00-20:30   |   地点：上海·北外滩·一滴水",
            size=16, color=WHITE, align=PP_ALIGN.CENTER)
add_textbox(s, 0.5, 6.3, 12.3, 0.5,
            "主办：北京大学经济学院上海校友会 · 复旦大学住房政策研究中心",
            size=14, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)
add_textbox(s, 0.5, 6.8, 12.3, 0.4,
            "晚宴冠名战略合作伙伴：绿城中国 |【绿城·潮鸣】星耀北外滩",
            size=14, color=GOLD, align=PP_ALIGN.CENTER)

# ============ 目录 ============
s = blank_slide()
add_header(s, "汇报目录", "Contents")
contents = [
    "01  活动总览与组织架构",
    "02  议程与流程",
    "03  VIP 出席嘉宾胸卡名单",
    "04  晚宴 · 圆桌分桌",
    "05  论坛 / 主席台台卡",
    "06  理事接龙与志愿者分组",
    "07  赞助方对接与物资清单",
    "08  赞助商口播脚本",
    "09  现场摄影特写需求",
    "10  关键时间节点 & 待补事项",
]
for i, t in enumerate(contents):
    add_rect(s, 1.5 + (i % 2) * 5.5, 1.5 + (i // 2) * 0.95, 5, 0.7, LIGHT_BLUE)
    add_textbox(s, 1.7 + (i % 2) * 5.5, 1.55 + (i // 2) * 0.95, 4.8, 0.6,
                t, size=16, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)

# ============ 01 总览 ============
s = blank_slide()
add_header(s, "01 / 活动总览与组织架构")
add_table(s, 0.5, 1.2, 12.3, 5.7,
          ["项目", "内容"],
          [
              ("活动主题", "重构与突围 — 寻找 AI 时代的超级个体、新质资产与资本新风口"),
              ("活动时间", "2026 年 5 月 22 日 13:00 - 20:30"),
              ("活动地点", "上海·北外滩·一滴水"),
              ("主办单位", "北京大学经济学院上海校友会、复旦大学住房政策研究中心"),
              ("协办单位", "复旦未来信息创新学院、上海云计算创新基地、杨浦/虹口科企联、东投资联盟 等"),
              ("战略合作伙伴", "腾讯云、福布斯中国人工智能科技企业评选委员会"),
              ("晚宴冠名战略合作伙伴", "绿城中国（晚宴：【绿城·潮鸣】星耀北外滩——AI 领袖定制晚宴，凭专属邀请码入场）"),
              ("特约赞助商", "蔚来汽车、泰隆银行"),
              ("赞助方", "茶活力、美年大健康、复金汇、古井贡酒·年份原浆"),
              ("媒体支持", "新华网 / 人民网 / 央视网 / 上海证券报 / 财联社 / SMG / 左明右理 等"),
              ("总联系人", "胡继刚 13262607888 · 王珏 13817642219 · 安淑娟 13661612711"),
          ], col_widths=[3, 9.5], body_font_size=11)

# ============ 02 议程 ============
s = blank_slide()
add_header(s, "02 / 议程与流程")
add_table(s, 0.4, 1.1, 12.5, 5.9,
          ["时间", "环节", "主持 / 主讲"],
          [
              ("13:00-13:30", "嘉宾签到与入场（凭邀请码核验）", "—"),
              ("13:30-13:35", "开幕致辞", "姚志勇 教授"),
              ("13:35-14:00", "主旨 1 · 大模型时代的底层架构突围与商业闭环", "白硕"),
              ("14:00-14:25", "主旨 2 · AI 如何改变我们的生活与投资", "夏春 博士"),
              ("14:25-14:40", "★ 第四届 2026 人工智能商业化落地颁奖典礼（高光启幕）", "—"),
              ("14:40-15:20", "圆桌一 · AI 硬核圆桌——技术突围与新质生产力落地", "主持：王珏"),
              ("15:20-16:00", "圆桌二 · 巅峰对话——从算力引擎到新质资产", "主持：胡继刚"),
              ("16:00-16:50", "陆家嘴交响乐团 · 时光音乐会", "—"),
              ("16:50-17:10", "主旨 3 · 特殊机遇/不良资产投资的新机遇", "王维军"),
              ("17:10-17:30", "主旨 4 · 从自然利率下降到 K 型经济", "寇文红"),
              ("17:30-17:50", "主旨 5 · 从致用之学到量子智能", "张露瑶 博士"),
              ("17:50-18:30", "圆桌三 · 投资大圆桌——寻找新质资产风口与闭幕仪式", "主持：黄欣"),
              ("18:30-20:30", "【绿城·潮鸣】星耀北外滩——AI 领袖定制晚宴（凭专属邀请码入场）", "冠名：绿城中国"),
          ], col_widths=[1.8, 7.7, 3], body_font_size=10)

# ============ 03 VIP 胸卡 ============
s = blank_slide()
add_header(s, "03 / VIP 出席嘉宾胸卡 — 按角色色标")
add_table(s, 0.4, 1.1, 12.5, 2.5,
          ["类别", "数量", "色标", "代表嘉宾"],
          [
              ("主席团 / 致辞", "1", "★ 深红", "姚志勇"),
              ("主旨演讲嘉宾", "5", "★ 金", "白硕、夏春、王维军、寇文红、张露瑶"),
              ("圆桌一 / 二 / 三 主持+嘉宾", "14", "● 蓝", "王珏 / 胡继刚 / 黄欣 + 三组嘉宾"),
              ("出席理事", "6", "● 紫", "王珏、王维军、何杭、童长征、冯旭南、张景瑞"),
              ("战略 / 冠名 / 赞助 VIP", "约 25-30", "● 橙", "绿城、腾讯云、泰隆、复金汇 等"),
              ("一般出席嘉宾", "60+", "● 绿", "投资 / 银行 / 产业 / 学术 / 媒体"),
              ("志愿者 / 工作人员", "23", "● 灰", "A-G 7 个分组"),
          ], col_widths=[3, 1.5, 1.5, 6.5])
add_textbox(s, 0.4, 3.8, 12.5, 0.4, "建议规格：90×130mm 横版；正面 姓名 + 公司 + 职务 + 角色色标；背面流程二维码 + 紧急联系电话",
            size=12, bold=True, color=NAVY)
add_textbox(s, 0.4, 4.4, 12.5, 0.4, "胸卡总数：约 500 张（含 30 张备用），由 张蒙（F 组）负责印刷，安淑娟 校对，5/21 18:00 前完成。",
            size=12, color=GRAY)
add_textbox(s, 0.4, 5.0, 12.5, 0.4, "待补：绿城 5 席名单；福布斯、蔚来、茶活力、美年大健康、古井贡酒 出席嘉宾名单。",
            size=12, color=RED)

# ============ 04 晚宴圆桌 ============
s = blank_slide()
add_header(s, "04 / 晚宴 · 圆桌分桌（T01-T10 重点桌）")
add_table(s, 0.4, 1.1, 12.5, 5.7,
          ["桌号", "主题", "核心嘉宾"],
          [
              ("T01", "VIP 主桌（主席团 / 主旨 / 冠名）", "姚志勇、白硕、夏春、王维军、寇文红、张露瑶、绿城×2、胡继刚、黄欣"),
              ("T02", "圆桌一 + 圆桌二（巅峰对话）嘉宾桌", "王珏、李龙、齐宝鑫、刘文平、刘帅华、徐泳泽、黎远、刘剑英、于敬、孔华威"),
              ("T03", "圆桌三（投资大圆桌+闭幕）+ 理事桌", "曲承东、刘胜利、饶雪莹、马俊杰 + 张景瑞、何杭、童长征、冯旭南、葛九明、朱震"),
              ("T04", "腾讯云 / 上海云基地 / 绿城联合桌", "周昭、孙总、李海洋、高歆、顾炯亮、陈艳 + 绿城预留 4 席"),
              ("T05", "泰隆银行虹口支行桌", "林骁、杨红江、汤梦薇、阎晓华、徐至柔、毛瑞香 等"),
              ("T06", "泰隆银行杨浦支行桌", "陈煜、赵逸飞、晋嘉乐 + 7 家科企客户"),
              ("T07", "投资机构桌 A", "李莺莺、王梦雅、郭紫奇、林雪、李克勤、刘中原、罗晓、邵彩霞 等"),
              ("T08", "银行 / 金融机构桌", "赵梓翔、陈颖芳、冯浩然、潘健慧、张涌、张稳、魏军、唐亮 等"),
              ("T09", "产业 / 科技企业桌", "张蒙、谈英豪、张戟、王梦、黄勇、田甜、唐伟、邱诗灵、姜逸斐、杨峰"),
              ("T10", "学术 / 政府 / 媒体桌", "周舟、杨知予、任广焕、刘晓、李柳阳 + 其余嘉宾"),
              ("T11-T50", "待排（北大校友 / 律所 / 媒体 / 家办 / 普通晚宴桌）", "5/21 22:00 前由 王珏 / 黄欣 / 安淑娟 定稿"),
          ], col_widths=[1.2, 4.3, 7], body_font_size=10)

# ============ 05 台卡 ============
s = blank_slide()
add_header(s, "05 / 论坛 · 主席台 台卡名单")
add_table(s, 0.4, 1.1, 12.5, 5.7,
          ["时间", "环节", "摆位 / 角色", "姓名", "单位"],
          [
              ("13:30", "开幕致辞", "主席台中央", "姚志勇 教授", "复旦管院 / 北大经院上海校友会会长"),
              ("13:35", "主旨 1", "演讲台", "白硕", "恒生电子研究院院长"),
              ("14:00", "主旨 2", "演讲台", "夏春 博士", "香港国际金融学会副会长"),
              ("14:25", "★ 颁奖典礼", "主礼嘉宾 ×5", "姚志勇、胡继刚、黄欣、福布斯-代表、绿城-代表", "高光启幕"),
              ("14:40", "圆桌一", "主持 + 嘉宾 ×4", "王珏 + 白硕、李龙、齐宝鑫、刘文平", "AI 硬核圆桌"),
              ("15:20", "圆桌二（巅峰对话）", "主持 + 嘉宾 ×5", "胡继刚 + 刘帅华、徐泳泽、黎远、刘剑英、于敬", "—"),
              ("16:50", "主旨 3", "演讲台", "王维军", "泓塬资产董事长"),
              ("17:10", "主旨 4", "演讲台", "寇文红", "—"),
              ("17:30", "主旨 5", "演讲台", "张露瑶 博士", "昆山杜克大学助理教授"),
              ("17:50", "圆桌三 + 闭幕", "主持 + 嘉宾 ×4", "黄欣 + 曲承东、刘胜利、饶雪莹、马俊杰", "投资大圆桌"),
          ], col_widths=[1.4, 2.5, 2.5, 3.5, 2.6], body_font_size=10)
add_textbox(s, 0.4, 6.85, 12.5, 0.3, "输出总数：32 张（含备用 5 张）；★ 颁奖典礼前移至 14:25-14:40；福布斯颁奖物料须 5/22 12:00 前到位。",
            size=11, bold=True, color=RED)

# ============ 06 理事 + 志愿者 ============
s = blank_slide()
add_header(s, "06 / 理事接龙 + 志愿者分组")
add_textbox(s, 0.4, 1.1, 12.5, 0.4, "出席理事（5/22 一滴水二楼论坛 6 位）", size=14, bold=True, color=NAVY)
add_table(s, 0.4, 1.55, 12.5, 1.7,
          ["#", "姓名", "校友", "到场", "全程"],
          [
              (1, "王珏", "—", "开场前 1 小时", "全程"),
              (2, "王维军", "93 国经", "正常出席", "全程"),
              (3, "何杭", "03 国贸", "开场前 1 小时", "全程"),
              (4, "童长征", "99 金融", "正常出席", "全程"),
              (5, "冯旭南", "02 硕士", "正常出席", "全程"),
              (6, "张景瑞", "98 硕士", "正常出席", "全程"),
          ], col_widths=[0.6, 1.5, 1.8, 2.5, 1.6])

add_textbox(s, 0.4, 3.35, 12.5, 0.4, "志愿者分组（A-G + 外部协作，共 23 人）", size=14, bold=True, color=NAVY)
add_table(s, 0.4, 3.8, 12.5, 3.2,
          ["组", "职责", "成员（含主联系人）"],
          [
              ("A", "签到引导", "韦佳玉 17301836967、李政春"),
              ("B", "台卡 / 物料", "王胜、朱铭喆、蔡杰、江梳桐"),
              ("C", "VIP 接待", "陈潇 Kelly、王珏（组长）、冯墨、黄露、吕志翔"),
              ("D", "圆桌对接 / 上下场", "葛九明、李正超"),
              ("E", "晚宴布置 / 席卡", "张卓、高辰辰、随圣博、安淑娟（13661612711 总校对）"),
              ("F", "印刷 / 物料", "朱俊峰、彭常丽、张蒙（席卡责任人）"),
              ("G", "现场调度 / 摄影支援", "刘严、李兵、蔡萍、马磊"),
              ("外部", "总联系 / 摄影 / 协作", "胡继刚 13262607888、洪小燕（摄影）15021488859、江江、皮尔德小C"),
          ], col_widths=[0.8, 3, 8.7], body_font_size=10)

# ============ 07 赞助物资 ============
s = blank_slide()
add_header(s, "07 / 赞助方对接与物资清单")
add_table(s, 0.4, 1.1, 12.5, 5.7,
          ["#", "赞助方（级别）", "品名 / 数量", "送达", "对接人", "签收"],
          [
              (1, "绿城中国（冠名）", "【绿城·潮鸣】KT 板 + 邀请码核验物料", "5/22 09:00", "待补", "葛九明"),
              (2, "绿城中国（冠名）", "晚宴定制礼品 500 份", "5/22 14:00", "待补", "张蒙"),
              (3, "腾讯云（战略）", "易拉宝 / 礼品", "5/22 10:00", "徐泳泽", "葛九明"),
              (4, "福布斯中国 AI 评选委员会（战略）", "★ 奖杯 / 证书 / 颁奖背景板", "5/22 12:00 必到", "待补", "张蒙"),
              (5, "蔚来汽车（特约）", "展车 / 易拉宝 / 礼品", "5/22 08:00 / 10:00", "待补", "葛九明、李正超"),
              (6, "泰隆银行（特约）", "易拉宝 / 客户伴手礼", "5/22 10:00", "林骁 13058861308 / 陈煜 13671896951", "葛九明"),
              (7, "茶活力（赞助）", "茶饮 / 体验台", "5/22 09:00", "待补", "蔡萍"),
              (8, "美年大健康（赞助）", "健康检测 / 礼品", "5/22 09:00", "待补", "蔡萍"),
              (9, "复金汇（赞助）", "礼品 / 易拉宝", "5/22 10:00", "袁一栋 15995991077", "马磊"),
              (10, "古井贡酒·年份原浆", "晚宴用酒 50 桌×2 + 礼盒", "5/22 14:00", "待补", "张蒙"),
          ], col_widths=[0.6, 2.7, 3.2, 1.8, 3, 1.4], body_font_size=10)

# ============ 08 口播 ============
s = blank_slide()
add_header(s, "08 / 赞助商口播脚本")
add_textbox(s, 0.4, 1.1, 12.5, 0.4, "A · 开场总致谢（13:30 姚志勇 致辞中嵌入）", size=14, bold=True, color=NAVY)
add_rect(s, 0.4, 1.55, 12.5, 0.95, LIGHT_BLUE)
add_textbox(s, 0.55, 1.62, 12.2, 0.85,
            "本届峰会得以呈现，特别感谢——晚宴冠名战略合作伙伴 绿城中国；战略合作伙伴 腾讯云、福布斯中国人工智能科技企业评选委员会；特约赞助商 蔚来汽车、泰隆银行；赞助方 茶活力、美年大健康、复金汇、古井贡酒·年份原浆。让我们再次以掌声致谢。",
            size=12, color=DARK_TEXT)

add_textbox(s, 0.4, 2.6, 12.5, 0.4, "B · 分时段定点口播（按官方最终定稿议程）", size=14, bold=True, color=NAVY)
add_table(s, 0.4, 3.05, 12.5, 3.9,
          ["时间", "环节", "口播内容（节选）", "口播人"],
          [
              ("14:25", "★ 颁奖典礼", "感谢福布斯中国 AI 评选委员会 联合颁布本届年度领军机构。", "姚志勇 / 胡继刚"),
              ("14:40", "圆桌一前", "感谢战略合作伙伴 腾讯云 的鼎力支持。", "王珏"),
              ("15:20", "圆桌二（巅峰对话）前", "巅峰对话【从算力引擎到新质资产】由 腾讯云 联合呈现。", "胡继刚"),
              ("16:00", "中场前", "中场由 茶活力 与 美年大健康 共同支持。", "主持人"),
              ("16:50", "主旨 3 前", "感谢特约赞助商 泰隆银行，深耕科技金融。", "主持人"),
              ("17:50", "圆桌三 + 闭幕前", "投资大圆桌由 复金汇 战略支持，本环节同时为闭幕仪式。", "黄欣"),
              ("18:30", "【绿城·潮鸣】晚宴开场", "晚宴由 绿城中国 冠名，古井贡酒 为晚宴用酒，蔚来汽车 为出行合作伙伴。", "王珏 / 胡继刚"),
          ], col_widths=[1.3, 2.6, 7.1, 1.5], body_font_size=10)

# ============ 09 摄影特写 ============
s = blank_slide()
add_header(s, "09 / 现场摄影特写需求（洪小燕 15021488859）")
add_table(s, 0.4, 1.1, 12.5, 5.8,
          ["#", "赞助方", "特写镜头", "触发节点", "输出"],
          [
              (1, "绿城中国", "【绿城·潮鸣】LOGO 全景、冠名牌特写、邀请码核验入场剪影、致辞特写", "18:00 / 18:30 / 致辞", "横竖各 5 + 视频 30s"),
              (2, "腾讯云", "LOGO 墙、徐泳泽圆桌二发言镜头 ≥6", "13:00 / 15:20-16:00 / 闭幕", "嘉宾 6 + 物料 4"),
              (3, "福布斯中国 AI 评选委员会", "★ 奖杯 / 证书、颁奖瞬间合影、获奖代表镜头", "★ 14:25-14:40 颁奖典礼", "合影 ≥5 + 1min 视频"),
              (4, "蔚来汽车", "展车 + 内饰 + 嘉宾驻足", "13:00 / 中场 16:00", "展车 ≥8"),
              (5, "泰隆 虹口 / 杨浦", "T05/T06 双桌全景 + 行长致辞", "18:30-19:30", "每支行各 5"),
              (6, "茶活力", "茶歇区品茶 + 体验台特写", "16:00-16:50", "≥6 + 15s 视频"),
              (7, "美年大健康", "健康检测 + 嘉宾参与", "全天展位", "≥6"),
              (8, "复金汇", "签到礼品台 + 嘉宾领取 + 圆桌三镜头", "13:00 / 17:50-18:30", "≥6"),
              (9, "古井贡酒", "餐桌酒瓶 + 嘉宾举杯", "18:30 / 19:00 / 19:30", "≥8"),
              (10, "全体合作伙伴", "闭幕大合影 + LOGO 墙合影", "18:30", "1-2 张高清"),
          ], col_widths=[0.6, 2.4, 5.5, 2.5, 1.5], body_font_size=10)

# ============ 10 时间节点 ============
s = blank_slide()
add_header(s, "10 / 关键时间节点 & 待补事项")
add_table(s, 0.4, 1.1, 12.5, 4.5,
          ["时间", "事项", "责任人"],
          [
              ("5/21 12:00", "赞助商口播文案、对接人、物资清单回收；福布斯获奖名单 / 颁奖嘉宾确认；绿城晚宴邀请码机制确认", "胡继刚 / 王珏"),
              ("5/21 18:00", "圆桌桌号牌、圆桌大卡、席卡、胸卡 完成印刷", "F 组"),
              ("5/21 22:00", "T11-T50 晚宴桌名单最终确认", "王珏 / 黄欣 / 安淑娟"),
              ("5/22 08:00-10:00", "各赞助方物资陆续到场签收 / 上架", "D / G 组"),
              ("5/22 12:00", "★ 台卡、主席台陈列就位；福布斯颁奖物料必到", "B 组 / 张蒙"),
              ("5/22 13:00-13:30", "嘉宾签到与入场（凭邀请码核验）", "A 组"),
              ("5/22 13:30", "峰会开幕（姚志勇 致辞）", "—"),
              ("5/22 14:25-14:40", "★ 第四届 2026 人工智能商业化落地颁奖典礼（高光启幕）", "—"),
              ("5/22 18:30", "【绿城·潮鸣】星耀北外滩——AI 领袖定制晚宴 开始", "—"),
          ], col_widths=[2.5, 7.3, 2.7], body_font_size=10)

add_rect(s, 0.4, 5.8, 12.5, 1.3, LIGHT_BLUE)
add_textbox(s, 0.6, 5.85, 12.1, 0.4, "5/21 前需各方回填的\u201c待补\u201d事项（红色 = 急）", size=13, bold=True, color=NAVY)
add_textbox(s, 0.6, 6.25, 12.1, 0.85,
            "1) 绿城中国：到场 5 席名单、物资清单、对接人、邀请码核验机制、致辞嘉宾；   2) 福布斯中国 AI 评选委员会：获奖名单、颁奖嘉宾、颁奖物料（须 5/22 12:00 前到位）；   3) 蔚来汽车：展车 / 礼品、对接人；   4) 茶活力 / 美年大健康 / 古井贡酒：物资 / 对接人；   5) 缺号志愿者手机号（韦佳玉回收）。",
            size=10, color=RED)

# ============ 致谢 ============
s = blank_slide()
add_rect(s, 0, 0, 13.333, 7.5, NAVY)
add_textbox(s, 0.5, 2.6, 12.3, 1, "感谢", size=64, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
add_textbox(s, 0.5, 3.8, 12.3, 0.5,
            "期待 2026 年 5 月 22 日 · 上海北外滩 · 一滴水 现场相会",
            size=18, color=WHITE, align=PP_ALIGN.CENTER)
add_textbox(s, 0.5, 4.5, 12.3, 0.5,
            "北大经院上海校友会 · 复旦大学住房政策研究中心",
            size=14, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)
add_textbox(s, 0.5, 5.3, 12.3, 0.5,
            "组委会总联系：胡继刚 13262607888 · 王珏 13817642219",
            size=14, color=WHITE, align=PP_ALIGN.CENTER)

prs.save(OUT)
print(f"OK → {OUT}")
