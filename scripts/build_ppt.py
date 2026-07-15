"""Build a PowerPoint summarizing Speaker 1's investment views.

NOTE: Inside Chinese text content we use curly quotes 「 」 to avoid
clashing with Python's straight-quote string delimiters.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

OUT = '/workspace/outputs/03_说话人1_投资观点.pptx'

NAVY = RGBColor(0x0B, 0x1F, 0x3A)
GOLD = RGBColor(0xC9, 0xA2, 0x27)
LIGHT = RGBColor(0xF5, 0xF1, 0xE6)
DARK = RGBColor(0x22, 0x22, 0x22)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT = RGBColor(0xB0, 0x1F, 0x24)
GREEN = RGBColor(0x1F, 0x7A, 0x3D)
GRAY = RGBColor(0x6B, 0x6B, 0x6B)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def add_rect(slide, x, y, w, h, fill):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_text(slide, x, y, w, h, text, *, size=18, bold=False, color=DARK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font='Microsoft YaHei'):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0); tf.margin_right = Emu(0)
    tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    for i, line in enumerate(text.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return tb


def add_bullets(slide, x, y, w, h, items, *, size=16, color=DARK,
                bullet_color=GOLD, line_spacing=1.25,
                font='Microsoft YaHei', bullet_char='●'):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0); tf.margin_right = Emu(0)
    tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = line_spacing
        r1 = p.add_run()
        r1.text = bullet_char + '  '
        r1.font.name = font
        r1.font.size = Pt(size)
        r1.font.color.rgb = bullet_color
        r2 = p.add_run()
        r2.text = item
        r2.font.name = font
        r2.font.size = Pt(size)
        r2.font.color.rgb = color
    return tb


def header(slide, title, subtitle=None):
    add_rect(slide, 0, 0, SW, Inches(0.9), NAVY)
    add_rect(slide, 0, Inches(0.9), SW, Inches(0.06), GOLD)
    add_text(slide, Inches(0.5), Inches(0.15), Inches(12), Inches(0.6),
             title, size=26, bold=True, color=WHITE)
    if subtitle:
        add_text(slide, Inches(0.5), Inches(0.55), Inches(12), Inches(0.35),
                 subtitle, size=12, color=LIGHT)


def footer(slide, page):
    add_text(slide, Inches(0.5), Inches(7.1), Inches(8), Inches(0.3),
             '说话人1 投资观点 · 2026-05-17 投资领域交流探讨',
             size=10, color=GRAY)
    add_text(slide, Inches(12.0), Inches(7.1), Inches(1), Inches(0.3),
             f'{page}', size=10, color=GRAY, align=PP_ALIGN.RIGHT)


# === Slide 1: Cover ===
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, SW, SH, NAVY)
add_rect(s, 0, Inches(3.2), SW, Inches(0.04), GOLD)
add_rect(s, Inches(11.5), Inches(0.6), Inches(1.0), Inches(0.04), GOLD)
add_rect(s, Inches(0.8), Inches(6.6), Inches(1.0), Inches(0.04), GOLD)
add_text(s, Inches(0.8), Inches(1.6), Inches(11.7), Inches(0.8),
         '穿透赛道 · 把握溢价', size=22, color=GOLD)
add_text(s, Inches(0.8), Inches(2.2), Inches(11.7), Inches(1.2),
         '说话人1 投资观点纪要', size=48, color=WHITE, bold=True)
add_text(s, Inches(0.8), Inches(3.6), Inches(11.7), Inches(0.6),
         'Speaker 1 · Investment Insights Brief',
         size=16, color=LIGHT)
add_text(s, Inches(0.8), Inches(4.6), Inches(11.7), Inches(0.5),
         '关注股票 · 投资逻辑 · 方法论体系', size=20, color=WHITE)
add_text(s, Inches(0.8), Inches(5.2), Inches(11.7), Inches(0.4),
         '—— 来源:2026-05-17 投资领域交流探讨', size=14, color=LIGHT)
add_rect(s, Inches(0.8), Inches(6.3), Inches(2.6), Inches(0.5), GOLD)
add_text(s, Inches(0.8), Inches(6.3), Inches(2.6), Inches(0.5),
         'INTERNAL USE', size=14, color=NAVY, bold=True,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# === Slide 2: Agenda ===
s = prs.slides.add_slide(BLANK)
header(s, '目录 Agenda', '本份纪要的结构与导览')
items = [
    ('01', '宏观判断', '特朗普访华 · 美股 · A股 · 牛市判断'),
    ('02', '核心方法论', '赛道价值 · 四重溢价 · 四次穿透'),
    ('03', '主线 + 七大赛道', '算力 → 电力 · 新增创新药赛道'),
    ('04', '重点股票', '买入清单 + 回避清单'),
    ('05', '特斯拉四大赛道', '机器人 / 无人驾驶 / 太空 / 消灭银行'),
    ('06', '老登行业八大贬值领域', '房地产 / 知识 / 学历 / 软件 ...'),
    ('07', 'AI · 外资 · 香港', '智能体 · 代币市场 · 资本回流'),
    ('08', '操作纪要与金句', '可执行清单'),
]
y0 = 1.3
for i, (num, t, sub) in enumerate(items):
    row, col = i // 2, i % 2
    x = Inches(0.6 + col * 6.3)
    y = Inches(y0 + row * 1.3)
    add_rect(s, x, y, Inches(0.9), Inches(0.9), GOLD)
    add_text(s, x, y, Inches(0.9), Inches(0.9), num,
             size=24, bold=True, color=NAVY,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(1.05), y - Inches(0.02),
             Inches(5.2), Inches(0.5), t, size=20, bold=True, color=NAVY)
    add_text(s, x + Inches(1.05), y + Inches(0.45),
             Inches(5.2), Inches(0.45), sub, size=12, color=GRAY)
footer(s, 2)


# === Slide 3: Macro ===
s = prs.slides.add_slide(BLANK)
header(s, '01 · 宏观判断', '特朗普访华三大成果')
cards = [
    ('关税战基本结束', '进入震荡期\n已作为案例'),
    ('企业 > 国家外交', '新时代开启\n给中国带来巨大机会'),
    ('帮扶选举', '美联储换帅\nQE → QT,通胀高企'),
]
for i, (t, sub) in enumerate(cards):
    x = Inches(0.7 + i * 4.2); y = Inches(1.5)
    add_rect(s, x, y, Inches(3.9), Inches(2.2), NAVY)
    add_rect(s, x, y, Inches(0.15), Inches(2.2), GOLD)
    add_text(s, x + Inches(0.35), y + Inches(0.25),
             Inches(3.5), Inches(0.6), t, size=22, bold=True, color=WHITE)
    add_text(s, x + Inches(0.35), y + Inches(1.0),
             Inches(3.5), Inches(1.1), sub, size=15, color=LIGHT)

add_rect(s, Inches(0.7), Inches(4.1), Inches(12.0), Inches(2.7), LIGHT)
add_text(s, Inches(0.9), Inches(4.25), Inches(11.5), Inches(0.5),
         '对市场的传导', size=20, bold=True, color=NAVY)
bullets = [
    '美股进入 QT 周期:开始下跌,但跌幅小于上次,不会转向。',
    '周一调整传导到 A 股 → 第三次「倒车接人」机会(此前两次点位 3800)。',
    '关键判断:本轮不会跌破 4000,4000 = 再次起跳的起点。',
    '牛市未结束:「小牛拉货车」—— 小牛 = 美国科技三巨头。',
]
add_bullets(s, Inches(0.9), Inches(4.8), Inches(11.6), Inches(2.0),
            bullets, size=15)
footer(s, 3)


# === Slide 4: Methodology — four-tier premium ===
s = prs.slides.add_slide(BLANK)
header(s, '02 · 核心方法论', '反对传统巴菲特派 —— 价值 = 赛道价值')
add_rect(s, Inches(0.7), Inches(1.4), Inches(12.0), Inches(0.9), NAVY)
add_text(s, Inches(0.9), Inches(1.55), Inches(11.6), Inches(0.6),
         '在风险中寻找机会,在泡沫中兑现财富',
         size=22, bold=True, color=GOLD)
add_text(s, Inches(0.7), Inches(2.5), Inches(12), Inches(0.5),
         '四重溢价模型', size=20, bold=True, color=NAVY)
tiers = [
    ('①', '基本面价值', '底层钢筋,≤20', GRAY),
    ('②', '赛道价值', '20 → 30', GOLD),
    ('③', '成长/创新/领先/海外', '30 → 60(核心收益区)', ACCENT),
    ('④', '情绪价值', '60 → 100', GREEN),
]
for i, (num, title, sub, col) in enumerate(tiers):
    x = Inches(0.7 + i * 3.05); y = Inches(3.1)
    add_rect(s, x, y, Inches(2.9), Inches(1.7), col)
    add_text(s, x, y + Inches(0.1), Inches(2.9), Inches(0.5), num,
             size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, x, y + Inches(0.7), Inches(2.9), Inches(0.5), title,
             size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, x, y + Inches(1.15), Inches(2.9), Inches(0.5), sub,
             size=12, color=WHITE, align=PP_ALIGN.CENTER)
add_text(s, Inches(0.7), Inches(5.05), Inches(12), Inches(0.5),
         '持仓与节奏', size=20, bold=True, color=NAVY)
add_bullets(s, Inches(0.9), Inches(5.55), Inches(12), Inches(1.5), [
    '短线干不过量化,长线干不过国家队 → 只做中间 3–5 个月。',
    '5500 只 A 股只跟踪 200 只(一半以上为美国三巨头中国生态伙伴);历史涨幅 1× ~ 46×。',
    '含科率:25% 以下=垃圾;30% 第一里程碑;40% A 股七巨头登场。',
], size=14)
footer(s, 4)


# === Slide 5: Four Penetrations ===
s = prs.slides.add_slide(BLANK)
header(s, '02 · 核心方法论', '四次穿透:跟踪标准的逻辑链')
stages = [
    ('穿透股价', '看企业'),
    ('穿透产业', '看创新'),
    ('穿透创新', '看企业家'),
    ('穿透企业家', '看中国生态伙伴'),
]
for i, (t, sub) in enumerate(stages):
    x = Inches(0.6 + i * 3.05); y = Inches(1.8)
    add_rect(s, x, y, Inches(2.85), Inches(1.6), NAVY)
    add_text(s, x, y + Inches(0.25), Inches(2.85), Inches(0.5), t,
             size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_rect(s, x + Inches(0.5), y + Inches(0.95),
             Inches(1.85), Inches(0.04), GOLD)
    add_text(s, x, y + Inches(1.05), Inches(2.85), Inches(0.5), sub,
             size=14, color=LIGHT, align=PP_ALIGN.CENTER)
    if i < 3:
        arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                   x + Inches(2.85), y + Inches(0.6),
                                   Inches(0.18), Inches(0.4))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = GOLD
        arrow.line.fill.background()

add_rect(s, Inches(0.6), Inches(3.7), Inches(12.1), Inches(3.0), LIGHT)
add_text(s, Inches(0.8), Inches(3.85), Inches(11.7), Inches(0.5),
         '为什么这样穿透?', size=18, bold=True, color=NAVY)
add_bullets(s, Inches(0.9), Inches(4.4), Inches(11.6), Inches(2.4), [
    '企业家的「中国生态伙伴选择」—— 就是我们的跟踪标准。',
    '美国 7 巨头 → 已分化为「三辆车」(三家公司),只看这 3 家。',
    '200 只跟踪池一半以上来自三巨头供应链/生态。',
    '搭对生态:过去最少翻 1 倍,最多翻 46 倍。',
    '搭错生态(如华为)→ 平台压榨,反向被收割。',
], size=14)
footer(s, 5)


# === Slide 6: Seven Tracks ===
s = prs.slides.add_slide(BLANK)
header(s, '03 · 主线 + 七大赛道', '主线:算力 → 电力')
rows = [
    ('01', '算力芯片', '几连法', '听过这词才入门', ACCENT),
    ('02', '光芯片', '东光华(东山精密+光迅科技+华工科技);原:易中天', '电信号→光信号大趋势', NAVY),
    ('03', '可控核聚变', '尚西安', '未来能源', NAVY),
    ('04', '存储', '明德 / 阳光', '—', NAVY),
    ('05', '电力', '阳光电源', '电力设备更新最看好', NAVY),
    ('06', '新材料', '云南锗业 / 力量钻石 / 黄河旋风 / 玻璃/PCB', '天天出爆款', NAVY),
    ('07', '储能 · 固态电池', '宁德时代', '钠电池替代锂电池', NAVY),
]
y0 = 1.4
for i, (num, t, stk, note, col) in enumerate(rows):
    y = Inches(y0 + i * 0.78)
    add_rect(s, Inches(0.6), y, Inches(0.9), Inches(0.7), col)
    add_text(s, Inches(0.6), y, Inches(0.9), Inches(0.7), num,
             size=20, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, Inches(1.5), y, Inches(11.2), Inches(0.7), LIGHT)
    add_text(s, Inches(1.7), y + Inches(0.05), Inches(2.5), Inches(0.6),
             t, size=15, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(4.3), y + Inches(0.05), Inches(5.5), Inches(0.6),
             stk, size=13, color=DARK, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(9.9), y + Inches(0.05), Inches(2.8), Inches(0.6),
             note, size=12, color=GRAY, anchor=MSO_ANCHOR.MIDDLE)
footer(s, 6)


# === Slide 7: Innovative drugs ===
s = prs.slides.add_slide(BLANK)
header(s, '03 · 主线 + 七大赛道', '新增第 9 条赛道:创新药 · 3 项严苛限定')
add_rect(s, Inches(0.7), Inches(1.4), Inches(12.0), Inches(1.0), NAVY)
add_text(s, Inches(0.9), Inches(1.55), Inches(11.6), Inches(0.7),
         '整条赛道不「整体起」,只取符合 3 项条件的个股',
         size=20, bold=True, color=GOLD)
cards = [
    ('① 出口占比 ≥ 50%', '并且持续上升'),
    ('② 不受美国制裁', '美国对中国创新药专门「特别舒服」豁免'),
    ('③ 万物五巨头的中国生态伙伴', '美国医药五巨头亲自挑选'),
]
for i, (t, sub) in enumerate(cards):
    x = Inches(0.7 + i * 4.0); y = Inches(2.7)
    add_rect(s, x, y, Inches(3.7), Inches(2.3), LIGHT)
    add_rect(s, x, y, Inches(3.7), Inches(0.2), GOLD)
    add_text(s, x + Inches(0.25), y + Inches(0.35),
             Inches(3.3), Inches(0.8), t, size=17, bold=True, color=NAVY)
    add_text(s, x + Inches(0.25), y + Inches(1.2),
             Inches(3.3), Inches(1.0), sub, size=13, color=DARK)
add_rect(s, Inches(0.7), Inches(5.2), Inches(12.0), Inches(1.6), NAVY)
add_text(s, Inches(0.9), Inches(5.4), Inches(11.6), Inches(0.6),
         '目标池', size=18, bold=True, color=GOLD)
add_text(s, Inches(0.9), Inches(5.9), Inches(11.6), Inches(0.8),
         '全中国(含港股)仅 10 家公司符合上述 3 项条件。',
         size=16, color=WHITE)
footer(s, 7)


# === Slide 8: BUY list ===
s = prs.slides.add_slide(BLANK)
header(s, '04 · 重点股票', '买入清单 · 投资逻辑')
buys = [
    ('三安光电', '周一买入·勿等',
     '磷化铟芯片;8个万亿级市场领袖;产业缺口30%;中国仅8家;\n董事长/CEO被留置创造机会;福建国资委想抢。'),
    ('阳光电源', '重仓 · 目标170-230',
     '全球电力设备更新最看好;海外占比4%→全球18国;\n下月香港上市;昨涨13%;从200调到120是上车机会。'),
    ('宁德时代', '继续买 · 高位也买',
     '「消灭汽车的杀手」;钠电池6个月内取代锂电池;\n进入电力设备更新市场。'),
    ('云南锗业', '跌至90是买点',
     '锗=银化锌正极;新材料赛道资源类龙头;\n中国排名第一;现价突破102后回撤。'),
    ('黄河旋风/力量钻石', '新材料黑马',
     '河南人造钻石;意外发现是光通信液冷载体/光模块新材料。'),
    ('长飞光纤', '10年翻10倍',
     '英伟达美国建3厂;先进制造业回流标的;\n电信号→光信号大趋势。'),
]
for i, (name, tag, logic) in enumerate(buys):
    col, row = i % 2, i // 2
    x = Inches(0.55 + col * 6.35); y = Inches(1.4 + row * 1.92)
    add_rect(s, x, y, Inches(6.15), Inches(1.78), LIGHT)
    add_rect(s, x, y, Inches(0.18), Inches(1.78), GREEN)
    add_text(s, x + Inches(0.35), y + Inches(0.1),
             Inches(3.5), Inches(0.5), name, size=18, bold=True, color=NAVY)
    add_text(s, x + Inches(4.0), y + Inches(0.15),
             Inches(2.1), Inches(0.5), tag,
             size=11, bold=True, color=ACCENT, align=PP_ALIGN.RIGHT)
    add_text(s, x + Inches(0.35), y + Inches(0.65),
             Inches(5.7), Inches(1.1), logic, size=12, color=DARK)
footer(s, 8)


# === Slide 8.5: 三安光电 三重打击 deep-dive ===
s = prs.slides.add_slide(BLANK)
header(s, '04 · 重点股票 · 案例深读', '三安光电(600703):为什么在「三重打击」下还要买?')

add_rect(s, Inches(0.6), Inches(1.4), Inches(12.1), Inches(0.7), NAVY)
add_text(s, Inches(0.8), Inches(1.4), Inches(11.7), Inches(0.7),
         '近期股价低迷的根源 = 治理 + 股权 + 业绩 三重打击集中爆发',
         size=18, bold=True, color=GOLD, anchor=MSO_ANCHOR.MIDDLE)

# Three columns
hits = [
    ('🌪️ 第一重\n管理层动荡',
     ACCENT,
     '2026年3–4月\n实控人林秀成、副董事长林科闯\n相继被监委留置\n林科闯主导重庆碳化硅项目'),
    ('🔗 第二重\n股权危机·被动减持',
     ACCENT,
     '控股股东及一致行动人\n100%股份司法冻结\n4月中起15天内\n5972万股司法强卖\n套现约8.45亿元'),
    ('📉 第三重\n业绩首亏',
     ACCENT,
     '2025归母净利润 -3.53亿元\n(2006上市18年来首亏)\n2026 Q1 营收-32.59%\n归母净利润-68.15%\nLumileds收购被CFIUS否决'),
]
for i, (t, col, body) in enumerate(hits):
    x = Inches(0.6 + i * 4.05); y = Inches(2.25)
    add_rect(s, x, y, Inches(3.95), Inches(2.7), LIGHT)
    add_rect(s, x, y, Inches(3.95), Inches(0.7), col)
    add_text(s, x + Inches(0.2), y + Inches(0.05), Inches(3.7),
             Inches(0.7), t, size=14, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.25), y + Inches(0.85), Inches(3.6),
             Inches(1.85), body, size=11, color=DARK)

# Why still buy
add_rect(s, Inches(0.6), Inches(5.1), Inches(12.1), Inches(1.75), NAVY)
add_text(s, Inches(0.8), Inches(5.2), Inches(11.7), Inches(0.45),
         '为什么老师在「三重打击」下还要买?',
         size=15, bold=True, color=GOLD)
add_bullets(s, Inches(0.85), Inches(5.65), Inches(11.8), Inches(1.15), [
    '① 利空已充分定价 + 刚回撤,正是上车机会',
    '② 磷化铟卡光通信「咽喉材料」;后面8个万亿赛道(SiC/InP/Mini-Micro LED/滤波器…)都是领袖',
    '③ 治理崩塌 ≠ 资产基本面崩塌:福建国资委想抢这块「肥肉」→ 国资接盘预期',
    '④ 被动减持 = 出清信号;15天卖完后筹码集中度反而上升',
], size=12, color=WHITE)
footer(s, 9)


# === Slide 9: AVOID list ===
s = prs.slides.add_slide(BLANK)
header(s, '04 · 重点股票', '回避清单 · 看空逻辑')
avoid = [
    ('华为及关联平台', '中国最无耻的公司;压榨供应链(同芯片国内2000元/欧洲2000欧元);赛力斯卖一辆车给华为14万;机柜90%仍是电信号→被光模块时代淘汰;AI迭代落后阿里/腾讯。'),
    ('比亚迪', '事实上已破产;拖欠/压榨供应商,业内口碑差;钠电池替代锂电池后冲击最大。'),
    ('银行业(整体)', '「玩银行的时代彻底结束」;特斯拉版微信(年息6%)进一步消灭银行。'),
    ('传统汽车', '房地产之后下一个全线崩溃赛道;特斯拉无人驾驶/机器人化将摧毁传统车企。'),
    ('房地产', '「50年起不来」;万科类大窟窿太多;商铺99%必死无疑。'),
    ('国企(整体)', '三光:赔光/偷光/抢光;万科即深圳国资委窟窿;基本都是烂账。'),
    ('东山精密(谨慎)', '唯一不太确定项;华西村出身;财报水分大;卖不动产59亿、有600亿债。'),
    ('青龙行者', '15种品种中3种被淘汰,不要碰。'),
]
y0 = 1.35
for i, (name, logic) in enumerate(avoid):
    y = Inches(y0 + i * 0.66)
    add_rect(s, Inches(0.6), y, Inches(0.5), Inches(0.58), ACCENT)
    add_text(s, Inches(0.6), y, Inches(0.5), Inches(0.58), '✕',
             size=22, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, Inches(1.1), y, Inches(11.6), Inches(0.58), LIGHT)
    add_text(s, Inches(1.3), y + Inches(0.04), Inches(3.0), Inches(0.5),
             name, size=14, bold=True, color=NAVY,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(4.4), y + Inches(0.04), Inches(8.2), Inches(0.5),
             logic, size=11, color=DARK, anchor=MSO_ANCHOR.MIDDLE)
footer(s, 9)


# === Slide 10: Tesla ===
s = prs.slides.add_slide(BLANK)
header(s, '05 · 特斯拉四大赛道', '引领全球商业的「四辆车」')
cards = [
    ('机器人', 'Optimus 持续迭代', NAVY),
    ('无人驾驶', '即将全线铺开', NAVY),
    ('太空法律 / 太空经济', '马上开放;芯片厂量产 7nm', NAVY),
    ('消灭银行', '特斯拉版微信·年息 6%\n随时转换为任意股票', ACCENT),
]
for i, (t, sub, col) in enumerate(cards):
    x = Inches(0.6 + (i % 2) * 6.4)
    y = Inches(1.5 + (i // 2) * 2.3)
    add_rect(s, x, y, Inches(6.0), Inches(2.0), col)
    add_text(s, x + Inches(0.3), y + Inches(0.25),
             Inches(5.5), Inches(0.6), t, size=22, bold=True, color=GOLD)
    add_text(s, x + Inches(0.3), y + Inches(0.95),
             Inches(5.5), Inches(1.0), sub, size=15, color=WHITE)
add_rect(s, Inches(0.6), Inches(6.2), Inches(12.1), Inches(0.6), GOLD)
add_text(s, Inches(0.8), Inches(6.2), Inches(11.7), Inches(0.6),
         '未来同时是:全球最大金融公司 · 芯片公司 · 太空公司 · 机器人公司',
         size=15, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
footer(s, 10)


# === Slide 11: 8 depreciating sectors ===
s = prs.slides.add_slide(BLANK)
header(s, '06 · 「老登」行业', '八大资产/领域贬值清单')
items = [
    ('01', '房地产', '99% 必死,仅约 1% 可活;50 年起不来'),
    ('02', '商铺', '99% 必死无疑'),
    ('03', '知识', '中低档大模型 > 90% 大专院校教授'),
    ('04', '学历', '美国改招高中生;个案出现 1 亿美金薪酬'),
    ('05', '软件 / 码农', 'AI 出代码;五大行/地产软件开源后无价值'),
    ('06', '传统车企', '下一个全线崩溃的赛道'),
    ('07', '银行', '玩银行的时代彻底结束'),
    ('08', '国企(普遍)', '三光:赔光/偷光/抢光'),
]
y0 = 1.4
for i, (num, t, sub) in enumerate(items):
    col, row = i % 2, i // 2
    x = Inches(0.6 + col * 6.35); y = Inches(y0 + row * 1.35)
    add_rect(s, x, y, Inches(6.1), Inches(1.2), LIGHT)
    add_rect(s, x, y, Inches(0.9), Inches(1.2), ACCENT)
    add_text(s, x, y, Inches(0.9), Inches(1.2), num,
             size=26, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(1.1), y + Inches(0.1),
             Inches(4.9), Inches(0.5), t, size=17, bold=True, color=NAVY)
    add_text(s, x + Inches(1.1), y + Inches(0.6),
             Inches(4.9), Inches(0.6), sub, size=12, color=DARK)
footer(s, 11)


# === Slide 12: AI / foreign / HK ===
s = prs.slides.add_slide(BLANK)
header(s, '07 · AI · 外资 · 香港', '其他重要趋势')

add_rect(s, Inches(0.6), Inches(1.4), Inches(6.05), Inches(5.4), NAVY)
add_text(s, Inches(0.8), Inches(1.55), Inches(5.8), Inches(0.5),
         'AI 智能体', size=20, bold=True, color=GOLD)
add_bullets(s, Inches(0.85), Inches(2.15), Inches(5.7), Inches(4.5), [
    '「看不见的智能体」= 真正的爆发点(指节盖大小)',
    '看得见的机器人(酒店/产线)= 低价值',
    '看不见的智能体 = 看得见的 × 10',
    'AI 时代中文重新重要:提示词越精准,完成度越高',
    '可用模型:DeepSeek、Kimi 2.6、赤浦、Cursor、Cloud 4.7',
    '豆包基本不能用',
    '大芯片新趋势:美国公司直接做「盘子那么大」芯片挑战英伟达',
], size=13, color=WHITE)

add_rect(s, Inches(6.75), Inches(1.4), Inches(5.95), Inches(2.6), LIGHT)
add_text(s, Inches(6.95), Inches(1.5), Inches(5.7), Inches(0.5),
         '外资回流 · 四大主流', size=18, bold=True, color=NAVY)
add_bullets(s, Inches(7.0), Inches(2.05), Inches(5.6), Inches(2.0), [
    '美国:贝莱德(14万亿)·重仓宁德/阳光电源',
    '瑞士:欧洲钱通过瑞士进入',
    '中东:阿布扎比',
    '中国回流:含潘石屹等',
], size=13)

add_rect(s, Inches(6.75), Inches(4.15), Inches(5.95), Inches(2.65), LIGHT)
add_text(s, Inches(6.95), Inches(4.25), Inches(5.7), Inches(0.5),
         '香港 · 全球资本市场枢纽', size=18, bold=True, color=NAVY)
add_bullets(s, Inches(7.0), Inches(4.8), Inches(5.6), Inches(2.0), [
    '由「金融自由港」升格为「全球资本枢纽」',
    '下半年大量代币入场,代币市场爆发',
    '代币资产起点 = 6%',
    '马斯克代币基本一周抢完',
], size=13)
footer(s, 12)


# === Slide 13: action notes ===
s = prs.slides.add_slide(BLANK)
header(s, '08 · 操作纪要', '可执行金句 · 行动清单')
add_rect(s, Inches(0.6), Inches(1.4), Inches(12.1), Inches(0.8), GOLD)
add_text(s, Inches(0.8), Inches(1.4), Inches(11.7), Inches(0.8),
         '周一一只票:三安光电。不要问为什么,先上车。',
         size=22, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
quotes = [
    '现在站在光里 —— 只有三安光电还能站在光里。',
    '现在只买行业龙头。',
    '回撤即上车。',
    '便宜没好货:低市盈率股票多是垃圾或陷阱。',
    '在风险中寻找机会,在泡沫中兑现财富。',
    '短线干不过量化,长线干不过国家队,只做中间一段。',
    '创造性毁灭:关注「创造」,远离「毁灭」。',
    '穿透股价看企业,穿透产业看创新,穿透创新看企业家,穿透企业家看中国生态伙伴。',
    '5500只票只跟踪200只;一年内,3个月内不翻倍几乎不可能。',
    '能上链的就上链;不能上链的回避(传统行业未来没机会)。',
]
add_bullets(s, Inches(0.9), Inches(2.5), Inches(12.0), Inches(4.2),
            quotes, size=15)
footer(s, 13)


# === Slide 14: closing ===
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, SW, SH, NAVY)
add_rect(s, 0, Inches(3.4), SW, Inches(0.04), GOLD)
add_text(s, Inches(0.8), Inches(2.4), Inches(11.7), Inches(0.8),
         'THANK YOU', size=48, color=WHITE, bold=True)
add_text(s, Inches(0.8), Inches(3.6), Inches(11.7), Inches(0.5),
         '穿透赛道,把握溢价。', size=22, color=GOLD)
add_text(s, Inches(0.8), Inches(4.2), Inches(11.7), Inches(0.4),
         '—— 说话人1 投资观点纪要 · 2026-05-17', size=14, color=LIGHT)

prs.save(OUT)
print('Saved:', OUT)
