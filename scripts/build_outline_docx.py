# -*- coding: utf-8 -*-
"""
生成《反者道之动》详细内容大纲 Word 文档。

运行：python3 scripts/build_outline_docx.py
输出：dist/反者道之动_详细内容大纲.docx
"""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
DIST = ROOT / "dist"
DIST.mkdir(parents=True, exist_ok=True)
OUT = DIST / "反者道之动_详细内容大纲.docx"

INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x5C, 0x5C, 0x5C)
ACCENT = RGBColor(0x2C, 0x4A, 0x3E)
RULE = RGBColor(0x8A, 0x9A, 0x8E)


def set_run_font(run, size=11, bold=False, color=INK, font="宋体", italic=False):
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.name = font
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.get_or_add_rFonts()
    west = "Times New Roman" if font in ("宋体", "楷体", "黑体", "仿宋") else font
    r_fonts.set(qn("w:ascii"), west)
    r_fonts.set(qn("w:hAnsi"), west)
    r_fonts.set(qn("w:eastAsia"), font)


def add_page_number(paragraph):
    run1 = paragraph.add_run("— ")
    set_run_font(run1, size=9, color=MUTED, font="宋体")
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run = paragraph.add_run()
    set_run_font(run, size=9, color=MUTED, font="Times New Roman")
    r = run._r
    r.append(fld_begin)
    r.append(instr)
    r.append(fld_sep)
    t = OxmlElement("w:t")
    t.text = "1"
    r.append(t)
    r.append(fld_end)
    run2 = paragraph.add_run(" —")
    set_run_font(run2, size=9, color=MUTED, font="宋体")


def setup_page(doc):
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.4)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.6)
    section.right_margin = Cm(2.4)
    section.different_first_page_header_footer = True

    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = hp.add_run("反者道之动  ·  详细内容大纲")
    set_run_font(run, size=9, color=ACCENT, font="宋体")

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_page_number(fp)


def p(doc, text, *, size=11, bold=False, color=INK, font="宋体",
      align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=8,
      first_line=True, italic=False, line_spacing=1.35):
    para = doc.add_paragraph()
    para.alignment = align
    pf = para.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line_spacing
    if first_line and align == WD_ALIGN_PARAGRAPH.LEFT:
        pf.first_line_indent = Cm(0.74)
    run = para.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color, font=font, italic=italic)
    return para


def h1(doc, text):
    p(doc, text, size=18, bold=True, color=ACCENT, font="黑体",
      align=WD_ALIGN_PARAGRAPH.CENTER, space_before=18, space_after=12,
      first_line=False, line_spacing=1.2)


def h2(doc, text):
    p(doc, text, size=14, bold=True, color=ACCENT, font="黑体",
      space_before=16, space_after=8, first_line=False, line_spacing=1.2)


def h3(doc, text):
    p(doc, text, size=12, bold=True, color=INK, font="黑体",
      space_before=12, space_after=6, first_line=False, line_spacing=1.2)


def h4(doc, text):
    p(doc, text, size=11, bold=True, color=ACCENT, font="楷体",
      space_before=8, space_after=4, first_line=False)


def bullet(doc, text, level=0):
    para = doc.add_paragraph(style="List Bullet")
    para.clear()
    pf = para.paragraph_format
    pf.space_before = Pt(1)
    pf.space_after = Pt(3)
    pf.line_spacing = 1.3
    pf.left_indent = Cm(0.5 + level * 0.5)
    run = para.add_run(text)
    set_run_font(run, size=10.5, color=INK, font="宋体")


def quote(doc, text):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = para.paragraph_format
    pf.space_before = Pt(8)
    pf.space_after = Pt(8)
    pf.left_indent = Cm(1.0)
    pf.right_indent = Cm(1.0)
    run = para.add_run(text)
    set_run_font(run, size=11, color=ACCENT, font="楷体", italic=False)


def divider(doc):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = para.paragraph_format
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    run = para.add_run("·  ·  ·")
    set_run_font(run, size=10, color=RULE, font="宋体")


def chapter_block(doc, title, zhangzhi, arguments, dao=None, antifragile=None,
                  examples=None, tip=None, extra_bullets=None):
    h3(doc, title)
    h4(doc, "章旨")
    p(doc, zhangzhi, size=10.5, space_after=6)
    h4(doc, "核心论点")
    if isinstance(arguments, list):
        for item in arguments:
            bullet(doc, item)
    else:
        p(doc, arguments, size=10.5, space_after=6)
    if dao:
        h4(doc, "《道德经》对接")
        if isinstance(dao, list):
            for item in dao:
                bullet(doc, item)
        else:
            p(doc, dao, size=10.5, space_after=6)
    if antifragile:
        h4(doc, "《反脆弱》对接")
        if isinstance(antifragile, list):
            for item in antifragile:
                bullet(doc, item)
        else:
            p(doc, antifragile, size=10.5, space_after=6)
    if examples:
        h4(doc, "例证 / 图像")
        for item in examples:
            bullet(doc, item)
    if tip:
        h4(doc, "短章写法提示")
        p(doc, tip, size=10.5, space_after=6, color=MUTED)
    if extra_bullets:
        for heading, items in extra_bullets:
            h4(doc, heading)
            for item in items:
                bullet(doc, item)


def build():
    doc = Document()
    setup_page(doc)

    # —— 封面 ——
    for _ in range(3):
        p(doc, "", space_after=0, first_line=False)
    p(doc, "书稿策划大纲", size=12, color=MUTED, font="楷体",
      align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, space_after=18)
    p(doc, "反者道之动", size=28, bold=True, color=ACCENT, font="黑体",
      align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, space_after=10, line_spacing=1.1)
    p(doc, "逆向、波动与柔弱之用", size=14, color=INK, font="楷体",
      align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, space_after=6)
    p(doc, "——从《道德经》读《反脆弱》", size=12, color=MUTED, font="宋体",
      align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, space_after=28)
    quote(doc, "反者道之动，弱者道之用。")
    p(doc, "《道德经》第四十章", size=10, color=MUTED, font="宋体",
      align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, space_after=36)
    p(doc, "体例参照《反脆弱》：分卷 · 多章 · 短小精悍", size=11, color=MUTED, font="宋体",
      align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False)
    p(doc, "七卷三十六章 + 前言后记与附录", size=11, color=MUTED, font="宋体",
      align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, space_after=24)
    doc.add_page_break()

    # —— 定位 ——
    h1(doc, "〇、全书定位")
    h2(doc, "一句话")
    p(doc, "这不是注经，也不是《反脆弱》的中文缩写本，而是用「反」与「弱」两条线索，"
      "把老子的逆向动力学与塔勒布的反脆弱思想织成一本可实践的短章哲思书。")

    h2(doc, "读者")
    bullet(doc, "读过《反脆弱》或对「不确定性」敏感的人")
    bullet(doc, "对《道德经》有兴趣但不想读成鸡汤或玄学的人")
    bullet(doc, "创业者、投资人、管理者、研究者——凡是必须在波动中决策的人")

    h2(doc, "写法原则（刻意模仿《反脆弱》）")
    bullet(doc, "章多而短：每章通常 1500–3500 字，一章一个锋利的想法。")
    bullet(doc, "先图像，后概念：隐喻优先于定义。")
    bullet(doc, "经文作锚，不作文抄：每章最多引用一两句，立刻落到现代情境。")
    bullet(doc, "拒绝预测崇拜：反对「精确预测 → 精密控制」的现代幻觉。")
    bullet(doc, "有皮肉的观点：敢下判断，敢得罪「脆弱化专家」。")
    bullet(doc, "卷末留「可带走的句子」：每卷结尾给 3–5 条可记诵的原则。")

    h2(doc, "《道德经》×《反脆弱》对应骨架")
    pairs = [
        ("反者道之动", "反脆弱：从波动中获益", "逆向是系统的引擎"),
        ("弱者道之用", "选择权、冗余、杠铃", "柔弱是可用的力量"),
        ("为道日损", "Via Negativa（否定法）", "少做错，比多做对更重要"),
        ("无为而无不为", "非预测式行动 / 试错", "设计可上可下的结构"),
        ("飘风不终朝", "黑天鹅、尾部风险", "极端事件比均值更重要"),
        ("知止不殆", "皮肤在局", "有代价的判断才可信"),
    ]
    for a, b, c in pairs:
        bullet(doc, f"{a}  ↔  {b}  →  {c}")

    h2(doc, "建议成书规模")
    bullet(doc, "七卷 + 前言 + 后记 + 附录")
    bullet(doc, "正文约 36 章（不含前言、后记）")
    bullet(doc, "全稿目标约 10–14 万字")
    divider(doc)

    # —— 目录 ——
    h1(doc, "全书目录总览")
    toc = [
        "前言　为什么「反」比「正」更接近真相",
        "",
        "第一卷　道动于反——何谓反脆弱的古典说法",
        "　1. 达摩克利斯、九头蛇与水",
        "　2. 「反」不是「反对」，是回返与反弹",
        "　3. 三态：脆弱、强韧、反者",
        "　4. 过补偿：受伤之后反而更强",
        "　5. 压力是信息，无菌是毒药",
        "　6. 我的脆弱，可能是别人的反脆弱",
        "",
        "第二卷　现代性为何怕「反」——fragilista 的时代",
        "　7. 市集与写字楼",
        "　8. 我爱（某些）随机性",
        "　9. 天真干预：灭火的人常是纵火者",
        "　10. 均值的暴政与尾部的真相",
        "　11. 稳定崇拜如何制造更大的崩溃",
        "",
        "第三卷　不预测，而备势——非预测的世界观",
        "　12. 预测是现代性的孩子",
        "　13. 胖子托尼与脆弱专家",
        "　14. 塞涅卡的上边与下边",
        "　15. 知不知，尚矣：无知的正确用法",
        "　16. 备势，而不是备答案",
        "",
        "第四卷　弱者道之用——选择权、冗余与柔韧",
        "　17. 别嫁给摇滚明星",
        "　18. 泰勒斯的葡萄：选择权比知识更古老",
        "　19. 先有走，才有飞：实践的智慧顺序",
        "　20. 两件看起来一样的事，往往不是一回事",
        "　21. 杠铃：极端的两端，避开致命的中间",
        "　22. 柔弱处上：redundancy 是美德",
        "",
        "第五卷　非线性——大石头与一千颗小石子",
        "　23. 大石头与一千颗小石子",
        "　24. 点金石与它的逆运算",
        "　25. 时间会暴露脆弱",
        "　26. 过犹不及：剂量决定一切",
        "　27. 凸性与凹性：你站在哪一边",
        "",
        "第六卷　为道日损——减法比加法更接近道",
        "　28. 否定法：先问「不做什么」",
        "　29. 长寿的艺术：少碰黑天鹅",
        "　30. 医学、凸性与无知",
        "　31. 活得久，但不要活得太满",
        "　32. 损之又损，以至于无为",
        "",
        "第七卷　有身者责——反脆弱的伦理",
        "　33. 皮肤在局：没有代价的建议是噪音",
        "　34. 信念的代价",
        "　35. 不对称的道德",
        "　36. 知止：边界即智慧",
        "",
        "后记　反者道之动，弱者道之用",
        "附录 A　关键概念对照表",
        "附录 B　可实践的「反—弱」清单",
        "附录 C　延伸阅读与争议说明",
    ]
    for line in toc:
        if not line:
            p(doc, "", space_after=2, first_line=False)
            continue
        p(doc, line, size=10.5, space_after=2, first_line=False, line_spacing=1.25,
          font="黑体" if not line.startswith("　") and line[0] in "前后附第" or line.startswith("第") or line.startswith("前言") or line.startswith("后记") or line.startswith("附录") else "宋体",
          bold=not line.startswith("　"))

    doc.add_page_break()

    # —— 前言 ——
    h1(doc, "前言　为什么「反」比「正」更接近真相")
    h4(doc, "章旨")
    p(doc, "说明写作动机：读者感到《反脆弱》与「反者道之动」在讲同一件事，却缺少一本把二者真正接上的书。")
    h4(doc, "要写的内容")
    bullet(doc, "开篇故事：一次「按计划本该顺利、却因意外反而更好」的经历。")
    bullet(doc, "指出现代人的默认语法是「正」：目标—计划—执行—控制；老子与塔勒布共同提醒：真正驱动系统的，常常是「反」。")
    bullet(doc, "澄清本书不是注经、不是读书笔记、不是「越挫越勇」成功学。")
    bullet(doc, "给出阅读地图：七卷各自回答一个问题。")
    bullet(doc, "约定术语：「反者」= 从波动中获益；「弱用」= 以柔、余、选项、减法发挥作用。")
    h4(doc, "收束句")
    quote(doc, "道不直行。能从逆向中获益的，才接近道的用法。")

    # —— 第一卷 ——
    h1(doc, "第一卷　道动于反")
    p(doc, "本卷任务：建立全书基础词汇。读完六章后，能用「脆弱 / 强韧 / 反者」三态看世界，"
      "并把《道德经》的「反」从文学修辞变成可操作的系统概念。",
      size=10.5, color=MUTED, first_line=False)

    chapter_block(
        doc, "第 1 章　达摩克利斯、九头蛇与水",
        "用三个古典图像定义三种对随机性的关系。",
        [
            "达摩克利斯：头顶悬剑——享受好处却暴露于毁灭性下行，是脆弱。",
            "九头蛇：砍一头长两头——从伤害中增益，是反者。",
            "水：至柔驰骋天下之至坚——不硬抗，而改形、迂回、积累，是弱用的原型。",
        ],
        dao=["天下莫柔弱于水，而攻坚强者莫之能胜。", "反者道之动——动不是直线推进，而是回旋、反弹、反向生成。"],
        antifragile="开篇三隐喻的中国化改写：把「反脆弱」从外来词变成可感的古典图像。",
        tip="全章几乎不定义术语，只讲三个故事图像；章末用极简对照表收束。",
    )

    chapter_block(
        doc, "第 2 章　「反」不是「反对」，是回返与反弹",
        "澄清汉语里「反」的多重含义，避免把全书读成「唱反调」。",
        [
            "回返：物极则返。",
            "相反相成：有无相生。",
            "反弹：压之愈深，起之愈猛（过补偿）。",
            "逆向获利：别人当噪音的波动，你当养分。",
            "现代人常只听到「对着干」；本书取以上四层。",
        ],
        dao=["有无相生，难易相成。", "将欲歙之，必固张之。"],
        examples=["市场恐慌抛售制造下一轮便宜筹码", "训练性微损伤带来超量恢复", "可恢复的小失败强化纠错能力"],
        tip="像词典词条一样写，干净、锋利，最后落到「动」字。",
    )

    chapter_block(
        doc, "第 3 章　三态：脆弱、强韧、反者",
        "给出全书最重要的分类装置。",
        [
            "脆弱：面对波动受损（薄冰、绷紧的弦）。",
            "强韧：面对波动不受影响（岩石）——只是中间态。",
            "反者：面对波动获益（水、草木、九头蛇）。",
            "真正可欲的是：在可承受波动中获益，同时隔离毁灭性风险。",
        ],
        dao=["坚强者死之徒，柔弱者生之徒——歌颂可变形的生命结构，而非软弱。"],
        antifragile="fragile / robust / antifragile 三分法。",
        tip="用玻璃杯、塑料杯、免疫系统举例；避免公式。",
    )

    chapter_block(
        doc, "第 4 章　过补偿：受伤之后反而更强",
        "把「反」落到生理学与组织行为的具体机制。",
        [
            "系统受压后若能过量修复，就会比受压前更强。",
            "条件：压力可吸收、允许恢复、不把所有波动熨平。",
            "过载会杀死反者——这是本章的必要警告。",
        ],
        dao=["飘风不终朝，骤雨不终日。", "弱者道之用——可弯才能吸收并转化冲击。"],
        examples=["创业团队经历可控危机后流程更清晰", "过度保护的孩子缺少过补偿机会"],
        tip="写得像科学随笔：一个机制，三个例子，一个警告。",
    )

    chapter_block(
        doc, "第 5 章　压力是信息，无菌是毒药",
        "攻击「消灭一切波动」的现代洁癖。",
        "波动、噪声、错误、小失败，是系统获取信息与保持活性的方式。彻底无菌的环境培养出超级脆弱的生命。",
        dao=["天地不仁，以万物为刍狗——拒绝过度拟人化的照料。", "多言数穷，不如守中。"],
        antifragile="急性压力 vs 慢性压力；hormesis（毒物兴奋效应）。",
        examples=["消灭街头「混乱」也消灭创新的意外相遇", "消灭异议也消灭纠错通道"],
    )

    chapter_block(
        doc, "第 6 章　我的脆弱，可能是别人的反脆弱",
        "引入系统视角：脆弱性会外溢与转移。",
        "一个人/机构通过把波动转嫁出去而「显得」稳健，整体系统却更脆弱。银行、大企业、官僚体系常如此。",
        dao=["天之道，损有余而补不足；人之道则不然，损不足以奉有余。"],
        antifragile="为第七卷伦理埋线：What kills me makes others stronger。",
        extra_bullets=[("本卷可带走的句子", [
            "反，是回返，也是反弹，更是逆向获益。",
            "强韧只是不死；反者是因乱而进。",
            "没有压力的系统，正在失去未来。",
            "先分清：你消灭的是风险，还是信息。",
        ])],
    )

    # —— 第二卷 ——
    h1(doc, "第二卷　现代性为何怕「反」")
    p(doc, "本卷任务：诊断。为什么当代制度、企业与个人生活系统性地消灭「反」与「弱」，却制造更大的脆断。",
      size=10.5, color=MUTED, first_line=False)

    chapter_block(
        doc, "第 7 章　市集与写字楼",
        "用两种空间隐喻对比有机秩序与设计秩序。",
        [
            "市集：局部试错、冗余、可替代——偏反者。",
            "写字楼/中央计划：效率高、耦合紧——偏脆弱。",
            "现代性常把市集改造成写字楼，再把崩溃称为「黑天鹅」。",
        ],
        dao=["治大国若烹小鲜。", "其政闷闷，其民淳淳；其政察察，其民缺缺。"],
        examples=["供应链「刚好及时」vs 适度库存", "算法把多样性压成单一最优解"],
    )

    chapter_block(
        doc, "第 8 章　我爱（某些）随机性",
        "区分有益随机与致命随机。",
        "热爱小而频繁、可学习的随机；憎恨大而稀有、可毁灭的随机被制度隐藏。fragilista 的错误是：消灭前者，放大后者。",
        dao=["道常无为而无不为——不为强行整形，故能成全万物自化。"],
        tip="标题带一点挑衅；正文给出「爱 / 不爱」两份清单。",
    )

    chapter_block(
        doc, "第 9 章　天真干预：灭火的人常是纵火者",
        "定义天真干预：用不理解系统的短期可见好处，换取长期隐性脆弱。",
        "干预本身不是恶；无皮肤、无反馈、无尾部意识的干预才是恶。许多「维稳」「平滑」「优化」属于此类。",
        dao=["天下多忌讳，而民弥贫……法令滋彰，盗贼多有。", "为之者败之，执之者失之。"],
        examples=["过度医疗与过度教养", "用补贴掩盖错误价格信号", "用债务推迟清算"],
    )

    chapter_block(
        doc, "第 10 章　均值的暴政与尾部的真相",
        "说明为什么「平均而言」常常是最危险的句子。",
        "在非线性与厚尾世界里，决定命运的是极值而非均值。以均值管理风险，等于用天气日报管理地震。",
        dao=["飘风不终朝——非常态事件主导叙事。", "勇于敢则杀，勇于不敢则活。"],
        antifragile=["中庸被误用", "标准差崇拜", "「没有证据表明有风险」≠「有证据表明没风险」"],
    )

    chapter_block(
        doc, "第 11 章　稳定崇拜如何制造更大的崩溃",
        "本卷结论：稳定不是安全，被压抑的波动会以更大振幅回来。",
        [
            "森林灭掉所有小火 → 积累成大火。",
            "市场消灭所有小波动 → 积累成崩盘。",
            "组织消灭所有小冲突 → 积累成政治爆炸。",
            "你越按住「反」，「动」越以极端形式爆发。",
        ],
        extra_bullets=[("本卷可带走的句子", [
            "效率经常是脆弱的别名。",
            "小乱是学费，大乱是破产。",
            "先问干预消灭了什么信息。",
            "怕随机的人，终将被更大的随机教育。",
        ])],
    )

    # —— 第三卷 ——
    h1(doc, "第三卷　不预测，而备势")
    p(doc, "本卷任务：换世界观。从「先知道未来再行动」转向「在不知道未来时仍能因结构而获益」。",
      size=10.5, color=MUTED, first_line=False)

    chapter_block(
        doc, "第 12 章　预测是现代性的孩子",
        "拆穿预测工业的心理与制度根源。",
        "预测满足焦虑与绩效叙事，不一定满足生存。复杂系统的点预测常像占卜；真正重要的是对预测错误的暴露度。",
        dao=["前识者，道之华，而愚之始。", "知不知，尚矣；不知知，病也。"],
    )

    chapter_block(
        doc, "第 13 章　胖子托尼与脆弱专家",
        "用人物对照写出两种知识态度。",
        [
            "脆弱专家：模型精美、无下行暴露、错了仍升职。",
            "实践者：不求解释世界，只求在错误中不被毁，并偶尔捡便宜。",
            "中国语境可改写为「PPT 战略家」vs「带仓位的老手」。",
        ],
        dao=["信言不美，美言不信。", "善者不辩，辩者不善。"],
    )

    chapter_block(
        doc, "第 14 章　塞涅卡的上边与下边",
        "引入不对称：好事的上边与坏事的下边如何安排。",
        "斯多葛不是忍辱，而是主动管理最坏情况，从而自由面对上行。先能守柔，才有资格言进取。",
        dao=["知足不辱，知止不殆。", "不敢为天下先——不把所有筹码押在「先手正确」上。"],
        antifragile="Seneca’s barbell；stoicism as antifragility technology。",
    )

    chapter_block(
        doc, "第 15 章　知不知，尚矣：无知的正确用法",
        "把「无知」从缺陷改写为资源管理问题。",
        [
            "承认不可知域。",
            "把决策做成对无知不敏感（或能从无知中受益）。",
            "警惕「假装知道」。",
        ],
        dao="整章围绕「知不知，尚矣」展开，作为第三卷的经眼。",
    )

    chapter_block(
        doc, "第 16 章　备势，而不是备答案",
        "提出第三卷的实践口号。",
        [
            "备答案：依赖预测与计划。",
            "备势：储备冗余、选项、现金、信任、健康、可逆性、小规模试错权。",
            "势接近道的可用状态：不执某一形，而保持可变之形。",
        ],
        extra_bullets=[("本卷可带走的句子", [
            "预测错不可怕，暴露错才可怕。",
            "专家的流畅，常常是脆弱的化妆。",
            "先安排下边，再谈上边。",
            "备势者生，备答案者脆。",
        ])],
    )

    # —— 第四卷 ——
    h1(doc, "第四卷　弱者道之用")
    p(doc, "本卷任务：正面建设。说明「弱」不是无能，而是选项、冗余、可退出、可变形——道真正「用」得上的那一层。",
      size=10.5, color=MUTED, first_line=False)

    chapter_block(
        doc, "第 17 章　别嫁给摇滚明星",
        "用情感与事业选择讲清路径锁定如何制造脆弱。",
        "绑定单一叙事、客户、技术、身份，等于自愿取消选择权。反者需要可退出、可转移、可并列的结构。",
        dao=["持而盈之，不如其已。", "功成身退，天之道。"],
    )

    chapter_block(
        doc, "第 18 章　泰勒斯的葡萄：选择权比知识更古老",
        "说明 optionality：权利与义务不对称。",
        "泰勒斯租下榨油机：上行敞开，下行有限。真正的智慧有时不是更准的预测，而是买到不对称。弱（不满仓、不一次押死）提供了使用选择权的空间。",
        dao=["无为而无不为——不做把选择权做死的事。", "夫唯不争，故天下莫能与之争。"],
    )

    chapter_block(
        doc, "第 19 章　先有走，才有飞：实践的智慧顺序",
        "批判「理论先行、实践随后」的现代迷信。",
        "许多技术与制度是试错涌现，后被理论追认。组织应保护「可飞的试错」，而不是先写完备说明书。",
        dao=["道可道，非常道。", "为学日益，为道日损——为第六卷埋伏笔。"],
    )

    chapter_block(
        doc, "第 20 章　两件看起来一样的事，往往不是一回事",
        "训练辨别力：表象相似，凸性不同。",
        "同样叫投资、改革、学习、勤奋，有的有反者结构，有的是脆弱结构。名称相同，命运相反。",
        examples=["分散试错 vs 伪多元化", "现金缓冲 vs 账面利润", "真正恢复 vs 假装勤奋的空转"],
    )

    chapter_block(
        doc, "第 21 章　杠铃：极端的两端，避开致命的中间",
        "给出最实用的结构工具之一。",
        "一端极度保守（生存），一端极度开放（小成本试错）；避开「中等风险却致命」的中间地带。用弱托住生存，用小而多的「反」捕捉上行。",
        dao=["知其雄，守其雌。", "曲则全，枉则直。"],
    )

    chapter_block(
        doc, "第 22 章　柔弱处上：redundancy 是美德",
        "为「浪费」「重复」「低效」平反。",
        "自然系统充满冗余。效率优化常剃掉冗余，等于剃掉反脆弱。",
        dao=["上善若水。", "柔弱胜刚强。"],
        extra_bullets=[("本卷可带走的句子", [
            "弱，是保留变形的权利。",
            "选择权是古老的智慧技术。",
            "理论常常迟到，试错常常先到。",
            "杠铃：先不死，再谈跃迁。",
            "冗余不是浪费，是活着的手续费。",
        ])],
    )

    # —— 第五卷 ——
    h1(doc, "第五卷　非线性")
    p(doc, "本卷任务：机制层解释——世界是非线性的，剂量、规模与时间会改变性质。",
      size=10.5, color=MUTED, first_line=False)

    chapter_block(
        doc, "第 23 章　大石头与一千颗小石子",
        "用直觉实验讲清规模与伤害的非线性。",
        "同样重量，一颗大石与一千颗小石效果完全不同。脆弱者害怕集中冲击；反者喜欢把冲击切碎。",
        dao=["合抱之木，生于毫末；九层之台，起于累土。", "图难于其易，为大于其细。"],
    )

    chapter_block(
        doc, "第 24 章　点金石与它的逆运算",
        "讨论转化：如何把波动点成收益，以及如何把收益点成脆弱。",
        "反者结构像点金石；脆弱化政策像逆点金术。关键不在愿望，而在暴露的凸性。",
        dao=["祸兮福之所倚，福兮祸之所伏。"],
    )

    chapter_block(
        doc, "第 25 章　时间会暴露脆弱",
        "引入时间作为终极检验。",
        "尚未破产不是证明稳健；只是样本路径还没走到尾部。能穿越长时间的，多半内置了反者或至少强韧。",
        dao=["天长地久……以其不自生，故能长生。", "企者不立，跨者不行。"],
    )

    chapter_block(
        doc, "第 26 章　过犹不及：剂量决定一切",
        "把 hormesis 与中庸从口号变成剂量学。",
        "同一因子：小剂量可能有益，大剂量致命。运动、压力、杠杆、监管皆然。「中」不是平均，而是落在有益剂量区。",
        dao=["物壮则老，是谓不道，不道早已。", "甚爱必大费，多藏必厚亡。"],
    )

    chapter_block(
        doc, "第 27 章　凸性与凹性：你站在哪一边",
        "用不数学的方式讲清：波动对你是朋友还是敌人。",
        [
            "凸性（反者）：波动增大，期望变好。",
            "凹性（脆弱）：波动增大，期望变差。",
            "人生、投资、职业、健康的核心问题：你的结构是凸的还是凹的？",
        ],
        dao=["反者道之动——道站在能从逆向波动中生成新秩序的一边。"],
        extra_bullets=[("本卷可带走的句子", [
            "切碎冲击，不要合并冲击。",
            "福祸转化，发生在曲线上，不发生在口号里。",
            "时间是最严厉的风控。",
            "先看剂量，再谈好坏。",
            "站到凸的一边。",
        ])],
    )

    # —— 第六卷 ——
    h1(doc, "第六卷　为道日损")
    p(doc, "本卷任务：行动方法论核心——Via Negativa。少做错，比多做对更重要。",
      size=10.5, color=MUTED, first_line=False)

    chapter_block(
        doc, "第 28 章　否定法：先问「不做什么」",
        "确立第六卷总原则。",
        "加法满足表演欲；减法提高存活率。战略、健康、关系、产品，优先删除致命因素，再谈优化。",
        dao=["为学日益，为道日损。损之又损，以至于无为。——本卷经眼。"],
        antifragile="Via Negativa 整卷精神。",
    )

    chapter_block(
        doc, "第 29 章　长寿的艺术：少碰黑天鹅",
        "谈生存策略：长寿往往来自避开毁灭，而非追求刺激均值。",
        "活得长的系统，秘密通常是「没做哪些事」。个人财务、企业战略、国家治理同构。",
        dao=["知止不殆，可以长久。", "减少致命敞口，而非封闭主义。"],
    )

    chapter_block(
        doc, "第 30 章　医学、凸性与无知",
        "用医疗决策作减法方法论的尖锐案例。",
        "在因果不透明领域，积极干预常具凹性：好处可见、伤害滞后且厚尾。真正需要动手时，要有巨大的不对称证据。",
        dao=["圣人不病，以其病病。", "治人、事天，莫若啬。"],
        tip="避免伪科学断言；强调决策结构与谦卑，而非具体处方。",
    )

    chapter_block(
        doc, "第 31 章　活得久，但不要活得太满",
        "讨论「满」与「留白」。",
        "日程、杠杆、社交、野心、资产都可以太满。满则失去吸收冲击的空间。留白是反脆弱的预算。",
        dao=["保此道者，不欲盈。", "大成若缺，其用不弊。"],
    )

    chapter_block(
        doc, "第 32 章　损之又损，以至于无为",
        "把「无为」从消极误读中解放出来。",
        "无为 = 不做脆弱化之事 + 不阻断自组织 + 让选择权与冗余工作。无不为 = 结构对了，结果自然涌现。这是两书在行动哲学上的最高会合点。",
        extra_bullets=[("本卷可带走的句子", [
            "先减法，后加法。",
            "长寿是避开毁灭的艺术。",
            "在无知领域，少动手常常更智慧。",
            "留白是战略资源。",
            "无为是结构正确后的不多事。",
        ])],
    )

    # —— 第七卷 ——
    h1(doc, "第七卷　有身者责")
    p(doc, "本卷任务：补上价值论。没有皮肤在局的「反脆弱」会沦为剥削他人脆弱的话术。",
      size=10.5, color=MUTED, first_line=False)

    chapter_block(
        doc, "第 33 章　皮肤在局：没有代价的建议是噪音",
        "建立伦理第一原则。",
        "谁承担错误的代价，谁的判断才值得听。把下行转嫁给他人、自己只吃上行的人，是系统性脆弱的制造者。",
        dao=["圣人无积……既以为人己愈有。", "受国之垢，是谓社稷主。"],
        antifragile="Skin in the Game：伦理是反脆弱的完成式。",
    )

    chapter_block(
        doc, "第 34 章　信念的代价",
        "谈真诚：你是否愿意为判断付代价。",
        "口头信念廉价；带仓位、带声誉、带时间的信念昂贵。社会进步需要提高「空口正确」的成本。",
        dao=["轻诺必寡信。", "信不足焉，有不信焉。"],
    )

    chapter_block(
        doc, "第 35 章　不对称的道德",
        "批判「我错了由你买单」的制度道德。",
        "现代大型组织容易制造：决策者反脆弱，被决策者脆弱。文明标准是禁止或至少暴露这种不对称。",
        dao=["天之道，损有余而补不足。", "民不畏威，则大威至。"],
    )

    chapter_block(
        doc, "第 36 章　知止：边界即智慧",
        "全书最后一章正文：把「止」作为反者的成熟形态。",
        "反者不等于无限扩张、无限波动成瘾。知道何处停、何处不碰、何处保持弱、何处允许反，才是「弱者道之用」的完成。",
        dao=["知止不殆。", "深根固柢，长生久视之道。"],
        extra_bullets=[("本卷可带走的句子", [
            "无代价的聪明，是危险的聪明。",
            "信念要有皮肤。",
            "不对称的上行是掠夺，不是智慧。",
            "止，是反与弱的最终纪律。",
        ])],
    )

    # —— 后记 ——
    h1(doc, "后记　反者道之动，弱者道之用")
    bullet(doc, "反者道之动：让波动成为引擎，而不是只当敌人。")
    bullet(doc, "弱者道之用：用选择权、冗余、减法与可变形性去用道。")
    bullet(doc, "回应开篇：你不必成为预言家；你需要成为结构正确的人。")
    bullet(doc, "给出「明天就能用」的五条清单（与附录 B 呼应）。")
    bullet(doc, "承认边界：古典智慧与现代概率论不能生硬一一对应；本书是桥梁。")
    quote(doc, "水善利万物而不争。知止，所以不殆。")

    # —— 附录 ——
    h1(doc, "附录")
    h2(doc, "附录 A　关键概念对照表")
    mapping = [
        ("反者", "反者道之动", "antifragile", "从波动中获益"),
        ("强韧", "坚强/固", "robust", "抗得住，但不增益"),
        ("脆弱", "飘风/盈/壮", "fragile", "波动下受损"),
        ("弱用", "弱者道之用", "optionality / redundancy", "柔、余、选项的力量"),
        ("日损", "为道日损", "via negativa", "减法优先"),
        ("无为", "无为而无不为", "non-predictive action", "不多事，让结构工作"),
        ("知止", "知止不殆", "skin in the game / bounds", "有边界的暴露"),
        ("备势", "道/势/水", "optionality + barbell", "准备不对称结构"),
        ("天真干预", "为之者败之", "naive intervention", "好心制造脆弱"),
        ("过补偿", "反弹/损而益", "overcompensation", "可吸收压力后的增强"),
    ]
    for term, dao_t, anti_t, one in mapping:
        bullet(doc, f"{term}｜{dao_t}｜{anti_t}｜{one}")

    h2(doc, "附录 B　可实践的「反—弱」清单（20 条）")
    checklist = [
        "每周保留空白时间（留白预算）。",
        "收入/客户不要单一到不可替代。",
        "任何建议先问：错了谁付钱？",
        "把大决策拆成可逆的小试验。",
        "现金与睡眠是被低估的对冲。",
        "删除一个让你凹性暴露的习惯。",
        "对「平均而言很安全」保持怀疑。",
        "在能力圈内激进，在圈外保守（杠铃）。",
        "允许小失败，禁止一次性致命失败。",
        "少读预测，多看暴露。",
        "保留可退出条款：合同、合作、身份叙事。",
        "用冗余对抗效率迷信：备份、库存、第二技能。",
        "把冲击切碎：分期、分仓、分阶段。",
        "先问不做什么，再问做什么。",
        "警惕流畅的专家语言。",
        "为恢复安排时间，而不仅仅为输出安排时间。",
        "检查你是否在吃别人的上行、转嫁下行。",
        "用时间检验：这个结构十年后还活得了吗？",
        "剂量管理：同一件好事，过量即毒。",
        "知止：写下三件你永远不碰的事。",
    ]
    for i, item in enumerate(checklist, 1):
        bullet(doc, f"{i}. {item}")

    h2(doc, "附录 C　延伸阅读与争议说明")
    bullet(doc, "对照阅读：《道德经》《反脆弱》《黑天鹅》《随机漫步的傻瓜》《非对称风险》。")
    bullet(doc, "可能争议：是否过度比附古典与现代？「无为」是否被误读为消极？「反脆弱」是否被鸡汤化？")
    bullet(doc, "本书立场：宁要可检验的结构讨论，不要心灵鸡汤；宁要短章锋利，不要体系炫耀。")

    doc.add_page_break()
    h1(doc, "写作与出版执行建议")
    h2(doc, "章节节奏")
    bullet(doc, "每章：一个隐喻 + 一个判断 + 两三个例证 + 一句可记诵的收束。")
    bullet(doc, "卷与卷之间用 1–2 页过渡短页，说明下一卷要解决什么问题。")
    bullet(doc, "可设贯穿人物：「脆弱专家」「市集老人」「带仓位的实践者」。")

    h2(doc, "语气")
    bullet(doc, "冷静、带刺、反感装腔；但不犬儒到虚无。")
    bullet(doc, "对《道德经》保持敬意，对滥用它的成功学保持讽刺。")
    bullet(doc, "取《反脆弱》的方法与气质，不复制其文化战争式情绪密度。")

    h2(doc, "建议写作顺序")
    bullet(doc, "先写第 1–3、18、21、28、33、36 章（定调与骨架）。")
    bullet(doc, "再补各卷诊断与例证章。")
    bullet(doc, "最后写前言、后记与附录。")
    bullet(doc, "统稿时删掉所有「正确但无锋」的段落。")

    h2(doc, "副标题备选")
    bullet(doc, "逆向、波动与柔弱之用（推荐）")
    bullet(doc, "从《道德经》到反脆弱")
    bullet(doc, "为什么越想稳定，越容易崩溃")

    divider(doc)
    h2(doc, "一页纸总纲")
    p(doc, "《反者道之动》用七卷短章回答七个问题：", first_line=False)
    for i, q in enumerate([
        "「反」到底是什么？——回返、反弹、逆向获益。",
        "现代人为何系统性消灭它？——效率、干预、均值崇拜。",
        "不预测如何行动？——备势，不备答案。",
        "「弱」如何有用？——选择权、冗余、杠铃、可退出。",
        "机制是什么？——非线性、剂量、时间、凸性。",
        "怎么做？——为道日损，先减法。",
        "如何不把智慧变成掠夺？——皮肤在局，知止。",
    ], 1):
        bullet(doc, f"{i}. {q}")
    quote(doc, "如果你喜欢《反脆弱》的锋利，也隐约觉得《道德经》在说同一件事——这本书就是那座桥。")

    doc.save(OUT)
    print(f"已生成：{OUT}")
    return OUT


if __name__ == "__main__":
    build()
