#!/usr/bin/env python3
"""按原文排版合作原则。不增字，不减字。"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "deliverables"
BODY_FONT = "宋体"
HEAD_FONT = "黑体"

PINE = "1F4E4A"
PINE_RGB = (0x1F, 0x4E, 0x4A)
INK = (0x1C, 0x1C, 0x1C)
RULE = "C8B48A"

# 来稿原文。标题上的 Markdown 加粗标记不属于正文。
ORIGINAL = """赵同学：

结合今天上午的沟通，我认真考虑了一下。现在距离11月28日大会只有14天，时间确实非常紧张。我愿意发挥自身及相关平台的资源优势，全力推动大会，但为了避免后续产生误解，也为了双方能够集中精力开展工作，有几个合作原则需要一次性明确。

一、关于双方合作定位

本次合作以共同推动大会成功举办为目标。我负责协调高校专家、养老科技、人工智能、机器人、金融投资等相关嘉宾及产业资源；你负责对接组委会、场地、会务组织及既有合作单位。

双方各司其职、相互配合，但不将一方此前对第三方作出的承诺转移给另一方承担。

二、关于复旦大学住房政策研究中心的责任边界

我以复旦大学住房政策研究中心相关工作对接代表的身份参与筹备，在实际授权范围内协调学术和产业资源。

中心的正式参与、对外署名及邀请函使用等，均须符合相应授权要求。

复旦大学及复旦大学住房政策研究中心不承担大会资金、场地费用、招商指标、嘉宾到场数量及活动举办效果等保证责任，也不为任何第三方作出的承诺兜底。

三、关于嘉宾邀请及三分之一目标

目前详细议程及邀请函已经准备完成，但相关嘉宾尚未正式启动集中邀请。

我愿意充分发挥个人和相关平台的资源优势，尽最大努力争取高质量嘉宾参会。

但嘉宾最终能否出席，取决于个人意愿、工作安排及档期等多方面因素。

议程中约三分之一嘉宾的目标，可以作为双方努力争取的方向，但不能作为我方必须完成的指标，更不能形成对赌、保底或未达标赔偿责任。

四、关于赞助、招商及展位费用

我可以协助对接相关企业、金融机构和潜在赞助方，但不承诺任何具体赞助金额，也不承担招商保底任务。

对于展位、场地及其他活动成本，由实际经营或签约责任主体与组委会协商落实。

未经我方事先书面同意，不得要求我个人或复旦大学住房政策研究中心承担任何费用。涉及新增费用或商业收益分配的，须另行协商确认。

五、关于主办方及组织架构

充分尊重上海星智耀科技有限公司作为主办单位的既有安排，以及你前期协调形成的合作关系。

但主办单位应承担其依法依约应承担的组织、运营和相关商务责任，不因高校及其他机构参与协办、资源协调而转移。

现有指导、主办、承办、协办、支持及学术支持单位名单，在相关单位同意的前提下，以双方最终确认的版本为准。

原则上11月15日起不再随意调整组织架构。

六、关于后续增补单位与嘉宾

如现有嘉宾邀请未达到预期，可以继续增补与大会主题高度相关的养老科技、人工智能、机器人、智慧医疗、科技服务及产业投资机构。

既有辰山、水乡客厅、艺术学院等合作资源应得到合理尊重和妥善安排。

但新增单位原则上以嘉宾或产业合作方身份参与，不再随意增加主办、承办及支持单位，避免影响大会整体定位和邀请工作的连续性。

七、关于工作决策及沟通机制

双方确认最终议程、邀请函和组织架构后，各自负责相应资源协调工作。

我负责的嘉宾邀请方式、渠道和沟通节奏，由我根据实际情况自主安排；你负责组委会及相关会务协调。

重大事项提前沟通，具体执行以书面确认的信息为准，尽量减少反复修改方案和重复性沟通。

八、关于合作风险及退出机制

双方均应尽合理努力推动大会举办，但不对嘉宾到场数量、招商金额及最终活动效果提供保证。

如果最终嘉宾、赞助或合作资源没有达到预期，双方应协商调整方案，不因此追究另一方未作出保证的结果责任。

如后续出现新的保底要求、费用承担或超出授权的责任要求，且双方无法达成一致，任何一方均可停止尚未开展的合作事项，并妥善处理已形成的对外沟通，不得继续未经授权使用另一方或相关机构名义。"""

HEADINGS = tuple(f"{label}、" for label in "一二三四五六七八")


def set_run_font(run, name: str, size: float, bold: bool = False, color: tuple | None = None) -> None:
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = False
    if color:
        run.font.color.rgb = RGBColor(*color)
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        r_fonts.set(qn(attr), name)


def set_style_font(style, name: str, size: float) -> None:
    style.font.name = name
    style.font.size = Pt(size)
    r_pr = style.element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        r_fonts.set(qn(attr), name)


def paragraph_border(paragraph, color: str = RULE) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def add_page_field(paragraph) -> None:
    run = paragraph.add_run()
    set_run_font(run, BODY_FONT, 9, color=(0x8A, 0x8A, 0x8A))
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


def setup(doc: Document) -> None:
    set_style_font(doc.styles["Normal"], BODY_FONT, 12)
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.15)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.header_distance = Cm(0.7)
    section.footer_distance = Cm(0.6)

    header = section.header.paragraphs[0]
    header.paragraph_format.space_before = Pt(0)
    header.paragraph_format.space_after = Pt(0)
    paragraph_border(header, PINE)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.paragraph_format.space_before = Pt(0)
    footer.paragraph_format.space_after = Pt(0)
    add_page_field(footer)

    core = doc.core_properties
    core.title = "合作原则"
    core.subject = "原文排版"
    core.category = "信函"
    core.comments = "正文与来稿逐字一致。"


def kind_of(text: str, index: int) -> str:
    if index == 0:
        return "salute"
    if text.startswith(HEADINGS):
        return "heading"
    return "body"


def write_paragraph(doc: Document, text: str, kind: str, keep_next: bool):
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    paragraph.paragraph_format.widow_control = True
    paragraph.paragraph_format.keep_with_next = keep_next
    if kind == "salute":
        paragraph.paragraph_format.space_before = Pt(2)
        paragraph.paragraph_format.space_after = Pt(12)
        paragraph.paragraph_format.line_spacing = Pt(24)
        set_run_font(paragraph.add_run(text), BODY_FONT, 14, bold=True, color=INK)
    elif kind == "heading":
        paragraph.paragraph_format.space_before = Pt(15)
        paragraph.paragraph_format.space_after = Pt(6)
        paragraph.paragraph_format.line_spacing = Pt(24)
        set_run_font(paragraph.add_run(text), HEAD_FONT, 13.5, bold=True, color=PINE_RGB)
        paragraph_border(paragraph)
    else:
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(8)
        paragraph.paragraph_format.line_spacing = Pt(26)
        set_run_font(paragraph.add_run(text), BODY_FONT, 12, color=INK)
    return paragraph


def build(path: Path) -> None:
    blocks = [block.strip() for block in ORIGINAL.split("\n\n")]
    if any(not block or "\n" in block for block in blocks):
        raise SystemExit("原文分段异常")
    doc = Document()
    setup(doc)
    for index, text in enumerate(blocks):
        following = blocks[index + 1] if index + 1 < len(blocks) else ""
        section_continues = bool(following) and not following.startswith(HEADINGS)
        write_paragraph(doc, text, kind_of(text, index), section_continues)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


def validate(path: Path) -> None:
    document = Document(path)
    texts = [paragraph.text for paragraph in document.paragraphs]
    # 页眉页脚不在 paragraphs 里。正文必须与原文逐段一致。
    rendered = "\n\n".join(texts)
    if rendered != ORIGINAL:
        original_chars = list(ORIGINAL)
        rendered_chars = list(rendered)
        limit = max(len(original_chars), len(rendered_chars))
        for index in range(limit):
            left = original_chars[index] if index < len(original_chars) else "∅"
            right = rendered_chars[index] if index < len(rendered_chars) else "∅"
            if left != right:
                raise SystemExit(
                    f"第 {index} 个字符不一致：原文 {left!r}，成稿 {right!r}\n"
                    f"原文字数 {len(ORIGINAL)}，成稿字数 {len(rendered)}"
                )
        raise SystemExit("正文与原文不一致")
    body_only = "".join(texts)
    source_only = ORIGINAL.replace("\n", "")
    if body_only != source_only:
        raise SystemExit("去掉换行后仍不一致")
    print(f"已生成 {path}")
    print(f"正文 {len(blocks_count(texts))} 段，{len(source_only)} 字，与原文逐字一致")


def blocks_count(texts: list[str]) -> list[str]:
    return texts


def main() -> None:
    destination = OUT / "合作原则.docx"
    build(destination)
    validate(destination)
    for stale in OUT.glob("*PRD*"):
        stale.unlink()
        print(f"已移除旧文件 {stale.name}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(exc, file=sys.stderr)
        raise
