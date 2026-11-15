"""P8—P11《本人任现职以来工作总结》的版式规格。

PDF 重排与可编辑 DOCX 共用这一份规格，保证两者的字号、行距、标题层级一致。
长度单位均为 pt（1 pt = 1/72 inch）。
"""

from __future__ import annotations

from dataclasses import dataclass

PAGE_W = 595.28
PAGE_H = 841.89
PAGE_COUNT = 14

# 外框沿用平台导出原表的几何：左右 54.0 / 529.5；首页框顶 33.0，
# 表头栏底线 90.0；续页框顶 18.8，框底 810.8。
FRAME_LEFT = 54.0
FRAME_RIGHT = 529.5
FRAME_TOP_FIRST = 33.0
FORM_HEAD_BOTTOM = 90.0
FRAME_TOP = 18.8
FRAME_BOTTOM = 810.8
FRAME_LINE_WIDTH = 0.75
FRAME_GRAY = 0.2

PAD_X = 9.0
PAD_TOP = 8.0
PAD_BOTTOM = 5.5

TEXT_LEFT = FRAME_LEFT + PAD_X
TEXT_RIGHT = FRAME_RIGHT - PAD_X
TEXT_WIDTH = TEXT_RIGHT - TEXT_LEFT

# 末页：正文结束后收框，落款置于框外（与平台导出原表一致）。
LAST_FRAME_PAD = 8.0
SIGN_GAP = 26.0

SUMMARY_FIRST_PAGE = 8
SUMMARY_PAGE_COUNT = 4


@dataclass(frozen=True)
class Style:
    """一个文字层级的规格。family 取 song（宋体）或 hei（黑体）。"""

    family: str
    size: float
    bold: bool
    leading: float
    before: float = 0.0
    after: float = 0.0
    indent_em: float = 0.0
    align: str = "justify"


STYLES: dict[str, Style] = {
    # 表头栏属于表格模板本身，沿用其余各页表头的黑体。
    "form_title": Style("hei", 20.0, True, 26.0, align="center"),
    "form_note": Style("song", 9.0, False, 12.0, align="center"),
    # 总结正文一律宋体 / 宋体加粗。
    "title": Style("song", 14.0, True, 20.0, before=0.0, after=5.0, align="center"),
    "h1": Style("song", 12.0, True, 17.0, before=6.0, after=2.0, align="left"),
    "h2": Style("song", 10.5, True, 14.75, before=4.5, after=0.5, align="left"),
    "p": Style("song", 10.5, False, 14.75, indent_em=2.0),
    "note": Style("song", 9.0, False, 13.0, before=4.0, indent_em=2.0),
    "table": Style("song", 9.0, False, 12.0, before=3.0, after=3.0, align="center"),
    "sign": Style("song", 10.5, False, 20.0, align="right"),
}

TABLE_CELL_PAD_Y = 1.5
TABLE_CELL_PAD_X = 3.0
TABLE_LINE_WIDTH = 0.75

# PDF 内嵌字体：本环境没有中易宋体（SimSun）的授权字体文件，
# 用开源的思源宋体 / 思源黑体承担“宋体 / 宋体加粗 / 黑体”，加粗使用真粗体字重而非描边伪粗。
# 这里是排版 HTML 内部使用的 CSS 家族名和字重，字体文件由 prepare_fonts.py 提供。
PDF_FONT = {
    ("song", False): ("ReviewSong", 400),
    ("song", True): ("ReviewSong", 700),
    ("hei", True): ("ReviewHei", 700),
}

# DOCX 里直接指定“宋体 / 黑体”，在装有中易字库的 Word / WPS 上会用真正的宋体显示。
DOCX_FONT = {"song": "宋体", "hei": "黑体"}

TEXT_COLOR = "#111111"

# 每页都要带的“非官方”标识。页脚一行 + 斜向浅色底纹，直接写进页面内容，
# 不做成可一键隐藏的批注。
MARK_FOOTER = (
    "非官方排版审阅版 · 不具官方效力 · 印章外观不等于有效电子签章 · "
    "正式用途以江苏省官方归档件为准"
)
MARK_DIAGONAL = "非官方排版审阅版　不具官方效力"
MARK_COLOR = (176, 32, 24)
MARK_FOOTER_SIZE = 7.5
MARK_FOOTER_BASELINE_FROM_BOTTOM = 12.0
MARK_DIAGONAL_SIZE = 26.0
MARK_DIAGONAL_OPACITY = 0.12
MARK_DIAGONAL_ANGLE = 28.0
MARK_DIAGONAL_CENTERS_Y = (190.0, 430.0, 660.0)
