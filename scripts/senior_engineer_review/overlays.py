"""审阅版里新增的两类图形：工作总结页的外框，以及每页都有的“非官方”标识。"""

from __future__ import annotations

import html
from pathlib import Path

import layout as L
import prepare_fonts
from chrome_pdf import print_to_pdf


def _y(top: float) -> float:
    """版面坐标自上而下，PDF 坐标自下而上。"""
    return L.PAGE_H - top


def frame_stream(summary_index: int, last_frame_bottom: float) -> bytes:
    """工作总结第 summary_index 页（从 0 起）的外框。

    首页外框自 33.0 起并带表头分隔线；续页自 18.8 起；末页在正文结束处收框，
    落款留在框外，和平台原表一致。
    """
    last = summary_index == L.SUMMARY_PAGE_COUNT - 1
    top = L.FRAME_TOP_FIRST if summary_index == 0 else L.FRAME_TOP
    bottom = last_frame_bottom if last else L.FRAME_BOTTOM
    width = L.FRAME_RIGHT - L.FRAME_LEFT
    ops = [
        f"{L.FRAME_GRAY} G {L.FRAME_LINE_WIDTH} w 0 J 0 j",
        f"{L.FRAME_LEFT:.2f} {_y(bottom):.2f} {width:.2f} {bottom - top:.2f} re S",
    ]
    if summary_index == 0:
        y = _y(L.FORM_HEAD_BOTTOM)
        ops.append(f"{L.FRAME_LEFT:.2f} {y:.2f} m {L.FRAME_RIGHT:.2f} {y:.2f} l S")
    return ("\n".join(ops) + "\n").encode("ascii")


def marks_html(font_dir: Path) -> str:
    family, weight = L.PDF_FONT[("hei", True)]
    red, green, blue = L.MARK_COLOR
    color = f"rgb({red},{green},{blue})"
    diagonals = "".join(
        f'<div class="diag" style="top:{y - L.MARK_DIAGONAL_SIZE / 2:.2f}pt">{html.escape(L.MARK_DIAGONAL)}</div>'
        for y in L.MARK_DIAGONAL_CENTERS_Y
    )
    # 页脚基线离纸底 MARK_FOOTER_BASELINE_FROM_BOTTOM；行框高度取字号，基线约在行框 0.88 处。
    footer_top = L.PAGE_H - L.MARK_FOOTER_BASELINE_FROM_BOTTOM - L.MARK_FOOTER_SIZE * 0.88
    css = "\n".join(
        [
            f"@page{{size:{L.PAGE_W}pt {L.PAGE_H}pt;margin:0;}}",
            "html,body{margin:0;padding:0;background:transparent;}",
            f"body{{width:{L.PAGE_W}pt;height:{L.PAGE_H - 1}pt;position:relative;overflow:hidden;"
            f'font-family:"{family}";font-weight:{weight};color:{color};font-synthesis:none;}}',
            f".diag{{position:absolute;left:0;width:{L.PAGE_W}pt;text-align:center;white-space:nowrap;"
            f"font-size:{L.MARK_DIAGONAL_SIZE}pt;line-height:{L.MARK_DIAGONAL_SIZE}pt;"
            f"opacity:{L.MARK_DIAGONAL_OPACITY};transform:rotate(-{L.MARK_DIAGONAL_ANGLE}deg);}}",
            f".foot{{position:absolute;left:0;width:{L.PAGE_W}pt;top:{footer_top:.2f}pt;text-align:center;"
            f"white-space:nowrap;font-size:{L.MARK_FOOTER_SIZE}pt;line-height:{L.MARK_FOOTER_SIZE}pt;}}",
        ]
    )
    fonts = prepare_fonts.font_face_css(font_dir, L.PDF_FONT)
    return (
        '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">'
        f"<style>{fonts}\n{css}</style></head><body>{diagonals}"
        f'<div class="foot">{html.escape(L.MARK_FOOTER)}</div></body></html>'
    )


def build_marks_pdf(out_pdf: Path, font_dir: Path) -> None:
    """生成一页透明底的标识页，之后作为表单对象叠到每一页上。"""
    path = prepare_fonts.prepare(font_dir)[("hei", True)]
    missing = prepare_fonts.missing_chars(path, L.MARK_FOOTER + L.MARK_DIAGONAL)
    if missing:
        raise RuntimeError(f"{path.name} 缺少字形：{''.join(missing)}")
    print_to_pdf(marks_html(font_dir), out_pdf)
