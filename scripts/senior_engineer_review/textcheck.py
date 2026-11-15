"""文字比对用的公共函数：证明重排没有增、删、改任何一个字。"""

from __future__ import annotations

import re
from collections import Counter

import pymupdf

import layout as L

WATERMARK_TEXT = "江苏省一体化平台"

# 空白、填空下划线不算内容；长横线在不同来源里可能是 — 或 －，统一后再比。
_DROP = re.compile(r"[\s_\u3000\u00a0]+")
_DASHES = str.maketrans({"－": "—", "–": "—", "―": "—", "‑": "-", "‐": "-"})


def normalize(text: str) -> str:
    return _DROP.sub("", text).translate(_DASHES)


def _is_mark(text: str) -> bool:
    stripped = normalize(text)
    return bool(stripped) and stripped in normalize(L.MARK_FOOTER + L.MARK_DIAGONAL)


def body_lines(page: pymupdf.Page) -> list[dict]:
    """页面正文行：去掉平台水印（斜排）和本工具加的非官方标识。"""
    lines = []
    for block in page.get_text("dict")["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            text = "".join(s["text"] for s in line["spans"])
            if not text.strip():
                continue
            dx, dy = line["dir"]
            rotated = abs(dy) > 0.01 or dx < 0.99
            if rotated and (set(normalize(text)) <= set(WATERMARK_TEXT) or _is_mark(text)):
                continue
            if not rotated and _is_mark(text) and line["bbox"][1] > L.FRAME_BOTTOM:
                continue
            lines.append({"text": text, "bbox": line["bbox"], "rotated": rotated})
    lines.sort(key=lambda item: (round(item["bbox"][1], 0), item["bbox"][0]))
    return lines


def body_text(page: pymupdf.Page) -> str:
    return normalize("".join(line["text"] for line in body_lines(page)))


def char_diff(expected: str, actual: str) -> dict:
    """按字符多重集比较；任何一个字多了或少了都会列出来。"""
    want, got = Counter(expected), Counter(actual)
    return {
        "expected_chars": len(expected),
        "actual_chars": len(actual),
        "missing": dict((want - got).most_common()),
        "extra": dict((got - want).most_common()),
        "same_sequence": expected == actual,
    }
