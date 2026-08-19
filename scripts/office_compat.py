#!/usr/bin/env python3
"""修复腾讯 WorkBuddy / python-pptx 生成的 Office 文件，使其能被 PowerPoint、WPS、Word、Excel 打开。

WorkBuddy 底层用 python-pptx。把默认 4:3 模板改成 16:9 时，会留下
``type="screen4x3"``，但宽高已是 16:9。Office / WPS 会判定文件损坏。
"""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

# ECMA-376 预设尺寸（EMU）
SCREEN_16X9 = (12192000, 6858000)
SCREEN_4X3 = (9144000, 6858000)


def apply_pptx_compat(prs) -> None:
    """在 python-pptx Presentation 保存前修正幻灯片尺寸标记。"""
    sld_sz = prs._element.find(
        "{http://schemas.openxmlformats.org/presentationml/2006/main}sldSz"
    )
    if sld_sz is None:
        return
    cx = int(sld_sz.get("cx", "0") or 0)
    cy = int(sld_sz.get("cy", "0") or 0)
    cx, cy, size_type = _normalize_size(cx, cy)
    sld_sz.set("cx", str(cx))
    sld_sz.set("cy", str(cy))
    if size_type:
        sld_sz.set("type", size_type)
    elif "type" in sld_sz.attrib:
        del sld_sz.attrib["type"]


def repair_office_file(src: Path, dst: Path | None = None) -> Path:
    """修复 pptx / docx / xlsx。默认覆盖写到同目录 ``*_已修复.ext``。"""
    src = Path(src)
    data = src.read_bytes()
    if not data.startswith(b"PK"):
        raise ValueError(f"不是 Office Open XML 压缩包：{src}")

    suffix = src.suffix.lower()
    if suffix == ".pptx":
        repaired = _repair_pptx(data)
    elif suffix == ".xlsx":
        repaired = _repair_xlsx(data)
    elif suffix == ".docx":
        repaired = _repair_docx(data)
    else:
        raise ValueError(f"不支持的类型：{suffix}")

    if dst is None:
        dst = src.with_name(f"{src.stem}_已修复{src.suffix}")
    dst = Path(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_bytes(repaired)
    return dst


def _normalize_size(cx: int, cy: int) -> tuple[int, int, str | None]:
    if cy <= 0:
        return cx, cy, None
    ratio = cx / cy
    if abs(ratio - 16 / 9) < 0.08:
        return (*SCREEN_16X9, "screen16x9")
    if abs(ratio - 4 / 3) < 0.08:
        return (*SCREEN_4X3, "screen4x3")
    return cx, cy, None


def _rewrite_zip(original: bytes, replacements: dict[str, bytes]) -> bytes:
    src = zipfile.ZipFile(io.BytesIO(original))
    names = src.namelist()
    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w") as dst:
        for name in names:
            data = replacements.get(name, src.read(name))
            info = zipfile.ZipInfo(filename=name)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 0
            dst.writestr(info, data)
    return out_buf.getvalue()


def _fix_sldsz_xml(xml: str) -> tuple[str, bool]:
    match = re.search(r"<p:sldSz\b[^>/]*/>", xml)
    if not match:
        return xml, False
    tag = match.group(0)
    cx_m = re.search(r'cx="(\d+)"', tag)
    cy_m = re.search(r'cy="(\d+)"', tag)
    if not cx_m or not cy_m:
        return xml, False
    cx, cy, size_type = _normalize_size(int(cx_m.group(1)), int(cy_m.group(1)))
    old_type = re.search(r'type="([^"]*)"', tag)
    already_ok = (
        int(cx_m.group(1)) == cx
        and int(cy_m.group(1)) == cy
        and old_type
        and old_type.group(1) == size_type
    )
    if size_type:
        new_tag = f'<p:sldSz cx="{cx}" cy="{cy}" type="{size_type}"/>'
    else:
        new_tag = f'<p:sldSz cx="{cx}" cy="{cy}"/>'
    if already_ok or tag == new_tag:
        return xml, False
    return xml[: match.start()] + new_tag + xml[match.end() :], True


def _fix_empty_shape_names(xml: str) -> str:
    counter = {"n": 0}

    def repl(match: re.Match[str]) -> str:
        counter["n"] += 1
        return match.group(1) + f'name="Shape {counter["n"]}"'

    return re.sub(r'(<p:cNvPr\b[^>]*?)\bname=""', repl, xml)


def _expand_micro_ext(xml: str) -> str:
    """把几乎为 0 的 a:ext 拉到可读尺寸，避免文字竖成一列。"""

    def repl(match: re.Match[str]) -> str:
        cx = int(match.group(1))
        cy = int(match.group(2))
        if 0 < cx < 50000 and 0 < cy < 50000:
            return '<a:ext cx="9144000" cy="685800"/>'
        return match.group(0)

    return re.sub(r'<a:ext cx="(\d+)" cy="(\d+)"/>', repl, xml)


def _repair_pptx(data: bytes) -> bytes:
    zf = zipfile.ZipFile(io.BytesIO(data))
    replacements: dict[str, bytes] = {}
    if "ppt/presentation.xml" in zf.namelist():
        xml = zf.read("ppt/presentation.xml").decode("utf-8")
        xml, _ = _fix_sldsz_xml(xml)
        replacements["ppt/presentation.xml"] = xml.encode("utf-8")
    for name in zf.namelist():
        if not name.startswith("ppt/slides/slide") or not name.endswith(".xml"):
            continue
        xml = zf.read(name).decode("utf-8")
        xml = _fix_empty_shape_names(xml)
        xml = _expand_micro_ext(xml)
        replacements[name] = xml.encode("utf-8")
    return _rewrite_zip(data, replacements)


def _repair_xlsx(data: bytes) -> bytes:
    zf = zipfile.ZipFile(io.BytesIO(data))
    replacements: dict[str, bytes] = {}
    rels_name = "xl/_rels/workbook.xml.rels"
    if rels_name in zf.namelist():
        xml = zf.read(rels_name).decode("utf-8")
        new_xml = xml.replace('Target="/xl/', 'Target="')
        if new_xml != xml:
            replacements[rels_name] = new_xml.encode("utf-8")
    return _rewrite_zip(data, replacements) if replacements else data


def _repair_docx(data: bytes) -> bytes:
    """Word 一般能打开；仍重打包去掉可能导致 WPS 报错的额外 ZIP 字段。"""
    return _rewrite_zip(data, {})
