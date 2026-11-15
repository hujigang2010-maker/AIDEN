"""准备排版用的静态 TrueType 字体（思源宋体 Regular / Bold、思源黑体 Bold）。

为什么不直接用系统里的 Noto CJK：系统包是 CFF 轮廓的 OTC，Chrome 打印 PDF 时
只能把它们写成 Type 3 字体，各阅读器对 Type 3 的渲染、搜索和复制表现不一致。
这里从 Google Fonts 取 TrueType 轮廓的可变字体，定在 400 / 700 字重生成静态实例，
Chrome 就能按标准的 CID TrueType 子集内嵌，并带 ToUnicode 映射。

字体以 SIL OFL 1.1 授权，允许内嵌进文档。

用法：
    python3 prepare_fonts.py --font-dir ./fonts
"""

from __future__ import annotations

import argparse
import hashlib
import urllib.request
from pathlib import Path

from fontTools.ttLib import TTFont
from fontTools.varLib import instancer

SOURCES = {
    "NotoSerifSC-VF.ttf": (
        "https://github.com/google/fonts/raw/main/ofl/notoserifsc/NotoSerifSC%5Bwght%5D.ttf",
        "050080d9255a86808f2945bffac582b31ef32bc36411ce29563b4961670c66f9",
    ),
    "NotoSansSC-VF.ttf": (
        "https://github.com/google/fonts/raw/main/ofl/notosanssc/NotoSansSC%5Bwght%5D.ttf",
        "a3041811a78c361b1de50f953c805e0244951c21c5bd412f7232ef0d899af0da",
    ),
}

# 输出文件名 -> (来源可变字体, 字重, 家族名, 样式名, PostScript 名)
INSTANCES = {
    "NotoSerifSC-Regular.ttf": ("NotoSerifSC-VF.ttf", 400, "Noto Serif SC", "Regular", "NotoSerifSC-Regular"),
    "NotoSerifSC-Bold.ttf": ("NotoSerifSC-VF.ttf", 700, "Noto Serif SC", "Bold", "NotoSerifSC-Bold"),
    "NotoSansSC-Bold.ttf": ("NotoSansSC-VF.ttf", 700, "Noto Sans SC", "Bold", "NotoSansSC-Bold"),
}

# 版式规格里的 (家族, 是否加粗) -> 静态字体文件
FONT_FILES = {
    ("song", False): "NotoSerifSC-Regular.ttf",
    ("song", True): "NotoSerifSC-Bold.ttf",
    ("hei", True): "NotoSansSC-Bold.ttf",
}


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _download(name: str, font_dir: Path) -> Path:
    url, digest = SOURCES[name]
    path = font_dir / name
    if not path.exists() or _sha256(path) != digest:
        with urllib.request.urlopen(url, timeout=300) as resp:
            path.write_bytes(resp.read())
    if _sha256(path) != digest:
        raise RuntimeError(f"{name} 的校验值与预期不符，上游字体可能已更新，请核对后更新 SOURCES")
    return path


def _set_names(font: TTFont, family: str, style: str, ps_name: str) -> None:
    full = f"{family} {style}"
    values = {1: family, 2: style, 3: f"{ps_name};instanced", 4: full, 6: ps_name}
    name = font["name"]
    name.names = [n for n in name.names if n.nameID not in (16, 17, 25)]
    for record in list(name.names):
        if record.nameID in values:
            name.setName(values[record.nameID], record.nameID, record.platformID, record.platEncID, record.langID)
    bold = style == "Bold"
    font["OS/2"].usWeightClass = 700 if bold else 400
    font["OS/2"].fsSelection = (font["OS/2"].fsSelection & ~0b1100001) | (0b100000 if bold else 0b1000000)
    font["head"].macStyle = 1 if bold else 0


def prepare(font_dir: Path) -> dict[tuple[str, bool], Path]:
    font_dir.mkdir(parents=True, exist_ok=True)
    for out_name, (src_name, weight, family, style, ps_name) in INSTANCES.items():
        out_path = font_dir / out_name
        if out_path.exists():
            continue
        src = _download(src_name, font_dir)
        static = instancer.instantiateVariableFont(TTFont(src), {"wght": weight}, inplace=False)
        _set_names(static, family, style, ps_name)
        static.save(out_path)
    return {key: font_dir / fname for key, fname in FONT_FILES.items()}


def missing_chars(font_path: Path, text: str) -> list[str]:
    """返回字体 cmap 里没有的字符，用来在排版前拦住缺字。"""
    cmap = TTFont(font_path, lazy=True).getBestCmap()
    return sorted({ch for ch in text if not ch.isspace() and ord(ch) not in cmap})


def font_face_css(font_dir: Path, pdf_font: dict[tuple[str, bool], tuple[str, int]]) -> str:
    """把静态字体文件登记成排版 HTML 里使用的 CSS 家族。"""
    rules = []
    for key, path in prepare(font_dir).items():
        family, weight = pdf_font[key]
        rules.append(
            f'@font-face{{font-family:"{family}";font-weight:{weight};font-style:normal;'
            f'src:url("{path.resolve().as_uri()}") format("truetype");}}'
        )
    return "\n".join(rules)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--font-dir", required=True, type=Path)
    args = parser.parse_args()
    for key, path in prepare(args.font_dir).items():
        print(key, path, path.stat().st_size)


if __name__ == "__main__":
    main()
