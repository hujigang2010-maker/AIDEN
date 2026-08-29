#!/usr/bin/env python3
"""把 deliverables 下全部文字、图片、PDF、Office 文件打成一个压缩包。"""

from __future__ import annotations

from datetime import date
from pathlib import Path
import zipfile

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "deliverables"
ZIP_NAME = "青岛抚顺路和哈尔滨路路口交通事故_全部材料_20260829.zip"
README_NAME = "压缩包目录说明.txt"

SKIP_SUFFIXES = {".zip"}
SKIP_NAMES = {".DS_Store"}

README_TEXT = """青岛抚顺路和哈尔滨路路口交通事故 · 全部材料压缩包
打包日期：2026-08-29

本压缩包为家属内部材料合集。打印给对方的文件和内部测算件请分开使用，不要混发。

【可以打印或复制给对方的】
- 青岛…给刘孝春的垫付与护理说明_20260819.docx / .pdf（给刘孝春当面核对，不含内部比例）
- 给美团骑手的跟进微信_2026-08-22.md（复制「请直接复制发出」段）
- 给平安理赔专员陈老师的回复_2026-08-29.md（复制首次回复段发给陈老师）
- 授权委托书_胡志远委托胡继刚_草稿.md（打印后由胡志远签字，交交警/保险出示）
- 给护工的无糖床头食谱（Markdown / Word / PDF，明天打印贴床头）

【只给家属和岳父看，不要发给骑手、刘孝春或任何保险公司】
- 青岛…完整解决方案_20260822.docx
- 青岛…赔付测算表_20260822.xlsx
- 全案整理_截至2026-08-22.md
- 家庭执行方案_2026-08-29.md
- 给妈妈跟刘孝春的接话卡（Word/PDF）及沟通口径
- 给律师的完整经过、病历图文、胡志远病例文件.pdf、chart-photos/

【怎么打开】
- Word / Excel / PPT / PDF：用 Microsoft Office、WPS 或预览打开
- .md：用记事本、Word 或 Typora 打开
- .html：用浏览器直接打开（hongfeng-guide.html 或 处理总览.html）
- chart-photos/：病历抽出页照片
- 事故3D复原_示范动画.mp4：示意图，不是证据，待用

【重新生成】在仓库根目录执行：python3 scripts/build_all.py
【重新打包】在仓库根目录执行：python3 scripts/pack_all.py
"""


def iter_files() -> list[Path]:
    files: list[Path] = []
    for path in sorted(OUT.rglob("*")):
        if not path.is_file():
            continue
        if path.name in SKIP_NAMES or path.suffix.lower() in SKIP_SUFFIXES:
            continue
        files.append(path)
    return files


def build_zip(dest: Path | None = None) -> Path:
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / README_NAME).write_text(README_TEXT, encoding="utf-8")
    dest = dest or (OUT / ZIP_NAME)
    files = iter_files()
    if dest in files:
        files.remove(dest)
    with zipfile.ZipFile(dest, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in files:
            zf.write(path, arcname=str(path.relative_to(OUT)))
    return dest


def main() -> None:
    dest = build_zip()
    with zipfile.ZipFile(dest) as zf:
        n = len(zf.namelist())
    print(f"已打包 {n} 个条目 → {dest}（{dest.stat().st_size} 字节）")
    print(f"打包日期标记：{date.today().isoformat()}")


if __name__ == "__main__":
    main()
