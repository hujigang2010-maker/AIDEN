#!/usr/bin/env python3
"""把 manuscript/ 下的分章 Markdown 合并为《知行觉醒》全书书稿。

输出：
- output/知行觉醒_书稿全文.md

合并时会把源稿中的半角引号逐行成对地规范为中文弯引号（“”‘’），源稿里照常输入半角引号即可。

用法：
- python3 scripts/build_book.py          合并并打印字数统计
- python3 scripts/build_book.py --check  合并后校验目录、章节编号、表格与引号，失败时返回非零
"""

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "manuscript"
OUT = ROOT / "output" / "知行觉醒_书稿全文.md"

TOC_PLACEHOLDER = "<!-- 目录 -->"
PART_PREFIXES = ("上篇", "中篇", "下篇", "终篇")
CHAPTER_NUMERALS = [
    "一", "二", "三", "四", "五", "六", "七", "八",
    "九", "十", "十一", "十二", "十三", "十四", "十五", "十六",
]
HEADING_RE = re.compile(r"^(#{1,6})\s+(.+?)\s*$")
CJK_RE = re.compile(r"[\u4e00-\u9fff]")
# 只转换内部含汉字的单引号对，避免误改英文撇号（如 Descartes' Error）
CJK_SINGLE_QUOTE_RE = re.compile(r"'([^'\n]*?[\u4e00-\u9fff][^'\n]*?)'")


def normalize_quotes(text: str) -> str:
    """逐行把成对的半角引号换成中文弯引号；代码块和引号数为奇数的行保持原样，交给校验报告。"""
    out, in_fence = [], False
    for line in text.splitlines():
        if line.startswith("```"):
            in_fence = not in_fence
        elif not in_fence and line.count('"') % 2 == 0:
            pieces = line.split('"')
            line = "".join(p + ("“" if i % 2 == 0 else "”") for i, p in enumerate(pieces[:-1]))
            line += pieces[-1]
            line = CJK_SINGLE_QUOTE_RE.sub(r"‘\1’", line)
        out.append(line)
    return "\n".join(out) + "\n"


def github_slug(text: str, seen: dict) -> str:
    """按 GitHub 的规则生成标题锚点：转小写、去标点、空格变连字符、重名追加序号。"""
    slug = re.sub(r"[^\w\- ]", "", text.strip().lower()).replace(" ", "-")
    count = seen.get(slug, 0)
    seen[slug] = count + 1
    return slug if count == 0 else f"{slug}-{count}"


def iter_headings(text: str):
    """逐个返回 (级别, 标题)，跳过代码块内的内容。"""
    in_fence = False
    for line in text.splitlines():
        if line.startswith("```"):
            in_fence = not in_fence
            continue
        m = HEADING_RE.match(line)
        if m and not in_fence:
            yield len(m.group(1)), m.group(2)


def build_toc(body: str) -> str:
    """以二级标题生成目录：篇作为一级条目，章与篇末答问缩进挂在所属的篇下面。"""
    seen = {"目录": 1}  # 目录标题本身插在最前面，先占位以保证后续锚点序号一致
    lines = ["## 目录", ""]
    in_part = False
    for level, title in iter_headings(body):
        slug = github_slug(title, seen)
        if level != 2:
            continue
        if in_part and (title.startswith("第") or title.endswith("答问")):
            lines.append(f"  - [{title}](#{slug})")
        elif title.startswith(PART_PREFIXES):
            in_part = True
            lines.append(f"- [{title}](#{slug})")
        else:
            in_part = False
            lines.append(f"- [{title}](#{slug})")
    return "\n".join(lines)


def merge() -> str:
    parts = [p.read_text(encoding="utf-8").strip() for p in sorted(SRC.glob("*.md"))]
    body = normalize_quotes("\n\n---\n\n".join(parts))
    if TOC_PLACEHOLDER not in body:
        raise SystemExit(f"首个源文件中缺少目录占位符 {TOC_PLACEHOLDER}")
    toc = build_toc(body.replace(TOC_PLACEHOLDER, ""))
    return body.replace(TOC_PLACEHOLDER, toc)


def check(book: str) -> list:
    """返回问题列表；为空表示通过。"""
    problems = []
    if TOC_PLACEHOLDER in book:
        problems.append("目录占位符未被替换")

    seen = {}
    anchors = {github_slug(title, seen) for _, title in iter_headings(book)}
    for target in re.findall(r"\]\(#([^)]+)\)", book):
        if target not in anchors:
            problems.append(f"目录链接找不到对应标题：#{target}")

    for line in book.splitlines():
        m = re.match(r"^(\s*)- \[(.+?)\]\(#", line)
        if m and not m.group(1) and (m.group(2).startswith("第") or m.group(2).endswith("答问")):
            problems.append(f"目录条目没有挂在所属的篇下：{m.group(2)}")

    chapters = [t for lv, t in iter_headings(book) if lv == 2 and re.match(r"^第.+章 ", t)]
    expected = [f"第{n}章" for n in CHAPTER_NUMERALS]
    actual = [t.split(" ")[0] for t in chapters]
    if actual != expected:
        problems.append(f"章节编号不连续：{actual}")

    table_cols = None
    for no, line in enumerate(book.splitlines(), 1):
        if line.startswith("|"):
            cols = line.count("|") - 1
            if table_cols is None:
                table_cols = cols
            elif cols != table_cols:
                problems.append(f"第 {no} 行表格列数为 {cols}，与表头 {table_cols} 列不一致")
        else:
            table_cols = None

    in_fence = False
    for no, line in enumerate(book.splitlines(), 1):
        if line.startswith("```"):
            in_fence = not in_fence
        if in_fence:
            continue
        if '"' in line:
            problems.append(f"第 {no} 行仍有半角双引号，请检查是否成对")
        if line.count("“") != line.count("”") or line.count("‘") != line.count("’"):
            problems.append(f"第 {no} 行中文引号不配对")
    return problems


def report(book: str):
    total = 0
    print(f"{'源文件':<36}{'汉字数':>8}")
    for path in sorted(SRC.glob("*.md")):
        n = len(CJK_RE.findall(path.read_text(encoding="utf-8")))
        total += n
        print(f"{path.name:<36}{n:>8}")
    print(f"{'合计':<36}{total:>8}")
    chars = len(re.sub(r"\s", "", book))
    print(f"全文字符数（含标点，不含空白）：{chars}")


def main():
    book = merge()
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(book, encoding="utf-8")
    print(f"已生成：{OUT.relative_to(ROOT)}")
    report(book)
    if "--check" in sys.argv:
        problems = check(book)
        for p in problems:
            print(f"[失败] {p}")
        if problems:
            sys.exit(1)
        print("[通过] 目录锚点与层级、章节编号、表格列数、中文引号校验全部通过")


if __name__ == "__main__":
    main()
