# -*- coding: utf-8 -*-
"""从第一遍生成的 PDF 中提取各章节页码，写入 toc.json 供第二遍回填目录页码。"""
import json
import re
import sys

from pypdf import PdfReader

import report_content as rc

PDF = "/workspace/report/赛普客户-房企十五五竞争的破局关键-客户服务体系化建设.pdf"
OUT = "/workspace/scripts/toc.json"


def norm(s):
    return "".join(s.split())


def main():
    reader = PdfReader(PDF)
    raw = [p.extract_text() or "" for p in reader.pages]
    pages = [norm(t) for t in raw]
    # 目录页识别：含 3 处以上连续 8 个点的点线引导
    dot_run = re.compile(r"\.{8,}")
    toc_pages = {i for i, t in enumerate(raw) if len(dot_run.findall(t)) >= 3}
    tocmap = {}
    titles = [blk[1] for blk in rc.BLOCKS if blk[0] in ("h1", "h2")]
    for title in titles:
        needle = norm(title)
        found = ""
        for i in range(len(pages)):
            if i in toc_pages:
                continue
            if needle in pages[i]:
                found = i + 1  # PDF 页序（封面=1）
                break
        tocmap[norm(title)] = found
    json.dump(tocmap, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    missing = [t for t in titles if not tocmap[norm(t)]]
    print("total pages:", len(pages), "| toc pages:", sorted(i + 1 for i in toc_pages))
    print("missing:", missing if missing else "none")
    print("saved", OUT)


if __name__ == "__main__":
    main()
