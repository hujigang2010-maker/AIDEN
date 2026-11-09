#!/usr/bin/env python3
"""把针灸科现场手册打印为 A4 PDF，并导出分页预览。

依赖：Google Chrome、poppler-utils（pdftoppm、pdfinfo、pdftotext）、
系统字体 Noto Sans CJK SC。在仓库根目录执行：

    python3 scripts/build_a4_guide.py
"""

from __future__ import annotations

import re
import shutil
import subprocess
import sys
from pathlib import Path

import segno

ROOT = Path(__file__).resolve().parents[1]
HTML = ROOT / "guide" / "index.html"
QR_DIR = ROOT / "guide" / "qr"
OUT = ROOT / "output"
PDF = OUT / "岳阳医院针灸科-VPN及国外AI软件安装使用指南.pdf"
PREVIEW = OUT / "preview"

QR_ITEMS = {
    "chatgpt": "https://chatgpt.com/download/",
    "claude": "https://claude.com/download",
    "gemini": "https://gemini.google.com",
    "notebook": "https://notebook.google.com",
    "perplexity": "https://www.perplexity.ai",
    "elicit": "https://elicit.com",
    "consensus": "https://consensus.app",
    "zotero": "https://www.zotero.org/download/",
    "grok": "https://grok.com",
    "openai-regions": "https://help.openai.com/en/articles/7947663-chatgpt-supported-countries",
}

REQUIRED_TEXT = [
    "VPN及国外AI软件",
    "文献检索",
    "研究设计",
    "论文整理",
    "ik11.net/23c586b",
    "lk11.net",
    "veee",
    "数字签名",
    "chatgpt.com/download",
    "claude.com/download",
    "notebook.google.com",
    "NotebookLM",
    "elicit.com",
    "consensus.app",
    "zotero.org/download",
    "随机对照试验",
    "慢性疼痛",
    "sham acupuncture",
    "PMID",
    "中国大陆",
    "不上传",
    "培训目标",
]


def write_qr_codes() -> None:
    QR_DIR.mkdir(parents=True, exist_ok=True)
    for name, url in QR_ITEMS.items():
        qr = segno.make(url, error="m")
        rows = list(qr.matrix)
        border = 2
        size = len(rows) + border * 2
        rects: list[str] = []
        for y, row in enumerate(rows):
            for x, dark in enumerate(row):
                if dark:
                    rects.append(
                        f'<rect x="{x + border}" y="{y + border}" width="1" height="1"/>'
                    )
        svg = (
            f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {size} {size}" '
            f'shape-rendering="crispEdges"><rect width="{size}" height="{size}" fill="#ffffff"/>'
            f'<g fill="#0e3f3c">{"".join(rects)}</g></svg>'
        )
        (QR_DIR / f"{name}.svg").write_text(svg, encoding="utf-8")


def build_pdf() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    PREVIEW.mkdir(parents=True, exist_ok=True)
    if PDF.exists():
        PDF.unlink()
    for old in PREVIEW.glob("*.png"):
        old.unlink()
    # 不用 PATH 里的 google-chrome 包装脚本，它会强行打开调试端口并卡住不退出。
    chrome = "/opt/google/chrome/chrome"
    profile = Path("/tmp/chrome-a4-guide")
    if profile.exists():
        shutil.rmtree(profile)
    profile.mkdir(parents=True)
    cmd = [
        chrome,
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--disable-extensions",
        "--disable-background-networking",
        "--no-first-run",
        "--no-default-browser-check",
        f"--user-data-dir={profile}",
        "--allow-file-access-from-files",
        "--no-pdf-header-footer",
        "--virtual-time-budget=10000",
        f"--print-to-pdf={PDF}",
        HTML.as_uri(),
    ]
    subprocess.run(cmd, check=True, timeout=60)
    if not PDF.exists() or PDF.stat().st_size < 10000:
        raise SystemExit("PDF 没有生成或文件过小")
    subprocess.run(
        ["pdftoppm", "-png", "-r", "140", str(PDF), str(PREVIEW / "page")],
        check=True,
    )


def check_pdf() -> None:
    info = subprocess.run(["pdfinfo", str(PDF)], check=True, capture_output=True, text=True)
    print(info.stdout)
    text = subprocess.run(
        ["pdftotext", "-layout", str(PDF), "-"],
        check=True,
        capture_output=True,
        text=True,
    ).stdout
    missing = [item for item in REQUIRED_TEXT if item not in text]
    pages = sorted(PREVIEW.glob("page-*.png"))
    if len(pages) != 13:
        raise SystemExit(f"页数应为 13，实际 {len(pages)}")
    if not re.search(r"Pages:\s+13\b", info.stdout):
        raise SystemExit("pdfinfo 的页数不是 13")
    size = re.search(r"Page size:\s+([\d.]+) x ([\d.]+) pts", info.stdout)
    if not size:
        raise SystemExit("pdfinfo 没有页面尺寸")
    width, height = float(size.group(1)), float(size.group(2))
    # Chrome 会把 A4 记成约 594.96 x 841.92 点。
    if not (590 <= width <= 600 and 836 <= height <= 848):
        raise SystemExit(f"页面不是 A4：{width} x {height}")
    if missing:
        raise SystemExit("PDF 文本缺少：\n" + "\n".join(missing))
    print(f"已生成 {PDF}")
    print(f"预览 {len(pages)} 页")


def main() -> None:
    write_qr_codes()
    build_pdf()
    check_pdf()


if __name__ == "__main__":
    try:
        main()
    except subprocess.CalledProcessError as exc:
        print(exc, file=sys.stderr)
        raise SystemExit(exc.returncode) from exc
