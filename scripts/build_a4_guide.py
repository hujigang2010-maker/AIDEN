#!/usr/bin/env python3
"""把针灸科现场手册和开讲前问卷打印为 A4 PDF，并导出分页预览。

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
SURVEY_HTML = ROOT / "guide" / "survey.html"
QR_DIR = ROOT / "guide" / "qr"
OUT = ROOT / "output"
PDF = OUT / "岳阳医院针灸科-VPN及国外AI软件安装使用指南.pdf"
PREVIEW = OUT / "preview"
SURVEY_PDF = OUT / "岳阳医院针灸科-开讲前AI需求问卷.pdf"
SURVEY_PREVIEW = OUT / "preview-survey"

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


def print_pdf(html: Path, pdf: Path, preview: Path, profile_name: str) -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    preview.mkdir(parents=True, exist_ok=True)
    if pdf.exists():
        pdf.unlink()
    for old in preview.glob("*.png"):
        old.unlink()
    # 不用 PATH 里的 google-chrome 包装脚本，它会强行打开调试端口并卡住不退出。
    chrome = "/opt/google/chrome/chrome"
    profile = Path("/tmp") / profile_name
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
        f"--print-to-pdf={pdf}",
        html.as_uri(),
    ]
    subprocess.run(cmd, check=True, timeout=60)
    if not pdf.exists() or pdf.stat().st_size < 10000:
        raise SystemExit(f"PDF 没有生成或文件过小：{pdf.name}")
    subprocess.run(
        ["pdftoppm", "-png", "-r", "140", str(pdf), str(preview / "page")],
        check=True,
    )


def assert_a4(info: str, pdf_name: str) -> None:
    size = re.search(r"Page size:\s+([\d.]+) x ([\d.]+) pts", info)
    if not size:
        raise SystemExit(f"{pdf_name} 没有页面尺寸")
    width, height = float(size.group(1)), float(size.group(2))
    # Chrome 会把 A4 记成约 594.96 x 841.92 点。
    if not (590 <= width <= 600 and 836 <= height <= 848):
        raise SystemExit(f"{pdf_name} 不是 A4：{width} x {height}")


def check_document(pdf: Path, preview: Path, expected_pages: int, required: list[str]) -> None:
    info = subprocess.run(["pdfinfo", str(pdf)], check=True, capture_output=True, text=True).stdout
    print(info)
    text = subprocess.run(
        ["pdftotext", "-layout", str(pdf), "-"],
        check=True,
        capture_output=True,
        text=True,
    ).stdout
    missing = [item for item in required if item not in text]
    pages = sorted(preview.glob("page-*.png"))
    if len(pages) != expected_pages:
        raise SystemExit(f"{pdf.name} 页数应为 {expected_pages}，实际 {len(pages)}")
    if not re.search(rf"Pages:\s+{expected_pages}\b", info):
        raise SystemExit(f"{pdf.name} 的 pdfinfo 页数不是 {expected_pages}")
    assert_a4(info, pdf.name)
    if missing:
        raise SystemExit(f"{pdf.name} 文本缺少：\n" + "\n".join(missing))
    print(f"已生成 {pdf}")
    print(f"预览 {len(pages)} 页")


def main() -> None:
    write_qr_codes()
    print_pdf(HTML, PDF, PREVIEW, "chrome-a4-guide")
    check_document(PDF, PREVIEW, 13, REQUIRED_TEXT)
    print_pdf(SURVEY_HTML, SURVEY_PDF, SURVEY_PREVIEW, "chrome-a4-survey")
    check_document(
        SURVEY_PDF,
        SURVEY_PREVIEW,
        2,
        [
            "使用 AI 的需求小问卷",
            "不要填写患者信息",
            "只选一项",
            "查文献，并核对原文",
            "研究设计",
            "Zotero",
            "讲者统计",
            "不要发给医生",
            "膝骨关节炎",
        ],
    )


if __name__ == "__main__":
    try:
        main()
    except subprocess.CalledProcessError as exc:
        print(exc, file=sys.stderr)
        raise SystemExit(exc.returncode) from exc
