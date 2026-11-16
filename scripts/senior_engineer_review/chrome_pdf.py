"""用无头 Chrome 把 HTML 打印成 PDF。"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path


def find_chrome() -> str:
    # google-chrome 在部分环境里是带图形界面参数的包装脚本，放在最后兜底。
    for name in ("google-chrome-stable", "chromium", "chromium-browser", "google-chrome"):
        path = shutil.which(name)
        if path:
            return path
    raise RuntimeError("找不到 Chrome / Chromium，无法排版")


def print_to_pdf(html_text: str, out_pdf: Path) -> None:
    with tempfile.TemporaryDirectory() as tmp:
        html_path = Path(tmp) / "page.html"
        html_path.write_text(html_text, encoding="utf-8")
        cmd = [
            find_chrome(),
            "--headless=new",
            "--no-sandbox",
            "--disable-gpu",
            "--disable-dev-shm-usage",
            "--allow-file-access-from-files",
            "--no-pdf-header-footer",
            f"--user-data-dir={tmp}/profile",
            f"--print-to-pdf={out_pdf}",
            html_path.as_uri(),
        ]
        subprocess.run(cmd, check=True, capture_output=True, timeout=180)
    if not out_pdf.exists():
        raise RuntimeError("Chrome 没有生成 PDF")
