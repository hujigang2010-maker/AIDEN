# -*- coding: utf-8 -*-
from pathlib import Path

from docx import Document
from openpyxl import load_workbook
from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1] / "exports"

REQUIRED = [
    "飞书DemoDay4_观摩速览.pptx",
    "飞书DemoDay4_观摩备忘录.docx",
    "飞书DemoDay4_观摩记录表.xlsx",
    "飞书DemoDay4_转发海报.png",
    "飞书DemoDay4.ics",
    "发群一页稿.md",
    "发群一页稿_短.txt",
    "发群一页稿_完整.txt",
    "发群一页稿_回放.txt",
]


def main():
    missing = [name for name in REQUIRED if not (ROOT / name).exists()]
    assert not missing, f"缺少文件: {missing}"

    prs = Presentation(ROOT / "飞书DemoDay4_观摩速览.pptx")
    assert len(prs.slides) == 14, len(prs.slides)
    assert abs(prs.slide_width - 12191695) < 100

    doc = Document(ROOT / "飞书DemoDay4_观摩备忘录.docx")
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "Agent 入会" in text
    assert "236d1f15f1294d960e52362d10b7503c" in text
    assert "张咋啦" in text and "向阳乔木" in text

    wb = load_workbook(ROOT / "飞书DemoDay4_观摩记录表.xlsx")
    assert set(wb.sheetnames) == {"活动信息", "观摩记录", "Agent入会清单", "转发口径"}

    ics = (ROOT / "飞书DemoDay4.ics").read_text(encoding="utf-8")
    assert "DTSTART;TZID=Asia/Shanghai:20260702T110000" in ics
    assert "DTEND;TZID=Asia/Shanghai:20260702T123000" in ics

    poster = ROOT / "飞书DemoDay4_转发海报.png"
    assert poster.stat().st_size > 20_000

    print("verify ok")


if __name__ == "__main__":
    main()
