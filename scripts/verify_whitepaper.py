# -*- coding: utf-8 -*-
"""校验白皮书 DOCX 是否生成完整。"""

from pathlib import Path

from docx import Document

ROOT = Path(__file__).resolve().parents[1]
DOCX = ROOT / "dist" / "复旦大学住房政策研究中心_2026年具身智能机器人白皮书.docx"

REQUIRED_HEADINGS = [
    "2026 年具身智能机器人白皮书",
    "复旦大学住房政策研究中心",
    "摘要：2026 年的五个判断与 2030 年的一条主线",
    "一、导论：住房政策为何必须正视具身智能",
    "三、WAIC 2026 深度观察：五个结构性趋势",
    "四、全球与国内展会扫描：一年之内的现场证据",
    "六、住房、社区与城市：当机器人成为“住户”",
    "七、展望 2030：四个情景与关键变量",
    "八、政策建议",
    "附录二  主要公开资料来源",
]

REQUIRED_KEYWORDS = [
    "WAIC",
    "CES 2026",
    "汉诺威",
    "亦庄",
    "世界机器人大会",
    "适老化",
    "机器人可达性",
    "彭志辉",
    "具身智能",
    "2030",
]


def main():
    assert DOCX.exists(), f"未找到文件：{DOCX}"
    size_kb = DOCX.stat().st_size / 1024
    assert size_kb > 40, f"文件过小：{size_kb:.1f} KB"

    doc = Document(str(DOCX))
    texts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    blob = "\n".join(texts)
    char_count = len(blob.replace(" ", "").replace("\n", ""))

    missing_h = [h for h in REQUIRED_HEADINGS if h not in blob]
    missing_k = [k for k in REQUIRED_KEYWORDS if k not in blob]

    table_count = len(doc.tables)
    print(f"文件：{DOCX.name}")
    print(f"大小：{size_kb:.1f} KB")
    print(f"段落数：{len(doc.paragraphs)}")
    print(f"非空段落：{len(texts)}")
    print(f"正文字符数（约）：{char_count}")
    print(f"表格数：{table_count}")

    assert not missing_h, f"缺少标题：{missing_h}"
    assert not missing_k, f"缺少关键词：{missing_k}"
    assert char_count >= 12000, f"正文过短：{char_count}"
    assert table_count >= 3, f"表格过少：{table_count}"
    print("校验通过。")


if __name__ == "__main__":
    main()
