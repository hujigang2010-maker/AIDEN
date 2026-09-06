#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""校验修订稿：字数上限、产品主线、脱敏、正式稿无内部话术。"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import content as C  # noqa: E402
from generate_all import DOCS  # noqa: E402

OUT = Path(__file__).resolve().parents[1] / "output"

PII = ["爱国路389", "吉林磐石"]
FORMAL_FORBIDDEN = [
    "三条诚实边界",
    "不是咨询公司翻牌",
    "防止身份写串",
    "写高了就变成虚构",
    "不要选「暂无法估算」",
    "不虚构客户收入",
    "示范智能体可演示",
    "控制在50万元以内",
    "控制全年上限",
    "上海市工商联房地产商会秘书长",
]
FORMAL_REQUIRED_BP = [
    "城市更新案例库",
    "RAG",
    "企业画像",
    "正在搭建",
    "预计首年营业收入低于50万元",
    "技术开发",
    "服务收费",
    ENTITY := C.ENTITY_LINE[:18],
]


def fail(msg):
    print(f"失败  {msg}")
    return 1


def ok(msg):
    print(f"通过  {msg}")
    return 0


def main() -> int:
    errors = 0
    if list(OUT.glob("星火社区_00_*")):
        errors += fail("仍存在00全套上传件")
    else:
        ok("已取消00全套上传件")

    for stem, *_ in DOCS:
        for ext in (".docx", ".pdf", ".html"):
            path = OUT / f"星火社区_{stem}{ext}"
            if not path.exists() or path.stat().st_size < 2000:
                errors += fail(f"缺失或过小：{path.name}")
            else:
                ok(f"{path.name}  {path.stat().st_size} bytes")
        path = OUT / f"星火社区_{stem}.docx"
        try:
            with zipfile.ZipFile(path) as z:
                assert "word/document.xml" in z.namelist()
            ok(f"{path.name} 为有效 docx")
        except Exception as e:
            errors += fail(f"{path.name} {e}")

    md = OUT / "星火社区_04_申请表可粘贴字段.md"
    if not md.exists():
        errors += fail("缺字段 Markdown")
    else:
        ok(f"{md.name}  {md.stat().st_size} bytes")

    bp = (OUT / "星火社区_03_业务计划书.html").read_text(encoding="utf-8")
    bio = (OUT / "星火社区_02_个人简介.html").read_text(encoding="utf-8")
    form = (OUT / "星火社区_04_申请表可粘贴字段.html").read_text(encoding="utf-8")
    guide = (OUT / "星火社区_01_填写说明与口径备忘.html").read_text(encoding="utf-8")
    all_text = bp + bio + form + guide + md.read_text(encoding="utf-8")

    for phrase in PII:
        if phrase in all_text or phrase in Path("/workspace/scripts/content.py").read_text():
            errors += fail(f"仍含敏感信息：{phrase}")
        else:
            ok(f"已脱敏：{phrase}")

    for phrase in FORMAL_FORBIDDEN:
        if phrase in bp or phrase in bio:
            errors += fail(f"正式稿出现内部话术：{phrase}")
        else:
            ok(f"正式稿无：{phrase}")

    for phrase in FORMAL_REQUIRED_BP:
        if phrase not in bp:
            errors += fail(f"计划书缺少：{phrase}")
        else:
            ok(f"计划书含：{phrase}")

    if "盈利模式" not in form or "技术开发" not in form:
        errors += fail("字段稿缺少盈利模式勾选")
    else:
        ok("字段稿含盈利模式：技术开发、服务收费")
    if "核心产品" not in form:
        errors += fail("字段稿缺少核心产品")
    else:
        ok("字段稿含核心产品／服务")

    if "吉林磐石" in bio or "爱国路" in bio:
        errors += fail("个人简介仍含户籍或门牌")
    else:
        ok("个人简介无吉林户籍、无详细门牌")

    limits = [
        ("PRODUCT_100", C.PRODUCT_100, 100),
        ("INTRO_300", C.INTRO_300, 300),
        ("BIO_200", C.BIO_200, 200),
        ("TEAM_150", C.TEAM_150, 150),
    ]
    for key, text, lim in limits:
        n = C.char_count(text)
        if n > lim:
            errors += fail(f"{key} {n} 超过上限 {lim}")
        else:
            ok(f"{key} {n}／{lim}")

    if C.PROJECT["是否已注册"] != "否" or C.PROJECT["团队人数"] != "1":
        errors += fail("注册或团队口径被改")
    else:
        ok("未注册、团队1人")

    tracked_text = all_text + Path("/workspace/scripts/content.py").read_text(encoding="utf-8")
    if re.search(r"\d{17}[\dXx]", tracked_text):
        errors += fail("口径或输出中疑似身份证号")
    else:
        ok("口径与输出无身份证号模式")

    print("-" * 40)
    if errors:
        print(f"合计失败 {errors} 项")
        return 1
    print("全部校验通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
