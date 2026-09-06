#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""校验最终稿：字数、产品主线、脱敏、正式语气、口径一致。"""

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
    "不倒签合同",
    "示范智能体可演示",
    "控制在50万元以内",
    "控制全年上限",
    "上海市工商联房地产商会秘书长",
    "长文本备份",
]
FORMAL_REQUIRED_BP = [
    "政策研究智能体",
    "产业研究智能体",
    "企业画像与招商匹配智能体",
    "城市更新案例库智能体",
    "RAG",
    "已开展AI辅助研究实践",
    "正在将",
    "预计首个完整经营年度营业收入低于50万元",
    "服务收费",
    "技术开发",
    "拟申请",
    "以实际获批为准",
    "1名创始人负责运营，初期不另聘全职员工",
    C.ENTITY_LINE[:18],
]


def fail(msg):
    print(f"失败  {msg}")
    return 1


def ok(msg):
    print(f"通过  {msg}")
    return 0


def main() -> int:
    errors = 0
    if list(OUT.glob("星火社区_00_全套材料*")):
        errors += fail("仍存在旧版00全套材料")
    else:
        ok("无旧版00全套材料")

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
    pack = (OUT / "星火社区_00_正式稿合订本.html").read_text(encoding="utf-8")
    md_text = md.read_text(encoding="utf-8")
    all_text = bp + bio + form + guide + pack + md_text
    src = Path("/workspace/scripts/content.py").read_text(encoding="utf-8")

    for phrase in PII:
        if phrase in all_text or phrase in src:
            errors += fail(f"仍含敏感信息：{phrase}")
        else:
            ok(f"已脱敏：{phrase}")

    for phrase in FORMAL_FORBIDDEN:
        if phrase in bp or phrase in bio or phrase in pack:
            errors += fail(f"正式稿出现内部话术：{phrase}")
        else:
            ok(f"正式稿无：{phrase}")

    if "备选稿" in form or "备选稿" in md_text:
        errors += fail("字段稿仍含备选稿")
    else:
        ok("字段稿每字段仅一版")

    if "建议填写" in pack or "长文本最终稿" in pack:
        errors += fail("合订本混入填表稿")
    else:
        ok("合订本只含正式正文结构")

    for phrase in FORMAL_REQUIRED_BP:
        if phrase not in bp:
            errors += fail(f"计划书缺少：{phrase}")
        else:
            ok(f"计划书含：{phrase}")

    if C.ENTITY_LINE not in bio or C.ENTITY_LINE not in bp:
        errors += fail("主体说明未按最终口径写入简介或计划书")
    else:
        ok("主体说明已写入简介与计划书")

    if C.PROJECT["首期原型"] != "政策研究智能体":
        errors += fail("首期产品不是政策研究智能体")
    else:
        ok("首期=政策研究智能体")

    if "盈利模式" not in form or "服务收费" not in form:
        errors += fail("字段稿缺少盈利模式")
    else:
        ok("字段稿含盈利模式：服务收费、技术开发")
    if "核心产品" not in form:
        errors += fail("字段稿缺少核心产品")
    else:
        ok("字段稿含核心产品／服务")
    if "与截图对应" not in form:
        errors += fail("字段稿缺少截图对应区块")
    else:
        ok("字段稿含截图对应最终答案")

    if "吉林磐石" in bio or "爱国路" in bio:
        errors += fail("个人简介仍含户籍或门牌")
    else:
        ok("个人简介无吉林户籍、无详细门牌")

    limits = [
        ("项目名称", C.PROJECT["项目名称"], 50, 1),
        ("PRODUCT_100", C.PRODUCT_100, 100, 1),
        ("INTRO_300", C.INTRO_300, 300, 100),
        ("BIO_200", C.BIO_200, 200, 1),
        ("TEAM_200", C.TEAM_200, 200, 1),
    ]
    for key, text, hi, lo in limits:
        n = C.char_count(text)
        if n > hi or n < lo:
            errors += fail(f"{key} {n} 不在 {lo}–{hi}")
        else:
            ok(f"{key} {n}／上限{hi}")

    if C.PROJECT["是否已注册"] != "否" or C.PROJECT["团队人数"] != "1":
        errors += fail("注册或团队口径被改")
    else:
        ok("未注册、团队1人")

    if C.PROJECT["项目名称"] != "AIDEN产业智能体":
        errors += fail("项目名称被改")
    else:
        ok("项目名称一致")

    tracked_text = all_text + src
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
