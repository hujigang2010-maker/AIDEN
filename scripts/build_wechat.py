# -*- coding: utf-8 -*-
"""生成可直接粘贴到微信群的课题短稿。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import content as C

OUT = Path(__file__).resolve().parents[1] / "exports"
OUT.mkdir(parents=True, exist_ok=True)
OUT_MD = OUT / "微信群发布稿.md"
OUT_TXT = OUT / "微信群发布稿.txt"


def render_lines():
    lines = [
        f"【{C.TITLE}】",
        C.SUBTITLE,
        "",
        f"计划 {C.RELEASE_DATE} 正式发出。下面十条都是已经开工的真实课题，不是为发榜临时编的。",
        "经过初筛的同学：每人报 1 个第一志愿 + 1 个备选，写清你能贡献什么。老师编组后拉项目组。",
        "",
        "—— 怎么选 ——",
        "A 软件与数据：T01 排版器 / T02 小程序 / T03 楼宇数据 / T04 报名协议",
        "B AI 与智能体：T05 知识工作室 / T06 文档工程 / T09 工作流主线",
        "C 产业研究：T07 具身智能与产业空间 / T08 住房学区数据 / T10 白皮书工作坊",
        "",
        "—— 十条课题（一句话） ——",
    ]
    for t in C.TOPICS:
        lines.append(f"{t['id']}  {t['name']}")
        lines.append(f"状态：{t['status']}  ·  {t['quota']}  ·  {t['weeks']}")
        lines.append(t["one_liner"])
        lines.append("")
    lines.extend(
        [
            "—— 约定 ——",
            "进组是做事，不开空证明。远程为主，上海可短期驻场。每周建议不少于 8 小时，有周报。",
            "导师给方向和验收，不代写。密钥、内部名单、个人隐私不进仓库。",
            "",
            "—— 请回 ——",
            "姓名 + 学院专业年级 + 第一志愿 + 备选 + 每周小时 + 能否赴沪 + 80 字以内「我能贡献什么」",
            "详细任务见 Word，宣讲见 PPT，填表见 Excel。",
            "",
            C.WECHAT_CONTACT_NOTE,
            f"{C.ORG_LINE}",
            f"对接：{C.MENTOR}  ·  {C.DATE_STR}  ·  {C.VERSION}",
        ]
    )
    return lines


def render_text():
    return "\n".join(render_lines()) + "\n"


def build():
    text = render_text()
    OUT_TXT.write_text(text, encoding="utf-8")
    OUT_MD.write_text(text, encoding="utf-8")
    print(f"已生成 {OUT_TXT}")
    print(f"已生成 {OUT_MD}")
    return OUT_TXT


if __name__ == "__main__":
    build()
