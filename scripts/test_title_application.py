#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""检查申报表排版稿是否保留原表栏目、照片和关键文字。"""

from generate_title_application import build, self_check


def main() -> None:
    path = build()
    self_check(path)
    print("测试通过", path)


if __name__ == "__main__":
    main()
