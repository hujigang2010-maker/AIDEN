# -*- coding: utf-8 -*-
from pathlib import Path

from content import DISCLAIMER, EVENT, FORWARD_LONG, FORWARD_REPLAY, FORWARD_SHORT

ROOT = Path(__file__).resolve().parents[1] / "exports"


def build():
    ROOT.mkdir(parents=True, exist_ok=True)
    (ROOT / "发群一页稿_短.txt").write_text(FORWARD_SHORT.strip() + "\n", encoding="utf-8")
    (ROOT / "发群一页稿_完整.txt").write_text(FORWARD_LONG.strip() + "\n", encoding="utf-8")
    (ROOT / "发群一页稿_回放.txt").write_text(FORWARD_REPLAY.strip() + "\n", encoding="utf-8")
    md = f"""# 飞书 Demo Day #4 发群一页稿

> {DISCLAIMER}

## 短口径（群公告）

```
{FORWARD_SHORT.strip()}
```

## 完整口径（朋友圈 / 长公告）

```
{FORWARD_LONG.strip()}
```

## 回放口径（日历显示已结束）

```
{FORWARD_REPLAY.strip()}
```

## 官方入口

- 分享日历：{EVENT["share_url"]}
- 场次时间（日历）：2026-07-02 11:00–12:30 GMT+8
- 组织者：{EVENT["organizer"]}
- 状态：{EVENT["status_zh"]}
- 往期回放检索：{EVENT["bilibili_search"]}
"""
    (ROOT / "发群一页稿.md").write_text(md, encoding="utf-8")
    print("wrote wechat copy")


if __name__ == "__main__":
    build()
