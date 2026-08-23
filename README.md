# 飞书 AI Builder Demo Day #4 观摩与转发材料

把用户转发的第四场直播预告，收成一套能发群、能预习、能现场记、能会后复盘的材料。

## 已核对事实

| 项 | 内容 |
| --- | --- |
| 系列 | 飞书 AI Builder Demo Day |
| 官方日历标题 | Demo Day #4 让 AI 接住你的日常小麻烦 2 |
| 时间 | 2026-07-02（周四）11:00–12:30 GMT+8 |
| 组织者 | Ni Dan |
| 状态 | 分享日历显示已结束 |
| 预约链接 | https://bytedance.larkoffice.com/calendar/share?token=236d1f15f1294d960e52362d10b7503c |

嘉宾（张咋啦、向阳乔木）和「Agent 入会将现场展示」来自用户预告，不是日历详情里的官方议程。

## 交付物

| 文件 | 用途 |
| --- | --- |
| `exports/飞书DemoDay4_观摩速览.pptx` | 14 页内部速览，适合会前 10 分钟过一遍 |
| `exports/飞书DemoDay4_观摩备忘录.docx` | 完整备忘：事实、嘉宾看点、Agent 对照、转发口径 |
| `exports/飞书DemoDay4_观摩记录表.xlsx` | 活动信息 / 观摩记录 / Agent 入会清单 / 转发口径 |
| `exports/飞书DemoDay4_转发海报.png` | 1080×1440 发群海报 |
| `exports/飞书DemoDay4.ics` | 日历文件，含开场前 30 分钟和 5 分钟提醒 |
| `exports/发群一页稿.md` | 短口径、完整口径、回放口径 |

## 重新生成

```bash
python3 -m pip install python-pptx python-docx openpyxl pillow
python3 scripts/build_all.py
python3 scripts/verify.py
```

## 口径不要改偏

- 先加会议。上一场爆满。
- 官方时间以日历为准；若页面显示已结束，用同一链接看回放或加场。
- 不要把预告嘉宾写成官方确认赞助。
- Agent 入会是本场最大看点，但公开 CLI 能力说明不能冒充现场议程。
