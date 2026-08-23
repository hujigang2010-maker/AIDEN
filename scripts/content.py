# -*- coding: utf-8 -*-
"""飞书 AI Builder Demo Day #4 观摩包口径。只写已核对事实，不编造议程。"""

from datetime import datetime
from zoneinfo import ZoneInfo

TZ = ZoneInfo("Asia/Shanghai")

EVENT = {
    "series": "飞书 AI Builder Demo Day",
    "episode": 4,
    "title": "Demo Day #4 让 AI 接住你的日常小麻烦",
    "title_share": "Demo Day #4 让 AI 接住你的日常小麻烦 2",
    "theme": "让 AI 接住你的日常小麻烦",
    "start": datetime(2026, 7, 2, 11, 0, tzinfo=TZ),
    "end": datetime(2026, 7, 2, 12, 30, tzinfo=TZ),
    "weekday": "周四",
    "duration": "90 分钟",
    "organizer": "Ni Dan",
    "format": "线上直播（飞书日历分享入会）",
    "status": "finished",
    "status_zh": "官方分享日历显示：已结束",
    "share_url": (
        "https://bytedance.larkoffice.com/calendar/share"
        "?token=236d1f15f1294d960e52362d10b7503c"
    ),
    "token": "236d1f15f1294d960e52362d10b7503c",
    "bilibili_account": "飞书（哔哩哔哩）",
    "bilibili_search": "https://search.bilibili.com/all?keyword=AI%20Builder%20Demo%20Day",
    "replay_ep2": {
        "title": "AI Builder Demo Day 全程回放",
        "bvid": "BV176EX6hEdA",
        "url": "https://www.bilibili.com/video/BV176EX6hEdA",
        "note": "飞书官方账号；简介写明为第二期",
    },
}

GUESTS = [
    {
        "name": "张咋啦",
        "aka": "Zara Zhang",
        "role": "现场 Demo 嘉宾（用户预告口径）",
        "why": (
            "中美 AI Builder 圈高频出现的实践者，公开分享过把飞书文档当 Agent brief、"
            "用 vibe coding 出 HTML 演示、以及把 coding agent 接到飞书里干活。"
        ),
        "watch": [
            "日常麻烦是哪一件，而不是又一个玩具 Demo",
            "人给判断、Agent 给执行，边界画在哪里",
            "飞书文档 / CLI / 群，分别承担什么角色",
        ],
    },
    {
        "name": "向阳乔木",
        "aka": "乔木",
        "role": "现场 Demo 嘉宾（用户预告口径）",
        "why": (
            "长期做 AI 内容与落地项目分享，近期公开组织过 Vibe Coding / FDE 一线项目直播，"
            "风格偏「只讲干货、去废话」。"
        ),
        "watch": [
            "Demo 是否能在自己的工作流里复现",
            "FDE / 交付现场踩过哪些坑",
            "工具链有没有可抄的最小闭环",
        ],
    },
]

AGENT_JOIN = {
    "name": "Agent 入会",
    "source": "用户预告：本场将按 Demo Day 惯例现场展示",
    "what": (
        "让应用机器人真实加入一场进行中的飞书视频会议："
        "出现在参会列表、拉取会中事件、必要时发会中消息，会后仍可走纪要 / 逐字稿。"
    ),
    "official_skill": "飞书 CLI 技能 lark-vc-agent（会中动作）",
    "join_cmd": "lark-cli vc +meeting-join --as bot --meeting-number <9位会议号>",
    "can": [
        "机器人真实入会（会留下入会记录）",
        "读取进行中会议的实时事件：进出、发言、聊天、屏幕共享",
        "发送会中文本或表情",
        "会后用 lark-vc / 妙记拉纪要、逐字稿、录制",
    ],
    "cannot": [
        "加入尚未开始或已结束的会议",
        "在等候室 / 入会审批未放行时直接现身",
        "把 9 位会议号当成 meeting_id 去拉事件",
    ],
    "watch_questions": [
        "Agent 是「旁听记笔记」还是「会中开口」？",
        "主持人要不要审批？外部 Agent 能不能进？",
        "转写延迟大概几秒，能不能跟上讨论？",
        "会中消息会不会打断真人节奏？",
        "纪要是会中实时出，还是会后才可用？",
        "哪些会绝对不该让 Agent 进（人事、融资、客户机密）？",
    ],
}

WHY_WATCH = [
    "系列定位是 Builder 现场演示，不是发布会 PPT。",
    "本场主题对准「日常小麻烦」，比空中楼阁更可抄。",
    "张咋啦、向阳乔木的 Demo 预告质量高，值得对照自己的工作流。",
    "Agent 入会是新能力：AI 从会后总结，变成会中同事。",
    "上一场容易约满，日历入口要先占上。",
]

SERIES_NOTES = [
    "飞书官方在哔哩哔哩持续上传 AI Builder Demo Day 回放。",
    "已核到第二期回放 BV176EX6hEdA；第三期在官方账号相关推荐里可见「Demo Day 3 全程回放」。",
    "第四场官方日历标题带「2」，更像同一主题的加场 / 重场，而不是另一套议题。",
    "飞行社活动页同期还有「豆包工作伙伴新品演示会」等，不要和 Demo Day 混成一场。",
]

WATCH_FLOW = {
    "会前": [
        "点开分享日历，确认能否加入；约满就换账号 / 等加场。",
        "准备一个自己的「日常小麻烦」样本（周会、纪要、待办、客户会）。",
        "若要试 Agent 入会：先准备测试会，不要拿正式客户会开刀。",
        "打开观摩记录表，只记可复现步骤，不记气氛。",
    ],
    "会中": [
        "每个 Demo 只抓三件事：麻烦是什么、人做哪一步、Agent 做哪一步。",
        "Agent 入会段落：截图入会瞬间、权限提示、会中发言、纪要出口。",
        "把「我做不到」和「我不想做」分开写。",
    ],
    "会后 24h": [
        "用自己的测试会复现最小闭环，不要直接上生产。",
        "把可抄步骤写进飞书文档，当成下一份 Agent brief。",
        "回放补漏：官方 B 站搜索 AI Builder Demo Day。",
    ],
}

FORWARD_LONG = """飞书的 Demo Day 第四场直播来了。

这场 Demo 质量看着依旧很高，有张咋啦、向阳乔木的演示。还有 Agent 入会这个新功能——按 Demo Day 惯例大概率现场展示。AI 帮你开会，会非常有意思。

建议先加会议。上一场爆满，后面根本约不上。

主题：让 AI 接住你的日常小麻烦
官方日历场次：2026-07-02（周四）11:00–12:30（GMT+8）
预约 / 入会：https://bytedance.larkoffice.com/calendar/share?token=236d1f15f1294d960e52362d10b7503c

先占入口，再决定看直播还是看回放。"""

FORWARD_SHORT = """飞书 AI Builder Demo Day #4｜让 AI 接住你的日常小麻烦
张咋啦、向阳乔木 Demo + Agent 入会现场展示（预告）
先加日历，上一场约满就进不去：
https://bytedance.larkoffice.com/calendar/share?token=236d1f15f1294d960e52362d10b7503c"""

FORWARD_REPLAY = """飞书 AI Builder Demo Day #4《让 AI 接住你的日常小麻烦》

官方分享日历显示这场（2026-07-02 11:00–12:30）已结束。如果没赶上直播：

1. 仍可点原日历链接，看是否放出回放 / 加场。
2. 去哔哩哔哩搜「AI Builder Demo Day」，看飞书官方账号往期全程回放。
3. 本场预告亮点：张咋啦、向阳乔木 Demo，以及 Agent 入会（AI 进会议室）。

日历：https://bytedance.larkoffice.com/calendar/share?token=236d1f15f1294d960e52362d10b7503c
往期回放入口：https://search.bilibili.com/all?keyword=AI%20Builder%20Demo%20Day"""

DISCLAIMER = (
    "嘉宾名单与「Agent 入会将现场展示」来自用户转发预告，不是飞书日历详情里的官方议程。"
    "日历接口只确认了标题、时间、组织者、已结束状态和分享链接。"
    "不要把观摩提纲写成官方日程表。"
)
