# -*- coding: utf-8 -*-
from pathlib import Path

from content import EVENT

OUT = Path(__file__).resolve().parents[1] / "exports" / "飞书DemoDay4.ics"


def fmt(dt):
    return dt.strftime("%Y%m%dT%H%M%S")


def build():
    stamp = EVENT["start"]
    ics = f"""BEGIN:VCALENDAR
PRODID:-//AIDEN//Feishu Demo Day 4//ZH
VERSION:2.0
CALSCALE:GREGORIAN
METHOD:PUBLISH
BEGIN:VTIMEZONE
TZID:Asia/Shanghai
X-LIC-LOCATION:Asia/Shanghai
BEGIN:STANDARD
TZOFFSETFROM:+0800
TZOFFSETTO:+0800
TZNAME:CST
DTSTART:19700101T000000
END:STANDARD
END:VTIMEZONE
BEGIN:VEVENT
UID:feishu-demo-day-4-{EVENT["token"]}@aiden
DTSTAMP:{fmt(stamp)}Z
DTSTART;TZID=Asia/Shanghai:{fmt(EVENT["start"])}
DTEND;TZID=Asia/Shanghai:{fmt(EVENT["end"])}
SUMMARY:{EVENT["title"]}
DESCRIPTION:飞书 AI Builder Demo Day #4。预约链接：{EVENT["share_url"]} 官方日历状态以页面为准。
LOCATION:飞书线上会议
URL:{EVENT["share_url"]}
STATUS:CONFIRMED
BEGIN:VALARM
TRIGGER:-PT30M
ACTION:DISPLAY
DESCRIPTION:Demo Day #4 30 分钟后开始
END:VALARM
BEGIN:VALARM
TRIGGER:-PT5M
ACTION:DISPLAY
DESCRIPTION:Demo Day #4 5 分钟后开始，准备入会
END:VALARM
END:VEVENT
END:VCALENDAR
"""
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(ics.replace("\n", "\r\n"), encoding="utf-8")
    print(f"wrote {OUT}")
    return OUT


if __name__ == "__main__":
    build()
