# 生成脚本数据 Schema

与 `build_volunteer_assignment.py` 对齐的约定。新建活动时可复制脚本后改数据块。

## 活动与议程

```python
EVENT_TITLE = "..."
EVENT_SUB = "..."
EVENT_DATE = "..."
EVENT_VENUE = "..."

AGENDA = [
    ("13:00-13:30", "嘉宾签到与入场", "C/B/D", "签到处、引导动线"),
    # (时段, 环节, 责任组字母, 关键动作)
]
```

## 组别

```python
GROUPS = [
    {
        "code": "A",
        "name": "统筹协调组",
        "leader": "姓名（角色）",
        "size": "3 人",
        "color": "1F4E79",  # 六位 hex，用于 Excel/PPT
        "duty_summary": "一句话",
        "sub_duties": [
            ("岗位名", "职责描述"),
        ],
    },
]
```

## 人员分配

```python
ASSIGNMENT = [
    {
        "name": "张三",
        "group": "C",           # A–G
        "role": "颁奖主导",
        "detail": "流程卡/奖杯/引导嘉宾上台",
        "backup": "李四",       # 可选
        "arrive": "11:30",      # 可选，可与 ARRIVAL_TIMES 同步
        "note": "5/18 决议从 B 调入",
    },
]
```

## 建议筛选

```python
OPT_SUGGESTIONS = [
    {
        "id": "S01",
        "source": "5/13 讨论会",
        "text": "建议增加对讲机至 10 台",
        "status": "采纳",       # 采纳 | 调整 | 暂缓
        "action": "已写入 MATERIALS，F 组负责领取",
    },
]
```

## 应急

```python
CONTINGENCIES = [
    {
        "scene": "主旨嘉宾迟到 >15 分钟",
        "level": "高",
        "owner": "A/C",
        "steps": "通知主持拉长暖场 → 调整下一项 → 接待组确认 ETA",
    },
]
```

## 其他常用表

| 变量 | 用途 |
|------|------|
| `BACKUP_ROLES` | 关键岗主/备 |
| `MILESTONES` | 倒排节点 |
| `MATERIALS` | 物资品类、数量、责任人 |
| `EXTERNAL_CONTACTS` | 外部对接人 |
| `EXTERNAL_GROUPS` | 外部群 |
| `BANQUET_FLOW` | 晚宴节点 |
| `ARRIVAL_TIMES` | 分组到场 |
| `LIAISON_LIST` | 对内联络 |
| `PHASE_MATRIX` | 阶段 × 组职责 |

## 版本记录建议

在脚本顶部或 PPT 决议页记录：`v主.次` + 变更 bullet（调组、物资、新增页）。
