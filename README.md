# 许俊贤技能清单

把聊天记录里甩过来的 13 条链接，整理成一份 16:9 可编辑 PPT：十二件视觉生产工具，按产物分成四条赛道。

这不是排行榜。上一份材料《七款中文 AI PPT Skill 横评》解决「哪款能做出像样的 PPT」；这份解决「地图有多宽」——生成之后怎么改、怎么上飞书白板、怎么交 XRD 图。

## 交付物

- [`output/许俊贤技能清单.pptx`](output/许俊贤技能清单.pptx) — 19 页，原生形状与文字，可在 PowerPoint / WPS 里继续改
- [`scripts/build_skill_catalog.py`](scripts/build_skill_catalog.py) — 生成脚本

## 怎么重新生成

```bash
pip install python-pptx
python3 scripts/build_skill_catalog.py
```

可选预览（把每页栅格化成 PNG）：

```bash
pip install pillow
python3 scripts/render_pptx_preview.py
```

预览图写到 `output/preview/`，不纳入版本库。

## 四条赛道

| 赛道 | 产物 | 清单里的工具 |
| --- | --- | --- |
| 演示稿 | PPTX / 单文件 HTML / 图片卡片 | ppt-master、frontend-slides、归藏、宝玉、花叔、Lewis、乔木 |
| HTML 模板 | 34 套可克隆视觉系统 | beautiful-html-templates |
| 改稿器 | 浏览器里点选改 HTML | htmledit / [htmledit.ai](https://htmledit.ai/) |
| 飞书白板 | 可编辑原生对象 | 张咋啦白板、larkboard-graphy、XRD-SKILL |

## 口径

| 项 | 说明 |
| --- | --- |
| 来源 | 许俊贤聊天记录中的五段链接，按原话分组还原 |
| 排名 | 不做社区锐评排名；演示七款细读仍以横评材料为准 |
| Star | 2026-08-23 从 GitHub 抓取 |
| 宝玉 Star | 整个 `JimLiu/baoyu-skills` 仓库 |
| 财猫 | 仓库与 `htmledit.ai` 算同一产品的两个入口；协议 CC-BY-NC-4.0 |
| XRD 链接 | 聊天原文 `XRD- SKILL` 有空格，仓库名为 `LucianaiB2004/XRD-SKILL` |

选型先问「这份文件第二天要交给谁、在哪个软件里打开」。
