# 七款中文 AI PPT Skill 横评

一份 16:9 可编辑 PPT，把社区里常被点名的七款中文 PPT Skill 摊开：交什么文件、适合什么场合、什么时候不该用。

## 交付物

- [`output/七款中文AI-PPT-Skill横评.pptx`](output/七款中文AI-PPT-Skill横评.pptx) — 18 页，原生形状与文字，可在 PowerPoint / WPS 里继续改
- [`scripts/build_ppt_skill_overview.py`](scripts/build_ppt_skill_overview.py) — 生成脚本

## 怎么重新生成

```bash
pip install python-pptx
python3 scripts/build_ppt_skill_overview.py
```

可选预览（把每页栅格化成 PNG）：

```bash
python3 scripts/render_pptx_preview.py
```

预览图写到 `output/preview/`，不纳入版本库。

## 口径

| 项 | 说明 |
| --- | --- |
| 排名 | 沿用社区锐评（彼得潘AI / 即刻），看的是「能不能做出像样的 PPT」 |
| Star | 2026-08-23 从 GitHub 抓取的实时值，与社区原文中的约数不同 |
| 宝玉 Star | 整个 `JimLiu/baoyu-skills` 仓库，不是 `baoyu-slide-deck` 单技能 |
| 歸藏 Star | 已超过花叔，锐评名次未按 Star 重排 |

## 七款与三条路线

| 社区排名 | Skill | 作者 | 产物 |
| --- | --- | --- | --- |
| 01 | [ppt-master](https://github.com/hugohe3/ppt-master) | hugohe | 可编辑 PPTX |
| 02 | [frontend-slides](https://github.com/zarazhangrui/frontend-slides) | 张咋啦 | 单文件 HTML |
| 03 | [huashu-design](https://github.com/alchaincyf/huashu-design) | 花叔 | HTML + 可编辑 PPTX |
| 04 | [guizang-ppt-skill](https://github.com/op7418/guizang-ppt-skill) | 歸藏 | 单文件 HTML |
| 05 | [html-ppt-skill](https://github.com/lewislulu/html-ppt-skill) | Lewis | 单文件 HTML |
| 06 | [baoyu-slide-deck](https://github.com/JimLiu/baoyu-skills/tree/main/skills/baoyu-slide-deck) | 宝玉 | 图片卡片 |
| 07 | [qiaomu-anything-to-notebooklm](https://github.com/joeseesun/qiaomu-anything-to-notebooklm) | 乔木 | 图片 / 多格式 |

选型先问「这份文件第二天要交给谁」：领导改稿走 PPTX，自己上台走 HTML，发群发社媒走图片。
