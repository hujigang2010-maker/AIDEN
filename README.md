# AIDEN

AIDEN 是一个多任务工作区。`main` 分支存放共享配置；具体应用和文档交付物位于各自的 `cursor/*` 功能分支。

## 显示语言

本仓库默认使用**简体中文**。首次打开项目时，请安装推荐的中文语言包，并在命令面板中选择 **Configure Display Language → 中文(简体)**，然后重启 Cursor。

更多说明见 [AGENTS.md](./AGENTS.md)。

## 本分支：《以韩为鉴》书稿

以《以日为鉴：衰退时代生存指南》的五篇结构为参照，围绕就业、经济环境、周期反馈与康波周期，整理韩国的民生经验，供中国借鉴。

| 路径 | 内容 |
| --- | --- |
| `manuscript/00_全书提纲.md` | 全书定位、与《以日为鉴》的结构对照、七篇二十二章提纲、核心论点 |
| `manuscript/01_前言.md` — `09_附录.md` | 分篇书稿（每章八段式：民生画像、发生了什么、数据说话、周期定位、代价与反馈、中国对照、借鉴与反思、本章要点） |
| `manuscript/figures/` | 11 张数据配图 |
| `output/以韩为鉴_书稿.docx` | Word 版书稿（封面、目录、配图、页码） |
| `output/以韩为鉴_书稿全文.md` | 合并后的全文 Markdown |
| `output/以韩为鉴_数据底稿.xlsx` | 数据底稿：韩国核心数据（含来源链接）、中韩对照、康波周期坐标 |

重新生成全部交付物：

```bash
pip install python-docx openpyxl matplotlib
./scripts/build_all_korea.sh
```
