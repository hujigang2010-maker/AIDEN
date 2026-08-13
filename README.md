# 复旦链接 · 开放协作平台（V3.0 集合版）

由**复旦大学住房政策研究中心**发起。视觉主色 **NYU Violet `#57068C`**。

**公式：** 复旦大学小程序（服务宫格）+ 管院校友中心栏目 + 教育中心公开课/定制 + 互动吧活动能力 → **复旦链接**。

## 请先打开集合版

| 交付 | 说明 |
|------|------|
| `demo/集合版.html` | **集合版网页**：需求清单 + 全部对照表 + 优质内容 |
| `output/复旦链接-集合版.xlsx` | **集合版表格**：一站式 Excel |
| `demo/index.html` | 小程序手机框演示 |
| `miniprogram/` | 可导入微信开发者工具 |

## 下载

- 集合版 Excel：https://github.com/hujigang2010-maker/AIDEN/raw/cursor/fudan-hprc-activity-platform-14f7/output/%E5%A4%8D%E6%97%A6%E9%93%BE%E6%8E%A5-%E9%9B%86%E5%90%88%E7%89%88.xlsx
- 完整蓝图：https://github.com/hujigang2010-maker/AIDEN/raw/cursor/fudan-hprc-activity-platform-14f7/output/%E5%A4%8D%E6%97%A6%E9%93%BE%E6%8E%A5%E5%BC%80%E6%94%BE%E5%8D%8F%E4%BD%9C%E5%B9%B3%E5%8F%B0-%E5%B0%8F%E7%A8%8B%E5%BA%8F%E5%8A%9F%E8%83%BD%E8%93%9D%E5%9B%BE.xlsx

## 重新生成

```bash
python3 scripts/build_collection.py
python3 scripts/build_feature_table.py
python3 scripts/build_hprc_platform_excel.py
```
