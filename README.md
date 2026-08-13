# 复旦链接 · 开放协作平台

由**复旦大学住房政策研究中心**发起。视觉主色 **NYU Violet `#57068C`**。

**主骨架**对标复旦大学小程序（服务宫格）+ 管院校友中心（组织/活动/终身学习/风采/支持）；**活动能力**叠加互动吧（多票种、海报、签到、裂变等）。对外品牌：复旦链接。

## Tab

首页｜服务｜活动｜我的

## 交付物

| 文件 | 说明 |
|------|------|
| `demo/展示页.html` | **功能对照网页展示页**（含表格，可浏览器直接打开） |
| `output/复旦链接-功能对照表.xlsx` | **功能对照简表**（三源对标 / IA / 互动吧 / 收费） |
| `output/复旦链接开放协作平台-小程序功能蓝图.xlsx` | 完整功能蓝图 |
| `miniprogram/` | 微信小程序演示 |
| `demo/index.html` | 小程序手机框网页演示 |

## 网页展示页

本地打开：`demo/展示页.html`

## 表格下载

- 功能对照表：https://github.com/hujigang2010-maker/AIDEN/raw/cursor/fudan-hprc-activity-platform-14f7/output/%E5%A4%8D%E6%97%A6%E9%93%BE%E6%8E%A5-%E5%8A%9F%E8%83%BD%E5%AF%B9%E7%85%A7%E8%A1%A8.xlsx
- 完整蓝图：https://github.com/hujigang2010-maker/AIDEN/raw/cursor/fudan-hprc-activity-platform-14f7/output/%E5%A4%8D%E6%97%A6%E9%93%BE%E6%8E%A5%E5%BC%80%E6%94%BE%E5%8D%8F%E4%BD%9C%E5%B9%B3%E5%8F%B0-%E5%B0%8F%E7%A8%8B%E5%BA%8F%E5%8A%9F%E8%83%BD%E8%93%9D%E5%9B%BE.xlsx

## 重新生成

```bash
python3 scripts/build_feature_table.py
python3 scripts/build_hprc_platform_excel.py
```
