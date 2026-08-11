# 复旦链接 · 开放协作平台

由**复旦大学住房政策研究中心**发起运营的开放协作小程序：**先搭平台**，再承载活动、产业机会与国际对接。  
住房与城市议题是平台垂直之一，但不束缚平台边界；所有专项可追溯至本平台。

> 文案中已去除管理学院 / 校友中心等外部机构表述。

## 交付物

| 文件 | 说明 |
|------|------|
| `output/复旦链接开放协作平台-小程序功能蓝图.xlsx` | 平台能力、产业机会、功能规划、**驻沪领事馆名录（姓名·无联系方式）**、命名规范 |
| `miniprogram/` | 微信小程序演示（首页｜活动｜机会｜我的） |
| `demo/index.html` | 浏览器交互演示 |
| `docs/产品说明.md` | 定位与表述规范 |

## 重新生成 Excel

```bash
python3 scripts/build_hprc_platform_excel.py
```

## Excel 下载

https://github.com/hujigang2010-maker/AIDEN/raw/cursor/fudan-hprc-activity-platform-14f7/output/%E5%A4%8D%E6%97%A6%E9%93%BE%E6%8E%A5%E5%BC%80%E6%94%BE%E5%8D%8F%E4%BD%9C%E5%B9%B3%E5%8F%B0-%E5%B0%8F%E7%A8%8B%E5%BA%8F%E5%8A%9F%E8%83%BD%E8%93%9D%E5%9B%BE.xlsx

## 微信开发者工具

导入 `miniprogram/`；支付为演示弹窗。领事馆页点击姓名可复制公开检索关键词。
