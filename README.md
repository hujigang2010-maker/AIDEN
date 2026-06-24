# 杨浦区楼宇经济数据平台

一个面向上海市杨浦区的楼宇经济数据可视化系统，从**楼宇维度**与**入驻企业维度**展示租赁、产业与空间数据，提供数据总览、明细表格与地图分布三大模块。

![tech](https://img.shields.io/badge/React-18-38bdf8) ![tech](https://img.shields.io/badge/Vite-5-6366f1) ![tech](https://img.shields.io/badge/ECharts-5-a855f7) ![map](https://img.shields.io/badge/高德地图-2.0-22c55e)

## 功能模块

### 1. 数据总览（Dashboard）
- 关键指标卡片：楼宇总数、总建筑面积、平均出租率、平均报价租金、欠租预警
- 各板块楼宇建筑面积分布
- 物业类型 / 物业等级构成（环形图）
- 入驻行业租赁面积 TOP8
- 报价租金 vs 成交租金散点分析

### 2. 楼宇数据（Table）
- 关键词搜索（楼宇 / 企业 / 行业）
- 按板块、物业类型、物业等级多维筛选
- 所有数值列支持升降序排序
- 行内出租率进度条
- 点击任意行查看楼宇详情抽屉

### 3. 地图分布（Map）
- 基于高德地图（Amap 2.0）暗色底图
- 楼宇按物业类型着色打点
- 悬浮查看摘要、点击联动详情
- 左侧楼宇列表点击定位

## 数据维度

**楼宇维度**：楼宇名称、所属板块、物业类型（写字楼 / 产业园 / 租赁住宅 / 混合）、物业等级（超甲级 / 甲级 / 乙级 / 丙级 / 其它）、报价租金、物业费、楼宇总建筑面积、竣工时间、标准层层高。

**入驻企业维度**：入驻企业名称、所属行业、入驻楼层、租赁面积、成交租金、免租期、企业经营状态、剩余租期。

> 数据为演示样例，可在 `src/data/buildings.ts` 中替换为真实业务数据。

## 本地运行

```bash
npm install
npm run dev      # 启动开发服务器 http://localhost:5173
npm run build    # 生产构建
npm run preview  # 预览构建产物
```

## 高德地图密钥

地图密钥配置位于 `src/config.ts`，默认使用内置密钥；也可通过环境变量覆盖：

```bash
# .env.local
VITE_AMAP_KEY=你的高德Web端Key
```

> 高德 JS API 2.0 若需使用搜索/路径等服务，需在 [高德控制台](https://lbs.amap.com) 申请并配置 `securityJsCode`；本系统仅用于底图与打点展示，普通 Key 即可正常使用。

## 技术栈

- React 18 + TypeScript + Vite
- ECharts（echarts-for-react）数据可视化
- 高德地图 @amap/amap-jsapi-loader
- 纯 CSS 自定义暗色主题，无额外 UI 框架依赖
