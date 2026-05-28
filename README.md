# 杨浦五角场 5km 商办竞品地图

一个基于**高德地图 JS API** 的商办竞品调研可视化网页 — 顶部总览看板、左侧交互地图、右侧数据清单三联动,精简而专业。

> **基准点：** 上海市五角场创新创业学院（杨浦区国权路18号）
> **半径：** 5 km
> **分层：** 1 km / 1–2 km / 2–3 km / 3–5 km

---

## 功能一览

| 模块 | 说明 |
| --- | --- |
| 总览看板 | 项目总数 + 4 个距离圈层项目数,可点击直接做距离筛选 |
| 交互地图 | 基准点呼吸式光晕 + 4 圈层 + 业态色彩标记 + 编号 + 弹窗 |
| 数据清单 | 序号、项目、业态、距离、🚗 驾车、🚇 公交/地铁,与地图双向联动 |
| 业态筛选 | 全部 / 商办写字楼 / 产业园区 / 众创空间 |
| 距离筛选 | 全部 / 1km / 1–2km / 2–3km / 3–5km |
| 搜索 | 名称 / 地址 / 标签 / 合并楼宇 / 开发商 模糊匹配 |
| 详情弹层 | 含 “在高德打开 / 路线规划” 一键直达 |
| CSV 导出 | 导出当前筛选结果（含 BOM,Excel 中文不乱码） |
| 响应式 | 1100px 以下自动切换上下布局 |

---

## 使用方式

直接用浏览器打开 `index.html` 即可（需联网加载高德 JS API）。
也可以用任意静态服务器运行：

```bash
python -m http.server 8000
# 然后访问 http://localhost:8000
```

> 高德 Key 已写入 `index.html`（`b96fe5c5cdb328eceb9a6bf1a6d1af7f`,
> 安全密钥 `84da8ce45182a0281c095a358456ab4e`,
> 通过 `window._AMapSecurityConfig` 注入,符合 2021 后的安全要求）。

---

## 目录结构

```
.
├── index.html                  # 入口页（含 dashboard + map + table 三段式布局）
├── css/
│   └── style.css               # 全部样式
├── js/
│   └── app.js                  # 地图初始化、渲染、筛选、联动、CSV 导出
└── data/
    └── competitors.js          # 竞品数据集
```

---

## 数据结构

`data/competitors.js` 中：

```js
window.BASE_POINT = {
  name: '上海市五角场创新创业学院',
  address: '上海市杨浦区国权路18号',
  lng: 121.5103,
  lat: 31.3008
};

window.COMPETITORS = [
  {
    id: 'A01',
    name: '合生国际广场',
    category: '商办写字楼',     // 商办写字楼 / 产业园区 / 众创空间
    address: '上海市杨浦区淞沪路18号',
    lng: 121.5076, lat: 31.2984,
    drive:   '4 分钟 / 1.1 km',
    transit: '步行 11 分钟 · 10号线五角场站',
    developer: '合生创展',
    area: '约 18 万㎡',
    year: 2008,
    subs:  ['T1 商务塔', 'T2 商务塔', 'L1-L4 商业裙房'],
    tags:  ['核心商圈', '甲级写字楼', '地铁口'],
    note:  '五角场核心商圈代表写字楼,甲级标准,层高 3.8m。',
    excluded: false                  // 标 true 则在统计与展示中剔除（如邻里参照项目）
  },
  // ...
];
```

> 距离与圈层在脚本中自动计算（haversine 球面距离）—— 你只需要维护坐标。

### 如何从 Excel 导入数据

1. 在 Excel 中维护好以下列：
   `项目名称 / 业态 / 地址 / 经度 / 纬度 / 驾车 / 公交 / 开发商 / 体量 / 投用 / 含楼宇 / 标签 / 备注`
2. 用 Python 一行转 JS：

   ```python
   import pandas as pd, json
   df = pd.read_excel('杨浦5km竞品调研数据汇总 3.xlsx', sheet_name='竞品调研数据（全量）')
   df = df.rename(columns={
       '项目名称':'name','业态':'category','地址':'address',
       '经度':'lng','纬度':'lat','驾车':'drive','公交':'transit',
       '开发商':'developer','体量':'area','投用':'year',
       '含楼宇':'subs','标签':'tags','备注':'note'
   })
   df['subs'] = df['subs'].fillna('').apply(lambda x: [s.strip() for s in str(x).replace('、',',').split(',') if s.strip()])
   df['tags'] = df['tags'].fillna('').apply(lambda x: [s.strip() for s in str(x).replace('、',',').split(',') if s.strip()])
   records = df.to_dict(orient='records')
   for i, r in enumerate(records, 1):
       r['id'] = f'X{i:03d}'
   print('window.COMPETITORS =', json.dumps(records, ensure_ascii=False, indent=2))
   ```

   将输出贴回 `data/competitors.js` 末尾即可（注意保留 `window.BASE_POINT`）。

---

## 截图描述

- 顶部：渐变蓝色看板,5 张统计卡（总数 + 4 个圈层）
- 左下：基准点橘色呼吸光晕 + 4 圈层（1/2/3/5km）虚线
- 标记：圆角水滴,颜色按圈层（蓝绿橙红）,内含业态 emoji + 编号
- 右侧：表格列出全部项目,行 hover/点击高亮,左侧蓝色指示条
- 点击标记/行 → 弹窗 + 右下角详情面板(可在高德打开 / 路线规划)

---

## 已知限制

- 数据为基于公开资料的代表性样本,如需以最新调研数据为准,请按上节方式从 Excel 重新生成。
- 公交时长字段为估算,正式版本可调用 `AMap.Transfer` 接口动态计算（脚本已加载该插件,可直接扩展）。
- 当前为纯静态页,无后端;若需多人协作维护数据,推荐配合 Google Sheet → JSON 工作流。
