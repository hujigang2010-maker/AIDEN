# 杨浦五角场 5km 商办竞品地图

基于**高德地图 JS API** 的商办竞品调研可视化网页 — 顶部总览看板、左侧交互地图、右侧【项目概览】/【企业明细】双 Tab，地图与表格全双向联动。

> **基准点：** 上海市五角场创新创业学院（杨浦区国权路18号）
> **半径：** 5 km
> **分层：** 1 km / 1–2 km / 2–3 km / 3–5 km
> **数据：** 39 个项目 · 136 家代表企业

---

## 功能一览

| 模块 | 说明 |
| --- | --- |
| 总览看板 | 项目总数 + 4 圈层项目数,可点击直接做距离筛选 |
| 交互地图 | 基准点呼吸光晕 + 4 圈层 + 业态色彩标记 + 编号 + 企业数 badge |
| Tab：项目概览 | 序号 / 项目 / 业态 / 距离 / 🚗 驾车 / 🚇 公交,与地图双向联动 |
| Tab：企业明细 | 项目+企业 / 地址 / 距离 / 办公体量 / 成交租金 / 免租期 / 主力客户行业 / 物业费 / 性质说明 |
| InfoWindow | 弹窗中显示项目信息 + 代表企业 mini 表格（最多 8 家） |
| 筛选 | 业态（商办写字楼 / 产业园区 / 众创空间） × 距离 |
| 搜索 | 项目名 / 企业名 / 地址 / 标签 / 行业 / 性质 全部命中 |
| 详情面板 | 含 "在高德打开 / 路线规划" 一键直达 |
| CSV 导出 | 根据当前 Tab 输出对应字段,UTF-8 BOM,Excel 中文不乱码 |
| 响应式 | 1180px 以下自动切换上下布局 |

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
├── index.html                  # 入口页（dashboard + map + Tab 表格 三段式）
├── css/
│   └── style.css               # 全部样式
├── js/
│   └── app.js                  # 地图初始化、双 Tab 渲染、筛选、联动、CSV
└── data/
    └── competitors.js          # 竞品 + 企业数据集
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
    category: '商办写字楼',          // 商办写字楼 / 产业园区 / 众创空间
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

    // —— 新增项目级字段 ——
    propertyFee: 28,                  // 物业费（元/㎡/月）
    nature: '民营 · 甲级写字楼',       // 性质说明（产权 + 物业等级 / 园区性质）

    // —— 下属企业 ——
    tenants: [
      {
        name: '招商银行 杨浦支行',     // 企业名称
        industry: '金融 / 银行',       // 主力客户行业
        area: 3200,                    // 办公体量（㎡）
        rent: 5.8,                     // 成交租金（元/㎡/天）
        rentFreeMonths: 3,             // 免租期（月）
        dealYear: 2023                 // 成交/在租年份
      },
      // ...
    ],

    excluded: false                    // 标 true 则在统计与展示中剔除（如邻里参照项目）
  },
  // ...
];
```

> 距离与圈层在脚本中自动计算（haversine 球面距离）—— 你只需要维护坐标即可。

### 企业明细表的列对应关系

| 表格列 | 来源字段 |
| --- | --- |
| # | 自动序号 |
| 项目名称 / 企业 | `park.name` + `tenant.name` |
| 地址 | `park.address` |
| 距离 / 距离段 | 由 `lng/lat` 与基准点计算 |
| 办公体量 (㎡) | `tenant.area` |
| 成交租金 (元/㎡/天) | `tenant.rent` |
| 免租期 | `tenant.rentFreeMonths` |
| 主力客户行业 | `tenant.industry` |
| 物业费 (元/㎡/月) | `park.propertyFee` |
| 性质说明 | `park.nature` |

### 如何从 Excel 导入数据

1. Excel 至少维护这些列：
   `项目名称 / 业态 / 地址 / 经度 / 纬度 / 物业费 / 性质 / 企业名称 / 行业 / 体量 / 租金 / 免租期 / 成交年份`
2. Python 把多行（每行一个企业）按项目分组后转 JS：

   ```python
   import pandas as pd, json, collections
   df = pd.read_excel('杨浦5km竞品调研数据汇总 3.xlsx', sheet_name='竞品调研数据（全量）')
   # rename + clean ...
   parks = collections.OrderedDict()
   for _, r in df.iterrows():
       pid = r['项目名称']
       if pid not in parks:
           parks[pid] = {
               'id': f'X{len(parks)+1:03d}',
               'name': r['项目名称'], 'category': r['业态'],
               'address': r['地址'], 'lng': r['经度'], 'lat': r['纬度'],
               'propertyFee': r['物业费'], 'nature': r['性质'],
               'tenants': []
           }
       if pd.notna(r.get('企业名称', None)):
           parks[pid]['tenants'].append({
               'name': r['企业名称'], 'industry': r['行业'],
               'area': r['体量'], 'rent': r['租金'],
               'rentFreeMonths': r['免租期'], 'dealYear': r.get('成交年份')
           })
   print('window.COMPETITORS =', json.dumps(list(parks.values()), ensure_ascii=False, indent=2))
   ```

   贴回 `data/competitors.js` 末尾即可。

---

## 当前数据规模

- **39** 个园区/写字楼项目（1km 内 7、1–2km 10、2–3km 10、3–5km 12）
- **136** 家代表性企业 / 锚租户
- 涵盖业态：商办写字楼 16 · 产业园区 16 · 众创空间 7

> 数据为基于公开资料整理的代表性样本,如需以最新 Excel 调研数据为准,请按上节方式从 Excel 重新生成。

---

## 已知限制

- 公交/驾车时长字段为估算,正式版可调用 `AMap.Driving` / `AMap.Transfer` 接口实时计算
- 当前为纯静态页,无后端;若需多人协作维护,推荐配合 Google Sheet / 飞书多维表 → JSON 工作流
- 部分企业（如 `初创团队（孵化期免租）`、`联合办公租户合计` 等）是合并条目,代表园区内同类租户的聚合
