import ReactECharts from 'echarts-for-react'
import type { Building } from '../types'
import { computeMetrics, formatWan, round } from '../utils/metrics'
import { gradeColor, typeColor } from '../utils/metrics'

const palette = ['#38bdf8', '#6366f1', '#a855f7', '#22c55e', '#f59e0b', '#f472b6', '#ef4444', '#14b8a6']

const baseAxis = {
  axisLine: { lineStyle: { color: 'rgba(120,150,200,0.3)' } },
  axisLabel: { color: '#93a2bd' },
  splitLine: { lineStyle: { color: 'rgba(120,150,200,0.1)' } },
}

const tooltipStyle = {
  backgroundColor: 'rgba(13,20,36,0.95)',
  borderColor: 'rgba(120,150,200,0.3)',
  textStyle: { color: '#e6ecf7' },
}

export default function Dashboard({ data }: { data: Building[] }) {
  const metrics = data.map((b) => ({ b, m: computeMetrics(b) }))

  const totalArea = data.reduce((s, b) => s + b.totalArea, 0)
  const leasedArea = metrics.reduce((s, x) => s + x.m.leasedArea, 0)
  const tenantCount = metrics.reduce((s, x) => s + x.m.tenantCount, 0)
  const avgOcc = round((leasedArea / totalArea) * 100, 1)
  const avgAsking = round(data.reduce((s, b) => s + b.askingRent, 0) / data.length, 2)
  const warnCount = data.reduce(
    (s, b) => s + b.tenants.filter((t) => t.status === '欠租预警').length,
    0,
  )

  // 板块分布
  const plateMap = new Map<string, { area: number; count: number; rentSum: number }>()
  data.forEach((b) => {
    const cur = plateMap.get(b.plate) ?? { area: 0, count: 0, rentSum: 0 }
    cur.area += b.totalArea
    cur.count += 1
    cur.rentSum += b.askingRent
    plateMap.set(b.plate, cur)
  })
  const plates = [...plateMap.entries()].sort((a, b) => b[1].area - a[1].area)

  // 物业类型分布
  const typeMap = new Map<string, number>()
  data.forEach((b) => typeMap.set(b.propertyType, (typeMap.get(b.propertyType) ?? 0) + 1))

  // 物业等级分布
  const gradeOrder = ['超甲级', '甲级', '乙级', '丙级', '其它']
  const gradeMap = new Map<string, number>()
  data.forEach((b) => gradeMap.set(b.grade, (gradeMap.get(b.grade) ?? 0) + 1))

  // 行业分布（按租赁面积）
  const indMap = new Map<string, number>()
  data.forEach((b) =>
    b.tenants
      .filter((t) => t.status !== '已退租')
      .forEach((t) => indMap.set(t.industry, (indMap.get(t.industry) ?? 0) + t.leaseArea)),
  )
  const industries = [...indMap.entries()].sort((a, b) => b[1] - a[1]).slice(0, 8)

  const plateBarOption = {
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, ...tooltipStyle },
    grid: { left: 0, right: 24, top: 18, bottom: 0, containLabel: true },
    xAxis: { type: 'value', name: '万㎡', nameTextStyle: { color: '#64748b' }, ...baseAxis },
    yAxis: {
      type: 'category',
      data: plates.map((p) => p[0]).reverse(),
      ...baseAxis,
      splitLine: { show: false },
    },
    series: [
      {
        type: 'bar',
        data: plates.map((p) => round(p[1].area / 10000, 1)).reverse(),
        barWidth: 16,
        itemStyle: {
          borderRadius: [0, 6, 6, 0],
          color: {
            type: 'linear',
            x: 0,
            y: 0,
            x2: 1,
            y2: 0,
            colorStops: [
              { offset: 0, color: '#6366f1' },
              { offset: 1, color: '#38bdf8' },
            ],
          },
        },
        label: { show: true, position: 'right', color: '#93a2bd', fontSize: 11 },
      },
    ],
  }

  const typePieOption = {
    tooltip: { trigger: 'item', ...tooltipStyle },
    legend: { bottom: 0, textStyle: { color: '#93a2bd' }, itemWidth: 10, itemHeight: 10 },
    series: [
      {
        type: 'pie',
        radius: ['45%', '70%'],
        center: ['50%', '44%'],
        avoidLabelOverlap: true,
        itemStyle: { borderColor: '#0d1424', borderWidth: 3 },
        label: { color: '#e6ecf7', fontSize: 12, formatter: '{b}\n{d}%' },
        data: [...typeMap.entries()].map(([k, v]) => ({
          name: k,
          value: v,
          itemStyle: { color: typeColor[k] },
        })),
      },
    ],
  }

  const gradePieOption = {
    tooltip: { trigger: 'item', ...tooltipStyle },
    legend: { bottom: 0, textStyle: { color: '#93a2bd' }, itemWidth: 10, itemHeight: 10 },
    series: [
      {
        type: 'pie',
        radius: ['45%', '70%'],
        center: ['50%', '44%'],
        itemStyle: { borderColor: '#0d1424', borderWidth: 3 },
        label: { color: '#e6ecf7', fontSize: 12, formatter: '{b}\n{c}栋' },
        data: gradeOrder
          .filter((g) => gradeMap.has(g))
          .map((g) => ({ name: g, value: gradeMap.get(g)!, itemStyle: { color: gradeColor[g] } })),
      },
    ],
  }

  const industryBarOption = {
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, ...tooltipStyle, valueFormatter: (v: number) => `${(v / 10000).toFixed(1)} 万㎡` },
    grid: { left: 0, right: 30, top: 18, bottom: 0, containLabel: true },
    xAxis: { type: 'value', name: '万㎡', nameTextStyle: { color: '#64748b' }, ...baseAxis },
    yAxis: {
      type: 'category',
      data: industries.map((i) => i[0]).reverse(),
      ...baseAxis,
      splitLine: { show: false },
    },
    series: [
      {
        type: 'bar',
        data: industries.map((i) => round(i[1] / 10000, 2)).reverse(),
        barWidth: 14,
        itemStyle: {
          borderRadius: [0, 6, 6, 0],
          color: {
            type: 'linear',
            x: 0,
            y: 0,
            x2: 1,
            y2: 0,
            colorStops: [
              { offset: 0, color: '#a855f7' },
              { offset: 1, color: '#f472b6' },
            ],
          },
        },
      },
    ],
  }

  // 报价 vs 成交租金散点
  const scatterOption = {
    tooltip: {
      trigger: 'item',
      ...tooltipStyle,
      formatter: (p: { data: [number, number, string] }) =>
        `<b>${p.data[2]}</b><br/>报价 ${p.data[0]} 元 · 成交 ${p.data[1]} 元/㎡/天`,
    },
    grid: { left: 6, right: 18, top: 24, bottom: 6, containLabel: true },
    xAxis: { type: 'value', name: '报价租金', nameTextStyle: { color: '#64748b' }, scale: true, ...baseAxis },
    yAxis: { type: 'value', name: '成交租金', nameTextStyle: { color: '#64748b' }, scale: true, ...baseAxis },
    series: [
      {
        type: 'scatter',
        symbolSize: (d: number[]) => Math.max(8, Math.sqrt(d[2] ?? 100) / 5),
        data: metrics
          .filter((x) => x.m.avgDealRent > 0)
          .map((x) => ({
            value: [x.b.askingRent, x.m.avgDealRent, x.b.name],
            itemStyle: { color: typeColor[x.b.propertyType], opacity: 0.85 },
          })),
      },
    ],
  }

  const officeCount = data.filter((b) => b.propertyType === '写字楼').length
  const parkCount = data.filter((b) => b.propertyType === '产业园').length

  const kpis = [
    { label: '楼宇总数', value: data.length, unit: '栋', sub: `写字楼 ${officeCount} · 产业园 ${parkCount} · 覆盖 ${plates.length} 板块`, c: '#38bdf8' },
    { label: '总建筑面积', value: formatWan(totalArea), sub: `已租 ${formatWan(leasedArea)}`, c: '#6366f1' },
    { label: '平均出租率', value: avgOcc, unit: '%', sub: `入驻企业 ${tenantCount} 家`, c: avgOcc >= 70 ? '#22c55e' : '#f59e0b' },
    { label: '平均报价租金', value: avgAsking, unit: '元/㎡/天', sub: '全区写字楼及园区', c: '#a855f7' },
    { label: '欠租预警', value: warnCount, unit: '家', sub: warnCount ? '需重点跟进' : '运行平稳', c: warnCount ? '#ef4444' : '#22c55e' },
  ]

  return (
    <div>
      <div className="kpi-grid">
        {kpis.map((k) => (
          <div className="kpi-card" key={k.label} style={{ ['--c' as string]: k.c }}>
            <div className="kpi-label">
              <span className="kpi-dot" style={{ ['--c' as string]: k.c }} />
              {k.label}
            </div>
            <div className="kpi-value">
              {typeof k.value === 'number' ? k.value.toLocaleString('zh-CN') : k.value}
              {k.unit && <small>{k.unit}</small>}
            </div>
            <div className="kpi-sub">{k.sub}</div>
          </div>
        ))}
      </div>

      <div className="chart-grid">
        <div className="panel span-2">
          <h3 className="panel-title">
            <span className="bar" />各板块楼宇建筑面积分布<span className="hint">单位：万㎡</span>
          </h3>
          <ReactECharts option={plateBarOption} className="echart" style={{ height: 280 }} />
        </div>

        <div className="panel">
          <h3 className="panel-title"><span className="bar" />物业类型构成</h3>
          <ReactECharts option={typePieOption} className="echart" />
        </div>

        <div className="panel">
          <h3 className="panel-title"><span className="bar" />物业等级构成</h3>
          <ReactECharts option={gradePieOption} className="echart" />
        </div>

        <div className="panel">
          <h3 className="panel-title">
            <span className="bar" />入驻行业租赁面积 TOP8<span className="hint">单位：万㎡</span>
          </h3>
          <ReactECharts option={industryBarOption} className="echart" />
        </div>

        <div className="panel">
          <h3 className="panel-title">
            <span className="bar" />报价 vs 成交租金<span className="hint">颜色代表物业类型</span>
          </h3>
          <ReactECharts option={scatterOption} className="echart" />
        </div>
      </div>
    </div>
  )
}

export { palette }
