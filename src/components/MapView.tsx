import { useEffect, useRef, useState } from 'react'
import AMapLoader from '@amap/amap-jsapi-loader'
import type { Building } from '../types'
import { computeMetrics, typeColor } from '../utils/metrics'
import { AMAP_KEY } from '../config'

interface Props {
  data: Building[]
  onSelect: (b: Building) => void
}

const typeList: Building['propertyType'][] = ['写字楼', '产业园', '租赁住宅', '混合']

export default function MapView({ data, onSelect }: Props) {
  const containerRef = useRef<HTMLDivElement>(null)
  const mapRef = useRef<any>(null)
  const markersRef = useRef<Record<string, any>>({})
  const [error, setError] = useState<string>('')
  const [activeId, setActiveId] = useState<string>('')

  useEffect(() => {
    let destroyed = false
    // 高德 JS API 2.0 安全配置（如未配置 code，仅影响部分服务，不影响打点展示）
    ;(window as any)._AMapSecurityConfig = { securityJsCode: '' }

    AMapLoader.load({
      key: AMAP_KEY,
      version: '2.0',
      plugins: ['AMap.ToolBar', 'AMap.Scale'],
    })
      .then((AMap: any) => {
        if (destroyed || !containerRef.current) return
        const map = new AMap.Map(containerRef.current, {
          zoom: 13,
          center: [121.5185, 31.2846],
          mapStyle: 'amap://styles/dark',
          viewMode: '2D',
        })
        map.addControl(new AMap.ToolBar({ position: { top: '12px', right: '12px' } }))
        map.addControl(new AMap.Scale())
        mapRef.current = map

        data.forEach((b) => {
          const m = computeMetrics(b)
          const color = typeColor[b.propertyType]
          const content = `
            <div style="position:relative;transform:translate(-50%,-100%);">
              <div style="background:${color};color:#0b1220;font-weight:700;font-size:11px;
                padding:3px 8px;border-radius:8px;white-space:nowrap;
                box-shadow:0 6px 14px -4px rgba(0,0,0,.6);border:1px solid rgba(255,255,255,.5);">
                ${b.name}
              </div>
              <div style="width:0;height:0;border-left:5px solid transparent;border-right:5px solid transparent;
                border-top:7px solid ${color};margin:0 auto;"></div>
            </div>`
          const marker = new AMap.Marker({
            position: [b.lng, b.lat],
            content,
            anchor: 'bottom-center',
            offset: new AMap.Pixel(0, 0),
          })
          const info = new AMap.InfoWindow({
            isCustom: false,
            offset: new AMap.Pixel(0, -36),
            content: `<div class="ap-infowindow">
              <b>${b.name}</b><br/>
              ${b.plate} · ${b.propertyType} · ${b.grade}<br/>
              报价租金：${b.askingRent} 元/㎡/天<br/>
              出租率：${m.occupancyRate}% · 企业 ${m.tenantCount} 家
            </div>`,
          })
          marker.on('mouseover', () => info.open(map, [b.lng, b.lat]))
          marker.on('mouseout', () => info.close())
          marker.on('click', () => {
            setActiveId(b.id)
            onSelect(b)
          })
          marker.setMap(map)
          markersRef.current[b.id] = marker
        })
      })
      .catch((e: unknown) => {
        console.error(e)
        setError('地图加载失败，请检查网络或高德地图密钥配置。下方仍可浏览楼宇列表。')
      })

    return () => {
      destroyed = true
      mapRef.current?.destroy?.()
      mapRef.current = null
      markersRef.current = {}
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const focus = (b: Building) => {
    setActiveId(b.id)
    const map = mapRef.current
    if (map) {
      map.setZoomAndCenter(15, [b.lng, b.lat])
    }
  }

  return (
    <div className="map-layout">
      <div className="map-side">
        <div className="panel">
          <h3 className="panel-title"><span className="bar" />图例</h3>
          <div className="map-legend">
            {typeList.map((t) => (
              <div className="legend-item" key={t}>
                <span style={{ width: 12, height: 12, borderRadius: 3, background: typeColor[t] }} />
                {t}
              </div>
            ))}
          </div>
        </div>
        {data.map((b) => {
          const m = computeMetrics(b)
          return (
            <div
              key={b.id}
              className={`map-list-item ${activeId === b.id ? 'active' : ''}`}
              onClick={() => focus(b)}
            >
              <div className="ml-name">{b.name}</div>
              <div className="ml-sub">
                <span>{b.plate} · {b.propertyType}</span>
                <span>出租 {m.occupancyRate}%</span>
              </div>
            </div>
          )
        })}
      </div>
      {error ? (
        <div className="map-fallback">
          <div>
            <div style={{ fontSize: 15, marginBottom: 8, color: 'var(--bad)' }}>⚠ 地图未能加载</div>
            <div>{error}</div>
          </div>
        </div>
      ) : (
        <div id="amap-container" ref={containerRef} />
      )}
    </div>
  )
}
