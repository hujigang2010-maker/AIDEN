import { useEffect } from 'react'
import type { Building } from '../types'
import { computeMetrics, formatArea, gradeColor, occupancyColor, statusColor, typeColor } from '../utils/metrics'
import { Tag } from './ui'

export default function DetailDrawer({
  building,
  onClose,
}: {
  building: Building | null
  onClose: () => void
}) {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && onClose()
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  if (!building) return null
  const m = computeMetrics(building)

  const info: { k: string; v: React.ReactNode }[] = [
    { k: '楼宇编号', v: building.id },
    { k: '所属板块', v: building.plate },
    { k: '报价租金', v: <>{building.askingRent}<small> 元/㎡/天</small></> },
    { k: '平均成交租金', v: m.avgDealRent ? <>{m.avgDealRent}<small> 元/㎡/天</small></> : '—' },
    { k: '物业费', v: building.propertyFee ? <>{building.propertyFee}<small> 元/㎡/月</small></> : '—' },
    { k: '总建筑面积', v: <>{building.totalArea.toLocaleString('zh-CN')}<small> ㎡</small></> },
    { k: '竣工时间', v: <>{building.completionYear}<small> 年</small></> },
    { k: '标准层层高', v: <>{building.floorHeight}<small> 米</small></> },
    { k: '出租率', v: <span style={{ color: occupancyColor(m.occupancyRate) }}>{m.occupancyRate}%</span> },
    { k: '已租面积', v: <>{m.leasedArea.toLocaleString('zh-CN')}<small> ㎡</small></> },
  ]

  return (
    <>
      <div className="drawer-mask" onClick={onClose} />
      <aside className="drawer" role="dialog" aria-label={building.name}>
        <div className="drawer-head">
          <button className="close" onClick={onClose} aria-label="关闭">×</button>
          <h2>{building.name}</h2>
          <div className="drawer-tags">
            <Tag color={typeColor[building.propertyType]}>{building.propertyType}</Tag>
            <Tag color={gradeColor[building.grade]}>{building.grade}</Tag>
            <Tag color="#38bdf8" dot={false}>{building.plate}</Tag>
          </div>
        </div>

        <div className="drawer-body">
          <div className="info-grid">
            {info.map((i) => (
              <div className="info-cell" key={i.k}>
                <div className="k">{i.k}</div>
                <div className="v">{i.v}</div>
              </div>
            ))}
          </div>

          <h3 className="sub-title">
            入驻企业明细 <span className="count">{building.tenants.length}</span>
          </h3>
          {building.tenants.map((t, idx) => (
            <div className="tenant-card" key={idx}>
              <div className="t-head">
                <span className="t-name">{t.name}</span>
                <Tag color={statusColor[t.status]}>{t.status}</Tag>
              </div>
              <div className="t-meta">
                <div className="row"><span>所属行业</span><span>{t.industry}</span></div>
                <div className="row"><span>入驻楼层</span><span>{t.floor}</span></div>
                <div className="row"><span>租赁面积</span><span>{formatArea(t.leaseArea)}</span></div>
                <div className="row"><span>成交租金</span><span>{t.dealRent} 元/㎡/天</span></div>
                <div className="row"><span>免租期</span><span>{t.rentFreeMonths} 个月</span></div>
                <div className="row">
                  <span>剩余租期</span>
                  <span style={{ color: t.remainingMonths > 0 && t.remainingMonths <= 6 ? '#f59e0b' : undefined }}>
                    {t.remainingMonths > 0 ? `${t.remainingMonths} 个月` : '已到期'}
                  </span>
                </div>
              </div>
            </div>
          ))}
        </div>
      </aside>
    </>
  )
}
