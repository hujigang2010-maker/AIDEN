import { useMemo, useState } from 'react'
import type { Building } from '../types'
import {
  computeMetrics,
  gradeColor,
  occupancyColor,
  typeColor,
} from '../utils/metrics'
import { IconSearch, Tag } from './ui'
import { buildDetailCsv, buildSummaryCsv, downloadCsv } from '../utils/csv'

type SortKey =
  | 'name'
  | 'plate'
  | 'askingRent'
  | 'propertyFee'
  | 'totalArea'
  | 'completionYear'
  | 'occupancy'
  | 'tenants'

interface Props {
  data: Building[]
  filters: ReturnType<typeof useTableFilters>
  onSelect: (b: Building) => void
}

export function useTableFilters() {
  const [search, setSearch] = useState('')
  const [plate, setPlate] = useState('全部')
  const [type, setType] = useState('全部')
  const [grade, setGrade] = useState('全部')
  const reset = () => {
    setSearch('')
    setPlate('全部')
    setType('全部')
    setGrade('全部')
  }
  return { search, setSearch, plate, setPlate, type, setType, grade, setGrade, reset }
}

const columns: { key: SortKey; label: string; num?: boolean }[] = [
  { key: 'name', label: '楼宇名称' },
  { key: 'plate', label: '所属板块' },
  { key: 'askingRent', label: '报价租金', num: true },
  { key: 'propertyFee', label: '物业费', num: true },
  { key: 'totalArea', label: '建筑面积(㎡)', num: true },
  { key: 'completionYear', label: '竣工', num: true },
  { key: 'occupancy', label: '出租率' },
  { key: 'tenants', label: '入驻企业', num: true },
]

export default function BuildingTable({ data, filters, onSelect }: Props) {
  const [sortKey, setSortKey] = useState<SortKey>('totalArea')
  const [asc, setAsc] = useState(false)

  const plateOptions = useMemo(
    () => ['全部', ...[...new Set(data.map((b) => b.plate))]],
    [data],
  )

  const filtered = useMemo(() => {
    const kw = filters.search.trim().toLowerCase()
    return data.filter((b) => {
      if (filters.plate !== '全部' && b.plate !== filters.plate) return false
      if (filters.type !== '全部' && b.propertyType !== filters.type) return false
      if (filters.grade !== '全部' && b.grade !== filters.grade) return false
      if (kw) {
        const hay = (
          b.name +
          b.plate +
          b.id +
          b.tenants.map((t) => t.name + t.industry).join('')
        ).toLowerCase()
        if (!hay.includes(kw)) return false
      }
      return true
    })
  }, [data, filters.search, filters.plate, filters.type, filters.grade])

  const rows = useMemo(() => {
    const withM = filtered.map((b) => ({ b, m: computeMetrics(b) }))
    withM.sort((x, y) => {
      let cmp = 0
      switch (sortKey) {
        case 'name':
          cmp = x.b.name.localeCompare(y.b.name, 'zh')
          break
        case 'plate':
          cmp = x.b.plate.localeCompare(y.b.plate, 'zh')
          break
        case 'occupancy':
          cmp = x.m.occupancyRate - y.m.occupancyRate
          break
        case 'tenants':
          cmp = x.m.tenantCount - y.m.tenantCount
          break
        default:
          cmp = (x.b[sortKey] as number) - (y.b[sortKey] as number)
      }
      return asc ? cmp : -cmp
    })
    return withM
  }, [filtered, sortKey, asc])

  const onSort = (k: SortKey) => {
    if (k === sortKey) setAsc(!asc)
    else {
      setSortKey(k)
      setAsc(false)
    }
  }

  return (
    <div>
      <div className="toolbar">
        <div className="search-box">
          <IconSearch />
          <input
            value={filters.search}
            onChange={(e) => filters.setSearch(e.target.value)}
            placeholder="搜索楼宇 / 企业 / 行业…"
          />
        </div>
        <Select value={filters.plate} onChange={filters.setPlate} options={plateOptions} />
        <Select
          value={filters.type}
          onChange={filters.setType}
          options={['全部', '写字楼', '产业园', '租赁住宅', '混合']}
        />
        <Select
          value={filters.grade}
          onChange={filters.setGrade}
          options={['全部', '超甲级', '甲级', '乙级', '丙级', '其它']}
        />
        <button className="btn-reset" onClick={filters.reset}>
          重置
        </button>
        <button
          className="btn-export"
          onClick={() => downloadCsv('杨浦区楼宇数据_简版.csv', buildSummaryCsv(filtered))}
          title="简版：每栋楼宇一行（楼宇维度核心字段）"
        >
          ⬇ 导出简版
        </button>
        <button
          className="btn-export"
          onClick={() => downloadCsv('杨浦区楼宇数据_详细版.csv', buildDetailCsv(filtered))}
          title="详细版：每个入驻企业一行（楼宇 + 企业全字段）"
        >
          ⬇ 导出详细版
        </button>
        <div className="result-count">
          共 <b>{rows.length}</b> 栋楼宇
        </div>
      </div>

      <div className="table-wrap">
        <div className="table-scroll">
          <table className="data">
            <thead>
              <tr>
                {columns.map((c) => (
                  <th
                    key={c.key}
                    className={`sortable ${c.num ? 'num' : ''} ${sortKey === c.key ? 'sorted' : ''}`}
                    onClick={() => onSort(c.key)}
                  >
                    {c.label}
                    <span className="arrow">
                      {sortKey === c.key ? (asc ? '▲' : '▼') : '↕'}
                    </span>
                  </th>
                ))}
                <th>物业 / 等级</th>
              </tr>
            </thead>
            <tbody>
              {rows.map(({ b, m }) => (
                <tr key={b.id} onClick={() => onSelect(b)}>
                  <td>
                    <div className="b-name">{b.name}</div>
                    <div className="b-id">{b.id}</div>
                  </td>
                  <td>{b.plate}</td>
                  <td className="num">{b.askingRent}</td>
                  <td className="num">{b.propertyFee || '—'}</td>
                  <td className="num">{b.totalArea.toLocaleString('zh-CN')}</td>
                  <td className="num">{b.completionYear}</td>
                  <td>
                    <div className="occ-bar">
                      <div className="occ-track">
                        <div
                          className="occ-fill"
                          style={{
                            width: `${m.occupancyRate}%`,
                            background: occupancyColor(m.occupancyRate),
                          }}
                        />
                      </div>
                      <span className="occ-val">{m.occupancyRate}%</span>
                    </div>
                  </td>
                  <td className="num">{m.tenantCount}</td>
                  <td>
                    <div style={{ display: 'flex', gap: 6 }}>
                      <Tag color={typeColor[b.propertyType]} dot={false}>
                        {b.propertyType}
                      </Tag>
                      <Tag color={gradeColor[b.grade]} dot={false}>
                        {b.grade}
                      </Tag>
                    </div>
                  </td>
                </tr>
              ))}
              {rows.length === 0 && (
                <tr>
                  <td colSpan={9} style={{ textAlign: 'center', padding: 40, color: 'var(--text-faint)' }}>
                    没有符合条件的楼宇，试试调整筛选条件
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}

function Select({
  value,
  onChange,
  options,
}: {
  value: string
  onChange: (v: string) => void
  options: string[]
}) {
  return (
    <div className="select-wrap">
      <select value={value} onChange={(e) => onChange(e.target.value)}>
        {options.map((o) => (
          <option key={o} value={o}>
            {o === '全部' ? '全部' : o}
          </option>
        ))}
      </select>
    </div>
  )
}
