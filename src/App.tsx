import { useState } from 'react'
import { buildings } from './data/buildings'
import type { Building } from './types'
import Dashboard from './components/Dashboard'
import BuildingTable, { useTableFilters } from './components/BuildingTable'
import MapView from './components/MapView'
import DetailDrawer from './components/DetailDrawer'
import { IconBuilding, IconDashboard, IconMap, IconTable } from './components/ui'

type Tab = 'dashboard' | 'table' | 'map'

const tabs: { id: Tab; label: string; icon: React.ReactNode }[] = [
  { id: 'dashboard', label: '数据总览', icon: <IconDashboard /> },
  { id: 'table', label: '楼宇数据', icon: <IconTable /> },
  { id: 'map', label: '地图分布', icon: <IconMap /> },
]

export default function App() {
  const [tab, setTab] = useState<Tab>('dashboard')
  const [selected, setSelected] = useState<Building | null>(null)
  const filters = useTableFilters()

  return (
    <div className="app">
      <header className="app-header">
        <div className="brand">
          <div className="logo">
            <IconBuilding />
          </div>
          <div>
            <h1>杨浦区楼宇经济数据平台</h1>
            <p>商业写字楼 + 工业/产业楼宇全维度 · 楼宇 {buildings.length} 栋 · 简版 / 详细版可一键导出</p>
          </div>
        </div>
        <nav className="nav-tabs">
          {tabs.map((t) => (
            <button
              key={t.id}
              className={tab === t.id ? 'active' : ''}
              onClick={() => setTab(t.id)}
            >
              {t.icon}
              <span>{t.label}</span>
            </button>
          ))}
        </nav>
      </header>

      <main className="main">
        {tab === 'dashboard' && <Dashboard data={buildings} />}
        {tab === 'table' && (
          <BuildingTable data={buildings} filters={filters} onSelect={setSelected} />
        )}
        {tab === 'map' && <MapView data={buildings} onSelect={setSelected} />}
      </main>

      <footer className="footer">
        数据综合自政府公开文件 / 园区公示 / 主流商办平台公开行情，仅供参考 · 共 {buildings.length} 栋楼宇 · 地图服务由
        <a href="https://lbs.amap.com" target="_blank" rel="noreferrer"> 高德地图 </a>
        提供
      </footer>

      <DetailDrawer building={selected} onClose={() => setSelected(null)} />
    </div>
  )
}
