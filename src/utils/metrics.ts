import type { Building, BuildingMetrics } from '../types'

export function computeMetrics(b: Building): BuildingMetrics {
  const active = b.tenants.filter((t) => t.status !== '已退租')
  const leasedArea = active.reduce((s, t) => s + t.leaseArea, 0)
  const dealRents = active.map((t) => t.dealRent).filter((r) => r > 0)
  const avgDealRent = dealRents.length
    ? dealRents.reduce((s, r) => s + r, 0) / dealRents.length
    : 0
  const occupancyRate = b.totalArea ? Math.min(100, (leasedArea / b.totalArea) * 100) : 0
  return {
    avgDealRent: round(avgDealRent, 2),
    leasedArea,
    occupancyRate: round(occupancyRate, 1),
    tenantCount: active.length,
  }
}

export function round(n: number, digits = 1): number {
  const f = 10 ** digits
  return Math.round(n * f) / f
}

export function formatArea(n: number): string {
  return n.toLocaleString('zh-CN') + ' ㎡'
}

export function formatWan(n: number): string {
  return (n / 10000).toLocaleString('zh-CN', { maximumFractionDigits: 1 }) + ' 万㎡'
}

export const occupancyColor = (rate: number): string => {
  if (rate >= 80) return '#22c55e'
  if (rate >= 50) return '#38bdf8'
  if (rate >= 30) return '#f59e0b'
  return '#ef4444'
}

export const gradeColor: Record<string, string> = {
  超甲级: '#a855f7',
  甲级: '#38bdf8',
  乙级: '#22c55e',
  丙级: '#f59e0b',
  其它: '#94a3b8',
}

export const typeColor: Record<string, string> = {
  写字楼: '#38bdf8',
  产业园: '#22c55e',
  租赁住宅: '#f472b6',
  混合: '#a855f7',
}

export const statusColor: Record<string, string> = {
  正常经营: '#22c55e',
  新签约: '#38bdf8',
  筹备装修: '#f59e0b',
  欠租预警: '#ef4444',
  已退租: '#64748b',
}
