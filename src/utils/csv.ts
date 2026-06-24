import type { Building } from '../types'
import { computeMetrics } from './metrics'

function esc(v: string | number): string {
  const s = String(v ?? '')
  if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"'
  return s
}

/** 楼宇 + 入驻企业明细表（每个企业一行，楼宇信息重复） */
export function buildDetailCsv(data: Building[]): string {
  const header = [
    '楼宇编号',
    '楼宇名称',
    '详细地址',
    '所属板块',
    '物业类型',
    '物业等级',
    '报价租金(元/㎡/天)',
    '物业费(元/㎡/月)',
    '楼宇总建筑面积(㎡)',
    '竣工时间',
    '标准层层高(米)',
    '出租率(%)',
    '入驻企业名称',
    '入驻企业所属行业',
    '入驻楼层',
    '租赁面积(㎡)',
    '成交租金(元/㎡/天)',
    '免租期(月)',
    '企业经营状态',
    '剩余租期(月)',
  ]
  const rows: string[] = [header.join(',')]
  data.forEach((b) => {
    const m = computeMetrics(b)
    if (b.tenants.length === 0) {
      rows.push(
        [b.id, b.name, b.address ?? '', b.plate, b.propertyType, b.grade, b.askingRent, b.propertyFee, b.totalArea, b.completionYear, b.floorHeight, m.occupancyRate, '', '', '', '', '', '', '', '']
          .map(esc)
          .join(','),
      )
    }
    b.tenants.forEach((t) => {
      rows.push(
        [
          b.id, b.name, b.address ?? '', b.plate, b.propertyType, b.grade, b.askingRent, b.propertyFee,
          b.totalArea, b.completionYear, b.floorHeight, m.occupancyRate,
          t.name, t.industry, t.floor, t.leaseArea, t.dealRent, t.rentFreeMonths, t.status, t.remainingMonths,
        ]
          .map(esc)
          .join(','),
      )
    })
  })
  return rows.join('\r\n')
}

/** 楼宇汇总表（每栋楼宇一行） */
export function buildSummaryCsv(data: Building[]): string {
  const header = [
    '楼宇编号', '楼宇名称', '详细地址', '所属板块', '物业类型', '物业等级',
    '报价租金(元/㎡/天)', '平均成交租金(元/㎡/天)', '物业费(元/㎡/月)',
    '楼宇总建筑面积(㎡)', '已租面积(㎡)', '出租率(%)', '入驻企业数',
    '竣工时间', '标准层层高(米)',
  ]
  const rows = [header.join(',')]
  data.forEach((b) => {
    const m = computeMetrics(b)
    rows.push(
      [
        b.id, b.name, b.address ?? '', b.plate, b.propertyType, b.grade,
        b.askingRent, m.avgDealRent, b.propertyFee,
        b.totalArea, m.leasedArea, m.occupancyRate, m.tenantCount,
        b.completionYear, b.floorHeight,
      ]
        .map(esc)
        .join(','),
    )
  })
  return rows.join('\r\n')
}

export function downloadCsv(filename: string, csv: string) {
  // 加 BOM 以便 Excel 正确识别 UTF-8 中文
  const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}
