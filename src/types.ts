// 物业类型
export type PropertyType = '写字楼' | '产业园' | '租赁住宅' | '混合'

// 物业等级
export type PropertyGrade = '超甲级' | '甲级' | '乙级' | '丙级' | '其它'

// 企业经营状态
export type TenantStatus = '正常经营' | '新签约' | '筹备装修' | '欠租预警' | '已退租'

// 入驻企业（楼宇 -> 多企业）
export interface Tenant {
  /** 入驻企业名称 */
  name: string
  /** 入驻企业所属行业 */
  industry: string
  /** 入驻楼层 */
  floor: string
  /** 租赁面积（㎡） */
  leaseArea: number
  /** 成交租金（元/㎡/天） */
  dealRent: number
  /** 免租期（月） */
  rentFreeMonths: number
  /** 企业经营状态 */
  status: TenantStatus
  /** 剩余租期（月） */
  remainingMonths: number
}

// 楼宇
export interface Building {
  id: string
  /** 楼宇名称 */
  name: string
  /** 所属板块 */
  plate: string
  /** 物业类型 */
  propertyType: PropertyType
  /** 物业等级 */
  grade: PropertyGrade
  /** 报价租金（元/㎡/天） */
  askingRent: number
  /** 物业费（元/㎡/月） */
  propertyFee: number
  /** 楼宇总建筑面积（㎡） */
  totalArea: number
  /** 竣工时间（年份） */
  completionYear: number
  /** 标准层层高（米） */
  floorHeight: number
  /** 经度 */
  lng: number
  /** 纬度 */
  lat: number
  /** 入驻企业列表 */
  tenants: Tenant[]
}

export interface BuildingMetrics {
  /** 已成交租金均值（元/㎡/天） */
  avgDealRent: number
  /** 总租赁面积（㎡） */
  leasedArea: number
  /** 出租率（%） */
  occupancyRate: number
  /** 入驻企业数 */
  tenantCount: number
}
