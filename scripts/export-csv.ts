import { copyFileSync, existsSync, mkdirSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { buildings } from '../src/data/buildings'
import { buildDetailCsv, buildSummaryCsv } from '../src/utils/csv'

const __dirname = dirname(fileURLToPath(import.meta.url))
const releaseDir = resolve(__dirname, '../release')
const outDir = resolve(releaseDir, 'tables')
mkdirSync(outDir, { recursive: true })

// 将单文件构建产物复制为友好文件名（可直接双击在浏览器打开）
// 同时生成 ASCII 文件名副本，方便生成干净、可直接点击的下载/预览链接
const builtHtml = resolve(releaseDir, 'standalone/index.html')
if (existsSync(builtHtml)) {
  copyFileSync(builtHtml, resolve(releaseDir, '杨浦区楼宇经济数据平台.html'))
  copyFileSync(builtHtml, resolve(releaseDir, 'yangpu-building-platform.html'))
  console.log('✓ 已生成网页打开版本：release/杨浦区楼宇经济数据平台.html (+ yangpu-building-platform.html)')
}

const BOM = '\ufeff'
const detail = BOM + buildDetailCsv(buildings)
const summary = BOM + buildSummaryCsv(buildings)

// 简版（楼宇汇总，每栋一行）与详细版（企业明细，每个企业一行）
writeFileSync(resolve(outDir, '杨浦区楼宇数据_简版.csv'), summary, 'utf8')
writeFileSync(resolve(outDir, '杨浦区楼宇数据_详细版.csv'), detail, 'utf8')
// ASCII 文件名副本（便于干净的下载链接）
writeFileSync(resolve(outDir, 'yangpu-buildings-summary-simple.csv'), summary, 'utf8')
writeFileSync(resolve(outDir, 'yangpu-buildings-detail-full.csv'), detail, 'utf8')

const tenantRows = buildings.reduce((s, b) => s + b.tenants.length, 0)
const office = buildings.filter((b) => b.propertyType === '写字楼').length
const park = buildings.filter((b) => b.propertyType === '产业园').length
const mixed = buildings.filter((b) => b.propertyType === '混合').length
console.log(`✓ 已导出 CSV 至 release/tables/`)
console.log(`  · 简版  杨浦区楼宇数据_简版.csv   （${buildings.length} 栋楼宇：写字楼 ${office} / 产业园 ${park} / 混合 ${mixed}）`)
console.log(`  · 详细版 杨浦区楼宇数据_详细版.csv （${tenantRows} 条入驻企业记录）`)
