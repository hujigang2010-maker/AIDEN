import { copyFileSync, existsSync, mkdirSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { buildings } from '../src/data/buildings'
import { buildDetailCsv, buildSummaryCsv } from '../src/utils/csv'

const __dirname = dirname(fileURLToPath(import.meta.url))
const releaseDir = resolve(__dirname, '../release')
const outDir = resolve(releaseDir, 'tables')
mkdirSync(outDir, { recursive: true })

// 将单文件构建产物复制为友好的中文文件名（可直接双击在浏览器打开）
const builtHtml = resolve(releaseDir, 'standalone/index.html')
if (existsSync(builtHtml)) {
  copyFileSync(builtHtml, resolve(releaseDir, '杨浦区楼宇经济数据平台.html'))
  console.log('✓ 已生成网页打开版本：release/杨浦区楼宇经济数据平台.html')
}

const BOM = '\ufeff'
const detail = BOM + buildDetailCsv(buildings)
const summary = BOM + buildSummaryCsv(buildings)

writeFileSync(resolve(outDir, '杨浦区楼宇企业明细表.csv'), detail, 'utf8')
writeFileSync(resolve(outDir, '杨浦区楼宇汇总表.csv'), summary, 'utf8')

const tenantRows = buildings.reduce((s, b) => s + b.tenants.length, 0)
console.log(`✓ 已导出 CSV 至 release/tables/`)
console.log(`  · 杨浦区楼宇汇总表.csv   （${buildings.length} 栋楼宇）`)
console.log(`  · 杨浦区楼宇企业明细表.csv （${tenantRows} 条入驻企业记录）`)
