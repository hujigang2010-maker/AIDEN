/**
 * 企查查 / 天眼查 企业信息代理服务（零依赖,纯 Node.js）
 * ============================================================================
 *  浏览器无法直接调用 天眼查 / 企查查 的开放 API（CORS + 暴露 Key 风险）,
 *  本脚本提供一个最简代理：前端调 /api/enrich?company=XXX&address=YYY,
 *  本服务带着你的 API Key/Token 去查 天眼查（优先）或 企查查,
 *  返回标准化的 { phone, regCapital, legalPerson, ... } JSON。
 *
 *  启动：
 *    1) 申请 API:
 *       - 天眼查开放平台 https://open.tianyancha.com（推荐,接口最简洁）
 *       - 企查查开放平台 https://openapi.qcc.com
 *    2) 设置环境变量后启动:
 *       export TYC_TOKEN="你的天眼查 Bearer Token"          # 天眼查
 *       export QCC_KEY="你的企查查 AppKey"                  # 企查查（备选）
 *       export QCC_SECRET="你的企查查 SecretKey"            # 企查查
 *       node server/qcc-proxy.js
 *
 *    默认监听 http://localhost:3001,可用 PORT=xxxx 改端口。
 *
 *  前端在网页右上角【⚙ API 配置】中填入 http://localhost:3001 即可。
 * ============================================================================ */

'use strict';

const http  = require('http');
const https = require('https');
const crypto = require('crypto');
const { URL } = require('url');

const PORT       = Number(process.env.PORT || 3001);
const TYC_TOKEN  = process.env.TYC_TOKEN  || '';
const QCC_KEY    = process.env.QCC_KEY    || '';
const QCC_SECRET = process.env.QCC_SECRET || '';

/* -------- 小工具 -------- */

function fetchJson(url, headers = {}) {
  return new Promise((resolve, reject) => {
    const req = https.get(url, { headers, timeout: 8000 }, (resp) => {
      let buf = '';
      resp.on('data', (d) => (buf += d));
      resp.on('end', () => {
        try { resolve(JSON.parse(buf)); }
        catch (e) { reject(new Error('JSON parse failed: ' + e.message + ' | body=' + buf.slice(0, 300))); }
      });
    });
    req.on('timeout', () => { req.destroy(new Error('timeout')); });
    req.on('error', reject);
  });
}

function sendJson(res, status, body) {
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin':  '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  });
  res.end(JSON.stringify(body));
}

/* -------- 天眼查（推荐） -------- */
// 文档: https://open.tianyancha.com/open/910
async function tycLookup(keyword) {
  if (!TYC_TOKEN) return null;
  const url = `https://open.api.tianyancha.com/services/open/search/2.0?word=${encodeURIComponent(keyword)}&pageNum=1&pageSize=5`;
  const data = await fetchJson(url, { Authorization: TYC_TOKEN });
  if (data && data.error_code === 0) {
    const item = ((data.result && data.result.items) || [])[0];
    if (item) {
      return {
        source: 'tianyancha',
        name: item.name,
        phone: item.telephone || null,
        email: item.email || null,
        legalPerson: item.legalPersonName || null,
        regCapital: item.regCapital || null,
        regTime: item.estiblishTime || null,
        city: item.city || null,
        regNumber: item.regNumber || item.creditCode || null,
        businessScope: item.businessScope || null,
        url: item.companyUrl || null
      };
    }
  }
  return { source: 'tianyancha', error: (data && (data.reason || data.message)) || 'no result' };
}

/* -------- 企查查（备选） -------- */
// 文档: https://openapi.qcc.com/dataApi
// 鉴权: header.Token = MD5(AppKey + Timestamp + SecretKey)
async function qccLookup(keyword) {
  if (!QCC_KEY || !QCC_SECRET) return null;
  const ts  = Math.floor(Date.now() / 1000);
  const sig = crypto.createHash('md5').update(QCC_KEY + ts + QCC_SECRET).digest('hex').toUpperCase();
  const url = `https://api.qichacha.com/ECIV4/Search?key=${QCC_KEY}&keyword=${encodeURIComponent(keyword)}`;
  const data = await fetchJson(url, { Token: sig, Timespan: String(ts) });
  if (data && data.Status === '200') {
    const item = (data.Result || [])[0];
    if (item) {
      return {
        source: 'qichacha',
        name: item.Name,
        phone: item.PhoneNumber || null,
        legalPerson: item.OperName || null,
        regCapital: item.RegistCapi || null,
        regTime: item.StartDate || null,
        regNumber: item.CreditCode || item.No || null
      };
    }
  }
  return { source: 'qichacha', error: (data && data.Message) || 'no result' };
}

/* -------- 主路由 -------- */

const server = http.createServer(async (req, res) => {
  // CORS preflight
  if (req.method === 'OPTIONS') {
    res.writeHead(204, {
      'Access-Control-Allow-Origin':  '*',
      'Access-Control-Allow-Methods': 'GET, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
    return res.end();
  }

  const u = new URL(req.url, 'http://localhost');

  // 健康检查
  if (u.pathname === '/' || u.pathname === '/health') {
    return sendJson(res, 200, {
      ok: true,
      message: '企业信息代理运行中',
      providers: {
        tianyancha: TYC_TOKEN ? 'configured' : 'not configured (set TYC_TOKEN)',
        qichacha:   (QCC_KEY && QCC_SECRET) ? 'configured' : 'not configured (set QCC_KEY/QCC_SECRET)'
      },
      endpoints: ['/api/enrich?company=XXX', '/api/search?company=XXX']
    });
  }

  // 单条企业补充
  if (u.pathname === '/api/enrich' || u.pathname === '/api/search') {
    const company = u.searchParams.get('company') || u.searchParams.get('name') || '';
    if (!company.trim()) return sendJson(res, 400, { error: 'company 参数必填' });

    try {
      let result = await tycLookup(company);
      if (!result || result.error) {
        const alt = await qccLookup(company);
        if (alt && !alt.error) result = alt;
      }
      if (!result) {
        return sendJson(res, 200, { error: '未配置任何 API,请先设置 TYC_TOKEN 或 QCC_KEY/SECRET 环境变量。' });
      }
      return sendJson(res, 200, result);
    } catch (e) {
      return sendJson(res, 500, { error: e.message });
    }
  }

  sendJson(res, 404, { error: 'not found' });
});

server.listen(PORT, () => {
  console.log(`╭─────────────────────────────────────────────────────────╮`);
  console.log(`│  企业信息代理已启动 → http://localhost:${PORT}              `);
  console.log(`├─────────────────────────────────────────────────────────┤`);
  console.log(`│  天眼查 (TYC_TOKEN):   ${TYC_TOKEN  ? '✅ 已配置' : '❌ 未设置'}`);
  console.log(`│  企查查 (QCC_KEY):     ${(QCC_KEY && QCC_SECRET) ? '✅ 已配置' : '❌ 未设置'}`);
  console.log(`╰─────────────────────────────────────────────────────────╯`);
  console.log(`  健康检查:  curl http://localhost:${PORT}/health`);
  console.log(`  查询企业:  curl 'http://localhost:${PORT}/api/enrich?company=合生创展'`);
  console.log(``);
  console.log(`  在网页右上角点击【⚙ API 配置】,代理地址填:  http://localhost:${PORT}`);
});
