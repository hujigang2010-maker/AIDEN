/**
 * 企查查 / 天眼查 企业信息代理服务（零依赖,纯 Node.js）
 * ============================================================================
 *  浏览器无法直接调用 天眼查 / 企查查 的开放 API（CORS + 暴露 Key 风险）,
 *  本脚本提供一个最简代理：前端调 /api/enrich?company=XXX&address=YYY,
 *  本服务带着你的 API Key/Token 去查 天眼查（优先）或 企查查,
 *  返回标准化的 { phone, regCapital, legalPerson, ... } JSON。
 *
 *  ⚠️ 重要：天眼查 WAF 会封大多数云服务器 IP（包括阿里云海外/AWS/腾讯云海外
 *     /Cursor Cloud Agent 等）。在这些环境下会返回 418 + HTML 页面而非 JSON。
 *     建议在你自己的笔记本 / 公司办公网络 IP 下运行本代理。
 *
 *  启动：
 *    1) 申请 API:
 *       - 天眼查开放平台 https://open.tianyancha.com（推荐,接口最简洁）
 *       - 企查查开放平台 https://openapi.qcc.com
 *    2) 设置环境变量后启动:
 *       cp .env.example .env             # 然后把 token 填进去
 *       node --env-file=.env server/qcc-proxy.js
 *       # 或:
 *       export TYC_TOKEN="你的天眼查 Token"
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

// 简易 .env 加载(无依赖)
try {
  const fs = require('fs'); const path = require('path');
  const envPath = path.join(__dirname, '..', '.env');
  if (fs.existsSync(envPath)) {
    fs.readFileSync(envPath, 'utf-8').split(/\r?\n/).forEach(line => {
      const m = line.match(/^\s*([A-Z_][A-Z0-9_]*)\s*=\s*(.*?)\s*$/);
      if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^["']|["']$/g, '');
    });
  }
} catch (e) {}

const PORT       = Number(process.env.PORT || 3001);
const TYC_TOKEN  = process.env.TYC_TOKEN  || '';
const QCC_KEY    = process.env.QCC_KEY    || '';
const QCC_SECRET = process.env.QCC_SECRET || '';

/* -------- 小工具 -------- */

function fetchRaw(url, headers = {}) {
  return new Promise((resolve, reject) => {
    const req = https.get(url, { headers, timeout: 10000 }, (resp) => {
      let buf = '';
      resp.on('data', (d) => (buf += d));
      resp.on('end', () => resolve({ status: resp.statusCode, body: buf, headers: resp.headers }));
    });
    req.on('timeout', () => { req.destroy(new Error('timeout')); });
    req.on('error', reject);
  });
}

function fetchJson(url, headers = {}) {
  return fetchRaw(url, headers).then((r) => {
    const trimmed = (r.body || '').trim();
    if (!trimmed.startsWith('{') && !trimmed.startsWith('[')) {
      // WAF 拦截或异常响应
      const isWaf = /HWWAF|Block-Event-Id|<title>天眼查<\/title>/i.test(r.body) || r.headers.server === 'CW';
      const err = new Error(isWaf
        ? `WAF blocked (HTTP ${r.status}) — 当前服务器 IP 可能被天眼查防火墙拦截。请改在本地笔记本/办公网络 IP 下运行。`
        : `Non-JSON response (HTTP ${r.status}): ${r.body.slice(0, 200)}`);
      err.code = isWaf ? 'WAF_BLOCKED' : 'BAD_RESPONSE';
      err.httpStatus = r.status;
      throw err;
    }
    return JSON.parse(trimmed);
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
  const data = await fetchJson(url, {
    'Authorization': TYC_TOKEN,
    'User-Agent': 'Mozilla/5.0 wujiaochang-map/1.0'
  });
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
        url: item.companyUrl || null,
        regStatus: item.regStatus || null
      };
    }
    return { source: 'tianyancha', error: 'no match' };
  }
  return { source: 'tianyancha', error: (data && (data.reason || data.message)) || 'unknown' };
}

/* -------- 企查查（备选） -------- */
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
  if (req.method === 'OPTIONS') {
    res.writeHead(204, {
      'Access-Control-Allow-Origin':  '*',
      'Access-Control-Allow-Methods': 'GET, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
    return res.end();
  }

  const u = new URL(req.url, 'http://localhost');

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
      return sendJson(res, e.code === 'WAF_BLOCKED' ? 451 : 500, {
        error: e.message,
        code: e.code || 'ERR',
        hint: e.code === 'WAF_BLOCKED'
          ? '天眼查 WAF 拦截了当前服务器 IP。请改在本地笔记本/办公网络 IP 下运行 (常见于 AWS/阿里云海外/Cursor Cloud Agent 等环境)。'
          : undefined
      });
    }
  }

  sendJson(res, 404, { error: 'not found' });
});

server.listen(PORT, () => {
  console.log(`╭─────────────────────────────────────────────────────────╮`);
  console.log(`│  企业信息代理已启动 → http://localhost:${PORT}              `);
  console.log(`├─────────────────────────────────────────────────────────┤`);
  console.log(`│  天眼查 (TYC_TOKEN):   ${TYC_TOKEN  ? '✅ 已配置 (' + TYC_TOKEN.slice(0, 8) + '...)' : '❌ 未设置'}`);
  console.log(`│  企查查 (QCC_KEY):     ${(QCC_KEY && QCC_SECRET) ? '✅ 已配置' : '❌ 未设置'}`);
  console.log(`╰─────────────────────────────────────────────────────────╯`);
  console.log(`  健康检查:  curl http://localhost:${PORT}/health`);
  console.log(`  查询企业:  curl 'http://localhost:${PORT}/api/enrich?company=合生创展'`);
  console.log(``);
  console.log(`  ⚠️  若返回 451 / WAF_BLOCKED, 说明当前 IP 被天眼查 WAF 拦截,`);
  console.log(`     请改在本地笔记本 / 办公网络 IP 下运行（不要用云服务器）。`);
  console.log(``);
  console.log(`  网页右上角【⚙ API 配置】, 代理地址填:  http://localhost:${PORT}`);
});
