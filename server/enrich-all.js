#!/usr/bin/env node
/**
 * 一次性批量补全所有企业的工商信息（电话/法人/注册资本/统一信用代码 …）
 * ============================================================================
 *  使用方法：
 *    1) 把 .env.example 复制为 .env, 填入 TYC_TOKEN
 *    2) 运行:  node server/enrich-all.js
 *    3) 结果写入 data/enriched-tenants.json
 *    4) 刷新网页, 企业明细自动加载补全后的电话等字段
 *
 *  特性：
 *    - 进度条 + ETA
 *    - 断点续传（重复运行只补还没拿到的;已成功的不再重复请求）
 *    - 节流 1.2s/条, 避免触发对方限流
 *    - 智能检测 WAF / IP 拦截, 提前退出并提示
 *    - 不需要先启动 qcc-proxy.js, 本脚本直连天眼查
 * ============================================================================ */

'use strict';

const fs    = require('fs');
const path  = require('path');
const https = require('https');

const ROOT = path.join(__dirname, '..');

// 加载 .env
(function loadEnv() {
  const envPath = path.join(ROOT, '.env');
  if (!fs.existsSync(envPath)) return;
  fs.readFileSync(envPath, 'utf-8').split(/\r?\n/).forEach(line => {
    const m = line.match(/^\s*([A-Z_][A-Z0-9_]*)\s*=\s*(.*?)\s*$/);
    if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^["']|["']$/g, '');
  });
})();

const TYC_TOKEN = process.env.TYC_TOKEN || '';
const THROTTLE_MS = Number(process.env.THROTTLE_MS || 1200);

if (!TYC_TOKEN) {
  console.error('❌ 未设置 TYC_TOKEN。请创建 .env 文件并填入:');
  console.error('   TYC_TOKEN=你的天眼查开放平台 token');
  process.exit(1);
}

/* -------- 加载数据(同前端) -------- */

function loadDataFile(filePath) {
  // 用沙箱式 eval 直接执行 data/*.js 取出 window.COMPETITORS / window.EXTRA_TENANTS
  const code = fs.readFileSync(filePath, 'utf-8');
  const sandbox = { window: {} };
  // eslint-disable-next-line no-new-func
  new Function('window', code)(sandbox.window);
  return sandbox.window;
}

const w1 = loadDataFile(path.join(ROOT, 'data', 'competitors.js'));
const w2 = loadDataFile(path.join(ROOT, 'data', 'extra-tenants.js'));
const parks = w1.COMPETITORS || [];
const extra = w2.EXTRA_TENANTS || {};

// 合并 + 扁平化所有需要补全的企业
const allTenants = [];
parks.forEach(p => {
  if (p.excluded) return;
  const tenants = (p.tenants || []).concat(extra[p.id] || []);
  tenants.forEach(t => {
    allTenants.push({
      parkId: p.id,
      parkName: p.name,
      address: p.address,
      name: t.name
    });
  });
});

console.log(`╭─────────────────────────────────────────────╮`);
console.log(`│  天眼查批量补全 · 共 ${String(allTenants.length).padStart(3,' ')} 家企业           │`);
console.log(`│  Token: ${TYC_TOKEN.slice(0, 8)}...${TYC_TOKEN.slice(-4)}             │`);
console.log(`│  节流: ${THROTTLE_MS}ms/条                          │`);
console.log(`╰─────────────────────────────────────────────╯`);

/* -------- 断点续传：读取已有结果 -------- */

const outPath = path.join(ROOT, 'data', 'enriched-tenants.json');
let enriched = { _meta: { source: '天眼查 open.api', enrichedAt: '', count: 0 }, tenants: {} };

if (fs.existsSync(outPath)) {
  try {
    enriched = JSON.parse(fs.readFileSync(outPath, 'utf-8'));
    if (!enriched.tenants) enriched.tenants = {};
  } catch (e) {
    console.warn('⚠️  enriched-tenants.json 解析失败, 将重新生成');
  }
}

const todo = allTenants.filter(t => {
  const key = t.parkId + '|' + t.name;
  const ent = enriched.tenants[key];
  return !ent || !ent.phone;
});

console.log(`  已有: ${allTenants.length - todo.length}    待补: ${todo.length}\n`);

if (!todo.length) {
  console.log('🎉 所有企业都已补全, 无需调用 API.');
  process.exit(0);
}

/* -------- 调用天眼查 -------- */

function tycSearch(keyword) {
  return new Promise((resolve, reject) => {
    const u = `https://open.api.tianyancha.com/services/open/search/2.0?word=${encodeURIComponent(keyword)}&pageNum=1&pageSize=5`;
    const opts = {
      headers: {
        'Authorization': TYC_TOKEN,
        'User-Agent': 'Mozilla/5.0 wujiaochang-map/1.0',
        'Accept': 'application/json'
      },
      timeout: 12000
    };
    const req = https.get(u, opts, (resp) => {
      let buf = '';
      resp.on('data', d => buf += d);
      resp.on('end', () => {
        const trimmed = (buf || '').trim();
        if (!trimmed.startsWith('{')) {
          const isWaf = /HWWAF|Block-Event-Id|<title>天眼查<\/title>/i.test(buf) || resp.headers.server === 'CW';
          const err = new Error(isWaf
            ? `WAF 拦截 (HTTP ${resp.statusCode}) — 当前 IP 被天眼查防火墙封禁。请改在本地笔记本/办公网络运行。`
            : `非 JSON 响应 (HTTP ${resp.statusCode})`);
          err.code = isWaf ? 'WAF_BLOCKED' : 'BAD_RESPONSE';
          return reject(err);
        }
        try { resolve(JSON.parse(trimmed)); } catch (e) { reject(e); }
      });
    });
    req.on('timeout', () => req.destroy(new Error('timeout')));
    req.on('error', reject);
  });
}

function fmtTime(ms) {
  const s = Math.round(ms / 1000);
  if (s < 60) return s + ' 秒';
  return Math.floor(s / 60) + ' 分 ' + (s % 60) + ' 秒';
}

(async () => {
  const startedAt = Date.now();
  let ok = 0, fail = 0, noResult = 0;
  let wafBlockedConfirmed = false;

  for (let i = 0; i < todo.length; i++) {
    const t = todo[i];
    const elapsed = Date.now() - startedAt;
    const rate = (i + 1) / Math.max(1, elapsed / 1000);
    const eta = (todo.length - i) / Math.max(0.0001, rate) * 1000;
    process.stdout.write(`\r[${String(i + 1).padStart(3,' ')}/${todo.length}] ${ok}✅ ${fail}❌ ${noResult}∅  ETA ${fmtTime(eta)}   ${t.name.slice(0, 24).padEnd(24)} `);

    try {
      const data = await tycSearch(t.name);
      if (data && data.error_code === 0) {
        const item = ((data.result && data.result.items) || [])[0];
        if (item) {
          enriched.tenants[t.parkId + '|' + t.name] = {
            queryName: t.name,
            tycName: item.name || null,
            phone: item.telephone || null,
            email: item.email || null,
            legalPerson: item.legalPersonName || null,
            regCapital: item.regCapital || null,
            regTime: item.estiblishTime || null,
            city: item.city || null,
            regNumber: item.regNumber || item.creditCode || null,
            regStatus: item.regStatus || null,
            businessScope: item.businessScope || null
          };
          ok++;
        } else {
          enriched.tenants[t.parkId + '|' + t.name] = { queryName: t.name, error: 'no match' };
          noResult++;
        }
      } else {
        fail++;
        if (data && (data.error_code === 300001 || /token/i.test(data.reason || ''))) {
          console.error(`\n❌ Token 异常: ${data.reason || data.message}`);
          process.exit(2);
        }
      }
    } catch (e) {
      fail++;
      if (e.code === 'WAF_BLOCKED') {
        if (!wafBlockedConfirmed) {
          console.error(`\n\n❌ ${e.message}\n`);
          console.error('   解决办法:');
          console.error('   1) 在你自己的笔记本上运行此脚本（最简单）');
          console.error('   2) 找一台国内非云服务器的机器运行');
          console.error('   3) 联系天眼查申请把当前服务器 IP 加入白名单\n');
          wafBlockedConfirmed = true;
        }
        if (i > 5) break; // 多次失败就提前退出
      }
    }

    // 每 10 条写一次盘,防止意外丢失
    if ((i + 1) % 10 === 0) writeOut();
    await new Promise(r => setTimeout(r, THROTTLE_MS));
  }

  writeOut();
  console.log(`\n\n  完成 → 成功 ${ok} · 失败 ${fail} · 无匹配 ${noResult}`);
  console.log(`  结果已写入: ${path.relative(ROOT, outPath)}\n`);

  function writeOut() {
    enriched._meta = {
      source: '天眼查 open.api',
      enrichedAt: new Date().toISOString(),
      count: Object.keys(enriched.tenants).length
    };
    fs.writeFileSync(outPath, JSON.stringify(enriched, null, 2), 'utf-8');
  }
})();
