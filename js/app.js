/* ==========================================================================
 * 杨浦五角场 5km 商办竞品地图 - 业务脚本
 * --------------------------------------------------------------------------
 *  - 初始化高德地图 + 标准/卫星图切换
 *  - 绘制基准点、4 个距离圈层
 *  - 计算每个竞品的距离并归属圈层
 *  - 渲染【项目概览】+【企业明细】两个 Tab
 *  - 企业明细包含：办公体量 / 成交租金 / 免租期 / 剩余租期 / 主力客户行业
 *                 物业费 / 性质说明 / 电话
 *  - 电话：localStorage 手动录入,或通过本地代理调用 天眼查 / 企查查 批量补充
 *  - 实现地图 ↔ 表格 双向联动
 * ========================================================================== */

(function () {
  'use strict';

  /* ====================================================================== */
  /* ============================  常量 & 工具  ============================ */
  /* ====================================================================== */

  const TIER_CONFIG = {
    '1km':    { label: '1 km 内', min: 0,    max: 1000, color: '#1976d2', cls: 'tier-1' },
    '1-2km':  { label: '1-2 km',  min: 1000, max: 2000, color: '#2e7d32', cls: 'tier-2' },
    '2-3km':  { label: '2-3 km',  min: 2000, max: 3000, color: '#ef6c00', cls: 'tier-3' },
    '3-5km':  { label: '3-5 km',  min: 3000, max: 5000, color: '#c62828', cls: 'tier-4' },
    'out':    { label: '>5 km',   min: 5000, max: 1e9,  color: '#9e9e9e', cls: 'tier-out' }
  };

  const CAT_ICON = {
    '商办写字楼': '🏢',
    '产业园区':   '🏭',
    '众创空间':   '💡',
    '商业综合体': '🛍️'
  };

  const LS_PHONE_KEY = 'wjc:tenant-phones:v1';   // localStorage 电话表
  const LS_API_KEY   = 'wjc:api-proxy-url:v1';   // localStorage 代理 URL

  function haversine(lng1, lat1, lng2, lat2) {
    const R = 6371000;
    const toRad = (d) => d * Math.PI / 180;
    const dLat = toRad(lat2 - lat1);
    const dLng = toRad(lng2 - lng1);
    const a = Math.sin(dLat / 2) ** 2 +
              Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
              Math.sin(dLng / 2) ** 2;
    return 2 * R * Math.asin(Math.sqrt(a));
  }

  function tierOf(distance) {
    if (distance <= 1000) return '1km';
    if (distance <= 2000) return '1-2km';
    if (distance <= 3000) return '2-3km';
    if (distance <= 5000) return '3-5km';
    return 'out';
  }

  function fmtDist(meters) {
    if (meters < 1000) return Math.round(meters) + ' m';
    return (meters / 1000).toFixed(2) + ' km';
  }

  function fmtArea(a) {
    if (a == null || a === '') return '—';
    return Number(a).toLocaleString('zh-CN');
  }

  function fmtRent(r) {
    if (r == null || r === '') return '—';
    if (r === 0) return '0';
    return Number(r).toFixed(1);
  }

  function fmtMonths(m) {
    if (m == null || m === '') return '—';
    return m + ' 月';
  }

  function escapeHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  /* ---------- 租约计算 ---------- */

  function computeLease(t) {
    // 返回 { startStr, endStr, remainingMonths, status: 'active'|'expiring'|'expired'|'unknown' }
    if (!t.leaseStart || !t.leaseTerm) return { status: 'unknown' };
    const start = new Date(t.leaseStart);
    if (isNaN(+start)) return { status: 'unknown' };
    const end = new Date(start);
    end.setFullYear(end.getFullYear() + Number(t.leaseTerm));
    const now = new Date();
    const ms = end - now;
    if (ms <= 0) {
      return { startStr: ymd(start), endStr: ymd(end), remainingMonths: 0, status: 'expired' };
    }
    const months = Math.round(ms / (30.44 * 24 * 3600 * 1000));
    return {
      startStr: ymd(start),
      endStr: ymd(end),
      remainingMonths: months,
      status: months <= 6 ? 'expiring' : 'active'
    };
  }
  function ymd(d) {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }
  function fmtRemaining(t) {
    const lease = computeLease(t);
    if (lease.status === 'unknown') return { html: '<span class="lease-unknown">—</span>', text: '—' };
    if (lease.status === 'expired') {
      return {
        html: `<span class="lease-pill lease-expired" title="${lease.startStr} → ${lease.endStr}">已到期</span>`,
        text: '已到期'
      };
    }
    const m = lease.remainingMonths;
    const ys = Math.floor(m / 12), mo = m % 12;
    const label = (ys ? ys + ' 年' : '') + (mo ? mo + ' 月' : (ys ? '' : '0 月'));
    const cls = lease.status === 'expiring' ? 'lease-expiring' : 'lease-active';
    return {
      html: `<span class="lease-pill ${cls}" title="${lease.startStr} → ${lease.endStr}">${label}</span>`,
      text: label + ` (${lease.startStr}→${lease.endStr})`
    };
  }

  /* ---------- localStorage 工具 ---------- */

  function loadPhones() {
    try { return JSON.parse(localStorage.getItem(LS_PHONE_KEY) || '{}'); } catch (e) { return {}; }
  }
  function savePhones(obj) {
    try { localStorage.setItem(LS_PHONE_KEY, JSON.stringify(obj)); } catch (e) {}
  }
  function phoneKey(parkId, tenantName) { return parkId + '|' + tenantName; }

  function getPhone(park, tenant) {
    const stored = loadPhones();
    return stored[phoneKey(park.id, tenant.name)] || tenant.phone || '';
  }
  function setPhone(park, tenant, value, extraEnrich) {
    const stored = loadPhones();
    const k = phoneKey(park.id, tenant.name);
    if (value && value.trim()) stored[k] = value.trim();
    else delete stored[k];
    if (extraEnrich) {
      Object.assign(tenant, extraEnrich);
    }
    savePhones(stored);
  }

  function getApiProxy() {
    return localStorage.getItem(LS_API_KEY) || '';
  }
  function setApiProxy(url) {
    if (url) localStorage.setItem(LS_API_KEY, url.replace(/\/+$/, ''));
    else localStorage.removeItem(LS_API_KEY);
  }

  /* ====================================================================== */
  /* ==============================  数据加工  ============================= */
  /* ====================================================================== */

  const base = window.BASE_POINT;
  const rawData = (window.COMPETITORS || []).filter(d => !d.excluded);

  // 合并 extra-tenants.js 中的扩充企业
  if (window.EXTRA_TENANTS) {
    rawData.forEach(park => {
      const extras = window.EXTRA_TENANTS[park.id];
      if (extras && extras.length) {
        park.tenants = (park.tenants || []).concat(extras);
      }
    });
  }

  // 计算距离 + 圈层
  rawData.forEach((d) => {
    d.distance = haversine(base.lng, base.lat, d.lng, d.lat);
    d.tier = tierOf(d.distance);
  });
  rawData.sort((a, b) => a.distance - b.distance);
  rawData.forEach((d, i) => { d.idx = i + 1; });

  // 字段归一化（phone / leaseStart / leaseTerm 缺失则补默认）
  function deterministicMonth(name) {
    let h = 0;
    for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) | 0;
    return ((Math.abs(h) % 12) + 1).toString().padStart(2, '0');
  }
  rawData.forEach(park => {
    (park.tenants || []).forEach(t => {
      if (!('phone' in t)) t.phone = null;
      if (!t.leaseStart && t.dealYear) {
        t.leaseStart = `${t.dealYear}-${deterministicMonth(t.name)}-01`;
      }
      if (!t.leaseTerm) {
        t.leaseTerm = ((t.rent || 0) >= 6 && (t.dealYear || 0) >= 2023) ? 5 : 3;
      }
    });
  });

  // 扁平化租户
  const tenantRows = [];
  rawData.forEach((park) => {
    if (park.tier === 'out') return;
    (park.tenants || []).forEach((t) => {
      tenantRows.push({ park, tenant: t, rowIdx: 0 });
    });
  });
  tenantRows.sort((a, b) => {
    if (a.park.distance !== b.park.distance) return a.park.distance - b.park.distance;
    if (a.park.id !== b.park.id) return a.park.id < b.park.id ? -1 : 1;
    return (b.tenant.area || 0) - (a.tenant.area || 0);
  });
  tenantRows.forEach((r, i) => { r.rowIdx = i + 1; });

  /* ====================================================================== */
  /* ================================ 状态 ================================= */
  /* ====================================================================== */

  const state = {
    tab: 'projects',
    filterCat: 'all',
    filterTier: 'all',
    search: '',
    activeId: null,
    mapMode: 'standard'   // standard | satellite
  };

  function matchKeyword(hay, kw) {
    if (!kw) return true;
    return hay.toLowerCase().indexOf(kw.toLowerCase()) !== -1;
  }

  function getFilteredProjects() {
    const kw = state.search.trim();
    return rawData.filter((d) => {
      if (d.tier === 'out') return false;
      if (state.filterCat !== 'all' && d.category !== state.filterCat) return false;
      if (state.filterTier !== 'all' && d.tier !== state.filterTier) return false;
      if (kw) {
        const hay = [d.name, d.address, (d.tags || []).join(' '),
                     (d.subs || []).join(' '), d.developer || '', d.nature || '',
                     (d.tenants || []).map(t => `${t.name} ${t.industry}`).join(' ')]
                    .join(' ');
        if (!matchKeyword(hay, kw)) return false;
      }
      return true;
    });
  }

  function getFilteredTenants() {
    const kw = state.search.trim();
    return tenantRows.filter((r) => {
      const p = r.park, t = r.tenant;
      if (p.tier === 'out') return false;
      if (state.filterCat !== 'all' && p.category !== state.filterCat) return false;
      if (state.filterTier !== 'all' && p.tier !== state.filterTier) return false;
      if (kw) {
        const hay = [p.name, p.address, p.nature || '',
                     (p.tags || []).join(' '), p.developer || '',
                     t.name, t.industry || '', t.note || '',
                     getPhone(p, t)].join(' ');
        if (!matchKeyword(hay, kw)) return false;
      }
      return true;
    });
  }

  /* ====================================================================== */
  /* ===============================  地图  ================================ */
  /* ====================================================================== */

  const map = new AMap.Map('map', {
    zoom: 13.5,
    center: [base.lng, base.lat],
    viewMode: '2D',
    mapStyle: 'amap://styles/whitesmoke'
  });

  AMap.plugin(['AMap.Scale', 'AMap.ToolBar'], () => {
    map.addControl(new AMap.Scale());
    map.addControl(new AMap.ToolBar({ position: { top: '12px', right: '12px' } }));
  });

  // 卫星图层 + 路网图层（按需启用）
  const satelliteLayer = new AMap.TileLayer.Satellite();
  const roadnetLayer   = new AMap.TileLayer.RoadNet();
  let satelliteOn = false;
  function setMapMode(mode) {
    state.mapMode = mode;
    if (mode === 'satellite' && !satelliteOn) {
      map.add([satelliteLayer, roadnetLayer]);
      satelliteOn = true;
      map.setMapStyle('amap://styles/normal');
    } else if (mode !== 'satellite' && satelliteOn) {
      map.remove([satelliteLayer, roadnetLayer]);
      satelliteOn = false;
      map.setMapStyle('amap://styles/whitesmoke');
    }
  }

  /* ---------- 基准点 + 圈层 ---------- */

  new AMap.Marker({
    position: [base.lng, base.lat],
    map: map,
    zIndex: 200,
    offset: new AMap.Pixel(-18, -42),
    content: `
      <div class="amap-base-marker" title="${escapeHtml(base.name)}">
        <div class="amap-base-pin">📍</div>
        <div class="amap-base-label">${escapeHtml(base.name)}</div>
      </div>
    `
  });

  const CIRCLES = [
    { r: 1000, color: '#1976d2' },
    { r: 2000, color: '#2e7d32' },
    { r: 3000, color: '#ef6c00' },
    { r: 5000, color: '#c62828' }
  ];
  CIRCLES.forEach((c, i) => {
    new AMap.Circle({
      center: [base.lng, base.lat],
      radius: c.r,
      strokeColor: c.color,
      strokeOpacity: 0.65,
      strokeWeight: 1.5,
      strokeStyle: i === 3 ? 'solid' : 'dashed',
      fillColor: c.color,
      fillOpacity: 0.04,
      map: map,
      zIndex: 10 + i
    });
    new AMap.Text({
      text: c.r === 1000 ? '1 km' : c.r === 2000 ? '2 km' : c.r === 3000 ? '3 km' : '5 km',
      position: offsetPoint(base.lng, base.lat, c.r, 60),
      offset: new AMap.Pixel(-12, -10),
      style: {
        background: '#fff', border: '1px solid ' + c.color, color: c.color,
        padding: '2px 8px', borderRadius: '10px', fontSize: '12px',
        boxShadow: '0 1px 3px rgba(0,0,0,.1)'
      },
      map: map
    });
  });

  function offsetPoint(lng, lat, distMeters, bearingDeg) {
    const R = 6371000;
    const br = bearingDeg * Math.PI / 180;
    const lat1 = lat * Math.PI / 180;
    const lng1 = lng * Math.PI / 180;
    const lat2 = Math.asin(Math.sin(lat1) * Math.cos(distMeters / R) +
                Math.cos(lat1) * Math.sin(distMeters / R) * Math.cos(br));
    const lng2 = lng1 + Math.atan2(Math.sin(br) * Math.sin(distMeters / R) * Math.cos(lat1),
                Math.cos(distMeters / R) - Math.sin(lat1) * Math.sin(lat2));
    return [lng2 * 180 / Math.PI, lat2 * 180 / Math.PI];
  }

  /* ---------- 竞品 marker ---------- */

  const markerMap = new Map();

  function buildMarkerContent(d) {
    const cfg = TIER_CONFIG[d.tier];
    const icon = CAT_ICON[d.category] || '🏢';
    const tCnt = (d.tenants || []).length;
    return `
      <div class="amap-comp-marker ${cfg.cls}" data-id="${d.id}">
        <div class="amap-comp-pin" style="background:${cfg.color}">
          <span class="amap-comp-icon">${icon}</span>
          <span class="amap-comp-num">${d.idx}</span>
        </div>
        ${tCnt ? `<span class="amap-comp-badge" title="下属企业">${tCnt}</span>` : ''}
        <div class="amap-comp-tail" style="border-top-color:${cfg.color}"></div>
      </div>
    `;
  }

  function buildInfoWindow(d) {
    const cfg = TIER_CONFIG[d.tier];
    const ts = d.tenants || [];
    const tenantHtml = ts.length ? `
      <div class="info-tenants">
        <div class="info-tenants-title">代表企业（共 ${ts.length} 家,展示前 10）</div>
        <table class="info-tenants-table">
          <thead><tr><th>企业</th><th>行业</th><th>体量</th><th>租金</th><th>剩余</th><th>📞</th></tr></thead>
          <tbody>
            ${ts.slice(0, 10).map(t => {
              const rem = fmtRemaining(t).html;
              const phone = getPhone(d, t);
              return `
                <tr>
                  <td>${escapeHtml(t.name)}</td>
                  <td>${escapeHtml(t.industry || '—')}</td>
                  <td>${fmtArea(t.area)} ㎡</td>
                  <td>${fmtRent(t.rent)}</td>
                  <td>${rem}</td>
                  <td>${phone ? escapeHtml(phone) : '—'}</td>
                </tr>`;
            }).join('')}
          </tbody>
        </table>
      </div>` : '';

    return `
      <div class="info-card">
        <div class="info-head" style="border-color:${cfg.color}">
          <span class="info-num" style="background:${cfg.color}">${d.idx}</span>
          <div class="info-title">
            <div class="info-name">${escapeHtml(d.name)}</div>
            <div class="info-cat">${CAT_ICON[d.category] || '🏢'} ${escapeHtml(d.category)}
              <span class="info-tier" style="color:${cfg.color}">· ${cfg.label}</span>
            </div>
          </div>
        </div>
        <div class="info-body">
          <div class="info-row"><span class="k">📍 地址</span><span class="v">${escapeHtml(d.address)}</span></div>
          <div class="info-row"><span class="k">📐 距离</span><span class="v">${fmtDist(d.distance)}（${cfg.label}）</span></div>
          <div class="info-row"><span class="k">🚗 驾车</span><span class="v">${escapeHtml(d.drive || '—')}</span></div>
          <div class="info-row"><span class="k">🚇 公交/地铁</span><span class="v">${escapeHtml(d.transit || '—')}</span></div>
          ${d.developer   ? `<div class="info-row"><span class="k">🏗️ 开发商</span><span class="v">${escapeHtml(d.developer)}</span></div>` : ''}
          ${d.area        ? `<div class="info-row"><span class="k">📦 体量</span><span class="v">${escapeHtml(d.area)}</span></div>` : ''}
          ${d.year        ? `<div class="info-row"><span class="k">📅 投用</span><span class="v">${escapeHtml(d.year + ' 年')}</span></div>` : ''}
          ${d.propertyFee ? `<div class="info-row"><span class="k">💰 物业费</span><span class="v"><b>${d.propertyFee}</b> 元/㎡/月</span></div>` : ''}
          ${d.nature      ? `<div class="info-row"><span class="k">🏷️ 性质</span><span class="v">${escapeHtml(d.nature)}</span></div>` : ''}
          ${d.subs && d.subs.length ? `<div class="info-row"><span class="k">🏬 含楼宇</span><span class="v">${d.subs.map(escapeHtml).join('、')}</span></div>` : ''}
          ${d.tags && d.tags.length ? `<div class="info-tags">${d.tags.map(t => `<span class="tag">${escapeHtml(t)}</span>`).join('')}</div>` : ''}
          ${d.note ? `<div class="info-note">📝 ${escapeHtml(d.note)}</div>` : ''}
          ${tenantHtml}
        </div>
        <div class="info-actions">
          <a class="btn-primary"   target="_blank" href="https://uri.amap.com/marker?position=${d.lng},${d.lat}&name=${encodeURIComponent(d.name)}&src=wujiaochang-map&coordinate=gaode">在高德打开</a>
          <a class="btn-secondary" target="_blank" href="https://uri.amap.com/navigation?from=${base.lng},${base.lat},${encodeURIComponent(base.name)}&to=${d.lng},${d.lat},${encodeURIComponent(d.name)}&mode=car&src=wujiaochang-map&coordinate=gaode">路线规划</a>
        </div>
      </div>
    `;
  }

  const infoWindow = new AMap.InfoWindow({
    isCustom: false,
    autoMove: true,
    closeWhenClickMap: true,
    offset: new AMap.Pixel(0, -42)
  });

  function renderMarkers() {
    markerMap.forEach(m => map.remove(m));
    markerMap.clear();

    const list = getFilteredProjects();
    list.forEach((d) => {
      const marker = new AMap.Marker({
        position: [d.lng, d.lat],
        map: map,
        zIndex: 100,
        offset: new AMap.Pixel(-22, -54),
        content: buildMarkerContent(d),
        extData: d
      });
      marker.on('click', () => focusItem(d.id, { source: 'marker' }));
      markerMap.set(d.id, marker);
    });
  }

  /* ====================================================================== */
  /* ============================ 项目概览表 =============================== */
  /* ====================================================================== */

  const projectsTbody = document.getElementById('projectsTbody');

  function renderProjectsTable() {
    const list = getFilteredProjects();
    if (!list.length) {
      projectsTbody.innerHTML = `<tr><td colspan="6" class="empty">没有符合条件的项目,试试调整筛选 / 搜索。</td></tr>`;
      return;
    }
    projectsTbody.innerHTML = list.map((d) => {
      const cfg = TIER_CONFIG[d.tier];
      const subInfo = d.subs && d.subs.length
        ? `<div class="row-subs">含：${d.subs.map(escapeHtml).join('、')}</div>` : '';
      const tags = (d.tags || []).slice(0, 3).map(t => `<span class="row-tag">${escapeHtml(t)}</span>`).join('');
      const tCnt = (d.tenants || []).length;
      return `
        <tr data-id="${d.id}" class="${state.activeId === d.id ? 'is-active' : ''}">
          <td class="row-num"><span class="row-badge" style="background:${cfg.color}">${d.idx}</span></td>
          <td class="row-main">
            <div class="row-name">${escapeHtml(d.name)}
              ${tCnt ? `<span class="row-tenant-cnt" title="下属企业">🏷️ ${tCnt}</span>` : ''}
            </div>
            <div class="row-addr">📍 ${escapeHtml(d.address)}</div>
            ${subInfo}
            <div class="row-tags">${tags}</div>
          </td>
          <td><span class="row-cat">${CAT_ICON[d.category] || ''} ${escapeHtml(d.category)}</span></td>
          <td class="row-dist">
            <div class="dist-num">${fmtDist(d.distance)}</div>
            <div class="dist-tier" style="color:${cfg.color}">${cfg.label}</div>
          </td>
          <td>🚗 ${escapeHtml(d.drive || '—')}</td>
          <td>🚇 ${escapeHtml(d.transit || '—')}</td>
        </tr>
      `;
    }).join('');
  }
  projectsTbody.addEventListener('click', (e) => {
    const tr = e.target.closest('tr[data-id]');
    if (!tr) return;
    focusItem(tr.dataset.id, { source: 'row' });
  });

  /* ====================================================================== */
  /* ============================ 企业明细表 =============================== */
  /* ====================================================================== */

  const tenantsTbody = document.getElementById('tenantsTbody');

  function phoneCellHtml(park, tenant) {
    const phone = getPhone(park, tenant);
    if (phone) {
      return `<span class="phone-val" title="点击修改">📞 ${escapeHtml(phone)}</span>
              <button class="phone-edit" title="编辑">✏️</button>`;
    }
    return `<button class="phone-add" title="点击录入电话">＋ 录入</button>`;
  }

  function renderTenantsTable() {
    const list = getFilteredTenants();
    if (!list.length) {
      tenantsTbody.innerHTML = `<tr><td colspan="12" class="empty">没有符合条件的企业,试试调整筛选 / 搜索。</td></tr>`;
      return;
    }
    tenantsTbody.innerHTML = list.map((r, i) => {
      const p = r.park, t = r.tenant;
      const cfg = TIER_CONFIG[p.tier];
      const remaining = fmtRemaining(t);
      return `
        <tr data-id="${p.id}" data-tenant="${escapeHtml(t.name)}" class="${state.activeId === p.id ? 'is-active' : ''}">
          <td class="row-num"><span class="row-badge tiny" style="background:${cfg.color}">${i + 1}</span></td>
          <td class="t-proj">
            <div class="t-proj-name">${CAT_ICON[p.category] || ''} ${escapeHtml(p.name)}</div>
            <div class="t-tenant-name">└ ${escapeHtml(t.name)}</div>
          </td>
          <td class="t-addr">${escapeHtml(p.address)}</td>
          <td class="t-dist">
            <div class="dist-num">${fmtDist(p.distance)}</div>
            <div class="dist-tier" style="color:${cfg.color}">${cfg.label}</div>
          </td>
          <td class="t-num">${fmtArea(t.area)}</td>
          <td class="t-num"><span class="rent-pill">${fmtRent(t.rent)}</span></td>
          <td class="t-num">${fmtMonths(t.rentFreeMonths)}</td>
          <td class="t-remaining">${remaining.html}</td>
          <td><span class="industry-pill">${escapeHtml(t.industry || '—')}</span></td>
          <td class="t-phone" data-phone-cell="1">${phoneCellHtml(p, t)}</td>
          <td class="t-num">${p.propertyFee != null ? p.propertyFee : '—'}</td>
          <td class="t-nature">${escapeHtml(p.nature || '—')}</td>
        </tr>
      `;
    }).join('');
  }

  tenantsTbody.addEventListener('click', (e) => {
    // 电话单元格的编辑/录入
    const phoneCell = e.target.closest('[data-phone-cell]');
    if (phoneCell && (e.target.classList.contains('phone-add') ||
                     e.target.classList.contains('phone-edit') ||
                     e.target.classList.contains('phone-val'))) {
      e.stopPropagation();
      const tr = phoneCell.closest('tr');
      const parkId = tr.dataset.id;
      const tenantName = tr.dataset.tenant;
      const park = rawData.find(p => p.id === parkId);
      const tenant = park && (park.tenants || []).find(t => t.name === tenantName);
      if (!park || !tenant) return;
      const cur = getPhone(park, tenant);
      const v = window.prompt(`录入【${tenant.name}】的联系电话：\n（留空可清除,会保存到浏览器本地）`, cur);
      if (v === null) return;
      setPhone(park, tenant, v);
      renderTenantsTable();
      return;
    }
    // 普通行点击 → 聚焦地图
    const tr = e.target.closest('tr[data-id]');
    if (!tr) return;
    focusItem(tr.dataset.id, { source: 'tenant-row' });
  });

  /* ====================================================================== */
  /* ============================  联动 & 详情  ============================ */
  /* ====================================================================== */

  function focusItem(id, opts = {}) {
    const d = rawData.find(x => x.id === id);
    if (!d) return;
    state.activeId = id;

    document.querySelectorAll('#projectsTbody tr, #tenantsTbody tr').forEach(tr => {
      tr.classList.toggle('is-active', tr.dataset.id === id);
    });

    if (opts.source !== 'row' && opts.source !== 'tenant-row') {
      const tbodyId = state.tab === 'projects' ? '#projectsTbody' : '#tenantsTbody';
      const activeRow = document.querySelector(`${tbodyId} tr[data-id="${id}"]`);
      if (activeRow) activeRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }

    infoWindow.setContent(buildInfoWindow(d));
    infoWindow.open(map, [d.lng, d.lat]);
    if (opts.source !== 'marker') {
      map.setZoomAndCenter(Math.max(map.getZoom(), 15), [d.lng, d.lat], false, 400);
    }
    showDetail(d);
  }

  const detailPanel = document.getElementById('detailPanel');
  const detailBody  = document.getElementById('detailBody');
  document.getElementById('detailClose').addEventListener('click', () => {
    detailPanel.hidden = true;
  });
  function showDetail(d) {
    detailBody.innerHTML = buildInfoWindow(d);
    detailPanel.hidden = false;
  }

  /* ====================================================================== */
  /* ===============================  Tab  ================================= */
  /* ====================================================================== */

  const paneProjects = document.getElementById('paneProjects');
  const paneTenants  = document.getElementById('paneTenants');
  const btnSync = document.getElementById('btnSync');

  document.querySelectorAll('.tab[data-tab]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const t = btn.dataset.tab;
      state.tab = t;
      document.querySelectorAll('.tab').forEach(b => b.classList.toggle('active', b.dataset.tab === t));
      paneProjects.hidden = t !== 'projects';
      paneTenants.hidden  = t !== 'tenants';
      btnSync.hidden = t !== 'tenants';
      refresh();
    });
  });

  /* ====================================================================== */
  /* ===========================  顶部 / 看板  ============================= */
  /* ====================================================================== */

  function renderDashboard() {
    const list = rawData.filter(d => d.tier !== 'out');
    const tenantCnt = list.reduce((a, d) => a + ((d.tenants || []).length), 0);

    document.getElementById('cnt-all').textContent = list.length;
    document.getElementById('cnt-1km').textContent = list.filter(d => d.tier === '1km').length;
    document.getElementById('cnt-12').textContent  = list.filter(d => d.tier === '1-2km').length;
    document.getElementById('cnt-23').textContent  = list.filter(d => d.tier === '2-3km').length;
    document.getElementById('cnt-35').textContent  = list.filter(d => d.tier === '3-5km').length;
    document.getElementById('basePointName').textContent = base.name;
    document.getElementById('topSummary').textContent = `${list.length} 个项目 / ${tenantCnt} 家企业`;

    const now = new Date();
    document.getElementById('updateTime').textContent =
      `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
  }

  function updateCounts() {
    const proj = getFilteredProjects();
    const ten  = getFilteredTenants();
    document.getElementById('tabCntProjects').textContent = proj.length;
    document.getElementById('tabCntTenants').textContent  = ten.length;
    document.getElementById('listCount').textContent =
      state.tab === 'projects'
        ? `共 ${proj.length} 个项目`
        : `共 ${ten.length} 家企业 · 涉 ${new Set(ten.map(r => r.park.id)).size} 个项目`;
  }

  /* ====================================================================== */
  /* =============================  筛选 / 搜索  =========================== */
  /* ====================================================================== */

  document.querySelectorAll('[data-filter-cat]').forEach((btn) => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('[data-filter-cat]').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.filterCat = btn.dataset.filterCat;
      refresh();
    });
  });
  document.querySelectorAll('[data-filter-tier]').forEach((btn) => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('[data-filter-tier]').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.filterTier = btn.dataset.filterTier;
      refresh();
    });
  });
  document.querySelectorAll('.stat[data-tier]').forEach((card) => {
    card.addEventListener('click', () => {
      const t = card.dataset.tier;
      document.querySelectorAll('[data-filter-tier]').forEach(b => {
        b.classList.toggle('active', b.dataset.filterTier === t);
      });
      state.filterTier = t;
      refresh();
    });
  });
  document.querySelectorAll('[data-map-mode]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const m = btn.dataset.mapMode;
      document.querySelectorAll('[data-map-mode]').forEach(b => b.classList.toggle('active', b.dataset.mapMode === m));
      setMapMode(m);
    });
  });

  let searchTimer = null;
  document.getElementById('searchInput').addEventListener('input', (e) => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      state.search = e.target.value || '';
      refresh();
    }, 180);
  });
  document.getElementById('btnReset').addEventListener('click', () => {
    state.filterCat = 'all';
    state.filterTier = 'all';
    state.search = '';
    document.getElementById('searchInput').value = '';
    document.querySelectorAll('[data-filter-cat]').forEach(b => b.classList.toggle('active', b.dataset.filterCat === 'all'));
    document.querySelectorAll('[data-filter-tier]').forEach(b => b.classList.toggle('active', b.dataset.filterTier === 'all'));
    refresh(true);
  });

  /* ====================================================================== */
  /* =============================  CSV 导出  ============================== */
  /* ====================================================================== */

  function csvOf(rows) {
    return rows.map(r =>
      r.map(c => {
        const s = String(c == null ? '' : c).replace(/"/g, '""');
        return /[",\n]/.test(s) ? `"${s}"` : s;
      }).join(',')
    ).join('\n');
  }
  function download(csv, filename) {
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  document.getElementById('btnExport').addEventListener('click', () => {
    const day = new Date().toISOString().slice(0, 10);
    if (state.tab === 'projects') {
      const list = getFilteredProjects();
      const headers = ['序号','项目名称','业态','地址','距离(米)','距离圈层',
                       '驾车','公交/地铁','开发商','体量','投用年份',
                       '物业费(元/㎡/月)','性质说明','含楼宇','标签','下属企业数','备注'];
      const rows = list.map(d => [
        d.idx, d.name, d.category, d.address, Math.round(d.distance), TIER_CONFIG[d.tier].label,
        d.drive || '', d.transit || '', d.developer || '', d.area || '', d.year || '',
        d.propertyFee != null ? d.propertyFee : '', d.nature || '',
        (d.subs || []).join('、'), (d.tags || []).join('、'),
        (d.tenants || []).length, d.note || ''
      ]);
      download(csvOf([headers, ...rows]),
               `杨浦五角场5km竞品-项目概览_${day}.csv`);
    } else {
      const list = getFilteredTenants();
      const headers = ['序号','项目名称','企业名称','地址','距离(米)','距离圈层',
                       '办公体量(㎡)','成交租金(元/㎡/天)','免租期(月)',
                       '租约开始','租约期(年)','到期日期','剩余(月)',
                       '主力客户行业','电话','物业费(元/㎡/月)','性质说明','成交年份','备注'];
      const rows = list.map((r, i) => {
        const p = r.park, t = r.tenant;
        const lease = computeLease(t);
        return [
          i + 1, p.name, t.name, p.address, Math.round(p.distance), TIER_CONFIG[p.tier].label,
          t.area || '', t.rent != null ? t.rent : '', t.rentFreeMonths != null ? t.rentFreeMonths : '',
          lease.startStr || '', t.leaseTerm || '', lease.endStr || '',
          lease.status === 'expired' ? 0 : (lease.remainingMonths != null ? lease.remainingMonths : ''),
          t.industry || '', getPhone(p, t) || '',
          p.propertyFee != null ? p.propertyFee : '', p.nature || '',
          t.dealYear || '', t.note || ''
        ];
      });
      download(csvOf([headers, ...rows]),
               `杨浦五角场5km竞品-企业明细_${day}.csv`);
    }
  });

  /* ====================================================================== */
  /* ============================  API 配置弹窗  =========================== */
  /* ====================================================================== */

  const apiModal = document.getElementById('apiModal');
  const apiProxyUrlInput = document.getElementById('apiProxyUrl');
  const apiStatus = document.getElementById('apiStatus');

  function openApiModal() {
    apiProxyUrlInput.value = getApiProxy();
    apiStatus.innerHTML = '';
    apiModal.hidden = false;
    setTimeout(() => apiProxyUrlInput.focus(), 50);
  }
  function closeApiModal() { apiModal.hidden = true; }

  document.getElementById('btnApiConfig').addEventListener('click', openApiModal);
  apiModal.addEventListener('click', (e) => {
    if (e.target.dataset.modalClose) closeApiModal();
  });

  document.getElementById('apiSave').addEventListener('click', async () => {
    const url = apiProxyUrlInput.value.trim();
    setApiProxy(url);
    apiStatus.innerHTML = '<span class="api-status-info">已保存,正在测试连接…</span>';
    if (!url) {
      apiStatus.innerHTML = '<span class="api-status-warn">已清空代理地址。如需补充电话,请填写代理地址或在表格中手动录入。</span>';
      return;
    }
    try {
      const res = await fetch(url.replace(/\/+$/, '') + '/health', { method: 'GET' });
      const data = await res.json();
      if (data.ok) {
        const tyc = data.providers && data.providers.tianyancha;
        const qcc = data.providers && data.providers.qichacha;
        apiStatus.innerHTML = `<span class="api-status-ok">✅ 代理可达 · 天眼查：${escapeHtml(tyc||'')} · 企查查：${escapeHtml(qcc||'')}</span>`;
      } else {
        apiStatus.innerHTML = '<span class="api-status-warn">⚠️ 代理已响应但返回异常,请检查 API Key 是否配置。</span>';
      }
    } catch (e) {
      apiStatus.innerHTML = `<span class="api-status-err">❌ 无法访问代理：${escapeHtml(e.message)}<br/>请先启动 <code>node server/qcc-proxy.js</code>。</span>`;
    }
  });

  /* ====================================================================== */
  /* =========================  批量电话/工商同步  ========================= */
  /* ====================================================================== */

  async function enrichBatch() {
    const proxy = getApiProxy();
    if (!proxy) {
      openApiModal();
      apiStatus.innerHTML = '<span class="api-status-warn">⚠️ 请先填写代理地址（详情见 server/README.md）。也可以在表格电话单元格中手动录入。</span>';
      return;
    }
    const list = getFilteredTenants().filter(r => !getPhone(r.park, r.tenant));
    if (!list.length) {
      alert('当前筛选的所有企业都已有电话信息了 🎉');
      return;
    }
    if (!confirm(`将通过代理 ${proxy} 同步 ${list.length} 家企业的工商信息（电话/法人/注册资本等）,逐条调用约 ${(list.length * 0.8).toFixed(0)} 秒。继续？`)) return;

    btnSync.disabled = true;
    let ok = 0, fail = 0;
    for (let i = 0; i < list.length; i++) {
      const r = list[i];
      btnSync.textContent = `🔄 同步中 ${i + 1}/${list.length}`;
      try {
        const url = proxy.replace(/\/+$/, '') + '/api/enrich?company=' + encodeURIComponent(r.tenant.name);
        const res = await fetch(url);
        const data = await res.json();
        if (data && data.phone) {
          setPhone(r.park, r.tenant, data.phone, {
            legalPerson: data.legalPerson || r.tenant.legalPerson,
            regCapital: data.regCapital || r.tenant.regCapital,
            regNumber:  data.regNumber  || r.tenant.regNumber,
            regTime:    data.regTime    || r.tenant.regTime
          });
          ok++;
        } else {
          fail++;
        }
      } catch (e) {
        fail++;
      }
      // 节流：800ms 一条,避免触发对方限流
      await new Promise(r => setTimeout(r, 800));
    }
    btnSync.disabled = false;
    btnSync.textContent = '🔄 同步电话';
    renderTenantsTable();
    alert(`同步完成 · 成功 ${ok} / 失败 ${fail}`);
  }
  btnSync.addEventListener('click', enrichBatch);

  /* ====================================================================== */
  /* ===============================  刷新  ================================ */
  /* ====================================================================== */

  function refresh(autoFit) {
    renderMarkers();
    renderProjectsTable();
    renderTenantsTable();
    updateCounts();
    if (autoFit) {
      map.setZoomAndCenter(13.5, [base.lng, base.lat], false, 400);
    }
  }

  /* ====================================================================== */
  /* ===============================  启动  ================================ */
  /* ====================================================================== */

  renderDashboard();
  refresh();
})();
