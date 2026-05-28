/* ==========================================================================
 * 杨浦五角场 5km 商办竞品地图 - 业务脚本
 * --------------------------------------------------------------------------
 *  - 初始化高德地图
 *  - 绘制基准点、4 个距离圈层
 *  - 计算每个竞品的距离并归属圈层
 *  - 渲染统计看板、数据表
 *  - 实现地图 ↔ 表格联动
 *  - 实现业态/距离筛选、搜索、CSV 导出
 * ========================================================================== */

(function () {
  'use strict';

  /* ---------- 工具函数 ---------- */

  const TIER_CONFIG = {
    '1km':    { label: '1 km 内', min: 0,    max: 1000, color: '#1976d2', cls: 'tier-1' },
    '1-2km':  { label: '1~2 km',  min: 1000, max: 2000, color: '#2e7d32', cls: 'tier-2' },
    '2-3km':  { label: '2~3 km',  min: 2000, max: 3000, color: '#ef6c00', cls: 'tier-3' },
    '3-5km':  { label: '3~5 km',  min: 3000, max: 5000, color: '#c62828', cls: 'tier-4' },
    'out':    { label: '>5 km',   min: 5000, max: 1e9,  color: '#9e9e9e', cls: 'tier-out' }
  };

  const CAT_ICON = {
    '商办写字楼': '🏢',
    '产业园区':   '🏭',
    '众创空间':   '💡',
    '商业综合体': '🛍️'
  };

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

  function escapeHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  /* ---------- 数据加工 ---------- */

  const base = window.BASE_POINT;
  const rawData = (window.COMPETITORS || []).filter(d => !d.excluded);

  // 计算距离 + 归属圈层
  rawData.forEach((d) => {
    d.distance = haversine(base.lng, base.lat, d.lng, d.lat);
    d.tier = tierOf(d.distance);
  });

  // 按距离升序排序,重新编号
  rawData.sort((a, b) => a.distance - b.distance);
  rawData.forEach((d, i) => { d.idx = i + 1; });

  /* ---------- 状态管理 ---------- */

  const state = {
    filterCat: 'all',
    filterTier: 'all',
    search: '',
    activeId: null
  };

  function getFilteredData() {
    const kw = state.search.trim().toLowerCase();
    return rawData.filter((d) => {
      if (d.tier === 'out') return false;
      if (state.filterCat !== 'all' && d.category !== state.filterCat) return false;
      if (state.filterTier !== 'all' && d.tier !== state.filterTier) return false;
      if (kw) {
        const hay = (d.name + ' ' + d.address + ' ' + (d.tags || []).join(' ') +
                     ' ' + (d.subs || []).join(' ') + ' ' + (d.developer || '')).toLowerCase();
        if (hay.indexOf(kw) === -1) return false;
      }
      return true;
    });
  }

  /* ---------- 地图初始化 ---------- */

  const map = new AMap.Map('map', {
    zoom: 13.5,
    center: [base.lng, base.lat],
    viewMode: '2D',
    mapStyle: 'amap://styles/whitesmoke',
    features: ['bg', 'road', 'building', 'point']
  });

  map.addControl && AMap.plugin(['AMap.Scale', 'AMap.ToolBar'], () => {
    map.addControl(new AMap.Scale());
    map.addControl(new AMap.ToolBar({ position: { top: '12px', right: '12px' } }));
  });

  /* ---------- 基准点 + 圈层 ---------- */

  const baseMarker = new AMap.Marker({
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

  // 4 个圈层 (1km / 2km / 3km / 5km)
  const CIRCLES = [
    { r: 1000, color: '#1976d2', fill: '#1976d2' },
    { r: 2000, color: '#2e7d32', fill: '#2e7d32' },
    { r: 3000, color: '#ef6c00', fill: '#ef6c00' },
    { r: 5000, color: '#c62828', fill: '#c62828' }
  ];
  CIRCLES.forEach((c, i) => {
    new AMap.Circle({
      center: [base.lng, base.lat],
      radius: c.r,
      strokeColor: c.color,
      strokeOpacity: 0.55,
      strokeWeight: 1.5,
      strokeStyle: i === 3 ? 'solid' : 'dashed',
      fillColor: c.fill,
      fillOpacity: 0.04,
      map: map,
      zIndex: 10 + i
    });
    // 圈层标签
    new AMap.Text({
      text: c.r === 1000 ? '1 km' : c.r === 2000 ? '2 km' : c.r === 3000 ? '3 km' : '5 km',
      position: offsetPoint(base.lng, base.lat, c.r, 60),
      offset: new AMap.Pixel(-12, -10),
      style: {
        background: '#fff',
        border: '1px solid ' + c.color,
        color: c.color,
        padding: '2px 8px',
        borderRadius: '10px',
        fontSize: '12px',
        boxShadow: '0 1px 3px rgba(0,0,0,.1)'
      },
      map: map
    });
  });

  // 给定基点 + 半径(米) + 方向角(度),计算地球表面新点(用于标签摆放)
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

  const markerMap = new Map(); // id -> AMap.Marker

  function buildMarkerContent(d, isActive) {
    const cfg = TIER_CONFIG[d.tier];
    const icon = CAT_ICON[d.category] || '🏢';
    return `
      <div class="amap-comp-marker ${cfg.cls} ${isActive ? 'is-active' : ''}"
           data-id="${d.id}">
        <div class="amap-comp-pin" style="background:${cfg.color}">
          <span class="amap-comp-icon">${icon}</span>
          <span class="amap-comp-num">${d.idx}</span>
        </div>
        <div class="amap-comp-tail" style="border-top-color:${cfg.color}"></div>
      </div>
    `;
  }

  function buildInfoWindow(d) {
    const cfg = TIER_CONFIG[d.tier];
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
          ${d.developer ? `<div class="info-row"><span class="k">🏗️ 开发商</span><span class="v">${escapeHtml(d.developer)}</span></div>` : ''}
          ${d.area      ? `<div class="info-row"><span class="k">📦 体量</span><span class="v">${escapeHtml(d.area)}</span></div>` : ''}
          ${d.year      ? `<div class="info-row"><span class="k">📅 投用</span><span class="v">${escapeHtml(d.year + ' 年')}</span></div>` : ''}
          ${d.subs && d.subs.length ? `<div class="info-row"><span class="k">🏬 含楼宇</span><span class="v">${d.subs.map(escapeHtml).join('、')}</span></div>` : ''}
          ${d.tags && d.tags.length ? `<div class="info-tags">${d.tags.map(t => `<span class="tag">${escapeHtml(t)}</span>`).join('')}</div>` : ''}
          ${d.note ? `<div class="info-note">📝 ${escapeHtml(d.note)}</div>` : ''}
        </div>
        <div class="info-actions">
          <a class="btn-primary" href="https://uri.amap.com/marker?position=${d.lng},${d.lat}&name=${encodeURIComponent(d.name)}&src=wujiaochang-map&coordinate=gaode" target="_blank">在高德打开</a>
          <a class="btn-secondary" href="https://uri.amap.com/navigation?from=${base.lng},${base.lat},${encodeURIComponent(base.name)}&to=${d.lng},${d.lat},${encodeURIComponent(d.name)}&mode=car&src=wujiaochang-map&coordinate=gaode" target="_blank">路线规划</a>
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
    // 清空
    markerMap.forEach(m => map.remove(m));
    markerMap.clear();

    const list = getFilteredData();
    list.forEach((d) => {
      const marker = new AMap.Marker({
        position: [d.lng, d.lat],
        map: map,
        zIndex: 100,
        offset: new AMap.Pixel(-22, -54),
        content: buildMarkerContent(d, false),
        extData: d
      });
      marker.on('click', () => {
        focusItem(d.id, { source: 'marker' });
      });
      markerMap.set(d.id, marker);
    });
  }

  /* ---------- 表格渲染 ---------- */

  const tbody = document.getElementById('compTbody');
  const listCount = document.getElementById('listCount');

  function renderTable() {
    const list = getFilteredData();
    listCount.textContent = `共 ${list.length} 个项目`;

    if (!list.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty">没有符合条件的项目,试试调整筛选 / 搜索。</td></tr>`;
      return;
    }

    tbody.innerHTML = list.map((d) => {
      const cfg = TIER_CONFIG[d.tier];
      const subInfo = d.subs && d.subs.length
        ? `<div class="row-subs">含：${d.subs.map(escapeHtml).join('、')}</div>`
        : '';
      const tags = (d.tags || []).slice(0, 3).map(t => `<span class="row-tag">${escapeHtml(t)}</span>`).join('');
      return `
        <tr data-id="${d.id}" class="${state.activeId === d.id ? 'is-active' : ''}">
          <td class="row-num"><span class="row-badge" style="background:${cfg.color}">${d.idx}</span></td>
          <td class="row-main">
            <div class="row-name">${escapeHtml(d.name)}</div>
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

  tbody.addEventListener('click', (e) => {
    const tr = e.target.closest('tr[data-id]');
    if (!tr) return;
    focusItem(tr.dataset.id, { source: 'row' });
  });

  /* ---------- 联动：点击行或 marker 都聚焦同一项 ---------- */

  function focusItem(id, opts = {}) {
    const d = rawData.find(x => x.id === id);
    if (!d) return;
    state.activeId = id;

    // 行高亮
    document.querySelectorAll('#compTbody tr').forEach(tr => {
      tr.classList.toggle('is-active', tr.dataset.id === id);
    });
    // 滚到可见
    const activeRow = document.querySelector(`#compTbody tr[data-id="${id}"]`);
    if (activeRow && opts.source !== 'row') {
      activeRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }

    // 地图：弹窗 + 居中
    infoWindow.setContent(buildInfoWindow(d));
    infoWindow.open(map, [d.lng, d.lat]);
    if (opts.source !== 'marker') {
      map.setZoomAndCenter(Math.max(map.getZoom(), 15), [d.lng, d.lat], false, 400);
    }

    // 详情侧栏
    showDetail(d);
  }

  /* ---------- 详情侧栏 ---------- */

  const detailPanel = document.getElementById('detailPanel');
  const detailBody  = document.getElementById('detailBody');
  document.getElementById('detailClose').addEventListener('click', () => {
    detailPanel.hidden = true;
  });
  function showDetail(d) {
    detailBody.innerHTML = buildInfoWindow(d);
    detailPanel.hidden = false;
  }

  /* ---------- 统计看板 ---------- */

  function renderDashboard() {
    const list = rawData.filter(d => d.tier !== 'out');
    document.getElementById('cnt-all').textContent = list.length;
    document.getElementById('cnt-1km').textContent = list.filter(d => d.tier === '1km').length;
    document.getElementById('cnt-12').textContent  = list.filter(d => d.tier === '1-2km').length;
    document.getElementById('cnt-23').textContent  = list.filter(d => d.tier === '2-3km').length;
    document.getElementById('cnt-35').textContent  = list.filter(d => d.tier === '3-5km').length;
    document.getElementById('basePointName').textContent = base.name;

    const now = new Date();
    document.getElementById('updateTime').textContent =
      `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
  }

  /* ---------- 筛选 / 搜索 ---------- */

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

  // 顶部统计卡也可作为距离筛选入口
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

  /* ---------- 导出 CSV ---------- */

  document.getElementById('btnExport').addEventListener('click', () => {
    const list = getFilteredData();
    const headers = ['序号', '项目名称', '业态', '地址', '距离(米)', '距离圈层',
                     '驾车', '公交/地铁', '开发商', '体量', '投用年份', '含楼宇', '标签', '备注'];
    const rows = list.map(d => [
      d.idx, d.name, d.category, d.address, Math.round(d.distance), TIER_CONFIG[d.tier].label,
      d.drive || '', d.transit || '', d.developer || '', d.area || '', d.year || '',
      (d.subs || []).join('、'), (d.tags || []).join('、'), d.note || ''
    ]);
    const csv = [headers, ...rows].map(r =>
      r.map(c => {
        const s = String(c == null ? '' : c).replace(/"/g, '""');
        return /[",\n]/.test(s) ? `"${s}"` : s;
      }).join(',')
    ).join('\n');
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `杨浦五角场5km竞品_${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

  /* ---------- 刷新流程 ---------- */

  function refresh(autoFit) {
    renderMarkers();
    renderTable();
    if (autoFit) {
      map.setZoomAndCenter(13.5, [base.lng, base.lat], false, 400);
    }
  }

  /* ---------- 启动 ---------- */

  renderDashboard();
  refresh();
})();
