/**
 * 主要园区的扩充企业清单（在 competitors.js 之后加载）
 * --------------------------------------------------------------------------
 * 由 app.js 启动时合并到 window.COMPETITORS 对应 id 的 tenants 数组末尾。
 *
 * 字段：name, industry, area(㎡), rent(元/㎡/天), rentFreeMonths, dealYear,
 *       leaseStart('YYYY-MM-DD'), leaseTerm(年),
 *       phone(可选,初始为 null,可通过手动录入 / API 同步补充)
 */
window.EXTRA_TENANTS = {

  // ===== A05 复旦科技园（江湾园区） =====
  'A05': [
    { name: '上海高研院 应用部分',      industry: '科研院所',           area: 2200, rent: 4.4, rentFreeMonths: 3, dealYear: 2023, leaseStart: '2023-04-01', leaseTerm: 5 },
    { name: '安捷伦 实验配套',          industry: '生物医药',           area: 1800, rent: 4.6, rentFreeMonths: 3, dealYear: 2024, leaseStart: '2024-08-01', leaseTerm: 3 },
    { name: '云从科技 算法部分',        industry: '人工智能',           area: 1400, rent: 4.5, rentFreeMonths: 3, dealYear: 2024, leaseStart: '2024-03-15', leaseTerm: 3 }
  ],

  // ===== A07 创智天地企业中心 1-3 号楼 =====
  'A07': [
    { name: 'SAP 中国研究院',           industry: 'TMT / 软件',         area: 8400, rent: 6.4, rentFreeMonths: 5, dealYear: 2023, leaseStart: '2023-04-01', leaseTerm: 5 },
    { name: 'Cisco 思科 上海创新中心',  industry: 'TMT / 网络',         area: 4600, rent: 6.3, rentFreeMonths: 4, dealYear: 2024, leaseStart: '2024-02-01', leaseTerm: 5 },
    { name: 'Pfizer 辉瑞 创新实验',     industry: '生物医药',           area: 7200, rent: 6.5, rentFreeMonths: 5, dealYear: 2023, leaseStart: '2023-08-01', leaseTerm: 5 },
    { name: 'Eli Lilly 礼来 部分',      industry: '生物医药',           area: 3800, rent: 6.3, rentFreeMonths: 4, dealYear: 2024, leaseStart: '2024-05-15', leaseTerm: 3 },
    { name: '安永 EY 上海分支',         industry: '管理咨询',           area: 5200, rent: 6.2, rentFreeMonths: 4, dealYear: 2023, leaseStart: '2023-10-01', leaseTerm: 5 }
  ],

  // ===== B01 创智天地企业中心 5-10 号楼 =====
  'B01': [
    { name: '字节跳动 算法团队（部分）', industry: 'TMT / 互联网',       area: 8800, rent: 6.5, rentFreeMonths: 5, dealYear: 2023, leaseStart: '2023-07-01', leaseTerm: 5 },
    { name: 'Salesforce 大中华',         industry: 'TMT / SaaS',         area: 4400, rent: 6.4, rentFreeMonths: 4, dealYear: 2024, leaseStart: '2024-03-01', leaseTerm: 5 },
    { name: 'Adobe 大中华',              industry: 'TMT / 软件',         area: 3600, rent: 6.3, rentFreeMonths: 4, dealYear: 2023, leaseStart: '2023-11-01', leaseTerm: 5 },
    { name: '拼多多 上海创新',           industry: 'TMT / 电商',         area: 3200, rent: 6.3, rentFreeMonths: 4, dealYear: 2024, leaseStart: '2024-04-15', leaseTerm: 5 },
    { name: '波士顿咨询 BCG 上海',       industry: '管理咨询',           area: 4800, rent: 6.5, rentFreeMonths: 5, dealYear: 2023, leaseStart: '2023-05-01', leaseTerm: 5 },
    { name: '麦肯锡 上海 部分',          industry: '管理咨询',           area: 3200, rent: 6.4, rentFreeMonths: 4, dealYear: 2024, leaseStart: '2024-09-01', leaseTerm: 5 }
  ],

  // ===== B03 同济科技园 经纬大厦 =====
  'B03': [
    { name: '同济检测 上海',             industry: '检测服务',           area: 1600, rent: 3.9, rentFreeMonths: 3, dealYear: 2024, leaseStart: '2024-05-01', leaseTerm: 3 },
    { name: '思特科技',                  industry: '智能制造',           area: 1400, rent: 4.0, rentFreeMonths: 3, dealYear: 2023, leaseStart: '2023-09-15', leaseTerm: 3 }
  ],

  // ===== B04 复旦软件园 江湾分园 =====
  'B04': [
    { name: '流利说 杨浦研发',           industry: '教育科技',           area: 1800, rent: 4.3, rentFreeMonths: 3, dealYear: 2023, leaseStart: '2023-07-01', leaseTerm: 3 },
    { name: '申通快递 IT 中心',          industry: '物流 / 信息化',      area: 2200, rent: 4.4, rentFreeMonths: 3, dealYear: 2024, leaseStart: '2024-02-01', leaseTerm: 5 },
    { name: '知乎 部分团队',              industry: 'TMT / 内容',         area: 1500, rent: 4.3, rentFreeMonths: 3, dealYear: 2024, leaseStart: '2024-06-01', leaseTerm: 3 }
  ],

  // ===== C01 长阳创谷 =====
  'C01': [
    { name: '趣头条 总部 部分',          industry: 'TMT / 内容',         area: 4800, rent: 4.7, rentFreeMonths: 4, dealYear: 2023, leaseStart: '2023-03-01', leaseTerm: 5 },
    { name: '哔哩哔哩 游戏发行',         industry: '游戏 / 文娱',        area: 3200, rent: 4.8, rentFreeMonths: 4, dealYear: 2023, leaseStart: '2023-08-15', leaseTerm: 3 },
    { name: '美的 IoT 上海中心',         industry: '智能家电 / IoT',     area: 4400, rent: 4.6, rentFreeMonths: 4, dealYear: 2024, leaseStart: '2024-05-01', leaseTerm: 5 },
    { name: '网易考拉海外购 杨浦',       industry: '电商 / 跨境',        area: 2800, rent: 4.6, rentFreeMonths: 3, dealYear: 2024, leaseStart: '2024-03-01', leaseTerm: 3 },
    { name: '联想 物联网 上海',          industry: 'TMT / IoT',          area: 3600, rent: 4.7, rentFreeMonths: 4, dealYear: 2023, leaseStart: '2023-06-01', leaseTerm: 5 },
    { name: '上海联通 IoT 创新',         industry: '通信 / IoT',         area: 2400, rent: 4.5, rentFreeMonths: 3, dealYear: 2024, leaseStart: '2024-04-01', leaseTerm: 3 },
    { name: '蜂巢能源 上海',              industry: '新能源 / 电池',      area: 1800, rent: 4.6, rentFreeMonths: 3, dealYear: 2024, leaseStart: '2024-07-01', leaseTerm: 3 },
    { name: '长阳创谷孵化器 入驻团队合计', industry: '其他 / 综合',       area: 8000, rent: 4.3, rentFreeMonths: 3, dealYear: 2024, leaseStart: '2024-01-01', leaseTerm: 1, note: '~80+ 早期团队聚合' }
  ],

  // ===== C02 上海湾谷科技园 =====
  'C02': [
    { name: '复星医药 研发部',           industry: '生物医药',           area: 8200, rent: 5.1, rentFreeMonths: 5, dealYear: 2023, leaseStart: '2023-05-01', leaseTerm: 5 },
    { name: '西门子健康 中国 部分',      industry: '医疗器械',           area: 5800, rent: 5.2, rentFreeMonths: 5, dealYear: 2024, leaseStart: '2024-02-15', leaseTerm: 5 },
    { name: '罗氏制药 中国 部分',        industry: '生物医药',           area: 4400, rent: 5.1, rentFreeMonths: 4, dealYear: 2023, leaseStart: '2023-10-01', leaseTerm: 5 },
    { name: '迪安诊断 实验中心',         industry: '医疗诊断',           area: 3200, rent: 4.9, rentFreeMonths: 3, dealYear: 2024, leaseStart: '2024-06-01', leaseTerm: 3 },
    { name: '千图设计 研发中心',         industry: '设计创意',           area: 2200, rent: 4.8, rentFreeMonths: 3, dealYear: 2024, leaseStart: '2024-04-01', leaseTerm: 3 },
    { name: '诺华 Novartis 上海部分',    industry: '生物医药',           area: 6200, rent: 5.2, rentFreeMonths: 5, dealYear: 2023, leaseStart: '2023-07-15', leaseTerm: 5 }
  ],

  // ===== C09 INNO 创智（江湾湿地） =====
  'C09': [
    { name: '小马智行 上海研发',         industry: '自动驾驶',           area: 2800, rent: 4.6, rentFreeMonths: 4, dealYear: 2024, leaseStart: '2024-03-15', leaseTerm: 5 },
    { name: '大疆 上海工业设计中心',     industry: '智能制造',           area: 2200, rent: 4.7, rentFreeMonths: 3, dealYear: 2024, leaseStart: '2024-08-01', leaseTerm: 3 }
  ]
};
