/**
 * 杨浦五角场 5km 竞品调研数据
 * --------------------------------------------------------------------------
 * 字段说明：
 *  - 项目级：
 *      id / name / category / address / lng / lat / drive / transit
 *      developer / area / year / subs / tags / note
 *      propertyFee  物业费（元/㎡/月）
 *      nature       性质说明（产权 + 物业等级 / 园区性质）
 *  - 距离与圈层在 app.js 中通过 haversine 自动计算,不需手填。
 *
 *  - tenants[]：园区/写字楼下属企业（代表性 anchor + 典型租户）
 *      name             企业名称
 *      industry         主力客户行业
 *      area             办公体量（㎡）
 *      rent             成交租金（元/㎡/天）
 *      rentFreeMonths   免租期（月）
 *      dealYear         成交/在租年份
 *      note             备注（可选）
 */

window.BASE_POINT = {
  name: '上海市五角场创新创业学院',
  address: '上海市杨浦区国权路18号',
  lng: 121.5103,
  lat: 31.3008
};

window.COMPETITORS = [
  /* ========================================================================
   * ===============================  1 km 内  ==============================
   * ====================================================================== */
  {
    id: 'A01',
    name: '合生国际广场',
    category: '商办写字楼',
    address: '上海市杨浦区淞沪路18号',
    lng: 121.5076, lat: 31.2984,
    drive: '4 分钟 / 1.1 km',
    transit: '步行 11 分钟 · 10号线五角场站',
    developer: '合生创展',
    area: '约 18 万㎡',
    year: 2008,
    subs: ['T1 商务塔', 'T2 商务塔', 'L1-L4 商业裙房'],
    tags: ['核心商圈', '甲级写字楼', '地铁口'],
    note: '五角场核心商圈代表写字楼,甲级标准,层高 3.8m。',
    propertyFee: 28,
    nature: '民营 · 甲级写字楼',
    tenants: [
      { name: '招商银行 杨浦支行',         industry: '金融 / 银行',        area: 3200, rent: 5.8, rentFreeMonths: 3, dealYear: 2023 },
      { name: '平安人寿 上海北分',         industry: '金融 / 保险',        area: 2400, rent: 5.5, rentFreeMonths: 3, dealYear: 2024 },
      { name: '德勤咨询 杨浦项目组',       industry: '管理咨询',           area: 1800, rent: 6.0, rentFreeMonths: 4, dealYear: 2023 },
      { name: '链家·贝壳 区域中心',        industry: '房地产服务',         area: 2600, rent: 5.4, rentFreeMonths: 2, dealYear: 2024 }
    ]
  },
  {
    id: 'A02',
    name: '五角场万达广场（写字楼）',
    category: '商办写字楼',
    address: '上海市杨浦区淞沪路77号',
    lng: 121.5128, lat: 31.3025,
    drive: '4 分钟 / 1.2 km',
    transit: '步行 8 分钟 · 10号线江湾体育场站',
    developer: '万达集团',
    area: '约 14 万㎡',
    year: 2007,
    subs: ['办公塔楼', '万达商业广场裙楼'],
    tags: ['商业综合体', '核心商圈'],
    note: '万达集团旗舰商业综合体,办公+商业一体化。',
    propertyFee: 25,
    nature: '民营 · 商业综合体办公',
    tenants: [
      { name: '万达电影 总部办公',         industry: '文娱传媒',           area: 3500, rent: 5.5, rentFreeMonths: 3, dealYear: 2023 },
      { name: '喜茶 上海总部',             industry: '消费品牌 / 餐饮',    area: 2200, rent: 5.6, rentFreeMonths: 3, dealYear: 2024 },
      { name: '海底捞 华东运营中心',       industry: '消费品牌 / 餐饮',    area: 1900, rent: 5.4, rentFreeMonths: 3, dealYear: 2024 },
      { name: '美的集团 上海分公司',       industry: '智能家电',           area: 1600, rent: 5.3, rentFreeMonths: 2, dealYear: 2022 }
    ]
  },
  {
    id: 'A03',
    name: '百联又一城（办公部分）',
    category: '商办写字楼',
    address: '上海市杨浦区淞沪路8号',
    lng: 121.5080, lat: 31.3007,
    drive: '3 分钟 / 0.7 km',
    transit: '步行 6 分钟 · 10号线五角场站',
    developer: '百联集团',
    area: '约 10 万㎡（办公）',
    year: 2006,
    subs: ['办公塔', '百联购物中心'],
    tags: ['核心商圈', '商业综合体'],
    note: '体量大、人流稳定,办公以中小企业为主。',
    propertyFee: 22,
    nature: '国资 · 商业综合体办公',
    tenants: [
      { name: '百联集团 管理总部',         industry: '商业零售',           area: 6000, rent: 5.0, rentFreeMonths: 3, dealYear: 2022 },
      { name: '携程旅行 杨浦中心',         industry: '互联网 / 出行',      area: 2800, rent: 5.2, rentFreeMonths: 3, dealYear: 2024 },
      { name: '上海家化 销售中心',         industry: '日化美妆',           area: 1500, rent: 4.9, rentFreeMonths: 2, dealYear: 2023 }
    ]
  },
  {
    id: 'A04',
    name: '财大科技园',
    category: '产业园区',
    address: '上海市杨浦区国定路335号',
    lng: 121.5118, lat: 31.2966,
    drive: '3 分钟 / 0.8 km',
    transit: '步行 9 分钟 · 10号线五角场站',
    developer: '上海财经大学',
    area: '约 6.5 万㎡',
    year: 2003,
    subs: ['1号楼', '2号楼', '3号楼', '创业孵化中心'],
    tags: ['国家级科技园', '高校系'],
    note: '依托上海财经大学,聚焦财经金融创新企业。',
    propertyFee: 16,
    nature: '高校系 · 国家级科技园',
    tenants: [
      { name: '立信会计师事务所 杨浦所',   industry: '财税咨询',           area: 2200, rent: 4.0, rentFreeMonths: 3, dealYear: 2023 },
      { name: '华泰证券 五角场营业部',     industry: '金融 / 证券',        area: 1400, rent: 4.2, rentFreeMonths: 3, dealYear: 2024 },
      { name: '财大教育发展',              industry: '教育培训',           area: 1800, rent: 3.8, rentFreeMonths: 2, dealYear: 2022 },
      { name: '财经数科',                  industry: '金融科技',           area: 1200, rent: 4.1, rentFreeMonths: 4, dealYear: 2024 }
    ]
  },
  {
    id: 'A05',
    name: '复旦科技园（江湾园区）',
    category: '产业园区',
    address: '上海市杨浦区淞沪路2005号',
    lng: 121.5060, lat: 31.3050,
    drive: '5 分钟 / 1.4 km',
    transit: '步行 13 分钟 · 10号线五角场站',
    developer: '复旦大学',
    area: '约 9.2 万㎡',
    year: 2002,
    subs: ['创新中心', '国创中心', '研发楼 A/B/C'],
    tags: ['国家级科技园', '高校系'],
    note: '复旦学术资源溢出地,生命科学+人工智能聚集。',
    propertyFee: 18,
    nature: '高校系 · 国家级科技园',
    tenants: [
      { name: '复旦科技园发展',            industry: '产业园区运营',       area: 4500, rent: 4.5, rentFreeMonths: 3, dealYear: 2022 },
      { name: '国家集成电路创新中心',      industry: '集成电路',           area: 3800, rent: 4.6, rentFreeMonths: 4, dealYear: 2023 },
      { name: '西默科技',                  industry: '集成电路 / 芯片',    area: 1600, rent: 4.4, rentFreeMonths: 3, dealYear: 2024 },
      { name: '复旦 AI Lab 商业化中心',    industry: '人工智能',           area: 2100, rent: 4.7, rentFreeMonths: 4, dealYear: 2024 },
      { name: '基因港 生物科技',           industry: '生物医药',           area: 1500, rent: 4.5, rentFreeMonths: 3, dealYear: 2023 }
    ]
  },
  {
    id: 'A06',
    name: '太平洋森活天地（创智天地一期）',
    category: '众创空间',
    address: '上海市杨浦区淞沪路333号',
    lng: 121.5082, lat: 31.2964,
    drive: '3 分钟 / 0.6 km',
    transit: '步行 5 分钟 · 10号线五角场站',
    developer: '瑞安房地产',
    area: '约 4 万㎡',
    year: 2012,
    subs: ['M 栋', 'N 栋', '中心广场'],
    tags: ['街区式办公', '生活配套'],
    note: '街区式办公商业混合体,深受创业者欢迎。',
    propertyFee: 20,
    nature: '民营 · 街区式联合办公',
    tenants: [
      { name: '苏宁易购 创新研发',         industry: '电商 / 零售',        area: 2200, rent: 5.0, rentFreeMonths: 3, dealYear: 2023 },
      { name: '大众点评 杨浦孵化点',       industry: '互联网 / 本地生活',  area: 1600, rent: 5.1, rentFreeMonths: 3, dealYear: 2024 },
      { name: '微创网络',                  industry: '互联网 / 软件',      area: 900,  rent: 4.9, rentFreeMonths: 2, dealYear: 2024 }
    ]
  },
  {
    id: 'A07',
    name: '创智天地企业中心 1-3 号楼',
    category: '商办写字楼',
    address: '上海市杨浦区淞沪路290号',
    lng: 121.5050, lat: 31.2955,
    drive: '4 分钟 / 1.0 km',
    transit: '步行 9 分钟 · 10号线江湾体育场站',
    developer: '瑞安房地产',
    area: '约 22 万㎡',
    year: 2010,
    subs: ['1号楼', '2号楼', '3号楼', '能源中心'],
    tags: ['甲级写字楼', '高科技园区'],
    note: 'KIC 核心办公组团,世界 500 强 R&D 集聚。',
    propertyFee: 30,
    nature: '民营 · 甲级写字楼',
    tenants: [
      { name: 'IBM 上海创新研发中心',      industry: 'TMT / 软件',         area: 12000, rent: 6.5, rentFreeMonths: 5, dealYear: 2022 },
      { name: 'Microsoft 加速器·上海',     industry: 'TMT / 软件',         area: 4800,  rent: 6.6, rentFreeMonths: 5, dealYear: 2023 },
      { name: 'AstraZeneca 中国 R&D',      industry: '生物医药',           area: 8500,  rent: 6.4, rentFreeMonths: 5, dealYear: 2022 },
      { name: 'EMC²（Dell）研发中心',      industry: 'TMT / 软件',         area: 6200,  rent: 6.3, rentFreeMonths: 4, dealYear: 2023 },
      { name: '艾迪斯翻译',                industry: '专业服务',           area: 1100,  rent: 6.0, rentFreeMonths: 3, dealYear: 2024 }
    ]
  },

  /* ========================================================================
   * ===============================  1~2 km  ===============================
   * ====================================================================== */
  {
    id: 'B01',
    name: '创智天地企业中心 5-10 号楼',
    category: '商办写字楼',
    address: '上海市杨浦区伟德路 / 大学路一带',
    lng: 121.5022, lat: 31.2940,
    drive: '6 分钟 / 1.8 km',
    transit: '步行 6 分钟 · 10号线江湾体育场站',
    developer: '瑞安房地产',
    area: '约 18 万㎡',
    year: 2013,
    subs: ['5号楼', '6号楼', '7号楼', '8号楼', '9号楼', '10号楼'],
    tags: ['甲级写字楼', 'TMT 集聚'],
    note: 'KIC 二期组团,腾讯、IBM、EMC 等驻扎。',
    propertyFee: 30,
    nature: '民营 · 甲级写字楼',
    tenants: [
      { name: '腾讯 上海创新中心',         industry: 'TMT / 互联网',       area: 10500, rent: 6.4, rentFreeMonths: 5, dealYear: 2022 },
      { name: '网易有道 杨浦研发',         industry: '教育科技',           area: 5200,  rent: 6.2, rentFreeMonths: 4, dealYear: 2023 },
      { name: '哔哩哔哩 部分团队',         industry: '文娱传媒',           area: 4800,  rent: 6.3, rentFreeMonths: 4, dealYear: 2023 },
      { name: '美团（大众点评事业群）',    industry: 'TMT / 本地生活',     area: 7600,  rent: 6.4, rentFreeMonths: 5, dealYear: 2022 },
      { name: '甲骨文 中国 R&D 部分',      industry: 'TMT / 软件',         area: 3200,  rent: 6.3, rentFreeMonths: 4, dealYear: 2023 }
    ]
  },
  {
    id: 'B02',
    name: '同济联合广场',
    category: '商办写字楼',
    address: '上海市杨浦区国康路100号',
    lng: 121.5023, lat: 31.2880,
    drive: '8 分钟 / 2.0 km',
    transit: '10号线同济大学站 步行 4 分钟',
    developer: '同济资产',
    area: '约 8 万㎡',
    year: 2009,
    subs: ['A 座', 'B 座'],
    tags: ['高校系', '设计创意'],
    note: '紧邻同济大学,设计、咨询行业为主。',
    propertyFee: 22,
    nature: '高校系 · 乙级写字楼',
    tenants: [
      { name: '同济建筑设计研究院 部分',   industry: '建筑设计',           area: 5800, rent: 5.0, rentFreeMonths: 3, dealYear: 2022 },
      { name: '启迪设计 上海公司',         industry: '建筑设计',           area: 2200, rent: 4.8, rentFreeMonths: 3, dealYear: 2024 },
      { name: '同济咨询',                  industry: '工程咨询',           area: 1800, rent: 4.6, rentFreeMonths: 2, dealYear: 2023 },
      { name: 'AECOM 上海项目部',          industry: '工程咨询 / 设计',    area: 1600, rent: 5.1, rentFreeMonths: 3, dealYear: 2024 }
    ]
  },
  {
    id: 'B03',
    name: '同济科技园 经纬大厦',
    category: '产业园区',
    address: '上海市杨浦区四平路1239号',
    lng: 121.4988, lat: 31.2845,
    drive: '8 分钟 / 2.1 km',
    transit: '10号线同济大学站 步行 7 分钟',
    developer: '同济科技园',
    area: '约 7.8 万㎡',
    year: 2005,
    subs: ['经纬大厦', '中赤大厦', '设计创意学院园'],
    tags: ['国家级科技园', '设计创意'],
    note: '环同济知识经济圈核心载体之一。',
    propertyFee: 17,
    nature: '高校系 · 国家级科技园',
    tenants: [
      { name: '同济科技股份',              industry: '产业园区运营',       area: 4200, rent: 4.2, rentFreeMonths: 3, dealYear: 2022 },
      { name: '易斯达科技',                industry: '智能制造',           area: 2200, rent: 4.0, rentFreeMonths: 3, dealYear: 2024 },
      { name: '申通设计研究院',            industry: '工程设计',           area: 1900, rent: 4.1, rentFreeMonths: 3, dealYear: 2023 },
      { name: '同济创业谷企业',            industry: '其他 / 综合',        area: 1500, rent: 3.9, rentFreeMonths: 4, dealYear: 2024 }
    ]
  },
  {
    id: 'B04',
    name: '复旦软件园 江湾分园',
    category: '产业园区',
    address: '上海市杨浦区国和路577号',
    lng: 121.5150, lat: 31.3070,
    drive: '6 分钟 / 1.7 km',
    transit: '步行 14 分钟 · 10号线五角场站',
    developer: '复旦科技园',
    area: '约 5.4 万㎡',
    year: 2011,
    subs: ['1号楼', '2号楼', '3号楼'],
    tags: ['软件信息', '高校系'],
    note: '专注软件信息服务、智能制造孵化。',
    propertyFee: 18,
    nature: '高校系 · 软件信息产业园',
    tenants: [
      { name: '上海软件产业研究院',        industry: 'TMT / 软件',         area: 3200, rent: 4.3, rentFreeMonths: 3, dealYear: 2023 },
      { name: '文思海辉 杨浦分部',         industry: 'TMT / IT 服务',      area: 2400, rent: 4.4, rentFreeMonths: 3, dealYear: 2024 },
      { name: '帆软软件 华东中心',         industry: 'TMT / 软件',         area: 1800, rent: 4.2, rentFreeMonths: 3, dealYear: 2024 }
    ]
  },
  {
    id: 'B05',
    name: 'INNO 新业坊（江湾）',
    category: '众创空间',
    address: '上海市杨浦区国权北路1688号',
    lng: 121.5167, lat: 31.2870,
    drive: '7 分钟 / 1.9 km',
    transit: '10号线五角场站换乘公交 ≈ 18 分钟',
    developer: '上海地产',
    area: '约 5.2 万㎡',
    year: 2018,
    subs: ['北区办公', '南区办公', '中央花园'],
    tags: ['老厂房改造', '低密办公'],
    note: '老工业遗存改造,深受文创、设计公司喜爱。',
    propertyFee: 19,
    nature: '国资 · 老厂房改造创意园',
    tenants: [
      { name: '万科 北杨浦研发',           industry: '房地产 / 建筑',      area: 3400, rent: 4.6, rentFreeMonths: 3, dealYear: 2023 },
      { name: '蔚来汽车 部分团队',         industry: '新能源汽车',         area: 2800, rent: 4.7, rentFreeMonths: 4, dealYear: 2024 },
      { name: '创业邦 上海',                industry: '文创 / 服务业',      area: 1600, rent: 4.4, rentFreeMonths: 3, dealYear: 2023 },
      { name: '良设设计 工作室',           industry: '设计创意',           area: 900,  rent: 4.5, rentFreeMonths: 2, dealYear: 2024 }
    ]
  },
  {
    id: 'B06',
    name: '中环国际广场',
    category: '商办写字楼',
    address: '上海市杨浦区国宾路18号',
    lng: 121.4980, lat: 31.3035,
    drive: '7 分钟 / 1.8 km',
    transit: '10号线三门路站 步行 6 分钟',
    developer: '中环集团',
    area: '约 12 万㎡',
    year: 2014,
    subs: ['A 座', 'B 座', 'C 座'],
    tags: ['甲级写字楼'],
    note: '中环线上视觉地标,以金融、贸易企业为主。',
    propertyFee: 26,
    nature: '民营 · 甲级写字楼',
    tenants: [
      { name: '华夏银行 杨浦支行',         industry: '金融 / 银行',        area: 2800, rent: 5.4, rentFreeMonths: 3, dealYear: 2023 },
      { name: '招商证券 区域中心',         industry: '金融 / 证券',        area: 2200, rent: 5.5, rentFreeMonths: 3, dealYear: 2024 },
      { name: '中信建投 杨浦营业部',       industry: '金融 / 证券',        area: 1800, rent: 5.3, rentFreeMonths: 3, dealYear: 2023 },
      { name: '上海贸易促进会 杨浦',       industry: '商协会 / 政府',      area: 1200, rent: 5.0, rentFreeMonths: 2, dealYear: 2022 }
    ]
  },
  {
    id: 'B07',
    name: '五角丰达广场',
    category: '商办写字楼',
    address: '上海市杨浦区翔殷路1099号',
    lng: 121.5260, lat: 31.3060,
    drive: '6 分钟 / 1.6 km',
    transit: '8号线翔殷路站 步行 9 分钟',
    developer: '丰盛集团',
    area: '约 9.5 万㎡',
    year: 2015,
    subs: ['1号楼', '2号楼', '商业裙房'],
    tags: ['商办综合体'],
    note: '近翔殷路高架,通达性强。',
    propertyFee: 22,
    nature: '民营 · 商办综合体',
    tenants: [
      { name: '锦江旅游 区域中心',         industry: '旅游 / 出行',        area: 2200, rent: 4.9, rentFreeMonths: 3, dealYear: 2023 },
      { name: '安永咨询 杨浦点',           industry: '管理咨询',           area: 1600, rent: 5.0, rentFreeMonths: 3, dealYear: 2024 },
      { name: '太平洋保险 北分',           industry: '金融 / 保险',        area: 2400, rent: 4.8, rentFreeMonths: 2, dealYear: 2022 }
    ]
  },
  {
    id: 'B08',
    name: '杨浦科技创业中心（虬江路）',
    category: '众创空间',
    address: '上海市杨浦区国康路46号',
    lng: 121.5008, lat: 31.2868,
    drive: '8 分钟 / 2.0 km',
    transit: '10号线同济大学站 步行 5 分钟',
    developer: '杨浦区科委',
    area: '约 3.6 万㎡',
    year: 2001,
    subs: ['A 楼', 'B 楼', '加速器中心'],
    tags: ['国家级孵化器', '政府运营'],
    note: '杨浦最早的国家级孵化器之一。',
    propertyFee: 14,
    nature: '政府运营 · 国家级孵化器',
    tenants: [
      { name: '杨浦科创集团 总部',         industry: '产业园区运营',       area: 3500, rent: 3.5, rentFreeMonths: 4, dealYear: 2022 },
      { name: '初创团队（孵化期免租）',    industry: '其他 / 综合',        area: 1200, rent: 0,   rentFreeMonths: 12, dealYear: 2024, note: '孵化期 12 个月 0 租金' },
      { name: '杨浦区科创服务中心',        industry: '政府服务',           area: 800,  rent: 3.0, rentFreeMonths: 6, dealYear: 2023 }
    ]
  },
  {
    id: 'B09',
    name: '苏宁生活广场（江湾）',
    category: '商办写字楼',
    address: '上海市杨浦区政立路477号',
    lng: 121.5208, lat: 31.3088,
    drive: '6 分钟 / 1.7 km',
    transit: '10号线江湾体育场站 步行 14 分钟',
    developer: '苏宁置业',
    area: '约 7 万㎡',
    year: 2013,
    subs: ['办公楼 A', '商业部分'],
    tags: ['商业综合体'],
    note: '商办+商业混合体,以中小企业及代理为主。',
    propertyFee: 20,
    nature: '民营 · 商业综合体办公',
    tenants: [
      { name: '苏宁易购 华东运营',         industry: '电商 / 零售',        area: 3800, rent: 4.6, rentFreeMonths: 3, dealYear: 2023 },
      { name: '苏宁体育 区域',              industry: '体育 / 文娱',        area: 2200, rent: 4.5, rentFreeMonths: 3, dealYear: 2024 },
      { name: '苏宁支付 部分',              industry: '金融科技',           area: 1400, rent: 4.7, rentFreeMonths: 3, dealYear: 2024 }
    ]
  },
  {
    id: 'B10',
    name: '国和路1号',
    category: '商办写字楼',
    address: '上海市杨浦区国和路1号',
    lng: 121.5140, lat: 31.2890,
    drive: '7 分钟 / 1.9 km',
    transit: '10号线五角场站 步行 13 分钟',
    developer: '杨浦城投',
    area: '约 6 万㎡',
    year: 2016,
    subs: ['主塔', '裙楼'],
    tags: ['总部办公'],
    note: '中小型总部企业聚集。',
    propertyFee: 21,
    nature: '国资 · 乙级写字楼',
    tenants: [
      { name: '上海建工 北区分公司',       industry: '建筑施工',           area: 3200, rent: 4.8, rentFreeMonths: 3, dealYear: 2023 },
      { name: '杨浦城投 子公司',           industry: '产业园区运营',       area: 2800, rent: 4.5, rentFreeMonths: 3, dealYear: 2022 },
      { name: '九派资本',                  industry: '金融 / 投资',        area: 1100, rent: 5.0, rentFreeMonths: 3, dealYear: 2024 }
    ]
  },

  /* ========================================================================
   * ===============================  2~3 km  ===============================
   * ====================================================================== */
  {
    id: 'C01',
    name: '长阳创谷',
    category: '产业园区',
    address: '上海市杨浦区长阳路1687号',
    lng: 121.5170, lat: 31.2735,
    drive: '12 分钟 / 3.0 km',
    transit: '12号线宁国路站 步行 16 分钟',
    developer: '长阳谷投资',
    area: '约 18 万㎡',
    year: 2016,
    subs: ['1号楼', '2号楼', '3号楼', '4号楼', '5号楼', '6号楼'],
    tags: ['人工智能', '老厂房改造', '国家级双创基地'],
    note: '原中国纺织机械厂改造,人工智能产业聚集地。',
    propertyFee: 18,
    nature: '国资 · 国家级双创示范基地',
    tenants: [
      { name: '商汤科技 长阳子公司',       industry: '人工智能',           area: 5800, rent: 4.8, rentFreeMonths: 4, dealYear: 2023 },
      { name: '优刻得 UCloud',             industry: '云计算 / TMT',       area: 7400, rent: 4.9, rentFreeMonths: 4, dealYear: 2022 },
      { name: '智齿科技',                  industry: 'TMT / SaaS',         area: 2200, rent: 4.6, rentFreeMonths: 3, dealYear: 2024 },
      { name: '启明医疗 上海',              industry: '生物医药',           area: 3000, rent: 4.7, rentFreeMonths: 4, dealYear: 2023 },
      { name: '依图科技 部分',              industry: '人工智能',           area: 2600, rent: 4.7, rentFreeMonths: 3, dealYear: 2023 },
      { name: '网龙网络 长阳运营',         industry: '游戏 / TMT',         area: 1800, rent: 4.5, rentFreeMonths: 3, dealYear: 2024 }
    ]
  },
  {
    id: 'C02',
    name: '上海湾谷科技园',
    category: '产业园区',
    address: '上海市杨浦区殷高东路66号',
    lng: 121.5300, lat: 31.3340,
    drive: '12 分钟 / 3.0 km',
    transit: '3号线殷高西路站 步行 12 分钟',
    developer: '湾谷集团 / 中体集团',
    area: '约 60 万㎡',
    year: 2017,
    subs: ['A 区', 'B 区', 'C 区', 'D 区', '研发中心'],
    tags: ['大体量', '生命科学', '总部基地'],
    note: '复旦张江、生命科学头部企业集聚地。',
    propertyFee: 22,
    nature: '民营 · 总部基地 · 甲级研发园',
    tenants: [
      { name: '复旦张江 上海总部',         industry: '生物医药',           area: 12000, rent: 5.2, rentFreeMonths: 5, dealYear: 2022 },
      { name: '君实生物 研发中心',         industry: '生物医药',           area: 8400,  rent: 5.1, rentFreeMonths: 5, dealYear: 2023 },
      { name: '兆易创新 上海',              industry: '集成电路',           area: 6500,  rent: 5.0, rentFreeMonths: 4, dealYear: 2023 },
      { name: '神州数码 华东中心',         industry: 'TMT / IT 服务',      area: 5200,  rent: 5.0, rentFreeMonths: 4, dealYear: 2024 },
      { name: '健麾信息',                  industry: '医疗信息化',         area: 3800,  rent: 4.9, rentFreeMonths: 3, dealYear: 2024 }
    ]
  },
  {
    id: 'C03',
    name: '上海理工大学国家科技园',
    category: '产业园区',
    address: '上海市杨浦区军工路1100号',
    lng: 121.5505, lat: 31.2870,
    drive: '14 分钟 / 3.5 km',
    transit: '8号线翔殷路站换乘公交 ≈ 25 分钟',
    developer: '上理工资产',
    area: '约 10 万㎡',
    year: 2008,
    subs: ['1号楼', '2号楼', '3号楼'],
    tags: ['国家级科技园', '高校系'],
    note: '机械、能源、医疗器械方向。',
    propertyFee: 15,
    nature: '高校系 · 国家级科技园',
    tenants: [
      { name: '联影医疗 部分团队',         industry: '医疗器械',           area: 4200, rent: 4.0, rentFreeMonths: 4, dealYear: 2023 },
      { name: '大族激光 上海',              industry: '智能制造 / 激光',    area: 2800, rent: 3.9, rentFreeMonths: 3, dealYear: 2024 },
      { name: '上理工 能源研究院产业化',   industry: '新能源',             area: 1800, rent: 3.7, rentFreeMonths: 3, dealYear: 2023 }
    ]
  },
  {
    id: 'C04',
    name: '大学路创意街区',
    category: '众创空间',
    address: '上海市杨浦区大学路 80-180 号',
    lng: 121.5070, lat: 31.2947,
    drive: '6 分钟 / 1.5 km',
    transit: '10号线江湾体育场站 步行 6 分钟',
    developer: '瑞安房地产',
    area: '约 4 万㎡（沿街）',
    year: 2010,
    subs: ['沿街多栋低层街铺+小型办公'],
    tags: ['街区办公', '小型工作室'],
    note: '上海最具人气的创业街区之一。',
    propertyFee: 18,
    nature: '民营 · 街区式创意办公',
    tenants: [
      { name: 'Seesaw Coffee 总部',        industry: '消费品牌 / 餐饮',    area: 600,  rent: 5.8, rentFreeMonths: 2, dealYear: 2024 },
      { name: '良设设计',                  industry: '设计创意',           area: 800,  rent: 5.5, rentFreeMonths: 2, dealYear: 2023 },
      { name: '小红书 杨浦小型工作室',     industry: 'TMT / 社交',         area: 700,  rent: 5.6, rentFreeMonths: 2, dealYear: 2024 },
      { name: '众多小型工作室合计',        industry: '文创 / 设计',        area: 8000, rent: 5.2, rentFreeMonths: 2, dealYear: 2024 }
    ]
  },
  {
    id: 'C05',
    name: '复地新都国际',
    category: '商办写字楼',
    address: '上海市杨浦区四平路1500号',
    lng: 121.4955, lat: 31.2790,
    drive: '12 分钟 / 2.9 km',
    transit: '10号线同济大学站 步行 11 分钟',
    developer: '复地集团',
    area: '约 8 万㎡',
    year: 2012,
    subs: ['办公塔', '商业'],
    tags: ['商办综合体'],
    note: '同济商圈办公载体之一。',
    propertyFee: 22,
    nature: '民营 · 商办综合体',
    tenants: [
      { name: '复地集团 北区',              industry: '房地产 / 建筑',      area: 3600, rent: 4.8, rentFreeMonths: 3, dealYear: 2022 },
      { name: '同程艺龙 杨浦',              industry: '互联网 / 出行',      area: 2200, rent: 4.9, rentFreeMonths: 3, dealYear: 2024 },
      { name: '中信建投 同济营业部',       industry: '金融 / 证券',        area: 1400, rent: 4.7, rentFreeMonths: 3, dealYear: 2023 }
    ]
  },
  {
    id: 'C06',
    name: '黄兴公园万达广场（办公）',
    category: '商办写字楼',
    address: '上海市杨浦区双阳路100号',
    lng: 121.5346, lat: 31.2884,
    drive: '12 分钟 / 2.9 km',
    transit: '8号线黄兴公园站 步行 5 分钟',
    developer: '万达集团',
    area: '约 6 万㎡（办公）',
    year: 2017,
    subs: ['办公塔', '万达广场商业'],
    tags: ['商业综合体'],
    note: '消费氛围浓厚,办公以服务业为主。',
    propertyFee: 21,
    nature: '民营 · 商业综合体办公',
    tenants: [
      { name: '万达广场 商管',              industry: '商业管理',           area: 2400, rent: 4.5, rentFreeMonths: 3, dealYear: 2023 },
      { name: '太平洋人寿 黄兴部',          industry: '金融 / 保险',        area: 2000, rent: 4.4, rentFreeMonths: 3, dealYear: 2024 },
      { name: '招商证券 黄兴',              industry: '金融 / 证券',        area: 1500, rent: 4.6, rentFreeMonths: 3, dealYear: 2023 }
    ]
  },
  {
    id: 'C07',
    name: '上海湾谷科技园（北区）',
    category: '产业园区',
    address: '上海市杨浦区殷高东路 99 号',
    lng: 121.5340, lat: 31.3380,
    drive: '13 分钟 / 3.1 km',
    transit: '3号线殷高西路站 步行 15 分钟',
    developer: '湾谷集团',
    area: '约 14 万㎡',
    year: 2019,
    subs: ['E 区', 'F 区', 'G 区'],
    tags: ['新落成', '生命科学'],
    note: '湾谷北区新增组团,室外低密。',
    propertyFee: 22,
    nature: '民营 · 总部基地 · 甲级研发园',
    tenants: [
      { name: '微创医疗 上海创新',         industry: '医疗器械',           area: 6800, rent: 5.0, rentFreeMonths: 5, dealYear: 2023 },
      { name: '西门子医疗 部分团队',       industry: '医疗器械',           area: 4500, rent: 5.1, rentFreeMonths: 5, dealYear: 2024 },
      { name: '罗氏诊断 中国部分',         industry: '生物医药',           area: 3800, rent: 5.0, rentFreeMonths: 4, dealYear: 2023 },
      { name: '麦科林生物',                industry: '生物医药',           area: 1800, rent: 4.8, rentFreeMonths: 4, dealYear: 2024 }
    ]
  },
  {
    id: 'C08',
    name: '创智坊（KIC 商住板块）',
    category: '商办写字楼',
    address: '上海市杨浦区伟德路100号',
    lng: 121.5008, lat: 31.2902,
    drive: '8 分钟 / 2.1 km',
    transit: '10号线江湾体育场站 步行 12 分钟',
    developer: '瑞安房地产',
    area: '约 9 万㎡',
    year: 2015,
    subs: ['办公组团', '住宅组团'],
    tags: ['街区办公', '高品质'],
    note: '商住一体的高端街区。',
    propertyFee: 26,
    nature: '民营 · 街区式甲级办公',
    tenants: [
      { name: '大众点评 部分团队',         industry: 'TMT / 本地生活',     area: 3400, rent: 5.6, rentFreeMonths: 4, dealYear: 2023 },
      { name: '默克 Merck 中国部分',       industry: '生物医药',           area: 2800, rent: 5.7, rentFreeMonths: 4, dealYear: 2024 },
      { name: '拼多多 (早期孵化)',          industry: 'TMT / 电商',         area: 1800, rent: 5.5, rentFreeMonths: 3, dealYear: 2022 }
    ]
  },
  {
    id: 'C09',
    name: 'INNO 创智（江湾湿地）',
    category: '产业园区',
    address: '上海市杨浦区国和路1688号',
    lng: 121.5260, lat: 31.2820,
    drive: '11 分钟 / 2.7 km',
    transit: '10号线江湾体育场站 步行 18 分钟',
    developer: '上海地产',
    area: '约 8.6 万㎡',
    year: 2020,
    subs: ['T1', 'T2', 'T3', '联合办公中心'],
    tags: ['低密办公', '生态办公'],
    note: '生态化新办公载体,初创及成长型企业聚集。',
    propertyFee: 19,
    nature: '国资 · 生态低密产业园',
    tenants: [
      { name: '蔚来汽车 设计中心',         industry: '新能源汽车',         area: 4200, rent: 4.5, rentFreeMonths: 4, dealYear: 2023 },
      { name: '海尔创牌 上海',              industry: '智能家电',           area: 2800, rent: 4.4, rentFreeMonths: 3, dealYear: 2024 },
      { name: '小马智行 上海',              industry: '自动驾驶',           area: 2400, rent: 4.6, rentFreeMonths: 4, dealYear: 2024 },
      { name: '联合办公租户合计',          industry: '其他 / 综合',        area: 5800, rent: 4.3, rentFreeMonths: 3, dealYear: 2024 }
    ]
  },
  {
    id: 'C10',
    name: '杨浦滨江 · 绿之丘',
    category: '众创空间',
    address: '上海市杨浦区杨树浦路 1500 号',
    lng: 121.5360, lat: 31.2720,
    drive: '14 分钟 / 3.0 km',
    transit: '12号线爱国路站 步行 18 分钟',
    developer: '杨浦滨江',
    area: '约 2.4 万㎡',
    year: 2019,
    subs: ['主馆', '展览中心'],
    tags: ['滨江文创', '改造项目'],
    note: '工业遗存改造的文创办公地标。',
    propertyFee: 16,
    nature: '国资 · 滨江工业遗存改造',
    tenants: [
      { name: '杨浦滨江 运营平台',         industry: '产业园区运营',       area: 1600, rent: 4.2, rentFreeMonths: 3, dealYear: 2022 },
      { name: '上海风语筑 部分',            industry: '文创 / 展览',        area: 1200, rent: 4.3, rentFreeMonths: 3, dealYear: 2024 },
      { name: '设计师工作室合计',          industry: '设计创意',           area: 2200, rent: 4.0, rentFreeMonths: 2, dealYear: 2024 }
    ]
  },

  /* ========================================================================
   * ===============================  3~5 km  ===============================
   * ====================================================================== */
  {
    id: 'D01',
    name: '复地 · 活力城（办公）',
    category: '商办写字楼',
    address: '上海市杨浦区控江路1188号',
    lng: 121.5310, lat: 31.2670,
    drive: '16 分钟 / 3.6 km',
    transit: '8号线黄兴公园站 步行 14 分钟',
    developer: '复地集团',
    area: '约 14 万㎡',
    year: 2014,
    subs: ['T1', 'T2', '商业街'],
    tags: ['控江商圈'],
    note: '控江路商圈代表性商办综合体。',
    propertyFee: 23,
    nature: '民营 · 商办综合体',
    tenants: [
      { name: '复地集团 杨浦区域',          industry: '房地产 / 建筑',      area: 3800, rent: 4.7, rentFreeMonths: 3, dealYear: 2022 },
      { name: '中国人寿 控江营业部',       industry: '金融 / 保险',        area: 2400, rent: 4.6, rentFreeMonths: 3, dealYear: 2023 },
      { name: '光明乳业 销售中心',          industry: '快消食品',           area: 1800, rent: 4.5, rentFreeMonths: 3, dealYear: 2024 }
    ]
  },
  {
    id: 'D02',
    name: '长阳谷创意园',
    category: '产业园区',
    address: '上海市杨浦区长阳路 1761 号',
    lng: 121.5188, lat: 31.2705,
    drive: '15 分钟 / 3.4 km',
    transit: '12号线宁国路站 步行 14 分钟',
    developer: '杨浦科创',
    area: '约 4.6 万㎡',
    year: 2018,
    subs: ['1号楼', '2号楼', '3号楼'],
    tags: ['老厂房改造', '创意设计'],
    note: '与长阳创谷相邻,差异化做创意设计聚集。',
    propertyFee: 16,
    nature: '民营 · 老厂房改造创意园',
    tenants: [
      { name: '风语筑展览 总部',            industry: '文创 / 展览',        area: 3200, rent: 4.2, rentFreeMonths: 3, dealYear: 2023 },
      { name: '哈罗工业设计',              industry: '设计创意',           area: 1600, rent: 4.0, rentFreeMonths: 3, dealYear: 2024 },
      { name: '小型文创团队合计',          industry: '文创 / 设计',        area: 4400, rent: 3.9, rentFreeMonths: 2, dealYear: 2024 }
    ]
  },
  {
    id: 'D03',
    name: '中环国际商务广场',
    category: '商办写字楼',
    address: '上海市杨浦区周家嘴路3115号',
    lng: 121.5470, lat: 31.2710,
    drive: '18 分钟 / 4.5 km',
    transit: '12号线爱国路站 步行 15 分钟',
    developer: '中环集团',
    area: '约 9 万㎡',
    year: 2016,
    subs: ['A 座', 'B 座'],
    tags: ['物流贸易'],
    note: '近港口,以物流、贸易类公司为主。',
    propertyFee: 20,
    nature: '民营 · 乙级写字楼',
    tenants: [
      { name: '中外运 杨浦分公司',         industry: '物流 / 贸易',        area: 3200, rent: 4.3, rentFreeMonths: 3, dealYear: 2023 },
      { name: '中谷海运 区域',              industry: '物流 / 贸易',        area: 2400, rent: 4.2, rentFreeMonths: 3, dealYear: 2024 },
      { name: '上海港集 大宗',              industry: '物流 / 贸易',        area: 1800, rent: 4.1, rentFreeMonths: 2, dealYear: 2024 }
    ]
  },
  {
    id: 'D04',
    name: '优族 188 创意园',
    category: '众创空间',
    address: '上海市杨浦区国权北路188号',
    lng: 121.5290, lat: 31.3260,
    drive: '14 分钟 / 3.6 km',
    transit: '3号线殷高西路站 步行 17 分钟',
    developer: '杨浦科技',
    area: '约 3.2 万㎡',
    year: 2014,
    subs: ['A 区', 'B 区'],
    tags: ['文创办公'],
    note: '小体量文创办公,租金性价比突出。',
    propertyFee: 13,
    nature: '民营 · 小体量文创园',
    tenants: [
      { name: '锐叠文化 工作室',            industry: '文创 / 影视',        area: 900,  rent: 3.5, rentFreeMonths: 3, dealYear: 2024 },
      { name: '独立设计工作室合计',        industry: '设计创意',           area: 3200, rent: 3.4, rentFreeMonths: 3, dealYear: 2024 }
    ]
  },
  {
    id: 'D05',
    name: '同济创意设计园（赤峰路）',
    category: '产业园区',
    address: '上海市杨浦区赤峰路63号',
    lng: 121.4920, lat: 31.2785,
    drive: '15 分钟 / 3.5 km',
    transit: '10号线同济大学站 步行 10 分钟',
    developer: '同济科技园',
    area: '约 6 万㎡',
    year: 2007,
    subs: ['沿街多栋设计创意工作室'],
    tags: ['设计创意', '街区办公'],
    note: '同济设计圈核心组团之一。',
    propertyFee: 15,
    nature: '高校系 · 设计创意街区',
    tenants: [
      { name: '同济设计研究院 部分',       industry: '建筑设计',           area: 3800, rent: 4.0, rentFreeMonths: 3, dealYear: 2022 },
      { name: '日清设计',                  industry: '建筑设计',           area: 1800, rent: 3.9, rentFreeMonths: 3, dealYear: 2024 },
      { name: '众多小型设计工作室',        industry: '设计创意',           area: 4200, rent: 3.7, rentFreeMonths: 2, dealYear: 2024 }
    ]
  },
  {
    id: 'D06',
    name: '上海财大科技园（武川路园）',
    category: '产业园区',
    address: '上海市杨浦区武川路111号',
    lng: 121.5085, lat: 31.2920,
    drive: '7 分钟 / 1.7 km',
    transit: '10号线江湾体育场站 步行 8 分钟',
    developer: '上海财经大学',
    area: '约 4 万㎡',
    year: 2010,
    subs: ['1号楼', '2号楼'],
    tags: ['高校系', '财经创新'],
    note: '财大科技园次园区,聚焦财税咨询。',
    propertyFee: 15,
    nature: '高校系 · 财税咨询园',
    tenants: [
      { name: '锦天城律师 杨浦点',          industry: '法律服务',           area: 1800, rent: 4.0, rentFreeMonths: 3, dealYear: 2023 },
      { name: '中汇会计 上海',              industry: '财税咨询',           area: 1400, rent: 3.9, rentFreeMonths: 3, dealYear: 2024 },
      { name: '财大税法研究院产业化',      industry: '财税咨询',           area: 1200, rent: 3.7, rentFreeMonths: 4, dealYear: 2023 }
    ]
  },
  {
    id: 'D07',
    name: '华东理工科技园（部分,东区）',
    category: '产业园区',
    address: '上海市杨浦区国和路300号',
    lng: 121.5188, lat: 31.2860,
    drive: '9 分钟 / 2.2 km',
    transit: '10号线江湾体育场站 步行 19 分钟',
    developer: '华东理工',
    area: '约 3.5 万㎡',
    year: 2009,
    subs: ['1号楼', '2号楼'],
    tags: ['高校系', '化学化工'],
    note: '化学、能源相关初创聚集。',
    propertyFee: 14,
    nature: '高校系 · 化工产业园',
    tenants: [
      { name: '华理工资产 经营',            industry: '产业园区运营',       area: 1800, rent: 3.5, rentFreeMonths: 3, dealYear: 2023 },
      { name: '能特科技',                  industry: '新材料',             area: 1200, rent: 3.6, rentFreeMonths: 3, dealYear: 2024 },
      { name: '化工初创团队合计',          industry: '新能源 / 化工',      area: 2400, rent: 3.4, rentFreeMonths: 4, dealYear: 2024 }
    ]
  },
  {
    id: 'D08',
    name: '杨树浦六厂老仓库（绿地 · 东外滩）',
    category: '众创空间',
    address: '上海市杨浦区杨树浦路2200号',
    lng: 121.5460, lat: 31.2680,
    drive: '20 分钟 / 4.7 km',
    transit: '12号线爱国路站 步行 22 分钟',
    developer: '绿地集团',
    area: '约 2.1 万㎡',
    year: 2021,
    subs: ['仓库 A', '仓库 B', '滨江平台'],
    tags: ['滨江', '工业遗存'],
    note: '滨江段工业遗存改造,新兴文创办公。',
    propertyFee: 17,
    nature: '民营 · 滨江改造文创园',
    tenants: [
      { name: '绿地文创 杨浦',              industry: '文创 / 展览',        area: 1600, rent: 4.5, rentFreeMonths: 3, dealYear: 2023 },
      { name: '滨江品牌孵化',              industry: '消费品牌',           area: 1200, rent: 4.4, rentFreeMonths: 3, dealYear: 2024 }
    ]
  },
  {
    id: 'D09',
    name: '杨浦 INNO 滨江',
    category: '产业园区',
    address: '上海市杨浦区杨树浦路2061号',
    lng: 121.5440, lat: 31.2660,
    drive: '19 分钟 / 4.6 km',
    transit: '12号线爱国路站 步行 17 分钟',
    developer: '上海地产',
    area: '约 7.4 万㎡',
    year: 2022,
    subs: ['T1', 'T2'],
    tags: ['滨江办公', '新一代写字楼'],
    note: '北外滩-杨浦滨江办公新势力。',
    propertyFee: 22,
    nature: '国资 · 新一代滨江办公',
    tenants: [
      { name: '上海地产 滨江运营',         industry: '产业园区运营',       area: 3500, rent: 5.0, rentFreeMonths: 4, dealYear: 2023 },
      { name: '新锐设计师事务所合计',      industry: '设计创意',           area: 4200, rent: 4.9, rentFreeMonths: 4, dealYear: 2024 },
      { name: '滨江总部企业',              industry: '其他 / 综合',        area: 3800, rent: 5.1, rentFreeMonths: 4, dealYear: 2024 }
    ]
  },
  {
    id: 'D10',
    name: '虹口足球场宝矿国际',
    category: '商办写字楼',
    address: '上海市虹口区东江湾路188号',
    lng: 121.4880, lat: 31.2745,
    drive: '18 分钟 / 4.2 km',
    transit: '3/8号线虹口足球场站 步行 5 分钟',
    developer: '宝矿集团',
    area: '约 11 万㎡',
    year: 2013,
    subs: ['1座', '2座'],
    tags: ['交通便利', '甲级写字楼'],
    note: '紧邻虹口足球场,3/8号线换乘。',
    propertyFee: 28,
    nature: '民营 · 甲级写字楼',
    tenants: [
      { name: '韩国 LG 上海中心 部分',     industry: '消费电子',           area: 4200, rent: 5.6, rentFreeMonths: 4, dealYear: 2023 },
      { name: '韩国乐天 上海',              industry: '消费零售',           area: 3200, rent: 5.5, rentFreeMonths: 4, dealYear: 2024 },
      { name: '宝矿集团 总部',              industry: '能源 / 矿业',        area: 2800, rent: 5.4, rentFreeMonths: 3, dealYear: 2022 }
    ]
  },
  {
    id: 'D11',
    name: '中山北二路绿地科技园',
    category: '产业园区',
    address: '上海市虹口区中山北二路1515号',
    lng: 121.4880, lat: 31.2820,
    drive: '17 分钟 / 4.0 km',
    transit: '8号线虹口足球场站 步行 11 分钟',
    developer: '绿地集团',
    area: '约 6.8 万㎡',
    year: 2015,
    subs: ['1号楼', '2号楼', '3号楼'],
    tags: ['绿地系', '科创产业'],
    note: '虹口与杨浦交界,辐射两区。',
    propertyFee: 18,
    nature: '民营 · 科创产业园',
    tenants: [
      { name: '绿地金创 信息技术',          industry: 'TMT / 软件',         area: 2800, rent: 4.4, rentFreeMonths: 3, dealYear: 2023 },
      { name: '绿地大基建 设计',            industry: '建筑设计',           area: 2200, rent: 4.3, rentFreeMonths: 3, dealYear: 2024 },
      { name: '中小型 IT 企业合计',         industry: 'TMT / 软件',         area: 3600, rent: 4.1, rentFreeMonths: 3, dealYear: 2024 }
    ]
  },
  {
    id: 'D13',
    name: '宝武+(BAOWU)创新园',
    category: '产业园区',
    address: '上海市杨浦区周家嘴路2999号',
    lng: 121.5530, lat: 31.2790,
    drive: '20 分钟 / 4.8 km',
    transit: '12号线巨峰路站 步行 26 分钟',
    developer: '宝武集团',
    area: '约 8 万㎡',
    year: 2020,
    subs: ['1号楼', '2号楼'],
    tags: ['国企背景', '材料科技'],
    note: '依托宝武集团,聚焦新材料、装备研发。',
    propertyFee: 17,
    nature: '国资 · 新材料产业园',
    tenants: [
      { name: '宝武集团 数字中心',          industry: '智能制造 / 钢铁',    area: 4200, rent: 4.5, rentFreeMonths: 4, dealYear: 2023 },
      { name: '宝武碳业',                  industry: '新材料',             area: 2600, rent: 4.4, rentFreeMonths: 4, dealYear: 2024 },
      { name: '宝信软件 部分',              industry: 'TMT / 工业软件',     area: 2200, rent: 4.6, rentFreeMonths: 4, dealYear: 2024 }
    ]
  }
];
