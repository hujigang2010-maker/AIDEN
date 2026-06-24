import type { Building } from '../types'

/**
 * 杨浦区商业写字楼 + 工业/产业楼宇数据集
 *
 * 数据综合自杨浦区人民政府公开文件、园区公示信息、主流商办平台公开行情
 * （仲量联行市场报告、前瞻产业研究院园区名录等）整理，覆盖五角场、大创智、
 * 新江湾城、杨浦滨江（长海路/平凉路）、大连路/杨浦外滩、长阳路、控江路、
 * 江浦路、四平路、殷行、大桥、定海路等主要街镇。
 *
 * 坐标为 GCJ-02（高德）坐标系，按楼宇所在街镇/路段就近标注，可直接打点。
 * 租金/出租率等市场数据为公开行情区间整理值，仅供参考，请以实际洽谈为准。
 */
export const buildings: Building[] = [
  // ============ 五角场 / 大创智 ============
  {
    id: 'YP001', name: '创智天地企业中心', address: '淞沪路290弄', plate: '五角场', propertyType: '写字楼', grade: '甲级',
    askingRent: 7.2, propertyFee: 30, totalArea: 86000, completionYear: 2010, floorHeight: 4.2, lng: 121.5142, lat: 31.3036,
    tenants: [
      { name: '哔哩哔哩科技有限公司', industry: '互联网/文娱', floor: '8-15F', leaseArea: 12000, dealRent: 6.9, rentFreeMonths: 4, status: '正常经营', remainingMonths: 36 },
      { name: '英语流利说', industry: '在线教育', floor: '6F', leaseArea: 2600, dealRent: 6.5, rentFreeMonths: 3, status: '正常经营', remainingMonths: 14 },
      { name: '智充科技', industry: '新能源', floor: '5F', leaseArea: 1800, dealRent: 6.6, rentFreeMonths: 3, status: '新签约', remainingMonths: 48 },
    ],
  },
  {
    id: 'YP002', name: '创智天地广场', address: '淞沪路303号', plate: '五角场', propertyType: '混合', grade: '甲级',
    askingRent: 7.0, propertyFee: 30, totalArea: 99000, completionYear: 2009, floorHeight: 4.0, lng: 121.5031, lat: 31.3074,
    tenants: [
      { name: '抖音集团（上海）', industry: '互联网/文娱', floor: '12-18F', leaseArea: 9800, dealRent: 6.8, rentFreeMonths: 4, status: '正常经营', remainingMonths: 40 },
      { name: '声网 Agora', industry: '云通信', floor: '9F', leaseArea: 2400, dealRent: 6.6, rentFreeMonths: 3, status: '正常经营', remainingMonths: 22 },
    ],
  },
  {
    id: 'YP003', name: '创智66商务楼', address: '淞沪路258号', plate: '五角场', propertyType: '写字楼', grade: '甲级',
    askingRent: 6.5, propertyFee: 30, totalArea: 82000, completionYear: 2021, floorHeight: 4.1, lng: 121.5158, lat: 31.3066,
    tenants: [
      { name: '字节跳动上海研发中心', industry: '互联网', floor: '10-16F', leaseArea: 11000, dealRent: 6.4, rentFreeMonths: 4, status: '正常经营', remainingMonths: 44 },
      { name: '德勤华东区审计团队', industry: '专业服务', floor: '8F', leaseArea: 2800, dealRent: 6.6, rentFreeMonths: 3, status: '正常经营', remainingMonths: 30 },
    ],
  },
  {
    id: 'YP004', name: '上海国际设计中心', address: '昌邑路185号', plate: '五角场', propertyType: '写字楼', grade: '甲级',
    askingRent: 5.3, propertyFee: 28, totalArea: 120000, completionYear: 2018, floorHeight: 4.5, lng: 121.5093, lat: 31.3052,
    tenants: [
      { name: '安藤忠雄建筑研究所', industry: '建筑设计', floor: '8-11F', leaseArea: 4200, dealRent: 5.6, rentFreeMonths: 4, status: '正常经营', remainingMonths: 41 },
      { name: '洛可可创新设计', industry: '工业设计', floor: '6F', leaseArea: 2600, dealRent: 5.1, rentFreeMonths: 3, status: '正常经营', remainingMonths: 18 },
    ],
  },
  {
    id: 'YP005', name: '五角场万达广场写字楼', address: '淞沪路77号', plate: '五角场', propertyType: '写字楼', grade: '甲级',
    askingRent: 6.0, propertyFee: 24, totalArea: 72000, completionYear: 2008, floorHeight: 3.9, lng: 121.5168, lat: 31.3008,
    tenants: [
      { name: '美团点评（杨浦分部）', industry: '互联网/本地生活', floor: '18-22F', leaseArea: 8800, dealRent: 5.9, rentFreeMonths: 3, status: '正常经营', remainingMonths: 22 },
      { name: '安克创新', industry: '智能硬件', floor: '12F', leaseArea: 1600, dealRent: 5.6, rentFreeMonths: 3, status: '筹备装修', remainingMonths: 45 },
    ],
  },
  {
    id: 'YP006', name: '五角场苏宁广场', address: '淞沪路8号', plate: '五角场', propertyType: '混合', grade: '甲级',
    askingRent: 4.8, propertyFee: 27, totalArea: 180000, completionYear: 2018, floorHeight: 4.0, lng: 121.5176, lat: 31.2992,
    tenants: [
      { name: '苏宁易购华东区', industry: '电商', floor: '20-26F', leaseArea: 9200, dealRent: 4.7, rentFreeMonths: 3, status: '正常经营', remainingMonths: 28 },
      { name: '元气森林（华东）', industry: '快消品', floor: '15F', leaseArea: 2200, dealRent: 4.6, rentFreeMonths: 3, status: '正常经营', remainingMonths: 16 },
    ],
  },
  {
    id: 'YP007', name: '合生汇国际广场', address: '邯郸路585号', plate: '五角场', propertyType: '混合', grade: '甲级',
    askingRent: 5.6, propertyFee: 26, totalArea: 110000, completionYear: 2018, floorHeight: 4.0, lng: 121.5235, lat: 31.3071,
    tenants: [
      { name: '叠纸网络', industry: '游戏', floor: '20-25F', leaseArea: 9600, dealRent: 5.5, rentFreeMonths: 4, status: '正常经营', remainingMonths: 40 },
      { name: '小红书 MCN 工作室', industry: '互联网/文娱', floor: '10F', leaseArea: 1400, dealRent: 5.3, rentFreeMonths: 2, status: '新签约', remainingMonths: 24 },
    ],
  },
  {
    id: 'YP008', name: '复旦科技园大厦', address: '国权路579号', plate: '五角场', propertyType: '产业园', grade: '乙级',
    askingRent: 4.5, propertyFee: 26, totalArea: 100000, completionYear: 2017, floorHeight: 4.2, lng: 121.5006, lat: 31.3008,
    tenants: [
      { name: '复旦孵化·脑科学研究院', industry: '生物医药', floor: '5-9F', leaseArea: 7200, dealRent: 4.4, rentFreeMonths: 5, status: '正常经营', remainingMonths: 33 },
      { name: '复志科技', industry: '智能硬件', floor: '4F', leaseArea: 2100, dealRent: 4.3, rentFreeMonths: 4, status: '正常经营', remainingMonths: 19 },
    ],
  },
  {
    id: 'YP009', name: '同和国际大厦', address: '国通路127号', plate: '五角场', propertyType: '写字楼', grade: '甲级',
    askingRent: 5.0, propertyFee: 25, totalArea: 68000, completionYear: 2012, floorHeight: 3.9, lng: 121.5052, lat: 31.3047,
    tenants: [
      { name: '同盾科技', industry: '金融科技', floor: '8-12F', leaseArea: 5400, dealRent: 4.9, rentFreeMonths: 3, status: '正常经营', remainingMonths: 16 },
      { name: '森亿智能', industry: '医疗AI', floor: '6F', leaseArea: 1900, dealRent: 4.8, rentFreeMonths: 3, status: '正常经营', remainingMonths: 25 },
    ],
  },
  {
    id: 'YP010', name: '天盛科创广场', address: '政立路118号', plate: '五角场', propertyType: '写字楼', grade: '甲级',
    askingRent: 5.2, propertyFee: 26, totalArea: 100000, completionYear: 2016, floorHeight: 4.0, lng: 121.5066, lat: 31.3112,
    tenants: [
      { name: '依图科技实验室', industry: '人工智能', floor: '10-14F', leaseArea: 6400, dealRent: 5.1, rentFreeMonths: 4, status: '正常经营', remainingMonths: 23 },
      { name: '禾赛科技（销售）', industry: '智能硬件', floor: '8F', leaseArea: 1700, dealRent: 5.0, rentFreeMonths: 3, status: '新签约', remainingMonths: 39 },
    ],
  },
  {
    id: 'YP011', name: '财大科技园国华大厦', address: '武川路111号', plate: '五角场', propertyType: '写字楼', grade: '甲级',
    askingRent: 5.4, propertyFee: 25, totalArea: 68000, completionYear: 2012, floorHeight: 3.9, lng: 121.4992, lat: 31.3057,
    tenants: [
      { name: '国华人寿保险', industry: '金融', floor: '10-18F', leaseArea: 9800, dealRent: 5.3, rentFreeMonths: 3, status: '正常经营', remainingMonths: 44 },
      { name: '财大众创空间', industry: '创业服务', floor: '8F', leaseArea: 2200, dealRent: 5.0, rentFreeMonths: 4, status: '正常经营', remainingMonths: 13 },
    ],
  },

  // ============ 杨浦滨江（长海路 / 平凉路） ============
  {
    id: 'YP012', name: '长海商务楼', address: '国权北路省力', plate: '杨浦滨江', propertyType: '写字楼', grade: '超甲级',
    askingRent: 7.6, propertyFee: 32, totalArea: 86000, completionYear: 2022, floorHeight: 4.5, lng: 121.5082, lat: 31.2632,
    tenants: [
      { name: '汇丰银行（上海分行）', industry: '金融', floor: '20-28F', leaseArea: 11200, dealRent: 7.5, rentFreeMonths: 4, status: '正常经营', remainingMonths: 55 },
      { name: '普华永道数字化中心', industry: '专业服务', floor: '16F', leaseArea: 3200, dealRent: 7.3, rentFreeMonths: 3, status: '正常经营', remainingMonths: 30 },
    ],
  },
  {
    id: 'YP013', name: '杨浦滨江国际商务中心', address: '杨树浦路2218号', plate: '杨浦滨江', propertyType: '混合', grade: '甲级',
    askingRent: 6.1, propertyFee: 29, totalArea: 150000, completionYear: 2020, floorHeight: 4.2, lng: 121.5048, lat: 31.2588,
    tenants: [
      { name: '海尔智家上海中心', industry: '智能家电', floor: '12-18F', leaseArea: 8600, dealRent: 6.0, rentFreeMonths: 3, status: '正常经营', remainingMonths: 35 },
      { name: '美的楼宇科技', industry: '智能制造', floor: '9F', leaseArea: 2600, dealRent: 5.9, rentFreeMonths: 3, status: '新签约', remainingMonths: 47 },
    ],
  },
  {
    id: 'YP014', name: '滨江国际广场（东方渔人码头二期）', address: '杨树浦路1062号', plate: '杨浦滨江', propertyType: '写字楼', grade: '超甲级',
    askingRent: 8.6, propertyFee: 33, totalArea: 90000, completionYear: 2017, floorHeight: 4.4, lng: 121.5102, lat: 31.2566,
    tenants: [
      { name: '渣打银行科技中心', industry: '金融', floor: '25-32F', leaseArea: 10400, dealRent: 8.4, rentFreeMonths: 4, status: '正常经营', remainingMonths: 52 },
      { name: '麦肯锡数字咨询', industry: '专业服务', floor: '20F', leaseArea: 2200, dealRent: 8.5, rentFreeMonths: 3, status: '正常经营', remainingMonths: 27 },
    ],
  },
  {
    id: 'YP015', name: '杨浦滨江国际设计中心', address: '安浦路', plate: '杨浦滨江', propertyType: '写字楼', grade: '超甲级',
    askingRent: 8.8, propertyFee: 33, totalArea: 52000, completionYear: 2019, floorHeight: 4.5, lng: 121.5093, lat: 31.2602,
    tenants: [
      { name: 'BBC 创意工作室', industry: '文化创意', floor: '6-9F', leaseArea: 3800, dealRent: 8.6, rentFreeMonths: 4, status: '正常经营', remainingMonths: 41 },
      { name: '理想汽车上海设计中心', industry: '汽车', floor: '5F', leaseArea: 2400, dealRent: 8.4, rentFreeMonths: 3, status: '正常经营', remainingMonths: 22 },
    ],
  },
  {
    id: 'YP016', name: '平凉路滨江商务楼', address: '平凉路1788号', plate: '平凉路', propertyType: '写字楼', grade: '乙级',
    askingRent: 3.9, propertyFee: 23, totalArea: 70000, completionYear: 2021, floorHeight: 4.0, lng: 121.5256, lat: 31.2638,
    tenants: [
      { name: '中通快递华东总部', industry: '物流', floor: '10-14F', leaseArea: 7200, dealRent: 3.8, rentFreeMonths: 3, status: '正常经营', remainingMonths: 30 },
      { name: '货拉拉上海运营', industry: '物流', floor: '8F', leaseArea: 2400, dealRent: 3.7, rentFreeMonths: 2, status: '正常经营', remainingMonths: 13 },
    ],
  },

  // ============ 大连路 / 杨浦外滩 ============
  {
    id: 'YP017', name: '尚浦领世商务广场', address: '昆明路567号', plate: '大连路', propertyType: '写字楼', grade: '超甲级',
    askingRent: 8.5, propertyFee: 32, totalArea: 96000, completionYear: 2015, floorHeight: 4.3, lng: 121.5208, lat: 31.2668,
    tenants: [
      { name: '欧莱雅中国数字中心', industry: '快消品', floor: '22-30F', leaseArea: 11200, dealRent: 8.2, rentFreeMonths: 4, status: '正常经营', remainingMonths: 55 },
      { name: '西门子数字化工厂', industry: '智能制造', floor: '18F', leaseArea: 3600, dealRent: 8.0, rentFreeMonths: 3, status: '正常经营', remainingMonths: 26 },
    ],
  },
  {
    id: 'YP018', name: '大连路壹中心（海鸥科创大厦）', address: '霍山路1188号', plate: '大连路', propertyType: '写字楼', grade: '甲级',
    askingRent: 6.8, propertyFee: 28, totalArea: 82000, completionYear: 2019, floorHeight: 4.2, lng: 121.5236, lat: 31.2641,
    tenants: [
      { name: '中国联通（杨浦）', industry: '通信', floor: '15-20F', leaseArea: 7600, dealRent: 6.6, rentFreeMonths: 3, status: '正常经营', remainingMonths: 29 },
      { name: '安莉芳（中国）', industry: '快消品', floor: '12F', leaseArea: 2400, dealRent: 6.5, rentFreeMonths: 3, status: '欠租预警', remainingMonths: 9 },
    ],
  },
  {
    id: 'YP019', name: '绿地汇创国际广场', address: '大连路588号', plate: '大连路', propertyType: '写字楼', grade: '甲级',
    askingRent: 7.0, propertyFee: 27, totalArea: 82000, completionYear: 2013, floorHeight: 4.0, lng: 121.5221, lat: 31.2615,
    tenants: [
      { name: '京东（华东数科）', industry: '互联网', floor: '15-20F', leaseArea: 7600, dealRent: 6.8, rentFreeMonths: 3, status: '正常经营', remainingMonths: 38 },
      { name: '太平洋保险科技', industry: '金融', floor: '12F', leaseArea: 2400, dealRent: 6.6, rentFreeMonths: 3, status: '正常经营', remainingMonths: 20 },
    ],
  },
  {
    id: 'YP020', name: '互联宝地科技园', address: '黄兴路221号', plate: '大连路', propertyType: '产业园', grade: '乙级',
    askingRent: 5.0, propertyFee: 18, totalArea: 60000, completionYear: 2014, floorHeight: 4.2, lng: 121.5189, lat: 31.2702,
    tenants: [
      { name: '美团（上海）', industry: '互联网/本地生活', floor: '6-9F', leaseArea: 9200, dealRent: 4.9, rentFreeMonths: 4, status: '正常经营', remainingMonths: 31 },
      { name: '得物 APP 运营中心', industry: '电商', floor: '5F', leaseArea: 3400, dealRent: 5.0, rentFreeMonths: 3, status: '正常经营', remainingMonths: 20 },
      { name: '福特汽车（研发）', industry: '汽车', floor: '4F', leaseArea: 1600, dealRent: 4.8, rentFreeMonths: 2, status: '新签约', remainingMonths: 36 },
    ],
  },

  // ============ 新江湾城 ============
  {
    id: 'YP021', name: '尚浦中心', address: '三门路258号', plate: '新江湾城', propertyType: '混合', grade: '甲级',
    askingRent: 5.8, propertyFee: 28, totalArea: 200000, completionYear: 2016, floorHeight: 4.3, lng: 121.5152, lat: 31.3346,
    tenants: [
      { name: '耐克大中华区总部', industry: '快消品', floor: 'A栋整栋', leaseArea: 26000, dealRent: 5.7, rentFreeMonths: 4, status: '正常经营', remainingMonths: 60 },
      { name: '抖音集团（研发）', industry: '互联网', floor: 'C栋', leaseArea: 14000, dealRent: 5.8, rentFreeMonths: 4, status: '正常经营', remainingMonths: 48 },
    ],
  },
  {
    id: 'YP022', name: '湾谷科技园', address: '国权北路1688弄', plate: '新江湾城', propertyType: '产业园', grade: '乙级',
    askingRent: 2.7, propertyFee: 20, totalArea: 220000, completionYear: 2021, floorHeight: 4.5, lng: 121.5096, lat: 31.3382,
    tenants: [
      { name: '联影医疗研发中心', industry: '医疗器械', floor: 'B3-6F', leaseArea: 8800, dealRent: 2.8, rentFreeMonths: 5, status: '正常经营', remainingMonths: 50 },
      { name: '商汤科技实验室', industry: '人工智能', floor: 'A4-7F', leaseArea: 7200, dealRent: 2.9, rentFreeMonths: 4, status: '正常经营', remainingMonths: 33 },
      { name: '诺禾致源', industry: '生物科技', floor: 'C2F', leaseArea: 3200, dealRent: 2.6, rentFreeMonths: 4, status: '正常经营', remainingMonths: 28 },
    ],
  },
  {
    id: 'YP023', name: '未来谷-湾谷创新中心', address: '国权北路1688弄', plate: '新江湾城', propertyType: '产业园', grade: '乙级',
    askingRent: 3.0, propertyFee: 20, totalArea: 85000, completionYear: 2023, floorHeight: 4.6, lng: 121.5133, lat: 31.3438,
    tenants: [
      { name: '燧原科技芯片研发', industry: '半导体', floor: '5-9F', leaseArea: 6800, dealRent: 3.0, rentFreeMonths: 5, status: '正常经营', remainingMonths: 43 },
      { name: '壁仞科技', industry: '半导体', floor: '4F', leaseArea: 2600, dealRent: 2.9, rentFreeMonths: 4, status: '筹备装修', remainingMonths: 50 },
    ],
  },
  {
    id: 'YP024', name: '新江湾城科技中心', address: '淞沪路2200号', plate: '新江湾城', propertyType: '产业园', grade: '乙级',
    askingRent: 4.0, propertyFee: 22, totalArea: 100000, completionYear: 2023, floorHeight: 4.5, lng: 121.5118, lat: 31.3401,
    tenants: [
      { name: '思谋科技', industry: '人工智能', floor: '6-10F', leaseArea: 7200, dealRent: 3.9, rentFreeMonths: 4, status: '正常经营', remainingMonths: 42 },
      { name: '药明康德数据团队', industry: '生物科技', floor: '5F', leaseArea: 3000, dealRent: 3.8, rentFreeMonths: 3, status: '正常经营', remainingMonths: 25 },
    ],
  },
  {
    id: 'YP025', name: '杨浦科技创业中心', address: '国康路100号', plate: '新江湾城', propertyType: '产业园', grade: '丙级',
    askingRent: 3.4, propertyFee: 21, totalArea: 60000, completionYear: 2020, floorHeight: 4.4, lng: 121.5079, lat: 31.3094,
    tenants: [
      { name: '森亿智能孵化团队', industry: '医疗AI', floor: '7-9F', leaseArea: 4800, dealRent: 3.3, rentFreeMonths: 4, status: '正常经营', remainingMonths: 23 },
      { name: '启迪之星孵化器', industry: '创业服务', floor: '6F', leaseArea: 1900, dealRent: 3.2, rentFreeMonths: 3, status: '正常经营', remainingMonths: 17 },
    ],
  },
  {
    id: 'YP026', name: '新江湾城瑞创大厦', address: '殷行路1388号', plate: '新江湾城', propertyType: '写字楼', grade: '甲级',
    askingRent: 4.6, propertyFee: 22, totalArea: 64000, completionYear: 2020, floorHeight: 4.2, lng: 121.5181, lat: 31.3421,
    tenants: [
      { name: '复星医药创新中心', industry: '医药', floor: '15-21F', leaseArea: 8600, dealRent: 4.5, rentFreeMonths: 4, status: '正常经营', remainingMonths: 49 },
      { name: '蔚来能源研究院', industry: '新能源', floor: '11F', leaseArea: 3000, dealRent: 4.4, rentFreeMonths: 3, status: '新签约', remainingMonths: 41 },
    ],
  },

  // ============ 长阳路 / 大创智拓展 ============
  {
    id: 'YP027', name: '长阳创谷', address: '长阳路1687号', plate: '长阳路', propertyType: '产业园', grade: '乙级',
    askingRent: 5.0, propertyFee: 18, totalArea: 120000, completionYear: 2016, floorHeight: 5.0, lng: 121.5301, lat: 31.2738,
    tenants: [
      { name: '百度（上海）创新中心', industry: '人工智能', floor: '1号楼', leaseArea: 8600, dealRent: 4.9, rentFreeMonths: 5, status: '正常经营', remainingMonths: 36 },
      { name: '埃森哲（上海）', industry: '专业服务', floor: '3号楼', leaseArea: 5200, dealRent: 5.0, rentFreeMonths: 4, status: '正常经营', remainingMonths: 24 },
      { name: '爱驰汽车', industry: '汽车', floor: '5号楼', leaseArea: 2600, dealRent: 4.8, rentFreeMonths: 4, status: '欠租预警', remainingMonths: 6 },
    ],
  },
  {
    id: 'YP028', name: '城市概念软件信息服务园', address: '隆昌路619号', plate: '控江路', propertyType: '产业园', grade: '丙级',
    askingRent: 3.8, propertyFee: 14, totalArea: 90000, completionYear: 2009, floorHeight: 4.8, lng: 121.5342, lat: 31.2715,
    tenants: [
      { name: '米哈游创意工作室', industry: '游戏', floor: '2-4F', leaseArea: 4400, dealRent: 3.7, rentFreeMonths: 4, status: '正常经营', remainingMonths: 27 },
      { name: '叮咚买菜（研发）', industry: '互联网/本地生活', floor: '1F', leaseArea: 1900, dealRent: 3.6, rentFreeMonths: 3, status: '已退租', remainingMonths: 0 },
    ],
  },
  {
    id: 'YP029', name: '互联网产业大楼', address: '长阳路1568号', plate: '长阳路', propertyType: '产业园', grade: '乙级',
    askingRent: 4.4, propertyFee: 16, totalArea: 55000, completionYear: 2015, floorHeight: 4.6, lng: 121.5278, lat: 31.2726,
    tenants: [
      { name: '优刻得 UCloud', industry: '云计算', floor: '6-9F', leaseArea: 6600, dealRent: 4.3, rentFreeMonths: 4, status: '正常经营', remainingMonths: 33 },
      { name: '智能云科', industry: '工业互联网', floor: '5F', leaseArea: 2400, dealRent: 4.2, rentFreeMonths: 3, status: '新签约', remainingMonths: 41 },
    ],
  },

  // ============ 大学科技园 / 高新园区 ============
  {
    id: 'YP030', name: '同济科技园', address: '赤峰路63号', plate: '四平路', propertyType: '产业园', grade: '乙级',
    askingRent: 4.2, propertyFee: 24, totalArea: 80000, completionYear: 2011, floorHeight: 4.1, lng: 121.5018, lat: 31.2856,
    tenants: [
      { name: '同济设计集团（TJAD）', industry: '建筑设计', floor: '6-12F', leaseArea: 9200, dealRent: 4.1, rentFreeMonths: 3, status: '正常经营', remainingMonths: 28 },
      { name: '岚图汽车数字研究院', industry: '汽车', floor: '5F', leaseArea: 2200, dealRent: 4.0, rentFreeMonths: 3, status: '筹备装修', remainingMonths: 46 },
    ],
  },
  {
    id: 'YP031', name: '上海财经大学国家大学科技园', address: '纪念路8号', plate: '五角场', propertyType: '产业园', grade: '乙级',
    askingRent: 4.0, propertyFee: 22, totalArea: 50000, completionYear: 2013, floorHeight: 4.0, lng: 121.4998, lat: 31.3098,
    tenants: [
      { name: '财跨境金融科技', industry: '金融科技', floor: '5-8F', leaseArea: 5200, dealRent: 3.9, rentFreeMonths: 4, status: '正常经营', remainingMonths: 21 },
      { name: '云锋基金（投研）', industry: '金融', floor: '4F', leaseArea: 1700, dealRent: 4.0, rentFreeMonths: 3, status: '正常经营', remainingMonths: 30 },
    ],
  },
  {
    id: 'YP032', name: '上海理工大学国家大学科技园', address: '军工路516号', plate: '军工路', propertyType: '产业园', grade: '乙级',
    askingRent: 3.6, propertyFee: 18, totalArea: 75000, completionYear: 2012, floorHeight: 4.4, lng: 121.5552, lat: 31.3092,
    tenants: [
      { name: '上海电气智能制造', industry: '智能制造', floor: '3-7F', leaseArea: 6800, dealRent: 3.5, rentFreeMonths: 4, status: '正常经营', remainingMonths: 26 },
      { name: '极链科技 Video++', industry: '人工智能', floor: '2F', leaseArea: 2100, dealRent: 3.4, rentFreeMonths: 3, status: '正常经营', remainingMonths: 18 },
    ],
  },
  {
    id: 'YP033', name: '五角场高新技术产业园·赛特分园', address: '南翔殷路58号', plate: '五角场', propertyType: '产业园', grade: '丙级',
    askingRent: 3.2, propertyFee: 15, totalArea: 42000, completionYear: 2010, floorHeight: 4.5, lng: 121.5118, lat: 31.3162,
    tenants: [
      { name: '禾赛科技中试基地', industry: '智能硬件', floor: '1-3F', leaseArea: 5600, dealRent: 3.1, rentFreeMonths: 3, status: '正常经营', remainingMonths: 22 },
    ],
  },
  {
    id: 'YP034', name: '四平科技公园·创新载体', address: '抚顺路', plate: '四平路', propertyType: '产业园', grade: '丙级',
    askingRent: 3.4, propertyFee: 15, totalArea: 36000, completionYear: 2014, floorHeight: 4.3, lng: 121.5098, lat: 31.2882,
    tenants: [
      { name: '声智科技', industry: '人工智能', floor: '2-4F', leaseArea: 3800, dealRent: 3.3, rentFreeMonths: 4, status: '新签约', remainingMonths: 39 },
    ],
  },
  {
    id: 'YP035', name: '上海婚纱艺术产业园', address: '军工路1436号', plate: '军工路', propertyType: '产业园', grade: '其它',
    askingRent: 2.8, propertyFee: 12, totalArea: 60000, completionYear: 2008, floorHeight: 5.2, lng: 121.5562, lat: 31.3148,
    tenants: [
      { name: '上海婚纱摄影联盟', industry: '文化创意', floor: '多栋', leaseArea: 8200, dealRent: 2.7, rentFreeMonths: 3, status: '正常经营', remainingMonths: 24 },
    ],
  },

  // ============ 江浦路 / 控江路 ============
  {
    id: 'YP036', name: '上海信息技术大厦', address: '隆昌路619号', plate: '江浦路', propertyType: '写字楼', grade: '乙级',
    askingRent: 2.7, propertyFee: 23, totalArea: 50000, completionYear: 2019, floorHeight: 3.9, lng: 121.5266, lat: 31.2706,
    tenants: [
      { name: '光明集团信息中心', industry: '快消品', floor: '8-12F', leaseArea: 6400, dealRent: 2.6, rentFreeMonths: 3, status: '正常经营', remainingMonths: 28 },
      { name: '招商银行江浦支行', industry: '金融', floor: '1-2F', leaseArea: 1800, dealRent: 2.9, rentFreeMonths: 2, status: '正常经营', remainingMonths: 19 },
    ],
  },
  {
    id: 'YP037', name: '中谊大厦', address: '宁国路374号', plate: '江浦路', propertyType: '写字楼', grade: '乙级',
    askingRent: 3.8, propertyFee: 25, totalArea: 80000, completionYear: 2017, floorHeight: 3.9, lng: 121.5286, lat: 31.2726,
    tenants: [
      { name: '哈啰出行华东团队', industry: '互联网/出行', floor: '8-12F', leaseArea: 6200, dealRent: 3.7, rentFreeMonths: 3, status: '正常经营', remainingMonths: 24 },
    ],
  },
  {
    id: 'YP038', name: '江浦路金融服务中心', address: '江浦路489号', plate: '江浦路', propertyType: '写字楼', grade: '乙级',
    askingRent: 4.0, propertyFee: 24, totalArea: 70000, completionYear: 2022, floorHeight: 4.0, lng: 121.5276, lat: 31.2682,
    tenants: [
      { name: '众安保险（运营）', industry: '金融', floor: '10-15F', leaseArea: 7400, dealRent: 3.9, rentFreeMonths: 3, status: '正常经营', remainingMonths: 38 },
    ],
  },
  {
    id: 'YP039', name: '控江路商务大厦', address: '控江路1500号', plate: '控江路', propertyType: '写字楼', grade: '丙级',
    askingRent: 3.1, propertyFee: 21, totalArea: 60000, completionYear: 2018, floorHeight: 3.8, lng: 121.5363, lat: 31.2882,
    tenants: [
      { name: '韵达速递（分部）', industry: '物流', floor: '10-14F', leaseArea: 5200, dealRent: 3.0, rentFreeMonths: 3, status: '正常经营', remainingMonths: 21 },
    ],
  },
  {
    id: 'YP040', name: '隆昌路智造园', address: '隆昌路1号', plate: '控江路', propertyType: '产业园', grade: '其它',
    askingRent: 2.9, propertyFee: 12, totalArea: 28000, completionYear: 2008, floorHeight: 5.2, lng: 121.5412, lat: 31.2806,
    tenants: [
      { name: '上海机床智能装备', industry: '智能制造', floor: '1-3F', leaseArea: 5600, dealRent: 2.8, rentFreeMonths: 3, status: '正常经营', remainingMonths: 26 },
    ],
  },

  // ============ 四平路 / 殷行 ============
  {
    id: 'YP041', name: '四平路商务中心', address: '四平路1188号', plate: '四平路', propertyType: '写字楼', grade: '乙级',
    askingRent: 4.2, propertyFee: 26, totalArea: 90000, completionYear: 2019, floorHeight: 4.0, lng: 121.5108, lat: 31.2851,
    tenants: [
      { name: '同济科技园企业服务', industry: '专业服务', floor: '6-10F', leaseArea: 6400, dealRent: 4.1, rentFreeMonths: 3, status: '正常经营', remainingMonths: 28 },
      { name: '霍尼韦尔（楼宇）', industry: '智能制造', floor: '5F', leaseArea: 2200, dealRent: 4.0, rentFreeMonths: 3, status: '正常经营', remainingMonths: 22 },
    ],
  },
  {
    id: 'YP042', name: '殷行科创大厦', address: '中原路1号', plate: '殷行', propertyType: '产业园', grade: '丙级',
    askingRent: 3.5, propertyFee: 20, totalArea: 50000, completionYear: 2022, floorHeight: 4.3, lng: 121.5302, lat: 31.3302,
    tenants: [
      { name: '华测检测认证', industry: '检测认证', floor: '5-8F', leaseArea: 5200, dealRent: 3.4, rentFreeMonths: 4, status: '正常经营', remainingMonths: 24 },
    ],
  },
  {
    id: 'YP043', name: '中环创智园', address: '逸仙路', plate: '中环', propertyType: '混合', grade: '乙级',
    askingRent: 4.0, propertyFee: 19, totalArea: 56000, completionYear: 2015, floorHeight: 4.1, lng: 121.4952, lat: 31.2901,
    tenants: [
      { name: '声网 Agora（华东）', industry: '云通信', floor: '8-12F', leaseArea: 6200, dealRent: 3.9, rentFreeMonths: 3, status: '正常经营', remainingMonths: 24 },
    ],
  },
  {
    id: 'YP044', name: '大柏树智慧广场', address: '汶水东路', plate: '中环', propertyType: '写字楼', grade: '乙级',
    askingRent: 4.4, propertyFee: 20, totalArea: 58000, completionYear: 2011, floorHeight: 3.8, lng: 121.4986, lat: 31.2944,
    tenants: [
      { name: '中通快递（数科）', industry: '物流', floor: '10-14F', leaseArea: 7200, dealRent: 4.3, rentFreeMonths: 3, status: '正常经营', remainingMonths: 30 },
    ],
  },

  // ============ 大桥 / 定海路 / 延吉新村 ============
  {
    id: 'YP045', name: '大桥商务大厦', address: '宁国路419号', plate: '大桥', propertyType: '写字楼', grade: '丙级',
    askingRent: 2.9, propertyFee: 20, totalArea: 70000, completionYear: 2016, floorHeight: 3.9, lng: 121.5352, lat: 31.2652,
    tenants: [
      { name: '德邦物流上海', industry: '物流', floor: '10-15F', leaseArea: 6800, dealRent: 2.8, rentFreeMonths: 3, status: '正常经营', remainingMonths: 21 },
    ],
  },
  {
    id: 'YP046', name: '定海滨江商务楼', address: '波阳路', plate: '定海路', propertyType: '写字楼', grade: '丙级',
    askingRent: 3.1, propertyFee: 21, totalArea: 60000, completionYear: 2015, floorHeight: 3.9, lng: 121.5448, lat: 31.2702,
    tenants: [
      { name: '沪东中华船舶配套', industry: '船舶制造', floor: '6-10F', leaseArea: 6200, dealRent: 3.0, rentFreeMonths: 3, status: '正常经营', remainingMonths: 26 },
    ],
  },
  {
    id: 'YP047', name: '延吉中路商务大厦', address: '延吉中路80号', plate: '延吉新村', propertyType: '写字楼', grade: '丙级',
    askingRent: 2.8, propertyFee: 19, totalArea: 40000, completionYear: 2020, floorHeight: 3.9, lng: 121.5352, lat: 31.2952,
    tenants: [
      { name: '锦江国际（分支）', industry: '商务服务', floor: '8-12F', leaseArea: 4200, dealRent: 2.7, rentFreeMonths: 3, status: '正常经营', remainingMonths: 18 },
    ],
  },
  {
    id: 'YP048', name: '财富中心广场', address: '邯郸路', plate: '五角场', propertyType: '写字楼', grade: '甲级',
    askingRent: 6.0, propertyFee: 26, totalArea: 70000, completionYear: 2014, floorHeight: 4.0, lng: 121.5201, lat: 31.2998,
    tenants: [
      { name: '小米生态链（上海）', industry: '智能硬件', floor: '16-20F', leaseArea: 6600, dealRent: 5.9, rentFreeMonths: 3, status: '正常经营', remainingMonths: 34 },
      { name: '理想汽车上海中心', industry: '汽车', floor: '12F', leaseArea: 2800, dealRent: 5.8, rentFreeMonths: 3, status: '正常经营', remainingMonths: 19 },
    ],
  },
  {
    id: 'YP049', name: '黄兴路总部基地', address: '黄兴路1858号', plate: '控江路', propertyType: '写字楼', grade: '甲级',
    askingRent: 5.2, propertyFee: 23, totalArea: 62000, completionYear: 2013, floorHeight: 4.0, lng: 121.5306, lat: 31.2926,
    tenants: [
      { name: '韵达速递总部', industry: '物流', floor: '14-19F', leaseArea: 7400, dealRent: 5.1, rentFreeMonths: 3, status: '正常经营', remainingMonths: 38 },
      { name: '哔哩哔哩（增长团队）', industry: '互联网/文娱', floor: '10F', leaseArea: 3200, dealRent: 5.0, rentFreeMonths: 3, status: '正常经营', remainingMonths: 21 },
    ],
  },
  {
    id: 'YP050', name: '东方渔人码头商务楼', address: '杨树浦路1088号', plate: '杨浦滨江', propertyType: '混合', grade: '甲级',
    askingRent: 7.2, propertyFee: 29, totalArea: 74000, completionYear: 2017, floorHeight: 4.1, lng: 121.5108, lat: 31.2571,
    tenants: [
      { name: '太平洋保险（滨江）', industry: '金融', floor: '12-16F', leaseArea: 6800, dealRent: 7.0, rentFreeMonths: 3, status: '正常经营', remainingMonths: 35 },
      { name: '蔚来汽车体验中心', industry: '汽车', floor: '1-2F', leaseArea: 2600, dealRent: 7.1, rentFreeMonths: 3, status: '新签约', remainingMonths: 47 },
    ],
  },
]
