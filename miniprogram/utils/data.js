/**
 * 复旦链接 · 开放协作平台 — 演示数据
 * 板块模式参照「高层管理教育 / 教育中心」：公开课程、定制培训、在线学习、活动报名、新闻、关于我们
 * 发起可追溯：复旦大学住房政策研究中心
 * 视觉：NYU Violet #57068C
 */

const platform = {
  name: "复旦链接",
  fullName: "复旦链接 · 开放协作平台",
  motto: "找到适合你的学习与协作项目",
  subtitle: "全面的、系统的、可持续的开放协作与研修体系",
  sponsorNote: "由复旦大学住房政策研究中心发起运营",
  vision: "以世界级的开放协作与研修服务，满足终身学习与产业链接诉求。",
  mission: "赋能研究者、管理者与组织，连接活动、课程、产业机会与国际对接。",
  values: ["开放", "专业", "好奇", "敏捷", "协作"],
  about:
    "板块结构学习教育中心模式：公开课程、定制培训、在线学习、活动报名、平台新闻与关于我们。先搭平台，再承载安居议题、人工智能、具身智能、展会与出海机会。",
  principles: ["公开课程", "定制培训", "活动报名", "终身学习"],
  address: "上海市杨浦区",
  email: "platform@fudan-link.example",
  consultTip: "课程与定制咨询请通过「我的-联系平台」留言（演示环境）"
};

/** EE 式首页入口：公开课程 / 定制 / 在线 / 活动 */
const quickEntries = [
  { id: "courses", title: "公开课程", desc: "个人发展 · 专题研修", path: "/pages/courses/courses", tab: true },
  { id: "customize", title: "定制培训", desc: "企业/政府量身方案", path: "/pages/customize/customize" },
  { id: "events", title: "活动报名", desc: "论坛 · 沙龙 · 产业日", path: "/pages/events/events", tab: true },
  { id: "online", title: "在线学习", desc: "回放 · 微课 · 资料", path: "/pages/courses/courses", tab: true, query: "tab=online" }
];

/** 公开课程（卡片字段对齐教育中心：开班时间、学制、课时、了解详情） */
const courses = [
  {
    id: "c1",
    category: "公开课程",
    title: "城市与安居公开研修 · 第 1 期",
    tagline: "面向干部与管理者的系统性城市/安居研修",
    start: "2026年10月",
    duration: "共 6 天 · 周末集中",
    hours: "6 天",
    price: 6800,
    deposit: 1000,
    summary: "覆盖住房保障、市场调控与城市更新，含案例研讨。",
    modules: ["住房制度与保障", "市场与金融", "城市更新与社区"]
  },
  {
    id: "c2",
    category: "公开课程",
    title: "人工智能产业观察研修 · 第 1 期",
    tagline: "大模型、智算与产业落地路径",
    start: "2026年9月",
    duration: "4 天 · 两周集中",
    hours: "4 天",
    price: 4800,
    deposit: 800,
    summary: "链接智算与应用场景，对标产业日协作资源。",
    modules: ["大模型基础", "智算与成本", "场景落地工作坊"]
  },
  {
    id: "c3",
    category: "公开课程",
    title: "具身智能与场景实训营",
    tagline: "机器人、外骨骼与实训场景",
    start: "2026年11月",
    duration: "3 天集中",
    hours: "3 天",
    price: 3980,
    deposit: 600,
    summary: "真机体验 + 产业路演旁听 + 场景合作意向。",
    modules: ["具身技术图谱", "场景实训", "合作对接"]
  },
  {
    id: "c4",
    category: "在线学习",
    title: "住房政策前沿微课包",
    tagline: "碎片化学习 · 可单课解锁",
    start: "随时开课",
    duration: "录播 · 永久回看（演示）",
    hours: "约 6 小时",
    price: 199,
    deposit: 0,
    summary: "精选论坛实录与政策解读微课。",
    modules: ["保障住房", "租赁市场", "城市更新"]
  },
  {
    id: "c5",
    category: "在线学习",
    title: "出海国别预习课（领事名录导读）",
    tagline: "结合驻沪领事检索名录做国别预习",
    start: "随时开课",
    duration: "录播",
    hours: "约 2 小时",
    price: 99,
    deposit: 0,
    summary: "不提供领事联系方式；教你如何公开检索与合规对接。",
    modules: ["名录使用", "国别预习", "合规提示"]
  }
];

const customizeSteps = [
  "初步沟通，明确组织特征与学习需求",
  "学习需求分析，形成标准定制 / 模块系列 / 深度定制草案",
  "多轮沟通优化方案",
  "确认方案并签订合作协议",
  "邀请师资，按需求深度开发",
  "项目实施与过程评估",
  "回访反馈培训成果"
];

const events = [
  {
    id: "e1",
    status: "报名中",
    type: "活动报名",
    title: "城市安居与高质量发展论坛",
    date: "2026.09.18",
    time: "13:30–17:00",
    place: "上海 · 杨浦",
    city: "上海",
    paid: true,
    priceFrom: 199,
    summary: "汇聚政策、学界与产业观点，讨论保障性住房与城市更新路径。",
    agenda: [
      { time: "13:30", item: "签到入场" },
      { time: "14:00", item: "主旨报告" },
      { time: "15:10", item: "圆桌讨论" },
      { time: "16:30", item: "交流合影" }
    ],
    tickets: [
      { id: "t1", name: "标准票", price: 199, desc: "含资料包与茶歇" },
      { id: "t2", name: "学生票", price: 99, desc: "凭有效学生证" },
      { id: "t3", name: "VIP 席", price: 680, desc: "前排 + 会后交流" }
    ]
  },
  {
    id: "e2",
    status: "报名中",
    type: "产业日",
    title: "人工智能 × 具身智能 产业对接日",
    date: "2026.09.12",
    time: "09:30–17:30",
    place: "上海 · 创客场景",
    city: "上海",
    paid: true,
    priceFrom: 128,
    summary: "链接大模型、智算与具身智能企业，设置路演、体验与对接席。",
    agenda: [
      { time: "09:30", item: "开场与机会墙导览" },
      { time: "10:30", item: "企业路演" },
      { time: "14:00", item: "具身体验层" },
      { time: "16:00", item: "一对一对接" }
    ],
    tickets: [
      { id: "p1", name: "参会票", price: 128, desc: "全天议程 + 体验层" },
      { id: "p2", name: "对接席", price: 980, desc: "含闭门对接时段" }
    ]
  },
  {
    id: "e3",
    status: "预约中",
    type: "出海专题",
    title: "出海开放协作 · 国别商务沙龙",
    date: "2026.09.28",
    time: "14:00–18:30",
    place: "上海",
    city: "上海",
    paid: false,
    priceFrom: 0,
    summary: "面向有真实出海需求的实体企业；可结合驻沪领事公开检索名录做国别预习。",
    agenda: [
      { time: "14:00", item: "开场" },
      { time: "14:30", item: "国别营商观察" },
      { time: "16:30", item: "对接交流" }
    ],
    tickets: [{ id: "t0", name: "审核制免费票", price: 0, desc: "报名后审核确认" }]
  }
];

const archives = [
  {
    id: "a1",
    type: "论坛回顾",
    title: "城市政策圆桌纪要",
    date: "往期",
    summary: "围绕城市住房与更新路径的讨论摘要。",
    highlights: ["保障与市场需协同", "青年住房是关键切口"]
  },
  {
    id: "a2",
    type: "产业回顾",
    title: "具身智能体验日精彩瞬间",
    date: "往期",
    summary: "真机互动、路演与对接现场回顾。",
    highlights: ["体验层", "企业路演", "对接席"]
  }
];

const opportunities = [
  {
    id: "o1",
    track: "人工智能",
    title: "大模型与智算协作位",
    orgs: "优刻得、智谱 AI、道客网络等（示例）",
    action: "圆桌发言 / 黑客松导师 / 算力协作意向",
    paid: false
  },
  {
    id: "o2",
    track: "具身智能",
    title: "人形机器人与外骨骼场景",
    orgs: "苏度科技、卓益得、傲鲨智能、清宝引擎等（示例）",
    action: "真机体验 / 路演 / 实训场景合作",
    paid: false
  },
  {
    id: "o3",
    track: "ChinaJoy",
    title: "ChinaJoy / 数字娱乐生态对接",
    orgs: "ChinaJoy、CDEC、AI 未来生态大会伙伴（示例）",
    action: "展会专区、招商名录权限、赞助位",
    paid: true
  },
  {
    id: "o4",
    track: "国际对接",
    title: "驻沪领事检索名录",
    orgs: "70+ 驻沪领事机构",
    action: "查看负责人姓名与公开检索提示（无电话邮箱）",
    paid: false,
    link: "/pages/consulates/consulates"
  }
];

const exhibits = [
  {
    id: "x1",
    title: "开放协作主题展",
    subtitle: "线上展览",
    summary: "以时间线呈现：平台如何连接课程、活动、产业机会与国际对接。",
    sections: [
      { title: "公开课程", body: "对标教育中心公开课卡片：开班、学制、了解详情、报名。" },
      { title: "定制培训", body: "七步定制流程，服务企业与政府组织能力发展。" },
      { title: "活动报名", body: "论坛、产业日与出海专题可持续运营。" },
      { title: "国际与产业", body: "领事姓名可检索；AI / 具身 / 展会机会可跟进。" }
    ]
  }
];

const feeds = [
  { type: "平台新闻", title: "住房政策研究中心揭牌成立", summary: "2009年8月29日于光华楼揭牌", ref: "news1" },
  { type: "推荐课程", title: "城市与安居公开研修 · 第 1 期", summary: "10 月开班 · 接受定金", ref: "course:c1" },
  { type: "活动报名", title: "城市安居与高质量发展论坛", summary: "9 月 18 日开放票种", ref: "e1" }
];

const newsList = [
  {
    id: "news1",
    date: "2009.08.29",
    tag: "揭牌成立",
    title: "复旦大学住房政策研究中心正式揭牌",
    summary:
      "中心于光华楼举行揭牌仪式。全国政协常委、中国房地产研究会会长、前建设部副部长刘志峰与时任复旦大学党委书记秦绍德共同揭牌。中心原则：学术性、中立性、公益性。",
    source: "公开报道",
    linkHint: "检索：复旦大学住房政策研究中心 揭牌"
  },
  {
    id: "news2",
    date: "2009.12",
    tag: "专题论坛",
    title: "第二期住房政策专题论坛：上海住房政策未来展望与改进分析",
    summary: "汇聚学界、政策部门与行业专家，讨论政策框架完善及上海“十二五”住房政策建议稿。",
    source: "中心公开活动报道",
    linkHint: "检索：复旦大学住房政策研究中心 第二期论坛"
  },
  {
    id: "news3",
    date: "2010–2011",
    tag: "研究成果",
    title: "发布「复旦-同策上海均质住房价格指数」试行成果",
    summary: "基于上海新建市场化商品住宅全样本数据测算均质价格指数，服务公共信息与政策评估。",
    source: "公开发布会报道",
    linkHint: "检索：复旦-同策 上海均质住房价格指数"
  },
  {
    id: "news4",
    date: "2011.08",
    tag: "国际研讨",
    title: "首届「公共住房的未来」国际研讨会（香港）",
    summary: "十余个国家与地区专家出席，讨论上海公共住房、公积金改革与国际公屋经验。",
    source: "公开会议报道",
    linkHint: "检索：公共住房的未来 国际研讨会 复旦"
  },
  {
    id: "news5",
    date: "2011–2012",
    tag: "国际交流",
    title: "印度住房部代表团来访交流公共住房经验",
    summary: "就中国公共住房框架、上海公租房实践、融资与管理等问题深入讨论。",
    source: "公开交流报道",
    linkHint: "检索：复旦大学住房政策研究中心 印度代表团"
  },
  {
    id: "news6",
    date: "2011.08",
    tag: "政策研讨",
    title: "保障房建设动机与国际教训专题讨论",
    summary: "讨论公共住房民生、预防与发展三重作用，强调可持续运营与避免低收入聚居。",
    source: "公开研讨报道",
    linkHint: "检索：复旦大学 公共住房的未来 保障房"
  }
];

const memberPlans = [
  { id: "m1", name: "复旦链接 · 年度会员", price: 365, desc: "优先报名公开课与活动、资料包、部分回放" },
  { id: "m2", name: "支持平台 · 任意金额", price: 0, custom: true, desc: "支持公共活动与主题展览" }
];

const consulates = [
  { country: "美国", type: "总领事馆", head: "王汉", search: "美国驻上海总领事 王汉", area: "徐汇区淮海中路一带" },
  { country: "日本", type: "总领事馆", head: "冈田胜", search: "日本驻上海总领事 冈田胜", area: "长宁区万山路一带" },
  { country: "德国", type: "总领事馆", head: "李德仁", search: "德国驻上海总领事 李德仁", area: "徐汇区永福路一带" },
  { country: "新加坡", type: "总领事馆", head: "陈子勤", search: "新加坡驻上海总领事 陈子勤", area: "长宁区万山路一带" },
  { country: "韩国", type: "总领事馆", head: "金英俊", search: "韩国驻上海总领事 金英俊", area: "长宁区万山路一带" },
  { country: "泰国", type: "总领事馆", head: "陈佩恩", search: "泰国驻上海总领事 陈佩恩", area: "长宁区万山路一带" },
  { country: "马来西亚", type: "总领事馆", head: "沙哈菲兹·沙哈里斯", search: "马来西亚驻上海总领事", area: "长宁区红宝石路一带" },
  { country: "越南", type: "总领事馆", head: "陈辉雄", search: "越南驻上海总领事 陈辉雄", area: "浦东新区浦东大道一带" },
  { country: "印尼", type: "总领事馆", head: "邯伯盼", search: "印度尼西亚驻上海总领事", area: "长宁区延安西路一带" },
  { country: "菲律宾", type: "总领事馆", head: "方明易", search: "菲律宾驻上海总领事 方明易", area: "长宁区延安西路一带" },
  { country: "斐济", type: "总领事馆", head: "陈玉茹", search: "斐济驻上海总领事 陈玉茹", area: "闵行区古北路一带", note: "驻沪领团团长" },
  { country: "英国", type: "总领事馆", head: "（请公开检索现任馆长）", search: "英国驻上海总领事", area: "静安区北京西路一带" },
  { country: "法国", type: "总领事馆", head: "（请公开检索现任馆长）", search: "法国驻上海总领事", area: "长宁区中山西路一带" },
  { country: "澳大利亚", type: "总领事馆", head: "（请公开检索现任馆长）", search: "澳大利亚驻上海总领事", area: "静安区南京西路一带" }
];

function getEventById(id) {
  return events.find((e) => e.id === id) || null;
}

function getCourseById(id) {
  return courses.find((c) => c.id === id) || null;
}

function getNewsById(id) {
  return newsList.find((n) => n.id === id) || null;
}

module.exports = {
  platform,
  center: platform,
  quickEntries,
  courses,
  customizeSteps,
  events,
  archives,
  opportunities,
  exhibits,
  feeds,
  newsList,
  memberPlans,
  consulates,
  getEventById,
  getCourseById,
  getNewsById
};
