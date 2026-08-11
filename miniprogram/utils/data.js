/**
 * 复旦链接 · 开放协作平台 — 演示数据
 * 发起可追溯：复旦大学住房政策研究中心
 * 不含管理学院 / 校友中心等外部机构表述
 */

const platform = {
  name: "复旦链接",
  fullName: "复旦链接 · 开放协作平台",
  motto: "先搭平台，再链接活动、产业与世界",
  sponsorNote: "由复旦大学住房政策研究中心发起运营",
  about:
    "本平台优先完成「开放协作」底座建设：发布与回顾活动、展示主题成果、链接人工智能 / 具身智能 / 展会与企业机会，并提供驻沪领事检索名录（仅姓名与公开检索线索，不含直接联系方式）。住房与城市议题是平台垂直之一，但不束缚平台边界。",
  principles: ["开放协作", "可追溯", "可持续"],
  address: "上海市杨浦区",
  email: "platform@fudan-link.example"
};

const quickEntries = [
  { id: "events", title: "活动报名", desc: "论坛 · 沙龙 · 产业日", path: "/pages/events/events", tab: true },
  { id: "archive", title: "活动回顾", desc: "纪要 · 相册 · 实录", path: "/pages/events/events", tab: true },
  { id: "opp", title: "产业机会", desc: "AI · 具身 · 展会", path: "/pages/opportunity/opportunity", tab: true },
  { id: "pay", title: "收费入口", desc: "票务 · 会员 · 支持", path: "/pages/mine/mine", tab: true, highlight: true }
];

const events = [
  {
    id: "e1",
    status: "报名中",
    type: "开放协作专项",
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
  },
  {
    id: "e4",
    status: "即将开班",
    type: "公开研修",
    title: "城市与产业公开研修 · 第 1 期",
    date: "2026.10.11",
    time: "周末集中 · 共 6 天",
    place: "上海",
    city: "上海",
    paid: true,
    priceFrom: 6800,
    summary: "覆盖城市治理、产业机会与国际对接模块，面向管理者与从业者。",
    agenda: [
      { time: "模块一", item: "城市与安居议题" },
      { time: "模块二", item: "AI 与具身产业观察" },
      { time: "模块三", item: "开放与出海协作" }
    ],
    tickets: [
      { id: "c1", name: "全款学费", price: 6800, desc: "含教材与证书" },
      { id: "c2", name: "报名定金", price: 1000, desc: "锁定名额，可抵扣" }
    ]
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
    track: "科创载体",
    title: "创客岛与园区通道",
    orgs: "全球创客岛、云基地、科创园区（示例）",
    action: "入孵意向、场景打样、一事一议",
    paid: false
  },
  {
    id: "o5",
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
    summary: "以时间线呈现：平台如何连接活动、产业机会与国际对接。",
    sections: [
      { title: "缘起", body: "先搭平台，形成可追溯的开放协作底座。" },
      { title: "活动", body: "论坛、沙龙、产业日与出海专题持续沉淀。" },
      { title: "机会", body: "AI、具身智能、ChinaJoy 与企业通道可被发现与跟进。" },
      { title: "国际", body: "驻沪领事姓名可检索，便于预习与合规对接。" }
    ]
  }
];

const feeds = [
  { type: "中心新闻", title: "住房政策研究中心揭牌成立", summary: "2009年8月29日于光华楼揭牌，坚持学术性、中立性、公益性", ref: "news1" },
  { type: "即将开始", title: "城市安居与高质量发展论坛", summary: "9 月 18 日开放票种报名", ref: "e1" },
  { type: "产业机会", title: "AI × 具身智能产业对接日", summary: "路演、体验层与对接席", ref: "e2" },
  { type: "国际对接", title: "驻沪领事检索名录已上线", summary: "仅提供姓名与公开检索线索", ref: "consulates" }
];

/** 公开新闻精选（据官网与公开报道整理，供平台动态与回顾） */
const newsList = [
  {
    id: "news1",
    date: "2009.08.29",
    tag: "揭牌成立",
    title: "复旦大学住房政策研究中心正式揭牌",
    summary:
      "中心于光华楼举行揭牌仪式。全国政协常委、中国房地产研究会会长、前建设部副部长刘志峰与时任复旦大学党委书记秦绍德共同揭牌。中心原则：学术性、中立性、公益性；研究方向涵盖住房保障、公共住房政策、住房市场调控与住房金融等。",
    source: "复旦大学社会科学高等研究院 / 公开报道",
    linkHint: "检索：复旦大学住房政策研究中心 揭牌"
  },
  {
    id: "news2",
    date: "2009.12",
    tag: "专题论坛",
    title: "第二期住房政策专题论坛：上海住房政策未来展望与改进分析",
    summary:
      "住房政策研究中心（CHPS）主办专题论坛，汇聚学界、政策部门与行业专家，围绕房价波动原因、政策框架完善及保障与市场两端协同展开讨论，并就上海“十二五”住房政策建议稿征求意见。",
    source: "中心公开活动报道",
    linkHint: "检索：复旦大学住房政策研究中心 第二期论坛"
  },
  {
    id: "news3",
    date: "2010–2011",
    tag: "研究成果",
    title: "发布「复旦-同策上海均质住房价格指数」试行成果",
    summary:
      "正值中心成立两周年前后，联合课题组基于上海新建市场化商品住宅全样本数据，测算均质价格指数，讨论房价指数对购房、政策评估与产业决策的公共价值，并计划拓展二手房等系列指数。",
    source: "公开发布会报道",
    linkHint: "检索：复旦-同策 上海均质住房价格指数"
  },
  {
    id: "news4",
    date: "2011.08",
    tag: "国际研讨",
    title: "首届「公共住房的未来」国际研讨会（香港）",
    summary:
      "中心与香港城市大学公共与社会管理系联合发起，十余个国家与地区专家出席。议题涵盖上海公共住房与住房公积金改革、香港公屋经验及公共住房的社会融合与城市竞争力等。",
    source: "公开会议报道",
    linkHint: "检索：公共住房的未来 国际研讨会 复旦"
  },
  {
    id: "news5",
    date: "2011–2012",
    tag: "国际交流",
    title: "印度住房部代表团来访交流公共住房经验",
    summary:
      "印度住房与城市减贫部代表团访问交流。双方就中国公共住房框架、上海公租房实践、融资与管理等问题深入讨论，并探讨城市化住房矛盾下的互鉴可能。",
    source: "公开交流报道",
    linkHint: "检索：复旦大学住房政策研究中心 印度代表团"
  },
  {
    id: "news6",
    date: "2011.08",
    tag: "政策研讨",
    title: "「公共住房的未来」研讨：保障房建设动机与国际教训",
    summary:
      "结合千万套级保障房建设背景，研讨公共住房的民生、预防与发展三重作用；强调避免低收入聚居、加强可持续运营，以及政府作为协调者与制度设计者的角色。",
    source: "公开研讨报道",
    linkHint: "检索：复旦大学 公共住房的未来 保障房"
  }
];

const memberPlans = [
  { id: "m1", name: "复旦链接 · 年度会员", price: 365, desc: "优先报名、资料包、部分回放" },
  { id: "m2", name: "支持平台 · 任意金额", price: 0, custom: true, desc: "支持公共活动与主题展览" }
];

/** 驻沪领事馆精简演示集（完整表见 Excel）；不含联系方式 */
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
  { country: "柬埔寨", type: "总领事馆", head: "宁维斯", search: "柬埔寨驻上海总领事 宁维斯", area: "静安区天目中路一带" },
  { country: "老挝", type: "总领事馆", head: "维达万·桥本控", search: "老挝驻上海总领事", area: "静安区武定路一带" },
  { country: "斐济", type: "总领事馆", head: "陈玉茹", search: "斐济驻上海总领事 陈玉茹", area: "闵行区古北路一带", note: "驻沪领团团长" },
  { country: "英国", type: "总领事馆", head: "（请公开检索现任馆长）", search: "英国驻上海总领事", area: "静安区北京西路一带" },
  { country: "法国", type: "总领事馆", head: "（请公开检索现任馆长）", search: "法国驻上海总领事", area: "长宁区中山西路一带" },
  { country: "澳大利亚", type: "总领事馆", head: "（请公开检索现任馆长）", search: "澳大利亚驻上海总领事", area: "静安区南京西路一带" },
  { country: "加拿大", type: "总领事馆", head: "（请公开检索现任馆长）", search: "加拿大驻上海总领事", area: "静安区南京西路一带" },
  { country: "卡塔尔", type: "总领事馆", head: "（请公开检索现任馆长）", search: "卡塔尔驻上海总领事", area: "浦东新区世纪大道一带" }
];

function getEventById(id) {
  return events.find((e) => e.id === id) || null;
}

module.exports = {
  platform,
  center: platform,
  quickEntries,
  events,
  archives,
  opportunities,
  exhibits,
  feeds,
  newsList,
  memberPlans,
  consulates,
  getEventById,
  getNewsById(id) {
    return newsList.find((n) => n.id === id) || null;
  }
};
