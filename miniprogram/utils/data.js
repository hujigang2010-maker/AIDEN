/**
 * 演示数据 —— 正式上线前替换为中心确认内容与后台接口
 * 定位：活动平台（未来活动 / 回顾 / 展示）+ 收费入口
 */

const center = {
  name: "复旦大学住房政策研究中心",
  en: "Fudan Center for Housing Policy Studies",
  motto: "以活动连接学术与实践",
  principles: ["学术性", "中立性", "公益性"],
  about:
    "中心挂靠复旦大学管理学院，致力于住房保障、公共住房政策、住房市场调控与住房金融等领域的跨学科研究，并通过论坛、研修、展览与公共服务，服务政策制定与社会关切。",
  address: "上海市杨浦区国顺路 670 号",
  email: "hprc@fudan.edu.cn",
  directions: ["住房保障与公共住房", "住房市场调控与金融", "城市更新与居住公平", "国际住房政策比较"]
};

const quickEntries = [
  { id: "events", title: "活动报名", desc: "论坛 · 沙龙 · 参访", path: "/pages/events/events", tab: true },
  { id: "archive", title: "活动回顾", desc: "纪要 · 相册 · 实录", path: "/pages/archive/archive", tab: true },
  { id: "exhibit", title: "主题展览", desc: "线上策展展示", path: "/pages/exhibit/exhibit" },
  { id: "pay", title: "收费入口", desc: "票务 · 培训 · 支持", path: "/pages/mine/mine", tab: true, highlight: true }
];

const events = [
  {
    id: "e1",
    status: "报名中",
    type: "主题论坛",
    title: "保障性住房高质量发展路径论坛",
    date: "2026.09.18",
    time: "13:30–17:00",
    place: "复旦大学李达三楼",
    city: "上海",
    coverTone: "teal",
    paid: true,
    priceFrom: 199,
    summary: "汇聚学界、主管部门与行业专家，讨论保障房供给、融资与社区治理。",
    agenda: [
      { time: "13:30", item: "签到入场" },
      { time: "14:00", item: "主旨报告：保障性住房的下一程" },
      { time: "15:10", item: "圆桌：融资机制与长期运营" },
      { time: "16:30", item: "交流与合影" }
    ],
    tickets: [
      { id: "t1", name: "标准票", price: 199, desc: "含资料包与茶歇" },
      { id: "t2", name: "学生票", price: 99, desc: "凭有效学生证入场" },
      { id: "t3", name: "VIP 席", price: 680, desc: "前排席位 + 会后交流" }
    ]
  },
  {
    id: "e2",
    status: "报名中",
    type: "政策沙龙",
    title: "租赁住房条例落地观察沙龙",
    date: "2026.08.28",
    time: "19:00–21:00",
    place: "线下 · 杨浦",
    city: "上海",
    coverTone: "gold",
    paid: false,
    priceFrom: 0,
    summary: "小型闭门沙龙，围绕租赁市场规范与新市民住房问题展开讨论。",
    agenda: [
      { time: "19:00", item: "开场与议题导入" },
      { time: "19:20", item: "嘉宾分享" },
      { time: "20:20", item: "自由讨论" }
    ],
    tickets: [{ id: "t0", name: "免费票", price: 0, desc: "名额有限，审核后确认" }]
  },
  {
    id: "e3",
    status: "即将开班",
    type: "专题培训",
    title: "住房政策公开课 · 第 1 期（对标 EE 公开课）",
    date: "2026.10.11",
    time: "周末集中 · 共 6 天",
    place: "复旦大学管理学院",
    city: "上海",
    coverTone: "primary",
    paid: true,
    priceFrom: 6800,
    summary: "面向政府干部与行业管理者的住房政策系统研修，含案例研讨与参访。",
    agenda: [
      { time: "模块一", item: "住房制度与保障体系" },
      { time: "模块二", item: "市场调控与住房金融" },
      { time: "模块三", item: "城市更新与社区治理" }
    ],
    tickets: [
      { id: "c1", name: "全款学费", price: 6800, desc: "含教材、证书与同学社群" },
      { id: "c2", name: "报名定金", price: 1000, desc: "锁定名额，可抵扣学费" }
    ]
  },
  {
    id: "e4",
    status: "预约中",
    type: "实地参访",
    title: "保障性租赁住房社区参访",
    date: "2026.09.05",
    time: "09:00–12:00",
    place: "上海市某保租房社区",
    city: "上海",
    coverTone: "teal",
    paid: true,
    priceFrom: 128,
    summary: "走进运营现场，理解筹集渠道、运营管理与居民服务。",
    agenda: [
      { time: "09:00", item: "集合出发" },
      { time: "09:40", item: "项目介绍与参观" },
      { time: "11:00", item: "座谈交流" }
    ],
    tickets: [{ id: "v1", name: "参访票", price: 128, desc: "含交通统筹与讲解" }]
  }
];

const archives = [
  {
    id: "a1",
    type: "论坛回顾",
    title: "上海住房政策未来展望与改进分析",
    date: "往期论坛",
    summary: "学界与业界围绕上海住房政策框架、保障与市场两端展开讨论。",
    highlights: ["政策框架需兼顾市场与保障", "金融与土地工具需协同", "青年与新市民住房是关键切口"]
  },
  {
    id: "a2",
    type: "活动相册",
    title: "中心揭牌仪式影像回顾",
    date: "成立纪念",
    summary: "记录中心揭牌与各界嘉宾莅临，见证住房政策研究机构的启程。",
    highlights: ["揭牌时刻", "嘉宾致辞", "合影留念"]
  },
  {
    id: "a3",
    type: "实录摘要",
    title: "城市更新中的居住公平圆桌纪要",
    date: "沙龙实录",
    summary: "从空间、社会与政策三维理解城市更新中的居住公平。",
    highlights: ["避免大拆大建", "公众参与机制", "存量提质路径"]
  }
];

const exhibits = [
  {
    id: "x1",
    title: "安居有方 · 住房保障主题展",
    subtitle: "线上展览",
    summary: "以时间线与案例墙，呈现中国住房保障制度演进与地方实践。",
    sections: [
      { title: "缘起", body: "从住房民生关切出发，理解保障体系为何成为公共政策核心议题。" },
      { title: "制度演进", body: "经适房、公租房、保租房到配售型保障房，供给工具持续迭代。" },
      { title: "地方实践", body: "选取典型城市样本，展示筹集、分配与运营的多样路径。" },
      { title: "研究回应", body: "中心成果如何为政策评估与优化提供证据与方案。" }
    ]
  }
];

const feeds = [
  { type: "即将开始", title: "保障性住房高质量发展路径论坛", summary: "9 月 18 日 · 开放三类票种报名", ref: "e1" },
  { type: "培训招生", title: "住房政策公开课第 1 期开始接受定金", summary: "对标管院 EE 公开课模式，面向干部与管理者", ref: "e3" },
  { type: "展览上新", title: "安居有方 · 住房保障主题展上线", summary: "在小程序内沉浸浏览制度演进与地方实践", ref: "exhibit" }
];

const memberPlans = [
  { id: "m1", name: "中心之友 · 年度会员", price: 365, desc: "优先报名、活动资料包、部分回放畅看" },
  { id: "m2", name: "支持中心 · 任意金额", price: 0, custom: true, desc: "支持论坛、青年学者与公共展览" }
];

function getEventById(id) {
  return events.find((e) => e.id === id) || null;
}

function getArchiveById(id) {
  return archives.find((a) => a.id === id) || null;
}

module.exports = {
  center,
  quickEntries,
  events,
  archives,
  exhibits,
  feeds,
  memberPlans,
  getEventById,
  getArchiveById
};
