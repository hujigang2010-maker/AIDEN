const data = require("../../utils/data");

Page({
  data: {
    filters: ["全部", "论坛沙龙", "产业日", "出海专题", "联谊聚会"],
    active: "全部",
    list: data.events,
    tip: "活动主轴叠加互动吧：多票种售票、海报邀请函、电子票签到、现场互动。"
  },
  onFilter(e) {
    const active = e.currentTarget.dataset.name;
    const list =
      active === "全部"
        ? data.events
        : data.events.filter((item) => item.type === active);
    this.setData({ active, list });
  },
  onOpen(e) {
    wx.navigateTo({
      url: `/pages/event-detail/event-detail?id=${e.currentTarget.dataset.id}`
    });
  }
});
