const data = require("../../utils/data");

Page({
  data: {
    filters: ["全部", "主题论坛", "政策沙龙", "专题培训", "实地参访"],
    active: "全部",
    list: data.events
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
