const data = require("../../utils/data");

Page({
  data: {
    filters: ["全部", "开放协作专项", "产业日", "出海专题", "公开研修"],
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
