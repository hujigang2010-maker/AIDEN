const data = require("../../utils/data");

Page({
  data: {
    center: data.center,
    quickEntries: data.quickEntries,
    feeds: data.feeds,
    upcoming: data.events.filter((e) => e.status !== "已结束").slice(0, 3)
  },
  onEntryTap(e) {
    const item = e.currentTarget.dataset.item;
    if (!item) return;
    if (item.tab) {
      wx.switchTab({ url: item.path });
    } else {
      wx.navigateTo({ url: item.path });
    }
  },
  onEventTap(e) {
    const id = e.currentTarget.dataset.id;
    wx.navigateTo({ url: `/pages/event-detail/event-detail?id=${id}` });
  },
  onFeedTap(e) {
    const ref = e.currentTarget.dataset.ref;
    if (ref === "exhibit") {
      wx.navigateTo({ url: "/pages/exhibit/exhibit" });
      return;
    }
    if (ref) {
      wx.navigateTo({ url: `/pages/event-detail/event-detail?id=${ref}` });
    }
  },
  goAbout() {
    wx.navigateTo({ url: "/pages/about/about" });
  }
});
