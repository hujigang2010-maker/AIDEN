const data = require("../../utils/data");

Page({
  data: {
    platform: data.platform,
    quickEntries: data.quickEntries,
    feeds: data.feeds,
    upcoming: data.events.slice(0, 3)
  },
  onEntryTap(e) {
    const item = e.currentTarget.dataset.item;
    if (!item) return;
    if (item.tab) wx.switchTab({ url: item.path });
    else wx.navigateTo({ url: item.path });
  },
  onEventTap(e) {
    wx.navigateTo({
      url: `/pages/event-detail/event-detail?id=${e.currentTarget.dataset.id}`
    });
  },
  onFeedTap(e) {
    const ref = e.currentTarget.dataset.ref;
    if (ref === "consulates") {
      wx.navigateTo({ url: "/pages/consulates/consulates" });
      return;
    }
    if (ref) wx.navigateTo({ url: `/pages/event-detail/event-detail?id=${ref}` });
  },
  goAbout() {
    wx.navigateTo({ url: "/pages/about/about" });
  }
});
