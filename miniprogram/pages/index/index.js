const data = require("../../utils/data");

Page({
  data: {
    platform: data.platform,
    quickEntries: data.quickEntries,
    feeds: data.feeds,
    newsList: data.newsList.slice(0, 4),
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
    if (ref && String(ref).startsWith("news")) {
      const news = data.getNewsById(ref);
      if (news) {
        wx.showModal({
          title: news.title,
          content: `${news.date} · ${news.tag}\n\n${news.summary}\n\n来源：${news.source}\n可检索：${news.linkHint}`,
          confirmText: "复制检索词",
          success: (res) => {
            if (res.confirm) {
              wx.setClipboardData({ data: news.linkHint });
            }
          }
        });
      }
      return;
    }
    if (ref) wx.navigateTo({ url: `/pages/event-detail/event-detail?id=${ref}` });
  },
  goAbout() {
    wx.navigateTo({ url: "/pages/about/about" });
  }
});
