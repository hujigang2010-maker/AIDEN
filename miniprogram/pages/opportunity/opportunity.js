const data = require("../../utils/data");

Page({
  data: {
    list: data.opportunities,
    archives: data.archives
  },
  openItem(e) {
    const item = e.currentTarget.dataset.item;
    if (item && item.link) {
      wx.navigateTo({ url: item.link });
      return;
    }
    wx.showModal({
      title: item.title,
      content: `${item.orgs}\n\n可链接动作：${item.action}\n\n该机会归属「复旦链接」开放协作专项，可追溯至平台发起主体。`,
      confirmText: item.paid ? "了解收费" : "知道了",
      success: (res) => {
        if (res.confirm && item.paid) {
          wx.switchTab({ url: "/pages/mine/mine" });
        }
      }
    });
  },
  openConsulates() {
    wx.navigateTo({ url: "/pages/consulates/consulates" });
  },
  openExhibit() {
    wx.navigateTo({ url: "/pages/exhibit/exhibit" });
  }
});
