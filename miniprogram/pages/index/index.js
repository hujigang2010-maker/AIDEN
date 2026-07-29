const { feeds } = require("../../utils/data");

Page({
  data: {
    feeds
  },

  goResearch() {
    wx.switchTab({ url: "/pages/research/research" });
  },

  goPolicy() {
    wx.switchTab({ url: "/pages/policy/policy" });
  },

  goAbout() {
    wx.switchTab({ url: "/pages/about/about" });
  },

  onMapTap() {
    wx.showToast({ title: "研究地图即将上线", icon: "none" });
  },

  openFeed(e) {
    const ref = e.currentTarget.dataset.ref;
    wx.navigateTo({
      url: `/pages/detail/detail?type=research&id=${ref}`
    });
  },

  onShareAppMessage() {
    const app = getApp();
    return {
      title: app.globalData.shareTitle,
      path: app.globalData.sharePath
    };
  }
});
