const data = require("../../utils/data");
Page({
  data: { platform: data.platform },
  goFeatures() {
    wx.navigateTo({ url: "/pages/features/features" });
  }
});
