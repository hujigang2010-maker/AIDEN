const data = require("../../utils/data");

Page({
  data: {
    features: data.activityFeatures,
    note: "能力学习互动吧全链路；对外品牌统一为「复旦链接」，不挂互动吧名义。"
  },
  goEvents() {
    wx.switchTab({ url: "/pages/events/events" });
  }
});
