const data = require("../../utils/data");

Page({
  data: {
    features: data.activityFeatures,
    note: "主骨架对标复旦 eHall + 管院校友中心；活动能力学习互动吧全链路。对外品牌统一为「复旦链接」。"
  },
  goEvents() {
    wx.switchTab({ url: "/pages/events/events" });
  }
});
