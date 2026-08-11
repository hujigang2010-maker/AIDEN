const data = require("../../utils/data");

Page({
  data: {
    platform: data.platform,
    plans: data.memberPlans,
    orders: [],
    customAmount: "100"
  },
  onShow() {
    const app = getApp();
    this.setData({ orders: app.globalData.orders || [] });
  },
  onAmountInput(e) {
    this.setData({ customAmount: e.detail.value });
  },
  buyMember(e) {
    const plan = e.currentTarget.dataset.plan;
    if (plan.custom) {
      const amount = Number(this.data.customAmount);
      if (!amount || amount <= 0) {
        wx.showToast({ title: "请输入金额", icon: "none" });
        return;
      }
      wx.navigateTo({
        url: `/pages/pay/pay?type=support&title=${encodeURIComponent(plan.name)}&price=${amount}`
      });
      return;
    }
    wx.navigateTo({
      url: `/pages/pay/pay?type=member&title=${encodeURIComponent(plan.name)}&price=${plan.price}`
    });
  },
  goEvents() {
    wx.switchTab({ url: "/pages/events/events" });
  },
  goAbout() {
    wx.navigateTo({ url: "/pages/about/about" });
  },
  goConsulates() {
    wx.navigateTo({ url: "/pages/consulates/consulates" });
  },
  contact() {
    wx.showModal({
      title: "联系平台",
      content: `${data.platform.email}\n${data.platform.address}\n${data.platform.sponsorNote}`,
      showCancel: false
    });
  }
});
