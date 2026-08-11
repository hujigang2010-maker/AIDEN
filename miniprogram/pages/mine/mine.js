const data = require("../../utils/data");

Page({
  data: {
    platform: data.platform,
    plans: data.memberPlans,
    orders: [],
    customAmount: "100"
  },
  onShow() {
    this.setData({ orders: getApp().globalData.orders || [] });
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
  goCourses() {
    wx.switchTab({ url: "/pages/courses/courses" });
  },
  goCustomize() {
    wx.navigateTo({ url: "/pages/customize/customize" });
  },
  goAbout() {
    wx.navigateTo({ url: "/pages/about/about" });
  },
  goConsulates() {
    wx.navigateTo({ url: "/pages/consulates/consulates" });
  },
  goOpp() {
    wx.navigateTo({ url: "/pages/opportunity/opportunity" });
  },
  goNews() {
    wx.navigateTo({ url: "/pages/news/news" });
  },
  contact() {
    wx.showModal({
      title: "联系平台",
      content: `${data.platform.email}\n${data.platform.address}\n${data.platform.sponsorNote}`,
      showCancel: false
    });
  }
});
