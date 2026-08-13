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
  updateProfile() {
    wx.showModal({
      title: "更新资料（演示）",
      content: "对标校友中心「更新我的信息」：姓名、单位、城市、兴趣标签。正式环境对接登录与审核。",
      showCancel: false
    });
  },
  goOrgs() {
    wx.navigateTo({ url: "/pages/orgs/orgs" });
  },
  goServices() {
    wx.switchTab({ url: "/pages/services/services" });
  },
  goCourses() {
    wx.navigateTo({ url: "/pages/courses/courses" });
  },
  goEvents() {
    wx.switchTab({ url: "/pages/events/events" });
  },
  goFeatures() {
    wx.navigateTo({ url: "/pages/features/features" });
  },
  goStories() {
    wx.navigateTo({ url: "/pages/stories/stories" });
  },
  showTicket(e) {
    const item = e.currentTarget.dataset.item;
    if (!item) return;
    wx.showModal({
      title: "电子票（演示）",
      content: `${item.title}\n${item.ticketName}\n核销码：FL-${(item.id || "DEMO").toString().slice(-6).toUpperCase()}\n到场前将推送提醒（规划）。`,
      showCancel: false
    });
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
