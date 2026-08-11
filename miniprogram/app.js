App({
  globalData: {
    brand: "复旦大学住房政策研究中心",
    motto: "以活动连接学术与实践",
    orders: []
  },
  onLaunch() {
    try {
      const orders = wx.getStorageSync("hprc_orders") || [];
      this.globalData.orders = orders;
    } catch (e) {
      this.globalData.orders = [];
    }
  },
  saveOrder(order) {
    const orders = this.globalData.orders || [];
    orders.unshift(order);
    this.globalData.orders = orders;
    try {
      wx.setStorageSync("hprc_orders", orders);
    } catch (e) {}
  }
});
