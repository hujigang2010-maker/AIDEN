App({
  globalData: {
    brand: "复旦链接 · 开放协作平台",
    motto: "先搭平台，再链接活动、产业与世界",
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
