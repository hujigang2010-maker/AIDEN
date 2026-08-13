Page({
  data: {
    type: "event",
    title: "",
    ticketName: "",
    price: 0,
    paying: false
  },
  onLoad(query) {
    this.setData({
      type: query.type || "event",
      eventId: query.eventId || "",
      ticketId: query.ticketId || "",
      title: decodeURIComponent(query.title || "订单"),
      ticketName: decodeURIComponent(query.ticketName || "标准项目"),
      price: Number(query.price || 0)
    });
  },
  submit() {
    if (this.data.paying) return;
    this.setData({ paying: true });
    const price = this.data.price;
    const done = () => {
      const order = {
        id: `o_${Date.now()}`,
        type: this.data.type,
        title: this.data.title,
        ticketName: this.data.ticketName,
        price,
        status: price > 0 ? "已支付（演示）" : "已报名（演示）",
        createdAt: new Date().toLocaleString()
      };
      getApp().saveOrder(order);
      wx.showToast({
        title: price > 0 ? "支付成功（演示）" : "报名成功",
        icon: "success"
      });
      setTimeout(() => {
        wx.switchTab({ url: "/pages/mine/mine" });
      }, 600);
    };

    // 演示环境：模拟微信支付确认，不调用真实 wx.requestPayment
    wx.showModal({
      title: price > 0 ? "模拟微信支付" : "确认报名",
      content:
        price > 0
          ? `将支付 ¥${price}（演示，不会产生真实扣款）`
          : "确认提交免费报名？",
      success: (res) => {
        if (res.confirm) {
          setTimeout(done, 300);
        } else {
          this.setData({ paying: false });
        }
      }
    });
  }
});
