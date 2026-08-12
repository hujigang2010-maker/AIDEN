const data = require("../../utils/data");

Page({
  data: {
    event: null,
    ticketIndex: 0
  },
  onLoad(query) {
    const event = data.getEventById(query.id || "e1");
    this.setData({ event, ticketIndex: 0 });
    if (event) {
      wx.setNavigationBarTitle({ title: event.type || "活动详情" });
    }
  },
  onSelectTicket(e) {
    this.setData({ ticketIndex: Number(e.currentTarget.dataset.index) });
  },
  onPoster() {
    const event = this.data.event;
    if (!event) return;
    wx.showModal({
      title: "海报 / 邀请函（演示）",
      content: `已生成「${event.title}」分享海报草稿（含活动二维码占位）。真实环境可保存相册或发朋友圈。`,
      confirmText: "复制文案",
      success: (res) => {
        if (res.confirm) {
          wx.setClipboardData({
            data: `【复旦链接】邀请你参加：${event.title}｜${event.date} ${event.time}｜${event.place}`
          });
        }
      }
    });
  },
  onShare() {
    const event = this.data.event;
    if (!event) return;
    wx.showModal({
      title: "裂变传播（演示）",
      content: "可分享给好友 / 朋友圈；召集官转发码用于渠道统计。品牌展示为复旦链接。",
      confirmText: "复制路径",
      success: (res) => {
        if (res.confirm) {
          wx.setClipboardData({
            data: `/pages/event-detail/event-detail?id=${event.id}&from=invite`
          });
        }
      }
    });
  },
  onTicket() {
    const event = this.data.event;
    if (!event) return;
    const ticket = event.tickets[this.data.ticketIndex];
    wx.showModal({
      title: "电子票 / 签到（演示）",
      content: `票种：${ticket.name}\n核销码：FL-${event.id.toUpperCase()}-${ticket.id.toUpperCase()}\n现场扫码验票入场。`,
      showCancel: false
    });
  },
  onLive() {
    wx.showActionSheet({
      itemList: ["签到墙上墙（演示）", "现场抽奖（演示）", "投票互动（演示）"],
      success: (res) => {
        const names = ["签到墙", "抽奖", "投票"];
        wx.showToast({
          title: `${names[res.tapIndex]}已打开（演示）`,
          icon: "none"
        });
      }
    });
  },
  onPay() {
    const event = this.data.event;
    if (!event) return;
    const ticket = event.tickets[this.data.ticketIndex];
    wx.navigateTo({
      url: `/pages/pay/pay?type=event&eventId=${event.id}&ticketId=${ticket.id}&title=${encodeURIComponent(event.title)}&ticketName=${encodeURIComponent(ticket.name)}&price=${ticket.price}`
    });
  }
});
