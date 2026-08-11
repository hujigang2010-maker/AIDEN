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
  onPay() {
    const event = this.data.event;
    if (!event) return;
    const ticket = event.tickets[this.data.ticketIndex];
    wx.navigateTo({
      url: `/pages/pay/pay?type=event&eventId=${event.id}&ticketId=${ticket.id}&title=${encodeURIComponent(event.title)}&ticketName=${encodeURIComponent(ticket.name)}&price=${ticket.price}`
    });
  }
});
