const data = require("../../utils/data");

Page({
  data: { course: null },
  onLoad(q) {
    const course = data.getCourseById(q.id || "c1");
    this.setData({ course });
    if (course) wx.setNavigationBarTitle({ title: course.category });
  },
  enrollFull() {
    const c = this.data.course;
    if (!c) return;
    wx.navigateTo({
      url: `/pages/pay/pay?type=course&title=${encodeURIComponent(c.title)}&ticketName=${encodeURIComponent("全款学费")}&price=${c.price}`
    });
  },
  enrollDeposit() {
    const c = this.data.course;
    if (!c || !c.deposit) {
      this.enrollFull();
      return;
    }
    wx.navigateTo({
      url: `/pages/pay/pay?type=course&title=${encodeURIComponent(c.title)}&ticketName=${encodeURIComponent("报名定金")}&price=${c.deposit}`
    });
  }
});
