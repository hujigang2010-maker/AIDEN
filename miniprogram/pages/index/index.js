const data = require("../../utils/data");

Page({
  data: {
    platform: data.platform,
    serviceGrid: data.serviceGrid,
    featured: data.courses.filter((c) => c.category === "公开课程").slice(0, 2),
    newsList: data.newsList.slice(0, 2),
    upcoming: data.events.slice(0, 3)
  },
  onEntryTap(e) {
    const item = e.currentTarget.dataset.item;
    if (!item) return;
    if (item.tab) wx.switchTab({ url: item.path });
    else wx.navigateTo({ url: item.path });
  },
  openCourse(e) {
    wx.navigateTo({
      url: `/pages/course-detail/course-detail?id=${e.currentTarget.dataset.id}`
    });
  },
  openEvent(e) {
    wx.navigateTo({
      url: `/pages/event-detail/event-detail?id=${e.currentTarget.dataset.id}`
    });
  },
  openNews(e) {
    const id = e.currentTarget.dataset.id;
    const news = data.getNewsById(id);
    if (!news) return;
    wx.showModal({
      title: news.title,
      content: `${news.date} · ${news.tag}\n\n${news.summary}\n\n可检索：${news.linkHint}`,
      confirmText: "复制检索词",
      success: (res) => {
        if (res.confirm) wx.setClipboardData({ data: news.linkHint });
      }
    });
  },
  goNews() {
    wx.navigateTo({ url: "/pages/news/news" });
  },
  goCourses() {
    wx.navigateTo({ url: "/pages/courses/courses" });
  },
  goEvents() {
    wx.switchTab({ url: "/pages/events/events" });
  },
  goServices() {
    wx.switchTab({ url: "/pages/services/services" });
  },
  goAbout() {
    wx.navigateTo({ url: "/pages/about/about" });
  }
});
