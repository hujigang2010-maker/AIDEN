const data = require("../../utils/data");

Page({
  data: {
    filters: ["全部", "公开课程", "在线学习"],
    active: "全部",
    list: data.courses
  },
  onShow() {
    // 支持从首页「在线学习」带 tab 参数的简易处理
  },
  onFilter(e) {
    const active = e.currentTarget.dataset.name;
    const list =
      active === "全部"
        ? data.courses
        : data.courses.filter((c) => c.category === active);
    this.setData({ active, list });
  },
  openCourse(e) {
    wx.navigateTo({
      url: `/pages/course-detail/course-detail?id=${e.currentTarget.dataset.id}`
    });
  },
  goCustomize() {
    wx.navigateTo({ url: "/pages/customize/customize" });
  }
});
