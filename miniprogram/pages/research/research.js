const { researchList } = require("../../utils/data");

Page({
  data: {
    chips: ["全部", "研究报告", "学术论文", "中心活动"],
    filter: "全部",
    list: researchList
  },

  onFilter(e) {
    const filter = e.currentTarget.dataset.chip;
    const list =
      filter === "全部"
        ? researchList
        : researchList.filter((item) => item.type === filter);
    this.setData({ filter, list });
  },

  openDetail(e) {
    const id = e.currentTarget.dataset.id;
    wx.navigateTo({
      url: `/pages/detail/detail?type=research&id=${id}`
    });
  },

  onShareAppMessage() {
    return {
      title: "复旦住房政策研究成果",
      path: "/pages/research/research"
    };
  }
});
