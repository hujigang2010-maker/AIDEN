const { getResearchById, getPolicyById } = require("../../utils/data");

Page({
  data: {
    item: null
  },

  onLoad(query) {
    const { type, id } = query || {};
    let item = null;
    if (type === "policy") {
      item = getPolicyById(id);
    } else {
      item = getResearchById(id);
    }
    this.setData({ item });
    if (item && item.title) {
      wx.setNavigationBarTitle({ title: "详情" });
    }
  },

  onShareAppMessage() {
    const item = this.data.item;
    return {
      title: item ? item.title : "复旦住房政策研究中心",
      path: `/pages/detail/detail?type=${item && item.org ? "policy" : "research"}&id=${item ? item.id : ""}`
    };
  }
});
