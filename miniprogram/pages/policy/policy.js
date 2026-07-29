const { policyList } = require("../../utils/data");

Page({
  data: {
    chips: ["全部", "住房保障", "租赁住房", "城市更新", "房地产市场"],
    filter: "全部",
    keyword: "",
    list: policyList
  },

  applyFilter() {
    const { filter, keyword } = this.data;
    const q = (keyword || "").trim().toLowerCase();
    const list = policyList.filter((item) => {
      const catOk = filter === "全部" || item.category === filter;
      const qOk =
        !q ||
        item.title.toLowerCase().includes(q) ||
        item.org.toLowerCase().includes(q);
      return catOk && qOk;
    });
    this.setData({ list });
  },

  onFilter(e) {
    this.setData({ filter: e.currentTarget.dataset.chip }, () => this.applyFilter());
  },

  onSearch(e) {
    this.setData({ keyword: e.detail.value }, () => this.applyFilter());
  },

  openDetail(e) {
    const id = e.currentTarget.dataset.id;
    wx.navigateTo({
      url: `/pages/detail/detail?type=policy&id=${id}`
    });
  },

  onShareAppMessage() {
    return {
      title: "住房政策数据库 · 复旦 HPRC",
      path: "/pages/policy/policy"
    };
  }
});
