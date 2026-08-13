const data = require("../../utils/data");

Page({
  data: {
    groups: data.serviceGroups,
    benchmarks: data.benchmarks
  },
  onOpen(e) {
    const item = e.currentTarget.dataset.item;
    if (!item || !item.path) return;
    if (item.tab) wx.switchTab({ url: item.path });
    else wx.navigateTo({ url: item.path });
  }
});
