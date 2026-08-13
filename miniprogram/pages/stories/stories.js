const data = require("../../utils/data");

Page({
  data: {
    list: data.stories,
    tip: "对标管院校友中心「校友故事 / 喜讯 / 观点」。"
  },
  onOpen(e) {
    const item = e.currentTarget.dataset.item;
    wx.showModal({
      title: item.title,
      content: `${item.date} · ${item.kind}\n\n${item.summary}`,
      showCancel: false
    });
  }
});
