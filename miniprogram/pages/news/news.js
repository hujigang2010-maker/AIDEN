const data = require("../../utils/data");

Page({
  data: { list: data.newsList },
  open(e) {
    const news = data.getNewsById(e.currentTarget.dataset.id);
    if (!news) return;
    wx.showModal({
      title: news.title,
      content: `${news.date} · ${news.tag}\n\n${news.summary}\n\n来源：${news.source}\n可检索：${news.linkHint}`,
      confirmText: "复制检索词",
      success: (res) => {
        if (res.confirm) wx.setClipboardData({ data: news.linkHint });
      }
    });
  }
});
