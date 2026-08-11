const data = require("../../utils/data");

Page({
  data: {
    archives: data.archives,
    exhibit: data.exhibits[0]
  },
  openExhibit() {
    wx.navigateTo({ url: "/pages/exhibit/exhibit" });
  },
  openArchive(e) {
    const id = e.currentTarget.dataset.id;
    const item = data.getArchiveById(id);
    if (!item) return;
    wx.showModal({
      title: item.title,
      content: (item.highlights || []).join("；") || item.summary,
      showCancel: false
    });
  }
});
