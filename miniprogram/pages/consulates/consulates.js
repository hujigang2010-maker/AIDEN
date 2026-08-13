const data = require("../../utils/data");

Page({
  data: {
    list: data.consulates,
    tip: "本页只展示机构与负责人姓名。点击姓名可复制公开检索关键词，请自行通过搜索引擎或馆方官网查找；平台不提供电话、传真、邮箱或微信。"
  },
  onCopy(e) {
    const item = e.currentTarget.dataset.item;
    const text = item.search || `${item.country}驻上海总领事 ${item.head}`;
    wx.setClipboardData({
      data: text,
      success: () => wx.showToast({ title: "已复制检索词", icon: "success" })
    });
  }
});
