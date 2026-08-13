const data = require("../../utils/data");

Page({
  data: {
    list: data.orgs,
    tip: "对标管院校友中心「地方联络处 / 俱乐部 / 同学会」；此处为开放协作组织演示。"
  },
  onJoin(e) {
    const item = e.currentTarget.dataset.item;
    wx.showModal({
      title: "申请加入（演示）",
      content: `已提交「${item.title}」加入意向。正式环境将对接审核与社群引导。`,
      showCancel: false
    });
  }
});
