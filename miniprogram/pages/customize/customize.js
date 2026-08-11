const data = require("../../utils/data");

Page({
  data: {
    steps: data.customizeSteps,
    form: { org: "", need: "", contact: "" }
  },
  onInput(e) {
    const key = e.currentTarget.dataset.key;
    this.setData({ [`form.${key}`]: e.detail.value });
  },
  submit() {
    const { org, need, contact } = this.data.form;
    if (!org || !need || !contact) {
      wx.showToast({ title: "请完整填写", icon: "none" });
      return;
    }
    wx.showModal({
      title: "意向已提交（演示）",
      content: "定制培训走线索→方案→协议，不在端内闭环大额支付。平台运营将与您联系。",
      showCancel: false
    });
  }
});
