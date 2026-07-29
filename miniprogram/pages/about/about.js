Page({
  data: {
    desc: "复旦大学住房政策研究中心致力于住房制度、房地产市场、城市更新与居住社会学等领域的学术研究和政策咨询，推动构建更加公平、可持续、有韧性的居住体系。",
    extra: "中心汇聚经济学、社会学、管理学、城市规划等多学科力量，持续开展基础研究、政策评估与公共传播。",
    expanded: false,
    directions: [
      { num: "01", name: "住房制度与住房保障" },
      { num: "02", name: "房地产市场与金融" },
      { num: "03", name: "城市更新与社区治理" }
    ]
  },

  toggleMore() {
    this.setData({ expanded: !this.data.expanded });
  },

  copyEmail() {
    wx.setClipboardData({
      data: "hprc@fudan.edu.cn",
      success() {
        wx.showToast({ title: "已复制到剪贴板", icon: "none" });
      }
    });
  },

  onShareAppMessage() {
    return {
      title: "复旦大学住房政策研究中心",
      path: "/pages/about/about"
    };
  }
});
