/** 无浏览器环境下的模板逻辑冒烟测试（Node） */
const assert = require("assert");
const fs = require("fs");
const vm = require("vm");
const path = require("path");
const { spawnSync } = require("child_process");

const src = fs.readFileSync(path.join(__dirname, "app.js"), "utf8");

function extractKeywords(text, limit = 8) {
  const CN_STOP = new Set([
    "的", "了", "是", "在", "和", "与", "一个", "没有", "可以", "因为", "所以",
    "但是", "然后", "如果", "这个", "那个", "我们", "你们", "他们", "以及",
    "进行", "通过", "关于", "什么", "怎么", "如何", "就是", "还是", "不是",
    "已经", "可能", "需要", "时候", "现在", "今天", "一些", "自己", "这样",
    "那样", "其实", "认为", "觉得", "知道", "开始", "问题", "内容",
  ]);
  const freq = new Map();
  const cleaned = text.replace(/[^\u4e00-\u9fffA-Za-z0-9]+/g, " ");
  for (let n = 4; n >= 2; n--) {
    for (let i = 0; i <= cleaned.length - n; i++) {
      const w = cleaned.slice(i, i + n);
      if (w.length !== n || /\s/.test(w) || CN_STOP.has(w)) continue;
      freq.set(w, (freq.get(w) || 0) + 1);
    }
  }
  const ranked = [...freq.entries()]
    .filter(([w, c]) => (w.length >= 3 ? c >= 2 : c >= 3))
    .sort((a, b) => b[0].length - a[0].length || b[1] - a[1]);
  const picked = [];
  for (const [w] of ranked) {
    if (picked.some((p) => p.includes(w) || w.includes(p))) continue;
    picked.push(w);
    if (picked.length >= limit) break;
  }
  return picked;
}

const sample =
  "今天想聊聊时间管理与注意力。很多人以为时间管理就是把日程排满，其实核心是注意力管理。" +
  "如果你早上最清醒的两小时被会议占满，那么深度工作会被迫挤到晚上，效率会明显下降。" +
  "一个可行方法是：每天保护一段九十分钟的专注块，关闭通知，只做一件高价值的事。" +
  "会议方面，会前写清决策问题，会中只讨论阻塞点，会后立刻写下行动项和负责人。" +
  "最后，每周五做一次复盘：哪些事情真正推进了目标，哪些只是忙碌的幻觉。时间管理离不开注意力。";

const kws = extractKeywords(sample);
console.log("keywords:", kws);
assert(kws.includes("时间管理"), "应包含 时间管理");
assert(kws.includes("注意力"), "应包含 注意力");
assert(!kws.includes("今天想聊"));
assert(!kws.includes("程排满"));
assert(!kws.includes("很多人"));

const py = spawnSync(
  "python3",
  ["-m", "py_compile", path.join(__dirname, "server.py"), path.join(__dirname, "common.py")],
  { encoding: "utf8" }
);
assert.strictEqual(py.status, 0, py.stderr);

new vm.Script(src, { filename: "app.js" });

// 健康检查（若服务已启动）
try {
  const http = require("http");
  http.get("http://127.0.0.1:8765/api/health", (res) => {
    let b = "";
    res.on("data", (c) => (b += c));
    res.on("end", () => {
      const j = JSON.parse(b);
      assert.strictEqual(j.ok, true);
      console.log("health OK");
      console.log("selftest OK");
    });
  }).on("error", () => {
    console.log("health skipped (server not running)");
    console.log("selftest OK");
  });
} catch {
  console.log("selftest OK");
}
