/* TicNote Studio — 本地知识库 + 转写模板 + TicNote API 同步 */
(() => {
  "use strict";

  const STORAGE = {
    kb: "tn_studio_kb_v1",
    settings: "tn_studio_settings_v1",
    session: "tn_studio_session_v1",
  };

  const AGENTS = [
    {
      id: "summarizer",
      name: "纪要官",
      desc: "提炼一句话摘要、核心要点与行动项，适合会后同步。",
      tpl: "summary",
    },
    {
      id: "mindmapper",
      name: "结构官",
      desc: "把长文拆成层级主题，生成可导出的思维导图。",
      tpl: "mindmap",
    },
    {
      id: "aha",
      name: "顿悟官",
      desc: "抓洞见、金句与待办，适合访谈与灵感录音。",
      tpl: "aha",
    },
    {
      id: "researcher",
      name: "研究员",
      desc: "主题、发现、概念与延伸问题；可对接 TicNote Deep Research。",
      tpl: "research",
      ticnote: true,
    },
    {
      id: "podcaster",
      name: "播客制片",
      desc: "分角色对话稿 + Show Notes，方便二次分发。",
      tpl: "podcast",
    },
    {
      id: "translator",
      name: "翻译官",
      desc: "整篇或多语对照翻译，服务跨境会议与网页剪藏。",
      tpl: "translate",
    },
  ];

  const CN_STOP = new Set([
    "的", "了", "是", "在", "和", "与", "及", "或", "而", "被", "把", "让", "向", "给",
    "一个", "没有", "可以", "因为", "所以", "但是", "然后", "如果", "这个", "那个",
    "我们", "你们", "他们", "以及", "进行", "通过", "关于", "什么", "怎么", "如何",
    "就是", "还是", "不是", "已经", "可能", "需要", "时候", "现在", "今天", "一些",
    "自己", "这样", "那样", "其实", "认为", "觉得", "知道", "开始", "问题", "内容",
    "他们", "她们", "它们", "其中", "以上", "如下", "例如", "比如", "等等",
  ]);

  const EN_STOP = new Set(
    "the a an and or but if in on at to for of as is are was were be been being this that these those it its we you they he she them our your their with from by into about over after before not no yes do does did done have has had can could should would will just so than then also more most other into such only own same too very".split(
      " "
    )
  );

  // ─── DOM ───
  const $ = (id) => document.getElementById(id);
  const els = {
    connStatus: $("connStatus"),
    mainNav: $("mainNav"),
    mainTitle: $("mainTitle"),
    timeline: $("timeline"),
    kbFilter: $("kbFilter"),
    transcript: $("transcript"),
    noteTitle: $("noteTitle"),
    output: $("output"),
    mindmapSvgWrap: $("mindmapSvgWrap"),
    tplBar: $("tplBar"),
    translateOpts: $("translateOpts"),
    translateLang: $("translateLang"),
    langSelect: $("langSelect"),
    btnMic: $("btnMic"),
    audioFile: $("audioFile"),
    audioMeta: $("audioMeta"),
    audioPlayer: $("audioPlayer"),
    btnAudio2Text: $("btnAudio2Text"),
    btnExportAll: $("btnExportAll"),
    btnSaveNote: $("btnSaveNote"),
    btnSyncTicnote: $("btnSyncTicnote"),
    btnSyncProject: $("btnSyncProject"),
    btnLoadSample: $("btnLoadSample"),
    btnRefreshKb: $("btnRefreshKb"),
    btnConnect: $("btnConnect"),
    btnDisconnect: $("btnDisconnect"),
    appkey: $("appkey"),
    projectSelect: $("projectSelect"),
    projectIdInput: $("projectIdInput"),
    projectFiles: $("projectFiles"),
    proxyBase: $("proxyBase"),
    webUrl: $("webUrl"),
    btnFetchWeb: $("btnFetchWeb"),
    chatBox: $("chatBox"),
    askInput: $("askInput"),
    btnAsk: $("btnAsk"),
    btnQuickTranslate: $("btnQuickTranslate"),
    agentList: $("agentList"),
    kbDetail: $("kbDetail"),
    toast: $("toast"),
    settingsModal: $("settingsModal"),
    aiEndpoint: $("aiEndpoint"),
    aiKey: $("aiKey"),
    aiModel: $("aiModel"),
  };

  const state = {
    view: "capture",
    tpl: "summary",
    kb: [],
    activeId: null,
    recognizing: false,
    recognition: null,
    audioBlob: null,
    audioName: "",
    session: { token: "", appkey: "", baseUrl: "", projects: [] },
    settings: { aiEndpoint: "", aiKey: "", aiModel: "gpt-4o-mini" },
    lastSvg: "",
  };

  // ─── utils ───
  function toast(msg) {
    els.toast.textContent = msg;
    els.toast.classList.add("show");
    clearTimeout(toast._t);
    toast._t = setTimeout(() => els.toast.classList.remove("show"), 2600);
  }

  function uid() {
    return "n_" + Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
  }

  function loadJSON(key, fallback) {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : fallback;
    } catch {
      return fallback;
    }
  }

  function saveJSON(key, val) {
    localStorage.setItem(key, JSON.stringify(val));
  }

  function proxy() {
    const v = (els.proxyBase.value || "").trim();
    if (v) return v.replace(/\/$/, "");
    if (location.protocol.startsWith("http")) return location.origin;
    return "http://127.0.0.1:8765";
  }

  async function api(path, body) {
    const res = await fetch(proxy() + path, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body || {}),
    });
    let data;
    try {
      data = await res.json();
    } catch {
      throw new Error("服务无响应，请先运行 python3 server.py");
    }
    if (!res.ok || data.ok === false) {
      const err = data.error;
      const msg =
        typeof err === "string"
          ? err
          : err?.msg || err?.message || JSON.stringify(err || data);
      throw new Error(msg || "请求失败");
    }
    return data;
  }

  function download(filename, content, type = "text/plain;charset=utf-8") {
    const blob = content instanceof Blob ? content : new Blob([content], { type });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }

  function todayKey(d = new Date()) {
    return d.toISOString().slice(0, 10);
  }

  function formatTime(iso) {
    if (!iso) return "";
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return String(iso);
    return d.toLocaleString("zh-CN", { hour12: false });
  }

  // ─── text analytics / templates ───
  function splitSentences(text) {
    return text
      .replace(/\s+/g, " ")
      .split(/(?<=[。！？!?；;\n])/)
      .map((s) => s.trim())
      .filter((s) => s.length >= 6);
  }

  function extractKeywords(text, limit = 8) {
    const isHans = /[\u4e00-\u9fff]/.test(text);
    const freq = new Map();
    if (isHans) {
      const cleaned = text.replace(/[^\u4e00-\u9fffA-Za-z0-9]+/g, " ");
      for (let n = 4; n >= 2; n--) {
        for (let i = 0; i <= cleaned.length - n; i++) {
          const w = cleaned.slice(i, i + n);
          if (w.length !== n || /\s/.test(w)) continue;
          if (CN_STOP.has(w)) continue;
          if (/^[A-Za-z0-9]+$/.test(w) && w.length < 3) continue;
          freq.set(w, (freq.get(w) || 0) + 1);
        }
      }
    } else {
      for (const w of text.toLowerCase().match(/[a-z][a-z\-']{2,}/g) || []) {
        if (EN_STOP.has(w)) continue;
        freq.set(w, (freq.get(w) || 0) + 1);
      }
    }
    // 先按长度再按频次，避免短词抢占「时间管理」这类长短语
    // 二字词提高阈值，减少「做一」「哪些」类噪声
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

  function buildMindmapTree(text) {
    const sents = splitSentences(text);
    const kws = extractKeywords(text, 6);
    const title = (els.noteTitle.value || "未命名主题").trim();
    const children = kws.map((k) => {
      const related = sents.filter((s) => s.includes(k)).slice(0, 3);
      return {
        name: k,
        children: (related.length ? related : sents.slice(0, 2)).map((s) => ({
          name: s.length > 36 ? s.slice(0, 36) + "…" : s,
        })),
      };
    });
    if (!children.length) {
      return {
        name: title,
        children: sents.slice(0, 5).map((s) => ({
          name: s.length > 40 ? s.slice(0, 40) + "…" : s,
        })),
      };
    }
    return { name: title, children };
  }

  function renderMindmapSvg(tree) {
    const W = 900;
    const H = 520;
    const cx = 140;
    const cy = H / 2;
    const branches = tree.children || [];
    const n = Math.max(branches.length, 1);
    let paths = "";
    let labels = "";
    branches.forEach((b, i) => {
      const t = n === 1 ? 0.5 : i / (n - 1);
      const ang = -Math.PI / 2.6 + t * (Math.PI * 1.15);
      const x1 = cx + Math.cos(ang) * 180;
      const y1 = cy + Math.sin(ang) * 170;
      paths += `<path d="M${cx},${cy} Q${cx + 90},${cy} ${x1},${y1}" fill="none" stroke="#0f766e" stroke-width="2.2" opacity=".75"/>`;
      labels += `<circle cx="${x1}" cy="${y1}" r="8" fill="#0f766e"/><text x="${x1 + 14}" y="${y1 + 5}" font-size="14" font-family="Sora,sans-serif" fill="#14201f">${escapeXml(b.name)}</text>`;
      (b.children || []).slice(0, 3).forEach((c, j) => {
        const x2 = x1 + 160;
        const y2 = y1 - 28 + j * 28;
        paths += `<path d="M${x1},${y1} C${x1 + 50},${y1} ${x2 - 40},${y2} ${x2},${y2}" fill="none" stroke="#c9842f" stroke-width="1.4" opacity=".7"/>`;
        labels += `<text x="${x2 + 6}" y="${y2 + 4}" font-size="12" font-family="Sora,sans-serif" fill="#5a6b69">${escapeXml(c.name)}</text>`;
      });
    });
    return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">
  <rect width="100%" height="100%" fill="#f7faf8"/>
  <circle cx="${cx}" cy="${cy}" r="46" fill="#0b4f4a"/>
  <text x="${cx}" y="${cy + 5}" text-anchor="middle" fill="#fff" font-size="13" font-family="Fraunces,serif">${escapeXml(tree.name.slice(0, 12))}</text>
  ${paths}${labels}
</svg>`;
  }

  function escapeXml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function tplSummary(text) {
    const sents = splitSentences(text);
    const kws = extractKeywords(text);
    const one = sents[0] || text.slice(0, 80);
    const points = sents.slice(0, 6).map((s, i) => `${i + 1}. ${s}`);
    return [
      "## 总结",
      "",
      "**一句话概括**",
      one,
      "",
      "**核心要点**",
      ...(points.length ? points : ["- （文本过短）"]),
      "",
      "**关键词**",
      kws.length ? kws.map((k) => `\`${k}\``).join(" · ") : "（暂无重复关键词）",
    ].join("\n");
  }

  function tplMindmap(text) {
    const tree = buildMindmapTree(text);
    state.lastSvg = renderMindmapSvg(tree);
    const lines = [`## 思维导图`, "", `- ${tree.name}`];
    for (const c of tree.children || []) {
      lines.push(`  - ${c.name}`);
      for (const g of c.children || []) lines.push(`    - ${g.name}`);
    }
    return lines.join("\n");
  }

  function tplAha(text) {
    const sents = splitSentences(text);
    const insights = sents.filter((s) => /因为|所以|关键|本质|其实|发现|原来|意味着|导致|如果/.test(s)).slice(0, 5);
    const quotes = [...sents].sort((a, b) => a.length - b.length).slice(0, 4);
    const actions = sents.filter((s) => /要|应该|需要|下一步|准备|计划|马上|记得|别忘|行动/.test(s)).slice(0, 5);
    return [
      "## 顿悟摘录",
      "",
      "### 洞见",
      ...(insights.length ? insights.map((s) => `- ${s}`) : ["- （未识别到明显洞见句，可补充原文）"]),
      "",
      "### 金句",
      ...(quotes.map((s) => `> ${s}`)),
      "",
      "### 行动项",
      ...(actions.length ? actions.map((s) => `- [ ] ${s}`) : ["- [ ] 复盘本次内容并列出 3 个下一步"]),
    ].join("\n");
  }

  function tplResearch(text) {
    const sents = splitSentences(text);
    const kws = extractKeywords(text, 8);
    const title = els.noteTitle.value.trim() || kws[0] || "研究主题";
    return [
      "## 深度研究报告",
      "",
      `**主题**：${title}`,
      "",
      "### 核心发现",
      ...sents.slice(0, 5).map((s) => `- ${s}`),
      "",
      "### 关键概念",
      ...(kws.length ? kws.map((k) => `- ${k}`) : ["- （待补充）"]),
      "",
      "### 引用线索",
      ...sents.slice(0, 3).map((s) => `- 「${s}」`),
      "",
      "### 延伸问题",
      `- ${title} 的边界条件是什么？`,
      `- 有哪些反例或风险？`,
      `- 下一步最值得验证的假设是什么？`,
    ].join("\n");
  }

  function tplPodcast(text) {
    const sents = splitSentences(text);
    const chunks = [];
    for (let i = 0; i < sents.length; i += 2) {
      chunks.push(sents.slice(i, i + 2));
    }
    const lines = [
      "## 播客节目稿",
      "",
      `**节目名**：${els.noteTitle.value.trim() || "TicNote 速记场"}`,
      "",
      "### 对话",
    ];
    let t = 0;
    chunks.slice(0, 10).forEach((pair, idx) => {
      const mm = String(Math.floor(t / 60)).padStart(2, "0");
      const ss = String(t % 60).padStart(2, "0");
      const host = idx % 2 === 0 ? "主持人" : "嘉宾";
      lines.push(`**[${mm}:${ss}] ${host}**：${pair.join(" ")}`);
      t += 18 + Math.min(40, pair.join(" ").length / 4);
    });
    lines.push("", "### Show Notes");
    extractKeywords(text, 6).forEach((k) => lines.push(`- ${k}`));
    return lines.join("\n");
  }

  const DICT = {
    en: {
      时间: "time",
      管理: "management",
      注意力: "attention",
      会议: "meeting",
      总结: "summary",
      行动: "action",
      问题: "question",
      今天: "today",
      我们: "we",
      需要: "need",
      可以: "can",
      重要: "important",
      研究: "research",
      播客: "podcast",
    },
    ja: { 时间: "時間", 管理: "管理", 注意力: "注意力", 会议: "会議", 总结: "まとめ" },
    ko: { 时间: "시간", 管理: "관리", 注意力: "집중력", 会议: "회의", 总结: "요약" },
    fr: { 时间: "temps", 管理: "gestion", 注意力: "attention", 会议: "réunion", 总结: "résumé" },
    de: { 时间: "Zeit", 管理: "Management", 注意力: "Aufmerksamkeit", 会议: "Meeting", 总结: "Zusammenfassung" },
    zh: {},
  };

  function tplTranslate(text, lang) {
    if (lang === "zh") {
      return ["## 翻译（中文）", "", text].join("\n");
    }
    const dict = DICT[lang] || DICT.en;
    let out = text;
    Object.keys(dict)
      .sort((a, b) => b.length - a.length)
      .forEach((k) => {
        out = out.split(k).join(dict[k]);
      });
    const note =
      state.settings.aiKey
        ? "（已配置 AI，可在设置后点模板再次生成以使用模型翻译）"
        : "（本地词典粗译；配置 AI 或使用 TicNote 翻译可获更高质量）";
    return [`## 翻译（${lang}）`, "", out, "", `_${note}_`].join("\n");
  }

  function applyTemplate() {
    const text = els.transcript.value.trim();
    if (!text) {
      els.output.textContent = "请先录入或同步一段文本。";
      els.mindmapSvgWrap.classList.add("hidden");
      return;
    }
    els.translateOpts.classList.toggle("hidden", state.tpl !== "translate");
    let md = "";
    if (state.tpl === "summary") md = tplSummary(text);
    else if (state.tpl === "mindmap") md = tplMindmap(text);
    else if (state.tpl === "aha") md = tplAha(text);
    else if (state.tpl === "research") md = tplResearch(text);
    else if (state.tpl === "podcast") md = tplPodcast(text);
    else if (state.tpl === "translate") md = tplTranslate(text, els.translateLang.value);

    els.output.textContent = md;
    if (state.tpl === "mindmap" && state.lastSvg) {
      els.mindmapSvgWrap.classList.remove("hidden");
      els.mindmapSvgWrap.innerHTML = state.lastSvg;
    } else {
      els.mindmapSvgWrap.classList.add("hidden");
    }
  }

  async function applyTemplateAI() {
    const text = els.transcript.value.trim();
    if (!text || !state.settings.aiKey || !state.settings.aiEndpoint) {
      applyTemplate();
      return;
    }
    const prompts = {
      summary: "请用中文输出：一句话概括、核心要点列表、关键词。",
      mindmap: "请用中文输出层级 Markdown 思维导图（- 缩进）。",
      aha: "请提取洞见、金句、行动项（checkbox）。",
      research: "请输出深度研究报告：主题、核心发现、关键概念、引用、延伸问题。",
      podcast: "请写成双人播客对话稿，带时间轴与 Show Notes。",
      translate: `请翻译为 ${els.translateLang.value}，保持段落结构。`,
    };
    try {
      toast("AI 正在整理…");
      const data = await api("/api/ai/chat", {
        endpoint: state.settings.aiEndpoint,
        apiKey: state.settings.aiKey,
        model: state.settings.aiModel,
        messages: [
          { role: "system", content: "你是 TicNote 风格的会议与知识整理助手，输出简洁 Markdown。" },
          { role: "user", content: `${prompts[state.tpl]}\n\n原文：\n${text.slice(0, 12000)}` },
        ],
      });
      els.output.textContent = data.content;
      if (state.tpl === "mindmap") {
        state.lastSvg = renderMindmapSvg(buildMindmapTree(text));
        els.mindmapSvgWrap.classList.remove("hidden");
        els.mindmapSvgWrap.innerHTML = state.lastSvg;
      }
    } catch (e) {
      toast("AI 失败，改用本地模板：" + e.message);
      applyTemplate();
    }
  }

  // ─── knowledge base ───
  function persistKb() {
    saveJSON(STORAGE.kb, state.kb);
  }

  function upsertNote(note) {
    const idx = state.kb.findIndex((n) => n.id === note.id);
    if (idx >= 0) state.kb[idx] = note;
    else state.kb.unshift(note);
    state.kb.sort((a, b) => (b.updatedAt || "").localeCompare(a.updatedAt || ""));
    persistKb();
    renderTimeline();
  }

  function currentNotePayload() {
    return {
      id: state.activeId || uid(),
      title: els.noteTitle.value.trim() || "未命名笔记",
      text: els.transcript.value.trim(),
      source: "local",
      tags: ["本地"],
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
  }

  function saveCurrentNote() {
    const note = currentNotePayload();
    if (!note.text) return toast("没有可保存的内容");
    const existing = state.kb.find((n) => n.id === note.id);
    if (existing) note.createdAt = existing.createdAt;
    note.templates = {
      summary: tplSummary(note.text),
      mindmap: tplMindmap(note.text),
      aha: tplAha(note.text),
      research: tplResearch(note.text),
      podcast: tplPodcast(note.text),
    };
    state.activeId = note.id;
    upsertNote(note);
    toast("已保存到知识库");
  }

  function renderTimeline() {
    const q = (els.kbFilter.value || "").trim().toLowerCase();
    const items = state.kb.filter((n) => {
      if (!q) return true;
      return (
        (n.title || "").toLowerCase().includes(q) ||
        (n.text || "").toLowerCase().includes(q) ||
        (n.source || "").toLowerCase().includes(q)
      );
    });
    if (!items.length) {
      els.timeline.innerHTML = `<div class="empty">没有匹配的条目。</div>`;
      return;
    }
    const groups = new Map();
    for (const n of items) {
      const day = (n.updatedAt || n.createdAt || "").slice(0, 10) || todayKey();
      if (!groups.has(day)) groups.set(day, []);
      groups.get(day).push(n);
    }
    const days = [...groups.keys()].sort((a, b) => b.localeCompare(a));
    els.timeline.innerHTML = days
      .map((day) => {
        const cards = groups
          .get(day)
          .map((n) => {
            const tag =
              n.source === "ticnote"
                ? `<span class="tag voice">TicNote</span>`
                : n.source === "web"
                  ? `<span class="tag web">网页</span>`
                  : n.source === "audio"
                    ? `<span class="tag voice">音频</span>`
                    : `<span class="tag">本地</span>`;
            return `<button type="button" class="kb-item ${n.id === state.activeId ? "active" : ""}" data-id="${n.id}">
              <div class="t">${escapeXml(n.title || "未命名")}</div>
              <div class="m">${tag}<span>${formatTime(n.updatedAt)}</span><span>${(n.text || "").length} 字</span></div>
            </button>`;
          })
          .join("");
        return `<div class="day-group"><h3>${day}</h3>${cards}</div>`;
      })
      .join("");
  }

  function openNote(id) {
    const n = state.kb.find((x) => x.id === id);
    if (!n) return;
    state.activeId = id;
    els.noteTitle.value = n.title || "";
    els.transcript.value = n.text || "";
    els.kbDetail.textContent = [
      `标题：${n.title}`,
      `来源：${n.source || "-"}`,
      `更新：${formatTime(n.updatedAt)}`,
      n.recordId ? `recordId：${n.recordId}` : "",
      n.projectId ? `projectId：${n.projectId}` : "",
      "",
      (n.text || "").slice(0, 4000),
    ]
      .filter(Boolean)
      .join("\n");
    renderTimeline();
    applyTemplate();
  }

  // ─── speech ───
  function getRecognition() {
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) return null;
    const rec = new SR();
    rec.continuous = true;
    rec.interimResults = true;
    rec.lang = els.langSelect.value;
    return rec;
  }

  function toggleMic() {
    if (state.recognizing) {
      state.recognition?.stop();
      state.recognizing = false;
      els.btnMic.classList.remove("recording");
      els.btnMic.textContent = "🎙 开始录音转写";
      toast("已停止听写");
      return;
    }
    const rec = getRecognition();
    if (!rec) return toast("当前浏览器不支持语音识别，请用 Chrome / Edge");
    state.recognition = rec;
    let base = els.transcript.value;
    if (base && !base.endsWith("\n")) base += "\n";
    rec.onresult = (ev) => {
      let interim = "";
      let final = "";
      for (let i = ev.resultIndex; i < ev.results.length; i++) {
        const t = ev.results[i][0].transcript;
        if (ev.results[i].isFinal) final += t;
        else interim += t;
      }
      if (final) {
        base += final + ( /[。！？.!?]$/.test(final) ? "" : "。" );
        els.transcript.value = base;
        applyTemplate();
      } else {
        els.transcript.value = base + interim;
      }
    };
    rec.onerror = (e) => {
      toast("识别错误：" + (e.error || "unknown"));
      state.recognizing = false;
      els.btnMic.classList.remove("recording");
      els.btnMic.textContent = "🎙 开始录音转写";
    };
    rec.onend = () => {
      if (state.recognizing) {
        try { rec.start(); } catch { /* ignore */ }
      }
    };
    rec.start();
    state.recognizing = true;
    els.btnMic.classList.add("recording");
    els.btnMic.textContent = "⏹ 停止";
    toast("正在听写…");
  }

  async function audioToText() {
    if (!state.audioBlob && !els.audioPlayer.src) {
      return toast("请先上传音频文件");
    }
    // AI whisper-compatible path
    if (state.settings.aiKey && state.settings.aiEndpoint && state.audioBlob) {
      try {
        toast("正在通过 AI 转录音频…");
        const endpoint = state.settings.aiEndpoint.replace(/\/$/, "");
        const form = new FormData();
        form.append("file", state.audioBlob, state.audioName || "audio.webm");
        form.append("model", "whisper-1");
        form.append("language", els.langSelect.value.slice(0, 2));
        const res = await fetch(endpoint + "/audio/transcriptions", {
          method: "POST",
          headers: { Authorization: "Bearer " + state.settings.aiKey },
          body: form,
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error?.message || "转录失败");
        els.transcript.value = (els.transcript.value + "\n" + (data.text || "")).trim();
        if (!els.noteTitle.value) els.noteTitle.value = state.audioName || "音频转写";
        applyTemplate();
        toast("AI 转录完成");
        return;
      } catch (e) {
        toast("AI 转录失败，尝试播放听写：" + e.message);
      }
    }
    // play-through + Web Speech
    const rec = getRecognition();
    if (!rec) return toast("无法听写：请用 Chrome/Edge，或配置 AI 端点");
    const player = els.audioPlayer;
    if (!player.src && state.audioBlob) {
      player.src = URL.createObjectURL(state.audioBlob);
    }
    let base = els.transcript.value;
    if (base && !base.endsWith("\n")) base += "\n";
    rec.continuous = true;
    rec.interimResults = true;
    rec.lang = els.langSelect.value;
    rec.onresult = (ev) => {
      let final = "";
      for (let i = ev.resultIndex; i < ev.results.length; i++) {
        if (ev.results[i].isFinal) final += ev.results[i][0].transcript;
      }
      if (final) {
        base += final + "。";
        els.transcript.value = base;
      }
    };
    try {
      rec.start();
      await player.play();
      toast("正在播放并实时听写…");
      await new Promise((resolve) => {
        player.onended = resolve;
      });
      rec.stop();
      applyTemplate();
      toast("播放听写结束");
    } catch (e) {
      try { rec.stop(); } catch { /* ignore */ }
      toast("播放听写失败：" + e.message);
    }
  }

  // ─── export ───
  function exportAll() {
    const text = els.transcript.value.trim();
    if (!text) return toast("没有可导出的内容");
    const title = els.noteTitle.value.trim() || "TicNote笔记";
    const safe = title.replace(/[\\/:*?"<>|]/g, "_").slice(0, 40);
    const md = [
      `# ${title}`,
      "",
      `> 由 TicNote Studio 导出 · ${new Date().toLocaleString("zh-CN")}`,
      "",
      tplSummary(text),
      "",
      tplMindmap(text),
      "",
      tplAha(text),
      "",
      tplResearch(text),
      "",
      tplPodcast(text),
      "",
      "## 转录原文",
      "",
      text,
    ].join("\n");
    download(`${safe}-完整笔记.md`, md, "text/markdown;charset=utf-8");
    if (!state.lastSvg) state.lastSvg = renderMindmapSvg(buildMindmapTree(text));
    download(`${safe}-思维导图.svg`, state.lastSvg, "image/svg+xml;charset=utf-8");
    download(`${safe}-转录原文.txt`, text);
    toast("已导出 Markdown / SVG / 原文");
  }

  // ─── TicNote sync ───
  function setConnected(ok) {
    els.connStatus.textContent = ok ? "已连接 TicNote" : "未连接 TicNote";
    els.connStatus.classList.toggle("offline", !ok);
    els.connStatus.classList.toggle("ok", ok);
  }

  async function connectTicnote() {
    const appkey = els.appkey.value.trim();
    if (!appkey) return toast("请填写 AppKey");
    try {
      toast("正在登录…");
      const data = await api("/api/ticnote/login", { appkey });
      state.session = {
        token: data.token,
        appkey,
        baseUrl: data.base_url,
        projects: [],
      };
      saveJSON(STORAGE.session, { appkey, token: data.token, baseUrl: data.base_url });
      setConnected(true);
      await loadProjects();
      toast("连接成功");
    } catch (e) {
      setConnected(false);
      toast("连接失败：" + e.message);
    }
  }

  function disconnectTicnote() {
    state.session = { token: "", appkey: "", baseUrl: "", projects: [] };
    localStorage.removeItem(STORAGE.session);
    setConnected(false);
    els.projectSelect.innerHTML = `<option value="">先连接以加载项目列表</option>`;
    els.projectFiles.innerHTML = "";
    toast("已断开");
  }

  async function loadProjects() {
    const data = await api("/api/ticnote/projects", {
      appkey: state.session.appkey,
      token: state.session.token,
    });
    state.session.projects = data.projects || [];
    const opts = ['<option value="">选择项目…</option>'];
    for (const p of state.session.projects) {
      const val = p.projectId || p.chatId;
      opts.push(
        `<option value="${escapeXml(String(val))}" data-chat="${escapeXml(String(p.chatId || ""))}">${escapeXml(p.name)}（${p.fileNum || 0}）</option>`
      );
    }
    els.projectSelect.innerHTML = opts.join("");
    // 预填用户提到的 chatId
    const hint = "2078873748446011394";
    if (!els.projectIdInput.value) {
      const hit = state.session.projects.find(
        (p) => String(p.chatId) === hint || String(p.projectId) === hint
      );
      if (hit) {
        els.projectSelect.value = String(hit.projectId || hit.chatId);
        els.projectIdInput.value = String(hit.projectId || hit.chatId);
      } else {
        els.projectIdInput.placeholder = hint;
      }
    }
  }

  function authBody(extra) {
    return {
      appkey: state.session.appkey,
      token: state.session.token,
      ...extra,
    };
  }

  async function syncProject() {
    if (!state.session.token) return toast("请先连接 TicNote");
    const projectId =
      els.projectIdInput.value.trim() ||
      els.projectSelect.value.trim() ||
      "2078873748446011394";
    if (!projectId) return toast("请选择或填写 projectId / chatId");
    try {
      toast("正在同步项目并尝试自动转写…");
      const data = await api(
        "/api/ticnote/sync-project",
        authBody({ projectId, chatId: projectId, autoTranscribe: true, language: "zh", limit: 30 })
      );
      let added = 0;
      for (const e of data.entries || []) {
        if (!e.text && !e.summary) {
          // 仍入库占位
        }
        const id = "tn_" + (e.recordId || e.fileId || uid());
        const note = {
          id,
          title: e.name || "TicNote 文件",
          text: e.text || e.summary || "（尚无转写正文，可稍后「拉取/触发转写」）",
          source: "ticnote",
          tags: ["TicNote", e.isVoice ? "音频" : "文档"],
          recordId: e.recordId,
          fileId: e.fileId,
          projectId: e.projectId,
          transcribeId: e.transcribeId,
          dprSessionId: e.dprSessionId,
          isVoice: e.isVoice,
          createdAt: e.updatedAt || new Date().toISOString(),
          updatedAt: e.updatedAt || new Date().toISOString(),
        };
        const prev = state.kb.find((n) => n.id === id);
        if (prev && prev.text && prev.text.length > note.text.length) {
          note.text = prev.text;
        }
        upsertNote(note);
        added++;
      }
      renderProjectFiles(data.entries || []);
      if (data.entries?.[0]) openNote("tn_" + (data.entries[0].recordId || data.entries[0].fileId));
      toast(`同步完成：${added} 个文件`);
      switchView("knowledge");
    } catch (e) {
      toast("同步失败：" + e.message);
    }
  }

  function renderProjectFiles(entries) {
    if (!entries.length) {
      els.projectFiles.innerHTML = `<div class="empty">项目为空或无权访问。</div>`;
      return;
    }
    els.projectFiles.innerHTML = entries
      .map((e) => {
        const st = e.text ? "已有正文" : e.transcodeStatus || "待转写";
        return `<button type="button" class="kb-item" data-record="${escapeXml(String(e.recordId || ""))}" data-file="${escapeXml(String(e.fileId || ""))}">
          <div class="t">${escapeXml(e.name || "文件")}</div>
          <div class="m"><span class="tag ${e.isVoice ? "voice" : ""}">${e.isVoice ? "音频" : "文档"}</span><span>${escapeXml(String(st))}</span></div>
        </button>`;
      })
      .join("");
  }

  async function pullTranscribeActive() {
    const n = state.kb.find((x) => x.id === state.activeId);
    if (!n?.recordId) return toast("当前条目不是 TicNote 文件");
    if (!state.session.token) return toast("请先连接");
    try {
      if (n.fileId && n.isVoice) {
        await api("/api/ticnote/transcribe", authBody({ fileId: n.fileId, language: "zh", hasSpeakers: true }));
      }
      toast("正在拉取详情…");
      const data = await api(
        "/api/ticnote/file-detail",
        authBody({ recordId: n.recordId, poll: true, timeout: 90, interval: 4 })
      );
      if (data.text) {
        n.text = data.text;
        n.updatedAt = new Date().toISOString();
        n.transcribeId = data.detail?.transcribeId || n.transcribeId;
        n.dprSessionId = data.detail?.dprSessionId || n.dprSessionId;
        upsertNote(n);
        openNote(n.id);
        toast("转写正文已更新");
      } else {
        toast("仍无正文，可能仍在处理或需要 VIP");
      }
    } catch (e) {
      toast("拉取失败：" + e.message);
    }
  }

  // ─── web clip ───
  async function fetchWeb() {
    const url = els.webUrl.value.trim();
    if (!url) return toast("请输入网址");
    try {
      toast("抓取网页中…");
      const data = await api("/api/fetch-url", { url });
      els.noteTitle.value = data.title || url;
      els.transcript.value = data.text || "";
      state.activeId = null;
      applyTemplate();
      const note = {
        id: uid(),
        title: data.title || url,
        text: data.text || "",
        source: "web",
        tags: ["网页"],
        url,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      };
      state.activeId = note.id;
      upsertNote(note);
      toast("网页已转入知识库");
    } catch (e) {
      toast("抓取失败：" + e.message);
    }
  }

  // ─── ask / agents ───
  function addBubble(role, text) {
    const div = document.createElement("div");
    div.className = "bubble " + (role === "user" ? "user" : "bot");
    div.textContent = text;
    els.chatBox.appendChild(div);
    els.chatBox.scrollTop = els.chatBox.scrollHeight;
  }

  function scopeText() {
    const scope = document.querySelector('input[name="askScope"]:checked')?.value || "current";
    if (scope === "kb") {
      return state.kb
        .slice(0, 20)
        .map((n) => `【${n.title}】\n${(n.text || "").slice(0, 800)}`)
        .join("\n\n");
    }
    return els.transcript.value.trim();
  }

  function localAnswer(question, context) {
    const q = question.trim();
    if (/翻译|译成|translate/i.test(q)) {
      const lang = /英|en/i.test(q) ? "en" : /日|ja/i.test(q) ? "ja" : /韩|ko/i.test(q) ? "ko" : "en";
      const target = q.split(/[:：]/).slice(1).join(":") || context.slice(0, 1200);
      return tplTranslate(target.trim(), lang).replace(/^## 翻译[^\n]*\n+/, "");
    }
    if (/待办|行动|todo/i.test(q)) {
      return tplAha(context).split("### 行动项")[1]?.trim() || "暂无行动项";
    }
    if (/总结|概括|摘要/i.test(q)) return tplSummary(context);
    if (/关键词/i.test(q)) return extractKeywords(context).join("、") || "暂无";
    const sents = splitSentences(context);
    const hits = sents.filter((s) => q.split(/\s+/).some((w) => w.length > 1 && s.includes(w)));
    if (hits.length) return "相关片段：\n" + hits.slice(0, 5).map((s) => "• " + s).join("\n");
    return (
      "基于当前范围的快速回答：\n" +
      tplSummary(context).split("\n").slice(0, 12).join("\n") +
      "\n\n（配置 AI 端点或使用 TicNote Deep Research 可获得更深回答）"
    );
  }

  async function ask() {
    const q = els.askInput.value.trim();
    if (!q) return;
    const ctx = scopeText();
    if (!ctx) return toast("没有可供提问的内容");
    addBubble("user", q);
    els.askInput.value = "";
    if (state.settings.aiKey && state.settings.aiEndpoint) {
      try {
        const data = await api("/api/ai/chat", {
          endpoint: state.settings.aiEndpoint,
          apiKey: state.settings.aiKey,
          model: state.settings.aiModel,
          messages: [
            {
              role: "system",
              content: "你是 TicNote Studio 助手。根据给定资料回答，不知道就说不知道。用简体中文。",
            },
            { role: "user", content: `资料：\n${ctx.slice(0, 14000)}\n\n问题：${q}` },
          ],
        });
        addBubble("bot", data.content);
        return;
      } catch (e) {
        addBubble("bot", "AI 调用失败，改用本地回答。\n" + localAnswer(q, ctx));
        return;
      }
    }
    addBubble("bot", localAnswer(q, ctx));
  }

  function renderAgents() {
    els.agentList.innerHTML = AGENTS.map(
      (a) => `<div class="agent-card" data-agent="${a.id}">
        <h4>${a.name}</h4>
        <p>${a.desc}</p>
        <div class="row">
          <button type="button" class="teal" data-run="${a.id}">用当前文稿运行</button>
          ${a.ticnote ? `<button type="button" class="amber" data-deep="${a.id}">TicNote 深研</button>` : ""}
        </div>
      </div>`
    ).join("");
  }

  async function runAgent(id) {
    const agent = AGENTS.find((a) => a.id === id);
    if (!agent) return;
    state.tpl = agent.tpl;
    [...els.tplBar.querySelectorAll(".tpl")].forEach((b) =>
      b.classList.toggle("active", b.dataset.tpl === state.tpl)
    );
    switchView("capture");
    if (state.settings.aiKey) await applyTemplateAI();
    else applyTemplate();
    toast(`智能体「${agent.name}」已生成`);
  }

  async function runDeepResearch() {
    const n = state.kb.find((x) => x.id === state.activeId);
    if (!state.session.token) return toast("请先连接 TicNote");
    if (!n?.recordId && !n?.dprSessionId) return toast("请先打开一条 TicNote 同步条目");
    const question =
      prompt("输入深研问题", `分析「${n.title}」的核心观点与行动建议`) || "";
    if (!question.trim()) return;
    try {
      toast("已提交 Deep Research…");
      const data = await api(
        "/api/ticnote/deep-research",
        authBody({
          recordId: n.recordId,
          sessionId: n.dprSessionId,
          question: question.trim(),
        })
      );
      addBubble("bot", "Deep Research 已提交：\n" + JSON.stringify(data.data, null, 2));
      switchView("ask");
      toast("已提交到 TicNote");
    } catch (e) {
      toast("深研失败：" + e.message);
    }
  }

  // ─── views ───
  function switchView(name) {
    state.view = name;
    [...els.mainNav.querySelectorAll("button")].forEach((b) =>
      b.classList.toggle("active", b.dataset.view === name)
    );
    ["capture", "knowledge", "agents", "ask"].forEach((v) => {
      const el = document.getElementById("view-" + v);
      if (!el) return;
      el.classList.toggle("hidden", v !== name);
    });
    const titles = {
      capture: "采集与转写",
      knowledge: "知识库详情",
      agents: "智能体",
      ask: "提问与翻译",
    };
    els.mainTitle.textContent = titles[name] || "";
  }

  function loadSample() {
    els.noteTitle.value = "注意力与时间管理工作坊";
    els.transcript.value = [
      "今天想聊聊时间管理与注意力。很多人以为时间管理就是把日程排满，其实核心是注意力管理。",
      "如果你早上最清醒的两小时被会议占满，那么深度工作会被迫挤到晚上，效率会明显下降。",
      "一个可行方法是：每天保护一段九十分钟的专注块，关闭通知，只做一件高价值的事。",
      "会议方面，会前写清决策问题，会中只讨论阻塞点，会后立刻写下行动项和负责人。",
      "最后，每周五做一次复盘：哪些事情真正推进了目标，哪些只是忙碌的幻觉。",
    ].join("");
    state.activeId = null;
    applyTemplate();
    saveCurrentNote();
    toast("已载入示例");
  }

  // ─── events ───
  function bind() {
    els.mainNav.addEventListener("click", (e) => {
      const btn = e.target.closest("button[data-view]");
      if (btn) switchView(btn.dataset.view);
    });
    els.tplBar.addEventListener("click", (e) => {
      const btn = e.target.closest(".tpl");
      if (!btn) return;
      state.tpl = btn.dataset.tpl;
      [...els.tplBar.querySelectorAll(".tpl")].forEach((b) =>
        b.classList.toggle("active", b === btn)
      );
      if (state.settings.aiKey && (state.tpl === "translate" || state.tpl === "research")) {
        applyTemplateAI();
      } else applyTemplate();
    });
    els.transcript.addEventListener("input", () => applyTemplate());
    els.translateLang.addEventListener("change", () => applyTemplate());
    els.kbFilter.addEventListener("input", renderTimeline);
    els.timeline.addEventListener("click", (e) => {
      const item = e.target.closest(".kb-item[data-id]");
      if (item) {
        openNote(item.dataset.id);
        if (state.view === "capture") applyTemplate();
        else switchView("knowledge");
      }
    });
    els.btnMic.addEventListener("click", toggleMic);
    els.btnAudio2Text.addEventListener("click", audioToText);
    els.btnExportAll.addEventListener("click", exportAll);
    els.btnSaveNote.addEventListener("click", saveCurrentNote);
    els.btnLoadSample.addEventListener("click", loadSample);
    els.btnRefreshKb.addEventListener("click", renderTimeline);
    els.btnConnect.addEventListener("click", connectTicnote);
    els.btnDisconnect.addEventListener("click", disconnectTicnote);
    els.btnSyncTicnote.addEventListener("click", () => {
      if (!state.session.token) {
        $("settingsModal");
        toast("请先在右侧填写 AppKey 并连接");
        return;
      }
      syncProject();
    });
    els.btnSyncProject.addEventListener("click", syncProject);
    els.projectSelect.addEventListener("change", () => {
      els.projectIdInput.value = els.projectSelect.value;
    });
    els.btnFetchWeb.addEventListener("click", fetchWeb);
    els.btnAsk.addEventListener("click", ask);
    els.askInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") ask();
    });
    els.btnQuickTranslate.addEventListener("click", () => {
      state.tpl = "translate";
      [...els.tplBar.querySelectorAll(".tpl")].forEach((b) =>
        b.classList.toggle("active", b.dataset.tpl === "translate")
      );
      switchView("capture");
      applyTemplate();
      addBubble("bot", els.output.textContent);
      switchView("ask");
    });
    els.audioFile.addEventListener("change", () => {
      const f = els.audioFile.files?.[0];
      if (!f) return;
      state.audioBlob = f;
      state.audioName = f.name;
      els.audioPlayer.src = URL.createObjectURL(f);
      els.audioMeta.classList.remove("hidden");
      els.audioMeta.textContent = `已选：${f.name}（${Math.round(f.size / 1024)} KB）→ 点「音频转文字」`;
      els.btnAudio2Text.classList.remove("hidden");
      if (!els.noteTitle.value) els.noteTitle.value = f.name.replace(/\.[^.]+$/, "");
      toast("音频已就绪");
    });
    $("btnCopyOut").addEventListener("click", async () => {
      await navigator.clipboard.writeText(els.output.textContent);
      toast("已复制");
    });
    $("btnExportMd").addEventListener("click", () => {
      const title = els.noteTitle.value.trim() || "note";
      download(`${title}.md`, els.output.textContent, "text/markdown;charset=utf-8");
    });
    $("btnExportSvg").addEventListener("click", () => {
      if (!state.lastSvg) applyTemplate();
      if (!state.lastSvg) return toast("请先生成思维导图");
      download("mindmap.svg", state.lastSvg, "image/svg+xml;charset=utf-8");
    });
    $("btnOpenInCapture").addEventListener("click", () => switchView("capture"));
    $("btnDeleteNote").addEventListener("click", () => {
      if (!state.activeId) return;
      state.kb = state.kb.filter((n) => n.id !== state.activeId);
      state.activeId = null;
      persistKb();
      renderTimeline();
      els.kbDetail.textContent = "已删除。";
      toast("已删除");
    });
    $("btnPullTranscribe").addEventListener("click", pullTranscribeActive);
    els.agentList.addEventListener("click", (e) => {
      const run = e.target.closest("[data-run]");
      const deep = e.target.closest("[data-deep]");
      if (run) runAgent(run.dataset.run);
      if (deep) runDeepResearch();
    });
    $("btnSettings").addEventListener("click", () => els.settingsModal.classList.add("show"));
    $("btnCloseSettings").addEventListener("click", () => els.settingsModal.classList.remove("show"));
    $("btnSaveSettings").addEventListener("click", () => {
      state.settings = {
        aiEndpoint: els.aiEndpoint.value.trim(),
        aiKey: els.aiKey.value.trim(),
        aiModel: els.aiModel.value.trim() || "gpt-4o-mini",
      };
      saveJSON(STORAGE.settings, state.settings);
      els.settingsModal.classList.remove("show");
      toast("设置已保存");
    });
    els.projectFiles.addEventListener("click", (e) => {
      const item = e.target.closest(".kb-item[data-record]");
      if (!item) return;
      const id = "tn_" + item.dataset.record;
      if (state.kb.some((n) => n.id === id)) openNote(id);
      else toast("请先点「同步并自动转写」入库");
    });
  }

  function init() {
    state.kb = loadJSON(STORAGE.kb, []);
    state.settings = { ...state.settings, ...loadJSON(STORAGE.settings, {}) };
    els.aiEndpoint.value = state.settings.aiEndpoint || "";
    els.aiKey.value = state.settings.aiKey || "";
    els.aiModel.value = state.settings.aiModel || "gpt-4o-mini";
    const sess = loadJSON(STORAGE.session, null);
    if (sess?.token && sess?.appkey) {
      state.session = { ...state.session, ...sess, projects: [] };
      els.appkey.value = sess.appkey;
      setConnected(true);
      loadProjects().catch(() => setConnected(false));
    }
    // 预填用户链接中的 chatId
    els.projectIdInput.value = els.projectIdInput.value || "";
    els.projectIdInput.placeholder = "2078873748446011394";
    renderAgents();
    renderTimeline();
    bind();
    switchView("capture");
    // 探测本地代理
    fetch(proxy() + "/api/health")
      .then((r) => (r.ok ? r.json() : null))
      .then((d) => {
        if (d?.ok) toast("本地服务已就绪");
      })
      .catch(() => {
        /* 允许纯本地离线使用 */
      });
  }

  init();
})();
