// GMath Word 加载项 —— 核心逻辑（纯原生公式 + 双向实时同步，尽力而为）
// 正向：MathLive 编辑 → MathML → mml2omml 转 OMML → Flat OPC → insertOoxml
// 反向：点选文档里的公式 → getOoxml → inspectOmathPackage → omml2latex → 载入面板
// 同步：新公式用隐藏内容控件 + UUID tag 锚定；旧公式首次回写前验证选区快照并建立锚点。
// Word.run 任务串行执行，选区变化会使尚未开始的旧任务失效，避免异步竞态误替换正文。

import { mml2ommlDetailed } from "./mathml2omml.js";
import { omml2latexDetailed, inspectOmathPackage } from "./omml2latex.js";
import { latexToOmml, hasDecoration } from "./latex2omml.js";
import { SYMBOL_CATEGORIES, CAT_EN } from "./symbols.js";
import { I18N, TIP_EN } from "./i18n.js";
import { buildFlatOpc } from "./ooxml.js";
import { LatestTaskQueue, makeAnchorTag, normalizeOmmlFingerprint } from "./sync-state.js";
import { VERSION } from "./version.js";
import { AI_LIMITS, cleanLatex, normalizeEndpoint } from "./ai-client.js";

const BUILD = VERSION;

const $ = (id) => document.getElementById(id);
const boot = window.__gmathBoot || { start: performance.now(), marks: [] };
const bootMark = (name) => {
  boot.marks.push({ name, ms: Math.round(performance.now() - boot.start) });
};
const bootReport = () =>
  ["GMath 启动耗时:", ...boot.marks.map((m) => `- ${m.name}: ${m.ms} ms`)].join("\n");

// ===== 国际化（中文 / English，本机记忆选择） =====
let lang = "zh";
try {
  const saved = localStorage.getItem("gmath.lang");
  if (saved === "en" || saved === "zh") lang = saved;
} catch { /* localStorage 不可用时默认中文 */ }

const T = (key) => (I18N[lang] && I18N[lang][key] != null ? I18N[lang][key] : key);
const catLabel = (idx) => (lang === "en" ? CAT_EN[idx] : SYMBOL_CATEGORIES[idx].name);
const tipText = (zh) => (lang === "en" ? TIP_EN[zh] || zh : zh);

function applyI18n() {
  document.querySelectorAll("[data-i18n]").forEach((el) => {
    const v = I18N[lang][el.dataset.i18n];
    if (v != null) el.textContent = v;
  });
  document.querySelectorAll("[data-i18n-html]").forEach((el) => {
    const v = I18N[lang][el.dataset.i18nHtml];
    if (v != null) el.innerHTML = v;
  });
  document.querySelectorAll("[data-i18n-ph]").forEach((el) => {
    const v = I18N[lang][el.dataset.i18nPh];
    if (v != null) el.placeholder = v;
  });
  document.querySelectorAll("[data-i18n-title]").forEach((el) => {
    const v = I18N[lang][el.dataset.i18nTitle];
    if (v != null) el.title = v;
  });
  document.querySelectorAll("[data-i18n-aria]").forEach((el) => {
    const v = I18N[lang][el.dataset.i18nAria];
    if (v != null) el.setAttribute("aria-label", v);
  });
  document.documentElement.lang = lang === "en" ? "en" : "zh-CN";
  const lb = $("langBtn");
  if (lb) lb.textContent = lang === "zh" ? "EN" : "中";
  refreshPaletteLang();
  if (modeLine) setMode(linked); // 模式徽章按当前连接状态重译
}

function refreshPaletteLang() {
  const tabs = $("catTabs");
  if (!tabs) return;
  tabs.querySelectorAll("button[data-cat]").forEach((b) => {
    b.textContent = catLabel(Number(b.dataset.cat));
  });
  const active = tabs.querySelector("button.active");
  if (active) showPaletteCat(Number(active.dataset.cat));
}

function setLang(l) {
  lang = l === "en" ? "en" : "zh";
  try { localStorage.setItem("gmath.lang", lang); } catch { /* ignore */ }
  applyI18n();
}

let mathfield, latexEl, statusEl, debugEl, insertBtn, layoutModeEl, eqNumberEl, eqNumberRow, modeLine;

// 同步状态机
let linked = false; // 面板当前是否“连接”着文档里的某个公式
let loadedLatex = null; // 最近一次从文档读到的 LaTeX（用于判断是否同一个公式、避免误重载）
let applying = false; // 正在回写（屏蔽选区监听，防自我触发）
let loadingDoc = false; // 正在把文档公式灌入面板（屏蔽 mathfield→回写）
let selTimer = null;
let syncTimer = null;
let selectionRevision = 0;
let selectionReadQueued = false;
let activeAnchorTag = null;
let loadedOmmlFingerprint = null;
const syncQueue = new LatestTaskQueue();

bootMark("taskpane module evaluating");

function setStatus(msg, kind = "") {
  statusEl.textContent = msg;
  statusEl.className = "status" + (kind ? " " + kind : "");
}

function setStartupStatus(msg, kind = "") {
  const el = $("status");
  if (!el) return;
  el.textContent = msg;
  el.className = "status" + (kind ? " " + kind : "");
}

function writeStartupDebug(prefix = "") {
  const el = $("debug");
  if (!el) return;
  el.value = (prefix ? prefix + "\n\n" : "") + bootReport();
}

function setMode(isLinked) {
  linked = isLinked;
  if (isLinked) {
    modeLine.textContent = T("modeLinked");
    modeLine.classList.add("linked");
  } else {
    clearTimeout(syncTimer);
    syncQueue.invalidate();
    loadedLatex = null;
    activeAnchorTag = null;
    loadedOmmlFingerprint = null;
    modeLine.textContent = T("modeNew");
    modeLine.classList.remove("linked");
  }
}

function currentOmml() {
  const latex = mathfield.getValue("latex");
  // 含装饰命令（\overline/\overbrace 等 MathLive 序列化会丢内容的）走补丁层；
  // 其余保持原 MathML 快路径，行为与旧版完全一致（零回归）。
  if (hasDecoration(latex) && window.MathLive?.convertLatexToMathMl) {
    const omml = latexToOmml(latex, {
      convertLatexToMathMl: (s) => window.MathLive.convertLatexToMathMl(s),
      mml2omml: (mathml) => {
        const result = mml2ommlDetailed(mathml);
        if (result.lossy) throw new Error(`${T("unsupportedFormula")} ${result.warnings.join(", ")}`);
        return result.omml;
      },
    });
    return omml && omml.trim() ? omml : null;
  }
  const mathml = mathfield.getValue("math-ml");
  if (!mathml || !mathml.trim()) return null;
  const result = mml2ommlDetailed(mathml);
  if (result.lossy) throw new Error(`${T("unsupportedFormula")} ${result.warnings.join(", ")}`);
  return result.omml;
}

const describeError = (e) => ({ code: e.code, message: e.message, debugInfo: e.debugInfo });

// 单个符号/模板按钮：文本标签 l + 点击插入模板 i（#0 → 插入后的光标占位）
function renderSymButton(it, catIdx) {
  const b = document.createElement("button");
  b.type = "button";
  b.dataset.insert = it.i;
  const title = it.t ? tipText(it.t) : it.i;
  b.title = title;
  b.setAttribute("aria-label", title);
  if (catIdx === 0 && it.d) {
    b.className = "template";
    const main = document.createElement("span");
    main.className = "sym-main";
    if (it.m) main.innerHTML = it.m;
    else main.textContent = it.l;
    const desc = document.createElement("span");
    desc.className = "sym-desc";
    desc.textContent = tipText(it.d);
    b.append(main, desc);
  } else {
    b.textContent = it.l;
  }
  return b;
}

function showPaletteCat(idx) {
  const grid = $("symGrid");
  if (!grid) return;
  grid.innerHTML = "";
  grid.hidden = false;
  grid.classList.toggle("common", idx === 0);
  SYMBOL_CATEGORIES[idx].items.forEach((it) => grid.appendChild(renderSymButton(it, idx)));
}

function collapsePalette() {
  const grid = $("symGrid");
  if (!grid) return;
  grid.innerHTML = "";
  grid.hidden = true;
}

// 渲染分类符号速选：上方分类标签，点选某类后在下方展开该类符号
function renderPalette() {
  const tabs = $("catTabs");
  if (!tabs || !$("symGrid")) return;

  SYMBOL_CATEGORIES.forEach((cat, idx) => {
    const t = document.createElement("button");
    t.type = "button";
    t.textContent = catLabel(idx);
    t.dataset.cat = String(idx);
    t.setAttribute("aria-controls", "symGrid");
    t.setAttribute("aria-expanded", idx === 0 ? "true" : "false");
    if (idx === 0) t.classList.add("active");
    tabs.appendChild(t);
  });

  tabs.addEventListener("click", (ev) => {
    const t = ev.target.closest("button[data-cat]");
    if (!t) return;
    const isOpen = t.classList.contains("active");
    tabs.querySelectorAll("button").forEach((b) => {
      b.classList.remove("active");
      b.setAttribute("aria-expanded", "false");
    });
    if (isOpen) {
      collapsePalette();
      return;
    }
    t.classList.add("active");
    t.setAttribute("aria-expanded", "true");
    showPaletteCat(Number(t.dataset.cat));
  });

  tabs.querySelector("button[data-cat='0']")?.setAttribute("aria-expanded", "true");
  showPaletteCat(0);
}

// 符号速选是纯 DOM，模块加载即渲染，立即可见可切换（此刻 MathLive 可能未就绪，
// 先用文本回退；待 wireUp 中 MathLive 确认就绪后再重渲染为真实公式预览）。
renderPalette();
applyI18n(); // 按记忆的语言完成首次本地化（含分类标签）

// 当前选中的版式：inline / display / numbered
function getLayout() {
  const active = layoutModeEl.querySelector("button.active");
  return active ? active.dataset.mode : "display";
}

function setLayout(mode) {
  layoutModeEl.querySelectorAll("button").forEach((b) => {
    b.classList.toggle("active", b.dataset.mode === mode);
    b.setAttribute("aria-pressed", b.dataset.mode === mode ? "true" : "false");
  });
  eqNumberRow.hidden = mode !== "numbered"; // 编号输入只在「右编号」时出现
  scheduleSync();
}

// 把一段 LaTeX 灌入面板（不触发回写）
function loadIntoPane(latex, display) {
  loadingDoc = true;
  try {
    mathfield.setValue(latex);
    latexEl.value = mathfield.getValue("latex");
    if (typeof display === "boolean") setLayout(display ? "display" : "inline");
  } finally {
    loadingDoc = false;
  }
}

function finishApplying() {
  applying = false;
  if (selectionReadQueued) {
    selectionReadQueued = false;
    clearTimeout(selTimer);
    selTimer = setTimeout(readFromSelection, 0);
  }
}

function anchorRange(range, tag) {
  const control = range.insertContentControl();
  control.tag = tag;
  control.title = "GMath formula";
  control.appearance = "Hidden";
  return control;
}

function conversionError(error, action) {
  setStatus(`${action}${error.message || error}`, "err");
  if (debugEl) debugEl.value = `[转换保护] ${error.stack || error}\n\n${debugEl.value}`;
}

// ===== 面板 → 文档：插入一个新公式（不连接） =====
async function insertNew() {
  let omml;
  try {
    omml = currentOmml();
  } catch (error) {
    conversionError(error, T("insertFail"));
    return;
  }
  if (!omml) {
    setStatus(T("emptyFormula"), "err");
    return;
  }
  const flatOpc = buildFlatOpc(omml, getLayout(), eqNumberEl.value);
  const anchorTag = makeAnchorTag();
  debugEl.value = "构建版本: " + BUILD + "\n\nOMML:\n" + omml + "\n\n完整 OOXML:\n" + flatOpc;
  applying = true;
  insertBtn.disabled = true;
  try {
    await Word.run(async (context) => {
      const r = context.document.getSelection().insertOoxml(flatOpc, Word.InsertLocation.replace);
      anchorRange(r, anchorTag);
      r.select(); // 选中新公式，便于随后继续在面板里改、自动同步
      await context.sync();
    });
    loadedLatex = mathfield.getValue("latex");
    loadedOmmlFingerprint = normalizeOmmlFingerprint(omml);
    activeAnchorTag = anchorTag;
    setMode(true);
    setStatus(T("insertedLinked"), "ok");
  } catch (e) {
    setStatus(T("insertFail") + (e.code || "") + " " + (e.message || e), "err");
    debugEl.value += "\n\n=== 错误详情 ===\n" + JSON.stringify(describeError(e), null, 2);
  } finally {
    insertBtn.disabled = false;
    finishApplying();
  }
}

// ===== 面板 → 文档：把改动同步回当前选中的公式（替换选区） =====
async function syncToDoc({ isLatest }) {
  if (!linked) return;
  let omml;
  try {
    omml = currentOmml();
  } catch (error) {
    conversionError(error, T("syncFail"));
    return;
  }
  if (!omml) return;
  const flatOpc = buildFlatOpc(omml, getLayout(), eqNumberEl.value);
  const anchorTag = activeAnchorTag;
  const expectedFingerprint = loadedOmmlFingerprint;
  let nextAnchorTag = anchorTag;
  let wrote = false;
  applying = true;
  try {
    await Word.run(async (context) => {
      let r;
      if (anchorTag) {
        const controls = context.document.contentControls.getByTag(anchorTag);
        controls.load("items");
        await context.sync();
        if (!isLatest()) return;
        if (controls.items.length !== 1) {
          const error = new Error(T("anchorLost"));
          error.code = "AnchorLost";
          throw error;
        }
        r = controls.items[0].insertOoxml(flatOpc, Word.InsertLocation.replace);
      } else {
        const selection = context.document.getSelection();
        const snapshot = selection.getOoxml();
        await context.sync();
        if (!isLatest()) return;
        const inspected = inspectOmathPackage(snapshot.value);
        const actualFingerprint = normalizeOmmlFingerprint(inspected.omml);
        if (!inspected.safe || !expectedFingerprint || actualFingerprint !== expectedFingerprint) {
          const error = new Error(T("selectionChanged"));
          error.code = "SelectionChanged";
          throw error;
        }
        r = selection.insertOoxml(flatOpc, Word.InsertLocation.replace);
        nextAnchorTag = makeAnchorTag();
        anchorRange(r, nextAnchorTag);
      }
      r.select(); // 替换后重新选中，保证下一次还能替换到它
      await context.sync();
      wrote = true;
    });
    if (!wrote) return;
    loadedLatex = mathfield.getValue("latex");
    loadedOmmlFingerprint = normalizeOmmlFingerprint(omml);
    activeAnchorTag = nextAnchorTag;
    setStatus(T("syncedDoc"), "ok");
  } catch (e) {
    if (e.code === "AnchorLost" || e.code === "SelectionChanged") setMode(false);
    setStatus(T("syncFail") + (e.code || "") + " " + (e.message || e), "err");
  } finally {
    finishApplying();
  }
}

function scheduleSync() {
  if (!linked || loadingDoc) return;
  clearTimeout(syncTimer);
  syncTimer = setTimeout(() => {
    syncQueue.enqueue(syncToDoc).catch((error) => conversionError(error, T("syncFail")));
  }, 600);
}

// ===== 文档 → 面板：选区变化时，若选中的是公式就载入面板 =====
function onSelectionChanged() {
  selectionRevision++;
  syncQueue.invalidate();
  if (applying) {
    selectionReadQueued = true;
    return;
  }
  clearTimeout(selTimer);
  selTimer = setTimeout(readFromSelection, 250);
}

async function readFromSelection() {
  if (applying) {
    selectionReadQueued = true;
    return;
  }
  const revision = selectionRevision;
  try {
    let ooxml = null;
    let selectedControlTag = null;
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const res = selection.getOoxml();
      const controls = selection.contentControls;
      controls.load("items/tag");
      await context.sync();
      ooxml = res.value;
      selectedControlTag = controls.items
        .map((control) => control.tag)
        .find((tag) => tag && tag.startsWith("gmath:")) || null;
    });
    if (revision !== selectionRevision) return;
    // 光标可能位于内容控件内部，此时 selection.contentControls 不一定包含父控件。
    // parentContentControl 属于 WordApi 1.1；未处于控件内时会抛 ItemNotFound，安全忽略。
    if (!selectedControlTag) {
      try {
        await Word.run(async (context) => {
          const parent = context.document.getSelection().parentContentControl;
          parent.load("tag");
          await context.sync();
          if (parent.tag?.startsWith("gmath:")) selectedControlTag = parent.tag;
        });
      } catch { /* 当前选区没有父内容控件 */ }
      if (revision !== selectionRevision) return;
    }
    const inspected = inspectOmathPackage(ooxml);
    if (!inspected.omml) {
      if (linked) setMode(false); // 选区离开了公式 → 断开，避免误同步
      return;
    }
    if (!inspected.safe) {
      setMode(false);
      setStatus(T("unsafeSelection"), "err");
      return;
    }
    const converted = omml2latexDetailed(inspected.omml);
    if (converted.lossy) {
      setMode(false);
      setStatus(`${T("unsupportedSelected")} ${converted.warnings.join(", ")}`, "err");
      return;
    }
    const latex = converted.latex;
    const observedAnchorTag = selectedControlTag || inspected.gmathTag;
    if (linked && latex === loadedLatex && observedAnchorTag === activeAnchorTag && observedAnchorTag) {
      return; // 仍是同一个带锚点公式，别覆盖正在改的内容
    }
    loadIntoPane(latex, /oMathPara/.test(inspected.omml));
    loadedLatex = latex;
    loadedOmmlFingerprint = normalizeOmmlFingerprint(inspected.omml);
    activeAnchorTag = observedAnchorTag;
    setMode(true);
    setStatus(T("loadedSel"), "ok");
  } catch (e) {
    debugEl.value = "[选区读取] " + (e.message || e) + "\n" + debugEl.value;
  }
}

// ===== 图片转公式（AI 识别，OpenAI 兼容视觉接口） =====
const AI_KEYS = {
  url: "gmath.ai.url",
  key: "gmath.ai.key",
  model: "gmath.ai.model",
  remember: "gmath.ai.remember",
};
let aiController = null;
let aiRequestRevision = 0;
let proxyTokenPromise = null;

function lsGet(k) {
  try { return localStorage.getItem(k) || ""; } catch { return ""; }
}
function lsSet(k, v) {
  try { localStorage.setItem(k, v); } catch { /* 某些 WebView 禁用 localStorage */ }
}
function lsRemove(k) {
  try { localStorage.removeItem(k); } catch { /* ignore */ }
}
function ssGet(k) {
  try { return sessionStorage.getItem(k) || ""; } catch { return ""; }
}
function ssSet(k, v) {
  try { sessionStorage.setItem(k, v); } catch { /* ignore */ }
}
function ssRemove(k) {
  try { sessionStorage.removeItem(k); } catch { /* ignore */ }
}

function setAiStatus(msg, kind = "") {
  const el = $("aiStatus");
  if (!el) return;
  el.hidden = !msg;
  el.textContent = msg || "";
  el.className = "ai-status" + (kind ? " " + kind : "");
}

function localProxyUrl(path = "/api/ai/chat/completions") {
  if (["localhost", "127.0.0.1", "[::1]"].includes(window.location.hostname)) {
    return path;
  }
  return `https://localhost:3000${path}`;
}

function aiPayload(model, dataUrl) {
  return JSON.stringify({
    model,
    temperature: 0,
    messages: [
      { role: "system", content: T("aiSysPrompt") },
      {
        role: "user",
        content: [
          { type: "text", text: T("aiUserPrompt") },
          { type: "image_url", image_url: { url: dataUrl } },
        ],
      },
    ],
  });
}

async function getProxyToken(signal) {
  if (!proxyTokenPromise) {
    proxyTokenPromise = fetch(localProxyUrl("/api/ai/token"), { signal })
      .then(async (response) => {
        if (!response.ok) throw new Error(`Local proxy token request failed (${response.status})`);
        const data = await response.json();
        if (!data?.token) throw new Error("Local proxy did not return a token");
        return data.token;
      })
      .catch((error) => {
        proxyTokenPromise = null;
        throw error;
      });
  }
  return proxyTokenPromise;
}

async function postAi(endpoint, key, body, useProxy, signal) {
  const headers = {
    "Content-Type": "application/json",
    Authorization: "Bearer " + key,
  };
  if (useProxy) {
    headers["X-GMath-AI-Endpoint"] = endpoint;
    headers["X-GMath-Proxy-Token"] = await getProxyToken(signal);
  }
  return fetch(useProxy ? localProxyUrl() : endpoint, {
    method: "POST",
    headers,
    body,
    signal,
  });
}

// 读取图片文件 → 压缩到最长边 ≤1600 的 dataURL（减小体积/费用，PNG 保边缘清晰）
function fileToDataUrl(file, maxSide = 1600) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error(T("readImgFail")));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error(T("parseImgFail")));
      img.onload = () => {
        if (img.width * img.height > AI_LIMITS.maxPixels) {
          reject(new Error(T("imageTooLarge")));
          return;
        }
        const scale = Math.min(1, maxSide / Math.max(img.width, img.height));
        if (scale === 1) return resolve(reader.result);
        const c = document.createElement("canvas");
        c.width = Math.round(img.width * scale);
        c.height = Math.round(img.height * scale);
        c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
        resolve(c.toDataURL("image/png"));
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

async function recognizeImage(dataUrl) {
  const url = $("aiUrl").value.trim();
  const key = $("aiKey").value.trim();
  const model = $("aiModel").value.trim();
  if (!url || !key || !model) {
    setAiStatus(T("aiNeedCfg"), "err");
    $("aiSettings").hidden = false;
    $("aiSettingsBtn").setAttribute("aria-expanded", "true");
    return;
  }
  aiController?.abort();
  const revision = ++aiRequestRevision;
  const controller = new AbortController();
  aiController = controller;
  const timeout = setTimeout(() => controller.abort(), AI_LIMITS.requestTimeoutMs);
  setAiStatus(T("aiBusy"), "busy");
  try {
    const endpoint = normalizeEndpoint(url);
    const body = aiPayload(model, dataUrl);
    let resp;
    if (new URL(endpoint).protocol === "https:") {
      try {
        resp = await postAi(endpoint, key, body, false, controller.signal);
      } catch (directErr) {
        if (controller.signal.aborted) throw directErr;
        resp = await postAi(endpoint, key, body, true, controller.signal).catch(() => {
          throw directErr;
        });
      }
    } else {
      // HTTP Key 绝不由任务窗直发；仅交给显式允许不安全目标的本地代理决定。
      resp = await postAi(endpoint, key, body, true, controller.signal);
    }
    if (revision !== aiRequestRevision) return;
    if (!resp.ok) {
      const body = await resp.text().catch(() => "");
      setAiStatus(`${T("aiHttp")}${resp.status}: ${body.slice(0, 200)}`, "err");
      return;
    }
    const data = await resp.json();
    if (revision !== aiRequestRevision) return;
    const content = data?.choices?.[0]?.message?.content;
    const text = typeof content === "string"
      ? content
      : Array.isArray(content) ? content.map((c) => c?.text || "").join("") : "";
    const latex = cleanLatex(text);
    if (!latex) {
      setAiStatus(T("aiNoFormula"), "err");
      return;
    }
    mathfield.setValue(latex);
    latexEl.value = mathfield.getValue("latex");
    setMode(false); // 作为新公式载入，等用户确认后再插入
    mathfield.focus();
    setAiStatus(T("aiDone"), "ok");
  } catch (e) {
    if (revision !== aiRequestRevision) return;
    if (controller.signal.aborted) {
      setAiStatus(T("aiTimeout"), "err");
      return;
    }
    setAiStatus(`${T("aiReqFail")}${e.message || e}${T("aiReqFailHint")}`, "err");
  } finally {
    clearTimeout(timeout);
    if (aiController === controller) aiController = null;
  }
}

async function handleImageFile(file) {
  if (!file || !file.type || !file.type.startsWith("image/")) return;
  if (file.size > AI_LIMITS.maxFileBytes) {
    setAiStatus(T("imageTooLarge"), "err");
    return;
  }
  try {
    const dataUrl = await fileToDataUrl(file);
    const thumb = $("aiThumb");
    thumb.src = dataUrl;
    thumb.hidden = false;
    $("aiHint").hidden = true;
    await recognizeImage(dataUrl);
  } catch (e) {
    setAiStatus(e.message || T("procImgFail"), "err");
  }
}

function wireAI() {
  const urlEl = $("aiUrl");
  const keyEl = $("aiKey");
  const modelEl = $("aiModel");
  const rememberEl = $("aiRememberKey");
  const drop = $("aiDrop");
  const fileEl = $("aiFile");

  // 载入已保存的配置
  urlEl.value = lsGet(AI_KEYS.url);
  let rememberedKey = lsGet(AI_KEYS.key);
  const explicitlyRemembered = lsGet(AI_KEYS.remember) === "yes";
  // 从旧版本迁移：没有明确同意长期保存的历史 Key 自动降级到会话存储。
  if (rememberedKey && !explicitlyRemembered) {
    ssSet(AI_KEYS.key, rememberedKey);
    lsRemove(AI_KEYS.key);
    rememberedKey = "";
  }
  keyEl.value = rememberedKey || ssGet(AI_KEYS.key);
  rememberEl.checked = !!rememberedKey && explicitlyRemembered;
  modelEl.value = lsGet(AI_KEYS.model);

  $("aiSettingsBtn").addEventListener("click", () => {
    const s = $("aiSettings");
    s.hidden = !s.hidden;
    $("aiSettingsBtn").setAttribute("aria-expanded", s.hidden ? "false" : "true");
  });
  $("aiSaveBtn").addEventListener("click", () => {
    lsSet(AI_KEYS.url, urlEl.value.trim());
    lsSet(AI_KEYS.model, modelEl.value.trim());
    const key = keyEl.value.trim();
    if (rememberEl.checked) {
      lsSet(AI_KEYS.key, key);
      lsSet(AI_KEYS.remember, "yes");
      ssRemove(AI_KEYS.key);
    } else {
      lsRemove(AI_KEYS.key);
      lsRemove(AI_KEYS.remember);
      ssSet(AI_KEYS.key, key);
    }
    setAiStatus(T("aiSaved"), "ok");
    $("aiSettings").hidden = true;
    $("aiSettingsBtn").setAttribute("aria-expanded", "false");
  });

  // 点击 / 拖拽 / 选择文件
  drop.addEventListener("click", () => fileEl.click());
  drop.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fileEl.click(); }
  });
  fileEl.addEventListener("change", () => {
    if (fileEl.files && fileEl.files[0]) handleImageFile(fileEl.files[0]);
    fileEl.value = "";
  });
  drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("drag"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("drag"));
  drop.addEventListener("drop", (e) => {
    e.preventDefault();
    drop.classList.remove("drag");
    const f = e.dataTransfer?.files?.[0];
    if (f) handleImageFile(f);
  });

  // 全局粘贴：剪贴板里有图片就拦下来识别（不影响在输入框里粘贴文本）
  document.addEventListener("paste", (e) => {
    const items = e.clipboardData?.items;
    if (!items) return;
    for (const it of items) {
      if (it.type && it.type.startsWith("image/")) {
        const f = it.getAsFile();
        if (f) { e.preventDefault(); handleImageFile(f); return; }
      }
    }
  });
}

function wireUp(inWord) {
  bootMark("wireUp start");
  mathfield = $("mathfield");
  // 关闭虚拟键盘：聚焦时不再自动弹出（角标已用 CSS 隐藏）
  mathfield.mathVirtualKeyboardPolicy = "manual";
  latexEl = $("latex");
  statusEl = $("status");
  debugEl = $("debug");
  insertBtn = $("insertBtn");
  layoutModeEl = $("layoutMode");
  eqNumberEl = $("eqNumber");
  eqNumberRow = $("eqNumberRow");
  modeLine = $("modeLine");

  latexEl.value = mathfield.getValue("latex");
  mathfield.addEventListener("input", () => {
    latexEl.value = mathfield.getValue("latex");
    scheduleSync();
  });
  latexEl.addEventListener("input", () => {
    mathfield.setValue(latexEl.value, { suppressChangeNotifications: true });
    scheduleSync();
  });
  layoutModeEl.addEventListener("click", (ev) => {
    const btn = ev.target.closest("button[data-mode]");
    if (!btn) return;
    setLayout(btn.dataset.mode);
  });
  eqNumberEl.addEventListener("input", scheduleSync);

  $("palette").addEventListener("click", (ev) => {
    const btn = ev.target.closest("button[data-insert]");
    if (!btn) return;
    mathfield.executeCommand(["insert", btn.dataset.insert]);
    mathfield.focus();
    latexEl.value = mathfield.getValue("latex");
    scheduleSync();
  });

  $("clearBtn").addEventListener("click", () => {
    mathfield.setValue("");
    latexEl.value = "";
    setMode(false); // 清空即断开连接，回到新建
    mathfield.focus();
  });

  insertBtn.addEventListener("click", () => {
    clearTimeout(syncTimer);
    syncQueue.enqueue(() => insertNew()).catch((error) => conversionError(error, T("insertFail")));
  });
  insertBtn.disabled = !inWord;

  $("langBtn").addEventListener("click", () => setLang(lang === "zh" ? "en" : "zh"));

  wireAI();

  if (inWord) {
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      onSelectionChanged
    );
  }
  setMode(false);
  bootMark("wireUp complete");
  writeStartupDebug();
}

function waitForMathField() {
  const timeoutMs = 8000;
  const warning = setTimeout(() => {
    bootMark("math-field wait exceeded " + timeoutMs + " ms");
    setStartupStatus(T("mathSlow"), "err");
    writeStartupDebug("启动阶段仍在等待 <math-field> 注册。");
  }, timeoutMs);

  return customElements.whenDefined("math-field").finally(() => {
    clearTimeout(warning);
    bootMark("math-field defined");
  });
}

if (!window.Office || !Office.onReady) {
  bootMark("office.js unavailable");
  setStartupStatus(T("officeUnavailable"), "err");
  writeStartupDebug();
} else {
  const officeReadyWarning = setTimeout(() => {
    bootMark("office ready wait exceeded 8000 ms");
    setStartupStatus(T("officeWaiting"), "err");
    writeStartupDebug();
  }, 8000);

  Office.onReady((info) => {
    clearTimeout(officeReadyWarning);
    bootMark("office ready");
    setStartupStatus(T("loadingEditor"));
    return waitForMathField().then(() => start(info));
  });
}

function start(info) {
  const inWord = info.host === Office.HostType.Word;
  wireUp(inWord);
  if (inWord) setStatus(T("startReady"), "ok");
  else setStatus(T("startNonWord"));
}
