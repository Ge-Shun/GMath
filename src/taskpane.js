// GMath Word 加载项 —— 核心逻辑（纯原生公式 + 双向实时同步，尽力而为）
// 正向：MathLive 编辑 → MathML → mml2omml 转 OMML → Flat OPC → insertOoxml
// 反向：点选文档里的公式 → getOoxml → extractOMath → omml2latex → 载入面板
// 同步：面板里改动 → 防抖 ~0.5s → 替换“当前选中的那个公式”，并重新选中以便继续同步
//
// 注意（纯原生、无锚点的固有限制）：回写靠“替换当前选区”。
//   想可靠同步，请把整个公式选中（拖选或在公式上点一下再用方向键/三击选中），
//   光标只是停在公式里时，替换可能变成插入。详见 README。

import { mml2omml } from "./mathml2omml.js";
import { omml2latex, extractOMath } from "./omml2latex.js";
import { latexToOmml, hasDecoration } from "./latex2omml.js";
import { SYMBOL_CATEGORIES, CAT_EN } from "./symbols.js";
import { I18N, TIP_EN } from "./i18n.js";

const BUILD = "2026-06-12-common-desc-cache-bust";

// XML 文本转义（用于把用户输入的编号安全嵌入 OOXML）
const escXml = (s) =>
  String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

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
    loadedLatex = null;
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
      mml2omml,
    });
    return omml && omml.trim() ? omml : null;
  }
  const mathml = mathfield.getValue("math-ml");
  if (!mathml || !mathml.trim()) return null;
  return mml2omml(mathml);
}

// 生成右编号的文字 run：留空→自动编号（Word SEQ 域，会自动续号）；否则用字面文本
function buildNumberRuns(numberText) {
  const t = (numberText || "").trim();
  if (t) {
    return `<w:r><w:t xml:space="preserve">${escXml(t)}</w:t></w:r>`;
  }
  // ( SEQ equation \* ARABIC ) → (1) (2) …，由 Word 维护并可整篇续号
  return (
    `<w:r><w:t xml:space="preserve">(</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> SEQ equation \\* ARABIC </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t>1</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
    `<w:r><w:t xml:space="preserve">)</w:t></w:r>`
  );
}

// layout: "inline"（行内）| "display"（行间居中）| "numbered"（居中＋右侧编号）
function buildFlatOpc(ommlMath, layout, numberText) {
  let oMath = ommlMath.replace(/^\s*<\?xml[^>]*\?>\s*/i, "").trim();
  let mathBlock;
  if (layout === "display") {
    mathBlock = `<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>${oMath}</m:oMathPara>`;
  } else if (layout === "numbered") {
    // 用「相对页边距的定位制表符」实现：公式居中、编号靠右贴右边距，与纸张/页宽无关
    mathBlock =
      `<w:r><w:ptab w:relativeTo="margin" w:alignment="center" w:leader="none"/></w:r>` +
      oMath +
      `<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>` +
      buildNumberRuns(numberText);
  } else {
    mathBlock = oMath; // 行内
  }
  const documentXml =
    `<w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">` +
    `<w:body><w:p>${mathBlock}</w:p></w:body>` +
    `</w:document>`;
  const relsXml =
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" ` +
    `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ` +
    `Target="word/document.xml"/>` +
    `</Relationships>`;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<?mso-application progid="Word.Document"?>` +
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    `<pkg:part pkg:name="/_rels/.rels" ` +
    `pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">` +
    `<pkg:xmlData>${relsXml}</pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" ` +
    `pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">` +
    `<pkg:xmlData>${documentXml}</pkg:xmlData></pkg:part>` +
    `</pkg:package>`
  );
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
  });
  eqNumberRow.hidden = mode !== "numbered"; // 编号输入只在「右编号」时出现
  scheduleSync();
}

// 把一段 LaTeX 灌入面板（不触发回写）
function loadIntoPane(latex, display) {
  loadingDoc = true;
  mathfield.setValue(latex);
  latexEl.value = mathfield.getValue("latex");
  if (typeof display === "boolean") setLayout(display ? "display" : "inline");
  loadingDoc = false;
}

// ===== 面板 → 文档：插入一个新公式（不连接） =====
async function insertNew() {
  const omml = currentOmml();
  if (!omml) {
    setStatus(T("emptyFormula"), "err");
    return;
  }
  const flatOpc = buildFlatOpc(omml, getLayout(), eqNumberEl.value);
  debugEl.value = "构建版本: " + BUILD + "\n\nOMML:\n" + omml + "\n\n完整 OOXML:\n" + flatOpc;
  applying = true;
  try {
    await Word.run(async (context) => {
      const r = context.document.getSelection().insertOoxml(flatOpc, Word.InsertLocation.replace);
      r.select(); // 选中新公式，便于随后继续在面板里改、自动同步
      await context.sync();
    });
    loadedLatex = mathfield.getValue("latex");
    setMode(true);
    setStatus(T("insertedLinked"), "ok");
  } catch (e) {
    setStatus(T("insertFail") + (e.code || "") + " " + (e.message || e), "err");
    debugEl.value += "\n\n=== 错误详情 ===\n" + JSON.stringify(describeError(e), null, 2);
  } finally {
    setTimeout(() => (applying = false), 500);
  }
}

// ===== 面板 → 文档：把改动同步回当前选中的公式（替换选区） =====
async function syncToDoc() {
  if (!linked) return;
  const omml = currentOmml();
  if (!omml) return;
  const flatOpc = buildFlatOpc(omml, getLayout(), eqNumberEl.value);
  applying = true;
  try {
    await Word.run(async (context) => {
      const r = context.document.getSelection().insertOoxml(flatOpc, Word.InsertLocation.replace);
      r.select(); // 替换后重新选中，保证下一次还能替换到它
      await context.sync();
    });
    loadedLatex = mathfield.getValue("latex");
    setStatus(T("syncedDoc"), "ok");
  } catch (e) {
    setStatus(T("syncFail") + (e.code || "") + " " + (e.message || e), "err");
  } finally {
    setTimeout(() => (applying = false), 500);
  }
}

function scheduleSync() {
  if (!linked || loadingDoc) return;
  clearTimeout(syncTimer);
  syncTimer = setTimeout(syncToDoc, 600);
}

// ===== 文档 → 面板：选区变化时，若选中的是公式就载入面板 =====
function onSelectionChanged() {
  if (applying) return;
  clearTimeout(selTimer);
  selTimer = setTimeout(readFromSelection, 250);
}

async function readFromSelection() {
  if (applying) return;
  try {
    let ooxml = null;
    await Word.run(async (context) => {
      const res = context.document.getSelection().getOoxml();
      await context.sync();
      ooxml = res.value;
    });
    const omath = extractOMath(ooxml);
    if (!omath) {
      if (linked) setMode(false); // 选区离开了公式 → 断开，避免误同步
      return;
    }
    const latex = omml2latex(omath);
    if (linked && latex === loadedLatex) return; // 还是同一个公式，别覆盖正在改的内容
    loadIntoPane(latex, /oMathPara/.test(omath));
    loadedLatex = latex;
    setMode(true);
    setStatus(T("loadedSel"), "ok");
  } catch (e) {
    debugEl.value = "[选区读取] " + (e.message || e) + "\n" + debugEl.value;
  }
}

// ===== 图片转公式（AI 识别，OpenAI 兼容视觉接口） =====
const AI_KEYS = { url: "gmath.ai.url", key: "gmath.ai.key", model: "gmath.ai.model" };

function lsGet(k) {
  try { return localStorage.getItem(k) || ""; } catch { return ""; }
}
function lsSet(k, v) {
  try { localStorage.setItem(k, v); } catch { /* 某些 WebView 禁用 localStorage */ }
}

function setAiStatus(msg, kind = "") {
  const el = $("aiStatus");
  if (!el) return;
  el.hidden = !msg;
  el.textContent = msg || "";
  el.className = "ai-status" + (kind ? " " + kind : "");
}

// 把接口地址规整成 chat/completions 端点
function normalizeEndpoint(url) {
  let u = url.trim().replace(/\/+$/, "");
  if (/\/chat\/completions$/.test(u)) return u;
  if (/\/v\d+$/.test(u)) return u + "/chat/completions";
  return u + "/v1/chat/completions";
}

function aiRequestUrl() {
  if (["localhost", "127.0.0.1", "[::1]"].includes(window.location.hostname)) {
    return "/api/ai/chat/completions";
  }
  return "https://localhost:3000/api/ai/chat/completions";
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

async function postAi(endpoint, key, body, useProxy) {
  const headers = {
    "Content-Type": "application/json",
    Authorization: "Bearer " + key,
  };
  if (useProxy) headers["X-GMath-AI-Endpoint"] = endpoint;
  return fetch(useProxy ? aiRequestUrl() : endpoint, {
    method: "POST",
    headers,
    body,
  });
}

// 清洗模型返回：去掉代码块围栏与 $ / \[ \] / \( \) 定界符
function cleanLatex(raw) {
  let s = (raw || "").trim();
  s = s.replace(/^```(?:latex|tex|math)?\s*/i, "").replace(/\s*```$/, "").trim();
  s = s
    .replace(/^\$\$([\s\S]*?)\$\$$/, "$1")
    .replace(/^\\\[([\s\S]*?)\\\]$/, "$1")
    .replace(/^\\\(([\s\S]*?)\\\)$/, "$1")
    .replace(/^\$([\s\S]*?)\$$/, "$1")
    .trim();
  return s;
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
    return;
  }
  setAiStatus(T("aiBusy"), "busy");
  try {
    const endpoint = normalizeEndpoint(url);
    const body = aiPayload(model, dataUrl);
    let resp;
    try {
      resp = await postAi(endpoint, key, body, false);
    } catch (directErr) {
      resp = await postAi(endpoint, key, body, true).catch(() => {
        throw directErr;
      });
    }
    if (!resp.ok) {
      const body = await resp.text().catch(() => "");
      setAiStatus(`${T("aiHttp")}${resp.status}: ${body.slice(0, 200)}`, "err");
      return;
    }
    const data = await resp.json();
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
    setAiStatus(`${T("aiReqFail")}${e.message || e}${T("aiReqFailHint")}`, "err");
  }
}

async function handleImageFile(file) {
  if (!file || !file.type || !file.type.startsWith("image/")) return;
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
  const drop = $("aiDrop");
  const fileEl = $("aiFile");

  // 载入已保存的配置
  urlEl.value = lsGet(AI_KEYS.url);
  keyEl.value = lsGet(AI_KEYS.key);
  modelEl.value = lsGet(AI_KEYS.model);

  $("aiSettingsBtn").addEventListener("click", () => {
    const s = $("aiSettings");
    s.hidden = !s.hidden;
  });
  $("aiSaveBtn").addEventListener("click", () => {
    lsSet(AI_KEYS.url, urlEl.value.trim());
    lsSet(AI_KEYS.key, keyEl.value.trim());
    lsSet(AI_KEYS.model, modelEl.value.trim());
    setAiStatus(T("aiSaved"), "ok");
    $("aiSettings").hidden = true;
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

  insertBtn.addEventListener("click", insertNew);
  insertBtn.disabled = false;

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
