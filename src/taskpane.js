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

const BUILD = "2026-06-10-a";

const $ = (id) => document.getElementById(id);
const boot = window.__gmathBoot || { start: performance.now(), marks: [] };
const bootMark = (name) => {
  boot.marks.push({ name, ms: Math.round(performance.now() - boot.start) });
};
const bootReport = () =>
  ["GMath 启动耗时:", ...boot.marks.map((m) => `- ${m.name}: ${m.ms} ms`)].join("\n");

let mathfield, latexEl, statusEl, debugEl, insertBtn, displayModeEl, modeLine;

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
    modeLine.textContent = "已连接文档中的公式 —— 在此修改会自动同步回去";
    modeLine.classList.add("linked");
  } else {
    loadedLatex = null;
    modeLine.textContent = "新建公式（点选文档里的公式可载入编辑）";
    modeLine.classList.remove("linked");
  }
}

function currentOmml() {
  const mathml = mathfield.getValue("math-ml");
  if (!mathml || !mathml.trim()) return null;
  return mml2omml(mathml);
}

function buildFlatOpc(ommlMath, display) {
  let oMath = ommlMath.replace(/^\s*<\?xml[^>]*\?>\s*/i, "").trim();
  const mathBlock = display
    ? `<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>${oMath}</m:oMathPara>`
    : oMath;
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

// 把一段 LaTeX 灌入面板（不触发回写）
function loadIntoPane(latex, display) {
  loadingDoc = true;
  mathfield.setValue(latex);
  latexEl.value = mathfield.getValue("latex");
  if (typeof display === "boolean") displayModeEl.checked = display;
  loadingDoc = false;
}

// ===== 面板 → 文档：插入一个新公式（不连接） =====
async function insertNew() {
  const omml = currentOmml();
  if (!omml) {
    setStatus("公式为空，请先输入内容。", "err");
    return;
  }
  const flatOpc = buildFlatOpc(omml, displayModeEl.checked);
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
    setStatus("已插入并连接该公式，之后在面板里改会自动同步。", "ok");
  } catch (e) {
    setStatus("插入失败：" + (e.code || "") + " " + (e.message || e), "err");
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
  const flatOpc = buildFlatOpc(omml, displayModeEl.checked);
  applying = true;
  try {
    await Word.run(async (context) => {
      const r = context.document.getSelection().insertOoxml(flatOpc, Word.InsertLocation.replace);
      r.select(); // 替换后重新选中，保证下一次还能替换到它
      await context.sync();
    });
    loadedLatex = mathfield.getValue("latex");
    setStatus("已同步到文档。", "ok");
  } catch (e) {
    setStatus("同步失败：" + (e.code || "") + " " + (e.message || e), "err");
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
    setStatus("已载入选中的公式，可直接修改（自动同步）。", "ok");
  } catch (e) {
    debugEl.value = "[选区读取] " + (e.message || e) + "\n" + debugEl.value;
  }
}

function wireUp(inWord) {
  bootMark("wireUp start");
  mathfield = $("mathfield");
  latexEl = $("latex");
  statusEl = $("status");
  debugEl = $("debug");
  insertBtn = $("insertBtn");
  displayModeEl = $("displayMode");
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
  displayModeEl.addEventListener("change", scheduleSync);

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
    setStartupStatus(
      "公式编辑器加载较慢，通常是 MathLive CDN 下载或注册耗时。请稍候；若长期不变，请检查网络或刷新任务面板。",
      "err"
    );
    writeStartupDebug("启动阶段仍在等待 <math-field> 注册。");
  }, timeoutMs);

  return customElements.whenDefined("math-field").finally(() => {
    clearTimeout(warning);
    bootMark("math-field defined");
  });
}

if (!window.Office || !Office.onReady) {
  bootMark("office.js unavailable");
  setStartupStatus("Office.js 未加载，无法初始化 Word 加载项。请检查网络后重新打开任务面板。", "err");
  writeStartupDebug();
} else {
  const officeReadyWarning = setTimeout(() => {
    bootMark("office ready wait exceeded 8000 ms");
    setStartupStatus("正在等待 Office 初始化，若长期不变通常是 Office.js 或 Word WebView 启动较慢。", "err");
    writeStartupDebug();
  }, 8000);

  Office.onReady((info) => {
    clearTimeout(officeReadyWarning);
    bootMark("office ready");
    setStartupStatus("正在加载公式编辑器…");
    return waitForMathField().then(() => start(info));
  });
}

function start(info) {
  const inWord = info.host === Office.HostType.Word;
  wireUp(inWord);
  if (inWord) setStatus("就绪（版本 " + BUILD + "）。新建公式，或点选文档里的公式来编辑。", "ok");
  else setStatus("（非 Word 环境）编辑器可用，但只有在 Word 中才能插入/同步。");
}
