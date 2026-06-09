// GMath Word 加载项 —— 核心逻辑
// 链路（正向）：MathLive 编辑 → MathML → mml2omml 转 OMML → Flat OPC → insertOoxml 插入 Word
// 链路（反向）：文档里的公式 → getOoxml → extractOMath → omml2latex → 载回 MathLive 面板
//
// MathType 式体验：插入的公式用一个 tag="GMath" 的内容控件包住（带框）。
// 在文档里点选它 → 选区监听读回其当前 OMML → 反推成 LaTeX → 填回面板 → 改完「更新」原地替换。

import { mml2omml } from "./mathml2omml.js";
import { omml2latex, extractOMath } from "./omml2latex.js";

const BUILD = "2026-06-09-f"; // 版本标记：确认面板加载的是不是最新代码
const CC_TAG = "GMath";

const $ = (id) => document.getElementById(id);

let mathfield, latexEl, statusEl, debugEl, displayModeEl;
let insertBtn, newBtn, importBtn, unwrapBtn, clearBtn, modeLine;

// 编辑态：当前面板正绑定到文档里哪个 GMath 内容控件
let activeCcId = null;
// 选区监听自我触发保护：我们自己插入/更新/解包时会改变选区，需临时屏蔽
let suppressSelection = false;
let selTimer = null;

function setStatus(msg, kind = "") {
  statusEl.textContent = msg;
  statusEl.className = "status" + (kind ? " " + kind : "");
}

// 切换「新建 / 编辑中」两态的 UI
function setMode(editing) {
  if (editing) {
    modeLine.textContent = "正在编辑文档中的公式（改完点「更新选中公式」）";
    modeLine.classList.add("editing");
    insertBtn.textContent = "更新选中公式";
    newBtn.hidden = false;
    unwrapBtn.hidden = false;
  } else {
    activeCcId = null;
    modeLine.textContent = "新建公式";
    modeLine.classList.remove("editing");
    insertBtn.textContent = "插入到 Word";
    newBtn.hidden = true;
    unwrapBtn.hidden = true;
  }
}

// 当前面板内容 → OMML（<m:oMath>…</m:oMath>）
function currentOmml() {
  const mathml = mathfield.getValue("math-ml");
  if (!mathml || !mathml.trim()) return null;
  return mml2omml(mathml);
}

// 把 OMML 包成 insertOoxml 需要的 Flat OPC（扁平 OPC 包）
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

function describeError(e) {
  return { code: e.code, message: e.message, debugInfo: e.debugInfo };
}

// 插入一个新的 GMath 公式（包内容控件、打标），并进入编辑态绑定到它
async function insertNew() {
  const omml = currentOmml();
  if (!omml) {
    setStatus("公式为空，请先输入内容。", "err");
    return;
  }
  const flatOpc = buildFlatOpc(omml, displayModeEl.checked);
  debugEl.value = "构建版本: " + BUILD + "\n\nOMML:\n" + omml + "\n\n完整 OOXML:\n" + flatOpc;

  suppressSelection = true;
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      const inserted = range.insertOoxml(flatOpc, Word.InsertLocation.replace);
      const cc = inserted.insertContentControl();
      cc.tag = CC_TAG;
      cc.title = "GMath 公式";
      cc.appearance = Word.ContentControlAppearance.boundingBox;
      cc.load("id");
      await context.sync();
      activeCcId = cc.id;
    });
    setMode(true);
    setStatus("已插入为 GMath 可编辑公式（点选它可回此面板再编辑）。", "ok");
  } catch (e) {
    setStatus("插入失败：" + (e.code || "") + " " + (e.message || e), "err");
    debugEl.value += "\n\n=== 错误详情 ===\n" + JSON.stringify(describeError(e), null, 2);
  } finally {
    suppressSelection = false;
  }
}

// 更新当前绑定的 GMath 公式（原地替换内容控件内容）
async function updateActive() {
  if (activeCcId == null) return insertNew();
  const omml = currentOmml();
  if (!omml) {
    setStatus("公式为空，请先输入内容。", "err");
    return;
  }
  const flatOpc = buildFlatOpc(omml, displayModeEl.checked);

  suppressSelection = true;
  try {
    await Word.run(async (context) => {
      const cc = context.document.contentControls.getById(activeCcId);
      cc.insertOoxml(flatOpc, Word.InsertLocation.replace);
      // 确保标记仍在（替换内容不应丢标，这里兜底重设）
      cc.tag = CC_TAG;
      cc.appearance = Word.ContentControlAppearance.boundingBox;
      await context.sync();
    });
    setStatus("已更新文档中的公式。", "ok");
  } catch (e) {
    setStatus("更新失败：" + (e.code || "") + " " + (e.message || e) + "（公式可能已被删除，可改用「插入为新公式」）", "err");
    debugEl.value += "\n\n=== 错误详情 ===\n" + JSON.stringify(describeError(e), null, 2);
  } finally {
    suppressSelection = false;
  }
}

// 主按钮：编辑态→更新，新建态→插入
function mainAction() {
  if (activeCcId != null) updateActive();
  else insertNew();
}

// 导入「选中的 Word 公式」到面板（支持任意原生公式或 GMath 公式）
async function importSelected() {
  try {
    let ooxml = null;
    let ccId = null;
    let ccTag = null;
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      const cc = sel.parentContentControlOrNullObject;
      cc.load("id,tag,isNullObject");
      const res = sel.getOoxml();
      await context.sync();
      ooxml = res.value;
      if (!cc.isNullObject) {
        ccId = cc.id;
        ccTag = cc.tag;
      }
    });
    const omath = extractOMath(ooxml);
    if (!omath) {
      setStatus("选区里没有找到公式。请先选中一个公式再导入。", "err");
      return;
    }
    const latex = omml2latex(omath);
    mathfield.setValue(latex);
    latexEl.value = mathfield.getValue("latex");
    displayModeEl.checked = /oMathPara/.test(omath);

    if (ccTag === CC_TAG && ccId != null) {
      // 选中的本就是 GMath 公式 → 进入编辑态，后续可「更新」
      activeCcId = ccId;
      setMode(true);
      setStatus("已载入选中的 GMath 公式。", "ok");
    } else {
      // 原生公式 → 作为新公式编辑；点「插入到 Word」会生成一个 GMath 版本
      setMode(false);
      setStatus("已导入选中的 Word 公式。点「插入到 Word」可生成 GMath 可编辑版本。", "ok");
    }
  } catch (e) {
    setStatus("导入失败：" + (e.code || "") + " " + (e.message || e), "err");
    debugEl.value += "\n\n=== 错误详情 ===\n" + JSON.stringify(describeError(e), null, 2);
  }
}

// 把当前 GMath 公式转成 Word 原生公式（去掉内容控件、保留公式本身）
async function unwrapToNative() {
  if (activeCcId == null) return;
  suppressSelection = true;
  try {
    await Word.run(async (context) => {
      const cc = context.document.contentControls.getById(activeCcId);
      cc.delete(true); // keepContent=true：保留公式，仅移除控件包装
      await context.sync();
    });
    setMode(false);
    setStatus("已转为 Word 原生公式（可用 Word 自带公式工具编辑）。", "ok");
  } catch (e) {
    setStatus("转换失败：" + (e.code || "") + " " + (e.message || e), "err");
  } finally {
    suppressSelection = false;
  }
}

// 选区变化 → 若进入某个 GMath 公式，自动把它载回面板
function onSelectionChanged() {
  if (suppressSelection) return;
  clearTimeout(selTimer);
  selTimer = setTimeout(loadFromSelection, 200);
}

async function loadFromSelection() {
  if (suppressSelection) return;
  try {
    let hit = null; // { id, ooxml }
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      const cc = sel.parentContentControlOrNullObject;
      cc.load("id,tag,isNullObject");
      await context.sync();
      if (cc.isNullObject || cc.tag !== CC_TAG) return;
      if (cc.id === activeCcId) return; // 已在编辑同一个，别覆盖正在改的内容
      const res = cc.getOoxml();
      await context.sync();
      hit = { id: cc.id, ooxml: res.value };
    });
    if (!hit) return;
    const omath = extractOMath(hit.ooxml);
    if (!omath) return;
    const latex = omml2latex(omath);
    mathfield.setValue(latex);
    latexEl.value = mathfield.getValue("latex");
    displayModeEl.checked = /oMathPara/.test(omath);
    activeCcId = hit.id;
    setMode(true);
    setStatus("已载入选中的 GMath 公式，可编辑后点「更新选中公式」。", "ok");
  } catch (e) {
    // 选区监听是后台行为，失败不打扰用户，仅记录到调试框
    debugEl.value = "[选区载入] " + (e.message || e) + "\n" + debugEl.value;
  }
}

function wireUp(inWord) {
  mathfield = $("mathfield");
  latexEl = $("latex");
  statusEl = $("status");
  debugEl = $("debug");
  displayModeEl = $("displayMode");
  insertBtn = $("insertBtn");
  newBtn = $("newBtn");
  importBtn = $("importBtn");
  unwrapBtn = $("unwrapBtn");
  clearBtn = $("clearBtn");
  modeLine = $("modeLine");

  // MathLive ↔ LaTeX 文本框双向同步
  latexEl.value = mathfield.getValue("latex");
  mathfield.addEventListener("input", () => {
    latexEl.value = mathfield.getValue("latex");
  });
  latexEl.addEventListener("input", () => {
    mathfield.setValue(latexEl.value, { suppressChangeNotifications: true });
  });

  // 快捷符号面板
  $("palette").addEventListener("click", (ev) => {
    const btn = ev.target.closest("button[data-insert]");
    if (!btn) return;
    mathfield.executeCommand(["insert", btn.dataset.insert]);
    mathfield.focus();
    latexEl.value = mathfield.getValue("latex");
  });

  clearBtn.addEventListener("click", () => {
    mathfield.setValue("");
    latexEl.value = "";
    setMode(false); // 清空即退出编辑态，回到新建
    mathfield.focus();
  });

  insertBtn.addEventListener("click", mainAction);
  newBtn.addEventListener("click", insertNew);
  importBtn.addEventListener("click", importSelected);
  unwrapBtn.addEventListener("click", unwrapToNative);
  insertBtn.disabled = false;

  if (inWord) {
    // 监听文档选区变化，实现「点选公式即载入面板」
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      onSelectionChanged
    );
  } else {
    importBtn.disabled = true;
    unwrapBtn.disabled = true;
  }
  setMode(false);
}

// 等待 Office 与 MathLive 都就绪
Office.onReady((info) => {
  const inWord = info.host === Office.HostType.Word;
  customElements.whenDefined("math-field").then(() => {
    wireUp(inWord);
    if (inWord) setStatus("就绪（版本 " + BUILD + "）。编辑公式后点「插入到 Word」。", "ok");
    else setStatus("（非 Word 环境）编辑器可用，但只有在 Word 中才能插入/导入。");
  });
});
