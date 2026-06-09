// GMath Word 加载项 —— 运行时核心
// 形态（最接近 MathType）：共享运行时常驻 + 弹窗编辑器
//   - 在文档里点选一个 GMath 公式 → 弹出独立编辑窗口（Dialog），载入其 LaTeX
//   - 改完点「更新」→ 窗口回传 LaTeX → 这里转成 OMML 原地替换该公式
//   - 面板自身用于「新建公式」与若干工具按钮
//
// 正向：MathLive → MathML → mml2omml → Flat OPC → insertOoxml
// 反向：文档公式 → getOoxml → extractOMath → omml2latex → 弹窗里的 LaTeX

import { mml2omml } from "./mathml2omml.js";
import { omml2latex, extractOMath } from "./omml2latex.js";

const BUILD = "2026-06-09-g";
const CC_TAG = "GMath";

const $ = (id) => document.getElementById(id);

let mathfield, latexEl, statusEl, debugEl, displayModeEl;
let insertBtn, importBtn, unwrapBtn, clearBtn, modeLine;
let inWord = false;

// 弹窗与选区监听状态
let activeDialog = null;
let dialogOpen = false;
let editingCcId = null; // 弹窗当前编辑的内容控件 id（新建时为 null）
let suppressUntil = 0; // 在此时间戳前忽略选区事件（防止插入/更新后自我触发）
let selTimer = null;

function setStatus(msg, kind = "") {
  statusEl.textContent = msg;
  statusEl.className = "status" + (kind ? " " + kind : "");
}

const appearanceBoundingBox = () => Word.ContentControlAppearance.boundingBox;

// 面板内容 → OMML
function paneOmml() {
  const mathml = mathfield.getValue("math-ml");
  if (!mathml || !mathml.trim()) return null;
  return mml2omml(mathml);
}

// LaTeX → OMML（借面板里的 mathfield 做 LaTeX→MathML）
function latexToOmml(latex) {
  mathfield.setValue(latex);
  const mathml = mathfield.getValue("math-ml");
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

// 插入一个新的 GMath 公式（带框内容控件），返回 cc id
async function insertNewEquation(omml, display) {
  const flatOpc = buildFlatOpc(omml, display);
  let newId = null;
  suppressUntil = Date.now() + 1500;
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    const inserted = range.insertOoxml(flatOpc, Word.InsertLocation.replace);
    const cc = inserted.insertContentControl();
    cc.tag = CC_TAG;
    cc.title = "GMath 公式";
    cc.appearance = appearanceBoundingBox();
    cc.load("id");
    await context.sync();
    newId = cc.id;
  });
  return newId;
}

// 原地更新某个 GMath 内容控件
async function updateEquation(ccId, omml, display) {
  const flatOpc = buildFlatOpc(omml, display);
  suppressUntil = Date.now() + 1500;
  await Word.run(async (context) => {
    const cc = context.document.contentControls.getById(ccId);
    cc.insertOoxml(flatOpc, Word.InsertLocation.replace);
    cc.tag = CC_TAG;
    cc.appearance = appearanceBoundingBox();
    await context.sync();
  });
}

// ===== 面板：新建公式 =====
async function insertFromPane() {
  const omml = paneOmml();
  if (!omml) {
    setStatus("公式为空，请先输入内容。", "err");
    return;
  }
  try {
    await insertNewEquation(omml, displayModeEl.checked);
    setStatus("已插入为 GMath 可编辑公式（点选它会弹出编辑窗口）。", "ok");
  } catch (e) {
    setStatus("插入失败：" + (e.code || "") + " " + (e.message || e), "err");
    debugEl.value = "OMML:\n" + omml + "\n\n=== 错误详情 ===\n" + JSON.stringify(describeError(e), null, 2);
  }
}

// ===== 面板：导入选中公式到面板（用于复用/另存为新） =====
async function importSelected() {
  try {
    let ooxml = null;
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      const res = sel.getOoxml();
      await context.sync();
      ooxml = res.value;
    });
    const omath = extractOMath(ooxml);
    if (!omath) {
      setStatus("选区里没有找到公式。请先选中一个公式。", "err");
      return;
    }
    const latex = omml2latex(omath);
    mathfield.setValue(latex);
    latexEl.value = mathfield.getValue("latex");
    displayModeEl.checked = /oMathPara/.test(omath);
    setStatus("已把选中公式载入面板。点「插入到 Word」可生成新的 GMath 公式。", "ok");
  } catch (e) {
    setStatus("导入失败：" + (e.code || "") + " " + (e.message || e), "err");
  }
}

// ===== 面板：把选中的 GMath 公式转为 Word 原生公式（去掉控件包装） =====
async function unwrapSelected() {
  try {
    let done = false;
    suppressUntil = Date.now() + 1500;
    await Word.run(async (context) => {
      const cc = context.document.getSelection().parentContentControlOrNullObject;
      cc.load("tag,isNullObject");
      await context.sync();
      if (cc.isNullObject || cc.tag !== CC_TAG) return;
      cc.delete(true); // keepContent=true：保留公式，仅移除控件
      done = true;
      await context.sync();
    });
    setStatus(
      done ? "已转为 Word 原生公式（可用 Word 自带公式工具编辑）。" : "请先在文档里选中一个 GMath 公式。",
      done ? "ok" : "err"
    );
  } catch (e) {
    setStatus("转换失败：" + (e.code || "") + " " + (e.message || e), "err");
  }
}

// ===== 弹窗编辑器 =====
function dialogUrl(latex, display, mode) {
  const url = new URL("./dialog.html", location.href);
  url.searchParams.set("latex", latex || "");
  url.searchParams.set("display", display ? "1" : "0");
  url.searchParams.set("mode", mode);
  return url.href;
}

function openDialog(latex, display, mode, ccId) {
  if (dialogOpen) return;
  dialogOpen = true;
  editingCcId = ccId ?? null;
  Office.context.ui.displayDialogAsync(
    dialogUrl(latex, display, mode),
    { height: 58, width: 46, promptBeforeOpen: false },
    (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        dialogOpen = false;
        setStatus("无法打开编辑窗口：" + (res.error && res.error.message), "err");
        return;
      }
      activeDialog = res.value;
      activeDialog.addEventHandler(Office.EventType.DialogMessageReceived, onDialogMessage);
      activeDialog.addEventHandler(Office.EventType.DialogEventReceived, onDialogEvent);
    }
  );
}

function closeDialog() {
  try {
    if (activeDialog) activeDialog.close();
  } catch {}
  activeDialog = null;
  dialogOpen = false;
  suppressUntil = Date.now() + 1000; // 关闭后冷却，避免立刻又被选区触发
}

async function onDialogMessage(arg) {
  let msg;
  try {
    msg = JSON.parse(arg.message);
  } catch {
    return;
  }
  if (msg.action === "cancel") {
    closeDialog();
    return;
  }
  if (msg.action === "submit") {
    const ccId = editingCcId;
    closeDialog();
    try {
      const omml = latexToOmml(msg.latex);
      if (ccId != null) {
        await updateEquation(ccId, omml, msg.display);
        setStatus("已更新文档中的公式。", "ok");
      } else {
        await insertNewEquation(omml, msg.display);
        setStatus("已插入公式。", "ok");
      }
    } catch (e) {
      setStatus("应用公式失败：" + (e.code || "") + " " + (e.message || e), "err");
      debugEl.value = "=== 错误详情 ===\n" + JSON.stringify(describeError(e), null, 2);
    }
  }
}

// 用户手动关闭弹窗（点 X、导航等）
function onDialogEvent() {
  activeDialog = null;
  dialogOpen = false;
  suppressUntil = Date.now() + 1000;
}

// ===== 选区监听：点选 GMath 公式 → 弹出编辑窗口 =====
function onSelectionChanged() {
  if (dialogOpen || Date.now() < suppressUntil) return;
  clearTimeout(selTimer);
  selTimer = setTimeout(maybeOpenEditor, 180);
}

async function maybeOpenEditor() {
  if (dialogOpen || Date.now() < suppressUntil) return;
  try {
    let hit = null;
    await Word.run(async (context) => {
      const cc = context.document.getSelection().parentContentControlOrNullObject;
      cc.load("id,tag,isNullObject");
      await context.sync();
      if (cc.isNullObject || cc.tag !== CC_TAG) return;
      const res = cc.getOoxml();
      await context.sync();
      hit = { id: cc.id, ooxml: res.value };
    });
    if (!hit) return;
    const omath = extractOMath(hit.ooxml);
    if (!omath) return;
    const latex = omml2latex(omath);
    const display = /oMathPara/.test(omath);
    openDialog(latex, display, "edit", hit.id);
  } catch (e) {
    debugEl.value = "[选区监听] " + (e.message || e) + "\n" + debugEl.value;
  }
}

function wireUp() {
  mathfield = $("mathfield");
  latexEl = $("latex");
  statusEl = $("status");
  debugEl = $("debug");
  displayModeEl = $("displayMode");
  insertBtn = $("insertBtn");
  importBtn = $("importBtn");
  unwrapBtn = $("unwrapBtn");
  clearBtn = $("clearBtn");
  modeLine = $("modeLine");

  modeLine.textContent = "新建公式（点选文档里的 GMath 公式会弹出编辑窗口）";

  latexEl.value = mathfield.getValue("latex");
  mathfield.addEventListener("input", () => {
    latexEl.value = mathfield.getValue("latex");
  });
  latexEl.addEventListener("input", () => {
    mathfield.setValue(latexEl.value, { suppressChangeNotifications: true });
  });

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
    mathfield.focus();
  });

  insertBtn.addEventListener("click", insertFromPane);
  importBtn.addEventListener("click", importSelected);
  unwrapBtn.addEventListener("click", unwrapSelected);
  insertBtn.disabled = false;

  if (inWord) {
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      onSelectionChanged
    );
  } else {
    importBtn.disabled = true;
    unwrapBtn.disabled = true;
  }
}

Office.onReady((info) => {
  inWord = info.host === Office.HostType.Word;
  customElements.whenDefined("math-field").then(() => {
    wireUp();
    if (inWord) setStatus("就绪（版本 " + BUILD + "）。可新建公式；点选文档里的公式会弹出编辑窗口。", "ok");
    else setStatus("（非 Word 环境）编辑器可用，但只有在 Word 中才能插入/编辑。");
  });
});
