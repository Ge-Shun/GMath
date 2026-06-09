// GMath 弹窗编辑器（Dialog API 子页面）
// 职责：用 MathLive 可视化/LaTeX 编辑公式，提交时把结果回传给“父”（运行时 taskpane.js）。
// 通信：Office.context.ui.messageParent(JSON) —— 对话框只能用这一条通道回主程序。
//
// 初始内容、模式、是否显示公式，通过 URL 查询参数带入：
//   ?latex=<编码后的 LaTeX>&display=1&mode=edit

const $ = (id) => document.getElementById(id);

let mathfield, latexEl, statusEl, displayModeEl, okBtn, cancelBtn, titleEl;

function setStatus(msg, kind = "") {
  statusEl.textContent = msg || "";
  statusEl.className = "status" + (kind ? " " + kind : "");
}

function send(payload) {
  try {
    Office.context.ui.messageParent(JSON.stringify(payload));
  } catch (e) {
    setStatus("无法回传给主程序：" + (e.message || e), "err");
  }
}

function wireUp() {
  mathfield = $("mathfield");
  latexEl = $("latex");
  statusEl = $("status");
  displayModeEl = $("displayMode");
  okBtn = $("okBtn");
  cancelBtn = $("cancelBtn");
  titleEl = $("title");

  const params = new URLSearchParams(location.search);
  const initLatex = params.get("latex") || "";
  const isEdit = params.get("mode") === "edit";
  displayModeEl.checked = params.get("display") !== "0";

  titleEl.textContent = isEdit ? "编辑公式" : "插入公式";
  okBtn.textContent = isEdit ? "更新" : "插入";

  if (initLatex) mathfield.setValue(initLatex);
  latexEl.value = mathfield.getValue("latex");

  // MathLive ↔ LaTeX 双向同步
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

  okBtn.addEventListener("click", () => {
    const latex = mathfield.getValue("latex");
    if (!latex || !latex.trim()) {
      setStatus("公式为空。", "err");
      return;
    }
    send({ action: "submit", latex, display: displayModeEl.checked });
  });

  cancelBtn.addEventListener("click", () => send({ action: "cancel" }));

  mathfield.focus();
}

Office.onReady(() => {
  customElements.whenDefined("math-field").then(wireUp);
});
