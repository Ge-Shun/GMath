// GMath Word 加载项 —— 核心逻辑
// 链路：MathLive 编辑 → 输出 MathML → mml2omml 转 OMML → 包成 Flat OPC → insertOoxml 插入 Word

import { mml2omml } from "./mathml2omml.js"; // 自研转换器，无第三方依赖

const BUILD = "2026-06-09-e"; // 版本标记：用于确认面板加载的是不是最新代码

const $ = (id) => document.getElementById(id);

let mathfield, latexEl, statusEl, debugEl, insertBtn, displayModeEl;

function setStatus(msg, kind = "") {
  statusEl.textContent = msg;
  statusEl.className = "status" + (kind ? " " + kind : "");
}

// 把 OMML 片段包装成 Office.js insertOoxml 需要的 Flat OPC（扁平 OPC 包）字符串
function buildFlatOpc(ommlMath, display) {
  // mml2omml 输出的是 <m:oMath>...</m:oMath>，去掉可能的 XML 声明
  let oMath = ommlMath.replace(/^\s*<\?xml[^>]*\?>\s*/i, "").trim();

  // 块级公式用 m:oMathPara 包裹（独立成行、默认居中）；行内公式直接放 oMath
  const mathBlock = display
    ? `<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>${oMath}</m:oMathPara>`
    : oMath;

  // 注意：pkg:xmlData 内的各部件不能再带 <?xml?> 声明，
  // 声明只允许出现在整个包的最开头，否则 insertOoxml 会抛 GeneralException。
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

async function insertEquation() {
  const mathml = mathfield.getValue("math-ml");
  if (!mathml || !mathml.trim()) {
    setStatus("公式为空，请先输入内容。", "err");
    return;
  }

  let omml;
  try {
    omml = mml2omml(mathml);
  } catch (e) {
    setStatus("MathML→OMML 转换失败：" + e.message, "err");
    debugEl.value = "MathML:\n" + mathml;
    return;
  }

  const display = displayModeEl.checked;
  const flatOpc = buildFlatOpc(omml, display);

  debugEl.value =
    "构建版本: " + BUILD + "\n\nMathML:\n" + mathml + "\n\nOMML:\n" + omml + "\n\n完整 OOXML:\n" + flatOpc;

  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.insertOoxml(flatOpc, Word.InsertLocation.replace);
      await context.sync();
    });
    setStatus("已插入到 Word（双击可继续编辑）。", "ok");
  } catch (e) {
    // Office 的错误对象带 code / debugInfo，比 message 详细得多
    const detail = {
      code: e.code,
      message: e.message,
      debugInfo: e.debugInfo,
    };
    setStatus("插入失败：" + (e.code || "") + " " + (e.message || e), "err");
    debugEl.value += "\n\n=== 错误详情 ===\n" + JSON.stringify(detail, null, 2);
  }
}

function wireUp() {
  mathfield = $("mathfield");
  latexEl = $("latex");
  statusEl = $("status");
  debugEl = $("debug");
  insertBtn = $("insertBtn");
  displayModeEl = $("displayMode");

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

  $("clearBtn").addEventListener("click", () => {
    mathfield.setValue("");
    latexEl.value = "";
    mathfield.focus();
  });

  insertBtn.addEventListener("click", insertEquation);
  insertBtn.disabled = false;
}

// 等待 Office 与 MathLive 都就绪
Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    // 仍允许在浏览器里预览编辑器，只是无法插入
    customElements.whenDefined("math-field").then(() => {
      wireUp();
      setStatus("（非 Word 环境）编辑器可用，但只有在 Word 中才能插入。");
    });
    return;
  }
  customElements.whenDefined("math-field").then(() => {
    wireUp();
    setStatus("就绪（版本 " + BUILD + "）。编辑公式后点击「插入到 Word」。", "ok");
  });
});
