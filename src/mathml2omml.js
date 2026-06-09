// 自研 MathML → OMML 转换器（不依赖任何第三方库）
// 原理：用浏览器原生 DOMParser 把 MathML 解析成 DOM 树，再递归地把
//       每种 MathML 标签翻译成对应的 OMML（m: 命名空间）标签。
//
// OMML 速查：
//   文本     <m:r><m:t>…</m:t></m:r>
//   分数     <m:f><m:num>…</m:num><m:den>…</m:den></m:f>
//   根号     <m:rad><m:deg>…</m:deg><m:e>…</m:e></m:rad>
//   上/下标  <m:sSup|sSub|sSubSup><m:e>底</m:e><m:sup|sub>…</m:sup|sub></…>
//   求和积分 <m:nary><m:naryPr><m:chr m:val="∑"/>…</m:naryPr><m:sub/><m:sup/><m:e/></m:nary>
//   括号     <m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>…</m:e></m:d>
//   矩阵     <m:m><m:mr><m:e>…</m:e>…</m:mr>…</m:m>

const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

// 作为 n-ary（大型运算符，下/上标是上下限）处理的字符
const NARY = new Set(["∑", "∏", "∐", "∫", "∬", "∭", "∮", "∯", "∰", "⋃", "⋂", "⋁", "⋀", "⨁", "⨂", "⨀"]);

const localName = (n) => n.localName || n.tagName.replace(/^.*:/, "");
const esc = (s) =>
  String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

// 仅取元素子节点（忽略空白文本节点）
const kidsOf = (node) => Array.from(node.childNodes).filter((n) => n.nodeType === 1);

// 一个文本 run
const run = (text) => `<m:r><m:t xml:space="preserve">${esc(text)}</m:t></m:r>`;

// 把若干元素子节点依次转换并拼接（用于 m:e / m:num 这类“一组内容”的槽位）
function seq(node) {
  return kidsOf(node).map(convert).join("");
}

// 把单个节点转成一段 OMML（mrow 会自动展开成序列）
function slot(node) {
  return node ? convert(node) : "";
}

// 上下标：msup / msub / msubsup
function scripts(kids, type) {
  const base = kids[0];
  // 大型运算符 + 上下标 → n-ary
  if (base && localName(base) === "mo" && NARY.has(base.textContent.trim())) {
    const chr = base.textContent.trim();
    const sub = type === "sub" || type === "subsup" ? slot(kids[1]) : "";
    const sup = type === "sup" ? slot(kids[1]) : type === "subsup" ? slot(kids[2]) : "";
    return naryXml(chr, "subSup", sub, sup);
  }
  const b = slot(base);
  if (type === "sup") return `<m:sSup><m:e>${b}</m:e><m:sup>${slot(kids[1])}</m:sup></m:sSup>`;
  if (type === "sub") return `<m:sSub><m:e>${b}</m:e><m:sub>${slot(kids[1])}</m:sub></m:sSub>`;
  return `<m:sSubSup><m:e>${b}</m:e><m:sub>${slot(kids[1])}</m:sub><m:sup>${slot(kids[2])}</m:sup></m:sSubSup>`;
}

// 上/下方标注：munder / mover / munderover（lim、∑ 的上下限等）
function underover(kids, type) {
  const base = kids[0];
  if (base && localName(base) === "mo" && NARY.has(base.textContent.trim())) {
    const chr = base.textContent.trim();
    const sub = type === "under" || type === "underover" ? slot(kids[1]) : "";
    const sup = type === "over" ? slot(kids[1]) : type === "underover" ? slot(kids[2]) : "";
    return naryXml(chr, "undOvr", sub, sup);
  }
  const b = slot(base);
  if (type === "under") return `<m:limLow><m:e>${b}</m:e><m:lim>${slot(kids[1])}</m:lim></m:limLow>`;
  if (type === "over") return `<m:limUpp><m:e>${b}</m:e><m:lim>${slot(kids[1])}</m:lim></m:limUpp>`;
  // underover：上下都有，嵌套实现
  return `<m:limUpp><m:e><m:limLow><m:e>${b}</m:e><m:lim>${slot(kids[1])}</m:lim></m:limLow></m:e><m:lim>${slot(kids[2])}</m:lim></m:limUpp>`;
}

function naryXml(chr, limLoc, sub, sup) {
  return (
    `<m:nary><m:naryPr><m:chr m:val="${esc(chr)}"/><m:limLoc m:val="${limLoc}"/><m:subHide m:val="${sub ? "off" : "on"}"/><m:supHide m:val="${sup ? "off" : "on"}"/></m:naryPr>` +
    `<m:sub>${sub}</m:sub><m:sup>${sup}</m:sup><m:e></m:e></m:nary>`
  );
}

// 括号 mfenced（已废弃但仍可能出现）
function fenced(node) {
  const open = node.getAttribute("open");
  const close = node.getAttribute("close");
  const begChr = open === null ? "(" : open;
  const endChr = close === null ? ")" : close;
  return `<m:d><m:dPr><m:begChr m:val="${esc(begChr)}"/><m:endChr m:val="${esc(endChr)}"/></m:dPr><m:e>${seq(node)}</m:e></m:d>`;
}

// 矩阵 mtable
function matrix(node) {
  const rows = kidsOf(node).filter((r) => localName(r) === "mtr");
  const body = rows
    .map((r) => {
      const cells = kidsOf(r)
        .filter((c) => localName(c) === "mtd")
        .map((c) => `<m:e>${seq(c)}</m:e>`)
        .join("");
      return `<m:mr>${cells}</m:mr>`;
    })
    .join("");
  return `<m:m>${body}</m:m>`;
}

// 核心递归：把一个节点翻译成 OMML 片段
function convert(node) {
  if (node.nodeType !== 1) return ""; // 文本节点由 token 元素的 textContent 处理
  const name = localName(node);
  const kids = kidsOf(node);

  switch (name) {
    case "math":
    case "mrow":
    case "mstyle":
    case "mpadded":
    case "menclose":
      return seq(node); // 透明容器：直接展开子节点

    case "mi":
    case "mn":
    case "mo":
    case "mtext":
    case "ms":
      return run(node.textContent);

    case "mspace":
    case "none":
      return "";

    case "mfrac":
      return `<m:f><m:num>${slot(kids[0])}</m:num><m:den>${slot(kids[1])}</m:den></m:f>`;

    case "msqrt":
      return `<m:rad><m:radPr><m:degHide m:val="on"/></m:radPr><m:deg/><m:e>${seq(node)}</m:e></m:rad>`;

    case "mroot":
      return `<m:rad><m:deg>${slot(kids[1])}</m:deg><m:e>${slot(kids[0])}</m:e></m:rad>`;

    case "msup":
      return scripts(kids, "sup");
    case "msub":
      return scripts(kids, "sub");
    case "msubsup":
      return scripts(kids, "subsup");

    case "munder":
      return underover(kids, "under");
    case "mover":
      return underover(kids, "over");
    case "munderover":
      return underover(kids, "underover");

    case "mfenced":
      return fenced(node);

    case "mtable":
      return matrix(node);

    default:
      // 未知标签：尽量保留其子内容，避免整段丢失
      return kids.length ? seq(node) : run(node.textContent);
  }
}

/**
 * 把 MathML 字符串转成 OMML 字符串（<m:oMath>…</m:oMath>）。
 * @param {string} mathmlString MathLive 的 getValue('math-ml') 输出
 * @returns {string} OMML
 */
export function mml2omml(mathmlString) {
  const doc = new DOMParser().parseFromString(mathmlString, "application/xml");
  const err = doc.getElementsByTagName("parsererror")[0];
  if (err) throw new Error("MathML 解析失败：" + err.textContent.trim());

  // 取根 <math>（若无则用文档根元素）
  let root = doc.documentElement;
  if (localName(root) !== "math") {
    const found = Array.from(doc.getElementsByTagName("*")).find((n) => localName(n) === "math");
    if (found) root = found;
  }

  return `<m:oMath xmlns:m="${M}">${seq(root)}</m:oMath>`;
}
