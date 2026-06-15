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

// 矩阵/分段的定界符。MathLive 会把 \begin{pmatrix}…\end{pmatrix} 之类输出成
// 「开括号 mo + mtable + 闭括号 mo」三个兄弟节点；分段 \begin{cases} 则是
// 「{ mo + mtable」（无右括号）。这里识别这些字符，把矩阵折叠进 m:d 定界符，
// 让 Word 里的括号能随矩阵高度自适应拉伸，而不是降级成普通文字。
const OPEN_FENCE = new Set(["(", "[", "{", "|", "‖", "⟨", "."]);
const CLOSE_FENCE = new Set([")", "]", "}", "|", "‖", "⟩", "."]);

const localName = (n) => n.localName || n.tagName.replace(/^.*:/, "");
// MathLive may emit invisible MathML operators such as U+2062 INVISIBLE TIMES
// for implicit multiplication (for example "4ac"). Word can render those
// controls as tofu boxes, so drop them before producing OMML text runs.
const INVISIBLE_MATH_CHARS = /[\u200b-\u200d\u2061-\u2064\ufeff\ufe00-\ufe0f]/g;

const normalizeTokenText = (s) => String(s).replace(INVISIBLE_MATH_CHARS, "").replace(/\u00a0/g, " ");

const esc = (s) =>
  normalizeTokenText(s).replace(/[&<>"\u007f-\u{10ffff}]/gu, (ch) => {
    switch (ch) {
      case "&":
        return "&amp;";
      case "<":
        return "&lt;";
      case ">":
        return "&gt;";
      case '"':
        return "&quot;";
      default:
        return `&#x${ch.codePointAt(0).toString(16).toUpperCase()};`;
    }
  });

// 仅取元素子节点（忽略空白文本节点）
const kidsOf = (node) => Array.from(node.childNodes).filter((n) => n.nodeType === 1);

// 一个文本 run
const run = (text) => {
  const value = esc(text);
  return value ? `<m:r><m:t xml:space="preserve">${value}</m:t></m:r>` : "";
};

const fenceChar = (node, set) => {
  if (!node || localName(node) !== "mo") return null;
  const t = node.textContent.trim();
  return set.has(t) ? t : null;
};

// 读 mtable 的 columnalign（如 "right left"）→ 每列对齐数组；缺省按列数补 "center"。
function columnAligns(mtableNode) {
  const cols = colCount(mtableNode);
  const raw = (mtableNode.getAttribute("columnalign") || "").trim();
  if (!raw) return Array(cols).fill("center");
  const toks = raw.split(/\s+/).filter(Boolean);
  return Array.from({ length: cols }, (_, i) => toks[Math.min(i, toks.length - 1)] || "center");
}

const colCount = (mtableNode) => {
  const firstRow = kidsOf(mtableNode).find((r) => localName(r) === "mtr");
  return firstRow ? kidsOf(firstRow).filter((c) => localName(c) === "mtd").length : 0;
};

// 对齐公式组（aligned/align）：裸 mtable，列对齐里出现 right/left（而非纯居中矩阵）。
const isAlignArray = (mtableNode) =>
  columnAligns(mtableNode).some((a) => a === "left" || a === "right");

// 把若干元素子节点依次转换并拼接（用于 m:e / m:num 这类“一组内容”的槽位）。
// 同时识别：①「开括号 + 矩阵 + 闭括号?」→ m:d 定界符包裹的矩阵；
//           ②裸的对齐公式组 → 带交替列对齐的 m:m。
function seq(node) {
  const kids = kidsOf(node);
  const out = [];
  for (let i = 0; i < kids.length; i++) {
    const open = fenceChar(kids[i], OPEN_FENCE);
    const next = kids[i + 1];
    if (open !== null && next && localName(next) === "mtable") {
      const close = fenceChar(kids[i + 2], CLOSE_FENCE);
      out.push(delimitedMatrix(next, open, close ?? ""));
      i += close !== null ? 2 : 1;
      continue;
    }
    const k = kids[i];
    if (localName(k) === "mtable" && isAlignArray(k)) {
      out.push(matrix(k, columnAligns(k)));
      continue;
    }
    out.push(convert(k));
  }
  return out.join("");
}

// 矩阵外加定界符。分段（左花括号、无右括号）用全列左对齐。
function delimitedMatrix(mtableNode, open, close) {
  const isCases = open === "{" && (close === "" || close === ".");
  const aligns = isCases ? Array(colCount(mtableNode)).fill("left") : null;
  const inner = matrix(mtableNode, aligns);
  return (
    `<m:d><m:dPr><m:begChr m:val="${esc(open)}"/><m:endChr m:val="${esc(close)}"/><m:grow/></m:dPr>` +
    `<m:e>${inner}</m:e></m:d>`
  );
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
  if (type === "sup") return `<m:sSup><m:sSupPr/><m:e>${b}</m:e><m:sup>${slot(kids[1])}</m:sup></m:sSup>`;
  if (type === "sub") return `<m:sSub><m:sSubPr/><m:e>${b}</m:e><m:sub>${slot(kids[1])}</m:sub></m:sSub>`;
  return `<m:sSubSup><m:sSubSupPr/><m:e>${b}</m:e><m:sub>${slot(kids[1])}</m:sub><m:sup>${slot(kids[2])}</m:sup></m:sSubSup>`;
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

// 把每列对齐数组（如 ["right","left"]）压成 OMML 的 m:mcs：相邻同对齐的列合并成
// 一个 m:mc，用 m:count 记列数。aligns 为 null 时返回空（默认居中，无 m:mPr）。
function matrixColsXml(aligns) {
  if (!aligns || !aligns.length) return "";
  const groups = [];
  for (const a of aligns) {
    const last = groups[groups.length - 1];
    if (last && last.jc === a) last.count++;
    else groups.push({ jc: a, count: 1 });
  }
  const mcs = groups
    .map((g) => `<m:mc><m:mcPr><m:count m:val="${g.count}"/><m:mcJc m:val="${g.jc}"/></m:mcPr></m:mc>`)
    .join("");
  return `<m:mPr><m:mcs>${mcs}</m:mcs></m:mPr>`;
}

// 矩阵 mtable。aligns 为每列对齐（"left"/"right"/"center"）数组，用于分段函数
// （全左对齐）与对齐公式组（交替右/左）；普通矩阵传 null（居中）。
function matrix(node, aligns = null) {
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
  return `<m:m>${matrixColsXml(aligns)}${body}</m:m>`;
}

// Content MathML: <apply><power/><ci>x</ci><cn>5</cn></apply>
// MathLive normally emits Presentation MathML for "math-ml", but supporting
// this shape prevents scripts from degrading to plain adjacent runs ("x5").
function applyContent(kids) {
  const op = kids[0] ? localName(kids[0]) : "";
  const args = kids.slice(1);

  switch (op) {
    case "power":
      return scripts(args, "sup");

    case "root": {
      const degree = args.find((arg) => localName(arg) === "degree");
      const radicand = args.find((arg) => localName(arg) !== "degree");
      if (degree) return `<m:rad><m:deg>${seq(degree)}</m:deg><m:e>${slot(radicand)}</m:e></m:rad>`;
      return `<m:rad><m:radPr><m:degHide m:val="on"/></m:radPr><m:deg/><m:e>${slot(radicand)}</m:e></m:rad>`;
    }

    case "divide":
      return `<m:f><m:num>${slot(args[0])}</m:num><m:den>${slot(args[1])}</m:den></m:f>`;

    case "plus":
      return args.map(slot).join(run("+"));

    case "times":
      return args.map(slot).join("");

    case "minus":
      return args.length === 1 ? run("-") + slot(args[0]) : args.map(slot).join(run("-"));

    case "eq":
      return args.map(slot).join(run("="));

    default:
      return args.map(slot).join("");
  }
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
    case "ci":
    case "cn":
      return run(node.textContent);

    case "mspace":
    case "none":
      return "";

    case "apply":
      return applyContent(kids);

    case "degree":
      return seq(node);

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
      // 裸 mtable 作为根直达此处（未经 seq）：对齐公式组保留列对齐，普通矩阵居中。
      return matrix(node, isAlignArray(node) ? columnAligns(node) : null);

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

  // 关键：必须 convert(root) 而不是 seq(root)。
  // MathLive 的 getValue('math-ml') 经常输出“裸”的单个元素做根（如 x^1 → <msup>…</msup>，
  // 不带 <math> 外壳）。seq(root) 会把 root 自身的标签丢掉、只转换它的子节点，
  // 导致 <msup>/<mfrac> 等结构降级成相邻的普通文字 run（x¹ 变 "x1"，a/b 变 "ab"）。
  // convert(root) 会先翻译 root 这一层；当 root 恰好是 <math>/<mrow> 这类透明容器时，
  // convert 内部本就会退化成 seq，行为与原来一致。
  return `<m:oMath xmlns:m="${M}">${convert(root)}</m:oMath>`;
}
