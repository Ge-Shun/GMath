// 反向转换器：OMML（Word 数学标记）→ LaTeX
// 用途：把文档里已有的公式（GMath 插入的或 Word 自带编辑器做的）读回 LaTeX，
//       直接喂给 MathLive 的 mathfield.setValue(latex) 载入面板继续编辑。
//
// 为什么直接产 LaTeX 而不是 MathML：
//   当前 MathLive 只提供 LaTeX→MathML（convertLatexToMathMl），没有 MathML→LaTeX，
//   setValue 也不支持 {format:"math-ml"}。所以反向必须自己产出 LaTeX。
//
// 设计：用 DOMParser 解析，localName 取标签名（忽略 m: 前缀），递归翻译，
//       与 src/mathml2omml.js 互为逆操作，覆盖同一子集。

const localName = (n) => n.localName || n.tagName.replace(/^.*:/, "");
const kidsOf = (node) => Array.from(node.childNodes).filter((n) => n.nodeType === 1);
const child = (node, name) => kidsOf(node).find((n) => localName(n) === name) || null;
const isPr = (name) => /Pr$/.test(name); // naryPr / radPr / sSupPr / dPr / mPr / rPr …
const attrVal = (el) => (el ? el.getAttribute("m:val") ?? el.getAttribute("val") : null);

// n-ary 运算符 → LaTeX 命令
const NARY_CMD = {
  "∑": "\\sum", "∏": "\\prod", "∐": "\\coprod",
  "∫": "\\int", "∬": "\\iint", "∭": "\\iiint",
  "∮": "\\oint", "∯": "\\oiint", "∰": "\\oiiint",
  "⋃": "\\bigcup", "⋂": "\\bigcap", "⋁": "\\bigvee", "⋀": "\\bigwedge",
  "⨁": "\\bigoplus", "⨂": "\\bigotimes", "⨀": "\\bigodot",
};

// 文本里若干 Unicode 数学符号 → 更稳妥的 LaTeX 写法（MathLive 对部分 Unicode 容忍度不一）
const SYMBOL = {
  "−": "-", "×": "\\times ", "÷": "\\div ", "±": "\\pm ", "∓": "\\mp ",
  "⋅": "\\cdot ", "∙": "\\cdot ", "→": "\\to ", "←": "\\leftarrow ",
  "≤": "\\le ", "≥": "\\ge ", "≠": "\\ne ", "≈": "\\approx ", "≡": "\\equiv ",
  "∞": "\\infty ", "∂": "\\partial ", "∇": "\\nabla ", "…": "\\ldots ", "⋯": "\\cdots ",
};

// LaTeX 里需要转义的特殊字符（出现在普通文本 token 时）
const LATEX_SPECIAL = /[\\#$%&_^{}~]/g;
const LATEX_ESCAPE = {
  "\\": "\\backslash ", "#": "\\#", $: "\\$", "%": "\\%",
  "&": "\\&", _: "\\_", "^": "\\^{}", "{": "\\{", "}": "\\}", "~": "\\~{}",
};

function textToLatex(s) {
  let out = "";
  for (const ch of String(s)) {
    if (ch === " ") { out += " "; continue; }
    if (SYMBOL[ch]) out += SYMBOL[ch];
    else if (LATEX_SPECIAL.test(ch)) { LATEX_SPECIAL.lastIndex = 0; out += LATEX_ESCAPE[ch]; }
    else out += ch;
  }
  return out;
}

// 文本 run：取其下所有 m:t 文本
function runText(node) {
  return kidsOf(node)
    .filter((n) => localName(n) === "t")
    .map((t) => textToLatex(t.textContent))
    .join("");
}

// 用 {} 包裹一个槽位（保证上下标/分子分母作为整体）
const grp = (s) => `{${s}}`;
const slot = (node) => (node ? seq(node) : "");

function seq(node) {
  return kidsOf(node)
    .filter((n) => !isPr(localName(n)))
    .map(convert)
    .join("");
}

function naryLatex(node) {
  const pr = child(node, "naryPr");
  const chr = (pr && attrVal(child(pr, "chr"))) || "∫";
  const cmd = NARY_CMD[chr] || chr;
  const sub = slot(child(node, "sub"));
  const sup = slot(child(node, "sup"));
  const e = slot(child(node, "e"));
  let out = cmd;
  if (sub) out += "_" + grp(sub);
  if (sup) out += "^" + grp(sup);
  return out + (e ? " " + e : "");
}

function delimLatex(node) {
  const pr = child(node, "dPr");
  const beg = (pr && attrVal(child(pr, "begChr"))) ?? "(";
  const end = (pr && attrVal(child(pr, "endChr"))) ?? ")";
  const items = kidsOf(node)
    .filter((n) => localName(n) === "e")
    .map((e) => seq(e))
    .join(",");
  const L = beg === "" ? "." : beg;
  const R = end === "" ? "." : end;
  return `\\left${L}${items}\\right${R}`;
}

function matrixLatex(node) {
  const rows = kidsOf(node).filter((r) => localName(r) === "mr");
  const body = rows
    .map((r) =>
      kidsOf(r)
        .filter((c) => localName(c) === "e")
        .map((c) => seq(c))
        .join(" & ")
    )
    .join(" \\\\ ");
  return `\\begin{matrix}${body}\\end{matrix}`;
}

function radLatex(node) {
  const pr = child(node, "radPr");
  const degHide = attrVal(child(pr || node, "degHide")) !== "off" && !!(pr && child(pr, "degHide"));
  const deg = child(node, "deg");
  const e = slot(child(node, "e"));
  const degText = deg ? seq(deg) : "";
  if (!degText) return `\\sqrt${grp(e)}`;
  return `\\sqrt[${degText}]${grp(e)}`;
}

function convert(node) {
  if (node.nodeType !== 1) return "";
  const name = localName(node);

  switch (name) {
    case "oMath":
    case "oMathPara":
    case "e":
    case "num":
    case "den":
    case "deg":
    case "sub":
    case "sup":
    case "lim":
      return seq(node);

    case "r":
      return runText(node);
    case "t":
      return textToLatex(node.textContent);

    case "f":
      return `\\frac${grp(slot(child(node, "num")))}${grp(slot(child(node, "den")))}`;

    case "rad":
      return radLatex(node);

    case "sSup":
      return `${grp(slot(child(node, "e")))}^${grp(slot(child(node, "sup")))}`;
    case "sSub":
      return `${grp(slot(child(node, "e")))}_${grp(slot(child(node, "sub")))}`;
    case "sSubSup":
      return `${grp(slot(child(node, "e")))}_${grp(slot(child(node, "sub")))}^${grp(slot(child(node, "sup")))}`;

    case "nary":
      return naryLatex(node);

    case "d":
      return delimLatex(node);

    case "m":
      return matrixLatex(node);

    case "limLow": // base 下方加注（如 lim_{x→0}）
      return `\\underset${grp(slot(child(node, "lim")))}${grp(slot(child(node, "e")))}`;
    case "limUpp":
      return `\\overset${grp(slot(child(node, "lim")))}${grp(slot(child(node, "e")))}`;

    default:
      if (isPr(name)) return "";
      return kidsOf(node).length ? seq(node) : textToLatex(node.textContent);
  }
}

/**
 * OMML 字符串（<m:oMath>… 或 oMathPara）→ LaTeX 字符串。
 * @param {string} ommlString
 * @returns {string}
 */
export function omml2latex(ommlString) {
  const doc = new DOMParser().parseFromString(ommlString, "application/xml");
  const err = doc.getElementsByTagName("parsererror")[0];
  if (err) throw new Error("OMML 解析失败：" + err.textContent.trim());

  let root = doc.documentElement;
  if (localName(root) !== "oMath" && localName(root) !== "oMathPara") {
    const found = Array.from(doc.getElementsByTagName("*")).find(
      (n) => localName(n) === "oMath" || localName(n) === "oMathPara"
    );
    if (found) root = found;
  }
  return convert(root).trim();
}

/**
 * 从 Word Range.getOoxml() / ContentControl.getOoxml() 返回的 Flat OPC 包里
 * 取出第一个数学区（优先 oMathPara，否则 oMath），序列化成字符串。
 * @param {string} ooxmlPackage
 * @returns {string|null} OMML 字符串；找不到返回 null
 */
export function extractOMath(ooxmlPackage) {
  const doc = new DOMParser().parseFromString(ooxmlPackage, "application/xml");
  const all = Array.from(doc.getElementsByTagName("*"));
  const para = all.find((n) => localName(n) === "oMathPara");
  const target = para || all.find((n) => localName(n) === "oMath");
  if (!target) return null;
  return new XMLSerializer().serializeToString(target);
}
