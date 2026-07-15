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
let warningSink = null;

function warnLossy(name) {
  if (warningSink) warningSink.add(name);
}

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

// OMML m:acc 的重音字符 → LaTeX 命令。键含组合字符（U+03xx/U+20D7），
// 同时兼容 Word 里可能直接存的间距型重音符（^ ~ ¯ ˙ ¨）。
const ACC_CMD = {
  "̂": "\\hat", "^": "\\hat", "ˆ": "\\hat",
  "̃": "\\tilde", "~": "\\tilde", "˜": "\\tilde",
  "́": "\\acute", "´": "\\acute",
  "̀": "\\grave", "`": "\\grave",
  "̇": "\\dot", "˙": "\\dot",
  "̈": "\\ddot", "¨": "\\ddot",
  "̄": "\\bar", "¯": "\\bar", "ˉ": "\\bar",
  "̌": "\\check", "ˇ": "\\check",
  "⃗": "\\vec", "→": "\\vec",
};

function accLatex(node) {
  const pr = child(node, "accPr");
  const chr = (pr && attrVal(child(pr, "chr"))) || "̂";
  const cmd = ACC_CMD[chr] || "\\hat";
  return `${cmd}${grp(slot(child(node, "e")))}`;
}

function barLatex(node) {
  const pr = child(node, "barPr");
  const pos = (pr && attrVal(child(pr, "pos"))) || "top";
  const cmd = pos === "bot" ? "\\underline" : "\\overline";
  return `${cmd}${grp(slot(child(node, "e")))}`;
}

function groupChrLatex(node) {
  const pr = child(node, "groupChrPr");
  const chr = (pr && attrVal(child(pr, "chr"))) || "";
  const pos = (pr && attrVal(child(pr, "pos"))) || "bot";
  const cmd = chr === "⏟" || (chr === "" && pos === "bot") ? "\\underbrace" : "\\overbrace";
  return `${cmd}${grp(slot(child(node, "e")))}`;
}

// 矩阵的列对齐里是否出现 left/right（→ 对齐公式组 aligned；纯居中则是普通矩阵）
function hasAlignCols(mNode) {
  const mPr = child(mNode, "mPr");
  const mcs = mPr && child(mPr, "mcs");
  if (!mcs) return false;
  return kidsOf(mcs).some((mc) => {
    const mcPr = child(mc, "mcPr");
    const jc = mcPr && attrVal(child(mcPr, "mcJc"));
    return jc === "left" || jc === "right";
  });
}

// 定界符的开闭字符 → 矩阵/分段环境名（与 mathml2omml 的折叠互逆）
function matrixEnv(beg, end) {
  const b = beg ?? "";
  const e = end ?? "";
  if (b === "(" && e === ")") return "pmatrix";
  if (b === "[" && e === "]") return "bmatrix";
  if (b === "{" && e === "}") return "Bmatrix";
  if (b === "|" && e === "|") return "vmatrix";
  if (b === "‖" && e === "‖") return "Vmatrix";
  if (b === "{" && (e === "" || e === ".")) return "cases";
  return null;
}

function delimLatex(node) {
  const pr = child(node, "dPr");
  const beg = (pr && attrVal(child(pr, "begChr"))) ?? "(";
  const end = (pr && attrVal(child(pr, "endChr"))) ?? ")";
  const es = kidsOf(node).filter((n) => localName(n) === "e");
  // 定界符内恰好只包一个矩阵 → 还原成 pmatrix/bmatrix/…/cases
  if (es.length === 1) {
    const inner = kidsOf(es[0]).filter((n) => !isPr(localName(n)));
    if (inner.length === 1 && localName(inner[0]) === "m") {
      const env = matrixEnv(beg, end);
      if (env) return matrixLatex(inner[0], env);
    }
  }
  const items = es.map((e) => seq(e)).join(",");
  const L = beg === "" ? "." : beg;
  const R = end === "" ? "." : end;
  return `\\left${L}${items}\\right${R}`;
}

function matrixLatex(node, env = "matrix") {
  const rows = kidsOf(node).filter((r) => localName(r) === "mr");
  const body = rows
    .map((r) =>
      kidsOf(r)
        .filter((c) => localName(c) === "e")
        .map((c) => seq(c))
        .join(" & ")
    )
    .join(" \\\\ ");
  return `\\begin{${env}}${body}\\end{${env}}`;
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
      // 裸矩阵：带 right/left 列对齐的是对齐公式组（aligned），否则普通 matrix。
      return matrixLatex(node, hasAlignCols(node) ? "aligned" : "matrix");

    case "acc": // 重音：\hat \vec \dot …
      return accLatex(node);
    case "bar": // 上/下划线：\overline \underline
      return barLatex(node);
    case "groupChr": // 上/下花括：\overbrace \underbrace
      return groupChrLatex(node);

    case "limLow": // base 下方加注（如 lim_{x→0}）
      return `\\underset${grp(slot(child(node, "lim")))}${grp(slot(child(node, "e")))}`;
    case "limUpp":
      return `\\overset${grp(slot(child(node, "lim")))}${grp(slot(child(node, "e")))}`;

    default:
      if (isPr(name)) return "";
      warnLossy(name);
      return kidsOf(node).length ? seq(node) : textToLatex(node.textContent);
  }
}

/**
 * OMML 字符串（<m:oMath>… 或 oMathPara）→ LaTeX 字符串。
 * @param {string} ommlString
 * @returns {string}
 */
export function omml2latex(ommlString) {
  return omml2latexDetailed(ommlString).latex;
}

/**
 * 与 omml2latex 相同，但同时报告被降级展开的 OMML 标签。
 * @returns {{latex:string, warnings:string[], lossy:boolean}}
 */
export function omml2latexDetailed(ommlString) {
  const previousSink = warningSink;
  const warnings = new Set();
  warningSink = warnings;
  try {
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
    const latex = convert(root).trim();
    return { latex, warnings: Array.from(warnings).sort(), lossy: warnings.size > 0 };
  } finally {
    warningSink = previousSink;
  }
}

const hasAncestorMath = (node) => {
  let cur = node.parentNode;
  while (cur && cur.nodeType === 1) {
    const name = localName(cur);
    if (name === "oMath" || name === "oMathPara") return true;
    cur = cur.parentNode;
  }
  return false;
};

/**
 * 检查 Word 选区 OOXML 是否只包含一个公式，并提取 GMath 内容控件标签。
 */
export function inspectOmathPackage(ooxmlPackage) {
  const doc = new DOMParser().parseFromString(ooxmlPackage, "application/xml");
  const all = Array.from(doc.getElementsByTagName("*"));
  const paras = all.filter((n) => localName(n) === "oMathPara");
  const standalone = all.filter(
    (n) => localName(n) === "oMath" && !hasAncestorMath(n)
  );
  const regions = [...paras, ...standalone];
  const outsideText = all
    .filter((n) => localName(n) === "t" && !hasAncestorMath(n))
    .map((n) => n.textContent || "")
    .join("")
    .trim();
  const tagNode = all.find((n) => {
    if (localName(n) !== "tag") return false;
    const value = attrVal(n) || n.getAttribute("w:val") || "";
    return value.startsWith("gmath:");
  });
  const gmathTag = tagNode
    ? attrVal(tagNode) || tagNode.getAttribute("w:val") || null
    : null;
  return {
    count: regions.length,
    omml: regions[0] ? new XMLSerializer().serializeToString(regions[0]) : null,
    outsideText,
    gmathTag,
    safe: regions.length === 1 && outsideText === "",
  };
}

/**
 * 从 Word Range.getOoxml() / ContentControl.getOoxml() 返回的 Flat OPC 包里
 * 取出第一个数学区（优先 oMathPara，否则 oMath），序列化成字符串。
 * @param {string} ooxmlPackage
 * @returns {string|null} OMML 字符串；找不到返回 null
 */
export function extractOMath(ooxmlPackage) {
  return inspectOmathPackage(ooxmlPackage).omml;
}
