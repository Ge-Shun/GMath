// 正向补丁层：处理 MathLive 自身无法序列化的「装饰命令」。
//
// 背景：MathLive 0.110 的 MathML 序列化对这几个命令是坏的——
//   \overline{x}/\underline{x}  → 输出空字符串（内容全丢）
//   \overbrace{x+y}             → <mover>⏞</mover>（base x+y 被丢掉）
//   \underbrace{x+y}            → <munder>⏟</munder>（同上）
//   \overrightarrow{AB}         → <mover>→</mover>（base AB 被丢掉）
// 内容在送进 mathml2omml 之前就没了，单靠 MathML→OMML 这步无法补救。
//
// 思路（占位符回填）：把每个装饰命令替换成一个私有区(PUA)占位字符，整体交给
// MathLive 建结构——此时装饰只是个叶子占位，frac/上下标/根号/矩阵等所有结构与
// 嵌套都由 MathLive 正确处理；得到 OMML 后，再把占位 run 换成真正的装饰 OMML，
// 其参数通过对本函数的递归转换得到，于是任意嵌套都成立。

// 装饰命令 → 生成对应 OMML 的包裹函数（e 为已转好的参数 OMML 片段）
const DECORATIONS = {
  "\\overline": (e) => `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>${e}</m:e></m:bar>`,
  "\\underline": (e) => `<m:bar><m:barPr><m:pos m:val="bot"/></m:barPr><m:e>${e}</m:e></m:bar>`,
  "\\overbrace": (e) =>
    `<m:groupChr><m:groupChrPr><m:chr m:val="⏞"/><m:pos m:val="top"/><m:vertJc m:val="bot"/></m:groupChrPr><m:e>${e}</m:e></m:groupChr>`,
  "\\underbrace": (e) =>
    `<m:groupChr><m:groupChrPr><m:chr m:val="⏟"/><m:pos m:val="bot"/><m:vertJc m:val="top"/></m:groupChrPr><m:e>${e}</m:e></m:groupChr>`,
  "\\overrightarrow": (e) => `<m:acc><m:accPr><m:chr m:val="⃗"/></m:accPr><m:e>${e}</m:e></m:acc>`,
  "\\overleftarrow": (e) => `<m:acc><m:accPr><m:chr m:val="⃖"/></m:accPr><m:e>${e}</m:e></m:acc>`,
};

const DECO_RE = /\\(overline|underline|overbrace|underbrace|overrightarrow|overleftarrow)\b/;
const PUA_BASE = 0xe000;

// 是否含需要特殊处理的装饰命令（含则走补丁层，否则走原 MathML 快路径）。
export function hasDecoration(latex) {
  return DECO_RE.test(latex || "");
}

// 从 latex 的 startBrace（必须是 "{"）开始，返回配平的 "}" 下标。找不到返回 -1。
function matchBrace(latex, openIdx) {
  let depth = 0;
  for (let i = openIdx; i < latex.length; i++) {
    const c = latex[i];
    if (c === "\\") {
      i++; // 跳过被转义的下一个字符（如 \{ \} ）
      continue;
    }
    if (c === "{") depth++;
    else if (c === "}") {
      depth--;
      if (depth === 0) return i;
    }
  }
  return -1;
}

// 把 latex 里「顶层遇到的」装饰命令替换成 PUA 占位字符；不深入其参数（参数留待
// 递归处理）。返回 { latex: 占位后的 latex, slots: [{ pua, wrap, arg }] }。
// 注意：扫描会穿过 \frac{…} 等其它命令的花括号，所以嵌在别的结构里的装饰也能被命中。
function replaceDecorations(latex) {
  const slots = [];
  let out = "";
  let i = 0;
  while (i < latex.length) {
    const c = latex[i];
    if (c !== "\\") {
      out += c;
      i++;
      continue;
    }
    // 读命令名：\ 后面连续的字母
    let j = i + 1;
    while (j < latex.length && /[a-zA-Z]/.test(latex[j])) j++;
    const cmd = latex.slice(i, j);
    const wrap = DECORATIONS[cmd];
    if (!wrap) {
      out += cmd; // 非装饰命令：原样保留，继续
      i = j;
      continue;
    }
    // 跳过命令名后的空白，期待一个 "{arg}"
    let k = j;
    while (k < latex.length && /\s/.test(latex[k])) k++;
    if (latex[k] !== "{") {
      out += cmd; // 没有花括号参数：原样保留交给 MathLive
      i = j;
      continue;
    }
    const close = matchBrace(latex, k);
    if (close === -1) {
      out += cmd; // 花括号不配平：放弃拦截
      i = j;
      continue;
    }
    const arg = latex.slice(k + 1, close);
    const pua = String.fromCodePoint(PUA_BASE + slots.length);
    slots.push({ pua, wrap, arg });
    out += pua;
    i = close + 1;
  }
  return { latex: out, slots };
}

// 取 mml2omml 输出（<m:oMath …>…</m:oMath>）里的内层内容，用于嵌入别的槽位。
function innerOMath(omml) {
  const m = omml.match(/<m:oMath\b[^>]*>([\s\S]*)<\/m:oMath>/);
  return m ? m[1] : omml;
}

// 把包住某个 PUA 占位字符的那个文本 run 整体替换成 replacement。
// mml2omml 的 run 形如 <m:r><m:t xml:space="preserve">&#xE000;</m:t></m:r>；
// 占位字符（>U+007F）会被 esc 成数值引用 &#xE000;。
function replacePlaceholderRun(omml, pua, replacement) {
  const hex = pua.codePointAt(0).toString(16).toUpperCase();
  const ref = `&#x${hex};`;
  // 匹配包含该字符引用、且 run 内不跨越其它 run 的最小文本 run
  const re = new RegExp(`<m:r>(?:(?!</m:r>)[\\s\\S])*?${ref}(?:(?!</m:r>)[\\s\\S])*?</m:r>`);
  return omml.replace(re, () => replacement);
}

/**
 * LaTeX → OMML，正确处理 \overline/\underline/\overbrace/\underbrace/
 * \overrightarrow/\overleftarrow（含任意嵌套），其余结构全部委托给 MathLive。
 * @param {string} latex
 * @param {{ convertLatexToMathMl:(s:string)=>string, mml2omml:(s:string)=>string }} deps
 * @returns {string} <m:oMath …>…</m:oMath>
 */
export function latexToOmml(latex, deps) {
  const { convertLatexToMathMl, mml2omml } = deps;
  const { latex: replaced, slots } = replaceDecorations(latex);

  // 无装饰：直接走原链路（与旧行为一致）
  if (!slots.length) return mml2omml(convertLatexToMathMl(latex));

  let omml = mml2omml(convertLatexToMathMl(replaced));
  for (const { pua, wrap, arg } of slots) {
    const argOmml = innerOMath(latexToOmml(arg, deps)); // 递归 → 支持嵌套
    omml = replacePlaceholderRun(omml, pua, wrap(argOmml));
  }
  return omml;
}
