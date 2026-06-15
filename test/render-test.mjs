// 全量公式渲染测试
// 思路：mml2omml.js 是整条链路里唯一可在 Node 中独立验证的环节
//   （MathLive 在浏览器里负责 LaTeX→MathML，Word 负责 OMML→可视渲染）。
// 这里用 @xmldom/xmldom 提供 DOMParser，喂入各类公式对应的、与 MathLive
// 实际输出同构的 Presentation MathML，校验：
//   1) 转换不抛异常；
//   2) 输出的 OMML 是合法（可被重新解析）的 XML；
//   3) 输出包含该类公式应有的关键 OMML 结构。

import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
globalThis.DOMParser = DOMParser; // mml2omml.js / omml2mml.js 依赖全局 DOMParser
globalThis.XMLSerializer = XMLSerializer; // omml2mml.js 的 extractOMath 依赖

const { mml2omml } = await import("../src/mathml2omml.js");
const { omml2latex } = await import("../src/omml2latex.js");

// MathLive 隐式乘法用的不可见字符
const IT = "⁢"; // INVISIBLE TIMES

// 每个用例：name=描述，latex=对应的 LaTeX，mathml=MathLive 同构输出，
// expect=输出 OMML 必须包含的关键标签/字符（数组，全部命中才算通过）
const cases = [
  {
    name: "单变量",
    latex: "x",
    mathml: `<mi>x</mi>`,
    expect: ["<m:r>", "x"],
  },
  {
    name: "分数 a/b",
    latex: "\\frac{a}{b}",
    mathml: `<mfrac><mi>a</mi><mi>b</mi></mfrac>`,
    expect: ["<m:f>", "<m:num>", "<m:den>"],
  },
  {
    name: "嵌套分数",
    latex: "\\frac{a}{\\frac{b}{c}}",
    mathml: `<mfrac><mi>a</mi><mfrac><mi>b</mi><mi>c</mi></mfrac></mfrac>`,
    expect: ["<m:f><m:num>", "<m:den><m:f>"],
  },
  {
    name: "平方根",
    latex: "\\sqrt{x}",
    mathml: `<msqrt><mi>x</mi></msqrt>`,
    expect: ["<m:rad>", "<m:degHide", "<m:e>"],
  },
  {
    name: "n 次根",
    latex: "\\sqrt[3]{x}",
    mathml: `<mroot><mi>x</mi><mn>3</mn></mroot>`,
    expect: ["<m:rad>", "<m:deg>", "3"],
  },
  {
    name: "上标 x^2",
    latex: "x^2",
    mathml: `<msup><mi>x</mi><mn>2</mn></msup>`,
    expect: ["<m:sSup>", "<m:sup>", "2"],
  },
  {
    name: "下标 x_n",
    latex: "x_n",
    mathml: `<msub><mi>x</mi><mi>n</mi></msub>`,
    expect: ["<m:sSub>", "<m:sub>"],
  },
  {
    name: "上下标 x_n^2",
    latex: "x_n^2",
    mathml: `<msubsup><mi>x</mi><mi>n</mi><mn>2</mn></msubsup>`,
    expect: ["<m:sSubSup>", "<m:sub>", "<m:sup>"],
  },
  {
    name: "求和（上下限，n-ary）",
    latex: "\\sum_{i=1}^{n} i",
    mathml: `<mrow><munderover><mo>∑</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></munderover><mi>i</mi></mrow>`,
    expect: ["<m:nary>", `m:chr m:val="∑"`, "<m:sub>", "<m:sup>"],
  },
  {
    name: "积分（上下限，n-ary）",
    latex: "\\int_0^1 x\\,dx",
    mathml: `<mrow><msubsup><mo>∫</mo><mn>0</mn><mn>1</mn></msubsup><mi>x</mi><mi>d</mi><mi>x</mi></mrow>`,
    expect: ["<m:nary>", `m:chr m:val="∫"`],
  },
  {
    name: "极限 lim",
    latex: "\\lim_{x\\to 0} f(x)",
    mathml: `<mrow><munder><mrow><mi>lim</mi></mrow><mrow><mi>x</mi><mo>→</mo><mn>0</mn></mrow></munder><mi>f</mi></mrow>`,
    expect: ["<m:limLow>", "<m:lim>", "lim"],
  },
  {
    name: "矩阵 pmatrix",
    latex: "\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}",
    mathml: `<mrow><mrow><mo>(</mo><mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable><mo>)</mo></mrow></mrow>`,
    expect: ["<m:d>", `m:begChr m:val="("`, `m:endChr m:val=")"`, "<m:m>", "<m:mr>", "<m:e>"],
  },
  {
    name: "矩阵 bmatrix（方括号）",
    latex: "\\begin{bmatrix}a&b\\\\c&d\\end{bmatrix}",
    mathml: `<mrow><mo>[</mo><mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable><mo>]</mo></mrow>`,
    expect: ["<m:d>", `m:begChr m:val="["`, `m:endChr m:val="]"`, "<m:m>"],
  },
  {
    name: "分段函数 cases",
    latex: "f(x)=\\begin{cases}1&x>0\\\\0&x\\le 0\\end{cases}",
    mathml: `<mrow><mo>{</mo><mtable columnalign="left left"><mtr><mtd><mn>1</mn></mtd><mtd><mi>x</mi><mo>&gt;</mo><mn>0</mn></mtd></mtr><mtr><mtd><mn>0</mn></mtd><mtd><mi>x</mi><mo>≤</mo><mn>0</mn></mtd></mtr></mtable></mrow>`,
    expect: ["<m:d>", `m:begChr m:val="{"`, `m:endChr m:val=""`, "<m:m>", "<m:mcJc m:val=\"left\""],
  },
  {
    name: "裸矩阵（无定界符）",
    latex: "\\begin{matrix}a&b\\\\c&d\\end{matrix}",
    mathml: `<mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable>`,
    expect: ["<m:m>", "<m:mr>"],
    forbid: ["<m:d>"],
  },
  {
    name: "对齐公式组 aligned",
    latex: "\\begin{aligned}x&=a+b\\\\y&=c\\end{aligned}",
    mathml: `<mtable columnalign="right left"><mtr><mtd><mi>x</mi></mtd><mtd><mo>=</mo><mi>a</mi><mo>+</mo><mi>b</mi></mtd></mtr><mtr><mtd><mi>y</mi></mtd><mtd><mo>=</mo><mi>c</mi></mtd></mtr></mtable>`,
    expect: ["<m:m>", `m:mcJc m:val="right"`, `m:mcJc m:val="left"`],
    forbid: ["<m:d>"],
  },
  {
    name: "希腊字母 α+β",
    latex: "\\alpha+\\beta",
    mathml: `<mrow><mi>α</mi><mo>+</mo><mi>β</mi></mrow>`,
    expect: ["α", "β", "+"],
  },
  {
    name: "质能方程 E=mc^2",
    latex: "E=mc^2",
    mathml: `<mrow><mi>E</mi><mo>=</mo><mi>m</mi><msup><mi>c</mi><mn>2</mn></msup></mrow>`,
    expect: ["E", "=", "<m:sSup>"],
  },
  {
    name: "一元二次求根公式",
    latex: "x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}",
    mathml: `<mrow><mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mo>${IT}</mo><mi>c</mi></mrow></msqrt></mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></mrow>`,
    expect: ["<m:f>", "<m:rad>", "<m:sSup>", "±"],
    forbid: [IT], // 不可见乘号必须被剔除
  },
  {
    name: "三角函数 sin x",
    latex: "\\sin x",
    mathml: `<mrow><mi>sin</mi><mo>${IT}</mo><mi>x</mi></mrow>`,
    expect: ["sin", "x"],
    forbid: [IT],
  },
  {
    name: "向量 \\vec{v}（重音/上方箭头）",
    latex: "\\vec{v}",
    mathml: `<mover><mi>v</mi><mo stretchy="true">→</mo></mover>`,
    expect: ["<m:limUpp>", "v"],
  },
  {
    name: "括号 mfenced",
    latex: "(a+b)",
    mathml: `<mfenced open="(" close=")"><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow></mfenced>`,
    expect: ["<m:d>", `m:begChr m:val="("`, `m:endChr m:val=")"`],
  },
  {
    name: "特殊字符转义 a<b & c",
    latex: "a<b",
    mathml: `<mrow><mi>a</mi><mo>&lt;</mo><mi>b</mi><mo>&amp;</mo><mi>c</mi></mrow>`,
    expect: ["&lt;", "&amp;"],
  },
];

// 把 OMML 包成 mml2omml 同样的根，验证可被重新解析（well-formed）。
// 只把致命错误（fatalError）当成不合法；warning/error 不影响 OMML 本身的良构性。
function isWellFormed(xml) {
  const wrapped = `<root xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${xml}</root>`;
  let fatal = false;
  const parser = new DOMParser({ onError: (level) => { if (level === "fatalError") fatal = true; } });
  try {
    const doc = parser.parseFromString(wrapped, "application/xml");
    if (fatal) return false;
    if (doc.getElementsByTagName("parsererror").length) return false;
    return true;
  } catch {
    return false;
  }
}

// esc() 会把非 ASCII 字符转成数值字符引用（如 ∑ → &#x2211;），这在 XML 中等价且合法。
// 断言关键字符时先解码回字面量，避免误判。
function decodeRefs(s) {
  return s
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)));
}

// 反向（OMML→LaTeX）往返时，每类公式回来的 LaTeX 应包含的关键命令/符号。
// 只为结构型用例设断言；纯文本类（希腊字母/特殊字符等）只验“能往返、非空”。
const latexExpect = {
  "分数 a/b": ["\\frac"],
  嵌套分数: ["\\frac{a}", "\\frac{b}"],
  平方根: ["\\sqrt"],
  "n 次根": ["\\sqrt[3]"],
  "上标 x^2": ["^"],
  "下标 x_n": ["_"],
  "上下标 x_n^2": ["_", "^"],
  "求和（上下限，n-ary）": ["\\sum", "_", "^"],
  "积分（上下限，n-ary）": ["\\int"],
  "极限 lim": ["\\underset"],
  "矩阵 pmatrix": ["\\begin{pmatrix}", "&", "\\\\", "\\end{pmatrix}"],
  "矩阵 bmatrix（方括号）": ["\\begin{bmatrix}", "\\end{bmatrix}"],
  "分段函数 cases": ["\\begin{cases}", "&", "\\\\", "\\end{cases}"],
  "裸矩阵（无定界符）": ["\\begin{matrix}", "\\end{matrix}"],
  "对齐公式组 aligned": ["\\begin{aligned}", "&", "\\\\", "\\end{aligned}"],
  "质能方程 E=mc^2": ["^"],
  一元二次求根公式: ["\\frac", "\\sqrt", "^", "\\pm"],
  "向量 \\vec{v}（重音/上方箭头）": ["\\overset"],
};

let pass = 0;
const fails = [];
const pad = (s, n) => (s + " ".repeat(n)).slice(0, n);

console.log("=".repeat(72));
console.log(pad("公式类型", 28), pad("LaTeX", 26), "结果");
console.log("-".repeat(72));

for (const c of cases) {
  let omml = "";
  const problems = [];
  try {
    omml = mml2omml(c.mathml);
  } catch (e) {
    problems.push("转换抛异常: " + e.message);
  }
  if (omml) {
    if (!isWellFormed(omml)) problems.push("OMML 非合法 XML");
    const decoded = decodeRefs(omml);
    for (const token of c.expect || []) {
      if (!decoded.includes(token)) problems.push("缺少结构: " + JSON.stringify(token));
    }
    for (const token of c.forbid || []) {
      if (decoded.includes(token)) problems.push("残留禁止字符: " + JSON.stringify(token));
    }

    // 反向往返：OMML → omml2latex → LaTeX（即「从文档读回公式」的链路）
    let backLatex = null;
    try {
      backLatex = omml2latex(decoded); // 用解码后的 OMML，符号为字面量
    } catch (e) {
      problems.push("反向转换抛异常: " + e.message);
    }
    if (backLatex != null) {
      if (!backLatex.trim()) problems.push("反向 LaTeX 为空");
      for (const token of latexExpect[c.name] || []) {
        if (!backLatex.includes(token)) problems.push("往返丢失: " + JSON.stringify(token));
      }
    }
  }
  const ok = problems.length === 0;
  if (ok) pass++;
  else fails.push({ name: c.name, problems, omml });
  console.log(pad(c.name, 28), pad(c.latex, 26), ok ? "✅ 通过" : "❌ 失败");
}

console.log("-".repeat(72));
console.log(`合计 ${cases.length} 项，通过 ${pass} 项，失败 ${cases.length - pass} 项。`);

if (fails.length) {
  console.log("\n失败详情：");
  for (const f of fails) {
    console.log(`\n● ${f.name}`);
    f.problems.forEach((p) => console.log("   - " + p));
    if (f.omml) console.log("   OMML: " + f.omml.slice(0, 300));
  }
  process.exit(1);
}
