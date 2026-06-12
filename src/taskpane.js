// GMath Word 加载项 —— 核心逻辑（纯原生公式 + 双向实时同步，尽力而为）
// 正向：MathLive 编辑 → MathML → mml2omml 转 OMML → Flat OPC → insertOoxml
// 反向：点选文档里的公式 → getOoxml → extractOMath → omml2latex → 载入面板
// 同步：面板里改动 → 防抖 ~0.5s → 替换“当前选中的那个公式”，并重新选中以便继续同步
//
// 注意（纯原生、无锚点的固有限制）：回写靠“替换当前选区”。
//   想可靠同步，请把整个公式选中（拖选或在公式上点一下再用方向键/三击选中），
//   光标只是停在公式里时，替换可能变成插入。详见 README。

import { mml2omml } from "./mathml2omml.js";
import { omml2latex, extractOMath } from "./omml2latex.js";

const BUILD = "2026-06-12-common-desc-cache-bust";

// XML 文本转义（用于把用户输入的编号安全嵌入 OOXML）
const escXml = (s) =>
  String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

// ===== 符号速选数据（分类，参考 MathType / AxMath） =====
// 每项 { l: 按钮显示, i: 插入的 LaTeX, t?: 悬停说明 }；#0 表示插入后的光标占位
const SYMBOL_CATEGORIES = [
  {
    name: "常用",
    items: [
      { l: "□/□", m: "<span class=\"sym-frac\"><span>□</span><span>□</span></span>", d: "分数", i: "\\frac{#0}{#0}", t: "分数" },
      { l: "√□", d: "平方根", i: "\\sqrt{#0}", t: "平方根" },
      { l: "ⁿ√□", m: "<sup>n</sup>√□", d: "n 次根", i: "\\sqrt[#0]{#0}", t: "n 次根" },
      { l: "x^□", m: "x<sup>□</sup>", d: "上标", i: "x^{#0}", t: "上标" },
      { l: "x_□", m: "<span class=\"sym-subonly\"><span>x</span><sub>□</sub></span>", d: "下标", i: "x_{#0}", t: "下标" },
      { l: "x_□^□", m: "<span class=\"sym-subsup\"><span>x</span><span><sup>□</sup><sub>□</sub></span></span>", d: "上下标", i: "x_{#0}^{#0}", t: "上下标" },
      { l: "(□)", d: "圆括号", i: "\\left(#0\\right)", t: "圆括号" },
      { l: "[□]", d: "方括号", i: "\\left[#0\\right]", t: "方括号" },
      { l: "{□}", d: "花括号", i: "\\left\\{#0\\right\\}", t: "花括号" },
      { l: "|□|", d: "绝对值", i: "\\left|#0\\right|", t: "绝对值" },
      { l: "‖□‖", d: "范数", i: "\\left\\|#0\\right\\|", t: "范数" },
      { l: "⌊□⌋", d: "向下取整", i: "\\left\\lfloor #0\\right\\rfloor", t: "向下取整" },
      { l: "⌈□⌉", d: "向上取整", i: "\\left\\lceil #0\\right\\rceil", t: "向上取整" },
      { l: "(□□)", m: "(<span class=\"sym-stack\"><span>□</span><span>□</span></span>)", d: "二项式", i: "\\binom{#0}{#0}", t: "二项式系数" },
      { l: "▦", d: "矩阵", i: "\\begin{pmatrix}#0\\end{pmatrix}", t: "矩阵" },
      { l: "{□", d: "分段", i: "\\begin{cases}#0\\end{cases}", t: "分段函数" },
      { l: "x⃗", m: "<span class=\"sym-accent\"><span class=\"sym-accent-mark arrow\">→</span><span>x</span></span>", d: "向量", i: "\\vec{x}", t: "向量" },
      { l: "AB⃗", m: "<span class=\"sym-accent wide\"><span class=\"sym-accent-mark arrow\">→</span><span>AB</span></span>", d: "长向量", i: "\\overrightarrow{AB}", t: "有向线段 / 长向量" },
      { l: "x̄", m: "<span class=\"sym-overline\">x</span>", d: "上划线", i: "\\overline{x}", t: "上划线" },
      { l: "x̲", m: "<span class=\"sym-underline\">x</span>", d: "下划线", i: "\\underline{x}", t: "下划线" },
      { l: "x̂", m: "<span class=\"sym-accent\"><span class=\"sym-accent-mark\">ˆ</span><span>x</span></span>", d: "帽", i: "\\hat{x}", t: "帽（estimate）" },
      { l: "x̃", m: "<span class=\"sym-accent\"><span class=\"sym-accent-mark\">˜</span><span>x</span></span>", d: "波浪号", i: "\\tilde{x}", t: "波浪号" },
      { l: "ẋ", m: "<span class=\"sym-accent\"><span class=\"sym-accent-mark\">˙</span><span>x</span></span>", d: "一阶导", i: "\\dot{x}", t: "一阶导（点）" },
      { l: "ẍ", m: "<span class=\"sym-accent\"><span class=\"sym-accent-mark\">¨</span><span>x</span></span>", d: "二阶导", i: "\\ddot{x}", t: "二阶导（双点）" },
      { l: "⏞□", m: "<span class=\"sym-brace over\"><span>⏞</span><span>□</span></span>", d: "上花括", i: "\\overbrace{#0}", t: "上花括（标注）" },
      { l: "⏟□", m: "<span class=\"sym-brace under\"><span>□</span><span>⏟</span></span>", d: "下花括", i: "\\underbrace{#0}", t: "下花括（标注）" },
    ],
  },
  {
    name: "运算符",
    items: [
      { l: "+", i: "+" },
      { l: "−", i: "-" },
      { l: "±", i: "\\pm" },
      { l: "∓", i: "\\mp" },
      { l: "×", i: "\\times" },
      { l: "÷", i: "\\div" },
      { l: "⋅", i: "\\cdot" },
      { l: "∗", i: "\\ast" },
      { l: "⋆", i: "\\star" },
      { l: "∘", i: "\\circ" },
      { l: "∙", i: "\\bullet" },
      { l: "⊕", i: "\\oplus" },
      { l: "⊖", i: "\\ominus" },
      { l: "⊗", i: "\\otimes" },
      { l: "⊙", i: "\\odot" },
      { l: "⊘", i: "\\oslash" },
      { l: "⊞", i: "\\boxplus" },
      { l: "⊠", i: "\\boxtimes" },
      { l: "⊎", i: "\\uplus" },
      { l: "⊓", i: "\\sqcap" },
      { l: "⊔", i: "\\sqcup" },
      { l: "⋄", i: "\\diamond" },
      { l: "△", i: "\\bigtriangleup" },
      { l: "▽", i: "\\bigtriangledown" },
      { l: "◁", i: "\\triangleleft" },
      { l: "▷", i: "\\triangleright" },
      { l: "†", i: "\\dagger" },
      { l: "‡", i: "\\ddagger" },
      { l: "∖", i: "\\setminus" },
      { l: "≀", i: "\\wr" },
      { l: "⨿", i: "\\amalg" },
    ],
  },
  {
    name: "关系符",
    items: [
      { l: "=", i: "=" },
      { l: "≠", i: "\\neq" },
      { l: "≈", i: "\\approx" },
      { l: "≡", i: "\\equiv" },
      { l: "≅", i: "\\cong" },
      { l: "∼", i: "\\sim" },
      { l: "≃", i: "\\simeq" },
      { l: "<", i: "<" },
      { l: ">", i: ">" },
      { l: "≤", i: "\\leq" },
      { l: "≥", i: "\\geq" },
      { l: "≪", i: "\\ll" },
      { l: "≫", i: "\\gg" },
      { l: "≺", i: "\\prec" },
      { l: "≻", i: "\\succ" },
      { l: "≼", i: "\\preceq" },
      { l: "≽", i: "\\succeq" },
      { l: "≰", i: "\\nleq" },
      { l: "≱", i: "\\ngeq" },
      { l: "∝", i: "\\propto" },
      { l: "⊥", i: "\\perp" },
      { l: "∥", i: "\\parallel" },
      { l: "≐", i: "\\doteq" },
      { l: "≍", i: "\\asymp" },
      { l: "≜", i: "\\triangleq", t: "定义为" },
      { l: "⊨", i: "\\models" },
      { l: "⊢", i: "\\vdash" },
      { l: "⊣", i: "\\dashv" },
      { l: "⋈", i: "\\bowtie" },
    ],
  },
  {
    name: "箭头",
    items: [
      { l: "→", i: "\\rightarrow" },
      { l: "←", i: "\\leftarrow" },
      { l: "↔", i: "\\leftrightarrow" },
      { l: "↑", i: "\\uparrow" },
      { l: "↓", i: "\\downarrow" },
      { l: "↕", i: "\\updownarrow" },
      { l: "⇑", i: "\\Uparrow" },
      { l: "⇓", i: "\\Downarrow" },
      { l: "⇕", i: "\\Updownarrow" },
      { l: "⇒", i: "\\Rightarrow" },
      { l: "⇐", i: "\\Leftarrow" },
      { l: "⇔", i: "\\Leftrightarrow" },
      { l: "↦", i: "\\mapsto" },
      { l: "⟼", i: "\\longmapsto" },
      { l: "⇌", i: "\\rightleftharpoons", t: "可逆 / 平衡" },
      { l: "⟶", i: "\\longrightarrow" },
      { l: "⟵", i: "\\longleftarrow" },
      { l: "↠", i: "\\twoheadrightarrow" },
      { l: "↗", i: "\\nearrow" },
      { l: "↘", i: "\\searrow" },
      { l: "↖", i: "\\nwarrow" },
      { l: "↙", i: "\\swarrow" },
      { l: "⇀", i: "\\rightharpoonup" },
      { l: "↪", i: "\\hookrightarrow" },
    ],
  },
  {
    name: "大型运算",
    items: [
      { l: "∑", i: "\\sum_{#0}^{#0}", t: "求和" },
      { l: "∏", i: "\\prod_{#0}^{#0}", t: "连乘" },
      { l: "∐", i: "\\coprod_{#0}^{#0}", t: "余积" },
      { l: "∫", i: "\\int_{#0}^{#0}", t: "积分" },
      { l: "∬", i: "\\iint", t: "二重积分" },
      { l: "∭", i: "\\iiint", t: "三重积分" },
      { l: "∮", i: "\\oint", t: "环路积分" },
      { l: "⋃", i: "\\bigcup_{#0}^{#0}", t: "并" },
      { l: "⋂", i: "\\bigcap_{#0}^{#0}", t: "交" },
      { l: "⊔", i: "\\bigsqcup_{#0}^{#0}", t: "不交并" },
      { l: "⋁", i: "\\bigvee_{#0}^{#0}", t: "析取" },
      { l: "⋀", i: "\\bigwedge_{#0}^{#0}", t: "合取" },
      { l: "⨄", i: "\\biguplus_{#0}^{#0}" },
      { l: "⨁", i: "\\bigoplus_{#0}^{#0}" },
      { l: "⨂", i: "\\bigotimes_{#0}^{#0}" },
      { l: "⨀", i: "\\bigodot_{#0}^{#0}" },
      { l: "lim", i: "\\lim_{#0}", t: "极限" },
      { l: "limsup", i: "\\limsup_{#0}", t: "上极限" },
      { l: "liminf", i: "\\liminf_{#0}", t: "下极限" },
      { l: "max", i: "\\max_{#0}", t: "最大值" },
      { l: "min", i: "\\min_{#0}", t: "最小值" },
      { l: "sup", i: "\\sup_{#0}", t: "上确界" },
      { l: "inf", i: "\\inf_{#0}", t: "下确界" },
    ],
  },
  {
    name: "集合逻辑",
    items: [
      { l: "∈", i: "\\in" },
      { l: "∉", i: "\\notin" },
      { l: "∋", i: "\\ni" },
      { l: "⊂", i: "\\subset" },
      { l: "⊆", i: "\\subseteq" },
      { l: "⊊", i: "\\subsetneq" },
      { l: "⊄", i: "\\not\\subset" },
      { l: "⊃", i: "\\supset" },
      { l: "⊇", i: "\\supseteq" },
      { l: "∪", i: "\\cup" },
      { l: "∩", i: "\\cap" },
      { l: "∅", i: "\\emptyset" },
      { l: "∁", i: "\\complement", t: "补集" },
      { l: "∀", i: "\\forall" },
      { l: "∃", i: "\\exists" },
      { l: "∄", i: "\\nexists" },
      { l: "∧", i: "\\land" },
      { l: "∨", i: "\\lor" },
      { l: "¬", i: "\\neg" },
      { l: "⊤", i: "\\top", t: "真" },
      { l: "⊥", i: "\\bot", t: "假" },
      { l: "⟹", i: "\\implies" },
      { l: "⟺", i: "\\iff" },
      { l: "∴", i: "\\therefore" },
      { l: "∵", i: "\\because" },
      { l: "ℕ", i: "\\mathbb{N}", t: "自然数集" },
      { l: "ℤ", i: "\\mathbb{Z}", t: "整数集" },
      { l: "ℚ", i: "\\mathbb{Q}", t: "有理数集" },
      { l: "ℝ", i: "\\mathbb{R}", t: "实数集" },
      { l: "ℂ", i: "\\mathbb{C}", t: "复数集" },
    ],
  },
  {
    name: "希腊字母",
    items: [
      { l: "α", i: "\\alpha" },
      { l: "β", i: "\\beta" },
      { l: "γ", i: "\\gamma" },
      { l: "δ", i: "\\delta" },
      { l: "ε", i: "\\varepsilon" },
      { l: "ϵ", i: "\\epsilon" },
      { l: "ζ", i: "\\zeta" },
      { l: "η", i: "\\eta" },
      { l: "θ", i: "\\theta" },
      { l: "ϑ", i: "\\vartheta" },
      { l: "ι", i: "\\iota" },
      { l: "κ", i: "\\kappa" },
      { l: "λ", i: "\\lambda" },
      { l: "μ", i: "\\mu" },
      { l: "ν", i: "\\nu" },
      { l: "ξ", i: "\\xi" },
      { l: "π", i: "\\pi" },
      { l: "ϖ", i: "\\varpi" },
      { l: "ρ", i: "\\rho" },
      { l: "ς", i: "\\varsigma" },
      { l: "σ", i: "\\sigma" },
      { l: "τ", i: "\\tau" },
      { l: "υ", i: "\\upsilon" },
      { l: "φ", i: "\\varphi" },
      { l: "ϕ", i: "\\phi" },
      { l: "χ", i: "\\chi" },
      { l: "ψ", i: "\\psi" },
      { l: "ω", i: "\\omega" },
      { l: "Γ", i: "\\Gamma" },
      { l: "Δ", i: "\\Delta" },
      { l: "Θ", i: "\\Theta" },
      { l: "Λ", i: "\\Lambda" },
      { l: "Ξ", i: "\\Xi" },
      { l: "Π", i: "\\Pi" },
      { l: "Σ", i: "\\Sigma" },
      { l: "Φ", i: "\\Phi" },
      { l: "Ψ", i: "\\Psi" },
      { l: "Ω", i: "\\Omega" },
    ],
  },
  {
    name: "其它",
    items: [
      { l: "∂", i: "\\partial", t: "偏导" },
      { l: "∇", i: "\\nabla", t: "梯度算子" },
      { l: "∞", i: "\\infty", t: "无穷" },
      { l: "′", i: "\\prime", t: "撇号" },
      { l: "∠", i: "\\angle", t: "角" },
      { l: "°", i: "^\\circ", t: "度" },
      { l: "∡", i: "\\measuredangle" },
      { l: "△", i: "\\triangle" },
      { l: "□", i: "\\square" },
      { l: "⋯", i: "\\cdots", t: "居中省略号" },
      { l: "…", i: "\\ldots", t: "底部省略号" },
      { l: "⋮", i: "\\vdots", t: "竖向省略号" },
      { l: "⋱", i: "\\ddots", t: "斜向省略号" },
      { l: "∢", i: "\\sphericalangle" },
      { l: "ℏ", i: "\\hbar" },
      { l: "ℓ", i: "\\ell" },
      { l: "ℜ", i: "\\Re" },
      { l: "ℑ", i: "\\Im" },
      { l: "℘", i: "\\wp" },
      { l: "ℵ", i: "\\aleph" },
      { l: "ı", i: "\\imath" },
      { l: "ȷ", i: "\\jmath" },
      { l: "♯", i: "\\sharp" },
      { l: "♭", i: "\\flat" },
      { l: "♮", i: "\\natural" },
      { l: "✓", i: "\\checkmark" },
      { l: "%", i: "\\%" },
      { l: "∎", i: "\\blacksquare", t: "证毕" },
    ],
  },
];

const $ = (id) => document.getElementById(id);
const boot = window.__gmathBoot || { start: performance.now(), marks: [] };
const bootMark = (name) => {
  boot.marks.push({ name, ms: Math.round(performance.now() - boot.start) });
};
const bootReport = () =>
  ["GMath 启动耗时:", ...boot.marks.map((m) => `- ${m.name}: ${m.ms} ms`)].join("\n");

// ===== 国际化（中文 / English，本机记忆选择） =====
let lang = "zh";
try {
  const saved = localStorage.getItem("gmath.lang");
  if (saved === "en" || saved === "zh") lang = saved;
} catch { /* localStorage 不可用时默认中文 */ }

const I18N = {
  zh: {
    appTitle: "公式",
    ebSymbols: "符号速选",
    ebImage: "图片转公式 · AI 识别",
    ebLatex: "LaTeX 源 · 双向同步",
    apiSettings: "⚙ 接口设置",
    apiSettingsTitle: "配置大模型接口",
    aiDropTitle: "粘贴或点击选择图片",
    aiHint: "粘贴截图（Ctrl/⌘+V）<br />或点击 / 拖入图片",
    aiUrlLabel: "API 地址",
    aiKeyLabel: "API Key",
    aiModelLabel: "模型（需支持图片/视觉）",
    aiUrlPh: "https://api.openai.com/v1/chat/completions",
    aiModelPh: "如 gpt-4o / glm-4v / qwen-vl-max",
    aiSave: "保存设置",
    aiNote: "仅存于本机浏览器；识别时把图片与 Key 发送到你填写的接口。",
    insert: "插入到 Word",
    clear: "清空",
    inline: "行内",
    display: "行间",
    numbered: "右编号",
    inlineTitle: "行内公式：嵌入正文行中",
    displayTitle: "行间公式：独立居中成行",
    numberedTitle: "右编号公式：居中成行，右侧加编号",
    eqNumberPh: "留空＝自动编号（Word SEQ，会自动续号）；或自定义如 (3.1)",
    latexPh: "在此输入 LaTeX",
    debugSummary: "查看转换得到的 OMML / MathML（调试用）",
    // 动态消息
    modeLinked: "已连接文档中的公式 —— 在此修改会自动同步回去",
    modeNew: "新建公式（点选文档里的公式可载入编辑）",
    emptyFormula: "公式为空，请先输入内容。",
    insertedLinked: "已插入并连接该公式，之后在面板里改会自动同步。",
    insertFail: "插入失败：",
    syncedDoc: "已同步到文档。",
    syncFail: "同步失败：",
    loadedSel: "已载入选中的公式，可直接修改（自动同步）。",
    aiNeedCfg: "请先在「接口设置」里填写 API 地址、Key 和模型。",
    aiBusy: "识别中…（正在调用大模型）",
    aiHttp: "接口返回 ",
    aiNoFormula: "没识别出公式，换一张更清晰、只含公式的图片再试。",
    aiDone: "识别完成，已载入编辑器，可修改后点「插入到 Word」。",
    aiReqFail: "请求失败：",
    aiReqFailHint: "（常见原因：接口不允许浏览器跨域 CORS，或网络/API 地址有误。可临时运行 npm run serve 使用本地代理兜底。）",
    aiSaved: "已保存接口设置。现在可以粘贴/拖入图片识别了。",
    readImgFail: "读取图片失败",
    parseImgFail: "图片解析失败",
    procImgFail: "处理图片失败。",
    startReady: "就绪。新建公式，或点选文档里的公式来编辑。",
    startNonWord: "（非 Word 环境）编辑器可用，但只有在 Word 中才能插入/同步。",
    loadingEditor: "正在加载公式编辑器…",
    officeUnavailable: "Office.js 未加载，无法初始化 Word 加载项。请检查网络后重新打开任务面板。",
    officeWaiting: "正在等待 Office 初始化，若长期不变通常是 Office.js 或 Word WebView 启动较慢。",
    mathSlow: "公式编辑器加载较慢，通常是 MathLive CDN 下载或注册耗时。请稍候；若长期不变，请检查网络或刷新任务面板。",
    aiSysPrompt:
      "你是数学公式 OCR。把图片中的数学公式转成 LaTeX，只输出 LaTeX 本身，" +
      "不要 $ 定界符、不要代码块、不要任何解释或多余文字。",
    aiUserPrompt: "识别图片中的数学公式，只输出 LaTeX。",
  },
  en: {
    appTitle: "Formula",
    ebSymbols: "Symbols",
    ebImage: "Image → Formula · AI",
    ebLatex: "LaTeX source · two-way sync",
    apiSettings: "⚙ API settings",
    apiSettingsTitle: "Configure the LLM endpoint",
    aiDropTitle: "Paste, or click to choose an image",
    aiHint: "Paste a screenshot (Ctrl/⌘+V)<br />or click / drop an image",
    aiUrlLabel: "API URL",
    aiKeyLabel: "API Key",
    aiModelLabel: "Model (must support vision)",
    aiUrlPh: "https://api.openai.com/v1/chat/completions",
    aiModelPh: "e.g. gpt-4o / glm-4v / qwen-vl-max",
    aiSave: "Save",
    aiNote: "Stored only in this browser; the image and key are sent to the endpoint you set.",
    insert: "Insert into Word",
    clear: "Clear",
    inline: "Inline",
    display: "Display",
    numbered: "Numbered",
    inlineTitle: "Inline equation: within the text line",
    displayTitle: "Display equation: centered on its own line",
    numberedTitle: "Numbered: centered line with a number on the right",
    eqNumberPh: "Empty = auto-number (Word SEQ); or custom like (3.1)",
    latexPh: "Type LaTeX here",
    debugSummary: "Show converted OMML / MathML (debug)",
    // dynamic
    modeLinked: "Linked to an equation in the document — edits here sync back automatically.",
    modeNew: "New equation (click an equation in the document to load it).",
    emptyFormula: "The equation is empty — type something first.",
    insertedLinked: "Inserted and linked; further edits here sync automatically.",
    insertFail: "Insert failed: ",
    syncedDoc: "Synced to the document.",
    syncFail: "Sync failed: ",
    loadedSel: "Loaded the selected equation; edit it directly (auto-sync).",
    aiNeedCfg: "Fill in API URL, Key and Model under “API settings” first.",
    aiBusy: "Recognizing… (calling the model)",
    aiHttp: "Endpoint returned ",
    aiNoFormula: "No formula recognized — try a clearer image with only the formula.",
    aiDone: "Done — loaded into the editor; edit and click “Insert into Word”.",
    aiReqFail: "Request failed: ",
    aiReqFailHint: " (often the endpoint blocks browser CORS, or the API URL/network is wrong. Run npm run serve temporarily to use the local proxy fallback.)",
    aiSaved: "Settings saved. You can now paste/drop an image to recognize.",
    readImgFail: "Failed to read the image",
    parseImgFail: "Failed to decode the image",
    procImgFail: "Failed to process the image.",
    startReady: "Ready. Create a new equation, or click an equation in the document to edit it.",
    startNonWord: "(Not in Word) The editor works, but inserting/syncing only works inside Word.",
    loadingEditor: "Loading the equation editor…",
    officeUnavailable: "Office.js did not load, so the Word add-in can't initialize. Check your network and reopen the task pane.",
    officeWaiting: "Waiting for Office to initialize; if it stalls, Office.js or the Word WebView is usually slow to start.",
    mathSlow: "The equation editor is slow to load (usually MathLive CDN download/registration). Please wait; if it stays, check your network or refresh the task pane.",
    aiSysPrompt:
      "You are a math OCR engine. Convert the math in the image to LaTeX. " +
      "Output only the LaTeX itself — no $ delimiters, no code fences, no explanation.",
    aiUserPrompt: "Recognize the math formula in the image and output only LaTeX.",
  },
};

// 分类名 / 悬停提示的英文（不改动 SYMBOL_CATEGORIES 数据本身）
const CAT_EN = ["Common", "Operators", "Relations", "Arrows", "Large operators", "Sets & Logic", "Greek", "Other"];
const TIP_EN = {
  "分数": "Fraction", "平方根": "Square root", "n 次根": "nth root",
  "上标": "Superscript", "下标": "Subscript", "上下标": "Sub & superscript",
  "圆括号": "Parentheses", "方括号": "Brackets", "花括号": "Braces",
  "绝对值": "Absolute value", "范数": "Norm", "向下取整": "Floor", "向上取整": "Ceiling",
  "二项式系数": "Binomial coefficient", "矩阵": "Matrix", "分段函数": "Piecewise (cases)",
  "向量": "Vector", "有向线段 / 长向量": "Directed segment / long vector",
  "上划线": "Overline", "下划线": "Underline", "帽（estimate）": "Hat (estimate)",
  "波浪号": "Tilde", "一阶导（点）": "First derivative (dot)", "二阶导（双点）": "Second derivative (double dot)",
  "上花括（标注）": "Overbrace (annotation)", "下花括（标注）": "Underbrace (annotation)",
  "定义为": "Defined as", "可逆 / 平衡": "Reversible / equilibrium",
  "求和": "Summation", "连乘": "Product", "余积": "Coproduct", "积分": "Integral",
  "二重积分": "Double integral", "三重积分": "Triple integral", "环路积分": "Contour integral",
  "并": "Union", "交": "Intersection", "不交并": "Disjoint union",
  "析取": "Disjunction (OR)", "合取": "Conjunction (AND)",
  "极限": "Limit", "上极限": "Limit superior", "下极限": "Limit inferior",
  "最大值": "Maximum", "最小值": "Minimum", "上确界": "Supremum", "下确界": "Infimum",
  "补集": "Complement", "真": "True", "假": "False",
  "自然数集": "Natural numbers", "整数集": "Integers", "有理数集": "Rationals",
  "实数集": "Reals", "复数集": "Complex numbers",
  "偏导": "Partial derivative", "梯度算子": "Gradient (nabla)", "无穷": "Infinity",
  "撇号": "Prime", "角": "Angle", "度": "Degree",
  "居中省略号": "Centered ellipsis", "底部省略号": "Baseline ellipsis",
  "竖向省略号": "Vertical ellipsis", "斜向省略号": "Diagonal ellipsis", "证毕": "QED",
};

const T = (key) => (I18N[lang] && I18N[lang][key] != null ? I18N[lang][key] : key);
const catLabel = (idx) => (lang === "en" ? CAT_EN[idx] : SYMBOL_CATEGORIES[idx].name);
const tipText = (zh) => (lang === "en" ? TIP_EN[zh] || zh : zh);

function applyI18n() {
  document.querySelectorAll("[data-i18n]").forEach((el) => {
    const v = I18N[lang][el.dataset.i18n];
    if (v != null) el.textContent = v;
  });
  document.querySelectorAll("[data-i18n-html]").forEach((el) => {
    const v = I18N[lang][el.dataset.i18nHtml];
    if (v != null) el.innerHTML = v;
  });
  document.querySelectorAll("[data-i18n-ph]").forEach((el) => {
    const v = I18N[lang][el.dataset.i18nPh];
    if (v != null) el.placeholder = v;
  });
  document.querySelectorAll("[data-i18n-title]").forEach((el) => {
    const v = I18N[lang][el.dataset.i18nTitle];
    if (v != null) el.title = v;
  });
  document.documentElement.lang = lang === "en" ? "en" : "zh-CN";
  const lb = $("langBtn");
  if (lb) lb.textContent = lang === "zh" ? "EN" : "中";
  refreshPaletteLang();
  if (modeLine) setMode(linked); // 模式徽章按当前连接状态重译
}

function refreshPaletteLang() {
  const tabs = $("catTabs");
  if (!tabs) return;
  tabs.querySelectorAll("button[data-cat]").forEach((b) => {
    b.textContent = catLabel(Number(b.dataset.cat));
  });
  const active = tabs.querySelector("button.active");
  if (active) showPaletteCat(Number(active.dataset.cat));
}

function setLang(l) {
  lang = l === "en" ? "en" : "zh";
  try { localStorage.setItem("gmath.lang", lang); } catch { /* ignore */ }
  applyI18n();
}

let mathfield, latexEl, statusEl, debugEl, insertBtn, layoutModeEl, eqNumberEl, eqNumberRow, modeLine;

// 同步状态机
let linked = false; // 面板当前是否“连接”着文档里的某个公式
let loadedLatex = null; // 最近一次从文档读到的 LaTeX（用于判断是否同一个公式、避免误重载）
let applying = false; // 正在回写（屏蔽选区监听，防自我触发）
let loadingDoc = false; // 正在把文档公式灌入面板（屏蔽 mathfield→回写）
let selTimer = null;
let syncTimer = null;

bootMark("taskpane module evaluating");

function setStatus(msg, kind = "") {
  statusEl.textContent = msg;
  statusEl.className = "status" + (kind ? " " + kind : "");
}

function setStartupStatus(msg, kind = "") {
  const el = $("status");
  if (!el) return;
  el.textContent = msg;
  el.className = "status" + (kind ? " " + kind : "");
}

function writeStartupDebug(prefix = "") {
  const el = $("debug");
  if (!el) return;
  el.value = (prefix ? prefix + "\n\n" : "") + bootReport();
}

function setMode(isLinked) {
  linked = isLinked;
  if (isLinked) {
    modeLine.textContent = T("modeLinked");
    modeLine.classList.add("linked");
  } else {
    loadedLatex = null;
    modeLine.textContent = T("modeNew");
    modeLine.classList.remove("linked");
  }
}

function currentOmml() {
  const mathml = mathfield.getValue("math-ml");
  if (!mathml || !mathml.trim()) return null;
  return mml2omml(mathml);
}

// 生成右编号的文字 run：留空→自动编号（Word SEQ 域，会自动续号）；否则用字面文本
function buildNumberRuns(numberText) {
  const t = (numberText || "").trim();
  if (t) {
    return `<w:r><w:t xml:space="preserve">${escXml(t)}</w:t></w:r>`;
  }
  // ( SEQ equation \* ARABIC ) → (1) (2) …，由 Word 维护并可整篇续号
  return (
    `<w:r><w:t xml:space="preserve">(</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> SEQ equation \\* ARABIC </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t>1</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
    `<w:r><w:t xml:space="preserve">)</w:t></w:r>`
  );
}

// layout: "inline"（行内）| "display"（行间居中）| "numbered"（居中＋右侧编号）
function buildFlatOpc(ommlMath, layout, numberText) {
  let oMath = ommlMath.replace(/^\s*<\?xml[^>]*\?>\s*/i, "").trim();
  let mathBlock;
  if (layout === "display") {
    mathBlock = `<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>${oMath}</m:oMathPara>`;
  } else if (layout === "numbered") {
    // 用「相对页边距的定位制表符」实现：公式居中、编号靠右贴右边距，与纸张/页宽无关
    mathBlock =
      `<w:r><w:ptab w:relativeTo="margin" w:alignment="center" w:leader="none"/></w:r>` +
      oMath +
      `<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>` +
      buildNumberRuns(numberText);
  } else {
    mathBlock = oMath; // 行内
  }
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

// 单个符号/模板按钮：文本标签 l + 点击插入模板 i（#0 → 插入后的光标占位）
function renderSymButton(it, catIdx) {
  const b = document.createElement("button");
  b.type = "button";
  b.dataset.insert = it.i;
  const title = it.t ? tipText(it.t) : it.i;
  b.title = title;
  b.setAttribute("aria-label", title);
  if (catIdx === 0 && it.d) {
    b.className = "template";
    const main = document.createElement("span");
    main.className = "sym-main";
    if (it.m) main.innerHTML = it.m;
    else main.textContent = it.l;
    const desc = document.createElement("span");
    desc.className = "sym-desc";
    desc.textContent = tipText(it.d);
    b.append(main, desc);
  } else {
    b.textContent = it.l;
  }
  return b;
}

function showPaletteCat(idx) {
  const grid = $("symGrid");
  if (!grid) return;
  grid.innerHTML = "";
  grid.hidden = false;
  grid.classList.toggle("common", idx === 0);
  SYMBOL_CATEGORIES[idx].items.forEach((it) => grid.appendChild(renderSymButton(it, idx)));
}

function collapsePalette() {
  const grid = $("symGrid");
  if (!grid) return;
  grid.innerHTML = "";
  grid.hidden = true;
}

// 渲染分类符号速选：上方分类标签，点选某类后在下方展开该类符号
function renderPalette() {
  const tabs = $("catTabs");
  if (!tabs || !$("symGrid")) return;

  SYMBOL_CATEGORIES.forEach((cat, idx) => {
    const t = document.createElement("button");
    t.type = "button";
    t.textContent = catLabel(idx);
    t.dataset.cat = String(idx);
    if (idx === 0) t.classList.add("active");
    tabs.appendChild(t);
  });

  tabs.addEventListener("click", (ev) => {
    const t = ev.target.closest("button[data-cat]");
    if (!t) return;
    const isOpen = t.classList.contains("active");
    tabs.querySelectorAll("button").forEach((b) => {
      b.classList.remove("active");
      b.setAttribute("aria-expanded", "false");
    });
    if (isOpen) {
      collapsePalette();
      return;
    }
    t.classList.add("active");
    t.setAttribute("aria-expanded", "true");
    showPaletteCat(Number(t.dataset.cat));
  });

  tabs.querySelector("button[data-cat='0']")?.setAttribute("aria-expanded", "true");
  showPaletteCat(0);
}

// 符号速选是纯 DOM，模块加载即渲染，立即可见可切换（此刻 MathLive 可能未就绪，
// 先用文本回退；待 wireUp 中 MathLive 确认就绪后再重渲染为真实公式预览）。
renderPalette();
applyI18n(); // 按记忆的语言完成首次本地化（含分类标签）

// 当前选中的版式：inline / display / numbered
function getLayout() {
  const active = layoutModeEl.querySelector("button.active");
  return active ? active.dataset.mode : "display";
}

function setLayout(mode) {
  layoutModeEl.querySelectorAll("button").forEach((b) => {
    b.classList.toggle("active", b.dataset.mode === mode);
  });
  eqNumberRow.hidden = mode !== "numbered"; // 编号输入只在「右编号」时出现
  scheduleSync();
}

// 把一段 LaTeX 灌入面板（不触发回写）
function loadIntoPane(latex, display) {
  loadingDoc = true;
  mathfield.setValue(latex);
  latexEl.value = mathfield.getValue("latex");
  if (typeof display === "boolean") setLayout(display ? "display" : "inline");
  loadingDoc = false;
}

// ===== 面板 → 文档：插入一个新公式（不连接） =====
async function insertNew() {
  const omml = currentOmml();
  if (!omml) {
    setStatus(T("emptyFormula"), "err");
    return;
  }
  const flatOpc = buildFlatOpc(omml, getLayout(), eqNumberEl.value);
  debugEl.value = "构建版本: " + BUILD + "\n\nOMML:\n" + omml + "\n\n完整 OOXML:\n" + flatOpc;
  applying = true;
  try {
    await Word.run(async (context) => {
      const r = context.document.getSelection().insertOoxml(flatOpc, Word.InsertLocation.replace);
      r.select(); // 选中新公式，便于随后继续在面板里改、自动同步
      await context.sync();
    });
    loadedLatex = mathfield.getValue("latex");
    setMode(true);
    setStatus(T("insertedLinked"), "ok");
  } catch (e) {
    setStatus(T("insertFail") + (e.code || "") + " " + (e.message || e), "err");
    debugEl.value += "\n\n=== 错误详情 ===\n" + JSON.stringify(describeError(e), null, 2);
  } finally {
    setTimeout(() => (applying = false), 500);
  }
}

// ===== 面板 → 文档：把改动同步回当前选中的公式（替换选区） =====
async function syncToDoc() {
  if (!linked) return;
  const omml = currentOmml();
  if (!omml) return;
  const flatOpc = buildFlatOpc(omml, getLayout(), eqNumberEl.value);
  applying = true;
  try {
    await Word.run(async (context) => {
      const r = context.document.getSelection().insertOoxml(flatOpc, Word.InsertLocation.replace);
      r.select(); // 替换后重新选中，保证下一次还能替换到它
      await context.sync();
    });
    loadedLatex = mathfield.getValue("latex");
    setStatus(T("syncedDoc"), "ok");
  } catch (e) {
    setStatus(T("syncFail") + (e.code || "") + " " + (e.message || e), "err");
  } finally {
    setTimeout(() => (applying = false), 500);
  }
}

function scheduleSync() {
  if (!linked || loadingDoc) return;
  clearTimeout(syncTimer);
  syncTimer = setTimeout(syncToDoc, 600);
}

// ===== 文档 → 面板：选区变化时，若选中的是公式就载入面板 =====
function onSelectionChanged() {
  if (applying) return;
  clearTimeout(selTimer);
  selTimer = setTimeout(readFromSelection, 250);
}

async function readFromSelection() {
  if (applying) return;
  try {
    let ooxml = null;
    await Word.run(async (context) => {
      const res = context.document.getSelection().getOoxml();
      await context.sync();
      ooxml = res.value;
    });
    const omath = extractOMath(ooxml);
    if (!omath) {
      if (linked) setMode(false); // 选区离开了公式 → 断开，避免误同步
      return;
    }
    const latex = omml2latex(omath);
    if (linked && latex === loadedLatex) return; // 还是同一个公式，别覆盖正在改的内容
    loadIntoPane(latex, /oMathPara/.test(omath));
    loadedLatex = latex;
    setMode(true);
    setStatus(T("loadedSel"), "ok");
  } catch (e) {
    debugEl.value = "[选区读取] " + (e.message || e) + "\n" + debugEl.value;
  }
}

// ===== 图片转公式（AI 识别，OpenAI 兼容视觉接口） =====
const AI_KEYS = { url: "gmath.ai.url", key: "gmath.ai.key", model: "gmath.ai.model" };

function lsGet(k) {
  try { return localStorage.getItem(k) || ""; } catch { return ""; }
}
function lsSet(k, v) {
  try { localStorage.setItem(k, v); } catch { /* 某些 WebView 禁用 localStorage */ }
}

function setAiStatus(msg, kind = "") {
  const el = $("aiStatus");
  if (!el) return;
  el.hidden = !msg;
  el.textContent = msg || "";
  el.className = "ai-status" + (kind ? " " + kind : "");
}

// 把接口地址规整成 chat/completions 端点
function normalizeEndpoint(url) {
  let u = url.trim().replace(/\/+$/, "");
  if (/\/chat\/completions$/.test(u)) return u;
  if (/\/v\d+$/.test(u)) return u + "/chat/completions";
  return u + "/v1/chat/completions";
}

function aiRequestUrl() {
  if (["localhost", "127.0.0.1", "[::1]"].includes(window.location.hostname)) {
    return "/api/ai/chat/completions";
  }
  return "https://localhost:3000/api/ai/chat/completions";
}

function aiPayload(model, dataUrl) {
  return JSON.stringify({
    model,
    temperature: 0,
    messages: [
      { role: "system", content: T("aiSysPrompt") },
      {
        role: "user",
        content: [
          { type: "text", text: T("aiUserPrompt") },
          { type: "image_url", image_url: { url: dataUrl } },
        ],
      },
    ],
  });
}

async function postAi(endpoint, key, body, useProxy) {
  const headers = {
    "Content-Type": "application/json",
    Authorization: "Bearer " + key,
  };
  if (useProxy) headers["X-GMath-AI-Endpoint"] = endpoint;
  return fetch(useProxy ? aiRequestUrl() : endpoint, {
    method: "POST",
    headers,
    body,
  });
}

// 清洗模型返回：去掉代码块围栏与 $ / \[ \] / \( \) 定界符
function cleanLatex(raw) {
  let s = (raw || "").trim();
  s = s.replace(/^```(?:latex|tex|math)?\s*/i, "").replace(/\s*```$/, "").trim();
  s = s
    .replace(/^\$\$([\s\S]*?)\$\$$/, "$1")
    .replace(/^\\\[([\s\S]*?)\\\]$/, "$1")
    .replace(/^\\\(([\s\S]*?)\\\)$/, "$1")
    .replace(/^\$([\s\S]*?)\$$/, "$1")
    .trim();
  return s;
}

// 读取图片文件 → 压缩到最长边 ≤1600 的 dataURL（减小体积/费用，PNG 保边缘清晰）
function fileToDataUrl(file, maxSide = 1600) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error(T("readImgFail")));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error(T("parseImgFail")));
      img.onload = () => {
        const scale = Math.min(1, maxSide / Math.max(img.width, img.height));
        if (scale === 1) return resolve(reader.result);
        const c = document.createElement("canvas");
        c.width = Math.round(img.width * scale);
        c.height = Math.round(img.height * scale);
        c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
        resolve(c.toDataURL("image/png"));
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

async function recognizeImage(dataUrl) {
  const url = $("aiUrl").value.trim();
  const key = $("aiKey").value.trim();
  const model = $("aiModel").value.trim();
  if (!url || !key || !model) {
    setAiStatus(T("aiNeedCfg"), "err");
    $("aiSettings").hidden = false;
    return;
  }
  setAiStatus(T("aiBusy"), "busy");
  try {
    const endpoint = normalizeEndpoint(url);
    const body = aiPayload(model, dataUrl);
    let resp;
    try {
      resp = await postAi(endpoint, key, body, false);
    } catch (directErr) {
      resp = await postAi(endpoint, key, body, true).catch(() => {
        throw directErr;
      });
    }
    if (!resp.ok) {
      const body = await resp.text().catch(() => "");
      setAiStatus(`${T("aiHttp")}${resp.status}: ${body.slice(0, 200)}`, "err");
      return;
    }
    const data = await resp.json();
    const content = data?.choices?.[0]?.message?.content;
    const text = typeof content === "string"
      ? content
      : Array.isArray(content) ? content.map((c) => c?.text || "").join("") : "";
    const latex = cleanLatex(text);
    if (!latex) {
      setAiStatus(T("aiNoFormula"), "err");
      return;
    }
    mathfield.setValue(latex);
    latexEl.value = mathfield.getValue("latex");
    setMode(false); // 作为新公式载入，等用户确认后再插入
    mathfield.focus();
    setAiStatus(T("aiDone"), "ok");
  } catch (e) {
    setAiStatus(`${T("aiReqFail")}${e.message || e}${T("aiReqFailHint")}`, "err");
  }
}

async function handleImageFile(file) {
  if (!file || !file.type || !file.type.startsWith("image/")) return;
  try {
    const dataUrl = await fileToDataUrl(file);
    const thumb = $("aiThumb");
    thumb.src = dataUrl;
    thumb.hidden = false;
    $("aiHint").hidden = true;
    await recognizeImage(dataUrl);
  } catch (e) {
    setAiStatus(e.message || T("procImgFail"), "err");
  }
}

function wireAI() {
  const urlEl = $("aiUrl");
  const keyEl = $("aiKey");
  const modelEl = $("aiModel");
  const drop = $("aiDrop");
  const fileEl = $("aiFile");

  // 载入已保存的配置
  urlEl.value = lsGet(AI_KEYS.url);
  keyEl.value = lsGet(AI_KEYS.key);
  modelEl.value = lsGet(AI_KEYS.model);

  $("aiSettingsBtn").addEventListener("click", () => {
    const s = $("aiSettings");
    s.hidden = !s.hidden;
  });
  $("aiSaveBtn").addEventListener("click", () => {
    lsSet(AI_KEYS.url, urlEl.value.trim());
    lsSet(AI_KEYS.key, keyEl.value.trim());
    lsSet(AI_KEYS.model, modelEl.value.trim());
    setAiStatus(T("aiSaved"), "ok");
    $("aiSettings").hidden = true;
  });

  // 点击 / 拖拽 / 选择文件
  drop.addEventListener("click", () => fileEl.click());
  drop.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fileEl.click(); }
  });
  fileEl.addEventListener("change", () => {
    if (fileEl.files && fileEl.files[0]) handleImageFile(fileEl.files[0]);
    fileEl.value = "";
  });
  drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("drag"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("drag"));
  drop.addEventListener("drop", (e) => {
    e.preventDefault();
    drop.classList.remove("drag");
    const f = e.dataTransfer?.files?.[0];
    if (f) handleImageFile(f);
  });

  // 全局粘贴：剪贴板里有图片就拦下来识别（不影响在输入框里粘贴文本）
  document.addEventListener("paste", (e) => {
    const items = e.clipboardData?.items;
    if (!items) return;
    for (const it of items) {
      if (it.type && it.type.startsWith("image/")) {
        const f = it.getAsFile();
        if (f) { e.preventDefault(); handleImageFile(f); return; }
      }
    }
  });
}

function wireUp(inWord) {
  bootMark("wireUp start");
  mathfield = $("mathfield");
  // 关闭虚拟键盘：聚焦时不再自动弹出（角标已用 CSS 隐藏）
  mathfield.mathVirtualKeyboardPolicy = "manual";
  latexEl = $("latex");
  statusEl = $("status");
  debugEl = $("debug");
  insertBtn = $("insertBtn");
  layoutModeEl = $("layoutMode");
  eqNumberEl = $("eqNumber");
  eqNumberRow = $("eqNumberRow");
  modeLine = $("modeLine");

  latexEl.value = mathfield.getValue("latex");
  mathfield.addEventListener("input", () => {
    latexEl.value = mathfield.getValue("latex");
    scheduleSync();
  });
  latexEl.addEventListener("input", () => {
    mathfield.setValue(latexEl.value, { suppressChangeNotifications: true });
    scheduleSync();
  });
  layoutModeEl.addEventListener("click", (ev) => {
    const btn = ev.target.closest("button[data-mode]");
    if (!btn) return;
    setLayout(btn.dataset.mode);
  });
  eqNumberEl.addEventListener("input", scheduleSync);

  $("palette").addEventListener("click", (ev) => {
    const btn = ev.target.closest("button[data-insert]");
    if (!btn) return;
    mathfield.executeCommand(["insert", btn.dataset.insert]);
    mathfield.focus();
    latexEl.value = mathfield.getValue("latex");
    scheduleSync();
  });

  $("clearBtn").addEventListener("click", () => {
    mathfield.setValue("");
    latexEl.value = "";
    setMode(false); // 清空即断开连接，回到新建
    mathfield.focus();
  });

  insertBtn.addEventListener("click", insertNew);
  insertBtn.disabled = false;

  $("langBtn").addEventListener("click", () => setLang(lang === "zh" ? "en" : "zh"));

  wireAI();

  if (inWord) {
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      onSelectionChanged
    );
  }
  setMode(false);
  bootMark("wireUp complete");
  writeStartupDebug();
}

function waitForMathField() {
  const timeoutMs = 8000;
  const warning = setTimeout(() => {
    bootMark("math-field wait exceeded " + timeoutMs + " ms");
    setStartupStatus(T("mathSlow"), "err");
    writeStartupDebug("启动阶段仍在等待 <math-field> 注册。");
  }, timeoutMs);

  return customElements.whenDefined("math-field").finally(() => {
    clearTimeout(warning);
    bootMark("math-field defined");
  });
}

if (!window.Office || !Office.onReady) {
  bootMark("office.js unavailable");
  setStartupStatus(T("officeUnavailable"), "err");
  writeStartupDebug();
} else {
  const officeReadyWarning = setTimeout(() => {
    bootMark("office ready wait exceeded 8000 ms");
    setStartupStatus(T("officeWaiting"), "err");
    writeStartupDebug();
  }, 8000);

  Office.onReady((info) => {
    clearTimeout(officeReadyWarning);
    bootMark("office ready");
    setStartupStatus(T("loadingEditor"));
    return waitForMathField().then(() => start(info));
  });
}

function start(info) {
  const inWord = info.host === Office.HostType.Word;
  wireUp(inWord);
  if (inWord) setStatus(T("startReady"), "ok");
  else setStatus(T("startNonWord"));
}
