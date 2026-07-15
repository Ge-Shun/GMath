import assert from "node:assert/strict";
import test from "node:test";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { convertLatexToMathMl } from "mathlive";

globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;

const { mml2ommlDetailed } = await import("../src/mathml2omml.js");
const { omml2latexDetailed } = await import("../src/omml2latex.js");
const { latexToOmml } = await import("../src/latex2omml.js");

const formulas = [
  ["fraction", "\\frac{a}{b}", "<m:f>"],
  ["root", "\\sqrt[3]{x}", "<m:rad>"],
  ["scripts", "x_n^2", "<m:sSubSup>"],
  ["sum", "\\sum_{i=1}^{n}i", "<m:nary>"],
  ["matrix", "\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}", "<m:m>"],
  ["cases", "\\begin{cases}1&x>0\\\\0&x\\le0\\end{cases}", "<m:d>"],
];

for (const [name, latex, expected] of formulas) {
  test(`real MathLive output round-trips: ${name}`, () => {
    const mathml = convertLatexToMathMl(latex);
    const forward = mml2ommlDetailed(mathml);
    assert.equal(forward.lossy, false, forward.warnings.join(", "));
    assert.match(forward.omml, new RegExp(expected.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
    const reverse = omml2latexDetailed(forward.omml);
    assert.equal(reverse.lossy, false, reverse.warnings.join(", "));
    assert.notEqual(reverse.latex, "");
  });
}

test("real MathLive decoration path keeps the decorated content", () => {
  const omml = latexToOmml("\\frac{\\overline{x}}{2}", {
    convertLatexToMathMl,
    mml2omml: (mathml) => mml2ommlDetailed(mathml).omml,
  });
  assert.match(omml, /<m:num><m:bar>/);
  assert.doesNotMatch(omml, /&#xE000;/);
});
