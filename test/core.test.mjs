import assert from "node:assert/strict";
import test from "node:test";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;

const { AI_LIMITS, cleanLatex, normalizeEndpoint } = await import("../src/ai-client.js");
const { buildFlatOpc } = await import("../src/ooxml.js");
const { inspectOmathPackage, omml2latexDetailed } = await import("../src/omml2latex.js");
const { LatestTaskQueue, normalizeOmmlFingerprint } = await import("../src/sync-state.js");

test("normalizes API endpoints without losing query parameters", () => {
  assert.equal(
    normalizeEndpoint("https://api.example.test/v1?tenant=demo"),
    "https://api.example.test/v1/chat/completions?tenant=demo",
  );
  assert.equal(
    normalizeEndpoint("https://api.example.test/custom/chat/completions?api-version=1"),
    "https://api.example.test/custom/chat/completions?api-version=1",
  );
  assert.throws(() => normalizeEndpoint("file:///tmp/key"));
  assert.equal(AI_LIMITS.maxFileBytes, 10 * 1024 * 1024);
});

test("cleans common model wrappers from LaTeX", () => {
  assert.equal(cleanLatex("```latex\n\\frac{a}{b}\n```"), "\\frac{a}{b}");
  assert.equal(cleanLatex("$$x^2$$"), "x^2");
});

test("escapes custom equation numbers in Flat OPC", () => {
  const xml = buildFlatOpc("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>", "numbered", "<3 & 4>");
  assert.match(xml, /&lt;3 &amp; 4&gt;/);
  assert.doesNotMatch(xml, /<3 & 4>/);
});

test("inspects safe, tagged and unsafe Word selections", () => {
  const one = `<pkg:package xmlns:pkg="p" xmlns:w="w" xmlns:m="m"><w:sdt><w:sdtPr><w:tag w:val="gmath:abc"/></w:sdtPr><w:sdtContent><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:sdtContent></w:sdt></pkg:package>`;
  const inspected = inspectOmathPackage(one);
  assert.equal(inspected.safe, true);
  assert.equal(inspected.count, 1);
  assert.equal(inspected.gmathTag, "gmath:abc");

  const mixed = `<root xmlns:w="w" xmlns:m="m"><w:r><w:t>prefix</w:t></w:r><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></root>`;
  assert.equal(inspectOmathPackage(mixed).safe, false);

  const multiple = `<root xmlns:m="m"><m:oMath/><m:oMath/></root>`;
  assert.equal(inspectOmathPackage(multiple).count, 2);
});

test("reports lossy OMML instead of silently flattening it", () => {
  const result = omml2latexDetailed(`<m:oMath xmlns:m="m"><m:eqArr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:eqArr></m:oMath>`);
  assert.equal(result.latex, "x");
  assert.equal(result.lossy, true);
  assert.deepEqual(result.warnings, ["eqArr"]);
});

test("normalizes namespace and whitespace differences in fingerprints", () => {
  const a = `<m:oMath xmlns:m="urn:m"> <m:r> <m:t>x</m:t> </m:r> </m:oMath>`;
  const b = `<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>`;
  assert.equal(normalizeOmmlFingerprint(a), normalizeOmmlFingerprint(b));
});

test("LatestTaskQueue serializes work and skips queued stale tasks", async () => {
  const queue = new LatestTaskQueue();
  const events = [];
  let release;
  let markStarted;
  const gate = new Promise((resolve) => { release = resolve; });
  const started = new Promise((resolve) => { markStarted = resolve; });
  const first = queue.enqueue(async () => {
    events.push("first:start");
    markStarted();
    await gate;
    events.push("first:end");
  });
  await started;
  const stale = queue.enqueue(async () => events.push("stale"));
  const latest = queue.enqueue(async () => events.push("latest"));
  release();
  await Promise.all([first, stale, latest]);
  assert.deepEqual(events, ["first:start", "first:end", "latest"]);
});
