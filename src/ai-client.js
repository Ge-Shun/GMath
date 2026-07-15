export const AI_LIMITS = {
  maxFileBytes: 10 * 1024 * 1024,
  maxPixels: 25_000_000,
  requestTimeoutMs: 45_000,
};

export function normalizeEndpoint(input) {
  const url = new URL(String(input || "").trim());
  if (url.protocol !== "https:" && url.protocol !== "http:") {
    throw new Error("API URL must use HTTP or HTTPS");
  }
  const path = url.pathname.replace(/\/+$/, "");
  if (/\/chat\/completions$/i.test(path)) {
    url.pathname = path;
  } else if (/\/v\d+$/i.test(path)) {
    url.pathname = `${path}/chat/completions`;
  } else {
    url.pathname = `${path}/v1/chat/completions`.replace(/\/{2,}/g, "/");
  }
  return url.toString();
}

export function cleanLatex(raw) {
  let value = String(raw || "").trim();
  value = value.replace(/^```(?:latex|tex|math)?\s*/i, "").replace(/\s*```$/, "").trim();
  return value
    .replace(/^\$\$([\s\S]*?)\$\$$/, "$1")
    .replace(/^\\\[([\s\S]*?)\\\]$/, "$1")
    .replace(/^\\\(([\s\S]*?)\\\)$/, "$1")
    .replace(/^\$([\s\S]*?)\$$/, "$1")
    .trim();
}
