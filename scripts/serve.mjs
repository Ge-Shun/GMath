import crypto from "node:crypto";
import dns from "node:dns/promises";
import fs from "node:fs";
import http from "node:http";
import https from "node:https";
import net from "node:net";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const defaultRoot = path.resolve(__dirname, "..");
const MAX_REQUEST_BYTES = 25 * 1024 * 1024;
const MAX_RESPONSE_BYTES = 10 * 1024 * 1024;
const UPSTREAM_TIMEOUT_MS = 45_000;

const mimeTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".txt": "text/plain; charset=utf-8",
  ".woff2": "font/woff2",
  ".xml": "application/xml; charset=utf-8",
};

export function allowedOrigins(port, extra = "") {
  return new Set([
    "https://ge-shun.github.io",
    `https://localhost:${port}`,
    `http://localhost:${port}`,
    `https://127.0.0.1:${port}`,
    `http://127.0.0.1:${port}`,
    ...String(extra).split(",").map((value) => value.trim()).filter(Boolean),
  ]);
}

function isLocalHostHeader(value, port) {
  return new Set([
    `localhost:${port}`,
    `127.0.0.1:${port}`,
    `[::1]:${port}`,
  ]).has(String(value || "").toLowerCase());
}

export function isRequestOriginAllowed(req, origins, port) {
  const origin = req.headers.origin;
  if (origin) return origins.has(origin);
  // Same-origin GETs may omit Origin. The server itself only binds to loopback;
  // also require a loopback Host so remote webpages can't use this fallback.
  return isLocalHostHeader(req.headers.host, port);
}

function responseHeaders(req, origins, port, extra = {}) {
  const headers = {
    "Cache-Control": "no-store",
    "X-Content-Type-Options": "nosniff",
    "Referrer-Policy": "no-referrer",
    ...extra,
  };
  const origin = req.headers.origin;
  if (origin && origins.has(origin)) {
    headers["Access-Control-Allow-Origin"] = origin;
    headers.Vary = "Origin";
  }
  return headers;
}

function send(req, res, status, body, origins, port, headers = {}) {
  res.writeHead(status, responseHeaders(req, origins, port, headers));
  res.end(body);
}

function sendJson(req, res, status, payload, origins, port) {
  send(req, res, status, JSON.stringify(payload), origins, port, {
    "Content-Type": "application/json; charset=utf-8",
  });
}

function readBody(req, maxBytes = MAX_REQUEST_BYTES) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > maxBytes) {
        reject(new Error("Request body is too large"));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

export function isPrivateAddress(address) {
  const value = String(address || "").toLowerCase().split("%")[0];
  const family = net.isIP(value);
  if (family === 4) {
    const parts = value.split(".").map(Number);
    const [a, b, c] = parts;
    return (
      a === 0 || a === 10 || a === 127 || a >= 224 ||
      (a === 100 && b >= 64 && b <= 127) ||
      (a === 169 && b === 254) ||
      (a === 172 && b >= 16 && b <= 31) ||
      (a === 192 && (b === 168 || (b === 0 && (c === 0 || c === 2)))) ||
      (a === 198 && (b === 18 || b === 19 || (b === 51 && c === 100))) ||
      (a === 203 && b === 0 && c === 113)
    );
  }
  if (family === 6) {
    if (value === "::" || value === "::1") return true;
    if (value.startsWith("::ffff:")) return isPrivateAddress(value.slice(7));
    return /^(f[cd]|fe[89ab]|2001:db8)/.test(value);
  }
  return true;
}

export async function validateTarget(raw, options = {}) {
  if (!raw || Array.isArray(raw)) throw new Error("Missing target endpoint");
  const target = new URL(String(raw));
  const allowInsecure = options.allowInsecure === true;
  const allowPrivate = options.allowPrivate === true;
  if (target.protocol !== "https:" && !(allowInsecure && target.protocol === "http:")) {
    throw new Error("Target endpoint must use HTTPS");
  }
  if (target.username || target.password) throw new Error("Target endpoint must not contain credentials");
  const host = target.hostname.replace(/^\[|\]$/g, "");
  if (options.allowedHosts?.size && !options.allowedHosts.has(host.toLowerCase())) {
    throw new Error("Target endpoint host is not in GMATH_AI_HOSTS");
  }
  const addresses = net.isIP(host)
    ? [{ address: host }]
    : await (options.lookup || dns.lookup)(host, { all: true, verbatim: true });
  if (!addresses.length) throw new Error("Target endpoint did not resolve to an address");
  if (!allowPrivate && addresses.some(({ address }) => isPrivateAddress(address))) {
    throw new Error("Target endpoint resolves to a private or reserved address");
  }
  return target.toString();
}

async function readUpstreamBody(response, maxBytes = MAX_RESPONSE_BYTES) {
  const declared = Number(response.headers.get("content-length") || 0);
  if (declared > maxBytes) throw new Error("AI response is too large");
  if (!response.body) return Buffer.alloc(0);
  const reader = response.body.getReader();
  const chunks = [];
  let size = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    size += value.byteLength;
    if (size > maxBytes) {
      await reader.cancel();
      throw new Error("AI response is too large");
    }
    chunks.push(Buffer.from(value));
  }
  return Buffer.concat(chunks);
}

function tokenMatches(actual, expected) {
  const left = Buffer.from(String(actual || ""));
  const right = Buffer.from(String(expected));
  return left.length === right.length && crypto.timingSafeEqual(left, right);
}

export function resolveStaticPath(root, requestPath) {
  const normalized = requestPath === "/" ? "/src/taskpane.html" : requestPath;
  const allowed = normalized === "/manifest.xml"
    || normalized.startsWith("/src/")
    || normalized.startsWith("/assets/");
  if (!allowed) return null;
  const filePath = path.resolve(root, `.${normalized}`);
  if (!filePath.startsWith(`${root}${path.sep}`)) return null;
  return filePath;
}

export function createRequestHandler(options = {}) {
  const root = options.root || defaultRoot;
  const port = Number(options.port || 3000);
  const origins = options.origins || allowedOrigins(port, process.env.GMATH_ALLOWED_ORIGINS);
  const proxyToken = options.proxyToken || crypto.randomBytes(32).toString("hex");
  const allowPrivate = options.allowPrivate ?? process.env.GMATH_ALLOW_PRIVATE_AI === "1";
  const allowInsecure = options.allowInsecure ?? process.env.GMATH_ALLOW_INSECURE_AI === "1";
  const configuredHosts = options.allowedHosts ?? process.env.GMATH_AI_HOSTS;
  const aiHosts = configuredHosts instanceof Set
    ? configuredHosts
    : new Set(String(configuredHosts || "").split(",").map((host) => host.trim().toLowerCase()).filter(Boolean));

  async function proxyAi(req, res) {
    if (!isRequestOriginAllowed(req, origins, port)) {
      sendJson(req, res, 403, { error: { message: "Origin is not allowed" } }, origins, port);
      return;
    }
    if (!tokenMatches(req.headers["x-gmath-proxy-token"], proxyToken)) {
      sendJson(req, res, 403, { error: { message: "Invalid proxy token" } }, origins, port);
      return;
    }

    let body;
    let target;
    try {
      body = await readBody(req);
      target = await validateTarget(req.headers["x-gmath-ai-endpoint"], {
        allowPrivate,
        allowInsecure,
        allowedHosts: aiHosts,
      });
    } catch (error) {
      if (!res.destroyed) sendJson(req, res, 400, { error: { message: error.message || String(error) } }, origins, port);
      return;
    }

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), UPSTREAM_TIMEOUT_MS);
    try {
      const upstream = await fetch(target, {
        method: "POST",
        headers: {
          "Content-Type": req.headers["content-type"] || "application/json",
          Authorization: req.headers.authorization || "",
        },
        body,
        redirect: "error",
        signal: controller.signal,
      });
      const upstreamBody = await readUpstreamBody(upstream);
      send(req, res, upstream.status, upstreamBody, origins, port, {
        "Content-Type": upstream.headers.get("content-type") || "application/json; charset=utf-8",
      });
    } catch (error) {
      const status = controller.signal.aborted ? 504 : 502;
      sendJson(req, res, status, {
        error: { message: `Local proxy could not reach the AI endpoint: ${error.message || String(error)}` },
      }, origins, port);
    } finally {
      clearTimeout(timeout);
    }
  }

  function serveStatic(req, res) {
    let pathname;
    try {
      const requestUrl = new URL(req.url || "/", `https://localhost:${port}`);
      pathname = decodeURIComponent(requestUrl.pathname);
    } catch {
      send(req, res, 400, "Bad request", origins, port, { "Content-Type": "text/plain; charset=utf-8" });
      return;
    }
    const filePath = resolveStaticPath(root, pathname);
    if (!filePath) {
      send(req, res, 404, "Not found", origins, port, { "Content-Type": "text/plain; charset=utf-8" });
      return;
    }
    fs.realpath(filePath, (realErr, realPath) => {
      if (realErr || !realPath.startsWith(`${root}${path.sep}`)) {
        send(req, res, 404, "Not found", origins, port, { "Content-Type": "text/plain; charset=utf-8" });
        return;
      }
      fs.stat(realPath, (statErr, stat) => {
        if (statErr || !stat.isFile()) {
          send(req, res, 404, "Not found", origins, port, { "Content-Type": "text/plain; charset=utf-8" });
          return;
        }
        const contentType = mimeTypes[path.extname(realPath)] || "application/octet-stream";
        res.writeHead(200, responseHeaders(req, origins, port, { "Content-Type": contentType }));
        if (req.method === "HEAD") res.end();
        else fs.createReadStream(realPath).pipe(res);
      });
    });
  }

  return async function requestHandler(req, res) {
    let url;
    try {
      url = new URL(req.url || "/", `https://localhost:${port}`);
    } catch {
      send(req, res, 400, "Bad request", origins, port, { "Content-Type": "text/plain; charset=utf-8" });
      return;
    }
    if (req.method === "OPTIONS") {
      if (!isRequestOriginAllowed(req, origins, port)) {
        send(req, res, 403, "Origin is not allowed", origins, port);
        return;
      }
      send(req, res, 204, "", origins, port, {
        "Access-Control-Allow-Headers": "Content-Type, Authorization, X-GMath-AI-Endpoint, X-GMath-Proxy-Token",
        "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      });
      return;
    }
    if (req.method === "GET" && url.pathname === "/api/ai/token") {
      if (!isRequestOriginAllowed(req, origins, port)) {
        sendJson(req, res, 403, { error: { message: "Origin is not allowed" } }, origins, port);
        return;
      }
      sendJson(req, res, 200, { token: proxyToken }, origins, port);
      return;
    }
    if (req.method === "POST" && url.pathname === "/api/ai/chat/completions") {
      await proxyAi(req, res);
      return;
    }
    if (req.method === "GET" || req.method === "HEAD") {
      serveStatic(req, res);
      return;
    }
    send(req, res, 405, "Method not allowed", origins, port, { "Content-Type": "text/plain; charset=utf-8" });
  };
}

export function createServer(options = {}) {
  const port = Number(options.port || process.env.PORT || 3000);
  const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
  const certPath = process.env.SSL_CRT_FILE || path.join(certDir, "localhost.crt");
  const keyPath = process.env.SSL_KEY_FILE || path.join(certDir, "localhost.key");
  const handler = createRequestHandler({ ...options, port });
  if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
    return {
      server: https.createServer({ cert: fs.readFileSync(certPath), key: fs.readFileSync(keyPath) }, handler),
      protocol: "https",
      port,
    };
  }
  console.warn("HTTPS certificate not found. Run `npm run dev-certs` for Word sideloading.");
  return { server: http.createServer(handler), protocol: "http", port };
}

const isMain = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isMain) {
  const { server, protocol, port } = createServer();
  server.listen(port, "localhost", () => {
    console.log(`GMath dev server running at ${protocol}://localhost:${port}/src/taskpane.html`);
    console.log("AI proxy is origin-restricted and available at /api/ai/chat/completions");
  });
}
