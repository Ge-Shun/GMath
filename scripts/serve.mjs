import fs from "node:fs";
import http from "node:http";
import https from "node:https";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");
const port = Number(process.env.PORT || 3000);
const certDir = path.join(process.env.HOME || "", ".office-addin-dev-certs");
const certPath = process.env.SSL_CRT_FILE || path.join(certDir, "localhost.crt");
const keyPath = process.env.SSL_KEY_FILE || path.join(certDir, "localhost.key");

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

function send(res, status, body, headers = {}) {
  res.writeHead(status, {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type, Authorization, X-GMath-AI-Endpoint",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Cache-Control": "no-store",
    ...headers,
  });
  res.end(body);
}

function sendJson(res, status, payload) {
  send(res, status, JSON.stringify(payload), { "Content-Type": "application/json; charset=utf-8" });
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > 25 * 1024 * 1024) {
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

function normalizeTarget(raw) {
  if (!raw || Array.isArray(raw)) throw new Error("Missing target endpoint");
  const target = new URL(String(raw));
  if (target.protocol !== "https:" && target.protocol !== "http:") {
    throw new Error("Target endpoint must be http or https");
  }
  return target.toString();
}

async function proxyAi(req, res) {
  let body;
  let target;
  try {
    body = await readBody(req);
    target = normalizeTarget(req.headers["x-gmath-ai-endpoint"]);
  } catch (err) {
    sendJson(res, 400, { error: { message: err.message || String(err) } });
    return;
  }

  try {
    const upstream = await fetch(target, {
      method: "POST",
      headers: {
        "Content-Type": req.headers["content-type"] || "application/json",
        Authorization: req.headers.authorization || "",
      },
      body,
    });
    const upstreamBody = Buffer.from(await upstream.arrayBuffer());
    send(res, upstream.status, upstreamBody, {
      "Content-Type": upstream.headers.get("content-type") || "application/json; charset=utf-8",
    });
  } catch (err) {
    sendJson(res, 502, {
      error: {
        message: `Local proxy could not reach the AI endpoint: ${err.message || String(err)}`,
      },
    });
  }
}

function serveStatic(req, res) {
  const requestUrl = new URL(req.url || "/", `https://localhost:${port}`);
  const decodedPath = decodeURIComponent(requestUrl.pathname);
  const requestPath = decodedPath === "/" ? "/src/taskpane.html" : decodedPath;
  const filePath = path.resolve(root, "." + requestPath);
  if (!filePath.startsWith(root + path.sep) && filePath !== root) {
    send(res, 403, "Forbidden", { "Content-Type": "text/plain; charset=utf-8" });
    return;
  }

  fs.stat(filePath, (statErr, stat) => {
    if (statErr || !stat.isFile()) {
      send(res, 404, "Not found", { "Content-Type": "text/plain; charset=utf-8" });
      return;
    }
    const contentType = mimeTypes[path.extname(filePath)] || "application/octet-stream";
    res.writeHead(200, {
      "Access-Control-Allow-Origin": "*",
      "Cache-Control": "no-store",
      "Content-Type": contentType,
    });
    if (req.method === "HEAD") {
      res.end();
      return;
    }
    fs.createReadStream(filePath).pipe(res);
  });
}

async function requestHandler(req, res) {
  if (req.method === "OPTIONS") {
    send(res, 204, "");
    return;
  }
  const url = new URL(req.url || "/", `https://localhost:${port}`);
  if (req.method === "POST" && url.pathname === "/api/ai/chat/completions") {
    await proxyAi(req, res);
    return;
  }
  if (req.method === "GET" || req.method === "HEAD") {
    serveStatic(req, res);
    return;
  }
  send(res, 405, "Method not allowed", { "Content-Type": "text/plain; charset=utf-8" });
}

function createServer() {
  if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
    return https.createServer({
      cert: fs.readFileSync(certPath),
      key: fs.readFileSync(keyPath),
    }, requestHandler);
  }

  console.warn("HTTPS certificate not found. Run `npm run dev-certs` for Word sideloading.");
  return http.createServer(requestHandler);
}

createServer().listen(port, "localhost", () => {
  const protocol = fs.existsSync(certPath) && fs.existsSync(keyPath) ? "https" : "http";
  console.log(`GMath dev server running at ${protocol}://localhost:${port}/src/taskpane.html`);
  console.log("AI proxy available at /api/ai/chat/completions");
});
