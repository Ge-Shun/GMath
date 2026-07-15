import assert from "node:assert/strict";
import { Readable } from "node:stream";
import test from "node:test";
import {
  allowedOrigins,
  createRequestHandler,
  isPrivateAddress,
  resolveStaticPath,
  validateTarget,
} from "../scripts/serve.mjs";

test("blocks private and reserved network addresses", () => {
  for (const address of ["127.0.0.1", "10.0.0.1", "169.254.1.1", "172.16.0.1", "192.168.1.1", "::1", "fd00::1"]) {
    assert.equal(isPrivateAddress(address), true, address);
  }
  assert.equal(isPrivateAddress("8.8.8.8"), false);
  assert.equal(isPrivateAddress("2606:4700:4700::1111"), false);
});

test("validates HTTPS targets after DNS resolution", async () => {
  const publicLookup = async () => [{ address: "8.8.8.8", family: 4 }];
  const privateLookup = async () => [{ address: "127.0.0.1", family: 4 }];
  assert.equal(
    await validateTarget("https://api.example.test/v1", { lookup: publicLookup }),
    "https://api.example.test/v1",
  );
  await assert.rejects(() => validateTarget("http://api.example.test", { lookup: publicLookup }), /HTTPS/);
  await assert.rejects(() => validateTarget("https://api.example.test", { lookup: privateLookup }), /private or reserved/);
  await assert.rejects(
    () => validateTarget("https://api.example.test", { lookup: publicLookup, allowedHosts: new Set(["other.example"]) }),
    /GMATH_AI_HOSTS/,
  );
});

test("static path allowlist excludes repository metadata", () => {
  const root = "/tmp/gmath";
  assert.equal(resolveStaticPath(root, "/src/taskpane.html"), "/tmp/gmath/src/taskpane.html");
  assert.equal(resolveStaticPath(root, "/.git/config"), null);
  assert.equal(resolveStaticPath(root, "/package.json"), null);
  assert.equal(resolveStaticPath(root, "/src/../../.git/config"), null);
});

function callHandler(handler, { method = "GET", url = "/", headers = {}, body = "" }) {
  const req = Readable.from(body ? [Buffer.from(body)] : []);
  Object.assign(req, { method, url, headers });
  return new Promise((resolve, reject) => {
    const response = { status: 0, headers: {}, body: Buffer.alloc(0) };
    const res = {
      destroyed: false,
      writeHead(status, responseHeaders) {
        response.status = status;
        response.headers = responseHeaders;
      },
      end(value = "") {
        response.body = Buffer.isBuffer(value) ? value : Buffer.from(String(value));
        resolve(response);
      },
    };
    Promise.resolve(handler(req, res)).catch(reject);
  });
}

test("HTTP handler enforces Origin and proxy token", async () => {
  const portForPolicy = 34567;
  const token = "test-token";
  const origins = allowedOrigins(portForPolicy);
  const handler = createRequestHandler({
    root: process.cwd(),
    port: portForPolicy,
    origins,
    proxyToken: token,
  });
  const denied = await callHandler(handler, {
    url: "/api/ai/token",
    headers: { origin: "https://evil.example", host: `localhost:${portForPolicy}` },
  });
  assert.equal(denied.status, 403);

  const tokenResponse = await callHandler(handler, {
    url: "/api/ai/token",
    headers: { origin: "https://ge-shun.github.io", host: `localhost:${portForPolicy}` },
  });
  assert.equal(tokenResponse.status, 200);
  assert.deepEqual(JSON.parse(tokenResponse.body), { token });

  const noToken = await callHandler(handler, {
    method: "POST",
    url: "/api/ai/chat/completions",
    headers: { origin: "https://ge-shun.github.io", host: `localhost:${portForPolicy}`, "content-type": "application/json" },
    body: "{}",
  });
  assert.equal(noToken.status, 403);

  const privateTarget = await callHandler(handler, {
    method: "POST",
    url: "/api/ai/chat/completions",
    headers: {
      origin: "https://ge-shun.github.io",
      host: `localhost:${portForPolicy}`,
      "content-type": "application/json",
      "x-gmath-proxy-token": token,
      "x-gmath-ai-endpoint": "https://127.0.0.1/private",
    },
    body: "{}",
  });
  assert.equal(privateTarget.status, 400);

  const metadata = await callHandler(handler, {
    url: "/.git/config",
    headers: { origin: "https://ge-shun.github.io", host: `localhost:${portForPolicy}` },
  });
  assert.equal(metadata.status, 404);
});
