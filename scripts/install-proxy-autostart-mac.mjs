import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");
const label = "com.gmath.local-proxy";
const launchAgentsDir = path.join(os.homedir(), "Library/LaunchAgents");
const logsDir = path.join(os.homedir(), "Library/Logs/GMath");
const plistPath = path.join(launchAgentsDir, `${label}.plist`);
const nodePath = process.execPath;
const serveScript = path.join(root, "scripts/serve.mjs");
const uid = process.getuid?.();
const guiTarget = typeof uid === "number" ? `gui/${uid}` : null;

function xmlEscape(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function runLaunchctl(args) {
  const result = spawnSync("launchctl", args, { encoding: "utf8" });
  return {
    ok: result.status === 0,
    stderr: result.stderr.trim(),
    stdout: result.stdout.trim(),
  };
}

fs.mkdirSync(launchAgentsDir, { recursive: true });
fs.mkdirSync(logsDir, { recursive: true });

const certPath = path.join(os.homedir(), ".office-addin-dev-certs/localhost.crt");
const keyPath = path.join(os.homedir(), ".office-addin-dev-certs/localhost.key");
if (!fs.existsSync(certPath) || !fs.existsSync(keyPath)) {
  console.warn("未找到 localhost HTTPS 证书。请先运行 `npm run dev-certs`，否则 Word/浏览器可能不信任本地代理。");
}

const plist = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>${xmlEscape(label)}</string>
  <key>ProgramArguments</key>
  <array>
    <string>${xmlEscape(nodePath)}</string>
    <string>${xmlEscape(serveScript)}</string>
  </array>
  <key>WorkingDirectory</key>
  <string>${xmlEscape(root)}</string>
  <key>EnvironmentVariables</key>
  <dict>
    <key>HOME</key>
    <string>${xmlEscape(os.homedir())}</string>
    <key>PORT</key>
    <string>3000</string>
  </dict>
  <key>RunAtLoad</key>
  <true/>
  <key>KeepAlive</key>
  <true/>
  <key>StandardOutPath</key>
  <string>${xmlEscape(path.join(logsDir, "local-proxy.out.log"))}</string>
  <key>StandardErrorPath</key>
  <string>${xmlEscape(path.join(logsDir, "local-proxy.err.log"))}</string>
</dict>
</plist>
`;

if (guiTarget) runLaunchctl(["bootout", guiTarget, plistPath]);
fs.writeFileSync(plistPath, plist);

if (!guiTarget) {
  console.log(`已写入自启动配置：${plistPath}`);
  console.log("请注销并重新登录，或手动运行 `npm run serve`。");
  process.exit(0);
}

const bootstrap = runLaunchctl(["bootstrap", guiTarget, plistPath]);
if (!bootstrap.ok) {
  console.error(`加载自启动服务失败：${bootstrap.stderr || bootstrap.stdout}`);
  process.exit(1);
}

runLaunchctl(["kickstart", "-k", `${guiTarget}/${label}`]);

console.log(`已安装并启动 GMath 本地 AI 代理：${plistPath}`);
console.log("之后登录 macOS 时会自动启动；Word 插件主界面不再依赖它，只有图片转公式会使用它。");
