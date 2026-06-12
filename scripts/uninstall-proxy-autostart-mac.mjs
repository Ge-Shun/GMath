import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const label = "com.gmath.local-proxy";
const plistPath = path.join(os.homedir(), "Library/LaunchAgents", `${label}.plist`);
const uid = process.getuid?.();
const guiTarget = typeof uid === "number" ? `gui/${uid}` : null;

if (guiTarget) {
  spawnSync("launchctl", ["bootout", guiTarget, plistPath], { encoding: "utf8" });
}

if (fs.existsSync(plistPath)) fs.unlinkSync(plistPath);

console.log("已移除 GMath 本地 AI 代理自启动配置。");
