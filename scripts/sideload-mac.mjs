import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");
const source = path.join(root, "manifest.xml");
const targetDir = path.join(
  os.homedir(),
  "Library/Containers/com.microsoft.Word/Data/Documents/wef",
);
const target = path.join(targetDir, "manifest.xml");
const port = process.env.PORT || "3000";
const localTaskpaneUrl = `https://localhost:${port}/src/taskpane.html`;
const mode = process.env.GMATH_TASKPANE || "hosted";

let manifest = fs.readFileSync(source, "utf8");
if (mode === "local") {
  manifest = manifest.replace(
    /https:\/\/ge-shun\.github\.io\/GMath\/src\/taskpane\.html/g,
    localTaskpaneUrl,
  );
}

fs.mkdirSync(targetDir, { recursive: true });
fs.writeFileSync(target, manifest);

console.log(`已复制 manifest 到 Word：${target}`);
if (mode === "local") {
  console.log(`任务窗格地址：${localTaskpaneUrl}`);
  console.log("本地开发模式需要 localhost 服务运行；请在另一个终端保持 npm run serve 运行。");
  console.log("然后完全退出并重新打开 Word。");
} else {
  console.log("任务窗格地址：https://ge-shun.github.io/GMath/src/taskpane.html");
  console.log("主界面不依赖本地服务；图片转公式会先直连接口，遇到 CORS 时可临时运行 npm run serve 使用本地代理。");
}
