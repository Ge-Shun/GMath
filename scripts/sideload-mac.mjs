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

let manifest = fs.readFileSync(source, "utf8");
manifest = manifest.replace(
  /https:\/\/ge-shun\.github\.io\/GMath\/src\/taskpane\.html/g,
  localTaskpaneUrl,
);

fs.mkdirSync(targetDir, { recursive: true });
fs.writeFileSync(target, manifest);

console.log(`已复制本地开发 manifest 到 Word：${target}`);
console.log(`任务窗格地址：${localTaskpaneUrl}`);
console.log("请保持 npm run serve 运行，然后完全退出并重新打开 Word。");
