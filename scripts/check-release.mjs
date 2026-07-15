import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const read = (name) => fs.readFileSync(path.join(root, name), "utf8");
const pkg = JSON.parse(read("package.json"));
const lock = JSON.parse(read("package-lock.json"));
const checks = [
  ["manifest.xml", new RegExp(`<Version>${pkg.version.replaceAll(".", "\\.")}</Version>`)],
  ["src/version.js", new RegExp(`VERSION = ["']${pkg.version.replaceAll(".", "\\.")}["']`)],
  ["src/taskpane.html", new RegExp(`taskpane\\.css\\?v=${pkg.version.replaceAll(".", "\\.")}`)],
  ["src/taskpane.html", new RegExp(`taskpane\\.js\\?v=${pkg.version.replaceAll(".", "\\.")}`)],
];

const failures = checks
  .filter(([file, pattern]) => !pattern.test(read(file)))
  .map(([file]) => file);

if (lock.version !== pkg.version || lock.packages?.[""]?.version !== pkg.version) {
  failures.push("package-lock.json");
}

const vendored = read("src/vendor/mathlive/mathlive-fonts.css").match(/^\/\*\s*([^ ]+)\s*\*\//)?.[1];
if (vendored !== pkg.dependencies.mathlive) failures.push("src/vendor/mathlive (version mismatch)");
for (const license of ["src/vendor/mathlive/LICENSE.txt", "src/vendor/fonts/OFL-1.1.txt"]) {
  if (!fs.existsSync(path.join(root, license))) failures.push(license);
}

if (failures.length) {
  console.error(`Release metadata is out of sync (${pkg.version}): ${failures.join(", ")}`);
  process.exit(1);
}
console.log(`Release metadata is synchronized at ${pkg.version}.`);
