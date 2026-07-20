import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const textExtensions = new Set([".js", ".jsx", ".mjs", ".html", ".css", ".md", ".json", ".yml", ".yaml"]);
const ignored = new Set(["node_modules", "dist", ".git"]);
const files = [];

function walk(directory) {
  for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
    if (ignored.has(entry.name)) continue;
    const full = path.join(directory, entry.name);
    if (entry.isDirectory()) walk(full);
    else if (textExtensions.has(path.extname(entry.name))) files.push(full);
  }
}

walk(root);
const contents = files.map((file) => [file, fs.readFileSync(file, "utf8")]);
const forbidden = [
  ["운영 학교코드", new RegExp(["HANYANG", "3520"].join(""), "i")],
  ["하드코딩 암호 키", /ADMIN_PASSWORD\s*[:=]/i],
  ["인라인 비밀키", /(?:api[_-]?key|secret)\s*[:=]\s*["'][^_"'][^"']{8,}/i],
];
const failures = [];
for (const [label, pattern] of forbidden) {
  for (const [file, content] of contents) {
    if (pattern.test(content)) failures.push(`${label}: ${path.relative(root, file)}`);
  }
}
const config = fs.readFileSync(path.join(root, "public", "assets", "config.js"), "utf8");
if (!config.includes("__APPS_SCRIPT_WEB_APP_URL__")) failures.push("config.js의 배포 URL 자리표시자가 없습니다.");
if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}
console.log(`정적 보안 검사 통과: ${files.length}개 공개 파일`);
