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
  ["하드코딩 학교코드", /(?:school[_-]?code|학교코드)\s*[:=]\s*["'][A-Z0-9_-]{6,}["']/i],
  ["하드코딩 암호 키", /ADMIN_PASSWORD\s*[:=]/i],
  ["인라인 비밀키", /(?:api[_-]?key|secret)\s*[:=]\s*["'][^_"'][^"']{8,}/i],
];
const privateSchoolCode = String(process.env.PRIVATE_SCHOOL_CODE || "").trim();
if (privateSchoolCode) {
  forbidden.push(["운영 학교코드", new RegExp(privateSchoolCode.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i")]);
}
const failures = [];
for (const [label, pattern] of forbidden) {
  for (const [file, content] of contents) {
    if (pattern.test(content)) failures.push(`${label}: ${path.relative(root, file)}`);
  }
}
const config = fs.readFileSync(path.join(root, "public", "assets", "config.js"), "utf8");
const appSource = fs.readFileSync(path.join(root, "src", "App.jsx"), "utf8");
const demoSource = fs.readFileSync(path.join(root, "src", "demo-data.js"), "utf8");
const decorativeEnglish = [
  "SCHOOL EXAM OPERATIONS",
  "ADMIN ACCESS",
  "TEACHER DESK",
  "EXAM HEADQUARTERS",
  "SEATING PLAN",
  "ADMIN SETTINGS",
  "PERSONAL EXAM SCHEDULE",
];
for (const phrase of decorativeEnglish) {
  if (appSource.includes(phrase)) failures.push(`장식용 영문 문구가 남아 있습니다: ${phrase}`);
}
if (appSource.includes('className="eyebrow"')) failures.push("장식용 영문 요소가 남아 있습니다.");
if (!demoSource.includes('school_name: "한양대학교사범대학부속고등학교"')) {
  failures.push("데모 학교명이 정식 학교명과 일치하지 않습니다.");
}
const apiUrl = config.match(/API_URL:\s*["']([^"']*)["']/)?.[1] || "";
const validPlaceholder = apiUrl === "__APPS_SCRIPT_WEB_APP_URL__";
const validDeployment = /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(apiUrl);
if (!validPlaceholder && !validDeployment) {
  failures.push("config.js의 API_URL은 배포 자리표시자 또는 Apps Script /exec 주소여야 합니다.");
}
if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}
console.log(`정적 보안 검사 통과: ${files.length}개 공개 파일`);
