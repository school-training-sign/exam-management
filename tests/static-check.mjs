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
const retiredSchoolCode = ["HANYANG", "3520"].join("");
const sourcePackagingSheetId = ["1ukMUIjGhRILLj", "XwpTmxm3NecSz", "-5cD3khRnD1Xk6Nx8"].join("");
forbidden.push(
  ["폐기된 운영 학교코드", new RegExp(retiredSchoolCode, "i")],
  ["외부 원안 포장 시트 ID", new RegExp(sourcePackagingSheetId, "i")],
  ["외부 Google Sheets 주소", /https:\/\/docs\.google\.com\/spreadsheets\/d\//i],
);
const failures = [];
for (const [label, pattern] of forbidden) {
  for (const [file, content] of contents) {
    if (pattern.test(content)) failures.push(`${label}: ${path.relative(root, file)}`);
  }
}
const config = fs.readFileSync(path.join(root, "public", "assets", "config.js"), "utf8");
const appSource = fs.readFileSync(path.join(root, "src", "App.jsx"), "utf8");
const apiSource = fs.readFileSync(path.join(root, "src", "api.js"), "utf8");
const demoSource = fs.readFileSync(path.join(root, "src", "demo-data.js"), "utf8");
const indexSource = fs.readFileSync(path.join(root, "index.html"), "utf8");
const systemName = "한양대학교사범대학부속고등학교 통합 시스템";
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
for (const [label, source] of [
  ["화면", appSource],
  ["데모 설정", demoSource],
  ["공개 설정", config],
  ["브라우저 제목", indexSource],
]) {
  if (!source.includes(systemName)) failures.push(`${label}의 시스템 이름이 정식 명칭과 일치하지 않습니다.`);
}
for (const phrase of [
  "데모 모드",
  "실제 데이터가 없는 데모로 둘러보기",
  "입력 정보는 학교 소유 시트에만 저장됩니다",
  "결시 입력부터 고사본부 현황, 자리배치와 출력까지 한 곳에서 관리합니다",
  'className="mode-badge',
  'className="app-footer',
]) {
  if (appSource.includes(phrase)) failures.push(`제거 대상 화면 문구 또는 요소가 남아 있습니다: ${phrase}`);
}
for (const legacyContract of ["school_login", "school_code", "school_session"]) {
  if (apiSource.includes(legacyContract) || appSource.includes(legacyContract)) {
    failures.push(`이전 학교코드 API 계약이 남아 있습니다: ${legacyContract}`);
  }
}
for (const requiredPackagingTerm of [
  "고사 원안 포장",
  "redeem_exam_packaging_invite",
  "save_exam_packaging_items",
  "exam-management:packaging-session",
]) {
  if (!appSource.includes(requiredPackagingTerm) && !apiSource.includes(requiredPackagingTerm)) {
    failures.push(`원안 포장 필수 구현이 없습니다: ${requiredPackagingTerm}`);
  }
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
