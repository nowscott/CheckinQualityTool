import { createHash } from "node:crypto";
import { readFileSync, readdirSync, renameSync, writeFileSync } from "node:fs";
import { basename, dirname, extname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";

const projectRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const sourceDirectory = resolve(projectRoot, "data/inspection/roster-source");
const workbookLibraryPath = resolve(projectRoot, "public/vendor/xlsx.full.min.js");
const sourceArgument = process.argv[2];

function sourceFilePath() {
  if (sourceArgument) return resolve(sourceArgument);
  const candidates = readdirSync(sourceDirectory)
    .filter((name) => extname(name).toLowerCase() === ".xlsx")
    .map((name) => resolve(sourceDirectory, name));
  if (candidates.length !== 1) {
    throw new Error(`在 ${sourceDirectory} 中需要且只能保留一份 .xlsx 文件；当前找到 ${candidates.length} 份。`);
  }
  return candidates[0];
}

function loadSheetJs() {
  const sandbox = { console };
  vm.createContext(sandbox);
  vm.runInContext(readFileSync(workbookLibraryPath, "utf8"), sandbox, { timeout: 10_000 });
  if (!sandbox.XLSX?.read) throw new Error("无法加载项目自带的 SheetJS。");
  return sandbox.XLSX;
}

function cellText(value) {
  return value == null ? "" : String(value).trim();
}

function emailValue(value) {
  return cellText(value).toLowerCase().match(/[\w.+-]+@[\w.-]+\.[a-z]{2,}/u)?.[0] ?? "";
}

function firstColumn(headers, aliases) {
  for (const alias of aliases) {
    const index = headers.indexOf(alias);
    if (index >= 0) return index;
  }
  return -1;
}

function snapshotDate(fileName) {
  const match = fileName.match(/(20\d{2})[-_年]?(\d{2})[-_月]?(\d{2})/u);
  if (!match) throw new Error(`文件名缺少 YYYYMMDD 快照日期：${fileName}`);
  const [, year, month, day] = match;
  const date = new Date(`${year}-${month}-${day}T00:00:00.000Z`);
  if (Number.isNaN(date.valueOf()) || date.toISOString().slice(0, 10) !== `${year}-${month}-${day}`) {
    throw new Error(`文件名中的快照日期无效：${fileName}`);
  }
  return `${year}-${month}-${day}`;
}

function atomicJsonWrite(path, value) {
  const temporaryPath = `${path}.tmp`;
  writeFileSync(temporaryPath, `${JSON.stringify(value, null, 2)}\n`);
  renameSync(temporaryPath, path);
}

const inputPath = sourceFilePath();
const inputBytes = readFileSync(inputPath);
const inputName = basename(inputPath);
const date = snapshotDate(inputName);
const XLSX = loadSheetJs();
const workbook = XLSX.read(inputBytes, {
  type: "buffer",
  dense: true,
  cellDates: true,
  cellText: false,
});

const candidates = workbook.SheetNames.map((name) => {
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], {
    header: 1,
    range: 0,
    blankrows: false,
    defval: "",
  });
  const headers = rows[0]?.map(cellText) ?? [];
  return {
    name,
    rows,
    emailIndex: firstColumn(headers, ["邮箱", "教师邮箱", "老师邮箱"]),
    roleIndex: firstColumn(headers, ["岗位描述", "岗位短描述", "职位描述", "职位"]),
  };
}).filter((candidate) => candidate.rows.length && candidate.emailIndex >= 0)
  .sort((left, right) => right.rows.length - left.rows.length);

if (!candidates.length) throw new Error("找不到含“邮箱”“教师邮箱”或“老师邮箱”的工作表。");
const selected = candidates[0];
const emails = new Set();
const roleExcludedEmails = new Set();
let matchedEmailRows = 0;
let unmatchedNonemptyEmailRows = 0;

for (const row of selected.rows.slice(1)) {
  const rawEmail = row[selected.emailIndex];
  const email = emailValue(rawEmail);
  if (!email) {
    if (cellText(rawEmail)) unmatchedNonemptyEmailRows += 1;
    continue;
  }
  matchedEmailRows += 1;
  emails.add(email);
  const role = selected.roleIndex >= 0 ? cellText(row[selected.roleIndex]) : "";
  if (/经理/u.test(role)) roleExcludedEmails.add(email);
}

if (!emails.size) throw new Error("邮箱列没有可识别的邮箱，未更新名单。");

const roster = {
  snapshotDate: date,
  sourceFile: inputName,
  sourceSha256: createHash("sha256").update(inputBytes).digest("hex"),
  rowCount: selected.rows.length - 1,
  matchedEmailRows,
  emails: [...emails].sort(),
};
const roleExclusions = {
  snapshotDate: date,
  sourceFile: inputName,
  emails: [...roleExcludedEmails].sort(),
};

atomicJsonWrite(resolve(projectRoot, "public/data/inspection-roster.json"), roster);
atomicJsonWrite(resolve(projectRoot, "public/data/inspection-role-exclusions.json"), roleExclusions);

console.log(`来源工作表：${selected.name}`);
console.log(`快照日期：${date}`);
console.log(`明细行：${roster.rowCount}；有效邮箱行：${matchedEmailRows}；唯一邮箱：${emails.size}`);
console.log(`经理岗位邮箱：${roleExcludedEmails.size}；非空但无法解析的邮箱单元格：${unmatchedNonemptyEmailRows}`);
