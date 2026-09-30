import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import vm from "node:vm";

globalThis.self = { postMessage() {} };
const { buildInspectionOutput } = await import("../worker/inspectionWriter.js");
const sandbox = { console };
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.resolve(process.cwd(), "public/vendor/xlsx.full.min.js"), "utf8"), sandbox);
const XLSX = sandbox.XLSX;

const basicColumns = ["教师姓名", "学员姓名", "课次日期", "课次时间"];

function workbookFrom(chunks) {
  const byteLength = chunks.reduce((total, chunk) => total + chunk.byteLength, 0);
  const bytes = new Uint8Array(byteLength);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return XLSX.read(bytes, { type: "array" });
}

const selection = {
  selectedRows: [{
    teacherName: "教师甲",
    studentName: "学员甲",
    lessonStart: "2026-09-08 10:00:00",
    lessonEnd: "2026-09-08 12:00:00",
    selectionOrder: 1,
    unsubmitted: false,
    selectionReason: "教师覆盖",
    sourceRowNumber: 8,
    courseId: "COURSE-1",
    teacherEmail: "teacher@example.com",
    source: { 报告链接H5: "https://example.com/report", 报告生成时间: "2026-09-08 10:00", 科目: "数学" },
    submittedValue: "是",
    productGroup: "高中一对一",
    projectGroup: "项目甲",
    campus: "校区甲",
  }],
  stats: {},
  sourceName: "feedback.xlsx",
  roster: { sourceName: "roster.xlsx", snapshotDate: "2026-09-22" },
  businessWeekStart: "2026-09-07",
  businessWeekEnd: "2026-09-13",
  attempt: 1,
  ruleVersion: "inspection-v10-unique-score-name-match",
  sampleLimit: 1000,
  priority: {
    mode: "coverage",
    focusTeacherCount: 0,
    matchedFocusTeacherCount: 0,
    unmatchedFocusTeacherCount: 0,
    ambiguousFocusTeacherCount: 0,
    scoreSourceTeacherCount: 1,
    matchedScoreTeacherCount: 1,
    missingScoreTeacherCount: 0,
    ambiguousScoreTeacherCount: 0,
  },
};

test("抽检名单默认保留4个基础信息列，不导出未生成报告页", () => {
  const output = buildInspectionOutput(selection, false, false);
  const workbook = workbookFrom(output.chunks);
  assert.deepEqual(Array.from(workbook.SheetNames), ["抽检名单"]);
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets["抽检名单"], { header: 1, defval: "" });
  assert.deepEqual(Array.from(rows[0]), basicColumns);
  assert.deepEqual(Array.from(rows[1]), ["教师甲", "学员甲", "2026/9/8", "10:00-12:00"]);
});

test("入选原因列由独立选项控制", () => {
  const workbook = workbookFrom(buildInspectionOutput(selection, false, true).chunks);
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets["抽检名单"], { header: 1, defval: "" });
  assert.deepEqual(Array.from(rows[0]), [...basicColumns, "入选原因"]);
  assert.deepEqual(Array.from(rows[1]), ["教师甲", "学员甲", "2026/9/8", "10:00-12:00", "教师覆盖"]);
});

test("处理说明页通过独立选项添加", () => {
  const workbook = workbookFrom(buildInspectionOutput(selection, true, false).chunks);
  assert.deepEqual(Array.from(workbook.SheetNames), ["抽检名单", "处理说明"]);
});
