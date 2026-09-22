import assert from "node:assert/strict";
import test from "node:test";

const { buildInspectionSelection } = await import("../worker/inspectionSampler.js");

const roster = {
  emails: new Set(["a@xdf.cn", "b@xdf.cn", "c@xdf.cn"]),
  sourceName: "在职教师明细20260915.xlsx",
  snapshotDate: "2026-09-15",
  rowCount: 3,
  matchedEmailRows: 3,
};

function row(index, teacherEmail, submittedValue = "是") {
  return {
    sourceRowNumber: index,
    source: { 老师姓名: teacherEmail.slice(0, 1), 老师邮箱: teacherEmail, 课次ID: `C${index}` },
    teacherName: teacherEmail.slice(0, 1),
    teacherEmail,
    studentName: `学员${index}`,
    studentId: `S${index}`,
    courseId: `C${index}`,
    lessonStart: `2026-09-${String(8 + index).padStart(2, "0")} 10:00:00`,
    lessonEnd: `2026-09-${String(8 + index).padStart(2, "0")} 12:00:00`,
    submittedValue,
    productGroup: "高中-普通一对一",
    campus: "广州学校",
    projectGroup: "测试组",
    unsubmitted: submittedValue === "否",
    selectionKey: `${teacherEmail}|C${index}|${index}`,
  };
}

const rows = [
  row(1, "a@xdf.cn", "否"),
  row(2, "a@xdf.cn", "是"),
  row(3, "b@xdf.cn", "是"),
  row(4, "c@xdf.cn", "否"),
  row(5, "no@xdf.cn", "否"),
];

function build(sampleCount, attempt = 1) {
  return buildInspectionSelection(rows, roster, {
    sampleCount,
    attempt,
    sourceSha256: "a".repeat(64),
    rosterSha256: "b".repeat(64),
    sourceName: "课程反馈.xlsx",
    sourceColumns: ["老师姓名", "老师邮箱", "课次ID"],
  });
}

test("抽检数不超过上限，优先覆盖教师并保留风险课程", () => {
  const result = build(3);
  assert.equal(result.selectedRows.length, 3);
  assert.equal(new Set(result.selectedRows.map((item) => item.teacherEmail)).size, 3);
  assert.equal(result.stats.unsubmittedRows, 2);
  assert.equal(result.stats.excludedNotInRosterRows, 1);
  assert.equal(result.riskRows.length, 2);
});

test("抽检数小于教师数时仍按风险优先选择不同教师", () => {
  const result = build(2);
  assert.equal(result.selectedRows.length, 2);
  assert.equal(new Set(result.selectedRows.map((item) => item.teacherEmail)).size, 2);
  assert.ok(result.selectedRows.every((item) => item.unsubmitted));
});

test("相同输入和尝试次数生成稳定名单", () => {
  const first = build(3);
  const second = build(3);
  assert.deepEqual(first.selectedRows.map((item) => item.courseId), second.selectedRows.map((item) => item.courseId));
  assert.equal(first.businessWeekStart, "2026-09-07");
  assert.equal(first.businessWeekEnd, "2026-09-13");
});

test("输出展示顺序将同一教师课程连续排列", () => {
  const result = build(4);
  const emails = result.selectedRows.map((item) => item.teacherEmail);
  const seen = new Set();
  let previous = "";
  for (const email of emails) {
    if (email !== previous) assert.equal(seen.has(email), false);
    seen.add(email);
    previous = email;
  }
});
