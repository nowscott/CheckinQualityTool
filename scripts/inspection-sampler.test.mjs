import assert from "node:assert/strict";
import test from "node:test";

const { buildInspectionSelection, historyItems } = await import("../worker/inspectionSampler.js");
const { hasExcludedInspectionRole } = await import("../worker/inspectionRoleRules.js");
const { displayTeacherName } = await import("../lib/teacherDisplay.js");

const roster = {
  emails: new Set(["a@xdf.cn", "b@xdf.cn", "c@xdf.cn"]),
  roleExcludedEmails: new Set(),
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

test("普通抽检不超过上限，未生成报告课程在上限外全部加抽", () => {
  const result = build(3);
  assert.equal(result.stats.normalSelectedRows, 2);
  assert.equal(result.stats.extraSelectedRows, 2);
  assert.equal(result.selectedRows.length, 4);
  assert.equal(new Set(result.selectedRows.map((item) => item.teacherEmail)).size, 3);
  assert.equal(result.stats.unsubmittedRows, 2);
  assert.equal(result.stats.excludedNotInRosterRows, 1);
  assert.equal(result.riskRows.length, 2);
  assert.ok(result.riskRows.every((item) => item.inspected));
});

test("普通抽检数小于教师数时先覆盖有普通课程的教师，风险课程仍全部加抽", () => {
  const result = build(2);
  assert.equal(result.stats.normalSelectedRows, 2);
  assert.equal(result.selectedRows.length, 4);
  assert.equal(new Set(result.selectedRows.map((item) => item.teacherEmail)).size, 3);
  assert.equal(result.selectedRows.filter((item) => !item.unsubmitted).length, 2);
  assert.equal(result.riskRows.length, 2);
});

test("普通抽检上限为零时仍抽取全部未生成报告课程", () => {
  const result = build(0);
  assert.equal(result.stats.normalSelectedRows, 0);
  assert.equal(result.stats.extraSelectedRows, 2);
  assert.equal(result.selectedRows.length, 2);
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

test("只排除经理岗位；主管岗位继续参与抽检", () => {
  assert.equal(hasExcludedInspectionRole("经理"), true);
  assert.equal(hasExcludedInspectionRole("课程经理"), true);
  assert.equal(hasExcludedInspectionRole("主管"), false);
  assert.equal(hasExcludedInspectionRole("助理主管"), false);
  assert.equal(hasExcludedInspectionRole("师训主管"), false);
});

test("经理邮箱名单中的教师从普通抽检和容量外加抽中全部排除", () => {
  const managerRows = [...rows, row(6, "b@xdf.cn", "否")];
  const result = buildInspectionSelection(
    managerRows,
    { ...roster, roleExcludedEmails: new Set(["b@xdf.cn"]) },
    {
      sampleCount: 20,
      attempt: 1,
      sourceSha256: "a".repeat(64),
      rosterSha256: "b".repeat(64),
      sourceName: "课程反馈.xlsx",
      sourceColumns: ["老师姓名", "老师邮箱", "课次ID"],
    },
  );
  assert.equal(result.stats.excludedManagementRows, 2);
  assert.equal(result.stats.excludedManagementTeachers, 1);
  assert.equal(result.stats.eligibleRows, 3);
  assert.equal(result.stats.eligibleTeachers, 2);
  assert.equal(result.stats.excludedRows, 3);
  assert.equal(result.stats.unsubmittedRows, 2);
  assert.equal(result.stats.unsubmittedSelectedRows, 2);
  assert.ok(result.selectedRows.every((item) => item.teacherEmail !== "b@xdf.cn"));
  assert.ok(result.allEligibleRows.every((item) => item.teacherEmail !== "b@xdf.cn"));
  assert.ok(result.riskRows.every((item) => item.teacherEmail !== "b@xdf.cn"));
  assert.ok(result.riskRows.every((item) => item.inspected));
});

test("教师展示姓名按邮箱末尾数字统一，历史记录复用同一规则", () => {
  assert.equal(displayTeacherName("吴君怡", "wujunyi7@xdf.cn"), "吴君怡7");
  assert.equal(displayTeacherName("吴君怡3", "wujunyi7@xdf.cn"), "吴君怡7");
  assert.equal(displayTeacherName("吴君怡", "wujunyi@xdf.cn"), "吴君怡");
  const items = historyItems([{
    selectionOrder: 1,
    teacherName: "吴君怡3",
    teacherEmail: "wujunyi7@xdf.cn",
    studentName: "学员",
    studentId: "S1",
    courseId: "C1",
    lessonStart: "2026-09-08 10:00:00",
    lessonEnd: "2026-09-08 12:00:00",
    submittedValue: "是",
    productGroup: "",
    campus: "",
    projectGroup: "",
    unsubmitted: false,
    selectionKey: "k",
    sourceRowNumber: 1,
    source: {},
    selectionReason: "测试",
  }]);
  assert.equal(items[0].teacherName, "吴君怡7");
});
