import assert from "node:assert/strict";
import test from "node:test";

const { buildInspectionSelection, historyItems } = await import("../worker/inspectionSampler.js");
const { hasExcludedInspectionRole } = await import("../worker/inspectionRoleRules.js");
const { displayTeacherName } = await import("../lib/teacherDisplay.js");

const roster = {
  emails: new Set(["a@xdf.cn", "b@xdf.cn", "c@xdf.cn", "d@xdf.cn"]),
  roleExcludedEmails: new Set(),
  sourceName: "在职教师明细20260915.xlsx",
  snapshotDate: "2026-09-15",
  rowCount: 4,
  matchedEmailRows: 4,
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

test("全覆盖后优先用未反馈教师的额外课程加频", () => {
  const focusRows = [
    row(10, "a@xdf.cn"),
    row(11, "a@xdf.cn"),
    row(12, "b@xdf.cn"),
    row(13, "b@xdf.cn"),
    row(14, "c@xdf.cn"),
    row(15, "c@xdf.cn"),
  ];
  const result = buildInspectionSelection(focusRows, roster, {
    sampleCount: 4,
    attempt: 1,
    sourceSha256: "a".repeat(64),
    rosterSha256: "b".repeat(64),
    sourceName: "课程反馈.xlsx",
    sourceColumns: ["老师姓名", "老师邮箱", "课次ID"],
    priorityMode: "coverage",
    focusTeacherNames: ["a"],
  });
  const selectedByTeacher = new Map();
  for (const selected of result.selectedRows) {
    selectedByTeacher.set(selected.teacherEmail, (selectedByTeacher.get(selected.teacherEmail) || 0) + 1);
  }
  assert.equal(result.stats.selectedTeachers, 3);
  assert.equal(result.stats.normalSelectedRows, 4);
  assert.equal(selectedByTeacher.get("a@xdf.cn"), 2);
  assert.equal(selectedByTeacher.get("b@xdf.cn"), 1);
  assert.equal(selectedByTeacher.get("c@xdf.cn"), 1);
  assert.equal(result.stats.focusTeacherExtraRows, 1);
  assert.ok(result.selectedRows.some((selected) => selected.teacherEmail === "a@xdf.cn"
    && selected.selectionReason.includes("未反馈教师剩余名额加频")));
  assert.equal(result.priority.focusTeacherCount, 1);
  assert.equal(result.priority.matchedFocusTeacherCount, 1);
});

test("名额不足时保留未反馈重点教师，再按低分优先覆盖", () => {
  const scoreRows = [
    row(20, "a@xdf.cn"),
    row(21, "b@xdf.cn"),
    row(22, "c@xdf.cn"),
    row(23, "d@xdf.cn"),
  ];
  const teacherScoresByEmail = {
    "a@xdf.cn": 100,
    "b@xdf.cn": 90,
    "c@xdf.cn": 60,
    "d@xdf.cn": 20,
  };
  for (const priorityMode of ["coverage", "unreported"]) {
    const result = buildInspectionSelection(scoreRows, roster, {
      sampleCount: 2,
      attempt: 1,
      sourceSha256: "a".repeat(64),
      rosterSha256: "b".repeat(64),
      sourceName: "课程反馈.xlsx",
      sourceColumns: ["老师姓名", "老师邮箱", "课次ID"],
      priorityMode,
      focusTeacherNames: ["a"],
      teacherScoresByEmail,
    });
    assert.deepEqual(
      result.selectedRows.map((selected) => selected.teacherEmail).sort(),
      ["a@xdf.cn", "d@xdf.cn"],
    );
    const focusRow = result.selectedRows.find((selected) => selected.teacherEmail === "a@xdf.cn");
    const lowScoreRow = result.selectedRows.find((selected) => selected.teacherEmail === "d@xdf.cn");
    assert.ok(focusRow?.selectionReason.includes(priorityMode === "coverage" ? "未反馈名单保护" : "本月未反馈教师优先"));
    assert.ok(lowScoreRow?.selectionReason.includes("低分优先保留"));
  }
});

test("余量时全覆盖模式先给未反馈教师加频，再按低分排序；未反馈模式按低分加频", () => {
  const scoreRows = [
    row(10, "a@xdf.cn"), row(11, "a@xdf.cn"),
    row(12, "b@xdf.cn"), row(13, "b@xdf.cn"),
    row(14, "c@xdf.cn"), row(15, "c@xdf.cn"),
    row(16, "d@xdf.cn"), row(17, "d@xdf.cn"),
  ];
  const teacherScoresByEmail = {
    "a@xdf.cn": 100,
    "b@xdf.cn": 90,
    "c@xdf.cn": 60,
    "d@xdf.cn": 20,
  };
  const buildScored = (priorityMode) => buildInspectionSelection(scoreRows, roster, {
    sampleCount: 6,
    attempt: 1,
    sourceSha256: "a".repeat(64),
    rosterSha256: "b".repeat(64),
    sourceName: "课程反馈.xlsx",
    sourceColumns: ["老师姓名", "老师邮箱", "课次ID"],
    priorityMode,
    focusTeacherNames: ["a"],
    teacherScoresByEmail,
  });
  const coverage = buildScored("coverage");
  const coverageCounts = new Map();
  for (const selected of coverage.selectedRows) {
    coverageCounts.set(selected.teacherEmail, (coverageCounts.get(selected.teacherEmail) || 0) + 1);
  }
  assert.equal(coverageCounts.get("a@xdf.cn"), 2);
  assert.equal(coverageCounts.get("d@xdf.cn"), 2);
  assert.equal(coverageCounts.get("b@xdf.cn"), 1);
  assert.equal(coverageCounts.get("c@xdf.cn"), 1);
  assert.equal(coverage.stats.focusTeacherExtraRows, 1);
  assert.ok(coverage.selectedRows.some((selected) => selected.teacherEmail === "a@xdf.cn"
    && selected.selectionReason.includes("未反馈教师剩余名额加频")));

  const unreported = buildScored("unreported");
  const unreportedCounts = new Map();
  for (const selected of unreported.selectedRows) {
    unreportedCounts.set(selected.teacherEmail, (unreportedCounts.get(selected.teacherEmail) || 0) + 1);
  }
  assert.equal(unreportedCounts.get("a@xdf.cn"), 1);
  assert.equal(unreportedCounts.get("d@xdf.cn"), 2);
  assert.equal(unreportedCounts.get("c@xdf.cn"), 2);
  assert.equal(unreportedCounts.get("b@xdf.cn"), 1);
  assert.equal(unreported.stats.focusTeacherExtraRows, 0);
});

test("名额不足时把高分教师顺位后移，但保留未反馈重点教师", () => {
  const scoreRows = [
    row(18, "a@xdf.cn"),
    row(19, "b@xdf.cn"),
    row(20, "c@xdf.cn"),
    row(21, "d@xdf.cn"),
  ];
  const teacherScoresByEmail = {
    "a@xdf.cn": 100,
    "b@xdf.cn": 90,
    "c@xdf.cn": 60,
    "d@xdf.cn": 20,
  };
  for (const priorityMode of ["coverage", "unreported"]) {
    const result = buildInspectionSelection(scoreRows, roster, {
      sampleCount: 2,
      attempt: 1,
      sourceSha256: "a".repeat(64),
      rosterSha256: "b".repeat(64),
      sourceName: "课程反馈.xlsx",
      sourceColumns: ["老师姓名", "老师邮箱", "课次ID"],
      priorityMode,
      focusTeacherNames: ["a"],
      teacherScoresByEmail,
    });
    assert.deepEqual(
      result.selectedRows.map((selected) => selected.teacherEmail).sort(),
      ["a@xdf.cn", "d@xdf.cn"],
    );
    const focused = result.selectedRows.find((selected) => selected.teacherEmail === "a@xdf.cn");
    const lowScored = result.selectedRows.find((selected) => selected.teacherEmail === "d@xdf.cn");
    assert.ok(focused?.selectionReason.includes(priorityMode === "coverage" ? "未反馈名单保护" : "本月未反馈教师优先"));
    assert.ok(lowScored?.selectionReason.includes("低分优先保留"));
  }
});

test("未接入分数时保持稳定顺序", () => {
  const first = build(2);
  const second = build(2);
  assert.deepEqual(first.selectedRows.map((selected) => selected.courseId), second.selectedRows.map((selected) => selected.courseId));
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
