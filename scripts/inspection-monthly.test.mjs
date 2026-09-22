import assert from "node:assert/strict";
import test from "node:test";
import fs from "node:fs";

const { buildMonthlyInspectionPlan } = await import("../worker/monthlyPlanner.js");
const snapshot = JSON.parse(fs.readFileSync("data/inspection/teaching-service-q1.json", "utf8"));

test("月度模拟先覆盖教师，再把额外次数给低分教师", () => {
  const plan = buildMonthlyInspectionPlan(snapshot, { month: "2026-09", items: [] }, 1000);
  assert.equal(plan.candidateTeachers, 837);
  assert.equal(plan.coverageShortfall, 0);
  assert.equal(plan.assignedSlots, 1000);
  assert.ok(plan.extraSlots > 0);
  assert.equal(plan.teachers[0].recommendedCount, 4);
  assert.ok(plan.frequencyShortfall > 0);
  assert.ok(plan.teachers[0].score < 70);
  assert.equal(plan.teachers[0].assignedCount, 2);
});

test("可用次数不足时显示覆盖缺口", () => {
  const plan = buildMonthlyInspectionPlan(snapshot, { month: "2026-09", items: [] }, 200);
  assert.equal(plan.assignedSlots, 200);
  assert.equal(plan.coverageShortfall, 637);
  assert.equal(plan.teachers.filter((teacher) => teacher.assignedCount > 0).length, 200);
});

test("季度无分数教师保底一次并单独标记", () => {
  const plan = buildMonthlyInspectionPlan(snapshot, { month: "2026-09", items: [] }, 1000);
  const missing = plan.teachers.filter((teacher) => teacher.scoreStatus !== "matched");
  assert.equal(missing.length, 9);
  assert.ok(missing.every((teacher) => teacher.recommendedCount === 1 && teacher.assignedCount >= 1));
});

test("可用次数足够时低分教师按每周频次完成", () => {
  const plan = buildMonthlyInspectionPlan(snapshot, { month: "2026-09", items: [] }, 1600);
  assert.equal(plan.frequencyShortfall, 0);
  assert.equal(plan.assignedSlots, 1600);
  assert.ok(plan.teachers.filter((teacher) => teacher.scoreStatus === "matched" && (teacher.score ?? 100) < 80).every((teacher) => teacher.assignedCount >= 4));
});
