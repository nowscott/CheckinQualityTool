import assert from "node:assert/strict";
import test from "node:test";
import { parseInspectionCoverage, inspectionDate, shanghaiMonth } from "../server/inspectionCoverage.js";
const header = "教师姓名,学员姓名,课次日期,课次时间,入选原因,抽检日期\n";
test("完整粘贴可带重复表头，按教师去重，按抽检日期而非课次日期跨月筛选", () => {
  const result = parseInspectionCoverage(header + "甲,学员1,2026/9/30,10:00-12:00,教师覆盖,2026-10-21\n甲,学员2,2026/10/1,10:00-12:00,,2026/10/21\n" + header + "乙,学员3,2026/10/5,10:00-12:00,,2026-09-30", "2026-10");
  assert.equal(result.rowCount, 2);
  assert.equal(result.teachers.length, 1);
  assert.equal(result.month, "2026-10");
  assert.deepEqual(Object.keys(result.teachers[0]), ["teacherEmail", "teacherName"]);
  assert.equal(parseInspectionCoverage(header, "2026-10").teachers.length, 0);
});
test("缺少日期和无效日期暂不计入覆盖，并返回补填提示", () => {
  const result = parseInspectionCoverage(header + "甲,学员1,2026/10/1,10:00-12:00,,\n乙,学员2,2026/10/1,10:00-12:00,,2026-02-30\n丙,学员3,2026/10/1,10:00-12:00,,2026-10-21", "2026-10");
  assert.equal(result.rowCount, 1);
  assert.equal(result.undatedRowCount, 1);
  assert.equal(result.invalidDateRowCount, 1);
  assert.deepEqual(result.invalidDateRows, [3]);
  assert.equal(result.teachers[0].teacherName, "丙");
});
test("兼容教师手填、表格自动日期序号与日期时间，拒绝无效日期", () => {
  for (const value of ["2026-10-21", "2026/10/21", "2026年10月21日", "46316", "2026-10-21 10:00:00"]) assert.equal(inspectionDate(value), "2026-10-21");
  assert.equal(inspectionDate("2026-02-30"), "");
  assert.equal(inspectionDate("昨天"), "");
  assert.equal(inspectionDate("2024-02-29"), "2024-02-29");
});
test("拒绝缺表头、残缺粘贴和错误月份", () => {
  assert.throws(() => parseInspectionCoverage("", "2026-10"), /抽检日期/);
  assert.throws(() => parseInspectionCoverage(header + "甲,,2026/10/1,10:00,,2026-10-21", "2026-10"), /完整粘贴/);
  assert.throws(() => parseInspectionCoverage(header, "2026-13"), /月份格式/);
});
test("月份以北京时间为准", () => {
  assert.equal(shanghaiMonth(new Date("2026-09-30T16:00:00Z")), "2026-10");
});

 test("腾讯文档补齐的空白范围不计入上限，满范围实际数据仍阻止静默截断", () => {
  assert.equal(parseInspectionCoverage(header + ",,,,,\n".repeat(9999), "2026-10").teachers.length, 0);
  assert.throws(() => parseInspectionCoverage(header + ",,,,,\n".repeat(9998) + "甲,学员,2026/10/1,10:00,,2026-10-21\n", "2026-10"), /超过读取范围/);
});
