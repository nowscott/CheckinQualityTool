import assert from "node:assert/strict";
import test from "node:test";

const { excelDate } = await import("../worker/utils.js");

test("Excel local-midnight Date keeps its local calendar day", () => {
  const lessonDate = new Date(2026, 6, 13, 0, 0, 0);
  assert.equal(excelDate(lessonDate), "2026-07-13");
});
