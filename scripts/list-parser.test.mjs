import assert from "node:assert/strict";
import test from "node:test";

globalThis.XLSX = {
  utils: {
    sheet_to_json(sheet) {
      return sheet.rows;
    },
  },
};

const { buildTargets } = await import("../worker/listParser.js");

test("完整课次时间表头优先选择含真实老师邮箱的课堂反馈 Sheet", () => {
  const workbook = {
    SheetNames: ["错位表", "课堂反馈主表"],
    Sheets: {
      错位表: {
        rows: [
          ["老师姓名", "学员姓名", "学员号", "课次开始时", "间", "间", "老师邮箱"],
          ["错表老师", "错表学员", "GZ0", "2026-07-20", "10:20", "12:20", "1065"],
        ],
      },
      课堂反馈主表: {
        rows: [
          ["老师姓名", "学员姓名", "学员号", "课次开始时间", "课次结束时间", "老师邮箱"],
          ["正确老师", "正确学员", "GZ1", new Date(2026, 6, 20, 10, 20), new Date(2026, 6, 20, 12, 20), "teacher@xdf.cn"],
        ],
      },
    },
  };

  const result = buildTargets(workbook);

  assert.equal(result.sheetName, "课堂反馈主表");
  assert.equal(result.targets.length, 1);
  assert.equal(result.targets[0].教师邮箱, "teacher@xdf.cn");
  assert.equal(result.targets[0].上课开始, "10:20");
  assert.equal(result.targets[0].上课结束, "12:20");
});
