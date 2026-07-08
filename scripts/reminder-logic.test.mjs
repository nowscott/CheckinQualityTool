import assert from "node:assert/strict";
import test from "node:test";

globalThis.XLSX = {
  utils: {
    sheet_to_json(sheet) {
      return sheet.__rows;
    },
  },
};

const { buildReminderTargets } = await import("../worker/reminderListParser.js");
const { matchReminderData } = await import("../worker/reminderMatching.js");
const { reminderProjectGroup } = await import("../worker/reminderProjectGroup.js");

function workbook(rows) {
  return {
    SheetNames: ["Sheet1"],
    Sheets: {
      Sheet1: { __rows: rows },
    },
  };
}

function chat(overrides) {
  return {
    有效教师邮箱: overrides.email || "teacher@xdf.cn",
    邮箱来源: "群聊发送人邮箱",
    发送人名称: overrides.sender || "老师",
    "群名/好友昵称": overrides.group || "",
    聊天时间: overrides.time || "2026-07-07 09:00:00",
    聊天内容: overrides.content || "",
    源聊天行号: overrides.row || 2,
  };
}

test("分母只按整行完全一致去重，同名同教师不同课时不误去重", () => {
  const list = buildReminderTargets(workbook([
    ["授课教师", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    ["张老师", "益智组", "王组长", "李主管", "陈一", "12"],
    ["张老师", "益智组", "王组长", "李主管", "陈一", "12"],
    ["张老师", "益智组", "王组长", "李主管", "陈一", "24"],
  ]));
  assert.equal(list.targets.length, 2);
  assert.equal(list.counts.整行重复行数, 1);
  assert.equal(list.targets[0].去重合并行号, "3");
  const result = matchReminderData(list, []);
  assert.equal(result.exceptionRows.length, 0);
});

test("群聊名称同时命中学员和教师时自动判定已发送", () => {
  const list = buildReminderTargets(workbook([
    ["授课教师", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    ["张老师", "益智组", "王组长", "李主管", "陈一", "12"],
  ]));
  const result = matchReminderData(list, [chat({ group: "张老师 陈一 新东方学习群" })]);
  assert.equal(result.studentRows[0].是否发送, "是");
  assert.equal(result.counts.群名教师学员命中, 1);
});

test("群聊名称命中唯一学员时自动判定已发送", () => {
  const list = buildReminderTargets(workbook([
    ["授课教师", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    ["张老师", "益智组", "王组长", "李主管", "陈一", "12"],
  ]));
  const result = matchReminderData(list, [chat({ group: "陈一新东方学习群" })]);
  assert.equal(result.studentRows[0].是否发送, "是");
  assert.equal(result.counts.群名唯一学员命中, 1);
});

test("聊天内容命中唯一学员时自动判定已发送", () => {
  const list = buildReminderTargets(workbook([
    ["授课教师", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    ["张老师", "益智组", "王组长", "李主管", "陈一", "12"],
  ]));
  const result = matchReminderData(list, [chat({ content: "陈一家长您好，暑假课马上开始。" })]);
  assert.equal(result.studentRows[0].是否发送, "是");
  assert.equal(result.counts.内容唯一学员命中, 1);
});

test("同名学员仅命中学生姓名时不进入异常，保持未发送", () => {
  const list = buildReminderTargets(workbook([
    ["授课教师", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    ["张老师", "益智组", "王组长", "李主管", "陈一", "12"],
    ["李老师", "益智组", "王组长", "赵主管", "陈一", "24"],
  ]));
  const result = matchReminderData(list, [chat({ group: "陈一新东方学习群" })]);
  assert.equal(result.studentRows[0].是否发送, "否");
  assert.equal(result.studentRows[1].是否发送, "否");
  assert.equal(result.counts.无法唯一匹配, 0);
  assert.equal(result.exceptionRows.length, 0);
});

test("同名学员可按聊天发送人与授课教师自动定位到老师-学生记录", () => {
  const list = buildReminderTargets(workbook([
    ["授课教师", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    ["张老师", "益智组", "王组长", "李主管", "陈一", "12"],
    ["李老师", "益智组", "王组长", "赵主管", "陈一", "24"],
  ]));
  const result = matchReminderData(list, [chat({ sender: "李老师", group: "陈一新东方学习群" })]);
  assert.equal(result.studentRows[0].是否发送, "否");
  assert.equal(result.studentRows[1].是否发送, "是");
  assert.equal(result.studentRows[1].匹配方式, "聊天发送人与授课教师一致，且群聊名称或聊天内容包含学员姓名");
  assert.equal(result.counts.发送人教师学员命中, 1);
  assert.equal(result.counts.无法唯一匹配, 0);
});

test("同名学员可按聊天内容中的授课教师和学员姓名自动定位", () => {
  const list = buildReminderTargets(workbook([
    ["授课教师", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    ["张老师", "益智组", "王组长", "李主管", "陈一", "12"],
    ["李老师", "益智组", "王组长", "赵主管", "陈一", "24"],
  ]));
  const result = matchReminderData(list, [chat({ content: "李老师 陈一家长您好，暑假课马上开始。" })]);
  assert.equal(result.studentRows[0].是否发送, "否");
  assert.equal(result.studentRows[1].是否发送, "是");
  assert.equal(result.studentRows[1].匹配方式, "群聊名称或聊天内容同时包含学员姓名和授课教师");
  assert.equal(result.counts.群名教师学员命中, 1);
  assert.equal(result.exceptionRows.length, 0);
});

test("多条质检记录命中同一分母记录时正常判定已发送且不进异常", () => {
  const list = buildReminderTargets(workbook([
    ["授课教师", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    ["张老师", "益智组", "王组长", "李主管", "陈一", "12"],
  ]));
  const result = matchReminderData(list, [
    chat({ group: "张老师 陈一 新东方学习群", time: "2026-07-07 09:00:00", row: 2 }),
    chat({ group: "张老师 陈一 暑假班学习群", time: "2026-07-07 09:30:00", row: 3 }),
  ]);
  assert.equal(result.studentRows[0].是否发送, "是");
  assert.equal(result.studentRows[0].匹配状态, "已发送");
  assert.equal(result.studentRows[0].异常原因, "");
  assert.equal(result.studentRows[0].匹配消息数, 2);
  assert.equal(result.counts.多条质检命中, 1);
  assert.equal(result.exceptionRows.length, 0);
});

test("聊天发送人姓名末尾数字不影响教师匹配", () => {
  const list = buildReminderTargets(workbook([
    ["授课教师", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    ["黄雪慧", "益智组", "王组长", "李主管", "陈一", "12"],
    ["张老师", "益智组", "王组长", "赵主管", "陈一", "24"],
  ]));
  const result = matchReminderData(list, [chat({ sender: "黄雪慧1", group: "陈一新东方学习群" })]);
  assert.equal(result.studentRows[0].是否发送, "是");
  assert.equal(result.studentRows[1].是否发送, "否");
  assert.equal(result.counts.发送人教师学员命中, 1);
});

test("助理主管维度按应发送和已发送静态汇总", () => {
  const list = buildReminderTargets(workbook([
    ["授课教师", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    ["张老师", "益智组", "王组长", "李主管", "陈一", "12"],
    ["李老师", "益智组", "王组长", "李主管", "王二", "24"],
  ]));
  const result = matchReminderData(list, [chat({ group: "张老师 陈一 新东方学习群" })]);
  assert.equal(result.assistantRows.length, 1);
  assert.equal(result.assistantRows[0].应发送数, 2);
  assert.equal(result.assistantRows[0].已发送数, 1);
  assert.equal(result.assistantRows[0].发送率, "50.0%");
  assert.equal(result.assistantRows[0].是否达标, "否");
});

test("开课提醒项目组归类将博文和实验字母组归入文理综", () => {
  assert.equal(reminderProjectGroup("博文G"), "文理综项目");
  assert.equal(reminderProjectGroup("博文Ｚ"), "文理综项目");
  assert.equal(reminderProjectGroup("实验P"), "文理综项目");
  assert.equal(reminderProjectGroup("实验Ｂ"), "文理综项目");
  assert.equal(reminderProjectGroup("文综"), "文理综项目");
  assert.equal(reminderProjectGroup("初中博文"), "博文项目");
  assert.equal(reminderProjectGroup("高中博文"), "博文项目");
  assert.equal(reminderProjectGroup("初中益智"), "益智项目");
});
