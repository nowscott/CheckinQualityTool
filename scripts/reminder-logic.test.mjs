import assert from "node:assert/strict";
import test from "node:test";

globalThis.XLSX = {
  utils: {
    sheet_to_json(sheet) {
      return sheet.__rows;
    },
  },
};

const exemptTeacher = decodeURIComponent("%E7%89%9B%E6%96%BD%E6%A1%A5");
const teacherSuffix = decodeURIComponent("%E8%80%81%E5%B8%88");

const { buildReminderTargets } = await import("../worker/reminderListParser.js");
const { buildReminderAppeals } = await import("../worker/reminderAppealParser.js");
const { matchReminderData } = await import("../worker/reminderMatching.js");
const { reminderProjectGroup } = await import("../worker/reminderProjectGroup.js");
const { buildWhitelist } = await import("../worker/whitelist.js");

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

test("指定教师开课提醒无需聊天命中也按已发送统计", () => {
  const list = buildReminderTargets(workbook([
    ["授课教师", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    [`${exemptTeacher}${teacherSuffix}`, "益智组", "王组长", "李主管", "陈一", "12"],
    [`${exemptTeacher}1`, "益智组", "王组长", "李主管", "王二", "24"],
    ["张老师", "益智组", "王组长", "赵主管", "赵三", "24"],
  ]));
  const result = matchReminderData(list, []);
  assert.deepEqual(result.studentRows.map((row) => row.是否发送), ["是", "是", "否"]);
  assert.equal(result.counts.应发送数, 3);
  assert.equal(result.counts.已发送数, 2);
  assert.equal(result.counts.未发送数, 1);
  assert.equal(result.teacherRows.find((row) => row.教师姓名 === `${exemptTeacher}${teacherSuffix}`).发送率, "100.0%");
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

test("开课提醒白名单别名可匹配系统登记名与实际群名不一致的学员", () => {
  const whitelist = buildWhitelist([
    "学员号,学员姓名,处理方式,匹配别名,说明",
    "GZ5443556611,何喜,别名,何梦涵,系统名与企微群名不一致",
  ].join("\n"));
  const list = buildReminderTargets(workbook([
    ["授课教师", "学员号", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    ["张老师", "GZ5443556611", "益智组", "王组长", "李主管", "何喜", "12"],
  ]), whitelist);
  const result = matchReminderData(list, [chat({ group: "张老师 何梦涵 新东方学习群" })]);
  assert.equal(result.studentRows[0].是否发送, "是");
  assert.equal(result.studentRows[0].白名单命中, "是");
  assert.equal(result.studentRows[0].白名单说明, "系统名与企微群名不一致");
  assert.equal(result.studentRows[0].命中关键词, "张老师+梦涵（白名单别名）");
  assert.equal(result.counts.白名单别名命中, 1);
});

test("开课提醒白名单保留原名先于姓名后缀清洗生效", () => {
  const whitelist = buildWhitelist([
    "学员号,学员姓名,处理方式,匹配别名,说明",
    "GZ6005060860,陈姜玉一,保留原名,,四字姓名末尾“一”为真实姓名，按学员号保留",
  ].join("\n"));
  const list = buildReminderTargets(workbook([
    ["授课教师", "学员号", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    ["张老师", "GZ6005060860", "益智组", "王组长", "李主管", "陈姜玉一", "12"],
  ]), whitelist);
  assert.equal(list.targets[0].匹配学员姓名, "陈姜玉一");
  assert.equal(list.targets[0].姓名清洗说明, "陈姜玉一（白名单保留原名）");
  const result = matchReminderData(list, [chat({ group: "张老师 陈姜玉一 新东方学习群" })]);
  assert.equal(result.studentRows[0].是否发送, "是");
});

test("申诉按教师和学生剔除分母并保留明细标注，不依赖是否通过", () => {
  const list = buildReminderTargets(workbook([
    ["授课教师", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    ["张老师", "益智组", "王组长", "李主管", "陈一", "12"],
    ["李老师", "益智组", "王组长", "赵主管", "王二", "24"],
  ]));
  const appeals = buildReminderAppeals(workbook([
    ["教师姓名", "学生姓名", "申诉原因", "申诉原因描述", "申诉是否通过"],
    ["张老师", "陈一", "已发送提醒话术", "截图已核实", "是"],
  ]));
  const result = matchReminderData(list, [], appeals);
  assert.equal(result.counts.应发送数, 1);
  assert.equal(result.counts.申诉数, 1);
  assert.equal(result.counts.未发送数, 1);
  assert.equal(result.studentRows[0].是否发送, "已申诉");
  assert.equal(result.studentRows[0].申诉情况说明, "已发送提醒话术：截图已核实");
  assert.equal(result.teacherRows.length, 2);
  const appealedTeacher = result.teacherRows.find((row) => row.教师姓名 === "张老师");
  const unsentTeacher = result.teacherRows.find((row) => row.教师姓名 === "李老师");
  assert.equal(appealedTeacher.应发送数, 0);
  assert.equal(appealedTeacher.申诉数, 1);
  assert.equal(unsentTeacher.应发送数, 1);
  assert.equal(unsentTeacher.申诉数, 0);
});

test("未标记通过的申诉同样公示并剔除分母", () => {
  const list = buildReminderTargets(workbook([
    ["授课教师", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    ["张老师", "益智组", "王组长", "李主管", "陈一", "12"],
  ]));
  const appeals = buildReminderAppeals(workbook([
    ["教师姓名", "学生姓名", "申诉原因", "申诉原因描述", "申诉是否通过"],
    ["张老师", "陈一", "老师已私发", "待复核截图", "否"],
  ]));
  const result = matchReminderData(list, [chat({ group: "张老师 陈一 新东方学习群" })], appeals);
  assert.equal(appeals.counts.申诉行数, 1);
  assert.equal(result.counts.应发送数, 0);
  assert.equal(result.counts.已发送数, 0);
  assert.equal(result.counts.申诉数, 1);
  assert.equal(result.studentRows[0].是否发送, "已申诉");
  assert.equal(result.studentRows[0].申诉情况说明, "老师已私发：待复核截图");
});

test("同一申诉记录填写多个学员时分别拆分匹配", () => {
  const list = buildReminderTargets(workbook([
    ["授课教师", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    ["赖老师", "益智组", "王组长", "李主管", "陈浩艺", "12"],
    ["赖老师", "益智组", "王组长", "李主管", "周秀雅", "12"],
    ["赖老师", "益智组", "王组长", "李主管", "周少哲", "12"],
    ["陈老师", "益智组", "王组长", "李主管", "全星熹", "12"],
    ["陈老师", "益智组", "王组长", "李主管", "王唯应", "12"],
  ]));
  const appeals = buildReminderAppeals(workbook([
    ["教师姓名", "学生姓名", "申诉原因", "申诉原因描述", "申诉是否通过"],
    ["赖老师", "陈浩艺，周秀雅，周少哲", "已发送提醒话术", "同一截图", ""],
    ["陈老师202", "全星熹 王唯应", "其他特殊情况", "未收到档案", "否"],
  ]));
  const result = matchReminderData(list, [], appeals);
  assert.equal(appeals.counts.原始申诉行数, 2);
  assert.equal(appeals.counts.申诉行数, 5);
  assert.equal(appeals.counts.申诉拆分学员数, 3);
  assert.equal(result.counts.应发送数, 0);
  assert.equal(result.counts.申诉数, 5);
  assert.deepEqual(result.studentRows.map((row) => row.是否发送), ["已申诉", "已申诉", "已申诉", "已申诉", "已申诉"]);
  assert.equal(result.studentRows[1].申诉情况说明, "已发送提醒话术：同一截图");
  assert.equal(result.studentRows[4].申诉情况说明, "其他特殊情况：未收到档案");
});

test("助理主管维度按主管自身汇总，跨教研组不重复主管分组", () => {
  const list = buildReminderTargets(workbook([
    ["授课教师", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    ["张老师", "益智组", "王组长", "李主管", "陈一", "12"],
    ["李老师", "双语组", "赵组长", "李主管", "王二", "24"],
  ]));
  const result = matchReminderData(list, [chat({ group: "张老师 陈一 新东方学习群" })]);
  assert.equal(result.assistantRows.length, 1);
  assert.equal(result.assistantRows[0].应发送数, 2);
  assert.equal(result.assistantRows[0].已发送数, 1);
  assert.equal(result.assistantRows[0].发送率, "50.0%");
  assert.equal(result.assistantRows[0].是否达标, "否");
});

test("助理主管维度使用主管本人教研组归属其管辖老师", () => {
  const list = buildReminderTargets(workbook([
    ["授课教师", "教研组", "师训组长", "师训助理主管/主管", "学员姓名", "课时"],
    ["黄主管", "初中博文", "王组长", "黄主管", "陈一", "12"],
    ["张老师", "高中博文", "王组长", "黄主管", "王二", "24"],
    ["李老师", "高中博文", "王组长", "黄主管", "赵三", "24"],
  ]));
  const result = matchReminderData(list, [chat({ sender: "黄主管", group: "陈一新东方学习群" })]);
  assert.equal(result.assistantRows.length, 1);
  assert.equal(result.assistantRows[0].助理主管, "黄主管");
  assert.equal(result.assistantRows[0].教研组, "初中博文");
  assert.equal(result.assistantRows[0].应发送数, 3);
  assert.equal(result.assistantRows[0].已发送数, 1);
});

test("开课提醒项目组归类将博文和实验字母组归入文理综", () => {
  assert.equal(reminderProjectGroup("博文G"), "文理综项目");
  assert.equal(reminderProjectGroup("博文Ｚ"), "文理综项目");
  assert.equal(reminderProjectGroup("实验P"), "文理综项目");
  assert.equal(reminderProjectGroup("实验Ｂ"), "文理综项目");
  assert.equal(reminderProjectGroup("文综"), "文理综项目");
  assert.equal(reminderProjectGroup("政史地生"), "文理综项目");
  assert.equal(reminderProjectGroup("史地生"), "文理综项目");
  assert.equal(reminderProjectGroup("初中博文"), "博文项目");
  assert.equal(reminderProjectGroup("高中博文"), "博文项目");
  assert.equal(reminderProjectGroup("初中益智"), "益智项目");
});
