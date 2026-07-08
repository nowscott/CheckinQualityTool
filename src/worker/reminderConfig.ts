export const REMINDER_PASS_RATE = 0.8;

export const REMINDER_LIST_HEADERS = {
  teacher: ["授课教师", "教师姓名", "老师姓名"],
  teachingGroup: ["教研组"],
  trainingLead: ["师训组长"],
  assistantLead: ["师训助理主管/主管", "助理主管", "师训助理主管"],
  student: ["学员姓名", "学生姓名"],
  studentType: ["新老生(季度)", "新老生（季度）", "新老生"],
  grade: ["年级"],
  counselor: ["学管", "学管姓名"],
  classHours: ["课时"],
  teacherEmail: ["教师邮箱", "老师邮箱", "邮箱"],
  campus: ["校区", "教学区"],
} as const;

export const REMINDER_CHAT_HEADERS = {
  group: ["群名/好友昵称", "群聊名称", "群名", "好友昵称"],
  content: ["聊天内容", "消息内容", "内容"],
  senderName: ["群聊发送人名称", "姓名", "发送人", "发送人名称"],
  senderEmail: ["群聊发送人邮箱", "邮箱"],
  chatTime: ["聊天时间", "发送时间", "消息时间"],
} as const;

export const REMINDER_REQUIRED_LIST_FIELDS = ["授课教师", "学员姓名"] as const;

export const REMINDER_MATCH_RULES = {
  senderTeacherAndStudent: "聊天发送人与授课教师一致，且群聊名称或聊天内容包含学员姓名",
  groupStudentAndTeacher: "群聊名称或聊天内容同时包含学员姓名和授课教师",
  uniqueStudentInGroup: "群聊名称包含学员姓名，且分母中该学员唯一",
  uniqueStudentInContent: "聊天内容包含学员姓名，且分母中该学员唯一",
} as const;
