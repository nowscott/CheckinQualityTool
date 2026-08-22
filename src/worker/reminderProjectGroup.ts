import { text } from "./utils";

function compact(value: unknown) {
  return text(value).replace(/\s+/g, "");
}

export function reminderProjectGroup(teachingGroup: unknown) {
  const value = compact(teachingGroup);
  if (!value) return "其他项目";
  if (
    /(博文|实验)[A-Za-zＡ-Ｚａ-ｚ]/u.test(value) ||
    /文综|理综|文理综|政史地生|政史地|史地生/u.test(value)
  ) {
    return "文理综项目";
  }
  if (value.includes("博文")) return "博文项目";
  if (value.includes("双语")) return "双语项目";
  if (value.includes("益智")) return "益智项目";
  return "其他项目";
}
