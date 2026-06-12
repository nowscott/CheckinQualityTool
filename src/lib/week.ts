import type { WeekLabel } from "../types/worker";

function weekOfMonthFromDate(date: Date) {
  const year = date.getFullYear();
  const month = date.getMonth();
  const firstDay = new Date(year, month, 1);
  const mondayOffset = (firstDay.getDay() + 6) % 7;
  return Math.min(5, Math.floor((date.getDate() + mondayOffset - 1) / 7) + 1);
}

function chineseWeek(number: number): Exclude<WeekLabel, "auto"> | "" {
  return (["", "第一周", "第二周", "第三周", "第四周", "第五周"][number] || "") as
    | Exclude<WeekLabel, "auto">
    | "";
}

export function inferWeekFromFilename(filename: string) {
  const matches = [...filename.matchAll(/(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})/g)];
  if (!matches.length) return "";
  const dates = matches.map((match) =>
    new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3])),
  );
  return chineseWeek(weekOfMonthFromDate(dates[Math.floor(dates.length / 2)]));
}
