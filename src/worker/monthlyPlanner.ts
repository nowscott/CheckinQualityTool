import type {
  MonthlyInspectionData,
  MonthlyInspectionPlan,
  MonthlyTeacherPlan,
  TeachingServiceSnapshot,
  TeachingServiceTeacher,
} from "./inspectionTypes";

function scoreValue(teacher: TeachingServiceTeacher) {
  return teacher.scoreStatus === "matched" && teacher.score != null ? teacher.score : Number.POSITIVE_INFINITY;
}

export function monthWeekCount(month: string) {
  const [year, monthNumber] = month.split("-").map(Number);
  const days = new Date(Date.UTC(year, monthNumber, 0)).getUTCDate();
  let mondays = 0;
  for (let day = 1; day <= days; day += 1) {
    if (new Date(Date.UTC(year, monthNumber - 1, day)).getUTCDay() === 1) mondays += 1;
  }
  return Math.max(1, mondays);
}

function recommendedCount(teacher: TeachingServiceTeacher, month: string) {
  if (teacher.scoreStatus !== "matched" || teacher.score == null) return 1;
  if (teacher.score >= 95) return 1;
  if (teacher.score >= 80) return Math.ceil(monthWeekCount(month) / 2);
  return monthWeekCount(month);
}

function priority(teacher: TeachingServiceTeacher, items: MonthlyInspectionData["items"]) {
  const teacherItems = items.filter((item) => item.teacherEmail === teacher.teacherEmail);
  const unsubmitted = teacherItems.filter((item) => item.submittedValue === "否").length;
  const lastWeek = teacherItems.map((item) => item.businessWeekStart).sort().at(-1) || "9999-99-99";
  return { score: scoreValue(teacher), unsubmitted, lastWeek };
}

function comparePriority(left: TeachingServiceTeacher, right: TeachingServiceTeacher, items: MonthlyInspectionData["items"]) {
  const a = priority(left, items);
  const b = priority(right, items);
  return b.unsubmitted - a.unsubmitted || a.score - b.score || a.lastWeek.localeCompare(b.lastWeek) || left.teacherEmail.localeCompare(right.teacherEmail);
}

function reason(teacher: TeachingServiceTeacher, assigned: number, actual: number, weekCount: number, unsubmittedCount: number, recommended: number) {
  if (assigned > recommended && unsubmittedCount > 0) return `普检未发送优先加抽 ${unsubmittedCount} 条`;
  if (assigned > recommended && teacher.scoreStatus === "matched" && (teacher.score ?? 100) < 80) return "季度低分重点，最低频次完成后继续加抽";
  if (teacher.scoreStatus !== "matched" || teacher.score == null) return actual >= assigned ? "季度无分数，已完成保底抽检" : "季度无分数，保底抽检待补足";
  if (teacher.score >= 95) return actual >= assigned ? "季度高分，完成月度保底抽检" : "季度高分，月度保底抽检待完成";
  if (teacher.score >= 80) return actual >= assigned ? `季度中高分，完成每两周${Math.ceil(weekCount / 2)}次抽检` : `季度中高分，应每两周抽检，当前缺${assigned - actual}次`;
  return actual >= assigned ? `季度低分重点，已完成每周${weekCount}次抽检` : `季度低分重点，应每周抽检，当前缺${assigned - actual}次`;
}

export function buildMonthlyInspectionPlan(
  snapshot: TeachingServiceSnapshot,
  data: MonthlyInspectionData,
  availableSlots: number,
): MonthlyInspectionPlan {
  const teachers = [...snapshot.teachers];
  const slots = Math.max(0, Math.floor(Number(availableSlots) || 0));
  const weekCount = monthWeekCount(data.month);
  const ordered = [...teachers].sort((left, right) => comparePriority(left, right, data.items));
  const assigned = new Map(teachers.map((teacher) => [teacher.teacherEmail, 0]));
  let remaining = slots;

  for (const teacher of ordered) {
    if (!remaining) break;
    assigned.set(teacher.teacherEmail, 1);
    remaining -= 1;
  }

  const mandatoryOrder = [...ordered].sort((left, right) => {
    const leftNeed = recommendedCount(left, data.month) - (assigned.get(left.teacherEmail) || 0);
    const rightNeed = recommendedCount(right, data.month) - (assigned.get(right.teacherEmail) || 0);
    return rightNeed - leftNeed || comparePriority(left, right, data.items);
  });
  while (remaining > 0 && mandatoryOrder.length) {
    let changed = false;
    for (const teacher of mandatoryOrder) {
      if (!remaining) break;
      const current = assigned.get(teacher.teacherEmail) || 0;
      const limit = recommendedCount(teacher, data.month);
      if (current < limit) {
        assigned.set(teacher.teacherEmail, current + 1);
        remaining -= 1;
        changed = true;
      }
    }
    if (!changed) break;
  }

  const riskOrder = [...ordered].sort((left, right) => {
    const leftRisk = data.items.filter((item) => item.teacherEmail === left.teacherEmail && item.submittedValue === "否").length;
    const rightRisk = data.items.filter((item) => item.teacherEmail === right.teacherEmail && item.submittedValue === "否").length;
    return rightRisk - leftRisk || comparePriority(left, right, data.items);
  });
  const lowScoreOrder = [...ordered].filter((teacher) => teacher.scoreStatus === "matched" && (teacher.score ?? 100) < 80);
  for (const extraOrder of [riskOrder, lowScoreOrder]) {
    const eligible = extraOrder.filter((teacher) => extraOrder !== riskOrder || data.items.some((item) => item.teacherEmail === teacher.teacherEmail && item.submittedValue === "否"));
    while (remaining > 0 && eligible.length) {
      for (const teacher of eligible) {
        if (!remaining) break;
        assigned.set(teacher.teacherEmail, (assigned.get(teacher.teacherEmail) || 0) + 1);
        remaining -= 1;
      }
    }
    if (!remaining) break;
  }

  const itemsByTeacher = new Map<string, MonthlyInspectionData["items"]>();
  for (const item of data.items) {
    const items = itemsByTeacher.get(item.teacherEmail) || [];
    items.push(item);
    itemsByTeacher.set(item.teacherEmail, items);
  }

  const rows: MonthlyTeacherPlan[] = teachers.map((teacher) => {
    const items = (itemsByTeacher.get(teacher.teacherEmail) || []).sort((left, right) => left.businessWeekStart.localeCompare(right.businessWeekStart) || left.position - right.position);
    const assignedCount = assigned.get(teacher.teacherEmail) || 0;
    const actualCount = items.length;
    const unsubmittedCount = items.filter((item) => item.submittedValue === "否").length;
    const recommended = recommendedCount(teacher, data.month);
    return {
      ...teacher,
      recommendedCount: recommended,
      assignedCount,
      actualCount,
      unsubmittedCount,
      remainingCount: Math.max(assignedCount - actualCount, 0),
      lastInspectionWeek: items.at(-1)?.businessWeekStart || "",
      focusReason: reason(teacher, assignedCount, actualCount, weekCount, unsubmittedCount, recommended),
      items,
    };
  }).sort((left, right) => right.assignedCount - left.assignedCount || comparePriority(left, right, data.items) || left.teacherName.localeCompare(right.teacherName));

  return {
    month: data.month,
    availableSlots: slots,
    candidateTeachers: teachers.length,
    assignedSlots: rows.reduce((total, row) => total + row.assignedCount, 0),
    coverageShortfall: Math.max(teachers.length - slots, 0),
    recommendedSlots: rows.reduce((total, row) => total + row.recommendedCount, 0),
    frequencyShortfall: rows.reduce((total, row) => total + Math.max(row.recommendedCount - row.assignedCount, 0), 0),
    extraSlots: Math.max(slots - teachers.length, 0),
    lowScoreTeachers: teachers.filter((teacher) => teacher.scoreStatus === "matched" && (teacher.score ?? 100) < 80).length,
    missingScoreTeachers: teachers.filter((teacher) => teacher.scoreStatus !== "matched").length,
    teachers: rows,
  };
}
