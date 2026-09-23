export interface InspectionRoleCorrection {
  teacherName: string;
  employeeId: string;
  teacherEmail: string;
  correctedRole: string;
  reason: string;
}

export const INSPECTION_ROLE_CORRECTIONS: InspectionRoleCorrection[] = [
  {
    teacherName: "祖万露",
    employeeId: "294086",
    teacherEmail: "zuwanlu@xdf.cn",
    correctedRole: "助理主管",
    reason: "用户纠偏：岗位描述误写为“益智新师师训助理”。",
  },
  {
    teacherName: "陈利",
    employeeId: "376177",
    teacherEmail: "chenli104@xdf.cn",
    correctedRole: "助理主管",
    reason: "用户纠偏：岗位描述存在误差，按助理主管处理。",
  },
];

export const INSPECTION_ROLE_CORRECTION_EMAILS = new Set(
  INSPECTION_ROLE_CORRECTIONS.map((correction) => correction.teacherEmail.trim().toLocaleLowerCase()),
);
