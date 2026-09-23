export interface TeacherScoreRow {
  teacherName: string;
  researchGroup: string;
  trainingLeader: string;
  score: number;
}

export interface ParsedTeacherScores {
  rows: TeacherScoreRow[];
  rowCount: number;
  missingNameRows: number;
  missingScoreRows: number;
}

const HEADER_ALIASES = {
  teacherName: ["教师姓名", "老师姓名", "姓名"],
  researchGroup: ["教研组", "教研组名称"],
  trainingLeader: ["师训组长", "师训负责人"],
  score: ["教学服务赋分", "教学服务评分", "教学服务成绩", "评分", "赋分"],
} as const;

function normalize(value: unknown) {
  return String(value ?? "").normalize("NFKC").replace(/[\p{White_Space}\p{Cf}]/gu, "").trim();
}

function scoreValue(value: unknown) {
  const normalized = normalize(value).replace(/,/gu, "").replace(/分$/u, "").replace(/%$/u, "");
  if (!normalized) return undefined;
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : undefined;
}

export function parseTeacherScoreCsv(csv: string): ParsedTeacherScores {
  const sourceRows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;
  const source = String(csv || "").replace(/^\uFEFF/u, "");

  for (let index = 0; index < source.length; index += 1) {
    const character = source[index];
    if (quoted) {
      if (character === '"' && source[index + 1] === '"') {
        field += '"';
        index += 1;
      } else if (character === '"') {
        quoted = false;
      } else {
        field += character;
      }
    } else if (character === '"') {
      quoted = true;
    } else if (character === ",") {
      row.push(field);
      field = "";
    } else if (character === "\n") {
      row.push(field.replace(/\r$/u, ""));
      sourceRows.push(row);
      row = [];
      field = "";
    } else {
      field += character;
    }
  }
  if (field || row.length) {
    row.push(field.replace(/\r$/u, ""));
    sourceRows.push(row);
  }

  const nonemptyRows = sourceRows.filter((candidate) => candidate.some((value) => normalize(value)));
  const headerIndex = nonemptyRows.findIndex((candidate) =>
    Object.values(HEADER_ALIASES).every((aliases) => aliases.some((alias) => candidate.some((cell) => normalize(cell) === alias))),
  );
  if (headerIndex < 0) {
    throw new Error("评分页表头不完整，需要教师姓名、教研组、师训组长、教学服务赋分四列。");
  }

  const headers = nonemptyRows[headerIndex].map(normalize);
  const indexOf = (aliases: readonly string[]) => aliases.map((alias) => headers.indexOf(alias)).find((index) => index >= 0) ?? -1;
  const indexes = {
    teacherName: indexOf(HEADER_ALIASES.teacherName),
    researchGroup: indexOf(HEADER_ALIASES.researchGroup),
    trainingLeader: indexOf(HEADER_ALIASES.trainingLeader),
    score: indexOf(HEADER_ALIASES.score),
  };
  const dataRows = nonemptyRows.slice(headerIndex + 1);
  let missingNameRows = 0;
  let missingScoreRows = 0;
  const rows: TeacherScoreRow[] = [];

  for (const candidate of dataRows) {
    const teacherName = String(candidate[indexes.teacherName] ?? "").trim();
    const score = scoreValue(candidate[indexes.score]);
    if (!teacherName) {
      missingNameRows += 1;
      continue;
    }
    if (score === undefined) {
      missingScoreRows += 1;
      continue;
    }
    rows.push({
      teacherName,
      researchGroup: String(candidate[indexes.researchGroup] ?? "").trim(),
      trainingLeader: String(candidate[indexes.trainingLeader] ?? "").trim(),
      score,
    });
  }

  if (!rows.length) throw new Error("评分页没有可用的教师赋分记录。");
  return { rows, rowCount: dataRows.length, missingNameRows, missingScoreRows };
}
