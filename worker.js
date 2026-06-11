importScripts("./vendor/xlsx.full.min.js");
importScripts("./vendor/fflate.js");

const QUOTE_SEPARATOR = /(?:-\s*){8,}|[—－-]{12,}/;
const QUOTE_PREFIX = /^\s*[「『][\s\S]{0,500}?[：:]/;

const FINAL_COLUMNS = [
  "序号", "教师姓名", "教师邮箱", "学员姓名", "上课日期", "上课时间", "该周课次数",
  "服务周", "发送情况", "匹配结论", "命中关键词", "命中位置", "命中群名",
  "命中聊天时间", "匹配消息数", "校区", "项目组", "科目", "源名单行号",
];
const DETAIL_COLUMNS = [
  "质检序号", "教师姓名", "教师邮箱", "学员姓名", "学员关键词_后两字", "学员关键词_末字",
  "匹配序号", "匹配强度", "命中位置", "命中关键词", "发送人名称", "有效教师邮箱",
  "邮箱来源", "群名/好友昵称", "聊天时间", "聊天内容", "源聊天行号",
];
const CHAT_COLUMNS = [
  "有效教师邮箱", "邮箱来源", "发送人名称", "群名/好友昵称",
  "聊天时间", "聊天内容", "源聊天行号",
];

function progress(title, message, value) {
  postMessage({ type: "progress", title, message, progress: value });
}

function text(value) {
  return value == null ? "" : String(value).trim();
}

function emailValue(value) {
  const match = text(value).toLowerCase().match(/[\w.+-]+@[\w.-]+\.[a-z]{2,}/);
  return match ? match[0] : "";
}

function cleanStudentName(value) {
  return text(value).replace(/\s+/g, "").replace(/(学员|同学|学生)$/u, "");
}

function excelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
  }
  const match = text(value).match(/\d{4}[-/.]\d{1,2}[-/.]\d{1,2}/);
  return match ? match[0].replace(/[/.]/g, "-") : text(value);
}

function displayValue(value) {
  if (!(value instanceof Date) || Number.isNaN(value.valueOf())) return text(value);
  const pad = (number) => String(number).padStart(2, "0");
  return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())} ` +
    `${pad(value.getHours())}:${pad(value.getMinutes())}:${pad(value.getSeconds())}`;
}

function sortDate(value) {
  const normalized = excelDate(value);
  const timestamp = Date.parse(normalized);
  return Number.isNaN(timestamp) ? Number.MAX_SAFE_INTEGER : timestamp;
}

function weekOfMonth(value) {
  const normalized = excelDate(value);
  const match = normalized.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return 0;
  const year = Number(match[1]);
  const month = Number(match[2]) - 1;
  const day = Number(match[3]);
  const firstDay = new Date(year, month, 1);
  const mondayOffset = (firstDay.getDay() + 6) % 7;
  return Math.min(5, Math.floor((day + mondayOffset - 1) / 7) + 1);
}

function chineseWeek(number) {
  return ["", "第一周", "第二周", "第三周", "第四周", "第五周"][number] || "";
}

function inferServiceWeek(counts) {
  const ranked = [...counts.entries()].sort((a, b) => b[1] - a[1] || a[0] - b[0]);
  return {
    label: chineseWeek(ranked[0]?.[0] || 0),
    distribution: ranked.map(([week, count]) => `${chineseWeek(week)} ${count}课次`).join("；"),
  };
}

function headerMap(headerRow) {
  const map = new Map();
  headerRow.forEach((value, index) => {
    const key = text(value);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(index);
  });
  return map;
}

function findSheet(workbook, requiredHeaders) {
  for (const name of workbook.SheetNames) {
    const sheet = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, range: 0, blankrows: false, defval: "" });
    if (!rows.length) continue;
    const headers = new Set(rows[0].map(text));
    if (requiredHeaders.every((header) => headers.has(header))) return { name, sheet, rows };
  }
  throw new Error(`找不到包含字段“${requiredHeaders.join("、")}”的工作表。`);
}

async function readWorkbook(file, stageStart, stageEnd, label) {
  progress(`正在读取${label}`, `${file.name} · ${(file.size / 1024 / 1024).toFixed(1)} MB`, stageStart);
  const data = await file.arrayBuffer();
  progress(`正在解析${label}`, "使用 dense 模式解析 Excel，此步骤耗时取决于文件大小。", stageStart + (stageEnd - stageStart) * 0.35);
  const workbook = XLSX.read(data, { type: "array", dense: true, cellDates: true });
  progress(`${label}解析完成`, "正在提取必要字段并释放原始工作簿。", stageEnd);
  return workbook;
}

function buildTargets(workbook) {
  const found = findSheet(workbook, ["老师姓名", "学员姓名", "老师邮箱", "课次开始时"]);
  const rows = found.rows;
  const map = headerMap(rows[0]);
  const index = (name, occurrence = 0) => (map.get(name) || [])[occurrence] ?? -1;
  const columns = {
    teacher: index("老师姓名"),
    student: index("学员姓名"),
    studentId: index("学员号"),
    lessonDate: index("课次开始时"),
    start: index("间", 0),
    end: index("间", 1),
    campus: index("校区"),
    project: index("项目组"),
    subject: index("科目"),
    email: index("老师邮箱"),
  };
  const grouped = new Map();
  const weekCounts = new Map();
  const counts = { 原始课次行数: Math.max(0, rows.length - 1), 跳过教师或学员为空: 0, 名单教师邮箱为空: 0 };

  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    const teacher = text(row[columns.teacher]);
    const student = cleanStudentName(row[columns.student]);
    const teacherEmail = emailValue(row[columns.email]);
    if (!teacher || !student) {
      counts.跳过教师或学员为空 += 1;
      continue;
    }
    if (!teacherEmail) counts.名单教师邮箱为空 += 1;
    const lessonWeek = weekOfMonth(row[columns.lessonDate]);
    if (lessonWeek) weekCounts.set(lessonWeek, (weekCounts.get(lessonWeek) || 0) + 1);
    const key = `${teacherEmail || `__name__:${teacher}`}\u0000${student}`;
    const record = {
      教师姓名: teacher,
      教师邮箱: teacherEmail,
      学员姓名: student,
      学员号: text(row[columns.studentId]),
      上课日期: row[columns.lessonDate],
      上课开始: text(row[columns.start]),
      上课结束: text(row[columns.end]),
      校区: text(row[columns.campus]),
      项目组: text(row[columns.project]),
      科目: text(row[columns.subject]),
      源名单行号: rowIndex + 1,
    };
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(record);
  }

  const targets = [];
  for (const records of grouped.values()) {
    records.sort((a, b) =>
      sortDate(a.上课日期) - sortDate(b.上课日期) ||
      a.上课开始.localeCompare(b.上课开始) ||
      a.源名单行号 - b.源名单行号
    );
    targets.push({ ...records[0], 该周课次数: records.length });
  }
  targets.sort((a, b) => a.教师姓名.localeCompare(b.教师姓名, "zh-CN") || a.学员姓名.localeCompare(b.学员姓名, "zh-CN"));
  counts.去重后质检人数 = targets.length;
  counts.合并的重复课次 = counts.原始课次行数 - targets.length - counts.跳过教师或学员为空;
  return { targets, counts, weekCounts, sheetName: found.name };
}

function preprocessChats(workbook) {
  const found = findSheet(workbook, ["聊天类型", "发送方", "聊天内容"]);
  const rows = found.rows;
  const map = headerMap(rows[0]);
  const index = (name) => (map.get(name) || [])[0] ?? -1;
  const columns = {
    name: index("姓名"),
    email: index("邮箱"),
    type: index("聊天类型"),
    sender: index("发送方"),
    group: index("群名/好友昵称"),
    groupSender: index("群聊发送人名称"),
    groupEmail: index("群聊发送人邮箱"),
    time: index("聊天时间"),
    content: index("聊天内容"),
  };
  const chats = [];
  const counts = {
    原始聊天行数: Math.max(0, rows.length - 1),
    检测到引用回复: 0,
    删除私聊: 0,
    删除无有效邮箱: 0,
    删除发送方非员工: 0,
    删除引用回复: 0,
  };

  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    const chatType = text(row[columns.type]);
    const sender = text(row[columns.sender]);
    const primaryEmail = emailValue(row[columns.groupEmail]);
    const fallbackEmail = emailValue(row[columns.email]);
    const effectiveEmail = primaryEmail || fallbackEmail;
    const content = text(row[columns.content]);
    const isQuotedReply = QUOTE_SEPARATOR.test(content) && QUOTE_PREFIX.test(content);
    if (isQuotedReply) counts.检测到引用回复 += 1;
    if (chatType === "私聊") {
      counts.删除私聊 += 1;
      continue;
    }
    if (!effectiveEmail) {
      counts.删除无有效邮箱 += 1;
      continue;
    }
    if (sender && sender !== "员工") {
      counts.删除发送方非员工 += 1;
      continue;
    }
    if (isQuotedReply) {
      counts.删除引用回复 += 1;
      continue;
    }
    chats.push({
      有效教师邮箱: effectiveEmail,
      邮箱来源: primaryEmail ? "群聊发送人邮箱" : "邮箱",
      发送人名称: text(row[columns.groupSender]) || text(row[columns.name]),
      "群名/好友昵称": text(row[columns.group]),
      聊天时间: row[columns.time],
      聊天内容: content,
      源聊天行号: rowIndex + 1,
    });
  }
  counts.清洗后聊天行数 = chats.length;
  return { chats, counts, sheetName: found.name };
}

function matchData(targets, chats, useSingle, weekLabel) {
  const chatsByEmail = new Map();
  chats.forEach((chat) => {
    if (!chatsByEmail.has(chat.有效教师邮箱)) chatsByEmail.set(chat.有效教师邮箱, []);
    chatsByEmail.get(chat.有效教师邮箱).push(chat);
  });

  const finalRows = [];
  const detailRows = [];
  const counts = { 已发送: 0, 未发送: 0, 强匹配: 0, 弱匹配: 0, 无匹配: 0 };

  targets.forEach((target, targetIndex) => {
    const strong = target.学员姓名.slice(-2);
    const weak = useSingle ? target.学员姓名.slice(-1) : "";
    const candidates = chatsByEmail.get(target.教师邮箱) || [];
    const matches = [];
    for (const chat of candidates) {
      const group = chat["群名/好友昵称"];
      const content = chat.聊天内容;
      let keyword = "";
      let strength = "";
      const locations = [];
      if (strong && group.includes(strong)) locations.push("群名");
      if (strong && content.includes(strong)) locations.push("聊天内容");
      if (locations.length) {
        keyword = strong;
        strength = "强匹配";
      } else if (weak) {
        if (group.includes(weak)) locations.push("群名");
        if (content.includes(weak)) locations.push("聊天内容");
        if (locations.length) {
          keyword = weak;
          strength = "弱匹配";
        }
      }
      if (locations.length) {
        matches.push({ ...chat, 匹配强度: strength, 命中位置: locations.join("+"), 命中关键词: keyword });
      }
    }
    matches.sort((a, b) =>
      (a.匹配强度 === "强匹配" ? 0 : 1) - (b.匹配强度 === "强匹配" ? 0 : 1) ||
      sortDate(a.聊天时间) - sortDate(b.聊天时间)
    );
    const best = matches[0];
    const sent = Boolean(best);
    const conclusion = best?.匹配强度 || "无匹配";
    counts[sent ? "已发送" : "未发送"] += 1;
    counts[conclusion] += 1;
    const id = targetIndex + 1;
    finalRows.push({
      序号: id,
      教师姓名: target.教师姓名,
      教师邮箱: target.教师邮箱,
      学员姓名: target.学员姓名,
      上课日期: excelDate(target.上课日期),
      上课时间: [target.上课开始, target.上课结束].filter(Boolean).join("-"),
      该周课次数: target.该周课次数,
      服务周: weekLabel,
      发送情况: sent ? "已发送" : "未发送",
      匹配结论: conclusion,
      命中关键词: best?.命中关键词 || "",
      命中位置: best?.命中位置 || "",
      命中群名: best?.["群名/好友昵称"] || "",
      命中聊天时间: best?.聊天时间 || "",
      匹配消息数: matches.length,
      校区: target.校区,
      项目组: target.项目组,
      科目: target.科目,
      源名单行号: target.源名单行号,
    });
    matches.forEach((match, matchIndex) => detailRows.push({
      质检序号: id,
      教师姓名: target.教师姓名,
      教师邮箱: target.教师邮箱,
      学员姓名: target.学员姓名,
      学员关键词_后两字: strong,
      学员关键词_末字: weak,
      匹配序号: matchIndex + 1,
      匹配强度: match.匹配强度,
      命中位置: match.命中位置,
      命中关键词: match.命中关键词,
      发送人名称: match.发送人名称,
      有效教师邮箱: match.有效教师邮箱,
      邮箱来源: match.邮箱来源,
      "群名/好友昵称": match["群名/好友昵称"],
      聊天时间: match.聊天时间,
      聊天内容: match.聊天内容,
      源聊天行号: match.源聊天行号,
    }));
  });
  counts.匹配明细行数 = detailRows.length;
  return { finalRows, detailRows, counts };
}

function xmlEscape(value) {
  return displayValue(value)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function excelColumn(index) {
  let result = "";
  for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) {
    result = String.fromCharCode(65 + ((value - 1) % 26)) + result;
  }
  return result;
}

function cellXml(value, columnIndex, rowIndex, style = 0) {
  const reference = `${excelColumn(columnIndex)}${rowIndex}`;
  return `<c r="${reference}" t="inlineStr"${style ? ` s="${style}"` : ""}><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
}

function addZipTextFile(zip, filename, chunks) {
  const entry = new fflate.ZipDeflate(filename, { level: 6 });
  zip.add(entry);
  if (typeof chunks === "string") {
    entry.push(fflate.strToU8(chunks), true);
    return;
  }
  for (const chunk of chunks) entry.push(fflate.strToU8(chunk), false);
  entry.push(new Uint8Array(0), true);
}

function* worksheetChunks(rows, columns, widths, rowStyle) {
  const lastCell = `${excelColumn(columns.length - 1)}${rows.length + 1}`;
  const colsXml = columns.map((column, index) => {
    const width = widths[column] || Math.min(Math.max(column.length * 2 + 2, 12), 24);
    return `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`;
  }).join("");
  yield (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<dimension ref="A1:${lastCell}"/><sheetViews><sheetView workbookViewId="0">` +
    `<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>` +
    `</sheetView></sheetViews><cols>${colsXml}</cols><sheetData>` +
    `<row r="1">${columns.map((column, index) => cellXml(column, index, 1, 1)).join("")}</row>`
  );
  let buffer = "";
  for (let rowOffset = 0; rowOffset < rows.length; rowOffset += 1) {
    const row = rows[rowOffset];
    const rowIndex = rowOffset + 2;
    const style = rowStyle ? rowStyle(row) : 0;
    buffer += `<row r="${rowIndex}">${columns.map((column, columnIndex) =>
      cellXml(row[column] ?? "", columnIndex, rowIndex, style)
    ).join("")}</row>`;
    if (buffer.length >= 512 * 1024) {
      yield buffer;
      buffer = "";
    }
  }
  if (buffer) yield buffer;
  yield `</sheetData><autoFilter ref="A1:${lastCell}"/></worksheet>`;
}

function concatenate(chunks) {
  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const output = new Uint8Array(total);
  let offset = 0;
  chunks.forEach((chunk) => {
    output.set(chunk, offset);
    offset += chunk.length;
  });
  return output;
}

function buildOutput(listInfo, chatInfo, matchInfo, useSingle, weekLabel, sourceNames) {
  const explanation = [
    { 项目: "生成时间", 值: new Date().toLocaleString("zh-CN", { hour12: false }) },
    { 项目: "质检名单文件", 值: sourceNames.list },
    { 项目: "聊天明细文件", 值: sourceNames.chat },
    { 项目: "服务周", 值: weekLabel || "未识别" },
    { 项目: "服务周识别方式", 值: listInfo.weekMode },
    { 项目: "课次周次分布", 值: listInfo.weekDistribution || "无有效日期" },
    { 项目: "名单工作表", 值: listInfo.sheetName },
    { 项目: "聊天工作表", 值: chatInfo.sheetName },
    { 项目: "名单去重规则", 值: "按教师邮箱+学员姓名合并，保留最早课次" },
    { 项目: "邮箱规则", 值: "优先群聊发送人邮箱，为空时回退邮箱" },
    { 项目: "聊天清洗规则", 值: "删除私聊、无有效邮箱、发送方非员工、引用回复" },
    { 项目: "强匹配规则", 值: "教师邮箱一致，且群名或聊天内容包含学员名后两字" },
    { 项目: "弱匹配规则", 值: useSingle ? "已启用：强匹配失败后使用学员名末字" : "未启用" },
    { 项目: "发送判定", 值: "强匹配或弱匹配任一命中即为已发送" },
  ];
  Object.entries(listInfo.counts).forEach(([key, value]) => explanation.push({ 项目: `名单_${key}`, 值: value }));
  Object.entries(chatInfo.counts).forEach(([key, value]) => explanation.push({ 项目: `聊天_${key}`, 值: value }));
  Object.entries(matchInfo.counts).forEach(([key, value]) => explanation.push({ 项目: `匹配_${key}`, 值: value }));

  const sheets = [
    {
      name: "打卡终版",
      rows: matchInfo.finalRows,
      columns: FINAL_COLUMNS,
      widths: { 教师邮箱: 28, 命中群名: 36, 项目组: 28, 校区: 24 },
      rowStyle: (row) => row.匹配结论 === "弱匹配" ? 4 : row.发送情况 === "已发送" ? 2 : 3,
    },
    {
      name: "匹配明细",
      rows: matchInfo.detailRows,
      columns: DETAIL_COLUMNS,
      widths: { 教师邮箱: 28, 有效教师邮箱: 28, "群名/好友昵称": 38, 聊天内容: 80 },
    },
    {
      name: "清洗后聊天",
      rows: chatInfo.chats,
      columns: CHAT_COLUMNS,
      widths: { 有效教师邮箱: 28, "群名/好友昵称": 38, 聊天内容: 80 },
    },
    {
      name: "处理说明",
      rows: explanation,
      columns: ["项目", "值"],
      widths: { 项目: 34, 值: 90 },
    },
  ];

  const outputChunks = [];
  const zip = new fflate.Zip((error, chunk) => {
    if (error) throw error;
    outputChunks.push(chunk);
  });
  const sheetOverrides = sheets.map((_, index) =>
    `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
  ).join("");
  addZipTextFile(zip, "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
    sheetOverrides + `</Types>`
  );
  addZipTextFile(zip, "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
    `</Relationships>`
  );
  addZipTextFile(zip, "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>` +
    sheets.map((sheet, index) => `<sheet name="${xmlEscape(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("") +
    `</sheets></workbook>`
  );
  addZipTextFile(zip, "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    sheets.map((_, index) =>
      `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`
    ).join("") +
    `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    `</Relationships>`
  );
  addZipTextFile(zip, "xl/styles.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts>` +
    `<fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>` +
    `<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>` +
    `<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill>` +
    `<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill>` +
    `<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill></fills>` +
    `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
    `<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
    `<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>` +
    `<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>` +
    `<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>` +
    `<xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/></cellXfs>` +
    `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
    `</styleSheet>`
  );
  sheets.forEach((sheet, index) => {
    progress("正在流式写入 Excel", `${sheet.name}：${sheet.rows.length.toLocaleString()} 行`, 85 + index * 3);
    addZipTextFile(
      zip,
      `xl/worksheets/sheet${index + 1}.xml`,
      worksheetChunks(sheet.rows, sheet.columns, sheet.widths, sheet.rowStyle),
    );
  });
  zip.end();
  return concatenate(outputChunks);
}

self.onmessage = async ({ data }) => {
  if (data.type !== "process") return;
  try {
    const listWorkbook = await readWorkbook(data.listFile, 3, 18, "课堂反馈名单");
    const listInfo = buildTargets(listWorkbook);
    progress("名单预处理完成", `原始 ${listInfo.counts.原始课次行数.toLocaleString()} 条，去重后 ${listInfo.targets.length.toLocaleString()} 人。`, 24);

    const chatWorkbook = await readWorkbook(data.chatFile, 25, 48, "聊天明细");
    const chatInfo = preprocessChats(chatWorkbook);
    progress(
      "聊天预处理完成",
      `原始 ${chatInfo.counts.原始聊天行数.toLocaleString()} 条，清洗后 ${chatInfo.chats.length.toLocaleString()} 条。`,
      58,
    );

    progress("正在匹配教师与学员", "按教师邮箱建立索引，再检查群名和聊天内容中的学员关键词。", 64);
    const inferredWeek = inferServiceWeek(listInfo.weekCounts);
    const selectedWeek = data.weekLabel === "auto" ? inferredWeek.label : data.weekLabel;
    listInfo.weekMode = data.weekLabel === "auto" ? "根据课次日期自动识别" : "人工指定";
    listInfo.weekDistribution = inferredWeek.distribution;
    if (!selectedWeek) throw new Error("无法从课次日期识别服务周，请手动选择第一周至第五周。");
    progress(
      "服务周识别完成",
      `${selectedWeek}（${listInfo.weekMode}）${inferredWeek.distribution ? `；${inferredWeek.distribution}` : ""}`,
      61,
    );

    const matchInfo = matchData(listInfo.targets, chatInfo.chats, data.useSingle, selectedWeek);
    progress(
      "匹配完成",
      `已发送 ${matchInfo.counts.已发送.toLocaleString()}，未发送 ${matchInfo.counts.未发送.toLocaleString()}。`,
      80,
    );

    progress("正在生成 Excel", "写入打卡终版、匹配明细、清洗后聊天和处理说明。", 84);
    const output = buildOutput(
      listInfo,
      chatInfo,
      matchInfo,
      data.useSingle,
      selectedWeek,
      { list: data.listFile.name, chat: data.chatFile.name },
    );
    const stamp = new Date().toISOString().slice(0, 19).replace(/[-:T]/g, "");
    postMessage({
      type: "complete",
      buffer: output.buffer,
      filename: `打卡质检结果_${stamp}.xlsx`,
      summary: {
        targets: listInfo.targets.length,
        sent: matchInfo.counts.已发送,
        unsent: matchInfo.counts.未发送,
        cleanChats: chatInfo.chats.length,
      },
    }, [output.buffer]);
  } catch (error) {
    postMessage({ type: "error", message: error?.stack || error?.message || String(error) });
  }
};
