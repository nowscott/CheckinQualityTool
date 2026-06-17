import { strToU8, Zip, ZipDeflate } from "fflate";
import { progress } from "./progress";
import type {
  ChatInfo,
  DataRow,
  ListInfo,
  MatchInfo,
  SheetDefinition,
  SourceNames,
  Whitelist,
} from "./types";
import { displayValue } from "./utils";
import { WHITELIST_COLUMNS } from "./whitelist";

const FINAL_COLUMNS = [
  "教师姓名", "上课日期", "上课时间", "学生姓名", "服务周", "发送情况",
  "序号", "教师邮箱", "匹配学员姓名", "姓名清洗说明", "该周课次数",
  "匹配结论", "命中关键词", "命中位置", "命中群名", "白名单命中", "白名单说明",
  "命中聊天时间", "匹配消息数", "校区", "项目组", "科目", "源名单行号",
] as const;

const DETAIL_COLUMNS = [
  "质检序号", "教师姓名", "教师邮箱", "原始学员姓名", "匹配学员姓名", "姓名清洗说明",
  "学员关键词_后两字", "学员关键词_末字",
  "匹配序号", "匹配强度", "命中位置", "命中关键词", "发送人名称", "有效教师邮箱",
  "邮箱来源", "群名/好友昵称", "聊天时间", "聊天内容", "源聊天行号",
] as const;

const CHAT_COLUMNS = [
  "有效教师邮箱", "邮箱来源", "发送人名称", "群名/好友昵称",
  "聊天时间", "聊天内容", "源聊天行号",
] as const;

function xmlEscape(value: unknown) {
  return displayValue(value)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function excelColumn(index: number) {
  let result = "";
  for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) {
    result = String.fromCharCode(65 + ((value - 1) % 26)) + result;
  }
  return result;
}

function cellXml(value: unknown, columnIndex: number, rowIndex: number, style = 0) {
  const reference = `${excelColumn(columnIndex)}${rowIndex}`;
  return `<c r="${reference}" t="inlineStr"${style ? ` s="${style}"` : ""}><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
}

function addZipTextFile(zip: Zip, filename: string, chunks: string | Iterable<string>) {
  const entry = new ZipDeflate(filename, { level: 6 });
  zip.add(entry);
  if (typeof chunks === "string") {
    entry.push(strToU8(chunks), true);
    return;
  }
  for (const chunk of chunks) entry.push(strToU8(chunk), false);
  entry.push(new Uint8Array(0), true);
}

function* worksheetChunks(
  rows: DataRow[],
  columns: readonly string[],
  widths: Record<string, number>,
  rowStyle?: (row: DataRow) => number,
) {
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

function concatenate(chunks: Uint8Array[]) {
  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const output = new Uint8Array(total);
  let offset = 0;
  chunks.forEach((chunk) => {
    output.set(chunk, offset);
    offset += chunk.length;
  });
  return output;
}

export function buildOutput(
  listInfo: ListInfo,
  chatInfo: ChatInfo,
  matchInfo: MatchInfo,
  whitelist: Whitelist,
  useSingle: boolean,
  weekLabel: string,
  sourceNames: SourceNames,
) {
  const explanation: DataRow[] = [
    { 项目: "生成时间", 值: new Date().toLocaleString("zh-CN", { hour12: false }) },
    { 项目: "质检名单文件", 值: sourceNames.list },
    { 项目: "聊天明细文件", 值: sourceNames.chat },
    { 项目: "服务周", 值: weekLabel || "未识别" },
    { 项目: "服务周识别方式", 值: listInfo.weekMode },
    { 项目: "课次周次分布", 值: listInfo.weekDistribution || "无有效日期" },
    { 项目: "名单工作表", 值: listInfo.sheetName },
    { 项目: "聊天工作表", 值: chatInfo.sheetName },
    { 项目: "名单去重规则", 值: "先清洗学员姓名，再按教师邮箱+匹配学员姓名合并，保留最早课次" },
    { 项目: "学员姓名清洗", 值: "白名单匹配学员姓名优先于清洗规则；未命中白名单时，四字中文姓名末尾为“一”则去除该标记，并循环去除空格及末尾数字、重复、学员/同学/学生、家长/妈妈/爸爸及兄弟姐妹称谓、初一至初三/高一至高三、括号备注等标记" },
    { 项目: "教师匹配规则", 值: "教师姓名仅用于展示；教师姓名带数字或与企微昵称不同时，仍直接按教师邮箱匹配" },
    { 项目: "教师姓名输出", 值: "以上传质检名单中的教师姓名为基准，追加名单教师邮箱@前本地部分末尾数字；聊天发送人昵称不参与生成" },
    { 项目: "邮箱规则", 值: "优先群聊发送人邮箱，为空时回退邮箱" },
    { 项目: "聊天清洗规则", 值: "删除私聊、无有效邮箱、发送方非员工、引用回复" },
    { 项目: "文本匹配规则", 值: "英文忽略大小写；中文按原字符匹配" },
    { 项目: "强匹配规则", 值: "教师邮箱一致，且群名或聊天内容包含学员名后两字" },
    { 项目: "弱匹配规则", 值: `清洗后不足两字自动使用末字；正常姓名末字弱匹配${useSingle ? "已启用" : "未启用"}` },
    { 项目: "白名单规则", 值: "优先按学员号关联；匹配学员姓名可在名单清洗前覆盖清洗结果。免检项不检查聊天，发送情况直接记为已发送，匹配结论保留“白名单免检”；别名项增加实际姓名后两字，聊天真实命中后才判已发送" },
    { 项目: "内置白名单数量", 值: whitelist.entries.length },
  ];
  Object.entries(listInfo.counts).forEach(([key, value]) => explanation.push({ 项目: `名单_${key}`, 值: value }));
  Object.entries(chatInfo.counts).forEach(([key, value]) => explanation.push({ 项目: `聊天_${key}`, 值: value }));
  Object.entries(matchInfo.counts).forEach(([key, value]) => explanation.push({ 项目: `匹配_${key}`, 值: value }));

  const sheets: SheetDefinition[] = [
    {
      name: "打卡结果",
      rows: matchInfo.finalRows,
      columns: FINAL_COLUMNS,
      widths: { 教师邮箱: 28, 命中群名: 36, 项目组: 28, 校区: 24 },
      rowStyle: (row) =>
        ["白名单免检", "别名匹配", "弱匹配"].includes(String(row.匹配结论))
          ? 4
          : row.发送情况 === "已发送" ? 2 : 3,
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
    {
      name: "内置白名单",
      rows: whitelist.entries,
      columns: WHITELIST_COLUMNS,
      widths: {
        学员号: 18, 学员姓名: 18, 匹配学员姓名: 18, 处理方式: 12, 匹配别名: 24, 说明: 60,
      },
    },
  ];

  const outputChunks: Uint8Array[] = [];
  const zip = new Zip((error, chunk) => {
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
    sheetOverrides + `</Types>`,
  );
  addZipTextFile(zip, "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
    `</Relationships>`,
  );
  addZipTextFile(zip, "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>` +
    sheets.map((sheet, index) => `<sheet name="${xmlEscape(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("") +
    `</sheets></workbook>`,
  );
  addZipTextFile(zip, "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    sheets.map((_, index) =>
      `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`
    ).join("") +
    `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    `</Relationships>`,
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
    `</styleSheet>`,
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
