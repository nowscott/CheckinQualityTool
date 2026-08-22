import type { ChatRow, DataRow, MatchInfo, TargetRow, Whitelist } from "./types";
import { excelDate, normalizeMatchText, sortDate } from "./utils";
import { findWhitelistEntry } from "./whitelist";

interface MatchedChat extends ChatRow {
  匹配强度: string;
  命中位置: string;
  命中关键词: string;
  命中关键词来源: string;
}

interface NormalizedChat {
  index: number;
  chat: ChatRow;
  group: string;
  content: string;
}

interface KeywordHit {
  item: NormalizedChat;
  locations: string[];
}

interface TargetPlan {
  target: TargetRow;
  strong: string;
  normalizedStrong: string;
  whitelistNameKeyword: string;
  weak: string;
  normalizedWeak: string;
  aliasKeywords: string[];
  whitelistEntry: ReturnType<typeof findWhitelistEntry>;
  whitelistExempt: boolean;
}

function listTeacherNameWithEmailSuffix(listTeacherName: string, listTeacherEmail: string) {
  const localPart = String(listTeacherEmail || "").split("@")[0];
  const emailDigits = localPart.match(/(\d+)$/)?.[1] || "";
  if (!emailDigits) return listTeacherName;
  return `${String(listTeacherName || "").replace(/[0-9０-９]+$/u, "")}${emailDigits}`;
}

function hitLocations(item: NormalizedChat, keyword: string) {
  const locations: string[] = [];
  if (!keyword) return locations;
  if (item.group.includes(keyword)) locations.push("群名");
  if (item.content.includes(keyword)) locations.push("聊天内容");
  return locations;
}

function addKeyword(keywordsByEmail: Map<string, Set<string>>, email: string, keyword: string) {
  if (!email || !keyword) return;
  if (!keywordsByEmail.has(email)) keywordsByEmail.set(email, new Set<string>());
  keywordsByEmail.get(email)!.add(keyword);
}

function trailingNameKeyword(value: string) {
  return normalizeMatchText([...String(value || "").replace(/\s+/g, "")].slice(-2).join(""));
}

function collectShortKeywordHits(textValue: string, keywords: Set<string>, output: Set<string>) {
  for (let index = 0; index < textValue.length; index += 1) {
    const one = textValue[index];
    if (keywords.has(one)) output.add(one);
    if (index + 2 <= textValue.length) {
      const two = textValue.slice(index, index + 2);
      if (keywords.has(two)) output.add(two);
    }
  }
}

function buildHitIndex(targetPlans: TargetPlan[], chats: ChatRow[]) {
  const keywordsByEmail = new Map<string, Set<string>>();
  targetPlans.forEach((plan) => {
    if (plan.whitelistExempt) return;
    plan.aliasKeywords.forEach((keyword) => addKeyword(keywordsByEmail, plan.target.教师邮箱, keyword));
    addKeyword(keywordsByEmail, plan.target.教师邮箱, plan.whitelistNameKeyword);
    addKeyword(keywordsByEmail, plan.target.教师邮箱, plan.normalizedStrong);
    addKeyword(keywordsByEmail, plan.target.教师邮箱, plan.normalizedWeak);
  });

  const chatsByEmail = new Map<string, NormalizedChat[]>();
  chats.forEach((chat, index) => {
    if (!keywordsByEmail.has(chat.有效教师邮箱)) return;
    const item: NormalizedChat = {
      index,
      chat,
      group: normalizeMatchText(chat["群名/好友昵称"]),
      content: normalizeMatchText(chat.聊天内容),
    };
    if (!chatsByEmail.has(chat.有效教师邮箱)) chatsByEmail.set(chat.有效教师邮箱, []);
    chatsByEmail.get(chat.有效教师邮箱)!.push(item);
  });

  const hitsByEmail = new Map<string, Map<string, KeywordHit[]>>();
  for (const [email, items] of chatsByEmail) {
    const keywords = keywordsByEmail.get(email);
    if (!keywords?.size) continue;
    const hitsByKeyword = new Map<string, KeywordHit[]>();
    const longKeywords = [...keywords].filter((keyword) => keyword.length > 2);
    for (const item of items) {
      const matchedKeywords = new Set<string>();
      collectShortKeywordHits(item.group, keywords, matchedKeywords);
      collectShortKeywordHits(item.content, keywords, matchedKeywords);
      longKeywords.forEach((keyword) => {
        if (item.group.includes(keyword) || item.content.includes(keyword)) matchedKeywords.add(keyword);
      });
      for (const keyword of matchedKeywords) {
        const locations = hitLocations(item, keyword);
        if (!locations.length) continue;
        if (!hitsByKeyword.has(keyword)) hitsByKeyword.set(keyword, []);
        hitsByKeyword.get(keyword)!.push({ item, locations });
      }
    }
    hitsByEmail.set(email, hitsByKeyword);
  }
  return hitsByEmail;
}

function addMatchesFromHits(
  matchesByChat: Map<number, MatchedChat>,
  hitsByKeyword: Map<string, KeywordHit[]> | undefined,
  keyword: string,
  displayKeyword: string,
  strength: string,
  source: string,
) {
  const hits = hitsByKeyword?.get(keyword);
  if (!hits?.length) return;
  hits.forEach((hit) => {
    if (matchesByChat.has(hit.item.index)) return;
    matchesByChat.set(hit.item.index, {
      ...hit.item.chat,
      匹配强度: strength,
      命中位置: hit.locations.join("+"),
      命中关键词: displayKeyword,
      命中关键词来源: source,
    });
  });
}

export function matchData(
  targets: TargetRow[],
  chats: ChatRow[],
  useSingle: boolean,
  weekLabel: string,
  whitelist: Whitelist,
): MatchInfo {
  const targetPlans = targets.map((target): TargetPlan => {
    const nameLength = [...target.学员姓名].length;
    const strong = nameLength >= 2 ? target.学员姓名.slice(-2) : "";
    const weakSource = target.学员姓名 || target.原始学员姓名;
    const automaticWeak = nameLength < 2;
    const weak = automaticWeak || useSingle ? weakSource.slice(-1) : "";
    const whitelistEntry = findWhitelistEntry(target, whitelist);
    return {
      target,
      strong,
      normalizedStrong: normalizeMatchText(strong),
      whitelistNameKeyword: trailingNameKeyword(whitelistEntry?.匹配学员姓名 || ""),
      weak,
      normalizedWeak: normalizeMatchText(weak),
      aliasKeywords: whitelistEntry?.处理方式 === "别名" ? whitelistEntry.匹配别名关键词.filter(Boolean) : [],
      whitelistEntry,
      whitelistExempt: whitelistEntry?.处理方式 === "免检",
    };
  });
  const hitsByEmail = buildHitIndex(targetPlans, chats);

  const finalRows: DataRow[] = [];
  const detailRows: DataRow[] = [];
  const counts = {
    已发送: 0,
    未发送: 0,
    免检: 0,
    强匹配: 0,
    弱匹配: 0,
    别名匹配: 0,
    白名单免检: 0,
    无匹配: 0,
    匹配明细行数: 0,
  };

  targetPlans.forEach((plan, targetIndex) => {
    const { target, strong, weak, whitelistEntry, whitelistExempt } = plan;
    const hitsByKeyword = whitelistExempt ? undefined : hitsByEmail.get(target.教师邮箱);
    const matchesByChat = new Map<number, MatchedChat>();
    plan.aliasKeywords.forEach((keyword) =>
      addMatchesFromHits(matchesByChat, hitsByKeyword, keyword, keyword, "别名匹配", "白名单别名"),
    );
    addMatchesFromHits(matchesByChat, hitsByKeyword, plan.whitelistNameKeyword, plan.whitelistNameKeyword, "强匹配", "白名单登记名");
    addMatchesFromHits(matchesByChat, hitsByKeyword, plan.normalizedStrong, strong, "强匹配", "名单姓名后两字");
    addMatchesFromHits(matchesByChat, hitsByKeyword, plan.normalizedWeak, weak, "弱匹配", "名单姓名末字");
    const matches = [...matchesByChat.values()];
    const matchPriority: Record<string, number> = {
      别名匹配: 0,
      强匹配: 1,
      弱匹配: 2,
    };
    matches.sort(
      (a, b) =>
        (matchPriority[a.匹配强度] ?? 9) - (matchPriority[b.匹配强度] ?? 9) ||
        sortDate(a.聊天时间) - sortDate(b.聊天时间),
    );
    const best = matches[0];
    const status = whitelistExempt || Boolean(best) ? "已发送" : "未发送";
    const conclusion = whitelistExempt ? "白名单免检" : best?.匹配强度 || "无匹配";
    counts[status as "已发送" | "未发送"] += 1;
    if (whitelistExempt) counts.免检 += 1;
    if (conclusion in counts) {
      counts[conclusion as "强匹配" | "弱匹配" | "别名匹配" | "白名单免检" | "无匹配"] += 1;
    }
    const id = targetIndex + 1;
    finalRows.push({
      序号: id,
      教师姓名: listTeacherNameWithEmailSuffix(target.教师姓名, target.教师邮箱),
      教师邮箱: target.教师邮箱,
      学生姓名: target.原始学员姓名,
      匹配学员姓名: target.学员姓名,
      姓名清洗说明: target.姓名清洗说明,
      上课日期: excelDate(target.上课日期),
      上课时间: [target.上课开始, target.上课结束].filter(Boolean).join("-"),
      该周课次数: target.该周课次数,
      服务周: weekLabel,
      发送情况: status,
      匹配结论: conclusion,
      命中关键词: best?.命中关键词 || "",
      命中关键词来源: best?.命中关键词来源 || "",
      命中位置: best?.命中位置 || "",
      命中群名: best?.["群名/好友昵称"] || "",
      白名单命中: whitelistEntry ? "是" : "否",
      白名单说明: whitelistEntry?.说明 || "",
      命中聊天时间: best?.聊天时间 || "",
      匹配消息数: matches.length,
      校区: target.校区,
      项目组: target.项目组,
      科目: target.科目,
      源名单行号: target.源名单行号,
    });
    if (best) {
      detailRows.push({
        质检序号: id,
        教师姓名: target.教师姓名,
        教师邮箱: target.教师邮箱,
        原始学员姓名: target.原始学员姓名,
        匹配学员姓名: target.学员姓名,
        姓名清洗说明: target.姓名清洗说明,
        学员关键词_后两字: strong,
        学员关键词_末字: weak,
        匹配序号: 1,
        匹配强度: best.匹配强度,
        命中位置: best.命中位置,
        命中关键词: best.命中关键词,
        命中关键词来源: best.命中关键词来源,
        发送人名称: best.发送人名称,
        有效教师邮箱: best.有效教师邮箱,
        邮箱来源: best.邮箱来源,
        "群名/好友昵称": best["群名/好友昵称"],
        聊天时间: best.聊天时间,
        聊天内容: best.聊天内容,
        源聊天行号: best.源聊天行号,
      });
    }
  });
  counts.匹配明细行数 = detailRows.length;
  return { finalRows, detailRows, counts };
}
