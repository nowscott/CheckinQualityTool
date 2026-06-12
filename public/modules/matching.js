function listTeacherNameWithEmailSuffix(listTeacherName, listTeacherEmail) {
  const localPart = String(listTeacherEmail || "").split("@")[0];
  const emailDigits = localPart.match(/(\d+)$/)?.[1] || "";
  if (!emailDigits) return listTeacherName;
  return `${String(listTeacherName || "").replace(/[0-9０-９]+$/u, "")}${emailDigits}`;
}

async function matchData(targets, chats, useSingle, weekLabel, whitelist) {
  const chatsByEmail = new Map();
  chats.forEach((chat) => {
    if (!chatsByEmail.has(chat.有效教师邮箱)) chatsByEmail.set(chat.有效教师邮箱, []);
    chatsByEmail.get(chat.有效教师邮箱).push(chat);
  });

  const finalRows = [];
  const detailRows = [];
  const counts = {
    已发送: 0, 未发送: 0, 免检: 0,
    强匹配: 0, 弱匹配: 0, 别名匹配: 0, 白名单免检: 0, 无匹配: 0,
  };

  targets.forEach((target, targetIndex) => {
    const nameLength = [...target.学员姓名].length;
    const strong = nameLength >= 2 ? target.学员姓名.slice(-2) : "";
    const weakSource = target.学员姓名 || target.原始学员姓名;
    const automaticWeak = nameLength < 2;
    const weak = automaticWeak || useSingle ? weakSource.slice(-1) : "";
    const whitelistEntry = findWhitelistEntry(target, whitelist);
    const candidates = whitelistEntry?.处理方式 === "免检" ? [] : chatsByEmail.get(target.教师邮箱) || [];
    const matches = [];
    for (const chat of candidates) {
      const group = chat["群名/好友昵称"];
      const content = chat.聊天内容;
      const normalizedGroup = normalizeMatchText(group);
      const normalizedContent = normalizeMatchText(content);
      const normalizedStrong = normalizeMatchText(strong);
      const normalizedWeak = normalizeMatchText(weak);
      let keyword = "";
      let strength = "";
      const locations = [];
      if (normalizedStrong && normalizedGroup.includes(normalizedStrong)) locations.push("群名");
      if (normalizedStrong && normalizedContent.includes(normalizedStrong)) locations.push("聊天内容");
      if (locations.length) {
        keyword = strong;
        strength = "强匹配";
      } else if (weak) {
        if (normalizedGroup.includes(normalizedWeak)) locations.push("群名");
        if (normalizedContent.includes(normalizedWeak)) locations.push("聊天内容");
        if (locations.length) {
          keyword = weak;
          strength = "弱匹配";
        }
      }
      if (!locations.length && whitelistEntry?.处理方式 === "别名") {
        const aliasKeyword = whitelistEntry.匹配别名关键词.find((value) =>
          normalizedGroup.includes(value) || normalizedContent.includes(value)
        );
        if (aliasKeyword && normalizedGroup.includes(aliasKeyword)) locations.push("群名");
        if (aliasKeyword && normalizedContent.includes(aliasKeyword)) locations.push("聊天内容");
        if (locations.length) {
          keyword = aliasKeyword;
          strength = "别名匹配";
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
    const isExempt = whitelistEntry?.处理方式 === "免检";
    const status = isExempt || sent ? "已发送" : "未发送";
    const conclusion = isExempt ? "白名单免检" : best?.匹配强度 || "无匹配";
    counts[status] += 1;
    if (isExempt) counts.免检 += 1;
    if (Object.hasOwn(counts, conclusion)) counts[conclusion] += 1;
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
    matches.forEach((match, matchIndex) => detailRows.push({
      质检序号: id,
      教师姓名: target.教师姓名,
      教师邮箱: target.教师邮箱,
      原始学员姓名: target.原始学员姓名,
      匹配学员姓名: target.学员姓名,
      姓名清洗说明: target.姓名清洗说明,
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
