const OUTPUT_SHEETS = [
  ["Sheet 1", "打卡结果", "核心结果列连续排列，清洗后单字姓名自动使用弱匹配。"],
  ["Sheet 2", "匹配明细", "记录命中关键词、位置、群名、聊天时间、内容和源行号。"],
  ["Sheet 3", "清洗后聊天", "仅保留后续匹配所需字段，移除无关列和无效记录。"],
  ["Sheet 4", "处理说明", "记录清洗规则、去重规则、各阶段数量和生成时间。"],
  ["Sheet 5", "内置白名单", "展示本次处理使用的学员号、别名和豁免说明。"],
] as const;

export function OutputGrid() {
  return (
    <section className="output-grid">
      {OUTPUT_SHEETS.map(([sheet, title, description]) => (
        <article key={sheet}>
          <span>{sheet}</span>
          <strong>{title}</strong>
          <p>{description}</p>
        </article>
      ))}
    </section>
  );
}
