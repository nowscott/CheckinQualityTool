const OUTPUT_SHEETS = [
  ["Sheet 1", "打卡结果", "核心结果列连续排列，清洗后单字姓名自动使用弱匹配。"],
  ["Sheet 2", "匹配明细", "记录命中关键词、位置、群名、聊天时间、内容和源行号。"],
  ["Sheet 3", "清洗后聊天", "仅保留后续匹配所需字段，移除无关列和无效记录。"],
  ["Sheet 4", "处理说明", "记录清洗规则、去重规则、各阶段数量和生成时间。"],
  ["Sheet 5", "内置白名单", "展示本次处理使用的学员号、别名和豁免说明。"],
] as const;

const REMINDER_OUTPUT_SHEETS = [
  ["Sheet 1", "学员名单", "对齐春季明细结构，保留教师、邮箱、教研组、学员、发送状态和申诉说明。"],
  ["Sheet 2", "教师维度发送进度", "按教师、教研组、师训组长和助理主管汇总发送进度。"],
  ["Sheet 3", "师训组维度", "按项目、教研组和师训组长分层汇总，阈值为 80%。"],
  ["Sheet 4", "助理主管维度", "沿用项目和教研组小计结构，助理主管按本人所属教研组归类。"],
  ["Sheet 5", "项目组维度", "按教研组和项目组汇总发送进度，结构贴近春季示例。"],
  ["辅助", "匹配核对-异常明细", "记录字段缺失、多命中和无法唯一匹配等核对项。"],
  ["可选", "清洗后聊天", "默认不输出；勾选后追加聊天来源、群名、时间、内容和源行号。"],
] as const;

interface OutputGridProps {
  mode?: "checkin" | "reminder";
}

export function OutputGrid({ mode = "checkin" }: OutputGridProps) {
  const sheets = mode === "reminder" ? REMINDER_OUTPUT_SHEETS : OUTPUT_SHEETS;
  return (
    <section className="output-grid">
      {sheets.map(([sheet, title, description]) => (
        <article key={sheet}>
          <span>{sheet}</span>
          <strong>{title}</strong>
          <p>{description}</p>
        </article>
      ))}
    </section>
  );
}
