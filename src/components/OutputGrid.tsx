const OUTPUT_SHEETS = [
  ["Sheet 1", "打卡结果", "核心结果列连续排列，清洗后单字姓名自动使用弱匹配。"],
  ["Sheet 2", "匹配明细", "记录命中关键词、位置、群名、聊天时间、内容和源行号。"],
  ["Sheet 3", "清洗后聊天", "仅保留后续匹配所需字段，移除无关列和无效记录。"],
  ["Sheet 4", "处理说明", "记录清洗规则、去重规则、各阶段数量和生成时间。"],
  ["Sheet 5", "内置白名单", "展示本次处理使用的学员号、别名和豁免说明。"],
] as const;

const REMINDER_OUTPUT_SHEETS = [
  ["Sheet 1", "学员名单", "对齐春季明细结构，保留教师、邮箱、教研组、学员、发送状态和申诉说明。"],
  ["Sheet 2", "教师维度发送进度", "按汇总文件累计触达客户和群聊，计算教师触达完成率。"],
  ["Sheet 3", "师训组维度", "由教师维度向上汇总有效触达数，阈值为 80%。"],
  ["Sheet 4", "助理主管维度", "沿用项目和教研组小计结构，助理主管按本人所属教研组归类。"],
  ["Sheet 5", "项目组维度", "按教研组和项目组汇总触达完成率，结构贴近春季示例。"],
  ["辅助", "匹配核对-异常明细", "记录字段缺失、多命中和无法唯一匹配等核对项。"],
  ["可选", "清洗后聊天", "默认不输出；勾选后追加聊天来源、群名、时间、内容和源行号。"],
] as const;

const STAGE_REPORT_OUTPUT_SHEETS = [
  ["Sheet 1", "发送明细", "首行保留钉钉表单原表头，只回填“是否发送阶段性报告”为是/否。"],
  ["Sheet 2", "教师维度", "按教师汇总应发、已发、发送率、未发和申诉数。"],
  ["Sheet 3", "师训组长维度", "按暑假督课公示样式展示师训组长、关联教研组和发送率数据条。"],
  ["Sheet 4", "助理主管维度", "完整保留教研组层级合并、教研组小计、项目小计和底部总计。"],
  ["Sheet 5", "教研组维度", "按教研组及项目汇总，负责人沿用助理主管口径。"],
] as const;

const STAGE_REPORT_BEAUTIFY_OUTPUT_SHEETS = [
  ["Sheet 1", "阶段性报告明细", "保留原始非窗口期明细与发送结论，统一表头、列宽和发送状态颜色。"],
  ["Sheet 2", "窗口期报告明细", "保留窗口期报告原始明细和发送结论，不覆盖任何业务字段。"],
  ["Sheet 3", "阶段性报告申诉情况", "保留阶段性报告申诉原始信息，不覆盖任何业务结论。"],
  ["Sheet 4", "教研组维度", "按项目、教研组分层汇总两类报告，不展示负责人。"],
  ["Sheet 5", "助理主管维度", "按项目、教研组、助理主管分层汇总两类报告，保留各级小计和总计。"],
  ["Sheet 6", "师训组长维度", "按师训组长整理两类报告发送率，附发送率数据条。"],
  ["Sheet 7", "教师维度", "保留阶段性报告与窗口期报告双口径，以及原始完成和通知字段。"],
] as const;

interface OutputGridProps {
  mode?: "checkin" | "reminder" | "stageReport" | "stageReportBeautify";
}

export function OutputGrid({ mode = "checkin" }: OutputGridProps) {
  const sheets = mode === "reminder"
    ? REMINDER_OUTPUT_SHEETS
    : mode === "stageReport"
      ? STAGE_REPORT_OUTPUT_SHEETS
      : mode === "stageReportBeautify"
        ? STAGE_REPORT_BEAUTIFY_OUTPUT_SHEETS
        : OUTPUT_SHEETS;
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
