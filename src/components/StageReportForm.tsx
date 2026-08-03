import type { FormEvent } from "react";
import { MultiUploadCard } from "./MultiUploadCard";
import { UploadCard } from "./UploadCard";

interface StageReportFormProps {
  denominatorFile: File | null;
  chatFiles: File[];
  processing: boolean;
  onDenominatorFileChange: (file: File | null) => void;
  onChatFilesChange: (files: File[]) => void;
  onSubmit: (event: FormEvent<HTMLFormElement>) => void;
}

export function StageReportForm({
  denominatorFile,
  chatFiles,
  processing,
  onDenominatorFileChange,
  onChatFilesChange,
  onSubmit,
}: StageReportFormProps) {
  return (
    <section className="card">
      <form id="stage-report-form" onSubmit={onSubmit}>
        <div className="grid">
          <UploadCard
            id="stage-report-denominator-file"
            name="stage_report_denominator_file"
            step="01"
            title="钉钉阶段性报告表单分母"
            description="直接上传钉钉导出的表单；明细保留原列，公示表完整复刻暑假督课样式"
            file={denominatorFile}
            onChange={onDenominatorFileChange}
          />
          <MultiUploadCard
            id="stage-report-chat-file"
            name="stage_report_chat_file"
            step="02"
            title="企微聊天质检结果"
            description="使用“阶段性报告”检索导出的 Excel；不重复检查关键词，学员按打卡口径匹配后两字"
            files={chatFiles}
            onChange={onChatFilesChange}
          />
        </div>

        <div className="rules reminder-rules">
          <span>检查口径</span>
          <b>邮箱优先/姓名兜底</b>
          <b>后两字优先匹配</b>
          <b>群名或正文命中</b>
          <b>五个公示 Sheet</b>
          <b>申诉原列不覆盖</b>
        </div>

        <button id="stage-report-submit-button" type="submit" disabled={processing}>
          <span>生成阶段性报告发送进度</span>
          <small>输出发送明细及教师、师训组长、助理主管、教研组公示表</small>
        </button>
      </form>
    </section>
  );
}
