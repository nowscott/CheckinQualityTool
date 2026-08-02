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
            title="阶段性报告分母"
            description="需要教师姓名、邮箱、学号、学员姓名；其余列完整保留"
            file={denominatorFile}
            onChange={onDenominatorFileChange}
          />
          <MultiUploadCard
            id="stage-report-chat-file"
            name="stage_report_chat_file"
            step="02"
            title="企微聊天质检结果"
            description="可同时选择多个 Excel；同条消息须含“阶段性报告”和学员姓名"
            files={chatFiles}
            onChange={onChatFilesChange}
          />
        </div>

        <div className="rules reminder-rules">
          <span>检查口径</span>
          <b>教师邮箱一致</b>
          <b>同条消息双命中</b>
          <b>群名仅作证据</b>
          <b>申诉原列不覆盖</b>
        </div>

        <button id="stage-report-submit-button" type="submit" disabled={processing}>
          <span>生成阶段性报告检查结果</span>
          <small>默认输出检查明细、教师发送汇总和处理说明</small>
        </button>
      </form>
    </section>
  );
}
