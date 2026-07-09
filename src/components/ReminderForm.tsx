import type { FormEvent } from "react";
import { MultiUploadCard } from "./MultiUploadCard";
import { UploadCard } from "./UploadCard";

interface ReminderFormProps {
  denominatorFile: File | null;
  appealFile: File | null;
  summaryFiles: File[];
  chatFiles: File[];
  includeCleanChats: boolean;
  includeResultColors: boolean;
  processing: boolean;
  onDenominatorFileChange: (file: File | null) => void;
  onAppealFileChange: (file: File | null) => void;
  onSummaryFilesChange: (files: File[]) => void;
  onChatFilesChange: (files: File[]) => void;
  onIncludeCleanChatsChange: (value: boolean) => void;
  onIncludeResultColorsChange: (value: boolean) => void;
  onSubmit: (event: FormEvent<HTMLFormElement>) => void;
}

export function ReminderForm({
  denominatorFile,
  appealFile,
  summaryFiles,
  chatFiles,
  includeCleanChats,
  includeResultColors,
  processing,
  onDenominatorFileChange,
  onAppealFileChange,
  onSummaryFilesChange,
  onChatFilesChange,
  onIncludeCleanChatsChange,
  onIncludeResultColorsChange,
  onSubmit,
}: ReminderFormProps) {
  return (
    <section className="card">
      <form id="reminder-form" onSubmit={onSubmit}>
        <div className="reminder-upload-layout">
          <div className="reminder-upload-column">
            <UploadCard
              id="reminder-list-file"
              name="reminder_list_file"
              step="01"
              title="开课提醒学员明细名单"
              description="授课教师、教研组、师训组长、助理主管、学员与课时"
              file={denominatorFile}
              onChange={onDenominatorFileChange}
            />
            <UploadCard
              id="reminder-appeal-file"
              name="reminder_appeal_file"
              step="02"
              title="申诉名单"
              description="可选；选择已发送的申诉计入分母，其他申诉公示原因并剔除分母"
              file={appealFile}
              required={false}
              onChange={onAppealFileChange}
            />
          </div>
          <div className="reminder-upload-column">
            <MultiUploadCard
              id="reminder-summary-file"
              name="reminder_summary_file"
              step="03"
              title="聊天质检汇总文件"
              description="可同时选择多个关键词汇总 Excel，教师及以上维度按触达完成率计算"
              files={summaryFiles}
              onChange={onSummaryFilesChange}
            />
            <MultiUploadCard
              id="reminder-chat-file"
              name="reminder_chat_file"
              step="04"
              title="企微聊天质检结果"
              description="可选；用于学员名单发送判断，不影响教师汇总；勾选下方才输出清洗后聊天"
              files={chatFiles}
              required={false}
              onChange={onChatFilesChange}
            />
          </div>
        </div>

        <div className="options reminder-options">
          <label className="switch-row">
            <span>
              <strong>包含聊天记录明细</strong>
              <small>默认不输出；需要核查命中来源时再勾选追加"清洗后聊天"Sheet</small>
            </span>
            <input
              id="include-reminder-chats"
              type="checkbox"
              checked={includeCleanChats}
              onChange={(event) => onIncludeCleanChatsChange(event.target.checked)}
            />
          </label>
          <label className="switch-row">
            <span>
              <strong>结果颜色标注</strong>
              <small>默认关闭；勾选后按通过、未通过、汇总行保留底色提示</small>
            </span>
            <input
              id="include-reminder-colors"
              type="checkbox"
              checked={includeResultColors}
              onChange={(event) => onIncludeResultColorsChange(event.target.checked)}
            />
          </label>
        </div>

        <div className="rules reminder-rules">
          <span>汇总口径</span>
          <b>整行去重</b>
          <b>多汇总合并</b>
          <b>客户+群聊触达</b>
          <b>按应发封顶</b>
          <b>明细仅参考</b>
        </div>

        <button id="reminder-submit-button" type="submit" disabled={processing}>
          <span>生成触达完成率公示</span>
          <small>
            默认输出公示表、学员名单、维度汇总、匹配核对/异常明细和处理说明
          </small>
        </button>
      </form>
    </section>
  );
}
