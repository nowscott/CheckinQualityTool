import type { FormEvent } from "react";
import { MultiUploadCard } from "./MultiUploadCard";
import { UploadCard } from "./UploadCard";

interface ReminderFormProps {
  reminderMode: "full" | "incremental";
  denominatorFile: File | null;
  previousFile: File | null;
  chatFiles: File[];
  includeCleanChats: boolean;
  processing: boolean;
  onReminderModeChange: (value: "full" | "incremental") => void;
  onDenominatorFileChange: (file: File | null) => void;
  onPreviousFileChange: (file: File | null) => void;
  onChatFilesChange: (files: File[]) => void;
  onIncludeCleanChatsChange: (value: boolean) => void;
  onSubmit: (event: FormEvent<HTMLFormElement>) => void;
}

export function ReminderForm({
  reminderMode,
  denominatorFile,
  previousFile,
  chatFiles,
  includeCleanChats,
  processing,
  onReminderModeChange,
  onDenominatorFileChange,
  onPreviousFileChange,
  onChatFilesChange,
  onIncludeCleanChatsChange,
  onSubmit,
}: ReminderFormProps) {
  return (
    <section className="card">
      <form id="reminder-form" onSubmit={onSubmit}>
        <div className="segmented">
          <label className={reminderMode === "full" ? "active" : ""}>
            <input
              type="radio"
              name="reminder_mode"
              value="full"
              checked={reminderMode === "full"}
              onChange={() => onReminderModeChange("full")}
            />
            <span>全量生成</span>
          </label>
          <label className={reminderMode === "incremental" ? "active" : ""}>
            <input
              type="radio"
              name="reminder_mode"
              value="incremental"
              checked={reminderMode === "incremental"}
              onChange={() => onReminderModeChange("incremental")}
            />
            <span>增量更新</span>
          </label>
        </div>
        <div className="grid">
          {reminderMode === "incremental" ? (
            <UploadCard
              id="reminder-previous-file"
              name="reminder_previous_file"
              step="01"
              title="上次编辑后的开课提醒文件"
              description="以其中“学员名单”为准，保留已申诉剔除的分母，只补未发送行"
              file={previousFile}
              onChange={onPreviousFileChange}
            />
          ) : (
            <UploadCard
              id="reminder-list-file"
              name="reminder_list_file"
              step="01"
              title="开课提醒学员明细名单"
              description="授课教师、教研组、师训组长、助理主管、学员与课时"
              file={denominatorFile}
              onChange={onDenominatorFileChange}
            />
          )}
          <MultiUploadCard
            id="reminder-chat-file"
            name="reminder_chat_file"
            step="02"
            title="企微聊天质检结果"
            description={
              reminderMode === "incremental"
                ? "上传聊天记录，系统只补齐旧表未发送行，不改已有发送结果"
                : "可同时选择多个聊天质检 Excel，系统会合并清洗后统一匹配"
            }
            files={chatFiles}
            onChange={onChatFilesChange}
          />
        </div>

        {reminderMode === "full" ? (
          <div className="options reminder-options">
            <label className="switch-row">
              <span>
                <strong>包含聊天记录明细</strong>
                <small>默认不输出；需要核查命中来源时再勾选追加“清洗后聊天”Sheet</small>
              </span>
              <input
                id="include-reminder-chats"
                type="checkbox"
                checked={includeCleanChats}
                onChange={(event) => onIncludeCleanChatsChange(event.target.checked)}
              />
            </label>
          </div>
        ) : null}

        <div className="rules reminder-rules">
          {reminderMode === "incremental" ? (
            <>
              <span>增量规则</span>
              <b>读取上次结果</b>
              <b>只补未发送</b>
              <b>保留已有</b>
              <b>保留剔除分母</b>
              <b>重算汇总</b>
            </>
          ) : (
            <>
              <span>自动匹配</span>
              <b>整行去重</b>
              <b>群名含教师+学员</b>
              <b>群名唯一学员</b>
              <b>内容唯一学员</b>
              <b>异常核对</b>
            </>
          )}
        </div>

        <button id="reminder-submit-button" type="submit" disabled={processing}>
          <span>{reminderMode === "incremental" ? "生成增量更新文件" : "生成发送率公示"}</span>
          <small>
            {reminderMode === "incremental"
              ? "输出新的公示表、学员名单、维度汇总、异常明细和处理说明；已有发送结果和已剔除分母不覆盖"
              : "默认输出公示表、学员名单、维度汇总、匹配核对/异常明细和处理说明"}
          </small>
        </button>
      </form>
    </section>
  );
}
