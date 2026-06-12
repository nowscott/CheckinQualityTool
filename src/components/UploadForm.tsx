import type { FormEvent } from "react";
import type { WeekLabel } from "../types/worker";
import { UploadCard } from "./UploadCard";

interface UploadFormProps {
  listFile: File | null;
  chatFile: File | null;
  weekLabel: WeekLabel;
  weekHint: string;
  useSingle: boolean;
  processing: boolean;
  onListFileChange: (file: File | null) => void;
  onChatFileChange: (file: File | null) => void;
  onWeekLabelChange: (value: WeekLabel) => void;
  onUseSingleChange: (value: boolean) => void;
  onSubmit: (event: FormEvent<HTMLFormElement>) => void;
}

const WEEK_OPTIONS: Array<{ value: WeekLabel; label: string }> = [
  { value: "auto", label: "自动识别（推荐）" },
  { value: "第一周", label: "第一周" },
  { value: "第二周", label: "第二周" },
  { value: "第三周", label: "第三周" },
  { value: "第四周", label: "第四周" },
  { value: "第五周", label: "第五周" },
];

export function UploadForm({
  listFile,
  chatFile,
  weekLabel,
  weekHint,
  useSingle,
  processing,
  onListFileChange,
  onChatFileChange,
  onWeekLabelChange,
  onUseSingleChange,
  onSubmit,
}: UploadFormProps) {
  return (
    <section className="card">
      <form id="process-form" onSubmit={onSubmit}>
        <div className="grid">
          <UploadCard
            id="list-file"
            name="list_file"
            step="01"
            title="课堂反馈质检名单"
            description="老师、学员、教师邮箱及课次信息"
            file={listFile}
            onChange={onListFileChange}
          />
          <UploadCard
            id="chat-file"
            name="chat_file"
            step="02"
            title="聊天质检明细"
            description="聊天类型、发送方、邮箱、群名及聊天内容"
            file={chatFile}
            onChange={onChatFileChange}
          />
        </div>

        <div className="options">
          <label className="week-field">
            <span className="option-title">服务周标记</span>
            <select
              id="week-label"
              name="week_label"
              value={weekLabel}
              onChange={(event) => onWeekLabelChange(event.target.value as WeekLabel)}
            >
              {WEEK_OPTIONS.map((option) => (
                <option key={option.value} value={option.value}>
                  {option.label}
                </option>
              ))}
            </select>
            <small id="week-hint">{weekHint}</small>
          </label>
          <label className="switch-row">
            <span>
              <strong>启用末字弱匹配</strong>
              <small>姓名至少两字时，后两字未命中再用末字匹配</small>
            </span>
            <input
              id="use-single"
              type="checkbox"
              checked={useSingle}
              onChange={(event) => onUseSingleChange(event.target.checked)}
            />
          </label>
        </div>

        <div className="rules">
          <span>自动清洗</span>
          <b>私聊</b>
          <b>无邮箱</b>
          <b>非员工发送</b>
          <b>引用回复</b>
          <b>姓名异常后缀</b>
          <b>内置白名单</b>
        </div>

        <button id="submit-button" type="submit" disabled={processing}>
          <span>生成质检结果</span>
          <small>输出打卡结果、匹配明细、清洗后聊天、处理说明、内置白名单</small>
        </button>
      </form>
    </section>
  );
}
