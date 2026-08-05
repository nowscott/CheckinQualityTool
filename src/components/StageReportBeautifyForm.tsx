import type { FormEvent } from "react";
import { UploadCard } from "./UploadCard";

interface StageReportBeautifyFormProps {
  sourceFile: File | null;
  processing: boolean;
  onSourceFileChange: (file: File | null) => void;
  onSubmit: (event: FormEvent<HTMLFormElement>) => void;
}

export function StageReportBeautifyForm({
  sourceFile,
  processing,
  onSourceFileChange,
  onSubmit,
}: StageReportBeautifyFormProps) {
  return (
    <section className="card">
      <form id="stage-report-beautify-form" onSubmit={onSubmit}>
        <div className="grid">
          <UploadCard
            id="stage-report-beautify-source-file"
            name="stage_report_beautify_source_file"
            step="01"
            title="窗口期＋非窗口期阶段性报告表单"
            description="上传原始明细表单；自动整理明细、教师、师训组长、助理主管、教研组及申诉数据"
            file={sourceFile}
            onChange={onSourceFileChange}
          />
          <div className="beautify-note">
            <span className="step">OUTPUT</span>
            <strong>统一公示风格</strong>
            <p>保留原始业务字段和发送口径，重建助理主管、教研组项目层级，并统一标题、列宽、数据条和状态颜色。</p>
            <div className="rules">
              <b>不改原始结论</b>
              <b>窗口期数据保留</b>
              <b>本地处理</b>
            </div>
          </div>
        </div>

        <button id="stage-report-beautify-submit-button" type="submit" disabled={processing}>
          <span>生成阶段性报告公示版</span>
          <small>输出明细与四个管理维度，另保留窗口期明细和申诉情况</small>
        </button>
      </form>
    </section>
  );
}
