import type { ChangeEvent } from "react";

interface MultiUploadCardProps {
  id: string;
  name: string;
  step: string;
  title: string;
  description: string;
  files: File[];
  onChange: (files: File[]) => void;
}

export function MultiUploadCard({
  id,
  name,
  step,
  title,
  description,
  files,
  onChange,
}: MultiUploadCardProps) {
  function handleChange(event: ChangeEvent<HTMLInputElement>) {
    onChange(Array.from(event.target.files || []));
  }

  const totalSize = files.reduce((sum, file) => sum + file.size, 0);
  const displayName = files.length
    ? `${files.length} 个 Excel 文件 · ${(totalSize / 1024 / 1024).toFixed(1)} MB`
    : "选择 Excel 文件";
  const fileTitle = files.map((file) => file.name).join("\n");

  return (
    <label className={`upload${files.length ? " has-file" : ""}`}>
      <span className="upload-top">
        <span className="step">{step}</span>
        <span className="file-type">XLSX</span>
      </span>
      <span className="upload-copy">
        <strong>{title}</strong>
        <small>{description}</small>
      </span>
      <input id={id} name={name} type="file" accept=".xlsx" multiple required onChange={handleChange} />
      <span className="file-name" id={`${id.replace("-file", "")}-name`}>
        <span className="file-name-text" title={fileTitle}>
          {displayName}
        </span>
        <span className="file-action">浏览</span>
      </span>
    </label>
  );
}

