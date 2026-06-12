import type { ChangeEvent } from "react";

interface UploadCardProps {
  id: string;
  name: string;
  step: string;
  title: string;
  description: string;
  file: File | null;
  onChange: (file: File | null) => void;
}

export function UploadCard({
  id,
  name,
  step,
  title,
  description,
  file,
  onChange,
}: UploadCardProps) {
  function handleChange(event: ChangeEvent<HTMLInputElement>) {
    onChange(event.target.files?.[0] || null);
  }

  const displayName = file
    ? `${file.name} · ${(file.size / 1024 / 1024).toFixed(1)} MB`
    : "选择 Excel 文件";

  return (
    <label className={`upload${file ? " has-file" : ""}`}>
      <span className="upload-top">
        <span className="step">{step}</span>
        <span className="file-type">XLSX</span>
      </span>
      <span className="upload-copy">
        <strong>{title}</strong>
        <small>{description}</small>
      </span>
      <input id={id} name={name} type="file" accept=".xlsx" required onChange={handleChange} />
      <span className="file-name" id={`${id.replace("-file", "")}-name`}>
        <span className="file-name-text" title={file?.name || ""}>
          {displayName}
        </span>
        <span className="file-action">浏览</span>
      </span>
    </label>
  );
}
