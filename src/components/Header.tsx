interface HeaderProps {
  onOpenGuide: () => void;
  onOpenChangelog: () => void;
}

export function Header({ onOpenGuide, onOpenChangelog }: HeaderProps) {
  return (
    <header>
      <div className="brand-line">
        <img
          className="brand-icon"
          src="/assets/checkin-icon.png"
          alt=""
          width="42"
          height="42"
        />
        <p className="eyebrow">LOCAL QUALITY CHECK</p>
      </div>
      <h1>打卡质检数据生成</h1>
      <p className="subtitle">
        上传课堂反馈名单与聊天导出数据，在浏览器本地完成清洗、匹配并生成可追溯的多 Sheet
        Excel。文件不会上传服务器。{" "}
        <button className="logic-button" id="logic-button" type="button" onClick={onOpenGuide}>
          查看匹配规则
        </button>
      </p>
      <button
        className="version-button"
        id="version-button"
        type="button"
        onClick={onOpenChangelog}
      >
        v2.0.0
      </button>
    </header>
  );
}
