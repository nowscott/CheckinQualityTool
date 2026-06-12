import type { Theme } from "../hooks/useTheme";

interface HeaderProps {
  theme: Theme;
  usesSystemTheme: boolean;
  onToggleTheme: () => void;
  onOpenGuide: () => void;
  onOpenChangelog: () => void;
}

export function Header({
  theme,
  usesSystemTheme,
  onToggleTheme,
  onOpenGuide,
  onOpenChangelog,
}: HeaderProps) {
  const nextThemeLabel = theme === "dark" ? "浅色" : "深色";
  const currentThemeLabel = theme === "dark" ? "深色" : "浅色";

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
      <div className="header-controls">
        <button
          className="theme-button"
          type="button"
          aria-label={`当前${currentThemeLabel}模式，切换为${nextThemeLabel}模式`}
          title={`${usesSystemTheme ? "当前跟随设备" : "当前手动设置"} · 切换为${nextThemeLabel}模式`}
          onClick={onToggleTheme}
        >
          <svg aria-hidden="true" viewBox="0 0 20 20" focusable="false">
            {theme === "dark" ? (
              <path d="M14.8 13.1A6 6 0 0 1 6.9 5.2 6.2 6.2 0 1 0 14.8 13.1Z" />
            ) : (
              <>
                <circle cx="10" cy="10" r="3.1" />
                <path d="M10 1.5V4M10 16v2.5M1.5 10H4M16 10h2.5M4 4l1.8 1.8M14.2 14.2 16 16M16 4l-1.8 1.8M5.8 14.2 4 16" />
              </>
            )}
          </svg>
          <span>{theme === "dark" ? "DARK" : "LIGHT"}</span>
        </button>
        <button
          className="version-button"
          id="version-button"
          type="button"
          onClick={onOpenChangelog}
        >
          v2.3.0
        </button>
      </div>
    </header>
  );
}
