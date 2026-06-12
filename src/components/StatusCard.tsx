import type { ProcessingStatus } from "../types/worker";

interface StatusCardProps {
  status: ProcessingStatus;
}

export function StatusCard({ status }: StatusCardProps) {
  const spinnerClass =
    status.mode === "done" ? "spinner done" : status.mode === "error" ? "spinner error" : "spinner";

  return (
    <section className="status-card" id="status-card" hidden={!status.visible}>
      <div className={spinnerClass} id="spinner" />
      <div>
        <strong id="status-title">{status.title}</strong>
        <p id="status-text">{status.message}</p>
        <div className="progress-track">
          <i id="progress-bar" style={{ width: `${Math.max(0, Math.min(100, status.progress))}%` }} />
        </div>
      </div>
    </section>
  );
}
