import { MATCHING_GUIDE } from "../data/modalContent";
import { ModalDialog } from "./ModalDialog";

interface MatchingGuideDialogProps {
  open: boolean;
  onClose: () => void;
}

function PairGrid({
  items,
  className,
}: {
  items: readonly (readonly [string, string])[];
  className: string;
}) {
  if (!items.length) return null;
  return (
    <div className={className}>
      {items.map(([title, description]) => (
        <div key={`${title}-${description}`}>
          <b>{title}</b>
          <span>{description}</span>
        </div>
      ))}
    </div>
  );
}

export function MatchingGuideDialog({ open, onClose }: MatchingGuideDialogProps) {
  const { intro, sections } = MATCHING_GUIDE;

  return (
    <ModalDialog
      id="logic-dialog"
      eyebrow="MATCHING GUIDE"
      title="打卡质检匹配规则"
      className="rules-dialog"
      open={open}
      onClose={onClose}
    >
      <div className="logic-intro">
        <strong>{intro.title}</strong>
        <p>{intro.description}</p>
        <code>{intro.formula}</code>
      </div>

      {sections.map((section, index) => (
        <section
          className={`logic-section${"result" in section && section.result ? " logic-result" : ""}`}
          key={section.title}
        >
          <span className="logic-number">{String(index + 1).padStart(2, "0")}</span>
          <div>
            <h3>{section.title}</h3>
            {"statuses" in section ? (
              <PairGrid items={section.statuses} className="logic-status-grid" />
            ) : null}
            <p>{section.description}</p>
            {"examples" in section
              ? section.examples.map((example) => (
                  <div
                    className={`logic-example${"tone" in example ? ` ${example.tone}` : ""}`}
                    key={`${section.title}-${example.title}`}
                  >
                    <b>{example.title}</b>
                    <p>{example.text}</p>
                  </div>
                ))
              : null}
            {"cases" in section ? <PairGrid items={section.cases} className="logic-cases" /> : null}
          </div>
        </section>
      ))}
    </ModalDialog>
  );
}
