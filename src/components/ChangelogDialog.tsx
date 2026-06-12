import { CHANGELOG_ENTRIES } from "../data/modalContent";
import { ModalDialog } from "./ModalDialog";

interface ChangelogDialogProps {
  open: boolean;
  onClose: () => void;
}

export function ChangelogDialog({ open, onClose }: ChangelogDialogProps) {
  return (
    <ModalDialog
      id="changelog-dialog"
      eyebrow="CHANGELOG"
      title="版本更新记录"
      open={open}
      onClose={onClose}
    >
      {CHANGELOG_ENTRIES.map((entry) => (
        <section key={entry.version}>
          <strong>{entry.version}</strong>
          <time dateTime={entry.date.replace(" ", "T")}>{entry.date}</time>
          {"items" in entry ? (
            <ul>
              {entry.items.map((item) => (
                <li key={item}>{item}</li>
              ))}
            </ul>
          ) : (
            <p>{entry.description}</p>
          )}
        </section>
      ))}
    </ModalDialog>
  );
}
