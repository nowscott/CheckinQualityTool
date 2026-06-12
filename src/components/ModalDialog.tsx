import { useEffect, useRef, type MouseEvent, type ReactNode } from "react";

interface ModalDialogProps {
  id: string;
  eyebrow: string;
  title: string;
  open: boolean;
  className?: string;
  children: ReactNode;
  onClose: () => void;
}

export function ModalDialog({
  id,
  eyebrow,
  title,
  open,
  className = "",
  children,
  onClose,
}: ModalDialogProps) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const scrollYRef = useRef(0);
  const closeId = `${id.replace("-dialog", "")}-close`;

  useEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog || !open) return;

    scrollYRef.current = window.scrollY;
    const scrollbarWidth = window.innerWidth - document.documentElement.clientWidth;
    document.body.style.top = `-${scrollYRef.current}px`;
    document.body.style.paddingRight = `${scrollbarWidth}px`;
    document.documentElement.classList.add("dialog-open");
    document.body.classList.add("dialog-open");

    if (!dialog.open) dialog.showModal();
    dialog.focus({ preventScroll: true });

    return () => {
      document.documentElement.classList.remove("dialog-open");
      document.body.classList.remove("dialog-open");
      document.body.style.top = "";
      document.body.style.paddingRight = "";
      window.scrollTo(0, scrollYRef.current);
    };
  }, [open]);

  function handleBackdropClick(event: MouseEvent<HTMLDialogElement>) {
    if (event.target === event.currentTarget) event.currentTarget.close();
  }

  return (
    <dialog
      ref={dialogRef}
      className={`changelog-dialog${className ? ` ${className}` : ""}`}
      id={id}
      tabIndex={-1}
      onClick={handleBackdropClick}
      onClose={onClose}
    >
      <div className="changelog-head">
        <div>
          <span>{eyebrow}</span>
          <h2>{title}</h2>
        </div>
        <button
          className="dialog-close"
          id={closeId}
          type="button"
          aria-label="关闭"
          onClick={() => dialogRef.current?.close()}
        >
          <svg aria-hidden="true" viewBox="0 0 20 20" focusable="false">
            <path d="M4 4 16 16M16 4 4 16" />
          </svg>
        </button>
      </div>
      {children}
    </dialog>
  );
}
