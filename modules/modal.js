import { CHANGELOG_ENTRIES, MATCHING_GUIDE } from "./modal-content.js";

function closeButton(id) {
  return `
    <button class="dialog-close" id="${id}" type="button" aria-label="关闭">
      <svg aria-hidden="true" viewBox="0 0 20 20" focusable="false">
        <path d="M4 4 16 16M16 4 4 16"></path>
      </svg>
    </button>
  `;
}

function modalHeader(eyebrow, title, closeId) {
  return `
    <div class="changelog-head">
      <div><span>${eyebrow}</span><h2>${title}</h2></div>
      ${closeButton(closeId)}
    </div>
  `;
}

function renderChangelog() {
  const entries = CHANGELOG_ENTRIES.map((entry) => `
    <section>
      <strong>${entry.version}</strong>
      <time datetime="${entry.date}">${entry.date}</time>
      ${entry.items
        ? `<ul>${entry.items.map((item) => `<li>${item}</li>`).join("")}</ul>`
        : `<p>${entry.description}</p>`}
    </section>
  `).join("");

  return `
    <dialog class="changelog-dialog" id="changelog-dialog" tabindex="-1">
      ${modalHeader("CHANGELOG", "版本更新记录", "changelog-close")}
      ${entries}
    </dialog>
  `;
}

function renderExamples(examples = []) {
  return examples.map((example) => `
    <div class="logic-example ${example.tone || ""}">
      <b>${example.title}</b>
      <p>${example.text}</p>
    </div>
  `).join("");
}

function renderPairs(items = [], className) {
  if (!items.length) return "";
  return `
    <div class="${className}">
      ${items.map(([title, description]) => `
        <div><b>${title}</b><span>${description}</span></div>
      `).join("")}
    </div>
  `;
}

function renderMatchingGuide() {
  const { intro, sections } = MATCHING_GUIDE;
  const content = sections.map((section, index) => `
    <section class="logic-section ${section.result ? "logic-result" : ""}">
      <span class="logic-number">${String(index + 1).padStart(2, "0")}</span>
      <div>
        <h3>${section.title}</h3>
        ${section.result ? renderPairs(section.statuses, "logic-status-grid") : ""}
        <p>${section.description}</p>
        ${renderExamples(section.examples)}
        ${renderPairs(section.cases, "logic-cases")}
      </div>
    </section>
  `).join("");

  return `
    <dialog class="changelog-dialog rules-dialog" id="logic-dialog" tabindex="-1">
      ${modalHeader("MATCHING GUIDE", "打卡质检匹配规则", "logic-close")}
      <div class="logic-intro">
        <strong>${intro.title}</strong>
        <p>${intro.description}</p>
        <code>${intro.formula}</code>
      </div>
      ${content}
    </dialog>
  `;
}

function createScrollLock() {
  let scrollY = 0;

  return {
    lock() {
      scrollY = window.scrollY;
      const scrollbarWidth = window.innerWidth - document.documentElement.clientWidth;
      document.body.style.top = `-${scrollY}px`;
      document.body.style.paddingRight = `${scrollbarWidth}px`;
      document.documentElement.classList.add("dialog-open");
      document.body.classList.add("dialog-open");
    },
    unlock() {
      if (!document.body.classList.contains("dialog-open")) return;
      document.documentElement.classList.remove("dialog-open");
      document.body.classList.remove("dialog-open");
      document.body.style.top = "";
      document.body.style.paddingRight = "";
      window.scrollTo(0, scrollY);
    },
  };
}

function bindDialog(trigger, dialog, close, scrollLock) {
  trigger.addEventListener("click", () => {
    scrollLock.lock();
    dialog.showModal();
    dialog.focus({ preventScroll: true });
  });
  close.addEventListener("click", () => dialog.close());
  dialog.addEventListener("click", (event) => {
    if (event.target === dialog) dialog.close();
  });
  dialog.addEventListener("close", scrollLock.unlock);
}

export function mountModalComponents(root) {
  root.innerHTML = `${renderChangelog()}${renderMatchingGuide()}`;

  const scrollLock = createScrollLock();
  bindDialog(
    document.querySelector("#version-button"),
    root.querySelector("#changelog-dialog"),
    root.querySelector("#changelog-close"),
    scrollLock,
  );
  bindDialog(
    document.querySelector("#logic-button"),
    root.querySelector("#logic-dialog"),
    root.querySelector("#logic-close"),
    scrollLock,
  );
}
