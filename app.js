const form = document.querySelector("#process-form");
const listFile = document.querySelector("#list-file");
const chatFile = document.querySelector("#chat-file");
const listName = document.querySelector("#list-name");
const chatName = document.querySelector("#chat-name");
const button = document.querySelector("#submit-button");
const statusCard = document.querySelector("#status-card");
const statusTitle = document.querySelector("#status-title");
const statusText = document.querySelector("#status-text");
const spinner = document.querySelector("#spinner");
const progressBar = document.querySelector("#progress-bar");
const weekSelect = document.querySelector("#week-label");
const weekHint = document.querySelector("#week-hint");
const versionButton = document.querySelector("#version-button");
const changelogDialog = document.querySelector("#changelog-dialog");
const changelogClose = document.querySelector("#changelog-close");

let worker;
let whitelistCsvPromise;
let dialogScrollY = 0;

function openChangelog() {
  dialogScrollY = window.scrollY;
  document.body.style.top = `-${dialogScrollY}px`;
  document.documentElement.classList.add("dialog-open");
  document.body.classList.add("dialog-open");
  changelogDialog.showModal();
  changelogDialog.focus({ preventScroll: true });
}

function unlockPageScroll() {
  if (!document.body.classList.contains("dialog-open")) return;
  document.documentElement.classList.remove("dialog-open");
  document.body.classList.remove("dialog-open");
  document.body.style.top = "";
  window.scrollTo(0, dialogScrollY);
}

function loadWhitelistCsv() {
  whitelistCsvPromise ||= fetch("./data/whitelist.csv", { cache: "no-store" }).then((response) => {
    if (!response.ok) throw new Error("内置白名单读取失败，请刷新页面后重试。");
    return response.text();
  });
  return whitelistCsvPromise;
}

function weekOfMonthFromDate(date) {
  const year = date.getFullYear();
  const month = date.getMonth();
  const firstDay = new Date(year, month, 1);
  const mondayOffset = (firstDay.getDay() + 6) % 7;
  return Math.min(5, Math.floor((date.getDate() + mondayOffset - 1) / 7) + 1);
}

function chineseWeek(number) {
  return ["", "第一周", "第二周", "第三周", "第四周", "第五周"][number] || "";
}

function inferWeekFromFilename(filename) {
  const matches = [...filename.matchAll(/(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})/g)];
  if (!matches.length) return "";
  const dates = matches.map((match) => new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
  const middleDate = dates[Math.floor(dates.length / 2)];
  return chineseWeek(weekOfMonthFromDate(middleDate));
}

function updateWeekHint() {
  if (weekSelect.value !== "auto") {
    weekHint.textContent = `已手动指定为${weekSelect.value}`;
    return;
  }
  const inferred = listFile.files[0] ? inferWeekFromFilename(listFile.files[0].name) : "";
  weekHint.textContent = inferred
    ? `根据文件日期预计为${inferred}，生成时会再用课次日期校验`
    : "选择名单后，将根据上课日期自动判断";
}

function bindFileName(input, output, onChange) {
  const label = input.closest(".upload");
  const nameText = output.querySelector(".file-name-text");
  input.addEventListener("change", () => {
    const file = input.files[0];
    const displayName = file ? `${file.name} · ${(file.size / 1024 / 1024).toFixed(1)} MB` : "选择 Excel 文件";
    nameText.textContent = displayName;
    nameText.title = file?.name || "";
    label.classList.toggle("has-file", Boolean(file));
    onChange?.();
  });
}

function setStatus(title, message, progress = 0, mode = "working") {
  statusCard.hidden = false;
  statusTitle.textContent = title;
  statusText.textContent = message;
  progressBar.style.width = `${Math.max(0, Math.min(100, progress))}%`;
  spinner.className = mode === "done" ? "spinner done" : mode === "error" ? "spinner error" : "spinner";
}

function downloadResult(buffer, filename) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

bindFileName(listFile, listName, updateWeekHint);
bindFileName(chatFile, chatName);
weekSelect.addEventListener("change", updateWeekHint);
versionButton.addEventListener("click", openChangelog);
changelogClose.addEventListener("click", () => changelogDialog.close());
changelogDialog.addEventListener("click", (event) => {
  if (event.target === changelogDialog) changelogDialog.close();
});
changelogDialog.addEventListener("close", unlockPageScroll);

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!listFile.files[0] || !chatFile.files[0]) return;

  worker?.terminate();
  worker = new Worker("./worker.js");
  button.disabled = true;
  setStatus("正在启动本地处理引擎", "所有文件只在当前浏览器中处理，不会上传。", 2);
  let whitelistCsv;
  try {
    whitelistCsv = await loadWhitelistCsv();
  } catch (error) {
    setStatus("处理失败", error.message, 100, "error");
    button.disabled = false;
    worker.terminate();
    worker = null;
    return;
  }

  worker.onmessage = ({ data }) => {
    if (data.type === "progress") {
      setStatus(data.title, data.message, data.progress);
      return;
    }
    if (data.type === "complete") {
      downloadResult(data.buffer, data.filename);
      setStatus(
        "处理完成，结果已下载",
        `质检 ${data.summary.targets.toLocaleString()} 人：已发送 ${data.summary.sent.toLocaleString()}，未发送 ${data.summary.unsent.toLocaleString()}，免检 ${data.summary.exempt.toLocaleString()}；清洗后聊天 ${data.summary.cleanChats.toLocaleString()} 条。`,
        100,
        "done",
      );
      button.disabled = false;
      worker.terminate();
      worker = null;
      return;
    }
    if (data.type === "error") {
      setStatus("处理失败", data.message, 100, "error");
      button.disabled = false;
      worker.terminate();
      worker = null;
    }
  };

  worker.onerror = (event) => {
    setStatus("处理失败", event.message || "浏览器工作线程发生错误。", 100, "error");
    button.disabled = false;
    worker?.terminate();
    worker = null;
  };

  worker.postMessage({
    type: "process",
    listFile: listFile.files[0],
    chatFile: chatFile.files[0],
    weekLabel: weekSelect.value,
    useSingle: document.querySelector("#use-single").checked,
    whitelistCsv,
  });
});
