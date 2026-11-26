const form = document.querySelector("#form");
const input = document.querySelector("#files");
const drop = document.querySelector("#drop");
const list = document.querySelector("#file-list");
const ledger = document.querySelector("#ledger");
const statusEl = document.querySelector("#status");
const meter = document.querySelector("#meter");
const meterBar = document.querySelector("#meter-bar");
const noteEl = document.querySelector("#note");
const result = document.querySelector("#result");
const tabs = document.querySelector("#tabs");
const grid = document.querySelector("#grid");
const go = document.querySelector("#go");
const download = document.querySelector("#download");

let chosen = [];

function setStatus(text, isError = false) {
  statusEl.textContent = text;
  statusEl.classList.toggle("is-error", isError);
}

function renderFiles() {
  list.innerHTML = "";
  if (!chosen.length) {
    list.hidden = true;
    return;
  }
  list.hidden = false;
  chosen.forEach((file, index) => {
    const li = document.createElement("li");
    const name = document.createElement("span");
    name.textContent = file.name;
    const remove = document.createElement("button");
    remove.type = "button";
    remove.textContent = "Bỏ";
    remove.addEventListener("click", () => {
      chosen.splice(index, 1);
      renderFiles();
    });
    li.append(name, remove);
    list.append(li);
  });
}

function addFiles(fileList) {
  for (const file of fileList) chosen.push(file);
  renderFiles();
}

input.addEventListener("change", () => {
  addFiles(input.files);
  input.value = "";
});

["dragenter", "dragover"].forEach((name) => {
  drop.addEventListener(name, (event) => {
    event.preventDefault();
    drop.classList.add("is-over");
  });
});
["dragleave", "drop"].forEach((name) => {
  drop.addEventListener(name, (event) => {
    event.preventDefault();
    drop.classList.remove("is-over");
  });
});
drop.addEventListener("drop", (event) => {
  addFiles(event.dataTransfer.files);
});

async function pollReady() {
  try {
    const res = await fetch("/api/ready");
    const data = await res.json();
    if (!ledger.classList.contains("is-reading")) setStatus(data.message);
    if (!data.ready) setTimeout(pollReady, 1500);
  } catch {
    setStatus("Không kết nối được app. Hãy chạy lại file CHAY.bat.", true);
  }
}
pollReady();

function showSheet(sheets, index) {
  tabs.querySelectorAll("button").forEach((button, i) => {
    button.setAttribute("aria-selected", i === index ? "true" : "false");
  });
  const sheet = sheets[index];
  grid.innerHTML = sheet.html;
  if (sheet.preview_limited) {
    const more = document.createElement("p");
    more.className = "note";
    more.textContent = "Xem trước 100 dòng đầu. File Excel có đủ phần còn lại.";
    grid.append(more);
  }
}

function showResult(job, jobId) {
  const sheets = job.sheets || [];
  result.hidden = false;
  download.hidden = false;
  download.href = `/api/jobs/${jobId}/xlsx`;
  tabs.innerHTML = "";
  sheets.forEach((sheet, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = `${sheet.name} (${sheet.rows} dòng)`;
    button.setAttribute("role", "tab");
    button.addEventListener("click", () => showSheet(sheets, index));
    tabs.append(button);
  });
  if (sheets.length) showSheet(sheets, 0);
  const bits = [];
  if (job.tables) bits.push(`${job.tables} bảng`);
  if (job.pages) bits.push(`${job.pages} trang`);
  setStatus(bits.length ? `Đã đọc xong. ${bits.join(", ")}.` : "Đã đọc xong.");
  noteEl.hidden = !job.note;
  noteEl.textContent = job.note || "";
}

async function watch(jobId) {
  const res = await fetch(`/api/jobs/${jobId}`);
  const job = await res.json();
  meter.hidden = false;
  meterBar.style.width = `${job.progress || 0}%`;
  if (job.status === "done") {
    ledger.classList.remove("is-reading");
    go.disabled = false;
    showResult(job, jobId);
    return;
  }
  if (job.status === "error") {
    ledger.classList.remove("is-reading");
    go.disabled = false;
    setStatus(job.message || "Không đọc được file.", true);
    return;
  }
  setStatus(job.message || "Đang đọc…");
  setTimeout(() => watch(jobId), 800);
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!chosen.length) {
    setStatus("Chọn hoặc thả ít nhất một file PDF hoặc ảnh.", true);
    return;
  }
  const body = new FormData();
  chosen.forEach((file) => body.append("files", file));
  body.append("borderless", document.querySelector("#borderless").checked ? "1" : "0");
  body.append("keep_outside", document.querySelector("#outside").checked ? "1" : "0");
  go.disabled = true;
  download.hidden = true;
  result.hidden = true;
  noteEl.hidden = true;
  ledger.classList.add("is-reading");
  meter.hidden = false;
  meterBar.style.width = "6%";
  setStatus("Đang gửi file vào bộ đọc…");
  try {
    const res = await fetch("/api/jobs", { method: "POST", body });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Không nhận file.");
    watch(data.id);
  } catch (error) {
    ledger.classList.remove("is-reading");
    go.disabled = false;
    setStatus(error.message, true);
  }
});
