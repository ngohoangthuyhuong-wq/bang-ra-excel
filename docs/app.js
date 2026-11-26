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

const MAX_PAGES = 20;
let chosen = [];
let workbook = null;
let workerPromise = null;

pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

function setStatus(text, isError = false) {
  statusEl.textContent = text;
  statusEl.classList.toggle("is-error", isError);
}

function renderFiles() {
  list.replaceChildren();
  list.hidden = chosen.length === 0;
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
drop.addEventListener("drop", (event) => addFiles(event.dataTransfer.files));

function fixDigits(value) {
  return value.replace(/[0-9oO]{2,}/g, (token) => {
    if (!/\d/.test(token)) return token;
    return token.replace(/o/g, "0").replace(/O/g, "0");
  });
}

function sheetName(raw, used) {
  let name = raw.replace(/[\\/?*[\]:]/g, " ").replace(/\s+/g, " ").trim() || "Bang";
  name = name.slice(0, 31);
  const base = name;
  let n = 2;
  while (used.has(name.toLowerCase())) {
    const suffix = ` ${n}`;
    name = base.slice(0, 31 - suffix.length) + suffix;
    n += 1;
  }
  used.add(name.toLowerCase());
  return name;
}

function cluster(words) {
  const usable = words.filter((word) => word.text && word.text.trim() && word.confidence >= 35);
  usable.sort((a, b) => (a.bbox.y0 + a.bbox.y1) / 2 - (b.bbox.y0 + b.bbox.y1) / 2);
  if (!usable.length) return [];
  const heights = usable.map((word) => word.bbox.y1 - word.bbox.y0).sort((a, b) => a - b);
  const medH = heights[Math.floor(heights.length / 2)] || 12;
  const rows = [];
  for (const word of usable) {
    const cy = (word.bbox.y0 + word.bbox.y1) / 2;
    const last = rows[rows.length - 1];
    if (!last || Math.abs(cy - last.cy) > medH * 0.65) {
      rows.push({ cy, words: [word] });
    } else {
      last.words.push(word);
      last.cy = last.words.reduce((sum, item) => sum + (item.bbox.y0 + item.bbox.y1) / 2, 0) / last.words.length;
    }
  }
  const widths = rows.flatMap((row) => row.words.map((word) => word.bbox.x1 - word.bbox.x0)).sort((a, b) => a - b);
  const gap = Math.max(16, (widths[Math.floor(widths.length / 2)] || 20) * 0.45);
  const edges = [];
  for (const word of rows.flatMap((row) => row.words).sort((a, b) => a.bbox.x0 - b.bbox.x0)) {
    if (!edges.length || word.bbox.x0 - edges[edges.length - 1] > gap) edges.push(word.bbox.x0);
  }
  return rows.map((row) => {
    const cells = edges.length > 1 ? Array(edges.length).fill("") : [""];
    for (const word of row.words.sort((a, b) => a.bbox.x0 - b.bbox.x0)) {
      let idx = 0;
      if (edges.length > 1) {
        let best = Infinity;
        edges.forEach((edge, index) => {
          const distance = Math.abs(word.bbox.x0 - edge);
          if (distance < best) {
            best = distance;
            idx = index;
          }
        });
      }
      cells[idx] = fixDigits(`${cells[idx]} ${word.text}`.trim());
    }
    return cells;
  }).filter((row) => row.some(Boolean));
}

async function getWorker(onProgress) {
  if (!workerPromise) {
    workerPromise = Tesseract.createWorker(["vie", "eng"], 1, {
      logger: (message) => {
        if (message.status === "loading language traineddata") onProgress("Đang tải bộ đọc tiếng Việt…");
      },
    });
  }
  return workerPromise;
}

function canvasFromBitmap(bitmap) {
  const longEdge = Math.max(bitmap.width, bitmap.height);
  let scale = 1;
  if (longEdge < 1800) scale = Math.min(2.5, 1800 / longEdge);
  if (longEdge * scale > 2600) scale = 2600 / longEdge;
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(bitmap.width * scale));
  canvas.height = Math.max(1, Math.round(bitmap.height * scale));
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  return canvas;
}

async function pagesOf(file) {
  if (file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf")) {
    const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
    const count = Math.min(pdf.numPages, MAX_PAGES);
    const pages = [];
    for (let number = 1; number <= count; number += 1) {
      const page = await pdf.getPage(number);
      const viewport = page.getViewport({ scale: 2 });
      const canvas = document.createElement("canvas");
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
      pages.push(canvas);
    }
    return { pages, truncated: pdf.numPages > MAX_PAGES, total: pdf.numPages };
  }
  const bitmap = await createImageBitmap(file);
  return { pages: [canvasFromBitmap(bitmap)], truncated: false, total: 1 };
}

function showSheet(sheets, index) {
  tabs.querySelectorAll("button").forEach((button, i) => {
    button.setAttribute("aria-selected", i === index ? "true" : "false");
  });
  const table = document.createElement("table");
  sheets[index].rows.slice(0, 100).forEach((row) => {
    const tr = document.createElement("tr");
    row.forEach((cell) => {
      const td = document.createElement("td");
      td.textContent = cell;
      tr.append(td);
    });
    table.append(tr);
  });
  grid.replaceChildren(table);
}

function showResult(sheets) {
  result.hidden = false;
  download.hidden = false;
  tabs.replaceChildren();
  sheets.forEach((sheet, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = `${sheet.name} (${sheet.rows.length} dòng)`;
    button.addEventListener("click", () => showSheet(sheets, index));
    tabs.append(button);
  });
  if (sheets.length) showSheet(sheets, 0);
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!chosen.length) {
    setStatus("Chọn hoặc thả ít nhất một file PDF hoặc ảnh.", true);
    return;
  }
  go.disabled = true;
  download.hidden = true;
  result.hidden = true;
  ledger.classList.add("is-reading");
  meter.hidden = false;
  meterBar.style.width = "8%";
  const book = XLSX.utils.book_new();
  const used = new Set();
  const sheets = [];
  const notes = [];
  try {
    const worker = await getWorker(setStatus);
    let done = 0;
    const jobs = chosen.length;
    for (const file of chosen) {
      setStatus(`Đang mở ${file.name}`);
      const { pages, truncated, total } = await pagesOf(file);
      if (truncated) notes.push(`${file.name} có ${total} trang, link này đọc ${MAX_PAGES} trang đầu.`);
      for (let index = 0; index < pages.length; index += 1) {
        setStatus(`Đang đọc ${file.name}, trang ${index + 1}/${pages.length}`);
        const recognized = await worker.recognize(pages[index]);
        const rows = cluster(recognized.data.words || []);
        if (rows.length) {
          const name = sheetName(pages.length > 1 || chosen.length > 1 ? `${file.name} ${index + 1}` : file.name, used);
          const ws = XLSX.utils.aoa_to_sheet(rows);
          XLSX.utils.book_append_sheet(book, ws, name);
          sheets.push({ name, rows });
        }
        done += 1;
        meterBar.style.width = `${Math.min(95, Math.round((done / Math.max(jobs, 1)) * 90))}%`;
      }
    }
    if (!sheets.length) throw new Error("Không thấy chữ. Thử ảnh sáng hơn, thẳng hơn.");
    workbook = book;
    meterBar.style.width = "100%";
    noteEl.hidden = notes.length === 0;
    noteEl.textContent = notes.join(" ");
    setStatus(`Đã đọc xong. ${sheets.length} bảng.`);
    showResult(sheets);
  } catch (error) {
    setStatus(error.message || "Không đọc được file.", true);
  } finally {
    ledger.classList.remove("is-reading");
    go.disabled = false;
  }
});

download.addEventListener("click", () => {
  if (workbook) XLSX.writeFile(workbook, "bang-ra-excel.xlsx");
});
