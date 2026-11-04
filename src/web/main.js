import * as docx from "docx";
import JSZip from "jszip";
import { createIcons, Download, FileText, Trash2, Upload } from "lucide";
import { buildDocument } from "../converter.js";

createIcons({ icons: { Download, FileText, Trash2, Upload } });

const dropZone = document.querySelector("#dropZone");
const fileInput = document.querySelector("#fileInput");
const fileList = document.querySelector("#fileList");
const downloadAllBtn = document.querySelector("#downloadAllBtn");
const clearBtn = document.querySelector("#clearBtn");

const MARKDOWN_EXTENSIONS = new Set([".md", ".markdown"]);
const items = new Map();

function isMarkdownFile(file) {
  const dot = file.name.lastIndexOf(".");
  if (dot === -1) {
    return false;
  }
  return MARKDOWN_EXTENSIONS.has(file.name.slice(dot).toLowerCase());
}

function outputName(fileName) {
  const dot = fileName.lastIndexOf(".");
  const base = dot === -1 ? fileName : fileName.slice(0, dot);
  return `${base}.docx`;
}

function humanSize(bytes) {
  if (bytes < 1024) {
    return `${bytes} B`;
  }
  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)} KB`;
  }
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function renderItem(id, data) {
  const existing = document.querySelector(`[data-file-id="${id}"]`);
  const node = existing || document.createElement("article");
  node.className = "file-item";
  node.dataset.fileId = id;

  const statusClass =
    data.status === "error"
      ? "error"
      : data.status === "done"
        ? "success"
        : "";
  const statusText =
    data.status === "error"
      ? data.error
      : data.status === "done"
        ? `已转换 · ${humanSize(data.blob.size)}`
        : "正在转换…";

  node.innerHTML = `
    <i data-lucide="file-text" class="file-icon"></i>
    <div class="file-meta">
      <div class="file-name" title="${escapeHtml(data.name)}">${escapeHtml(data.name)}</div>
      <div class="file-status ${statusClass}">${escapeHtml(statusText)}</div>
    </div>
    <div class="file-actions">
      ${
        data.status === "done"
          ? `<button class="btn btn-primary" data-action="download"><i data-lucide="download"></i><span>下载</span></button>`
          : ""
      }
    </div>
  `;

  if (!existing) {
    fileList.appendChild(node);
  }
  createIcons({ icons: { Download, FileText, Trash2, Upload } });
  return node;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function renderEmptyState() {
  const empty = document.createElement("p");
  empty.className = "empty-state";
  empty.textContent = "还没有文件，拖入或选择 Markdown 文件开始。";
  fileList.appendChild(empty);
}

function updateControls() {
  const ready = [...items.values()].filter((item) => item.status === "done");
  downloadAllBtn.disabled = ready.length === 0;
  clearBtn.disabled = items.size === 0;
}

async function convertFile(file) {
  const id = `${file.name}-${file.lastModified}-${file.size}`;
  items.set(id, { id, file, name: file.name, status: "loading" });
  renderItem(id, items.get(id));
  updateControls();

  if (!isMarkdownFile(file)) {
    items.set(id, {
      ...items.get(id),
      status: "error",
      error: "格式不支持，仅支持 .md 和 .markdown",
    });
    renderItem(id, items.get(id));
    updateControls();
    return;
  }

  try {
    const markdown = await file.text();
    const blob = await docx.Packer.toBlob(buildDocument(markdown));
    items.set(id, { ...items.get(id), status: "done", blob });
  } catch (error) {
    items.set(id, {
      ...items.get(id),
      status: "error",
      error: `转换失败：${error.message}`,
    });
  }
  renderItem(id, items.get(id));
  updateControls();
}

async function addFiles(files) {
  fileList.querySelectorAll(".empty-state").forEach((node) => node.remove());
  for (const file of files) {
    await convertFile(file);
  }
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

fileList.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-action]");
  if (!button) {
    return;
  }
  const itemNode = button.closest("[data-file-id]");
  const item = items.get(itemNode.dataset.fileId);
  if (button.dataset.action === "download" && item?.blob) {
    downloadBlob(item.blob, outputName(item.name));
  }
});

downloadAllBtn.addEventListener("click", async () => {
  const ready = [...items.values()].filter((item) => item.status === "done");
  if (ready.length === 0) {
    return;
  }
  const zip = new JSZip();
  for (const item of ready) {
    zip.file(outputName(item.name), item.blob);
  }
  const blob = await zip.generateAsync({ type: "blob" });
  downloadBlob(blob, "markdown-word.zip");
});

clearBtn.addEventListener("click", () => {
  items.clear();
  fileList.innerHTML = "";
  renderEmptyState();
  updateControls();
});

dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    fileInput.click();
  }
});

fileInput.addEventListener("change", () => {
  addFiles([...fileInput.files]);
  fileInput.value = "";
});

["dragenter", "dragover"].forEach((type) => {
  dropZone.addEventListener(type, (event) => {
    event.preventDefault();
    dropZone.classList.add("dragging");
  });
});

["dragleave", "drop"].forEach((type) => {
  dropZone.addEventListener(type, (event) => {
    event.preventDefault();
    dropZone.classList.remove("dragging");
  });
});

dropZone.addEventListener("drop", (event) => {
  addFiles([...event.dataTransfer.files]);
});

renderEmptyState();
updateControls();
