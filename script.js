const form = document.querySelector("#uploadForm");
const input = document.querySelector("#fileInput");
const dropzone = document.querySelector("#dropzone");
const dropTitle = document.querySelector("#dropTitle");
const dropSubtitle = document.querySelector("#dropSubtitle");
const generateBtn = document.querySelector("#generateBtn");
const downloadLink = document.querySelector("#downloadLink");
const statusPill = document.querySelector("#statusPill");
const resultBox = document.querySelector("#resultBox");
const resultTitle = document.querySelector("#resultTitle");
const resultText = document.querySelector("#resultText");
const steps = {
  upload: document.querySelector("#stepUpload"),
  generate: document.querySelector("#stepGenerate"),
  download: document.querySelector("#stepDownload"),
};

let selectedFile = null;

function setStep(step) {
  Object.values(steps).forEach((item) => item.classList.remove("active"));
  steps[step].classList.add("active");
}

function setFile(file) {
  selectedFile = file;
  input.files = makeFileList(file);
  dropTitle.textContent = file.name;
  dropSubtitle.textContent = `${formatBytes(file.size)} selected`;
  generateBtn.disabled = false;
  statusPill.textContent = "Record selected";
  setStep("generate");
  downloadLink.classList.add("hidden");
  resultBox.hidden = true;
}

function makeFileList(file) {
  const transfer = new DataTransfer();
  transfer.items.add(file);
  return transfer.files;
}

function formatBytes(bytes) {
  if (!bytes) return "0 KB";
  const kb = bytes / 1024;
  return kb < 1024 ? `${kb.toFixed(1)} KB` : `${(kb / 1024).toFixed(2)} MB`;
}

input.addEventListener("change", () => {
  if (input.files?.[0]) setFile(input.files[0]);
});

["dragenter", "dragover"].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.add("dragging");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.remove("dragging");
  });
});

dropzone.addEventListener("drop", (event) => {
  const file = event.dataTransfer.files?.[0];
  if (file) setFile(file);
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!selectedFile) return;

  generateBtn.disabled = true;
  statusPill.textContent = "Generating";
  resultBox.hidden = true;
  downloadLink.classList.add("hidden");

  const body = new FormData();
  body.append("file", selectedFile);

  try {
    const response = await fetch("/api/format", { method: "POST", body });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "Unable to generate the Excel file.");

    downloadLink.href = payload.downloadUrl;
    downloadLink.download = payload.fileName;
    downloadLink.classList.remove("hidden");
    statusPill.textContent = "Ready to download";
    resultBox.classList.remove("error");
    resultBox.hidden = false;
    resultTitle.textContent = "Relayouted file is ready";
    resultText.textContent = `${payload.summary.employees} employees formatted across ${payload.summary.days} DTR days.`;
    setStep("download");
  } catch (error) {
    statusPill.textContent = "Needs attention";
    resultBox.classList.add("error");
    resultBox.hidden = false;
    resultTitle.textContent = "Could not generate file";
    resultText.textContent = error.message;
    generateBtn.disabled = false;
  }
});
