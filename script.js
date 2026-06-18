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
let generatedDownloadUrl = null;

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
  if (generatedDownloadUrl) URL.revokeObjectURL(generatedDownloadUrl);
  generatedDownloadUrl = null;
  downloadLink.classList.add("hidden");
  downloadLink.removeAttribute("href");
  downloadLink.removeAttribute("download");
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
    if (!response.ok) throw new Error(await readErrorMessage(response));

    const blob = await response.blob();
    const fileName = getDownloadFileName(response) || "DTR_RECORD_relayouted.xlsx";
    if (generatedDownloadUrl) URL.revokeObjectURL(generatedDownloadUrl);
    generatedDownloadUrl = URL.createObjectURL(blob);

    downloadLink.href = generatedDownloadUrl;
    downloadLink.download = fileName;
    downloadLink.classList.remove("hidden");
    downloadLink.click();
    statusPill.textContent = "Ready to download";
    resultBox.classList.remove("error");
    resultBox.hidden = false;
    resultTitle.textContent = "Relayouted file is ready";
    resultText.textContent = `${fileName} has been generated and downloaded.`;
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

async function readErrorMessage(response) {
  const fallback = "Unable to generate the Excel file.";
  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) {
    const payload = await response.json().catch(() => null);
    return payload?.error || fallback;
  }
  const text = await response.text().catch(() => "");
  return text || fallback;
}

function getDownloadFileName(response) {
  const disposition = response.headers.get("content-disposition") || "";
  const utf8Match = disposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8Match) return decodeURIComponent(utf8Match[1]);
  const plainMatch = disposition.match(/filename="?([^"]+)"?/i);
  return plainMatch?.[1] || "";
}
