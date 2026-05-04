const form = document.getElementById("compare-form");
const statusPanel = document.getElementById("status-panel");
const submitButton = document.getElementById("submit-button");
const statusTitle = document.getElementById("status-title");
const statusBadge = document.getElementById("status-badge");
const statusStep = document.getElementById("status-step");
const statusMeta = document.getElementById("status-meta");
const meterBar = document.getElementById("meter-bar");
const downloadLink = document.getElementById("download-link");
const errorCopy = document.getElementById("error-copy");
const fontPanel = document.getElementById("font-panel");
const fontSummary = document.getElementById("font-summary");
const fontUniqueSummary = document.getElementById("font-unique-summary");
const fontUniqueDetails = document.getElementById("font-unique-details");
const fontUniqueList = document.getElementById("font-unique-list");
const fontGrid = document.getElementById("font-grid");

const uploadInput = document.getElementById("upload-input");
const pdfInput = document.getElementById("pdf-input");
const pptxInput = document.getElementById("pptx-input");
const pdfName = document.getElementById("pdf-name");
const pptxName = document.getElementById("pptx-name");
const submitStateNote = document.getElementById("submit-state-note");
const dropzone = document.querySelector(".dropzone");

let pollTimer = null;
let isJobActive = false;
let activeSubmissionFingerprint = null;
let lastCompletedSubmission = null;
const lastCompletedStorageKey = "pdf-to-pptx-last-completed-pair";
const apiOrigin =
  window.location.protocol === "http:" || window.location.protocol === "https:"
    ? window.location.origin
    : null;

const isPdfFile = (file) => file.name?.toLowerCase().endsWith(".pdf");
const isPptxFile = (file) => file.name?.toLowerCase().endsWith(".pptx");
const buildApiUrl = (path) => (apiOrigin ? `${apiOrigin}${path}` : path);

const backendUnavailableMessage = () => {
  if (!apiOrigin) {
    return "This page was opened directly from disk. Open the packaged app or run launch_app.py instead of opening index.html.";
  }
  return `The local worker was not found at ${apiOrigin}. Reopen the app and use the browser window it opens automatically.`;
};

const readLastCompletedSubmission = () => {
  try {
    const rawValue = window.localStorage.getItem(lastCompletedStorageKey);
    if (!rawValue) {
      return null;
    }
    const parsed = JSON.parse(rawValue);
    return parsed && typeof parsed.fingerprint === "string" ? parsed : null;
  } catch {
    return null;
  }
};

const writeLastCompletedSubmission = (submission) => {
  try {
    window.localStorage.setItem(lastCompletedStorageKey, JSON.stringify(submission));
  } catch {
    // Ignore storage failures and keep the in-memory copy.
  }
  lastCompletedSubmission = submission;
};

const buildFileFingerprint = (file) => {
  if (!file) {
    return null;
  }
  return [file.name || "", file.size || 0, file.lastModified || 0].join("::");
};

const buildSelectedPairFingerprint = (pdfFile, pptxFile) => {
  if (!pdfFile || !pptxFile) {
    return null;
  }
  return {
    fingerprint: `${buildFileFingerprint(pdfFile)}__${buildFileFingerprint(pptxFile)}`,
    pdfName: pdfFile.name || "reference.pdf",
    pptxName: pptxFile.name || "candidate.pptx",
    completedAt: new Date().toISOString(),
  };
};

const formatCompletedAt = (isoValue) => {
  if (!isoValue) {
    return "recently";
  }
  const value = new Date(isoValue);
  if (Number.isNaN(value.getTime())) {
    return "recently";
  }
  return value.toLocaleString(undefined, {
    dateStyle: "medium",
    timeStyle: "short",
  });
};

const getCurrentSelectionFingerprint = () => {
  return buildSelectedPairFingerprint(pdfInput.files?.[0], pptxInput.files?.[0]);
};

const updateSubmitAvailability = () => {
  const currentSelection = getCurrentSelectionFingerprint();
  const isDuplicateSelection = Boolean(
    currentSelection &&
      lastCompletedSubmission &&
      currentSelection.fingerprint === lastCompletedSubmission.fingerprint,
  );

  submitButton.disabled = isJobActive || isDuplicateSelection;
  submitButton.dataset.state = isJobActive ? "processing" : isDuplicateSelection ? "duplicate" : "ready";
  submitButton.textContent = isJobActive
    ? "Creating Updated Deck..."
    : isDuplicateSelection
      ? "Already Converted"
      : "Create Updated Deck";

  if (isDuplicateSelection && lastCompletedSubmission) {
    submitStateNote.textContent = `This exact PDF/PPTX pair was already converted on ${formatCompletedAt(lastCompletedSubmission.completedAt)}. Change one of the files to create a new deck.`;
    submitStateNote.classList.remove("hidden");
  } else {
    submitStateNote.textContent = "";
    submitStateNote.classList.add("hidden");
  }
};

const setSingleFile = (input, file) => {
  const transfer = new DataTransfer();
  if (file) {
    transfer.items.add(file);
  }
  input.files = transfer.files;
};

const refreshSelectedFiles = () => {
  const pdfFile = pdfInput.files?.[0];
  const pptxFile = pptxInput.files?.[0];
  pdfName.textContent = pdfFile ? `PDF: ${pdfFile.name}` : "PDF: none selected";
  pptxName.textContent = pptxFile ? `PPTX: ${pptxFile.name}` : "PPTX: none selected";
};

const setDragState = (element, active) => {
  element.classList.toggle("drag-active", active);
};

const escapeHtml = (value) =>
  String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");

const parsePageCountMap = (value) => {
  const counts = {};
  for (const [pageKey, pageCount] of Object.entries(value || {})) {
    const pageNumber = Number.parseInt(pageKey, 10);
    if (!Number.isNaN(pageNumber)) {
      counts[pageNumber] = Number(pageCount) || 0;
    }
  }
  return counts;
};

const sumCounts = (counts) => Object.values(counts || {}).reduce((total, count) => total + (Number(count) || 0), 0);

const allocatePercentages = (items, total, getCount) => {
  if (!items.length || total <= 0) {
    return items.map(() => 0);
  }

  const working = items.map((item, index) => {
    const raw = (getCount(item) / total) * 100;
    const floored = Math.floor(raw);
    return {
      index,
      floored,
      fraction: raw - floored,
    };
  });

  let remainder = 100 - working.reduce((totalFloor, item) => totalFloor + item.floored, 0);
  working.sort((left, right) => right.fraction - left.fraction || left.index - right.index);
  while (remainder > 0 && working.length) {
    for (const item of working) {
      if (remainder <= 0) {
        break;
      }
      item.floored += 1;
      remainder -= 1;
    }
  }

  const percentages = new Array(items.length).fill(0);
  working.forEach((item) => {
    percentages[item.index] = item.floored;
  });
  return percentages;
};

const assignSelectedFiles = (fileList) => {
  const files = Array.from(fileList || []);
  if (!files.length) {
    return;
  }

  const pdfFiles = files.filter(isPdfFile);
  const pptxFiles = files.filter(isPptxFile);
  const unsupportedFiles = files.filter((file) => !isPdfFile(file) && !isPptxFile(file));

  if (unsupportedFiles.length) {
    throw new Error("Only one PDF and one PPTX file are supported.");
  }
  if (pdfFiles.length > 1) {
    throw new Error("Please drop only one PDF file.");
  }
  if (pptxFiles.length > 1) {
    throw new Error("Please drop only one PPTX file.");
  }

  if (!pdfFiles.length && !pptxFiles.length) {
    throw new Error("Please choose a PDF and a PPTX file.");
  }

  if (pdfFiles[0]) {
    setSingleFile(pdfInput, pdfFiles[0]);
  }
  if (pptxFiles[0]) {
    setSingleFile(pptxInput, pptxFiles[0]);
  }

  refreshSelectedFiles();
  updateSubmitAvailability();
  errorCopy.classList.add("hidden");
};

const renderFonts = (fonts, pdfPageCount = 0, pdfPageCharacterTotals = {}) => {
  const pageTotals = parsePageCountMap(pdfPageCharacterTotals);
  const pageNumbers =
    pdfPageCount > 0
      ? Array.from({ length: pdfPageCount }, (_, index) => index + 1)
      : [...new Set(fonts.flatMap((font) => font.pageNumbers || []))].sort((left, right) => left - right);

  if (!fonts?.length && !pageNumbers.length) {
    fontPanel.classList.add("hidden");
    fontUniqueSummary.textContent = "Unique fonts";
    fontUniqueList.innerHTML = "";
    fontGrid.innerHTML = "";
    return;
  }

  const embeddedCount = fonts.filter((font) => font.embedded).length;
  const totalDocumentCharacters = sumCounts(pageTotals);
  fontSummary.textContent =
    totalDocumentCharacters > 0
      ? `${embeddedCount} embedded / ${fonts.length} unique fonts across ${pageNumbers.length} pages · percentages based on extracted text`
      : `${embeddedCount} embedded / ${fonts.length} unique fonts across ${pageNumbers.length} pages · font presence detected, but no extractable text split available`;
  fontUniqueSummary.textContent = `Unique fonts (${fonts.length})`;
  fontUniqueDetails.open = totalDocumentCharacters <= 0;
  const overallPercentages = allocatePercentages(fonts, totalDocumentCharacters, (font) => sumCounts(font.pageCharacterCounts));
  fontUniqueList.innerHTML = fonts.length
    ? fonts
        .map((font, index) => {
          const documentShare = overallPercentages[index];
          const countedCharacters = sumCounts(font.pageCharacterCounts);
          const pagesLabel = font.pageNumbers?.length ? font.pageNumbers.join(", ") : "none";
          const flags = [];
          flags.push(font.embedded ? "embedded" : "not embedded");
          if (font.subset) {
            flags.push("subset");
          }
          const detailLine = totalDocumentCharacters > 0
            ? `${documentShare}% of extracted text · ${countedCharacters} chars · pages ${pagesLabel}`
            : `detected on pages ${pagesLabel} · ${flags.join(" · ")}`;
          return `<article class="font-card"><strong>${escapeHtml(font.name)}</strong><span>${escapeHtml(font.fontType)} · ${escapeHtml(flags.join(" · "))}</span><small>${escapeHtml(detailLine)}</small></article>`;
        })
        .join("")
    : `<p class="font-empty-copy">No extractable text fonts detected in the PDF.</p>`;

  const rows = pageNumbers
    .map((pageNumber) => {
      const pageTotal = pageTotals[pageNumber] || 0;
      const pageFonts = fonts
        .filter((font) => Number(font.pageCharacterCounts?.[pageNumber] || 0) > 0)
        .sort((left, right) => {
          const rightCount = Number(right.pageCharacterCounts?.[pageNumber] || 0);
          const leftCount = Number(left.pageCharacterCounts?.[pageNumber] || 0);
          return rightCount - leftCount || left.name.localeCompare(right.name);
        });

      if (!pageTotal || !pageFonts.length) {
        const detectedFonts = fonts
          .filter((font) => (font.pageNumbers || []).includes(pageNumber))
          .sort((left, right) => left.name.localeCompare(right.name));
        if (!detectedFonts.length) {
          return `<tr><th scope="row">Page ${pageNumber}</th><td><span class="font-tag font-tag-empty">No fonts detected</span></td></tr>`;
        }
        const fallbackTags = detectedFonts
          .map((font) => {
            return `<span class="font-tag font-tag-fallback" title="${escapeHtml(font.fontType)}">${escapeHtml(font.name)} <small>detected</small></span>`;
          })
          .join("");
        return `<tr><th scope="row">Page ${pageNumber}</th><td>${fallbackTags}</td></tr>`;
      }

      const percentages = allocatePercentages(pageFonts, pageTotal, (font) => Number(font.pageCharacterCounts?.[pageNumber] || 0));
      const tags = pageFonts
        .map((font, index) => {
          return `<span class="font-tag" title="${escapeHtml(font.fontType)}">${escapeHtml(font.name)} <small>${percentages[index]}%</small></span>`;
        })
        .join("");
      return `<tr><th scope="row">Page ${pageNumber}</th><td>${tags}</td></tr>`;
    })
    .join("");
  fontGrid.innerHTML = `<thead><tr><th scope="col">Page</th><th scope="col">Fonts used on page</th></tr></thead><tbody>${rows}</tbody>`;
  fontPanel.classList.remove("hidden");
};

const ensureBackendAvailable = async () => {
  if (!apiOrigin) {
    throw new Error(backendUnavailableMessage());
  }

  try {
    const response = await fetch(buildApiUrl("/api/health"), { cache: "no-store" });
    if (!response.ok) {
      throw new Error(backendUnavailableMessage());
    }
  } catch (error) {
    if (error instanceof Error && error.message) {
      throw error;
    }
    throw new Error(backendUnavailableMessage());
  }
};

dropzone.addEventListener("dragenter", () => setDragState(dropzone, true));
dropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  setDragState(dropzone, true);
});
dropzone.addEventListener("dragleave", () => setDragState(dropzone, false));
dropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  setDragState(dropzone, false);
  try {
    assignSelectedFiles(event.dataTransfer?.files);
  } catch (error) {
    errorCopy.textContent = error.message;
    errorCopy.classList.remove("hidden");
  }
});

uploadInput.addEventListener("change", () => {
  try {
    assignSelectedFiles(uploadInput.files);
  } catch (error) {
    errorCopy.textContent = error.message;
    errorCopy.classList.remove("hidden");
  }
});

const updateStatus = (job) => {
  statusPanel.classList.remove("hidden");
  statusBadge.textContent = job.status;
  statusStep.textContent = job.step;
  statusMeta.textContent = `${job.slideProgress} / ${job.slideCount} slides processed`;

  const slideCount = Math.max(job.slideCount || 0, 1);
  const percent = Math.min(100, Math.round(((job.slideProgress || 0) / slideCount) * 100));
  meterBar.style.width = `${percent}%`;

  if (job.status === "queued") {
    statusTitle.textContent = "Waiting in queue";
  } else if (job.status === "processing") {
    statusTitle.textContent = "Parking PDF references";
  } else if (job.status === "completed") {
    statusTitle.textContent = "Updated deck ready";
  } else {
    statusTitle.textContent = "Job failed";
  }

  statusBadge.dataset.state = job.status;
  renderFonts(job.pdfFonts || [], job.pdfPageCount || 0, job.pdfPageCharacterTotals || {});

  if (job.outputReady) {
    downloadLink.href = buildApiUrl(`/api/jobs/${job.jobId}/download`);
    downloadLink.classList.remove("hidden");
  } else {
    downloadLink.classList.add("hidden");
  }

  if (job.error) {
    errorCopy.textContent = job.error;
    errorCopy.classList.remove("hidden");
  } else {
    errorCopy.classList.add("hidden");
  }
};

const pollJob = async (jobId) => {
  try {
    const response = await fetch(buildApiUrl(`/api/jobs/${jobId}`));
    if (!response.ok) {
      throw new Error(`Failed to fetch job status from ${apiOrigin || "the local worker"}.`);
    }

    const job = await response.json();
    updateStatus(job);

    if (job.status === "completed" || job.status === "failed") {
      isJobActive = false;
      if (job.status === "completed" && activeSubmissionFingerprint) {
        writeLastCompletedSubmission({
          ...activeSubmissionFingerprint,
          completedAt: new Date().toISOString(),
        });
      }
      activeSubmissionFingerprint = null;
      updateSubmitAvailability();
      if (pollTimer) {
        clearTimeout(pollTimer);
      }
      return;
    }

    pollTimer = window.setTimeout(() => pollJob(jobId), 1200);
  } catch (error) {
    isJobActive = false;
    activeSubmissionFingerprint = null;
    updateSubmitAvailability();
    errorCopy.textContent = error.message;
    errorCopy.classList.remove("hidden");
  }
};

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  errorCopy.classList.add("hidden");
  downloadLink.classList.add("hidden");

  if (!pdfInput.files?.length || !pptxInput.files?.length) {
    errorCopy.textContent = "Please choose both a PDF and a PPTX file.";
    errorCopy.classList.remove("hidden");
    return;
  }

  activeSubmissionFingerprint = getCurrentSelectionFingerprint();
  isJobActive = true;
  updateSubmitAvailability();
  statusPanel.classList.remove("hidden");
  statusTitle.textContent = "Uploading";
  statusStep.textContent = "Sending files to the local worker…";
  statusMeta.textContent = "0 / 0 slides processed";
  meterBar.style.width = "6%";

  const data = new FormData();
  data.append("pdf", pdfInput.files[0]);
  data.append("pptx", pptxInput.files[0]);

  try {
    await ensureBackendAvailable();

    const response = await fetch(buildApiUrl("/api/compare"), {
      method: "POST",
      body: data,
    });

    if (!response.ok) {
      if (response.status === 404) {
        throw new Error(backendUnavailableMessage());
      }
      const error = await response.json().catch(() => ({ detail: "Upload failed." }));
      throw new Error(error.detail || "Upload failed.");
    }

    const payload = await response.json();
    updateStatus({
      jobId: payload.jobId,
      status: payload.status,
      step: "Queued for processing",
      slideProgress: 0,
      slideCount: 0,
      outputReady: false,
      pdfPageCount: 0,
      pdfPageCharacterTotals: {},
      pdfFonts: [],
      error: null,
    });
    pollJob(payload.jobId);
  } catch (error) {
    isJobActive = false;
    activeSubmissionFingerprint = null;
    updateSubmitAvailability();
    errorCopy.textContent = error.message;
    errorCopy.classList.remove("hidden");
  }
});

window.addEventListener("DOMContentLoaded", async () => {
  lastCompletedSubmission = readLastCompletedSubmission();
  updateSubmitAvailability();
  try {
    await ensureBackendAvailable();
    errorCopy.classList.add("hidden");
  } catch (error) {
    errorCopy.textContent = error.message;
    errorCopy.classList.remove("hidden");
  }
});
