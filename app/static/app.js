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

const uploadInput = document.getElementById("upload-input");
const pdfInput = document.getElementById("pdf-input");
const pptxInput = document.getElementById("pptx-input");
const pdfName = document.getElementById("pdf-name");
const pptxName = document.getElementById("pptx-name");
const runtimeHelper = document.getElementById("runtime-helper");
const runtimeNote = document.getElementById("runtime-note");
const submitStateNote = document.getElementById("submit-state-note");
const dropzone = document.querySelector(".dropzone");
const aiQcToggle = document.getElementById("ai-qc-toggle");
const aiQcToggleRow = document.querySelector('label[for="ai-qc-toggle"]');
const aiQcToggleNote = document.querySelector(".toggle-note");
const promptOpenButton = document.getElementById("prompt-open-button");
const promptModal = document.getElementById("prompt-modal");
const promptModalBackdrop = document.getElementById("prompt-modal-backdrop");
const promptCloseButton = document.getElementById("prompt-close-button");
const qcGeneralSystemPrompt = document.getElementById("qc-general-system-prompt");
const qcGeneralUserPrompt = document.getElementById("qc-general-user-prompt");
const qcTextSystemPrompt = document.getElementById("qc-text-system-prompt");
const qcTextUserPrompt = document.getElementById("qc-text-user-prompt");
const promptResetButton = document.getElementById("prompt-reset-button");
const promptSaveButton = document.getElementById("prompt-save-button");
const promptSaveStatus = document.getElementById("prompt-save-status");

const debugLog = (...args) => console.log("[prompt-editor]", ...args);
const debugError = (...args) => console.error("[prompt-editor]", ...args);

let pollTimer = null;
let isJobActive = false;
let activeSubmissionFingerprint = null;
let lastCompletedSubmission = null;
let rendererCanConvert = true;
let rendererMessage = "";
let aiQcSupported = true;
let aiQcAvailable = false;
const lastCompletedStorageKey = "pdf-to-pptx-last-completed-pair";
const promptConfigStorageKey = "pdf-to-pptx-qc-prompt-config-v1";
const apiOrigin =
  window.location.protocol === "http:" || window.location.protocol === "https:"
    ? window.location.origin
    : null;

const isPdfFile = (file) => file.name?.toLowerCase().endsWith(".pdf");
const isPptxFile = (file) => file.name?.toLowerCase().endsWith(".pptx");
const buildApiUrl = (path) => (apiOrigin ? `${apiOrigin}${path}` : path);
const hasFilePayload = (event) => Array.from(event.dataTransfer?.types || []).includes("Files");

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

let defaultPromptConfig = null;

const buildSelectedPairFingerprint = (
  pdfFile,
  pptxFile,
  aiQcEnabled = aiQcToggle?.checked ?? false,
  promptConfig = getCurrentPromptConfig(),
) => {
  if (!pdfFile || !pptxFile) {
    return null;
  }
  const normalizedPromptConfig = aiQcEnabled ? JSON.stringify(promptConfig) : "";
  return {
    fingerprint: `${buildFileFingerprint(pdfFile)}__${buildFileFingerprint(pptxFile)}__ai:${aiQcEnabled ? "on" : "off"}__prompt:${normalizedPromptConfig}`,
    pdfName: pdfFile.name || "reference.pdf",
    pptxName: pptxFile.name || "candidate.pptx",
    aiQcEnabled,
    promptConfig: aiQcEnabled ? promptConfig : {},
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
  return buildSelectedPairFingerprint(pdfInput.files?.[0], pptxInput.files?.[0], aiQcToggle?.checked ?? false);
};

const getDefaultPromptConfig = () =>
  defaultPromptConfig || {
    generalSystemPrompt: "",
    generalUserPrompt: "",
    textSystemPrompt: "",
    textUserPrompt: "",
  };

const getCurrentPromptConfig = () => ({
  generalSystemPrompt: qcGeneralSystemPrompt?.value || "",
  generalUserPrompt: qcGeneralUserPrompt?.value || "",
  textSystemPrompt: qcTextSystemPrompt?.value || "",
  textUserPrompt: qcTextUserPrompt?.value || "",
});

const applyPromptConfig = (config) => {
  if (qcGeneralSystemPrompt) {
    qcGeneralSystemPrompt.value = config.generalSystemPrompt || "";
  }
  if (qcGeneralUserPrompt) {
    qcGeneralUserPrompt.value = config.generalUserPrompt || "";
  }
  if (qcTextSystemPrompt) {
    qcTextSystemPrompt.value = config.textSystemPrompt || "";
  }
  if (qcTextUserPrompt) {
    qcTextUserPrompt.value = config.textUserPrompt || "";
  }
};

const promptsMatch = (left, right) =>
  (left?.generalSystemPrompt || "") === (right?.generalSystemPrompt || "") &&
  (left?.generalUserPrompt || "") === (right?.generalUserPrompt || "") &&
  (left?.textSystemPrompt || "") === (right?.textSystemPrompt || "") &&
  (left?.textUserPrompt || "") === (right?.textUserPrompt || "");

const readSavedPromptConfig = () => {
  try {
    const rawValue = window.localStorage.getItem(promptConfigStorageKey);
    if (!rawValue) {
      return null;
    }
    const parsed = JSON.parse(rawValue);
    return parsed && typeof parsed === "object" ? parsed : null;
  } catch {
    return null;
  }
};

const writeSavedPromptConfig = (config) => {
  try {
    window.localStorage.setItem(promptConfigStorageKey, JSON.stringify(config));
  } catch {
    // Ignore storage failures and keep the in-memory editor state.
  }
};

const clearSavedPromptConfig = () => {
  try {
    window.localStorage.removeItem(promptConfigStorageKey);
  } catch {
    // Ignore storage failures.
  }
};

const updatePromptEditorState = () => {
  if (!promptSaveStatus) {
    return;
  }
  const isDefault = promptsMatch(getCurrentPromptConfig(), getDefaultPromptConfig());
  promptSaveStatus.textContent = isDefault
    ? "The editor currently matches the default prompts."
    : "The editor currently contains custom prompts.";
};

const updatePromptStatus = () => {
  const isDefault = promptsMatch(getCurrentPromptConfig(), getDefaultPromptConfig());
  if (!promptSaveStatus) {
    return;
  }
  if (!(aiQcToggle?.checked ?? false)) {
    promptSaveStatus.textContent = isDefault
      ? "AI QC is off. The editor currently matches the default prompts."
      : "AI QC is off. The editor currently contains custom prompts.";
    return;
  }
  promptSaveStatus.textContent = isDefault
    ? "The editor currently matches the default prompts."
    : "The editor currently contains custom prompts.";
};

const setPromptModalOpen = (isOpen) => {
  if (!promptModal) {
    debugError("setPromptModalOpen called but promptModal was not found.");
    return;
  }
  debugLog("setPromptModalOpen", {
    isOpen,
    hadHiddenClass: promptModal.classList.contains("hidden"),
    ariaHidden: promptModal.getAttribute("aria-hidden"),
  });
  promptModal.classList.toggle("hidden", !isOpen);
  promptModal.style.display = isOpen ? "block" : "";
  promptModal.setAttribute("aria-hidden", isOpen ? "false" : "true");
  debugLog("setPromptModalOpen applied", {
    isOpen,
    hasHiddenClass: promptModal.classList.contains("hidden"),
    computedDisplay: window.getComputedStyle(promptModal).display,
    ariaHidden: promptModal.getAttribute("aria-hidden"),
  });
};

window.__openPromptEditor = (source = "unknown") => {
  debugLog("__openPromptEditor invoked", {
    source,
    buttonDisabled: promptOpenButton?.disabled ?? null,
    modalFound: Boolean(promptModal),
  });
  if (promptOpenButton?.disabled) {
    debugLog("__openPromptEditor aborted because button is disabled.");
    return;
  }
  setPromptModalOpen(true);
};

const updatePromptAvailability = () => {
  const aiEnabled = (aiQcToggle?.checked ?? false) && aiQcSupported && aiQcAvailable;
  debugLog("updatePromptAvailability", { aiEnabled });
  [qcGeneralSystemPrompt, qcGeneralUserPrompt, qcTextSystemPrompt, qcTextUserPrompt].forEach((field) => {
    if (field) {
      field.disabled = !aiEnabled;
    }
  });
  if (promptResetButton) {
    promptResetButton.disabled = !aiEnabled;
  }
  if (promptSaveButton) {
    promptSaveButton.disabled = !aiEnabled;
  }
  if (promptOpenButton) {
    promptOpenButton.disabled = !aiEnabled;
  }
  if (!aiEnabled) {
    setPromptModalOpen(false);
  }
  updatePromptStatus();
};

const applyRendererStatus = (renderer) => {
  rendererCanConvert = Boolean(renderer?.canConvert);
  rendererMessage = renderer?.message || "";

  if (runtimeHelper && rendererMessage) {
    runtimeHelper.textContent = rendererMessage;
  }

  if (!runtimeNote) {
    return;
  }

  if (rendererCanConvert) {
    runtimeNote.textContent = "";
    runtimeNote.classList.add("hidden");
  } else {
    runtimeNote.textContent = rendererMessage || "No supported converter is available on this machine.";
    runtimeNote.classList.remove("hidden");
  }
};

const applyAiQcAvailability = ({ supported, available }) => {
  aiQcSupported = Boolean(supported);
  aiQcAvailable = Boolean(available);
  const aiControlsVisible = aiQcSupported && aiQcAvailable;

  if (aiQcToggle) {
    if (!aiControlsVisible) {
      aiQcToggle.checked = false;
    }
    aiQcToggle.disabled = !aiControlsVisible;
  }
  if (aiQcToggleRow) {
    aiQcToggleRow.classList.toggle("hidden", !aiControlsVisible);
  }
  if (aiQcToggleNote) {
    aiQcToggleNote.classList.toggle("hidden", !aiControlsVisible);
  }
  if (promptOpenButton) {
    promptOpenButton.classList.toggle("hidden", !aiControlsVisible);
  }

  if (!aiQcSupported) {
    runtimeHelper.textContent =
      "This build is configured for PDF reference insertion only. AI QC is disabled.";
  } else if (!aiQcAvailable) {
    runtimeHelper.textContent =
      "AI QC is unavailable in this build because no OpenAI API key is configured.";
  }

  updatePromptAvailability();
};

const updateSubmitAvailability = () => {
  const currentSelection = getCurrentSelectionFingerprint();
  const isDuplicateSelection = Boolean(
    currentSelection &&
      lastCompletedSubmission &&
      currentSelection.fingerprint === lastCompletedSubmission.fingerprint,
  );

  submitButton.disabled = isJobActive || isDuplicateSelection || !rendererCanConvert;
  submitButton.dataset.state = isJobActive
    ? "processing"
    : !rendererCanConvert
      ? "unavailable"
      : isDuplicateSelection
        ? "duplicate"
        : "ready";
  submitButton.textContent = isJobActive
    ? "Creating Updated Deck..."
    : !rendererCanConvert
      ? "Install LibreOffice or PowerPoint"
    : isDuplicateSelection
      ? "Already Converted"
      : "Create Updated Deck";

  if (!submitStateNote) {
    return;
  }

  if (!rendererCanConvert) {
    submitStateNote.textContent = rendererMessage || "A supported converter is not available on this machine.";
    submitStateNote.classList.remove("hidden");
  } else if (isDuplicateSelection && lastCompletedSubmission) {
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
  if (!element) {
    return;
  }
  element.classList.toggle("drag-active", active);
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

const ensureBackendAvailable = async () => {
  if (!apiOrigin) {
    throw new Error(backendUnavailableMessage());
  }

  try {
    const response = await fetch(buildApiUrl("/api/health"), { cache: "no-store" });
    if (!response.ok) {
      throw new Error(backendUnavailableMessage());
    }
    const payload = await response.json();
    applyRendererStatus(payload.renderer);
    applyAiQcAvailability({
      supported: payload.aiQcSupported,
      available: payload.aiQcAvailable,
    });
    updateSubmitAvailability();
    if (!payload.renderer?.canConvert) {
      throw new Error(payload.renderer?.message || "No supported converter is available on this machine.");
    }
  } catch (error) {
    if (error instanceof Error && error.message) {
      throw error;
    }
    throw new Error(backendUnavailableMessage());
  }
};

const handleDroppedFiles = (files) => {
  try {
    assignSelectedFiles(files);
  } catch (error) {
    errorCopy.textContent = error.message;
    errorCopy.classList.remove("hidden");
  }
};

if (dropzone) {
  dropzone.addEventListener("dragenter", (event) => {
    if (!hasFilePayload(event)) {
      return;
    }
    event.preventDefault();
    setDragState(dropzone, true);
  });
  dropzone.addEventListener("dragover", (event) => {
    if (!hasFilePayload(event)) {
      return;
    }
    event.preventDefault();
    setDragState(dropzone, true);
  });
  dropzone.addEventListener("dragleave", () => setDragState(dropzone, false));
  dropzone.addEventListener("drop", (event) => {
    if (!hasFilePayload(event)) {
      return;
    }
    event.preventDefault();
    setDragState(dropzone, false);
    handleDroppedFiles(event.dataTransfer?.files);
  });
}

window.addEventListener("dragover", (event) => {
  if (!hasFilePayload(event)) {
    return;
  }
  event.preventDefault();
});

window.addEventListener("drop", (event) => {
  if (!hasFilePayload(event)) {
    return;
  }
  event.preventDefault();
  setDragState(dropzone, false);

  const droppedInsideZone = dropzone && event.target instanceof Node && dropzone.contains(event.target);
  if (!droppedInsideZone) {
    handleDroppedFiles(event.dataTransfer?.files);
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

if (aiQcToggle) {
  aiQcToggle.addEventListener("change", () => {
    debugLog("AI QC toggle changed", { checked: aiQcToggle.checked });
    updatePromptAvailability();
    updateSubmitAvailability();
  });
}

[
  qcGeneralSystemPrompt,
  qcGeneralUserPrompt,
  qcTextSystemPrompt,
  qcTextUserPrompt,
].forEach((field) => {
  if (!field) {
    return;
  }
  field.addEventListener("input", () => {
    updatePromptStatus();
    updatePromptEditorState();
    updateSubmitAvailability();
  });
});

if (promptOpenButton) {
  promptOpenButton.addEventListener("click", () => {
    debugLog("Edit Prompt button clicked", {
      disabled: promptOpenButton.disabled,
      modalFound: Boolean(promptModal),
    });
    window.__openPromptEditor("listener");
  });
} else {
  debugError("Prompt open button was not found during script setup.");
}

if (promptModalBackdrop) {
  promptModalBackdrop.addEventListener("click", () => {
    debugLog("Prompt modal backdrop clicked.");
    setPromptModalOpen(false);
  });
} else {
  debugError("Prompt modal backdrop was not found during script setup.");
}

if (promptCloseButton) {
  promptCloseButton.addEventListener("click", () => {
    debugLog("Prompt modal close button clicked.");
    setPromptModalOpen(false);
  });
} else {
  debugError("Prompt modal close button was not found during script setup.");
}

if (promptSaveButton) {
  promptSaveButton.addEventListener("click", () => {
    debugLog("Prompt save clicked.");
    writeSavedPromptConfig(getCurrentPromptConfig());
    if (promptSaveStatus) {
      promptSaveStatus.textContent = "Custom prompts saved locally in this browser.";
    }
    updateSubmitAvailability();
  });
}

if (promptResetButton) {
  promptResetButton.addEventListener("click", () => {
    debugLog("Prompt reset clicked.");
    applyPromptConfig(getDefaultPromptConfig());
    clearSavedPromptConfig();
    if (promptSaveStatus) {
      promptSaveStatus.textContent = "Prompt editor reset to the default prompts.";
    }
    updateSubmitAvailability();
    qcGeneralSystemPrompt?.focus();
  });
}

const updateStatus = (job) => {
  statusPanel?.classList.remove("hidden");
  if (statusBadge) {
    statusBadge.textContent = job.status;
  }
  if (statusStep) {
    statusStep.textContent = job.step;
  }
  if (statusMeta) {
    statusMeta.textContent = `${job.slideProgress} / ${job.slideCount} slides processed`;
  }

  const slideCount = Math.max(job.slideCount || 0, 1);
  const percent = Math.min(100, Math.round(((job.slideProgress || 0) / slideCount) * 100));
  if (meterBar) {
    meterBar.style.width = `${percent}%`;
  }

  if (job.status === "queued") {
    if (statusTitle) {
      statusTitle.textContent = "Waiting in queue";
    }
  } else if (job.status === "processing") {
    if (statusTitle) {
      statusTitle.textContent = job.aiQcEnabled === false ? "Preparing PDF reference deck" : "Parking PDF references";
    }
  } else if (job.status === "completed") {
    if (statusTitle) {
      statusTitle.textContent = "Updated deck ready";
    }
  } else {
    if (statusTitle) {
      statusTitle.textContent = "Job failed";
    }
  }

  if (statusBadge) {
    statusBadge.dataset.state = job.status;
  }

  if (job.outputReady && downloadLink) {
    downloadLink.href = buildApiUrl(`/api/jobs/${job.jobId}/download`);
    downloadLink.classList.remove("hidden");
  } else if (downloadLink) {
    downloadLink.classList.add("hidden");
  }

  if (job.error && errorCopy) {
    errorCopy.textContent = job.error;
    errorCopy.classList.remove("hidden");
  } else if (errorCopy) {
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
  const aiQcEnabledForSubmission = (aiQcToggle?.checked ?? false) && aiQcSupported && aiQcAvailable;
  data.append("enable_ai_qc", aiQcEnabledForSubmission ? "true" : "false");
  if (aiQcEnabledForSubmission) {
    const promptConfig = getCurrentPromptConfig();
    data.append("qc_general_system_prompt", promptConfig.generalSystemPrompt);
    data.append("qc_general_user_prompt", promptConfig.generalUserPrompt);
    data.append("qc_text_system_prompt", promptConfig.textSystemPrompt);
    data.append("qc_text_user_prompt", promptConfig.textUserPrompt);
  }

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
      step: payload.aiQcEnabled ? "Queued for processing" : "Queued for PDF insertion",
      slideProgress: 0,
      slideCount: 0,
      outputReady: false,
      aiQcEnabled: payload.aiQcEnabled,
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
  debugLog("DOMContentLoaded", {
    apiOrigin,
    promptOpenButtonFound: Boolean(promptOpenButton),
    promptModalFound: Boolean(promptModal),
    promptModalBackdropFound: Boolean(promptModalBackdrop),
    promptCloseButtonFound: Boolean(promptCloseButton),
  });
  lastCompletedSubmission = readLastCompletedSubmission();
  setPromptModalOpen(false);
  updateSubmitAvailability();
  try {
    debugLog("Checking backend availability.");
    await ensureBackendAvailable();
    if (aiQcSupported) {
      debugLog("Backend available. Loading /api/qc-prompts.");
      const promptResponse = await fetch(buildApiUrl("/api/qc-prompts"), { cache: "no-store" });
      if (!promptResponse.ok) {
        throw new Error("Could not load the AI QC prompts.");
      }
      defaultPromptConfig = await promptResponse.json();
      debugLog("Loaded prompt defaults.", {
        generalSystemLength: defaultPromptConfig?.generalSystemPrompt?.length || 0,
        generalUserLength: defaultPromptConfig?.generalUserPrompt?.length || 0,
        textSystemLength: defaultPromptConfig?.textSystemPrompt?.length || 0,
        textUserLength: defaultPromptConfig?.textUserPrompt?.length || 0,
      });
      applyPromptConfig(readSavedPromptConfig() || defaultPromptConfig);
    } else {
      defaultPromptConfig = getDefaultPromptConfig();
      applyPromptConfig(defaultPromptConfig);
    }
    updatePromptAvailability();
    updatePromptStatus();
    updatePromptEditorState();
    errorCopy.classList.add("hidden");
  } catch (error) {
    debugError("DOMContentLoaded initialization failed.", error);
    errorCopy.textContent = error.message;
    errorCopy.classList.remove("hidden");
  }
});

window.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    debugLog("Escape pressed. Closing prompt modal.");
    setPromptModalOpen(false);
  }
});

window.addEventListener("error", (event) => {
  debugError("Unhandled window error.", event.error || event.message || event);
});

window.addEventListener("unhandledrejection", (event) => {
  debugError("Unhandled promise rejection.", event.reason || event);
});

debugLog("Script loaded.", {
  promptOpenButtonFound: Boolean(promptOpenButton),
  promptModalFound: Boolean(promptModal),
});
