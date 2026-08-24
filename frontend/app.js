const API_BASE = "/api";
const EXCEL_EXTENSIONS = [".xlsx", ".xlsm"];
const LOT_PATTERN = /(?:^|[^a-z0-9])(?:lot\s*(?:n|no|num|numero|n°)?\s*)?(\d{1,3})(?:\s*(?:bis|ter)|[a-z])?(?=[^a-z0-9]|$)/i;

const state = {
  dceFiles: [],
  actFiles: [],
  fileLots: {},
  actEnterpriseNames: [],
  progressTimer: null,
  progressStepIndex: 0,
  activeRunContext: "files",
};

const DROPZONE_PROGRESS_STEPS = [
  "Préparation des fichiers Excel déposés.",
  "Contrôle des extensions et des numéros de lot dans les noms de fichiers.",
  "Regroupement des fichiers par lot.",
  "Lecture des classeurs exploitables.",
  "Création du tableau restructuré.",
  "Finalisation du classeur et préparation du téléchargement.",
];

const elements = {
  helpToggle: document.getElementById("helpToggle"),
  helpPanel: document.getElementById("helpPanel"),
  dceDropzone: document.getElementById("dceDropzone"),
  dceInput: document.getElementById("dceInput"),
  dceFileList: document.getElementById("dceFileList"),
  actDropzone: document.getElementById("actDropzone"),
  actInput: document.getElementById("actInput"),
  actFileList: document.getElementById("actFileList"),
  runFilesButton: document.getElementById("runFilesButton"),
  dropHint: document.getElementById("dropHint"),
  analysisReport: document.getElementById("analysisReport"),
  filesXlsxStatus: document.getElementById("filesXlsxStatus"),
  filesXlsxDownload: document.getElementById("filesXlsxDownload"),
  lotFilesError: document.getElementById("lotFilesError"),
  runOfferAnalysisButton: document.getElementById("runOfferAnalysisButton"),
  offerAnalysisStatus: document.getElementById("offerAnalysisStatus"),
  analysisXlsxDownload: document.getElementById("analysisXlsxDownload"),
  analysisDropzone: document.getElementById("analysisDropzone"),
  analysisInput: document.getElementById("analysisInput"),
  analysisSelectedFile: document.getElementById("analysisSelectedFile"),
  analysisSelectedFileName: document.getElementById("analysisSelectedFileName"),
  removeAnalysisFileButton: document.getElementById("removeAnalysisFileButton"),
  enterpriseNamesError: document.getElementById("enterpriseNamesError"),
};

function fileExtension(fileName) {
  const match = String(fileName || "").toLowerCase().match(/\.[^.]+$/);
  return match ? match[0] : "";
}

function isExcelFile(file) {
  return EXCEL_EXTENSIONS.includes(fileExtension(file.name));
}

function extractLotNumber(fileName) {
  const normalized = String(fileName || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
  // Détection stricte : les nombres d'une date ou d'un numéro d'affaire ne sont pas des lots.
  const explicit = normalized.match(/(?:^|[^a-z0-9])(?:lot|l)\s*(?:n|no|num|numero|n°)?\s*[-_ ]*(\d{1,3})(?=[^a-z0-9]|$)/i);
  const value = explicit?.[1] || "";
  return value ? value.padStart(2, "0") : "";
}

function normalizeLotNumber(value) {
  const match = String(value || "").trim().match(/\d{1,3}/);
  return match ? match[0].padStart(2, "0") : "";
}

function assignedLot(file) {
  return state.fileLots[fileKey(file)] || extractLotNumber(file.name);
}

function fileKey(file) {
  return `${file.name}::${file.size}::${file.lastModified}`;
}

function mergeFiles(existingFiles, incomingFiles) {
  const merged = [...existingFiles];
  const seen = new Set(existingFiles.map(fileKey));
  incomingFiles.forEach((file) => {
    const key = fileKey(file);
    if (!seen.has(key)) {
      seen.add(key);
      merged.push(file);
    }
  });
  return merged;
}

function describeFile(file) {
  const isExcel = isExcelFile(file);
  const lot = assignedLot(file);
  const issues = [];
  if (!isExcel) issues.push("format non Excel");
  if (!lot) issues.push("lot interne détecté côté serveur");
  return { isExcel, lot, issues };
}

function renderFiles(listElement, files, type) {
  listElement.innerHTML = "";
  if (!files.length) {
    const li = document.createElement("li");
    li.textContent = "Aucun fichier sélectionné";
    listElement.appendChild(li);
    return;
  }

  files.forEach((file, index) => {
    const desc = describeFile(file);
    const li = document.createElement("li");
    li.className = `file-list__row lot-file-row${desc.issues.length ? " is-missing" : ""}`;

    const nameBlock = document.createElement("div");
    nameBlock.className = "lot-file-row__name";

    const strong = document.createElement("strong");
    strong.textContent = file.name;
    nameBlock.appendChild(strong);

    const hint = document.createElement("span");
    hint.textContent = desc.lot ? `Lot ${desc.lot}` : (type === "dce" ? "DCE général / lots internes" : desc.issues.join(" + "));
    nameBlock.appendChild(hint);

    const removeButton = document.createElement("button");
    removeButton.className = "trash-button";
    removeButton.type = "button";
    removeButton.title = "Retirer ce fichier de la sélection";
    removeButton.setAttribute("aria-label", `Retirer ${file.name} de la sélection`);
    removeButton.textContent = "X";
    removeButton.addEventListener("click", (event) => {
      event.stopPropagation();
      removeSelectedFile(type, index);
    });

    li.appendChild(nameBlock);
    li.appendChild(removeButton);
    listElement.appendChild(li);
  });
}


function validateIncomingFiles(files) {
  const invalidExcel = files.filter((file) => !isExcelFile(file));
  if (invalidExcel.length) {
    const details = invalidExcel.map((file) => `- ${file.name}`).join("\n");
    window.alert(
      "Fichier(s) refusé(s) : format non Excel.\n\n" + details +
      "\n\nFormats autorisés : .xlsx ou .xlsm."
    );
  }
  return files.filter(isExcelFile);
}

function syncActEnterpriseNames() {
  const previous = new Map(state.actEnterpriseNames.map((item) => [item.key, item]));
  state.actEnterpriseNames = state.actFiles.map((file, index) => {
    const key = fileKey(file);
    const old = previous.get(key);
    if (old) {
      old.index = index;
      old.original_name = file.name;
      old.lot_number = assignedLot(file);
      return old;
    }
    return {
      key,
      index,
      original_name: file.name,
      enterprise_name: file.name.replace(/\.(xlsx|xlsm)$/i, "").replace(/(?:dce|dpgf|chiffrage|lot\s*\d{1,3})/gi, " ").replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim(),
      lot_number: assignedLot(file),
      source: "user_input_required",
    };
  });
}

function renderActFiles() {
  syncActEnterpriseNames();
  elements.actFileList.innerHTML = "";
  if (!state.actFiles.length) {
    const li = document.createElement("li");
    li.textContent = "Aucun fichier sélectionné";
    elements.actFileList.appendChild(li);
    elements.enterpriseNamesError?.classList.add("hidden");
    return;
  }

  state.actFiles.forEach((file, index) => {
    const item = state.actEnterpriseNames[index];
    const lot = assignedLot(file);
    const li = document.createElement("li");
    li.className = `act-file-row${!item.enterprise_name ? " is-missing" : ""}`;

    const nameBlock = document.createElement("div");
    nameBlock.className = "act-file-row__name";
    const strong = document.createElement("strong");
    strong.textContent = file.name;
    const hint = document.createElement("span");
    hint.textContent = lot ? `Lot ${lot}` : "Lots internes détectés côté serveur";
    nameBlock.appendChild(strong);
    nameBlock.appendChild(hint);

    const lotInput = document.createElement("input");
    lotInput.className = "text-input lot-input";
    lotInput.type = "text";
    lotInput.inputMode = "numeric";
    lotInput.placeholder = "N° lot";
    lotInput.value = lot;
    lotInput.addEventListener("input", () => {
      const value = normalizeLotNumber(lotInput.value);
      state.fileLots[fileKey(file)] = value;
      state.actEnterpriseNames[index].lot_number = value;
      hint.textContent = value ? `Lot ${value}` : "Lot à renseigner";
      li.classList.toggle("is-missing", !state.actEnterpriseNames[index].enterprise_name);
      updateRunButtonState();
    });

    const enterpriseInput = document.createElement("input");
    enterpriseInput.className = "text-input";
    enterpriseInput.type = "text";
    enterpriseInput.placeholder = "Nom de l’entreprise";
    enterpriseInput.value = item.enterprise_name || "";
    enterpriseInput.addEventListener("input", () => {
      const value = enterpriseInput.value.trim();
      state.actEnterpriseNames[index].enterprise_name = value;
      state.actEnterpriseNames[index].source = value ? "user_input" : "user_input_required";
      li.classList.toggle("is-missing", !value);
      elements.enterpriseNamesError?.classList.add("hidden");
      updateRunButtonState();
    });

    const removeButton = document.createElement("button");
    removeButton.className = "trash-button";
    removeButton.type = "button";
    removeButton.textContent = "X";
    removeButton.setAttribute("aria-label", `Retirer ${file.name}`);
    removeButton.addEventListener("click", (event) => {
      event.stopPropagation();
      removeSelectedFile("act", index);
    });

    li.appendChild(nameBlock);
    li.appendChild(lotInput);
    li.appendChild(enterpriseInput);
    li.appendChild(removeButton);
    elements.actFileList.appendChild(li);
  });
}

function validateEnterpriseNames() {
  syncActEnterpriseNames();
  const valid = state.actEnterpriseNames.every((item) => String(item.enterprise_name || "").trim());
  elements.enterpriseNamesError?.classList.toggle("hidden", valid || !state.actFiles.length);
  return valid;
}

function enterpriseNamePayload() {
  syncActEnterpriseNames();
  return state.actEnterpriseNames.map((item, index) => ({
    index,
    original_name: item.original_name,
    enterprise_name: String(item.enterprise_name || "").trim(),
    lot_number: assignedLot(state.actFiles[index]),
    source: item.source || "user_input",
  }));
}

function validateDroppedFiles() {
  const dceValid = state.dceFiles.every((file) => isExcelFile(file));
  const actValid = state.actFiles.every((file) => isExcelFile(file)); // TCE_VIRTUAL_SHEETS_V22 : lot facultatif, catalogue serveur
  return state.dceFiles.length > 0 && state.actFiles.length > 0 && dceValid && actValid;
}

function updateRunButtonState() {
  const hasRequiredZones = state.dceFiles.length > 0 && state.actFiles.length > 0;
  const filesValid = validateDroppedFiles();
  const namesValid = state.actFiles.length > 0 && state.actEnterpriseNames.length === state.actFiles.length && state.actEnterpriseNames.every((item) => String(item.enterprise_name || "").trim());
  elements.runFilesButton.disabled = !(hasRequiredZones && filesValid && namesValid);

  if (!hasRequiredZones) {
    elements.dropHint.textContent = "Déposez au moins un DCE Excel et un ACT Excel pour lancer la restructuration.";
  } else if (!filesValid) {
    elements.dropHint.textContent = "Chaque retour doit être un Excel. Les lots mono-lot ou TCE sont détectés côté serveur.";
  } else if (!namesValid) {
    elements.dropHint.textContent = "Renseignez le nom de l’entreprise pour chaque fichier ACT.";
  } else {
    const lots = [...new Set([...state.dceFiles, ...state.actFiles].map(assignedLot).filter(Boolean))].sort();
    elements.dropHint.textContent = `Lots renseignés : ${lots.join(", ")}. Vous pouvez lancer la restructuration.`;
  }

  if (elements.lotFilesError) {
    elements.lotFilesError.classList.toggle("hidden", filesValid || !hasRequiredZones);
  }
}

function setLoading(isLoading) {
  const namesValid = state.actEnterpriseNames.length === state.actFiles.length && state.actEnterpriseNames.every((item) => String(item.enterprise_name || "").trim());
  elements.runFilesButton.disabled = isLoading || !(state.dceFiles.length > 0 && state.actFiles.length > 0 && validateDroppedFiles() && namesValid);
  elements.runFilesButton.textContent = isLoading ? "Traitement en cours..." : "Lancer la restructuration des fichiers déposés";
}

function clearProgressFeedback() {
  if (state.progressTimer) {
    clearInterval(state.progressTimer);
    state.progressTimer = null;
  }
}

function startProgressFeedback() {
  clearProgressFeedback();
  state.progressStepIndex = 0;
  elements.analysisReport.className = "report-box report-status--warning report-box--progress";
  elements.analysisReport.innerHTML = "";

  const wrapper = document.createElement("div");
  wrapper.className = "progress-panel";
  const title = document.createElement("h3");
  title.textContent = "Restructuration des fichiers déposés";
  const intro = document.createElement("p");
  intro.className = "progress-panel__intro";
  intro.textContent = "Les fichiers Excel sont envoyés au serveur. Le rapport complet sera affiché dès que le classeur restructuré sera prêt.";
  const bar = document.createElement("div");
  bar.className = "progress-bar";
  bar.setAttribute("role", "progressbar");
  bar.setAttribute("aria-valuemin", "0");
  bar.setAttribute("aria-valuemax", "100");
  bar.setAttribute("aria-valuenow", "0");
  const fill = document.createElement("div");
  fill.className = "progress-bar__fill";
  bar.appendChild(fill);
  const current = document.createElement("p");
  current.className = "progress-panel__current";
  const list = document.createElement("ol");
  list.className = "progress-steps";
  DROPZONE_PROGRESS_STEPS.forEach((step) => {
    const item = document.createElement("li");
    item.textContent = step;
    list.appendChild(item);
  });

  wrapper.appendChild(title);
  wrapper.appendChild(intro);
  wrapper.appendChild(bar);
  wrapper.appendChild(current);
  wrapper.appendChild(list);
  elements.analysisReport.appendChild(wrapper);

  const stepItems = Array.from(list.children);
  const updateProgress = () => {
    const index = Math.min(state.progressStepIndex, DROPZONE_PROGRESS_STEPS.length - 1);
    const percent = Math.min(92, Math.round(((index + 1) / DROPZONE_PROGRESS_STEPS.length) * 88));
    fill.style.width = `${percent}%`;
    bar.setAttribute("aria-valuenow", String(percent));
    current.textContent = DROPZONE_PROGRESS_STEPS[index];
    stepItems.forEach((item, itemIndex) => {
      item.classList.toggle("is-done", itemIndex < index);
      item.classList.toggle("is-current", itemIndex === index);
    });
    if (state.progressStepIndex < DROPZONE_PROGRESS_STEPS.length - 1) state.progressStepIndex += 1;
  };
  updateProgress();
  state.progressTimer = setInterval(updateProgress, 4200);
}

function setReportMessage(message, statusClass = "") {
  clearProgressFeedback();
  elements.analysisReport.className = "report-box";
  if (statusClass) elements.analysisReport.classList.add(statusClass);
  elements.analysisReport.innerHTML = "";
  const p = document.createElement("p");
  p.textContent = message;
  elements.analysisReport.appendChild(p);
}

function arrayify(value) {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function uniqueStrings(value) {
  return [...new Set(arrayify(value).map((item) => String(item)).filter(Boolean))];
}

function normalizeBackendResult(rawData) {
  const payload = rawData?.data?.result || rawData?.data?.run || rawData?.data || rawData || {};
  const primaryExport = Array.isArray(payload.creation_de_tableau_exports)
    ? payload.creation_de_tableau_exports.find((item) => item?.workbook_path || item?.download_url) || {}
    : {};
  return {
    status: payload.status || primaryExport.status || (rawData?.success ? "OK" : "ERROR"),
    workbookPath: payload.workbook_path || payload.xlsx || payload.xlsx_path || payload.workbook || primaryExport.workbook_path || "",
    warnings: uniqueStrings([...arrayify(payload.warnings), ...arrayify(rawData?.warnings)]),
    errors: uniqueStrings([...arrayify(payload.errors), ...arrayify(rawData?.errors), ...arrayify(primaryExport.error)]),
    downloadUrl: payload.download_url || payload.xlsx_url || payload.workbook_url || primaryExport.download_url || primaryExport.workbook_url || "",
  };
}

function createReportSection(title, text, modifier = "") {
  const section = document.createElement("section");
  section.className = `report-insight${modifier ? ` report-insight--${modifier}` : ""}`;
  const heading = document.createElement("h3");
  heading.textContent = title;
  const paragraph = document.createElement("p");
  paragraph.textContent = text;
  section.appendChild(heading);
  section.appendChild(paragraph);
  return section;
}

function createReportListSection(title, items, modifier = "") {
  if (!items || !items.length) return null;
  const section = document.createElement("section");
  section.className = `report-insight${modifier ? ` report-insight--${modifier}` : ""}`;
  const heading = document.createElement("h3");
  heading.textContent = title;
  const list = document.createElement("ul");
  items.forEach((item) => {
    const li = document.createElement("li");
    li.textContent = String(item).replace(/_/g, " ");
    list.appendChild(li);
  });
  section.appendChild(heading);
  section.appendChild(list);
  return section;
}

function setReportResult(rawData, responseOk) {
  clearProgressFeedback();
  const result = normalizeBackendResult(rawData);
  const isOk = responseOk && result.status === "OK";
  const hasFile = Boolean(result.downloadUrl || result.workbookPath);
  const errors = isOk ? [] : uniqueStrings(result.errors);

  elements.analysisReport.className = "report-box";
  elements.analysisReport.classList.add(isOk ? "report-status--success" : "report-status--error");
  elements.analysisReport.innerHTML = "";
  const summary = document.createElement("div");
  summary.className = "report-summary report-summary--readable";
  const title = document.createElement("h3");
  title.className = "report-summary__title";
  title.textContent = isOk ? "Traitement terminé" : "Traitement à vérifier";
  summary.appendChild(title);
  summary.appendChild(createReportSection(
    "Résultat",
    isOk ? "La restructuration des fichiers déposés est terminée." : "La restructuration n'a pas pu aller au bout.",
    isOk ? "success" : "error",
  ));
  summary.appendChild(createReportSection(
    "Fichier de sortie",
    hasFile ? "Un fichier est disponible via le bouton de téléchargement." : "Aucun fichier téléchargeable n'a été produit pour le moment.",
    hasFile ? "success" : "warning",
  ));
  if (errors.length) summary.appendChild(createReportListSection("Points bloquants", errors, "error"));
  summary.appendChild(createReportSection(
    "Action conseillée",
    errors.length ? "Corrigez le point bloquant puis relancez le traitement." : "Téléchargez le fichier généré, puis contrôlez le classeur.",
    "next",
  ));
  elements.analysisReport.appendChild(summary);
  updateDownloadFromResult(result);
}

function resetDownloadState() {
  elements.filesXlsxStatus.textContent = "Aucun XLSX restructuration disponible.";
  elements.filesXlsxDownload.href = "#";
  elements.filesXlsxDownload.removeAttribute("data-download-href");
  elements.filesXlsxDownload.classList.add("btn--disabled");
  elements.filesXlsxDownload.setAttribute("aria-disabled", "true");
}

function updateDownloadFromResult(result) {
  const fallbackUrl = result.workbookPath ? `${API_BASE}/download?path=${encodeURIComponent(result.workbookPath)}` : "";
  const href = result.downloadUrl || fallbackUrl;
  setDownloadState(href ? "XLSX restructuré prêt : choisissez son nom au téléchargement." : "Aucun XLSX restructuration disponible.", href);
}

function setDownloadState(message, href = "") {
  elements.filesXlsxStatus.textContent = message;
  if (href) {
    setAnalysisAvailable(href);
    resetAnalysisDownload();
    elements.filesXlsxDownload.href = href;
    elements.filesXlsxDownload.dataset.downloadHref = href;
    elements.filesXlsxDownload.classList.remove("btn--disabled");
    elements.filesXlsxDownload.setAttribute("aria-disabled", "false");
  } else {
    elements.filesXlsxDownload.href = "#";
    elements.filesXlsxDownload.removeAttribute("data-download-href");
    elements.filesXlsxDownload.classList.add("btn--disabled");
    elements.filesXlsxDownload.setAttribute("aria-disabled", "true");
  }
}

function preferredDownloadName() {
  const lots = [...new Set([...state.dceFiles, ...state.actFiles].map(assignedLot).filter(Boolean))].sort();
  const suffix = lots.length ? `_lots_${lots.join("-")}` : "";
  return `tableau_restructure${suffix}.xlsx`;
}

function sanitizeDownloadName(name) {
  let cleaned = String(name || "").trim().replace(/[\\/:*?"<>|]/g, "_");
  if (!cleaned) cleaned = preferredDownloadName();
  if (!/\.xlsx$/i.test(cleaned)) cleaned = cleaned.replace(/\.(xlsm|xls)$/i, "") + ".xlsx";
  return cleaned;
}

async function confirmDownloadName(event) {
  const link = event.currentTarget;
  const href = link.dataset.downloadHref || link.href;
  if (!href || href.endsWith("#") || link.getAttribute("aria-disabled") === "true") {
    event.preventDefault();
    return;
  }

  event.preventDefault();
  const wantedName = window.prompt("Nom du fichier à télécharger", preferredDownloadName());
  if (wantedName === null) return;
  const finalName = sanitizeDownloadName(wantedName);

  const previousText = link.textContent;
  link.setAttribute("aria-busy", "true");
  link.textContent = "Préparation du téléchargement...";

  try {
    // Le fichier est récupéré en mémoire, puis téléchargé via une URL Blob.
    // Le nom est alors imposé par le navigateur depuis finalName, sans dépendre
    // du nom physique UUID utilisé par Flask sur le serveur.
    const response = await fetch(new URL(href, window.location.origin), {
      method: "GET",
      cache: "no-store",
    });
    if (!response.ok) {
      throw new Error(`Téléchargement impossible (${response.status}).`);
    }

    const blob = await response.blob();
    const objectUrl = URL.createObjectURL(blob);
    const temporaryLink = document.createElement("a");
    temporaryLink.href = objectUrl;
    temporaryLink.download = finalName;
    temporaryLink.style.display = "none";
    document.body.appendChild(temporaryLink);
    temporaryLink.click();
    temporaryLink.remove();
    setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
  } catch (error) {
    window.alert(error?.message || "Le téléchargement du fichier a échoué.");
  } finally {
    link.removeAttribute("aria-busy");
    link.textContent = previousText;
  }
}

function bindDropzone(dropzone, input, type) {
  dropzone.addEventListener("click", () => input.click());
  dropzone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      input.click();
    }
  });
  ["dragenter", "dragover"].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.add("is-dragover");
    });
  });
  ["dragleave", "dragend", "drop"].forEach((eventName) => {
    dropzone.addEventListener(eventName, () => dropzone.classList.remove("is-dragover"));
  });
  dropzone.addEventListener("drop", (event) => {
    event.preventDefault();
    setFiles(type, Array.from(event.dataTransfer.files || []));
  });
  input.addEventListener("change", (event) => {
    setFiles(type, Array.from(event.target.files || []));
    event.target.value = "";
  });
}

function setFiles(type, files) {
  const acceptedFiles = validateIncomingFiles(files);
  if (!acceptedFiles.length) return;
  acceptedFiles.forEach((file) => {
    const detected = extractLotNumber(file.name);
    if (detected) state.fileLots[fileKey(file)] = detected;
  });
  if (type === "dce") {
    state.dceFiles = mergeFiles(state.dceFiles, acceptedFiles);
    renderFiles(elements.dceFileList, state.dceFiles, "dce");
  } else if (type === "act") {
    state.actFiles = mergeFiles(state.actFiles, acceptedFiles);
    renderActFiles();
  }
  updateRunButtonState();
}

function removeSelectedFile(type, index) {
  if (type === "dce") {
    state.dceFiles.splice(index, 1);
    renderFiles(elements.dceFileList, state.dceFiles, "dce");
  } else if (type === "act") {
    const removed = state.actFiles[index];
    state.actFiles.splice(index, 1);
    state.actEnterpriseNames.splice(index, 1);
    if (removed) delete state.fileLots[fileKey(removed)];
    renderActFiles();
  }
  updateRunButtonState();
}

async function parseJsonResponse(response) {
  const text = await response.text();
  if (!text) return { status: response.ok ? "OK" : "ERROR" };
  try {
    return JSON.parse(text);
  } catch (_error) {
    return { status: "ERROR", errors: ["Échec de la connexion. Essayez de rafraîchir la page."] };
  }
}

async function postForm(endpoint, form) {
  const response = await fetch(`${API_BASE}${endpoint}`, { method: "POST", body: form });
  return { response, data: await parseJsonResponse(response) };
}

function lotMetadataPayload() {
  const build = (file, index, role) => ({
    index,
    role,
    original_name: file.name,
    lot_number: assignedLot(file),
  });
  return [
    ...state.dceFiles.map((file, index) => build(file, index, "dce")),
    ...state.actFiles.map((file, index) => build(file, index, "act")),
  ];
}


async function runByFiles() {
  if (!state.dceFiles.length || !state.actFiles.length) {
    setReportMessage("La restructuration exige au moins un fichier DCE Excel et un fichier ACT Excel.", "report-status--warning");
    return;
  }
  if (!validateDroppedFiles()) {
    window.alert("Traitement bloqué : vérifiez que tous les fichiers déposés sont des classeurs Excel.");
    setReportMessage("Les lots peuvent être détectés dans les feuilles par le serveur.", "report-status--warning");
    updateRunButtonState();
    return;
  }
  if (!validateEnterpriseNames()) {
    window.alert("Traitement bloqué : renseignez le nom de l’entreprise pour chaque fichier ACT.");
    setReportMessage("Renseignez le nom de l’entreprise pour chaque fichier ACT avant de lancer la restructuration.", "report-status--warning");
    updateRunButtonState();
    return;
  }

  setLoading(true);
  startProgressFeedback();
  resetDownloadState();
  try {
    const form = new FormData();
    state.dceFiles.forEach((file) => form.append("dce_files", file));
    state.actFiles.forEach((file) => form.append("act_files", file));
    form.append("dropped_file_lots", JSON.stringify(lotMetadataPayload()));
    form.append("act_enterprise_names", JSON.stringify(enterpriseNamePayload()));
    const { response, data } = await postForm("/run/files", form);
    setReportResult(data, response.ok);
  } catch (_error) {
    setReportMessage("Échec de la connexion. Essayez de rafraîchir la page.", "report-status--error");
    resetDownloadState();
  } finally {
    setLoading(false);
  }
}

function toggleHelp() {
  const isHidden = elements.helpPanel.classList.toggle("hidden");
  elements.helpPanel.setAttribute("aria-hidden", String(isHidden));
  elements.helpToggle.setAttribute("aria-expanded", String(!isHidden));
}




// ---------- Dropzone analyse autonome V3 ----------
let manualAnalysisFile = null;

function isAnalysisWorkbook(file) {
  return Boolean(file) && String(file.name || "").toLowerCase().endsWith(".xlsx");
}

function renderManualAnalysisFile() {
  if (!elements.analysisSelectedFile || !elements.analysisSelectedFileName) return;
  if (manualAnalysisFile) {
    elements.analysisSelectedFileName.textContent = manualAnalysisFile.name;
    elements.analysisSelectedFile.classList.remove("hidden");
    if (elements.offerAnalysisStatus) {
      elements.offerAnalysisStatus.textContent = "Classeur restructuré déposé. Vous pouvez lancer l'analyse.";
    }
  } else {
    elements.analysisSelectedFileName.textContent = "";
    elements.analysisSelectedFile.classList.add("hidden");
    if (elements.offerAnalysisStatus) {
      elements.offerAnalysisStatus.textContent = restructuredWorkbookHref
        ? "Le classeur restructuré est prêt. Vous pouvez lancer l'analyse."
        : "Restructurez les fichiers ou déposez directement un classeur restructuré.";
    }
  }
  if (elements.runOfferAnalysisButton) {
    elements.runOfferAnalysisButton.disabled = !(manualAnalysisFile || restructuredWorkbookHref);
  }
}

function acceptManualAnalysisFile(files) {
  const candidates = Array.from(files || []);
  if (!candidates.length) return;
  if (candidates.length > 1) {
    window.alert("Déposez un seul classeur restructuré à analyser.");
    return;
  }
  const file = candidates[0];
  if (!isAnalysisWorkbook(file)) {
    window.alert(`${file.name} : le fichier d'analyse doit être au format .xlsx.`);
    return;
  }
  manualAnalysisFile = file;
  resetAnalysisDownload();
  renderManualAnalysisFile();
}

function bindAnalysisDropzone() {
  if (!elements.analysisDropzone || !elements.analysisInput) return;
  const dropzone = elements.analysisDropzone;
  const input = elements.analysisInput;
  dropzone.addEventListener("click", () => input.click());
  dropzone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      input.click();
    }
  });
  ["dragenter", "dragover"].forEach((name) => {
    dropzone.addEventListener(name, (event) => {
      event.preventDefault();
      dropzone.classList.add("is-dragover");
    });
  });
  ["dragleave", "dragend", "drop"].forEach((name) => {
    dropzone.addEventListener(name, () => dropzone.classList.remove("is-dragover"));
  });
  dropzone.addEventListener("drop", (event) => {
    event.preventDefault();
    acceptManualAnalysisFile(event.dataTransfer.files);
  });
  input.addEventListener("change", (event) => {
    acceptManualAnalysisFile(event.target.files);
    event.target.value = "";
  });
  elements.removeAnalysisFileButton?.addEventListener("click", (event) => {
    event.stopPropagation();
    manualAnalysisFile = null;
    renderManualAnalysisFile();
  });
  renderManualAnalysisFile();
}



// STABILITY_FIX_AFTER_RESTORE_V1
let analysisFeedbackTimer = null;
let analysisFeedbackStartedAt = 0;
let analysisFeedbackStep = 0;
const ANALYSIS_FEEDBACK_STEPS = [
  "Préparation du classeur restructuré.",
  "Envoi du classeur au moteur d’analyse.",
  "Analyse des feuilles et des blocs entreprises.",
  "Vérification des quantités, unités, prix et montants.",
  "Application des couleurs, commentaires et synthèses.",
  "Finalisation du classeur analysé.",
];

function formatElapsed(seconds) {
  const minutes = Math.floor(seconds / 60);
  const rest = seconds % 60;
  return `${String(minutes).padStart(2, "0")}:${String(rest).padStart(2, "0")}`;
}

function renderAnalysisFeedback() {
  if (!elements.analysisReport) return;
  const elapsed = Math.max(0, Math.floor((Date.now() - analysisFeedbackStartedAt) / 1000));
  const step = ANALYSIS_FEEDBACK_STEPS[analysisFeedbackStep % ANALYSIS_FEEDBACK_STEPS.length];
  elements.analysisReport.className = "report-box report-status--warning report-box--progress";
  elements.analysisReport.innerHTML = "";
  const title = document.createElement("h3");
  title.textContent = "Analyse des offres en cours";
  const current = document.createElement("p");
  current.innerHTML = `<strong>Étape indicative :</strong> ${step}`;
  const duration = document.createElement("p");
  duration.textContent = `Analyse active depuis ${formatElapsed(elapsed)}. Ne fermez pas cette page.`;
  const note = document.createElement("p");
  note.className = "progress-panel__intro";
  note.textContent = "La progression est indicative : le serveur renverra le rapport complet après l’enregistrement du classeur.";
  elements.analysisReport.append(title, current, duration, note);
}

function startAnalysisFeedback() {
  stopAnalysisFeedback();
  analysisFeedbackStartedAt = Date.now();
  analysisFeedbackStep = 0;
  renderAnalysisFeedback();
  analysisFeedbackTimer = window.setInterval(() => {
    const elapsed = Math.floor((Date.now() - analysisFeedbackStartedAt) / 1000);
    analysisFeedbackStep = Math.floor(elapsed / 12);
    renderAnalysisFeedback();
  }, 1000);
}

function stopAnalysisFeedback() {
  if (analysisFeedbackTimer) {
    window.clearInterval(analysisFeedbackTimer);
    analysisFeedbackTimer = null;
  }
}

// ---------- Analyse des offres frontend ----------
let restructuredWorkbookHref = "";
let analysisWorkbookHref = "";

function setAnalysisAvailable(href) {
  restructuredWorkbookHref = href || "";
  if (!elements.runOfferAnalysisButton) return;
  elements.runOfferAnalysisButton.disabled = !restructuredWorkbookHref;
  if (elements.offerAnalysisStatus) {
    elements.offerAnalysisStatus.textContent = restructuredWorkbookHref
      ? "Le classeur restructuré est prêt. Vous pouvez lancer l'analyse."
      : "Restructurez d'abord les fichiers pour activer l'analyse.";
  }
}

function resetAnalysisDownload() {
  analysisWorkbookHref = "";
  if (!elements.analysisXlsxDownload) return;
  elements.analysisXlsxDownload.href = "#";
  elements.analysisXlsxDownload.classList.add("btn--disabled");
  elements.analysisXlsxDownload.setAttribute("aria-disabled", "true");
}

async function runOfferAnalysis() {
  if ((!manualAnalysisFile && !restructuredWorkbookHref) || !elements.runOfferAnalysisButton) return;
  const button = elements.runOfferAnalysisButton;
  const previousText = button.textContent;
  button.disabled = true;
  button.textContent = "Analyse en cours...";
  resetAnalysisDownload();
  startAnalysisFeedback();
  if (elements.offerAnalysisStatus) {
    elements.offerAnalysisStatus.textContent = "Lecture du classeur restructuré et application des contrôles métier...";
  }

  try {
    let workbookFile = manualAnalysisFile;
    if (!workbookFile) {
      const workbookResponse = await fetch(new URL(restructuredWorkbookHref, window.location.origin), {
        cache: "no-store",
      });
      if (!workbookResponse.ok) {
        throw new Error(`Impossible de récupérer le classeur restructuré (${workbookResponse.status}).`);
      }
      const workbookBlob = await workbookResponse.blob();
      workbookFile = new File([workbookBlob], "classeur_restructure.xlsx", {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
    }
    const form = new FormData();
    form.append("workbook", workbookFile, workbookFile.name);

    const response = await fetch(`${API_BASE}/run/offer-analysis`, {
      method: "POST",
      body: form,
    });
    const data = await parseJsonResponse(response);
    if (!response.ok || data.status !== "OK") {
      const errors = Array.isArray(data.errors) ? data.errors.join("\n") : "L'analyse a échoué.";
      throw new Error(errors);
    }

    analysisWorkbookHref = data.download_url || "";
    if (!analysisWorkbookHref) {
      throw new Error("Le backend n'a retourné aucun fichier analysé.");
    }
    elements.analysisXlsxDownload.href = analysisWorkbookHref;
    elements.analysisXlsxDownload.classList.remove("btn--disabled");
    elements.analysisXlsxDownload.setAttribute("aria-disabled", "false");

    const report = Array.isArray(data.analysis_report) ? data.analysis_report : [];
    const analyzedSheets = report.filter((item) => item.status === "OK").length;
    const issueCount = report.reduce((sum, item) => sum + Number(item.issues || 0), 0);
    stopAnalysisFeedback();
    if (elements.offerAnalysisStatus) {
      elements.offerAnalysisStatus.textContent =
        `Analyse terminée : ${analyzedSheets} feuille(s), ${issueCount} alerte(s) métier détectée(s).`;
    }
    setReportMessage(`Analyse terminée : ${analyzedSheets} feuille(s), ${issueCount} alerte(s) métier. Le classeur analysé est disponible.`, "report-status--success");
    appendContextualRegistryDiagnostics(report);
  } catch (error) {
    stopAnalysisFeedback();
    if (elements.offerAnalysisStatus) {
      elements.offerAnalysisStatus.textContent = error?.message || "L'analyse a échoué.";
    }
    window.alert(error?.message || "L'analyse a échoué.");
  } finally {
    stopAnalysisFeedback();
    button.disabled = !(manualAnalysisFile || restructuredWorkbookHref);
    button.textContent = previousText;
  }
}

async function downloadAnalyzedWorkbook(event) {
  event.preventDefault();
  if (!analysisWorkbookHref || event.currentTarget.getAttribute("aria-disabled") === "true") return;
  const wanted = window.prompt("Nom du fichier analysé à télécharger", "analyse_offres.xlsx");
  if (wanted === null) return;
  const finalName = sanitizeDownloadName(wanted);
  try {
    const response = await fetch(new URL(analysisWorkbookHref, window.location.origin), { cache: "no-store" });
    if (!response.ok) throw new Error(`Téléchargement impossible (${response.status}).`);
    const blob = await response.blob();
    const objectUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = objectUrl;
    link.download = finalName;
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
  } catch (error) {
    window.alert(error?.message || "Téléchargement impossible.");
  }
}

function init() {
  bindDropzone(elements.dceDropzone, elements.dceInput, "dce");
  bindDropzone(elements.actDropzone, elements.actInput, "act");
  renderFiles(elements.dceFileList, [], "dce");
  renderActFiles();
  updateRunButtonState();
  resetDownloadState();
  elements.helpToggle.addEventListener("click", toggleHelp);
  elements.runFilesButton.addEventListener("click", runByFiles);
  elements.filesXlsxDownload.addEventListener("click", confirmDownloadName);
  elements.runOfferAnalysisButton?.addEventListener("click", runOfferAnalysis);
  elements.analysisXlsxDownload?.addEventListener("click", downloadAnalyzedWorkbook);
  bindAnalysisDropzone();
}

document.addEventListener("DOMContentLoaded", init);

// TCE_VIRTUAL_SHEETS_V22

// BEGIN_CONTEXTUAL_REGISTRY_V2
function appendContextualRegistryDiagnostics(report) {
  if (!elements.analysisReport || !Array.isArray(report) || !report.length) return;
  const enabled = report.filter((item) => item.contextual_registry_enabled === true);
  const failed = report.filter((item) => item.contextual_registry_enabled === false);
  if (!enabled.length && !failed.length) return;
  const total = (key) => enabled.reduce((sum, item) => sum + Number(item.contextual_registry?.[key] || 0), 0);
  const wrapper = document.createElement("details");
  wrapper.className = "contextual-diagnostic";
  const summary = document.createElement("summary");
  const unique = total("unique_records");
  const duplicates = total("duplicates_ignored");
  summary.textContent = `Diagnostic contextuel V2 : ${enabled.length}/${report.length} feuille(s), ${unique} anomalie(s) unique(s), ${duplicates} doublon(s) ignoré(s)`;
  const intro = document.createElement("p");
  intro.textContent = "Contrôle parallèle en lecture seule : aucune cellule, formule, couleur, légende ou synthèse Excel n’est modifiée.";
  wrapper.append(summary, intro);
  enabled.forEach((item) => {
    const d = item.contextual_registry || {};
    const line = document.createElement("p");
    line.className = "contextual-diagnostic__line";
    line.textContent = `${item.sheet} : ${Number(d.unique_records || 0)} unique(s), ${Number(d.duplicates_ignored || 0)} doublon(s), couverture bâtiment ${Number(d.scope_coverage_percent || 0).toLocaleString("fr-FR")} %.`;
    wrapper.appendChild(line);
  });
  failed.forEach((item) => {
    const line = document.createElement("p");
    line.className = "contextual-diagnostic__error";
    line.textContent = `${item.sheet || "Feuille"} : repli historique actif (${item.contextual_registry_fallback || "diagnostic indisponible"}).`;
    wrapper.appendChild(line);
  });
  elements.analysisReport.appendChild(wrapper);
}
// END_CONTEXTUAL_REGISTRY_V2
