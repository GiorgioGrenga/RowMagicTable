const MAX_FILE_SIZE = 15 * 1024 * 1024;
const state = {
  documentFile: null,
  workbookFile: null,
  originalTable: null,
  headers: [],
  rows: [],
  workbook: null,
  mapping: [],
};

const documentInput = document.querySelector("#document-input");
const workbookInput = document.querySelector("#workbook-input");
const extractButton = document.querySelector("#extract-button");
const resultsSection = document.querySelector("#results-section");
const mergeSection = document.querySelector("#merge-section");
const toast = document.querySelector("#toast");

function showToast(message, isError = false) {
  toast.textContent = message;
  toast.classList.toggle("error-toast", isError);
  toast.classList.add("visible");
  window.clearTimeout(showToast.timeout);
  showToast.timeout = window.setTimeout(() => toast.classList.remove("visible"), 4200);
}

function setFile(kind, file) {
  const isDocument = kind === "document";
  const error = document.querySelector("#input-error");
  error.hidden = true;
  if (!file) return;
  const supported = isDocument
    ? /\.(png|jpe?g|pdf)$/i.test(file.name)
    : /\.xlsx$/i.test(file.name);
  if (!supported) {
    error.textContent = isDocument
      ? "Scegli un’immagine PNG, JPG o un documento PDF."
      : "Scegli un file Excel nel formato .xlsx.";
    error.hidden = false;
    return;
  }
  if (file.size > MAX_FILE_SIZE) {
    error.textContent = "Il file supera il limite di 15 MB.";
    error.hidden = false;
    return;
  }
  resultsSection.hidden = true;
  mergeSection.hidden = true;
  state.workbook = null;
  state.originalTable = null;
  if (isDocument) state.documentFile = file;
  else state.workbookFile = file;
  const output = document.querySelector(isDocument ? "#document-selected" : "#workbook-selected");
  output.textContent = `✓  ${file.name} · ${formatSize(file.size)}`;
  output.hidden = false;
  const dropzone = document.querySelector(isDocument ? "#document-drop" : "#workbook-drop");
  dropzone.classList.add("file-loaded");
  document.querySelector(isDocument ? "#remove-document" : "#remove-workbook").hidden = false;
  if (!isDocument) document.querySelector("#workbook-drop-main").textContent = "File Excel selezionato";
  extractButton.disabled = !state.documentFile;
}

function removeFile(kind) {
  const isDocument = kind === "document";
  if (isDocument) {
    state.documentFile = null;
    documentInput.value = "";
  } else {
    state.workbookFile = null;
    workbookInput.value = "";
  }
  document.querySelector(isDocument ? "#document-selected" : "#workbook-selected").hidden = true;
  document.querySelector(isDocument ? "#remove-document" : "#remove-workbook").hidden = true;
  document.querySelector(isDocument ? "#document-drop" : "#workbook-drop").classList.remove("file-loaded");
  if (!isDocument) document.querySelector("#workbook-drop-main").textContent = "Trascina qui il tuo Excel";
  resultsSection.hidden = true;
  mergeSection.hidden = true;
  state.workbook = null;
  state.originalTable = null;
  extractButton.disabled = !state.documentFile;
}

function formatSize(bytes) {
  return bytes < 1024 * 1024
    ? `${Math.max(1, Math.round(bytes / 1024))} KB`
    : `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

documentInput.addEventListener("change", () => setFile("document", documentInput.files[0]));
workbookInput.addEventListener("change", () => setFile("workbook", workbookInput.files[0]));
document.querySelector("#remove-document").addEventListener("click", () => removeFile("document"));
document.querySelector("#remove-workbook").addEventListener("click", () => removeFile("workbook"));

for (const [input, dropzone, kind] of [
  [documentInput, document.querySelector("#document-drop"), "document"],
  [workbookInput, document.querySelector("#workbook-drop"), "workbook"],
]) {
  dropzone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      input.click();
    }
  });
  for (const eventName of ["dragenter", "dragover"]) {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.add("drag-over");
    });
  }
  for (const eventName of ["dragleave", "drop"]) {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.remove("drag-over");
    });
  }
  dropzone.addEventListener("drop", (event) => {
    const file = event.dataTransfer.files[0];
    if (file) setFile(kind, file);
  });
}

function updateRowsCounter() {
  const rowCount = state.rows.length;
  document.querySelector("#row-count").textContent = `${rowCount} ${rowCount === 1 ? "riga" : "righe"}`;
  document.querySelector("#table-summary").textContent =
    `${rowCount} ${rowCount === 1 ? "riga" : "righe"} · ${state.headers.length} ${state.headers.length === 1 ? "colonna" : "colonne"}`;
}

function renderTable() {
  const table = document.querySelector("#data-table");
  table.replaceChildren();
  const thead = document.createElement("thead");
  const headingRow = document.createElement("tr");
  for (let index = 0; index < state.headers.length; index += 1) {
    const th = document.createElement("th");
    const input = document.createElement("input");
    input.value = state.headers[index];
    input.setAttribute("aria-label", `Intestazione colonna ${index + 1}`);
    input.addEventListener("input", () => {
      state.headers[index] = input.value;
      if (state.workbook) renderMapping();
    });
    th.append(input);
    headingRow.append(th);
  }
  const deleteHeading = document.createElement("th");
  deleteHeading.className = "delete-heading";
  deleteHeading.setAttribute("aria-label", "Rimuovi riga");
  headingRow.append(deleteHeading);
  thead.append(headingRow);
  table.append(thead);

  const tbody = document.createElement("tbody");
  state.rows.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    for (let columnIndex = 0; columnIndex < state.headers.length; columnIndex += 1) {
      const td = document.createElement("td");
      const input = document.createElement("input");
      input.value = row[columnIndex] == null ? "" : String(row[columnIndex]);
      input.setAttribute("aria-label", `${state.headers[columnIndex]}, riga ${rowIndex + 1}`);
      input.addEventListener("input", () => {
        const value = input.value;
        if (value.trim() === "") row[columnIndex] = null;
        else if (/^-?\d+(?:[.,]\d+)?$/.test(value.trim())) {
          row[columnIndex] = Number(value.trim().replace(",", "."));
        } else row[columnIndex] = value;
      });
      td.append(input);
      tr.append(td);
    }
    const actionCell = document.createElement("td");
    actionCell.className = "delete-cell";
    const deleteButton = document.createElement("button");
    deleteButton.type = "button";
    deleteButton.className = "text-button row-delete";
    deleteButton.textContent = "×";
    deleteButton.setAttribute("aria-label", `Elimina riga ${rowIndex + 1}`);
    deleteButton.addEventListener("click", () => {
      state.rows.splice(rowIndex, 1);
      renderTable();
    });
    actionCell.append(deleteButton);
    tr.append(actionCell);
    tbody.append(tr);
  });
  table.append(tbody);
  updateRowsCounter();
}

function initializeMapping(sheet) {
  const headers = sheet.headers || [];
  state.mapping = state.headers.map((source) => {
    const match = headers.find((target) => target.trim().toLocaleLowerCase() === source.trim().toLocaleLowerCase());
    return match
      ? { mode: "map", target: match, name: source }
      : { mode: "add", target: "", name: source };
  });
}

function selectedSheet() {
  const select = document.querySelector("#sheet-select");
  return state.workbook?.sheets.find((sheet) => sheet.name === select.value);
}

function renderMapping() {
  const sheet = selectedSheet();
  if (!sheet) return;
  const preview = document.querySelector("#workbook-preview");
  preview.replaceChildren();
  if (sheet.headers.length && sheet.preview.length) {
    const previewTitle = document.createElement("div");
    previewTitle.className = "workbook-preview-title";
    previewTitle.textContent = "Un’anteprima delle righe già presenti";
    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const heading = document.createElement("tr");
    sheet.headers.forEach((header) => {
      const th = document.createElement("th");
      th.textContent = header;
      heading.append(th);
    });
    thead.append(heading);
    table.append(thead);
    const tbody = document.createElement("tbody");
    sheet.preview.forEach((previewRow) => {
      const tr = document.createElement("tr");
      previewRow.forEach((value) => {
        const td = document.createElement("td");
        td.textContent = value;
        tr.append(td);
      });
      tbody.append(tr);
    });
    table.append(tbody);
    preview.append(previewTitle, table);
    preview.hidden = false;
  } else preview.hidden = true;
  const rows = document.querySelector("#mapping-rows");
  rows.replaceChildren();
  const matches = state.mapping.filter((item) => item.mode === "map").length;
  const added = state.mapping.filter((item) => item.mode === "add").length;
  document.querySelector("#match-summary").innerHTML =
    `<strong>${matches} ${matches === 1 ? "intestazione abbinata" : "intestazioni abbinate"}</strong>` +
    ` · ${added} ${added === 1 ? "nuova colonna" : "nuove colonne"}` +
    (sheet.empty ? " · foglio vuoto: verrà creata la riga d’intestazione" : "");

  state.headers.forEach((source, index) => {
    const row = document.createElement("tr");
    const sourceCell = document.createElement("td");
    const sourceLabel = document.createElement("div");
    sourceLabel.className = "source-column";
    const badge = document.createElement("span");
    badge.className = "column-type";
    badge.textContent = "Aa";
    const name = document.createElement("span");
    name.textContent = source || `Colonna ${index + 1}`;
    sourceLabel.append(badge, name);
    sourceCell.append(sourceLabel);

    const actionCell = document.createElement("td");
    const controls = document.createElement("div");
    controls.className = "mapping-controls";
    const select = document.createElement("select");
    select.className = "mapping-select";
    select.setAttribute("aria-label", `Destinazione per ${source || `colonna ${index + 1}`}`);
    const options = [
      { value: "add", text: "＋  Aggiungi come nuova colonna" },
      ...sheet.headers.map((header) => ({ value: `map:${header}`, text: `↔  ${header}` })),
      { value: "skip", text: "—  Non importare" },
    ];
    for (const optionData of options) {
      const option = document.createElement("option");
      option.value = optionData.value;
      option.textContent = optionData.text;
      select.append(option);
    }
    const current = state.mapping[index];
    select.value = current.mode === "map" ? `map:${current.target}` : current.mode;
    select.addEventListener("change", () => {
      if (select.value === "add") state.mapping[index] = { mode: "add", name: state.headers[index] };
      else if (select.value === "skip") state.mapping[index] = { mode: "skip" };
      else state.mapping[index] = { mode: "map", target: select.value.slice(4) };
      renderMapping();
    });
    controls.append(select);
    if (current.mode === "add") {
      const input = document.createElement("input");
      input.className = "mapping-name";
      input.value = current.name || "";
      input.placeholder = "Nome nuova colonna";
      input.setAttribute("aria-label", `Nome nuova colonna per ${source}`);
      input.addEventListener("input", () => { current.name = input.value; });
      controls.append(input);
    }
    actionCell.append(controls);
    row.append(sourceCell, actionCell);
    rows.append(row);
  });
}

function setupWorkbook(workbook) {
  state.workbook = workbook;
  const select = document.querySelector("#sheet-select");
  select.replaceChildren();
  workbook.sheets.forEach((sheet) => {
    const option = document.createElement("option");
    option.value = sheet.name;
    option.textContent = `${sheet.name}${sheet.headers.length ? ` · ${sheet.headers.length} colonne` : " · vuoto"}`;
    select.append(option);
  });
  document.querySelector("#merge-filename").textContent = state.workbookFile.name;
  document.querySelector("#merge-sheet-count").textContent =
    `${workbook.sheets.length} ${workbook.sheets.length === 1 ? "foglio disponibile" : "fogli disponibili"}`;
  select.addEventListener("change", () => {
    initializeMapping(selectedSheet());
    renderMapping();
  });
  initializeMapping(workbook.sheets[0]);
  renderMapping();
  mergeSection.hidden = false;
  document.querySelector("#export-description").textContent = "Le nuove righe saranno aggiunte al file Excel selezionato.";
}

extractButton.addEventListener("click", async () => {
  if (!state.documentFile) return;
  const error = document.querySelector("#input-error");
  error.hidden = true;
  extractButton.disabled = true;
  extractButton.innerHTML = '<span class="loading-spinner" aria-hidden="true"></span><span>Leggo la tabella…</span>';
  const formData = new FormData();
  formData.append("document_file", state.documentFile);
  if (state.workbookFile) formData.append("workbook_file", state.workbookFile);
  try {
    const response = await fetch("/api/extract", { method: "POST", body: formData });
    const result = await response.json();
    if (!response.ok) throw new Error(result.detail || "Non riesco a leggere questo file.");
    state.originalTable = structuredClone(result.table);
    state.headers = [...result.table.headers];
    state.rows = result.table.rows.map((row) => [...row]);
    renderTable();
    if (result.workbook) setupWorkbook(result.workbook);
    else {
      state.workbook = null;
      mergeSection.hidden = true;
      document.querySelector("#export-description").textContent = "Scarica la tabella in formato Excel.";
    }
    resultsSection.hidden = false;
    resultsSection.scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (requestError) {
    error.textContent = requestError.message || "Si è verificato un errore. Riprova.";
    error.hidden = false;
  } finally {
    extractButton.disabled = !state.documentFile;
    extractButton.innerHTML = '<span>Estrai la tabella</span><span class="button-arrow">→</span>';
  }
});

document.querySelector("#add-row-button").addEventListener("click", () => {
  state.rows.push(state.headers.map(() => null));
  renderTable();
  document.querySelector("#data-table tbody tr:last-child input")?.focus();
});

document.querySelector("#reset-table-button").addEventListener("click", () => {
  if (!state.originalTable) return;
  state.headers = [...state.originalTable.headers];
  state.rows = state.originalTable.rows.map((row) => [...row]);
  if (state.workbook) initializeMapping(selectedSheet());
  renderTable();
  if (state.workbook) renderMapping();
  showToast("Tabella ripristinata come nell’estrazione iniziale.");
});

document.querySelector("#export-button").addEventListener("click", async (event) => {
  const button = event.currentTarget;
  const error = document.querySelector("#export-error");
  error.hidden = true;
  const table = { headers: state.headers, rows: state.rows };
  const mapping = state.mapping.map((item) => ({ ...item }));
  if (state.workbook && !mapping.some((item) => item.mode !== "skip")) {
    error.textContent = "Scegli almeno una colonna da aggiungere al foglio.";
    error.hidden = false;
    return;
  }
  button.disabled = true;
  button.innerHTML = '<span class="loading-spinner" aria-hidden="true"></span><span>Creo il file…</span>';
  const formData = new FormData();
  formData.append("table_json", JSON.stringify(table));
  formData.append("filename", state.documentFile?.name || "tabella.xlsx");
  if (state.workbook) {
    formData.append("workbook_file", state.workbookFile);
    formData.append("sheet_name", document.querySelector("#sheet-select").value);
    formData.append("mapping_json", JSON.stringify(mapping));
  }
  try {
    const response = await fetch("/api/export", { method: "POST", body: formData });
    if (!response.ok) {
      const result = await response.json();
      throw new Error(result.detail || "Non riesco a creare il file Excel.");
    }
    const blob = await response.blob();
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${(state.documentFile?.name || "tabella").replace(/\.[^.]+$/, "")}.xlsx`;
    document.body.append(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(link.href);
    showToast("Il tuo Excel è pronto per il download.");
  } catch (requestError) {
    error.textContent = requestError.message || "Si è verificato un errore. Riprova.";
    error.hidden = false;
  } finally {
    button.disabled = false;
    button.innerHTML = '<span>Scarica Excel</span><span class="download-icon">↓</span>';
  }
});
