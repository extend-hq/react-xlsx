import { createXlsxViewer, setWasmSource } from "@extend-ai/react-xlsx/vanilla";
import {
  createToast,
  hydrateIcons,
  icon,
  keepSelection,
  setupFileDrop,
  setupMenu,
  setupTabs,
  setupTheme,
  setupTooltips,
  updateSlider,
} from "./ui.js";

setWasmSource("/duke_sheets_wasm_bg.wasm");

const SAMPLE_URL = "/examples/crazy-chart-zoo-subtypes.xlsx";
const SAMPLE_NAME = "crazy-chart-zoo-subtypes.xlsx";
const FONT_FAMILIES = ["Aptos", "Calibri", "Arial", "Georgia", "Times New Roman", "Courier New"];
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];
const NUMBER_FORMATS = [
  { label: "General", sample: "abc", style: { numberFormat: { formatType: "general" } } },
  { label: "Number", sample: "1,234.00", style: { numberFormat: { formatType: "custom", formatString: "#,##0.00" } } },
  { label: "Currency", sample: "$1,234.00", style: { numberFormat: { formatType: "custom", formatString: "$#,##0.00" } } },
  { label: "Percentage", sample: "12.00%", style: { numberFormat: { formatType: "custom", formatString: "0.00%" } } },
  { label: "Short Date", sample: "1/31/2026", style: { numberFormat: { formatType: "custom", formatString: "m/d/yyyy" } } },
  { label: "Time", sample: "1:30 PM", style: { numberFormat: { formatType: "custom", formatString: "h:mm AM/PM" } } },
];
const STYLE_ACTIONS = {
  bold: { font: { bold: true } },
  italic: { font: { italic: true } },
  underline: { font: { underline: "single" } },
  strike: { font: { strikethrough: true } },
  "align-left": { alignment: { horizontal: "left" } },
  "align-center": { alignment: { horizontal: "center" } },
  "align-right": { alignment: { horizontal: "right" } },
  wrap: { alignment: { wrapText: true } },
  currency: NUMBER_FORMATS[2].style,
  percent: NUMBER_FORMATS[3].style,
  comma: { numberFormat: { formatType: "custom", formatString: "#,##0" } },
};

const byId = (id) => document.getElementById(id);
const fileInput = byId("file-input");
const readOnlyInput = byId("read-only");
const darkSheetInput = byId("dark-sheet");
const fontFamilySelect = byId("font-family");
const fontSizeSelect = byId("font-size");
const formulaInput = byId("formula-input");
const zoomSlider = byId("zoom-slider");
const showToast = createToast();

hydrateIcons();
setupTabs();
setupTooltips();
setupTheme(byId("theme-toggle"));
for (const element of document.querySelectorAll("[data-keep]")) keepSelection(element);
for (const family of FONT_FAMILIES) fontFamilySelect.append(new Option(family, family));
for (const size of FONT_SIZES) fontSizeSelect.append(new Option(String(size), String(size)));
fontSizeSelect.value = "11";

let loadSequence = 0;

/* ---------------------------- Viewer states ----------------------------- */

function createEmptyState() {
  const element = document.createElement("div");
  element.className = "empty-state";
  element.innerHTML = `
    <div class="empty-icon">${icon("file-spreadsheet")}</div>
    <div>
      <div class="empty-title">Open a workbook</div>
      <p class="empty-text">Drop an .xlsx file anywhere, or pick one to get started.</p>
    </div>
    <div class="empty-actions">
      <button class="btn btn-primary" type="button" data-action="open">${icon("upload")}Upload file</button>
      <button class="btn btn-outline" type="button" data-action="sample">${icon("sparkles")}Try the sample</button>
    </div>`;
  element.querySelector('[data-action="open"]').addEventListener("click", () => fileInput.click());
  element.querySelector('[data-action="sample"]').addEventListener("click", loadSample);
  return element;
}

function createLoadingState() {
  const element = document.createElement("div");
  element.className = "loading-state";
  element.innerHTML = '<span class="spinner"></span>';
  return element;
}

const viewer = createXlsxViewer(byId("viewer"), {
  src: SAMPLE_URL,
  fileName: SAMPLE_NAME,
  height: "100%",
  rounded: false,
  showDefaultToolbar: false,
  selectionColor: "#107C41",
  readOnlyAboveBytes: 5 * 1024 * 1024,
  emptyState: createEmptyState(),
  loadingState: createLoadingState(),
  renderChartLoading: () => createLoadingState(),
});

const current = () => viewer.getController();

/* ------------------------------ Formatting ------------------------------ */

function canEdit(controller = current()) {
  return Boolean(controller?.activeCellAddress) && !readOnlyInput.checked && !controller.readOnly;
}

function applyStyle(style) {
  const controller = current();
  if (!controller || !canEdit(controller)) return;
  if (controller.selection) controller.setRangeStyle(controller.selection, style);
  else controller.setSelectedCellStyle(style);
}

for (const button of document.querySelectorAll("[data-style]")) {
  button.addEventListener("click", () => applyStyle(STYLE_ACTIONS[button.dataset.style]));
}
fontFamilySelect.addEventListener("change", () => applyStyle({ font: { name: fontFamilySelect.value } }));
fontSizeSelect.addEventListener("change", () => applyStyle({ font: { size: Number(fontSizeSelect.value) } }));
byId("undo").addEventListener("click", () => current()?.undo());
byId("redo").addEventListener("click", () => current()?.redo());

const mergeDisabled = () => !current()?.selection || readOnlyInput.checked || Boolean(current()?.readOnly);
setupMenu(byId("merge"), [
  { label: "Merge &amp; Center", onSelect: () => current()?.mergeSelection("center"), disabled: mergeDisabled },
  { label: "Merge Across", onSelect: () => current()?.mergeSelection("across"), disabled: mergeDisabled },
  { label: "Merge Cells", onSelect: () => current()?.mergeSelection(), disabled: mergeDisabled },
  "separator",
  { label: "Unmerge Cells", onSelect: () => current()?.unmergeSelection(), disabled: mergeDisabled },
]);
setupMenu(
  byId("number-format"),
  NUMBER_FORMATS.map((format) => ({ label: format.label, hint: format.sample, onSelect: () => applyStyle(format.style) })),
);
setupMenu(byId("download"), [
  { label: "Excel workbook (.xlsx)", onSelect: () => current()?.exportXlsx(), disabled: () => !current()?.canExport },
  { label: "CSV, current sheet (.csv)", onSelect: () => current()?.exportCsv(), disabled: () => !current()?.canExport },
  "separator",
  { label: "Original file", onSelect: () => current()?.download(), disabled: () => !current()?.canDownload },
], { align: "end" });

/* ------------------------------ Formula bar ----------------------------- */

let formulaFocused = false;
let formulaInitial = "";

function commitFormula() {
  const controller = current();
  if (controller && !formulaInput.disabled && formulaInput.value !== formulaInitial) {
    controller.setSelectedFormula(formulaInput.value);
  }
  formulaFocused = false;
  syncFormulaButtons();
}

function cancelFormula() {
  formulaInput.value = formulaInitial;
  formulaFocused = false;
  syncFormulaButtons();
}

function syncFormulaButtons() {
  byId("formula-cancel").disabled = !formulaFocused;
  byId("formula-enter").disabled = !formulaFocused;
}

formulaInput.addEventListener("focus", () => {
  formulaInitial = formulaInput.value;
  formulaFocused = true;
  syncFormulaButtons();
});
formulaInput.addEventListener("blur", () => formulaFocused && commitFormula());
formulaInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    commitFormula();
    formulaInput.blur();
  } else if (event.key === "Escape") {
    event.preventDefault();
    cancelFormula();
    formulaInput.blur();
  }
});
for (const [id, action] of [["formula-cancel", cancelFormula], ["formula-enter", commitFormula]]) {
  const button = byId(id);
  button.addEventListener("mousedown", (event) => {
    event.preventDefault();
    action();
    formulaInput.blur();
  });
}

/* ------------------------------- Sheet tabs ----------------------------- */

const sheetList = byId("sheet-list");
let sheetSignature = "";

function renderSheetTabs(controller) {
  const signature = `${controller.activeSheetIndex}|${controller.sheets.map((sheet) => sheet.name).join("\u0000")}`;
  if (signature === sheetSignature) return;
  sheetSignature = signature;
  sheetList.replaceChildren(
    ...controller.sheets.map((sheet, index) => {
      const tab = document.createElement("button");
      tab.type = "button";
      tab.className = "sheet-tab";
      tab.setAttribute("role", "tab");
      tab.setAttribute("aria-selected", String(index === controller.activeSheetIndex));
      tab.textContent = sheet.name;
      tab.addEventListener("click", () => current()?.setActiveSheetIndex(index));
      return tab;
    }),
  );
  sheetList.querySelector('[aria-selected="true"]')?.scrollIntoView({ block: "nearest", inline: "nearest" });
}

byId("sheet-prev").addEventListener("click", () => {
  const controller = current();
  if (controller) controller.setActiveSheetIndex(controller.activeSheetIndex - 1);
});
byId("sheet-next").addEventListener("click", () => {
  const controller = current();
  if (controller) controller.setActiveSheetIndex(controller.activeSheetIndex + 1);
});
byId("sheet-add").addEventListener("click", () => current()?.addSheet());

/* --------------------------------- Zoom --------------------------------- */

byId("zoom-in").addEventListener("click", () => current()?.zoomIn());
byId("zoom-out").addEventListener("click", () => current()?.zoomOut());
byId("status-zoom-in").addEventListener("click", () => current()?.zoomIn());
byId("status-zoom-out").addEventListener("click", () => current()?.zoomOut());
byId("zoom-reset").addEventListener("click", () => current()?.resetZoom());
byId("zoom-value").addEventListener("click", () => current()?.resetZoom());
zoomSlider.addEventListener("input", () => {
  updateSlider(zoomSlider, Number(zoomSlider.value));
  current()?.setZoomScale(Number(zoomSlider.value));
});

/* -------------------------------- Render -------------------------------- */

function render(controller) {
  const hasWorkbook = controller.sheets.length > 0;
  const readOnly = readOnlyInput.checked || controller.readOnly;
  const editable = canEdit(controller);
  const name = hasWorkbook ? controller.displayFileName : "Book1";

  byId("file-name").textContent = name;
  document.title = `${name} · Vanilla XLSX Playground`;
  byId("read-only-badge").hidden = !readOnlyInput.checked;
  byId("download").disabled = !controller.canExport && !controller.canDownload;

  byId("undo").disabled = !controller.canUndo;
  byId("redo").disabled = !controller.canRedo;
  for (const element of document.querySelectorAll("[data-edit]")) element.disabled = !editable;
  byId("merge").disabled = !controller.selection || readOnly;

  // Formula bar
  const target = controller.selectedFormulaTarget;
  const isChartSeries = target?.kind === "chartSeries";
  const hasTarget = Boolean(controller.activeCellAddress) || isChartSeries;
  const isPivotCell = !isChartSeries && Boolean(controller.activeCell && controller.getPivotTableAtCell?.(controller.activeCell));
  byId("name-box").value = isChartSeries
    ? `Series ${target.seriesIndex + 1}`
    : (controller.selectedRangeAddress?.split("!").pop() ?? controller.activeCellAddress ?? "");
  formulaInput.disabled = !hasTarget || readOnly || isPivotCell;
  formulaInput.title = isPivotCell ? "Change PivotTable values through its field settings or source data." : "";
  if (!formulaFocused) formulaInput.value = controller.selectedFormula ?? "";

  // Sheet tabs
  renderSheetTabs(controller);
  byId("sheet-prev").disabled = controller.activeSheetIndex <= 0;
  byId("sheet-next").disabled = controller.activeSheetIndex >= controller.sheets.length - 1;
  byId("sheet-add").disabled = !hasWorkbook || readOnly;

  // Review
  byId("highlight-button").disabled = !hasWorkbook || controller.isLoading;
  byId("clear-highlights").disabled = controller.highlightedRanges.length === 0;

  // Status bar
  byId("status-spinner").hidden = !controller.isLoading;
  byId("status-text").textContent = controller.error
    ? "Could not open workbook"
    : controller.isLoading ? "Loading…" : readOnly ? "Read-only" : "Ready";
  byId("selection").textContent = controller.selectedRangeAddress ?? "";
  const runtime = controller.isLoading ? "Loading" : controller.isLoadDeferred ? "Deferred" : hasWorkbook ? "Ready" : "Idle";
  byId("runtime-label").textContent = runtime;
  byId("runtime-dot").dataset.state = hasWorkbook ? "ready" : "idle";

  // Zoom
  const zoom = Math.round(controller.zoomScale);
  zoomSlider.min = String(controller.minZoomScale);
  zoomSlider.max = String(controller.maxZoomScale);
  updateSlider(zoomSlider, zoom);
  byId("zoom-value").textContent = `${zoom}%`;
  for (const element of document.querySelectorAll("[data-zoom]")) element.disabled = !hasWorkbook;
  byId("zoom-in").disabled = byId("status-zoom-in").disabled = !hasWorkbook || !controller.canZoomIn;
  byId("zoom-out").disabled = byId("status-zoom-out").disabled = !hasWorkbook || !controller.canZoomOut;

  if (controller.error && controller.error !== lastError) showToast(controller.error.message, true);
  lastError = controller.error;
}

let lastError = null;

viewer.subscribe(render);

/* ------------------------------- Workbooks ------------------------------ */

async function openFile(file) {
  const sequence = ++loadSequence;
  try {
    const buffer = await file.arrayBuffer();
    if (sequence !== loadSequence) return;
    viewer.setOptions({ file: buffer, src: undefined, fileName: file.name });
  } catch (error) {
    if (sequence === loadSequence) showToast(error instanceof Error ? error.message : "Could not read the workbook.", true);
  }
}

function loadSample() {
  loadSequence += 1;
  viewer.setOptions({ file: undefined, src: SAMPLE_URL, fileName: SAMPLE_NAME });
  viewer.reload();
}

byId("open-file").addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  fileInput.value = "";
  if (file) void openFile(file);
});
setupFileDrop(byId("app"), byId("drop-overlay"), (file) => void openFile(file));
byId("load-sample").addEventListener("click", loadSample);

byId("url-form").addEventListener("submit", (event) => {
  event.preventDefault();
  const src = byId("workbook-url").value.trim();
  if (!src) return;
  loadSequence += 1;
  viewer.setOptions({ file: undefined, src, fileName: undefined });
});

byId("highlight-form").addEventListener("submit", (event) => {
  event.preventDefault();
  const ranges = byId("range-input").value.split(",").map((range) => range.trim()).filter(Boolean);
  if (ranges.length === 0) return;
  const accepted = current()?.highlightRanges(ranges);
  showToast(accepted ? `Highlighted ${ranges.join(", ")}.` : "Enter valid ranges on visible worksheets.", !accepted);
});
byId("clear-highlights").addEventListener("click", () => current()?.clearHighlightedRanges());

readOnlyInput.addEventListener("change", () => viewer.setOptions({ readOnly: readOnlyInput.checked }));
darkSheetInput.addEventListener("change", () => viewer.setOptions({ isDark: darkSheetInput.checked }));
