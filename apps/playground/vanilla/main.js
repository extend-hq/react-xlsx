import { createXlsxViewer, setWasmSource } from "@extend-ai/react-xlsx/vanilla";

setWasmSource("/duke_sheets_wasm_bg.wasm");

const byId = (id) => document.getElementById(id);
const fileInput = byId("file-input");
const urlInput = byId("workbook-url");
const rangeInput = byId("range-input");
const readOnlyInput = byId("read-only");
const darkModeInput = byId("dark-mode");
const status = byId("status");
const statusText = byId("status-text");
const message = byId("message");
const sampleUrl = "/examples/welcome.xlsx";
let loadSequence = 0;

const viewer = createXlsxViewer(byId("viewer"), {
  src: sampleUrl,
  height: "100%",
  rounded: false,
  readOnlyAboveBytes: 5 * 1024 * 1024
});

function showMessage(text, isError = false) {
  message.textContent = text;
  message.dataset.error = String(isError);
}

function updateStatus(controller) {
  const loaded = controller.sheets.length > 0 || controller.chartsheets.length > 0;
  byId("file-name").textContent = controller.displayFileName || "Workbook";
  byId("sheet-name").textContent = controller.activeTab?.name || (controller.isLoading ? "Opening workbook" : "No workbook open");
  byId("selection").textContent = controller.selectedRangeAddress || controller.activeCellAddress || "—";
  byId("export").disabled = !controller.canExport;
  byId("highlight-button").disabled = !loaded || controller.isLoading;
  byId("clear-highlights").disabled = controller.highlightedRanges.length === 0;

  if (controller.error) {
    status.dataset.state = "error";
    statusText.textContent = "Could not open";
    showMessage(controller.error.message, true);
  } else if (controller.isLoading) {
    status.dataset.state = "loading";
    statusText.textContent = "Loading";
  } else if (controller.isLoadDeferred) {
    status.dataset.state = "loading";
    statusText.textContent = "Waiting to load";
  } else {
    status.dataset.state = "ready";
    statusText.textContent = loaded ? (controller.readOnly ? "Read only" : "Ready to edit") : "No workbook";
  }
}

viewer.subscribe(updateStatus);

byId("open-file").addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", async () => {
  const file = fileInput.files?.[0];
  if (!file) return;
  const sequence = ++loadSequence;
  showMessage("Opening workbook…");
  try {
    const buffer = await file.arrayBuffer();
    if (sequence !== loadSequence) return;
    viewer.setOptions({ file: buffer, fileName: file.name });
    showMessage("Select a cell to inspect or edit it.");
  } catch (error) {
    if (sequence === loadSequence) showMessage(error instanceof Error ? error.message : "Could not read the workbook.", true);
  }
  fileInput.value = "";
});

byId("load-sample").addEventListener("click", () => {
  loadSequence += 1;
  viewer.setOptions({ src: sampleUrl, fileName: undefined });
  viewer.reload();
  showMessage("Select a cell to inspect or edit it.");
});

byId("url-form").addEventListener("submit", (event) => {
  event.preventDefault();
  const src = urlInput.value.trim();
  if (!src) return;
  loadSequence += 1;
  viewer.setOptions({ src, fileName: undefined });
  showMessage("Opening workbook…");
});

readOnlyInput.addEventListener("change", () => {
  viewer.setOptions({ readOnly: readOnlyInput.checked });
});

darkModeInput.addEventListener("change", () => {
  viewer.setOptions({ isDark: darkModeInput.checked });
});

byId("highlight-form").addEventListener("submit", (event) => {
  event.preventDefault();
  const range = rangeInput.value.trim();
  if (!range) return;
  const accepted = viewer.getController()?.highlightRanges(range);
  showMessage(accepted ? `Highlighted ${range}.` : "Enter a valid range on a visible worksheet.", !accepted);
});

byId("clear-highlights").addEventListener("click", () => {
  viewer.getController()?.clearHighlightedRanges();
  showMessage("Highlights cleared.");
});

byId("export").addEventListener("click", () => viewer.getController()?.exportXlsx());
