import * as React from "react";
import { useTheme } from "next-themes";
import {
  useXlsxViewer,
  useXlsxViewerController,
  useXlsxViewerEditing,
  useXlsxViewerSelection,
  XlsxViewer,
  XlsxViewerProvider
} from "react-xlsx";

const AUTO_READ_ONLY_THRESHOLD_BYTES = 5 * 1024 * 1024;
import {
  ChevronLeft,
  ChevronRight,
  Download,
  FileSpreadsheet,
  Link2,
  Moon,
  Plus,
  Redo2,
  RefreshCcw,
  Sun,
  Trash2,
  Undo2,
  Upload
} from "lucide-react";
import { Button } from "./components/ui/button";
import { ButtonGroup } from "./components/ui/button-group";
import { Input } from "./components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "./components/ui/select";
import { Switch } from "./components/ui/switch";

type ViewerSource =
  | {
      file: ArrayBuffer;
      fileName: string;
      type: "file";
    }
  | {
      fileName?: string;
      src: string;
      type: "url";
    }
  | null;

function ThemeToggle() {
  const { resolvedTheme, setTheme, theme } = useTheme();
  const [ready, setReady] = React.useState(false);

  React.useEffect(() => {
    setReady(true);
  }, []);

  const currentTheme = (resolvedTheme ?? theme ?? "light") as "light" | "dark";
  const isDark = currentTheme === "dark";

  return (
    <Button
      aria-label="Toggle theme"
      onClick={() => setTheme(isDark ? "light" : "dark")}
      size="icon-sm"
      variant="outline"
    >
      {ready && isDark ? <Sun /> : <Moon />}
    </Button>
  );
}

function RibbonGroup({
  children,
  label,
}: {
  children: React.ReactNode;
  label: string;
}) {
  return (
    <div className="flex flex-col items-stretch">
      <div className="flex items-center gap-1 px-2 py-1">
        {children}
      </div>
      <div className="text-muted-foreground border-t px-2 py-0.5 text-center text-[10px] font-medium">
        {label}
      </div>
    </div>
  );
}

function RibbonSeparator() {
  return <div className="bg-border mx-1 h-[52px] w-px shrink-0 self-center" />;
}

function ViewerEmptyState() {
  return (
    <div className="flex h-full w-full items-center justify-center px-6">
      <div className="mx-auto max-w-sm text-center">
        <div className="mx-auto flex size-12 items-center justify-center rounded-2xl border bg-muted/50 shadow-sm">
          <FileSpreadsheet className="text-muted-foreground size-5" />
        </div>
        <div className="mt-4 text-sm font-medium">Open an XLSX workbook</div>
        <p className="text-muted-foreground mt-2 text-xs leading-5">
          Use the ribbon above to upload a local file or load a workbook from a URL.
        </p>
      </div>
    </div>
  );
}

function WorkbookToolbar({
  onClear,
  onLoadUrl,
  onOpenFile,
  readOnly,
  remoteUrl,
  setReadOnly,
  setRemoteUrl,
}: {
  onClear: () => void;
  onLoadUrl: () => void;
  onOpenFile: () => void;
  readOnly: boolean;
  remoteUrl: string;
  setReadOnly: (value: boolean) => void;
  setRemoteUrl: (value: string) => void;
}) {
  const {
    activeSheet,
    activeSheetIndex,
    canDownload,
    canExport,
    displayFileName,
    download,
    exportCsv,
    exportXlsx,
    recalculate,
    setActiveSheetIndex,
    sheets,
  } = useXlsxViewer();
  const { activeCell, activeCellAddress, selection } = useXlsxViewerSelection();
  const {
    addSheet,
    canRedo,
    canUndo,
    defineNamedRange,
    mergeSelection,
    removeActiveSheet,
    readOnly: viewerReadOnly,
    redo,
    selectedFormula,
    setCellFormula,
    undo,
    unmergeSelection,
  } = useXlsxViewerEditing();
  const hasWorkbook = sheets.length > 0;
  const hasSelection = Boolean(selection);
  const hasActiveCell = Boolean(activeCellAddress);
  const isReadOnly = readOnly || viewerReadOnly;
  const [formulaDraft, setFormulaDraft] = React.useState("");
  const [namedRangeDraft, setNamedRangeDraft] = React.useState("");
  const [focusedField, setFocusedField] = React.useState<"formula" | null>(null);
  const formulaEditCellRef = React.useRef<typeof activeCell>(null);
  const formulaInitialValueRef = React.useRef("");

  React.useEffect(() => {
    if (focusedField === "formula") {
      return;
    }
    setFormulaDraft(selectedFormula);
  }, [selectedFormula, activeCellAddress, focusedField]);

  const commitFormula = React.useCallback((targetCell?: typeof activeCell | null, nextFormula?: string) => {
    const resolvedCell = targetCell ?? formulaEditCellRef.current;
    const resolvedFormula = nextFormula ?? formulaDraft;
    if (!resolvedCell) {
      return;
    }

    if (resolvedFormula === formulaInitialValueRef.current) {
      return;
    }

    setCellFormula(resolvedCell, resolvedFormula);
  }, [formulaDraft, setCellFormula]);

  const handleDefineNamedRange = React.useCallback(() => {
    const nextName = namedRangeDraft.trim();
    if (!nextName || !selection) {
      return;
    }

    defineNamedRange(nextName, selection);
    setNamedRangeDraft("");
  }, [defineNamedRange, namedRangeDraft, selection]);

  return (
    <div className="bg-background/95 backdrop-blur supports-[backdrop-filter]:bg-background/80">
      {/* Title bar */}
      <div className="flex min-h-10 items-center justify-between gap-3 border-b px-4 py-1.5">
        <div className="flex min-w-0 items-center gap-3">
          <div className="flex size-6 items-center justify-center rounded-md bg-emerald-600 text-white shadow-sm">
            <FileSpreadsheet className="size-3.5" />
          </div>
          <div className="truncate text-sm font-medium">{displayFileName}</div>
          {activeSheet ? (
            <>
              <div className="bg-border h-3 w-px shrink-0" />
              <div className="text-muted-foreground truncate text-[11px]">{activeSheet.name}</div>
            </>
          ) : null}
        </div>
        <div className="flex shrink-0 items-center gap-2">
          <div className="flex items-center gap-1.5 rounded-md border px-2 py-1">
            <span className="text-muted-foreground text-[11px] font-medium">Read only</span>
            <Switch
              aria-label="Toggle read only mode"
              checked={isReadOnly}
              onCheckedChange={setReadOnly}
              size="sm"
            />
          </div>
          <ThemeToggle />
        </div>
      </div>

      {/* Ribbon */}
      <div className="flex items-stretch border-b bg-muted/30 px-2">
        <RibbonGroup label="File">
          <Button onClick={onOpenFile} size="sm">
            <Upload />
            Open
          </Button>
          <Button disabled={!canDownload} onClick={download} size="sm" variant="outline">
            <Download />
            Source
          </Button>
          <Button disabled={!canExport} onClick={exportXlsx} size="sm" variant="outline">
            <Download />
            XLSX
          </Button>
          <Button disabled={!canExport} onClick={exportCsv} size="sm" variant="outline">
            <Download />
            CSV
          </Button>
        </RibbonGroup>

        <RibbonSeparator />

        <RibbonGroup label="URL">
          <Input
            className="min-w-[200px]"
            onChange={(event) => setRemoteUrl(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === "Enter") {
                event.preventDefault();
                onLoadUrl();
              }
            }}
            placeholder="https://example.com/report.xlsx"
            value={remoteUrl}
          />
          <Button onClick={onLoadUrl} size="sm" variant="outline">
            <Link2 />
            Load
          </Button>
        </RibbonGroup>

        <RibbonSeparator />

        <RibbonGroup label="Clipboard">
          <ButtonGroup>
            <Button disabled={!canUndo} onClick={undo} size="sm" variant="outline">
              <Undo2 />
            </Button>
            <Button disabled={!canRedo} onClick={redo} size="sm" variant="outline">
              <Redo2 />
            </Button>
          </ButtonGroup>
        </RibbonGroup>

        <RibbonSeparator />

        <RibbonGroup label="Cells">
          <Button disabled={!hasSelection || isReadOnly} onClick={mergeSelection} size="sm" variant="outline">
            Merge
          </Button>
          <Button disabled={!hasSelection || isReadOnly} onClick={unmergeSelection} size="sm" variant="outline">
            Unmerge
          </Button>
        </RibbonGroup>

        <RibbonSeparator />

        <RibbonGroup label="Sheets">
          <Button disabled={!hasWorkbook || isReadOnly} onClick={() => addSheet()} size="sm" variant="outline">
            <Plus />
          </Button>
          <Button disabled={sheets.length <= 1 || isReadOnly} onClick={removeActiveSheet} size="sm" variant="outline">
            <Trash2 />
          </Button>
          <ButtonGroup>
            <Button disabled={!activeSheet || activeSheetIndex <= 0} onClick={() => setActiveSheetIndex(activeSheetIndex - 1)} size="sm" variant="outline">
              <ChevronLeft />
            </Button>
            <Button disabled={!activeSheet || activeSheetIndex >= sheets.length - 1} onClick={() => setActiveSheetIndex(activeSheetIndex + 1)} size="sm" variant="outline">
              <ChevronRight />
            </Button>
          </ButtonGroup>
          <Select
            disabled={sheets.length === 0}
            onValueChange={(value) => setActiveSheetIndex(Number(value))}
            value={String(activeSheetIndex)}
          >
            <SelectTrigger className="min-w-[140px]" size="sm">
              <SelectValue placeholder="Select sheet" />
            </SelectTrigger>
            <SelectContent align="start">
              {sheets.map((sheet, index) => (
                <SelectItem key={sheet.name} value={String(index)}>
                  {sheet.name}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        </RibbonGroup>

        <RibbonSeparator />

        <RibbonGroup label="Names">
          <Input
            className="min-w-[120px]"
            onChange={(event) => setNamedRangeDraft(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === "Enter") {
                event.preventDefault();
                handleDefineNamedRange();
              }
            }}
            placeholder="Named range"
            value={namedRangeDraft}
          />
          <Button disabled={!hasSelection || !namedRangeDraft.trim() || isReadOnly} onClick={handleDefineNamedRange} size="sm" variant="outline">
            Define
          </Button>
        </RibbonGroup>

        <RibbonSeparator />

        <RibbonGroup label="Tools">
          <Button disabled={!canExport} onClick={recalculate} size="sm" variant="outline">
            <RefreshCcw />
            Recalc
          </Button>
          <Button disabled={!hasWorkbook} onClick={onClear} size="sm" variant="outline">
            <Trash2 />
            Clear
          </Button>
        </RibbonGroup>
      </div>

      {/* Formula bar */}
      <div className="flex items-center gap-px border-b bg-background px-2 py-1">
        <Input
          className="w-[90px] shrink-0 border-r font-mono text-xs"
          readOnly
          value={activeCellAddress ?? ""}
        />
        <div className="text-muted-foreground flex h-7 w-8 shrink-0 items-center justify-center border-r text-[11px] font-semibold italic">
          fx
        </div>
        <Input
          className="flex-1 border-0 shadow-none focus-visible:ring-0"
          disabled={!hasActiveCell || isReadOnly}
          onBlur={() => {
            commitFormula();
            setFocusedField(null);
          }}
          onChange={(event) => setFormulaDraft(event.target.value)}
          onFocus={() => {
            formulaEditCellRef.current = activeCell;
            formulaInitialValueRef.current = formulaDraft;
            setFocusedField("formula");
          }}
          onKeyDown={(event) => {
            if (event.key === "Enter") {
              event.preventDefault();
              commitFormula();
              setFocusedField(null);
            }
          }}
          placeholder="Enter a formula or value"
          value={formulaDraft}
        />
      </div>
    </div>
  );
}

function SheetTabs() {
  const { activeSheetIndex, setActiveSheetIndex, sheets } = useXlsxViewer();

  if (sheets.length === 0) {
    return null;
  }

  return (
    <div className="flex items-center gap-1 overflow-x-auto border-t bg-muted/35 px-3 py-2">
      {sheets.map((sheet, index) => (
        <button
          className={`rounded-md border px-3 py-1.5 text-xs font-medium transition-colors ${
            index === activeSheetIndex
              ? "border-border bg-background text-foreground shadow-sm"
              : "border-transparent bg-transparent text-muted-foreground hover:bg-muted"
          }`}
          key={sheet.name}
          onClick={() => setActiveSheetIndex(index)}
          type="button"
        >
          {sheet.name}
        </button>
      ))}
    </div>
  );
}

export function App() {
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [remoteUrl, setRemoteUrl] = React.useState("");
  const [source, setSource] = React.useState<ViewerSource>(null);
  const [isReadingFile, setIsReadingFile] = React.useState(false);
  const [isDragActive, setIsDragActive] = React.useState(false);
  const [isReadOnly, setIsReadOnly] = React.useState(false);
  const dragDepthRef = React.useRef(0);

  const controller = useXlsxViewerController(
    source?.type === "file"
      ? { file: source.file, fileName: source.fileName, readOnly: isReadOnly, readOnlyAboveBytes: AUTO_READ_ONLY_THRESHOLD_BYTES }
      : source?.type === "url"
        ? { src: source.src, fileName: source.fileName, readOnly: isReadOnly, readOnlyAboveBytes: AUTO_READ_ONLY_THRESHOLD_BYTES }
        : { readOnly: isReadOnly, readOnlyAboveBytes: AUTO_READ_ONLY_THRESHOLD_BYTES }
  );

  const loadWorkbookFile = React.useCallback(async (nextFile: File) => {
    setIsReadingFile(true);
    try {
      const fileBuffer = await nextFile.arrayBuffer();
      setSource({
        file: fileBuffer,
        fileName: nextFile.name,
        type: "file"
      });
    } finally {
      setIsReadingFile(false);
    }
  }, []);

  const handleFileChange = React.useCallback(async (event: React.ChangeEvent<HTMLInputElement>) => {
    const nextFile = event.target.files?.[0];
    event.target.value = "";
    if (!nextFile) {
      return;
    }

    await loadWorkbookFile(nextFile);
  }, [loadWorkbookFile]);

  const handleLoadUrl = React.useCallback(() => {
    const trimmed = remoteUrl.trim();
    if (!trimmed) {
      return;
    }

    setSource({
      src: trimmed,
      type: "url"
    });
  }, [remoteUrl]);

  const handleClear = React.useCallback(() => {
    setSource(null);
  }, []);

  const handleDragEnter = React.useCallback((event: React.DragEvent<HTMLDivElement>) => {
    if (!event.dataTransfer.types.includes("Files")) {
      return;
    }

    event.preventDefault();
    dragDepthRef.current += 1;
    setIsDragActive(true);
  }, []);

  const handleDragOver = React.useCallback((event: React.DragEvent<HTMLDivElement>) => {
    if (!event.dataTransfer.types.includes("Files")) {
      return;
    }

    event.preventDefault();
    event.dataTransfer.dropEffect = "copy";
    if (!isDragActive) {
      setIsDragActive(true);
    }
  }, [isDragActive]);

  const handleDragLeave = React.useCallback((event: React.DragEvent<HTMLDivElement>) => {
    if (!event.dataTransfer.types.includes("Files")) {
      return;
    }

    event.preventDefault();
    dragDepthRef.current = Math.max(0, dragDepthRef.current - 1);
    if (dragDepthRef.current === 0) {
      setIsDragActive(false);
    }
  }, []);

  const handleDrop = React.useCallback(async (event: React.DragEvent<HTMLDivElement>) => {
    if (!event.dataTransfer.files.length) {
      return;
    }

    event.preventDefault();
    dragDepthRef.current = 0;
    setIsDragActive(false);

    const nextFile = event.dataTransfer.files[0];
    if (!nextFile) {
      return;
    }

    await loadWorkbookFile(nextFile);
  }, [loadWorkbookFile]);

  return (
    <div className="bg-background text-foreground h-[100dvh] overflow-hidden">
      <input
        ref={fileInputRef}
        accept=".xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        className="hidden"
        onChange={handleFileChange}
        type="file"
      />

      <div className="mx-auto flex h-full min-h-0 max-w-[1800px] flex-col overflow-hidden px-3 py-3 md:px-5 md:py-4">
        <div
          className="relative flex h-full min-h-0 min-w-0 flex-1 flex-col overflow-hidden rounded-2xl border bg-background shadow-sm"
          onDragEnter={handleDragEnter}
          onDragLeave={handleDragLeave}
          onDragOver={handleDragOver}
          onDrop={handleDrop}
        >
          <XlsxViewerProvider controller={controller}>
            <WorkbookToolbar
              onClear={handleClear}
              onLoadUrl={handleLoadUrl}
              onOpenFile={() => fileInputRef.current?.click()}
              readOnly={isReadOnly}
              remoteUrl={remoteUrl}
              setReadOnly={setIsReadOnly}
              setRemoteUrl={setRemoteUrl}
            />
            <div className="min-h-0 min-w-0 flex-1 overflow-hidden bg-muted/20 p-2">
              <div className="min-h-0 min-w-0 flex h-full w-full overflow-hidden rounded-lg border bg-muted/40 p-2.5">
                <XlsxViewer
                  className="h-full min-h-0 min-w-0 flex-1"
                  emptyState={<ViewerEmptyState />}
                  height="100%"
                  loadingState={
                    <div className="text-muted-foreground flex h-full items-center justify-center text-sm">
                      {isReadingFile ? "Reading workbook..." : "Loading workbook..."}
                    </div>
                  }
                  readOnly={isReadOnly}
                  rounded={true}
                  showDefaultToolbar={false}
                />
              </div>
            </div>
            <SheetTabs />
          </XlsxViewerProvider>
          {isDragActive ? (
            <div className="bg-background/82 absolute inset-0 z-30 flex items-center justify-center backdrop-blur-sm">
              <div className="bg-background/96 ring-border/80 flex min-w-[320px] max-w-md flex-col items-center gap-3 rounded-2xl border border-dashed px-8 py-10 text-center shadow-lg ring-1">
                <div className="bg-emerald-600 text-white flex size-12 items-center justify-center rounded-2xl shadow-sm">
                  <Upload className="size-5" />
                </div>
                <div>
                  <div className="text-sm font-medium">Drop workbook to open</div>
                  <p className="text-muted-foreground mt-2 text-xs leading-5">
                    Release to import the first `.xls` or `.xlsx` file into the viewer.
                  </p>
                </div>
              </div>
            </div>
          ) : null}
        </div>
      </div>
    </div>
  );
}
