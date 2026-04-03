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
import { ButtonGroup, ButtonGroupText } from "./components/ui/button-group";
import { Input } from "./components/ui/input";
import {
  Menubar,
  MenubarContent,
  MenubarItem,
  MenubarMenu,
  MenubarSeparator,
  MenubarTrigger
} from "./components/ui/menubar";
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
    <div className="flex min-w-0 flex-col gap-2 rounded-lg border bg-background/70 px-3 py-3 shadow-sm">
      <div className="flex min-w-0 flex-wrap items-center gap-2">{children}</div>
      <div className="text-muted-foreground text-[10px] font-medium uppercase tracking-[0.18em]">{label}</div>
    </div>
  );
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
  const { activeCellAddress, selectedRangeAddress, selection } = useXlsxViewerSelection();
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
    selectedValue,
    setSelectedCellFormula,
    setSelectedCellValue,
    undo,
    unmergeSelection,
  } = useXlsxViewerEditing();
  const hasWorkbook = sheets.length > 0;
  const hasSelection = Boolean(selection);
  const hasActiveCell = Boolean(activeCellAddress);
  const isReadOnly = readOnly || viewerReadOnly;
  const [formulaDraft, setFormulaDraft] = React.useState("");
  const [valueDraft, setValueDraft] = React.useState("");
  const [namedRangeDraft, setNamedRangeDraft] = React.useState("");
  const selectionLabel = selectedRangeAddress ?? activeCellAddress ?? "No selection";

  React.useEffect(() => {
    setFormulaDraft(selectedFormula);
  }, [selectedFormula, activeCellAddress]);

  React.useEffect(() => {
    setValueDraft(selectedValue);
  }, [selectedValue, activeCellAddress]);

  const commitFormula = React.useCallback(() => {
    if (!hasActiveCell) {
      return;
    }

    setSelectedCellFormula(formulaDraft);
  }, [formulaDraft, hasActiveCell, setSelectedCellFormula]);

  const commitValue = React.useCallback(() => {
    if (!hasActiveCell) {
      return;
    }

    setSelectedCellValue(valueDraft);
  }, [hasActiveCell, setSelectedCellValue, valueDraft]);

  const handleDefineNamedRange = React.useCallback(() => {
    const nextName = namedRangeDraft.trim();
    if (!nextName || !selection) {
      return;
    }

    defineNamedRange(nextName, selection);
    setNamedRangeDraft("");
  }, [defineNamedRange, namedRangeDraft, selection]);

  return (
    <div className="border-b bg-background/95 backdrop-blur supports-[backdrop-filter]:bg-background/80">
      <div className="flex h-12 items-center gap-3 border-b px-4">
        <div className="flex min-w-0 items-center gap-3">
          <div className="flex size-7 items-center justify-center rounded-md bg-emerald-600 text-white shadow-sm">
            <FileSpreadsheet className="size-4" />
          </div>
          <div className="min-w-0">
            <div className="truncate text-sm font-medium">{displayFileName}</div>
            {activeSheet ? (
              <div className="text-muted-foreground truncate text-[11px]">{activeSheet.name}</div>
            ) : null}
          </div>
        </div>
      </div>

      <div className="flex items-center justify-between gap-4 border-b px-4 py-2">
        <Menubar className="h-8">
          <MenubarMenu>
            <MenubarTrigger>File</MenubarTrigger>
            <MenubarContent>
              <MenubarItem onClick={onOpenFile}>
                <Upload />
                Open Workbook
              </MenubarItem>
              <MenubarItem disabled={!canDownload} onClick={download}>
                <Download />
                Download Source
              </MenubarItem>
              <MenubarItem disabled={!canExport} onClick={exportXlsx}>
                <Download />
                Export XLSX
              </MenubarItem>
              <MenubarItem disabled={!canExport} onClick={exportCsv}>
                <Download />
                Export CSV
              </MenubarItem>
              <MenubarSeparator />
              <MenubarItem disabled={!hasWorkbook} onClick={onClear}>
                Clear Workbook
              </MenubarItem>
            </MenubarContent>
          </MenubarMenu>
          <MenubarMenu>
            <MenubarTrigger>Workbook</MenubarTrigger>
            <MenubarContent>
              <MenubarItem disabled={!canExport} onClick={recalculate}>
                <RefreshCcw />
                Recalculate Formulas
              </MenubarItem>
              <MenubarItem disabled={!hasWorkbook || isReadOnly} onClick={() => addSheet()}>
                <Plus />
                New Sheet
              </MenubarItem>
              <MenubarItem disabled={sheets.length <= 1 || isReadOnly} onClick={removeActiveSheet}>
                <Trash2 />
                Delete Sheet
              </MenubarItem>
              <MenubarItem disabled={!activeSheet || activeSheetIndex <= 0} onClick={() => setActiveSheetIndex(activeSheetIndex - 1)}>
                <ChevronLeft />
                Previous Sheet
              </MenubarItem>
              <MenubarItem
                disabled={!activeSheet || activeSheetIndex >= sheets.length - 1}
                onClick={() => setActiveSheetIndex(activeSheetIndex + 1)}
              >
                <ChevronRight />
                Next Sheet
              </MenubarItem>
            </MenubarContent>
          </MenubarMenu>
          <MenubarMenu>
            <MenubarTrigger>Edit</MenubarTrigger>
            <MenubarContent>
              <MenubarItem disabled={!hasSelection || isReadOnly} onClick={mergeSelection}>
                Merge Selection
              </MenubarItem>
              <MenubarItem disabled={!hasSelection || isReadOnly} onClick={unmergeSelection}>
                Unmerge Selection
              </MenubarItem>
              <MenubarSeparator />
              <MenubarItem disabled={!canUndo} onClick={undo}>
                <Undo2 />
                Undo
              </MenubarItem>
              <MenubarItem disabled={!canRedo} onClick={redo}>
                <Redo2 />
                Redo
              </MenubarItem>
              <MenubarSeparator />
              <MenubarItem disabled={!hasActiveCell || isReadOnly} onClick={commitValue}>
                Apply Cell Value
              </MenubarItem>
              <MenubarItem disabled={!hasActiveCell || isReadOnly} onClick={commitFormula}>
                Apply Formula
              </MenubarItem>
            </MenubarContent>
          </MenubarMenu>
        </Menubar>
        <div className="flex items-center gap-2">
          <span className="text-muted-foreground text-xs font-medium">Read only</span>
          <Switch
            aria-label="Toggle read only mode"
            checked={isReadOnly}
            onCheckedChange={setReadOnly}
            size="sm"
          />
        </div>
      </div>

      <div className="grid gap-3 px-4 py-3 xl:grid-cols-[minmax(340px,1.2fr)_minmax(320px,1fr)_minmax(360px,1.2fr)_auto]">
        <RibbonGroup label="Workbook">
          <Button onClick={onOpenFile} size="sm">
            <Upload />
            Open
          </Button>
          <Input
            className="min-w-[240px] xl:min-w-[300px]"
            onChange={(event) => setRemoteUrl(event.target.value)}
            placeholder="https://example.com/report.xlsx"
            value={remoteUrl}
          />
          <Button onClick={onLoadUrl} size="sm" variant="outline">
            <Link2 />
            Load URL
          </Button>
          <ThemeToggle />
        </RibbonGroup>

        <RibbonGroup label="Selection">
          <ButtonGroup>
            <Button disabled={!canUndo} onClick={undo} size="sm" variant="outline">
              <Undo2 />
              Undo
            </Button>
            <Button disabled={!canRedo} onClick={redo} size="sm" variant="outline">
              <Redo2 />
              Redo
            </Button>
          </ButtonGroup>
          <ButtonGroup>
            <Button disabled={!hasWorkbook || isReadOnly} onClick={() => addSheet()} size="sm" variant="outline">
              <Plus />
              Sheet
            </Button>
            <Button disabled={sheets.length <= 1 || isReadOnly} onClick={removeActiveSheet} size="sm" variant="outline">
              <Trash2 />
              Delete
            </Button>
          </ButtonGroup>
          <Button disabled={!hasSelection || isReadOnly} onClick={mergeSelection} size="sm" variant="outline">
            Merge
          </Button>
          <Button disabled={!hasSelection || isReadOnly} onClick={unmergeSelection} size="sm" variant="outline">
            Unmerge
          </Button>
          <ButtonGroupText>{selectionLabel}</ButtonGroupText>
          <Input
            className="min-w-[140px]"
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
            Save Name
          </Button>
        </RibbonGroup>

        <RibbonGroup label="Edit">
          <div className="grid min-w-0 flex-1 gap-2 md:grid-cols-[90px_minmax(180px,1fr)]">
            <Input className="font-mono text-xs" readOnly value={activeCellAddress ?? ""} />
            <Input
              disabled={!hasActiveCell || isReadOnly}
              onBlur={commitFormula}
              onChange={(event) => setFormulaDraft(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === "Enter") {
                  event.preventDefault();
                  commitFormula();
                }
              }}
              placeholder="Formula"
              value={formulaDraft}
            />
            <Input className="font-mono text-xs" readOnly value={selectedRangeAddress ?? activeCellAddress ?? ""} />
            <Input
              disabled={!hasActiveCell || isReadOnly}
              onBlur={commitValue}
              onChange={(event) => setValueDraft(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === "Enter") {
                  event.preventDefault();
                  commitValue();
                }
              }}
              placeholder="Value"
              value={valueDraft}
            />
          </div>
        </RibbonGroup>

        <RibbonGroup label="Workbook">
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
          <Button disabled={!canExport} onClick={recalculate} size="sm" variant="outline">
            <RefreshCcw />
            Recalc
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
            <SelectTrigger className="min-w-[180px]" size="sm">
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
          <ButtonGroupText>{activeSheet?.name ?? "No sheet"}</ButtonGroupText>
        </RibbonGroup>
      </div>
    </div>
  );
}

function SheetTabs() {
  const { activeSheetIndex, setActiveSheetIndex, sheets } = useXlsxViewer();

  if (sheets.length <= 1) {
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
      ? { file: source.file, fileName: source.fileName, readOnly: isReadOnly }
      : source?.type === "url"
        ? { src: source.src, fileName: source.fileName, readOnly: isReadOnly }
        : { readOnly: isReadOnly }
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

      <div className="mx-auto flex h-full min-h-0 max-w-[1800px] flex-col overflow-hidden px-4 py-4 md:px-6">
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
            <div className="min-h-0 min-w-0 flex-1 overflow-hidden bg-muted/20 p-3">
              <div className="min-h-0 min-w-0 flex h-full w-full overflow-hidden rounded-lg border bg-muted/40 p-4">
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
