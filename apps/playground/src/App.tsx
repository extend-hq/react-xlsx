import * as React from "react";
import { useTheme } from "next-themes";
import {
  useXlsxViewer,
  useXlsxViewerController,
  useXlsxViewerEditing,
  useXlsxViewerSelection,
  useXlsxViewerThumbnails,
  useXlsxViewerZoom,
  XlsxViewer,
  XlsxViewerProvider
} from "@extend-ai/react-xlsx";
import {
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Download,
  FileSpreadsheet,
  Link2,
  Minus,
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
import { DropdownMenu, DropdownMenuContent, DropdownMenuGroup, DropdownMenuItem, DropdownMenuLabel, DropdownMenuSeparator, DropdownMenuTrigger } from "./components/ui/dropdown-menu";
import { Input } from "./components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "./components/ui/select";
import { Switch } from "./components/ui/switch";
import { Tooltip, TooltipContent, TooltipTrigger } from "./components/ui/tooltip";

const AUTO_READ_ONLY_THRESHOLD_BYTES = 5 * 1024 * 1024;
const PLAYGROUND_SAMPLE_URL = "/examples/welcome.xlsx";

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

function ToolbarCluster({ children }: { children: React.ReactNode }) {
  return (
    <div className="bg-background/80 flex min-w-0 items-center gap-1 rounded-lg border px-2 py-1 shadow-sm">
      {children}
    </div>
  );
}

function ViewerEmptyState() {
  return (
    <div className="flex h-full w-full items-center justify-center px-6">
      <div className="mx-auto max-w-sm text-center">
        <div className="text-sm font-medium">Open an XLSX workbook</div>
        <p className="text-muted-foreground mt-2 text-xs leading-5">
          Use the ribbon above to upload a local file or load a workbook from a URL.
        </p>
      </div>
    </div>
  );
}

function ViewerFileTooLargeState() {
  return (
    <div className="grid h-full w-full place-items-center">
      <div className="text-sm font-medium">File is too large</div>
    </div>
  );
}

function WorkbookToolbar({
  experimentalCanvas,
  isDocumentDark,
  onClear,
  onLoadExampleUrl,
  onLoadUrl,
  onOpenFile,
  readOnly,
  remoteUrl,
  setExperimentalCanvas,
  setIsDocumentDark,
  setReadOnly,
  setRemoteUrl,
}: {
  experimentalCanvas: boolean;
  isDocumentDark: boolean;
  onClear: () => void;
  onLoadExampleUrl: () => void;
  onLoadUrl: () => void;
  onOpenFile: () => void;
  readOnly: boolean;
  remoteUrl: string;
  setExperimentalCanvas: (value: boolean) => void;
  setIsDocumentDark: (value: boolean) => void;
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
  const { canZoomIn, canZoomOut, defaultZoomScale, resetZoom, setZoomScale, zoomIn, zoomOut, zoomScale } = useXlsxViewerZoom();
  const hasWorkbook = sheets.length > 0;
  const hasSelection = Boolean(selection);
  const hasActiveCell = Boolean(activeCellAddress);
  const isReadOnly = readOnly || viewerReadOnly;
  const zoomChoices = React.useMemo(() => {
    const presets = [50, 75, 100, 125, 150, 200];
    if (presets.includes(Math.round(zoomScale))) {
      return presets;
    }

    return [...presets, Math.round(zoomScale)].sort((left, right) => left - right);
  }, [zoomScale]);
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
            <span className="text-muted-foreground text-[11px] font-medium">Document dark</span>
            <Switch
              aria-label="Toggle document dark mode"
              checked={isDocumentDark}
              onCheckedChange={setIsDocumentDark}
              size="sm"
            />
          </div>
          <div className="flex items-center gap-1.5 rounded-md border px-2 py-1">
            <span className="text-muted-foreground text-[11px] font-medium">Canvas</span>
            <Switch
              aria-label="Toggle experimental canvas renderer"
              checked={experimentalCanvas}
              onCheckedChange={setExperimentalCanvas}
              size="sm"
            />
          </div>
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

      {/* Toolbar */}
      <div className="flex flex-wrap items-center gap-2 border-b bg-muted/30 px-2 py-2">
        <ToolbarCluster>
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
        </ToolbarCluster>

        <ToolbarCluster>
          <Input
            className="min-w-[220px]"
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
          <Button onClick={onLoadExampleUrl} size="sm" variant="outline">
            Sample
          </Button>
        </ToolbarCluster>

        <ToolbarCluster>
          <ButtonGroup>
            <Button aria-label="Undo" disabled={!canUndo} onClick={undo} size="sm" variant="outline">
              <Undo2 />
            </Button>
            <Button aria-label="Redo" disabled={!canRedo} onClick={redo} size="sm" variant="outline">
              <Redo2 />
            </Button>
          </ButtonGroup>
        </ToolbarCluster>

        <ToolbarCluster>
          <Button disabled={!hasSelection || isReadOnly} onClick={mergeSelection} size="sm" variant="outline">
            Merge
          </Button>
          <Button disabled={!hasSelection || isReadOnly} onClick={unmergeSelection} size="sm" variant="outline">
            Unmerge
          </Button>
        </ToolbarCluster>

        <ToolbarCluster>
          <Button aria-label="Add sheet" disabled={!hasWorkbook || isReadOnly} onClick={() => addSheet()} size="sm" variant="outline">
            <Plus />
          </Button>
          <Button aria-label="Delete active sheet" disabled={sheets.length <= 1 || isReadOnly} onClick={removeActiveSheet} size="sm" variant="outline">
            <Trash2 />
          </Button>
          <ButtonGroup>
            <Button aria-label="Previous sheet" disabled={!activeSheet || activeSheetIndex <= 0} onClick={() => setActiveSheetIndex(activeSheetIndex - 1)} size="sm" variant="outline">
              <ChevronLeft />
            </Button>
            <Button aria-label="Next sheet" disabled={!activeSheet || activeSheetIndex >= sheets.length - 1} onClick={() => setActiveSheetIndex(activeSheetIndex + 1)} size="sm" variant="outline">
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
        </ToolbarCluster>

        <ToolbarCluster>
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
        </ToolbarCluster>

        <ToolbarCluster>
          <ButtonGroup>
            <Button aria-label="Zoom out" disabled={!hasWorkbook || !canZoomOut} onClick={zoomOut} size="sm" variant="outline">
              <Minus />
            </Button>
            <Select
              disabled={!hasWorkbook}
              onValueChange={(value) => setZoomScale(Number(value))}
              value={String(Math.round(zoomScale))}
            >
              <SelectTrigger className="min-w-[92px]" size="sm">
                <SelectValue placeholder="Zoom" />
              </SelectTrigger>
              <SelectContent align="start">
                {zoomChoices.map((choice) => (
                  <SelectItem key={choice} value={String(choice)}>
                    {choice}%
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
            <Button aria-label="Zoom in" disabled={!hasWorkbook || !canZoomIn} onClick={zoomIn} size="sm" variant="outline">
              <Plus />
            </Button>
          </ButtonGroup>
          <Button
            disabled={!hasWorkbook || Math.round(zoomScale) === Math.round(defaultZoomScale)}
            onClick={resetZoom}
            size="sm"
            variant="outline"
          >
            Reset
          </Button>
        </ToolbarCluster>

        <ToolbarCluster>
          <Button disabled={!canExport} onClick={recalculate} size="sm" variant="outline">
            <RefreshCcw />
            Recalc
          </Button>
          <Button disabled={!hasWorkbook} onClick={onClear} size="sm" variant="outline">
            <Trash2 />
            Clear
          </Button>
        </ToolbarCluster>
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
  const { thumbnails } = useXlsxViewerThumbnails({ includeHeaders: true, resolution: { maxHeight: 132, maxWidth: 200 } });

  if (sheets.length === 0) {
    return null;
  }

  return (
    <div className="flex items-center gap-1 overflow-x-auto border-t bg-muted/35 px-3 py-2">
      {sheets.map((sheet, index) => (
        <Tooltip key={sheet.name}>
          <TooltipTrigger render={<span className="inline-flex" />}>
            <button
              className={`rounded-md border px-3 py-1.5 text-xs font-medium transition-colors ${
                index === activeSheetIndex
                  ? "border-border bg-background text-foreground shadow-sm"
                  : "border-transparent bg-transparent text-muted-foreground hover:bg-muted"
              }`}
              onClick={() => setActiveSheetIndex(index)}
              type="button"
            >
              {sheet.name}
            </button>
          </TooltipTrigger>
          <TooltipContent
            align="center"
            className="bg-popover text-popover-foreground rounded-xl border p-2 shadow-lg"
            side="top"
            sideOffset={8}
          >
            <SheetTabThumbnail
              isActive={index === activeSheetIndex}
              name={sheet.name}
              thumbnail={thumbnails[index] ?? null}
            />
          </TooltipContent>
        </Tooltip>
      ))}
    </div>
  );
}

function SheetTabThumbnail({
  isActive,
  name,
  thumbnail
}: {
  isActive: boolean;
  name: string;
  thumbnail: ReturnType<typeof useXlsxViewerThumbnails>["thumbnails"][number] | null;
}) {
  const canvasRef = React.useRef<HTMLCanvasElement | null>(null);

  React.useEffect(() => {
    thumbnail?.paint(canvasRef.current);
  }, [thumbnail]);

  return (
    <div className="flex w-[212px] flex-col gap-2">
      <div className="flex items-center justify-between gap-3">
        <div className="truncate text-[11px] font-medium">{name}</div>
        <div className="text-muted-foreground shrink-0 text-[10px] uppercase tracking-[0.16em]">
          {isActive ? "Active" : "Sheet"}
        </div>
      </div>
      <div className="bg-muted/40 overflow-hidden rounded-lg border shadow-sm">
        {thumbnail ? (
          <canvas
            className="block h-auto w-full"
            height={thumbnail.height}
            ref={canvasRef}
            width={thumbnail.width}
          />
        ) : (
          <div className="text-muted-foreground flex h-[120px] items-center justify-center text-[11px]">
            No preview
          </div>
        )}
      </div>
    </div>
  );
}

export function App() {
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [remoteUrl, setRemoteUrl] = React.useState("");
  const [source, setSource] = React.useState<ViewerSource>(null);
  const [isReadingFile, setIsReadingFile] = React.useState(false);
  const [isDragActive, setIsDragActive] = React.useState(false);
  const [isDocumentDark, setIsDocumentDark] = React.useState(false);
  const [experimentalCanvas, setExperimentalCanvas] = React.useState(false);
  const [isReadOnly, setIsReadOnly] = React.useState(false);
  const dragDepthRef = React.useRef(0);

  const controller = useXlsxViewerController(
    source?.type === "file"
      ? { file: source.file, fileName: source.fileName, readOnly: isReadOnly, readOnlyAboveBytes: AUTO_READ_ONLY_THRESHOLD_BYTES }
      : source?.type === "url"
        ? { src: source.src, fileName: source.fileName, readOnly: isReadOnly, readOnlyAboveBytes: AUTO_READ_ONLY_THRESHOLD_BYTES }
        : { readOnly: isReadOnly, readOnlyAboveBytes: AUTO_READ_ONLY_THRESHOLD_BYTES }
  );
  const zoomInitializedForSourceRef = React.useRef<string | null>(null);
  const sourceKey = React.useMemo(() => {
    if (!source) {
      return null;
    }

    return source.type === "file"
      ? `file:${source.fileName}:${source.file.byteLength}`
      : `url:${source.fileName ?? ""}:${source.src}`;
  }, [source]);

  React.useEffect(() => {
    if (!sourceKey) {
      zoomInitializedForSourceRef.current = null;
      return;
    }

    if (controller.tabs.length === 0 || zoomInitializedForSourceRef.current === sourceKey) {
      return;
    }

    controller.setZoomScale(100);
    zoomInitializedForSourceRef.current = sourceKey;
  }, [controller, sourceKey]);

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

  const loadRemoteUrl = React.useCallback((nextUrl: string) => {
    const trimmed = nextUrl.trim();
    if (!trimmed) {
      return;
    }

    setSource({
      src: trimmed,
      type: "url"
    });
  }, []);

  const handleLoadUrl = React.useCallback(() => {
    loadRemoteUrl(remoteUrl);
  }, [loadRemoteUrl, remoteUrl]);

  const handleLoadExampleUrl = React.useCallback(() => {
    setRemoteUrl(PLAYGROUND_SAMPLE_URL);
    loadRemoteUrl(PLAYGROUND_SAMPLE_URL);
  }, [loadRemoteUrl]);

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
          <XlsxViewerProvider controller={controller} isDark={isDocumentDark}>
            <WorkbookToolbar
              experimentalCanvas={experimentalCanvas}
              isDocumentDark={isDocumentDark}
              onClear={handleClear}
              onLoadExampleUrl={handleLoadExampleUrl}
              onLoadUrl={handleLoadUrl}
              onOpenFile={() => fileInputRef.current?.click()}
              readOnly={isReadOnly}
              remoteUrl={remoteUrl}
              setExperimentalCanvas={setExperimentalCanvas}
              setIsDocumentDark={setIsDocumentDark}
              setReadOnly={setIsReadOnly}
              setRemoteUrl={setRemoteUrl}
            />
            <div className="min-h-0 min-w-0 flex-1 overflow-hidden bg-muted/20 p-2">
              <div className="min-h-0 min-w-0 flex h-full w-full overflow-hidden rounded-lg border bg-muted/40 p-2.5">
                <XlsxViewer
                  className="h-full min-h-0 min-w-0 flex-1"
                  emptyState={<ViewerEmptyState />}
                  fileTooLargeState={<ViewerFileTooLargeState />}
                  height="100%"
                  allowResizeInReadOnly
                  isDark={isDocumentDark}
                  loadingState={
                    <div className="text-muted-foreground flex h-full w-full items-center justify-center text-sm">
                      Loading...
                    </div>
                  }
                  experimentalCanvas={experimentalCanvas}
                  readOnly={isReadOnly}
                  renderTableHeaderMenu={({ column, direction, sortAscending, sortDescending, triggerProps }) => (
                    <DropdownMenu>
                      <DropdownMenuTrigger
                        render={<Button size="icon-xs" variant="ghost" />}
                        {...triggerProps}
                      >
                        <ChevronDown className="size-3" />
                      </DropdownMenuTrigger>
                      <DropdownMenuContent align="end" className="w-40">
                        <DropdownMenuGroup>
                          <DropdownMenuLabel>{column.name}</DropdownMenuLabel>
                        </DropdownMenuGroup>
                        <DropdownMenuSeparator />
                        <DropdownMenuItem onClick={sortAscending}>
                          Sort A to Z{direction === "ascending" ? " ✓" : ""}
                        </DropdownMenuItem>
                        <DropdownMenuItem onClick={sortDescending}>
                          Sort Z to A{direction === "descending" ? " ✓" : ""}
                        </DropdownMenuItem>
                      </DropdownMenuContent>
                    </DropdownMenu>
                  )}
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
