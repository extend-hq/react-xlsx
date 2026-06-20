import * as React from "react";
import {
  useXlsxViewer,
  useXlsxViewerController,
  useXlsxViewerEditing,
  useXlsxViewerSelection,
  useXlsxViewerThumbnails,
  useXlsxViewerZoom,
  XlsxViewer,
  XlsxViewerProvider,
  type XlsxCellStyleInput,
  type XlsxViewerProps
} from "@extend-ai/react-xlsx";
import { Button } from "./components/ui/button";
import { ButtonGroup } from "./components/ui/button-group";
import { ColorPicker } from "./components/ui/color-picker";
import {
  PlaygroundCustomizerPanel,
  PlaygroundIcon,
  usePlaygroundCustomizer,
} from "./components/playground-customizer";
import { DropdownMenu, DropdownMenuContent, DropdownMenuGroup, DropdownMenuItem, DropdownMenuLabel, DropdownMenuSeparator, DropdownMenuTrigger } from "./components/ui/dropdown-menu";
import { Input } from "./components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "./components/ui/select";
import { Switch } from "./components/ui/switch";
import { Tooltip, TooltipContent, TooltipTrigger } from "./components/ui/tooltip";

const AUTO_READ_ONLY_THRESHOLD_BYTES = 5 * 1024 * 1024;
const PLAYGROUND_SAMPLE_URL = "/examples/welcome.xlsx";

const RIBBON_TABS = ["Home", "Insert", "Page Layout", "Formulas", "Data", "View"] as const;
const FONT_FAMILIES = ["Aptos", "Calibri", "Arial", "Georgia", "Times New Roman", "Courier New"] as const;
const FONT_SIZES = [9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36] as const;
const NUMBER_FORMATS = [
  { label: "General", style: { numberFormat: { formatType: "general" } } },
  { label: "Number", style: { numberFormat: { formatType: "custom", formatString: "#,##0.00" } } },
  { label: "Currency", style: { numberFormat: { formatType: "custom", formatString: "$#,##0.00" } } },
  { label: "Percent", style: { numberFormat: { formatType: "custom", formatString: "0.00%" } } },
  { label: "Date", style: { numberFormat: { formatType: "custom", formatString: "m/d/yyyy" } } },
  { label: "Time", style: { numberFormat: { formatType: "custom", formatString: "h:mm AM/PM" } } },
] satisfies Array<{ label: string; style: XlsxCellStyleInput }>;
const CELL_STYLE_PRESETS = [
  {
    label: "Good",
    style: {
      fill: { fillType: "solid", color: { colorType: "rgb", hex: "E2F0D9" } },
      font: { color: { colorType: "rgb", hex: "375623" } },
    },
  },
  {
    label: "Neutral",
    style: {
      fill: { fillType: "solid", color: { colorType: "rgb", hex: "FFF2CC" } },
      font: { color: { colorType: "rgb", hex: "7F6000" } },
    },
  },
  {
    label: "Bad",
    style: {
      fill: { fillType: "solid", color: { colorType: "rgb", hex: "FCE4D6" } },
      font: { color: { colorType: "rgb", hex: "9C0006" } },
    },
  },
  {
    label: "Heading",
    style: {
      alignment: { horizontal: "center", vertical: "center" },
      border: { bottom: { style: "medium", color: { colorType: "rgb", hex: "5B9BD5" } } },
      font: { bold: true, color: { colorType: "rgb", hex: "1F4E79" }, size: 14 },
    },
  },
] satisfies Array<{ label: string; style: XlsxCellStyleInput }>;

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

type RibbonTab = (typeof RIBBON_TABS)[number];

function hexToStyleColor(hex: string): NonNullable<XlsxCellStyleInput["font"]>["color"] {
  return {
    colorType: "rgb",
    hex: hex.replace("#", "").toUpperCase(),
  };
}

function ThemeToggle() {
  const {
    resolvedAppearance,
    settings,
    updateSettings,
  } = usePlaygroundCustomizer();
  const isDark = resolvedAppearance === "dark";

  return (
    <Button
      aria-label="Toggle theme"
      onClick={() =>
        updateSettings({
          appearance: settings.appearance === "system" ? (isDark ? "light" : "dark") : isDark ? "light" : "dark",
        })
      }
      size="icon-sm"
      variant="outline"
    >
      {isDark ? <PlaygroundIcon name="sun" /> : <PlaygroundIcon name="moon" />}
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

function RibbonGroup({
  children,
  label,
}: {
  children: React.ReactNode;
  label: string;
}) {
  return (
    <div className="bg-background/80 flex min-h-[72px] min-w-fit flex-col justify-between gap-1 rounded-lg border px-2 py-1.5 shadow-sm">
      <div className="flex items-center gap-1">{children}</div>
      <div className="text-muted-foreground truncate text-center text-[10px] font-medium">{label}</div>
    </div>
  );
}

function RibbonButton({
  children,
  label,
  tooltip,
  ...props
}: React.ComponentProps<typeof Button> & {
  label: string;
  tooltip?: string;
}) {
  const button = (
    <Button aria-label={label} size="sm" variant="outline" {...props}>
      {children}
    </Button>
  );

  if (!tooltip) {
    return button;
  }

  return (
    <Tooltip>
      <TooltipTrigger render={<span className="inline-flex" />}>{button}</TooltipTrigger>
      <TooltipContent side="bottom">{tooltip}</TooltipContent>
    </Tooltip>
  );
}

function StyleSwatchButton({
  disabled,
  label,
  onClick,
  style,
}: {
  disabled?: boolean;
  label: string;
  onClick: () => void;
  style: XlsxCellStyleInput;
}) {
  const background = style.fill?.color?.hex ? `#${style.fill.color.hex.slice(-6)}` : "transparent";
  const color = style.font?.color?.hex ? `#${style.font.color.hex.slice(-6)}` : "currentColor";

  return (
    <Button
      className="h-11 w-[68px] flex-col gap-0.5 px-1.5"
      disabled={disabled}
      onClick={onClick}
      size="sm"
      variant="outline"
    >
      <span
        className="h-3.5 w-full rounded-sm border"
        style={{ backgroundColor: background, borderColor: color }}
      />
      <span className="truncate text-[10px]">{label}</span>
    </Button>
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
  highlightCells,
  isDocumentDark,
  onClear,
  onLoadExampleUrl,
  onLoadUrl,
  onOpenFile,
  readOnly,
  remoteUrl,
  setExperimentalCanvas,
  setHighlightCells,
  setIsDocumentDark,
  setReadOnly,
  setRemoteUrl,
}: {
  experimentalCanvas: boolean;
  highlightCells: boolean;
  isDocumentDark: boolean;
  onClear: () => void;
  onLoadExampleUrl: () => void;
  onLoadUrl: () => void;
  onOpenFile: () => void;
  readOnly: boolean;
  remoteUrl: string;
  setExperimentalCanvas: (value: boolean) => void;
  setHighlightCells: (value: boolean) => void;
  setIsDocumentDark: (value: boolean) => void;
  setReadOnly: (value: boolean) => void;
  setRemoteUrl: (value: string) => void;
}) {
  const [activeRibbonTab, setActiveRibbonTab] = React.useState<RibbonTab>("Home");
  const [fontFamily, setFontFamily] = React.useState<(typeof FONT_FAMILIES)[number]>("Aptos");
  const [fontSize, setFontSize] = React.useState("11");
  const [fontColor, setFontColor] = React.useState("#1F2937");
  const [fillColor, setFillColor] = React.useState("#DBEAFE");
  const [borderColor, setBorderColor] = React.useState("#4B5563");
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
    setRangeStyle,
    setSelectedCellStyle,
    undo,
    unmergeSelection,
  } = useXlsxViewerEditing();
  const { canZoomIn, canZoomOut, defaultZoomScale, resetZoom, setZoomScale, zoomIn, zoomOut, zoomScale } = useXlsxViewerZoom();
  const hasWorkbook = sheets.length > 0;
  const hasSelection = Boolean(selection);
  const hasActiveCell = Boolean(activeCellAddress);
  const isReadOnly = readOnly || viewerReadOnly;
  const canEditSelection = hasActiveCell && !isReadOnly;
  const applyStyle = React.useCallback((style: XlsxCellStyleInput) => {
    if (!canEditSelection) {
      return;
    }

    if (selection) {
      setRangeStyle(selection, style);
      return;
    }

    setSelectedCellStyle(style);
  }, [canEditSelection, selection, setRangeStyle, setSelectedCellStyle]);
  const applyFontStyle = React.useCallback((style: NonNullable<XlsxCellStyleInput["font"]>) => {
    applyStyle({ font: style });
  }, [applyStyle]);
  const applyAlignmentStyle = React.useCallback((style: NonNullable<XlsxCellStyleInput["alignment"]>) => {
    applyStyle({ alignment: style });
  }, [applyStyle]);
  const applyBorderStyle = React.useCallback((style: NonNullable<XlsxCellStyleInput["border"]>) => {
    applyStyle({ border: style });
  }, [applyStyle]);
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
            <PlaygroundIcon className="size-3.5" name="spreadsheet" />
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
          <Button disabled={!canExport} onClick={exportXlsx} size="sm" variant="outline">
            <PlaygroundIcon name="download" />
            Download
          </Button>
          <ThemeToggle />
          <PlaygroundCustomizerPanel />
        </div>
      </div>

      {/* Ribbon */}
      <div className="border-b bg-muted/30">
        <div className="flex items-center gap-1 overflow-x-auto border-b bg-background px-2 pt-1">
          {RIBBON_TABS.map((tab) => (
            <button
              className={`h-7 whitespace-nowrap rounded-t-md border border-b-0 px-3 text-xs font-medium transition-colors ${
                activeRibbonTab === tab
                  ? "bg-muted/30 text-foreground"
                  : "border-transparent text-muted-foreground hover:bg-muted/40 hover:text-foreground"
              }`}
              key={tab}
              onClick={() => setActiveRibbonTab(tab)}
              type="button"
            >
              {tab}
            </button>
          ))}
        </div>

        <div className="flex min-h-[92px] items-stretch gap-2 overflow-x-auto px-2 py-2">
          {activeRibbonTab === "Home" ? (
            <>
              <RibbonGroup label="Clipboard">
                <ButtonGroup>
                  <RibbonButton disabled={!canUndo} label="Undo" onClick={undo}>
                    <PlaygroundIcon name="undo" />
                  </RibbonButton>
                  <RibbonButton disabled={!canRedo} label="Redo" onClick={redo}>
                    <PlaygroundIcon name="redo" />
                  </RibbonButton>
                </ButtonGroup>
              </RibbonGroup>

              <RibbonGroup label="Font">
                <div className="flex flex-col gap-1">
                  <div className="flex gap-1">
                    <Select
                      disabled={!canEditSelection}
                      onValueChange={(value) => {
                        const nextFamily = value as (typeof FONT_FAMILIES)[number];
                        setFontFamily(nextFamily);
                        applyFontStyle({ name: nextFamily });
                      }}
                      value={fontFamily}
                    >
                      <SelectTrigger className="w-[132px]" size="sm">
                        <SelectValue placeholder="Font" />
                      </SelectTrigger>
                      <SelectContent align="start">
                        {FONT_FAMILIES.map((family) => (
                          <SelectItem key={family} value={family}>
                            {family}
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                    <Select
                      disabled={!canEditSelection}
                      onValueChange={(value) => {
                        if (!value) {
                          return;
                        }
                        setFontSize(value);
                        applyFontStyle({ size: Number(value) });
                      }}
                      value={fontSize}
                    >
                      <SelectTrigger className="w-[68px]" size="sm">
                        <SelectValue placeholder="Size" />
                      </SelectTrigger>
                      <SelectContent align="start">
                        {FONT_SIZES.map((size) => (
                          <SelectItem key={size} value={String(size)}>
                            {size}
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>
                  <div className="flex gap-1">
                    <ButtonGroup>
                      <RibbonButton disabled={!canEditSelection} label="Bold" onClick={() => applyFontStyle({ bold: true })}>
                        <span className="font-bold">B</span>
                      </RibbonButton>
                      <RibbonButton disabled={!canEditSelection} label="Italic" onClick={() => applyFontStyle({ italic: true })}>
                        <span className="italic">I</span>
                      </RibbonButton>
                      <RibbonButton disabled={!canEditSelection} label="Underline" onClick={() => applyFontStyle({ underline: "single" })}>
                        <span className="underline">U</span>
                      </RibbonButton>
                      <RibbonButton disabled={!canEditSelection} label="Strikethrough" onClick={() => applyFontStyle({ strikethrough: true })}>
                        <span className="line-through">S</span>
                      </RibbonButton>
                    </ButtonGroup>
                    <ColorPicker
                      color={fontColor}
                      disabled={!canEditSelection}
                      onChange={(color) => {
                        setFontColor(color);
                        applyFontStyle({ color: hexToStyleColor(color) });
                      }}
                      triggerClassName="h-6 w-[92px] px-2 text-[11px]"
                    />
                    <ColorPicker
                      color={fillColor}
                      disabled={!canEditSelection}
                      onChange={(color) => {
                        setFillColor(color);
                        applyStyle({ fill: { fillType: "solid", color: hexToStyleColor(color) } });
                      }}
                      triggerClassName="h-6 w-[92px] px-2 text-[11px]"
                    />
                  </div>
                </div>
              </RibbonGroup>

              <RibbonGroup label="Alignment">
                <div className="flex flex-col gap-1">
                  <ButtonGroup>
                    <RibbonButton disabled={!canEditSelection} label="Align left" onClick={() => applyAlignmentStyle({ horizontal: "left" })}>
                      Left
                    </RibbonButton>
                    <RibbonButton disabled={!canEditSelection} label="Center" onClick={() => applyAlignmentStyle({ horizontal: "center" })}>
                      Center
                    </RibbonButton>
                    <RibbonButton disabled={!canEditSelection} label="Align right" onClick={() => applyAlignmentStyle({ horizontal: "right" })}>
                      Right
                    </RibbonButton>
                  </ButtonGroup>
                  <ButtonGroup>
                    <RibbonButton disabled={!canEditSelection} label="Middle align" onClick={() => applyAlignmentStyle({ vertical: "center" })}>
                      Middle
                    </RibbonButton>
                    <RibbonButton disabled={!canEditSelection} label="Wrap text" onClick={() => applyAlignmentStyle({ wrapText: true })}>
                      Wrap
                    </RibbonButton>
                    <RibbonButton disabled={!canEditSelection} label="Rotate text" onClick={() => applyAlignmentStyle({ rotation: 45 })}>
                      Rotate
                    </RibbonButton>
                  </ButtonGroup>
                </div>
              </RibbonGroup>

              <RibbonGroup label="Number">
                <div className="flex flex-col gap-1">
                  <Select
                    disabled={!canEditSelection}
                    onValueChange={(value) => {
                      const format = NUMBER_FORMATS.find((item) => item.label === value);
                      if (format) {
                        applyStyle(format.style);
                      }
                    }}
                    value="General"
                  >
                    <SelectTrigger className="w-[124px]" size="sm">
                      <SelectValue placeholder="Format" />
                    </SelectTrigger>
                    <SelectContent align="start">
                      {NUMBER_FORMATS.map((format) => (
                        <SelectItem key={format.label} value={format.label}>
                          {format.label}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                  <ButtonGroup>
                    <RibbonButton disabled={!canEditSelection} label="Currency" onClick={() => applyStyle(NUMBER_FORMATS[2].style)}>
                      $
                    </RibbonButton>
                    <RibbonButton disabled={!canEditSelection} label="Percent" onClick={() => applyStyle(NUMBER_FORMATS[3].style)}>
                      %
                    </RibbonButton>
                    <RibbonButton disabled={!canEditSelection} label="Thousands" onClick={() => applyStyle({ numberFormat: { formatType: "custom", formatString: "#,##0" } })}>
                      000
                    </RibbonButton>
                  </ButtonGroup>
                </div>
              </RibbonGroup>

              <RibbonGroup label="Styles">
                <div className="grid grid-cols-2 gap-1">
                  {CELL_STYLE_PRESETS.map((preset) => (
                    <StyleSwatchButton
                      disabled={!canEditSelection}
                      key={preset.label}
                      label={preset.label}
                      onClick={() => applyStyle(preset.style)}
                      style={preset.style}
                    />
                  ))}
                </div>
              </RibbonGroup>

              <RibbonGroup label="Cells">
                <div className="flex flex-col gap-1">
                  <ButtonGroup>
                    <RibbonButton disabled={!hasSelection || isReadOnly} label="Merge selection" onClick={mergeSelection}>
                      Merge
                    </RibbonButton>
                    <RibbonButton disabled={!hasSelection || isReadOnly} label="Unmerge selection" onClick={unmergeSelection}>
                      Unmerge
                    </RibbonButton>
                  </ButtonGroup>
                  <DropdownMenu>
                    <DropdownMenuTrigger render={<Button disabled={!canEditSelection} size="sm" variant="outline" />}>
                      Borders
                      <PlaygroundIcon name="chevron-down" />
                    </DropdownMenuTrigger>
                    <DropdownMenuContent align="start" className="w-44">
                      <DropdownMenuLabel>Border color</DropdownMenuLabel>
                      <div className="px-2 py-1">
                        <ColorPicker
                          color={borderColor}
                          disabled={!canEditSelection}
                          onChange={setBorderColor}
                          triggerClassName="h-7 w-full"
                        />
                      </div>
                      <DropdownMenuSeparator />
                      <DropdownMenuItem onClick={() => applyBorderStyle({ bottom: { style: "thin", color: hexToStyleColor(borderColor) } })}>
                        Bottom border
                      </DropdownMenuItem>
                      <DropdownMenuItem onClick={() => applyBorderStyle({
                        bottom: { style: "thin", color: hexToStyleColor(borderColor) },
                        left: { style: "thin", color: hexToStyleColor(borderColor) },
                        right: { style: "thin", color: hexToStyleColor(borderColor) },
                        top: { style: "thin", color: hexToStyleColor(borderColor) },
                      })}>
                        All borders
                      </DropdownMenuItem>
                      <DropdownMenuItem onClick={() => applyBorderStyle({ bottom: { style: "none" }, left: { style: "none" }, right: { style: "none" }, top: { style: "none" } })}>
                        No border
                      </DropdownMenuItem>
                    </DropdownMenuContent>
                  </DropdownMenu>
                </div>
              </RibbonGroup>
            </>
          ) : null}

          {activeRibbonTab === "Insert" ? (
            <>
              <RibbonGroup label="Workbook">
                <ButtonGroup>
                  <RibbonButton disabled={!hasWorkbook || isReadOnly} label="Add sheet" onClick={() => addSheet()}>
                    <PlaygroundIcon name="plus" />
                    Sheet
                  </RibbonButton>
                  <RibbonButton disabled={sheets.length <= 1 || isReadOnly} label="Delete active sheet" onClick={removeActiveSheet}>
                    <PlaygroundIcon name="trash" />
                    Delete
                  </RibbonButton>
                </ButtonGroup>
              </RibbonGroup>
              <RibbonGroup label="Open">
                <Button onClick={onOpenFile} size="sm">
                  <PlaygroundIcon name="open" />
                  Open
                </Button>
                <Button onClick={onLoadExampleUrl} size="sm" variant="outline">
                  Sample
                </Button>
              </RibbonGroup>
              <RibbonGroup label="Source">
                <Input
                  className="w-[280px]"
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
                  <PlaygroundIcon name="link" />
                  Load
                </Button>
              </RibbonGroup>
            </>
          ) : null}

          {activeRibbonTab === "Page Layout" ? (
            <>
              <RibbonGroup label="Theme">
                <div className="flex items-center gap-1.5 rounded-md border px-2 py-1">
                  <span className="text-muted-foreground text-[11px] font-medium">Document dark</span>
                  <Switch aria-label="Toggle document dark mode" checked={isDocumentDark} onCheckedChange={setIsDocumentDark} size="sm" />
                </div>
                <ThemeToggle />
              </RibbonGroup>
              <RibbonGroup label="Export">
                <Button disabled={!canDownload} onClick={download} size="sm" variant="outline">
                  <PlaygroundIcon name="download" />
                  Source
                </Button>
                <Button disabled={!canExport} onClick={exportXlsx} size="sm" variant="outline">
                  <PlaygroundIcon name="download" />
                  XLSX
                </Button>
                <Button disabled={!canExport} onClick={exportCsv} size="sm" variant="outline">
                  <PlaygroundIcon name="download" />
                  CSV
                </Button>
              </RibbonGroup>
            </>
          ) : null}

          {activeRibbonTab === "Formulas" ? (
            <>
              <RibbonGroup label="Calculation">
                <Button disabled={!canExport} onClick={recalculate} size="sm" variant="outline">
                  <PlaygroundIcon name="refresh" />
                  Recalc
                </Button>
              </RibbonGroup>
              <RibbonGroup label="Defined Names">
                <Input
                  className="w-[148px]"
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
            </>
          ) : null}

          {activeRibbonTab === "Data" ? (
            <>
              <RibbonGroup label="Tables">
                <div className="text-muted-foreground max-w-[220px] px-1 text-xs leading-5">
                  Table sort controls appear in table header menus.
                </div>
              </RibbonGroup>
              <RibbonGroup label="Refresh">
                <Button disabled={!canExport} onClick={recalculate} size="sm" variant="outline">
                  <PlaygroundIcon name="refresh" />
                  Recalc
                </Button>
              </RibbonGroup>
              <RibbonGroup label="Workbook">
                <Button disabled={!hasWorkbook} onClick={onClear} size="sm" variant="outline">
                  <PlaygroundIcon name="trash" />
                  Clear
                </Button>
              </RibbonGroup>
            </>
          ) : null}

          {activeRibbonTab === "View" ? (
            <>
              <RibbonGroup label="Zoom">
                <ButtonGroup>
                  <RibbonButton disabled={!hasWorkbook || !canZoomOut} label="Zoom out" onClick={zoomOut}>
                    <PlaygroundIcon name="minus" />
                  </RibbonButton>
                  <Select disabled={!hasWorkbook} onValueChange={(value) => setZoomScale(Number(value))} value={String(Math.round(zoomScale))}>
                    <SelectTrigger className="w-[92px]" size="sm">
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
                  <RibbonButton disabled={!hasWorkbook || !canZoomIn} label="Zoom in" onClick={zoomIn}>
                    <PlaygroundIcon name="plus" />
                  </RibbonButton>
                </ButtonGroup>
                <Button disabled={!hasWorkbook || Math.round(zoomScale) === Math.round(defaultZoomScale)} onClick={resetZoom} size="sm" variant="outline">
                  Reset
                </Button>
              </RibbonGroup>
              <RibbonGroup label="Sheets">
                <ButtonGroup>
                  <RibbonButton disabled={!activeSheet || activeSheetIndex <= 0} label="Previous sheet" onClick={() => setActiveSheetIndex(activeSheetIndex - 1)}>
                    <PlaygroundIcon name="chevron-left" />
                  </RibbonButton>
                  <RibbonButton disabled={!activeSheet || activeSheetIndex >= sheets.length - 1} label="Next sheet" onClick={() => setActiveSheetIndex(activeSheetIndex + 1)}>
                    <PlaygroundIcon name="chevron-right" />
                  </RibbonButton>
                </ButtonGroup>
                <Select disabled={sheets.length === 0} onValueChange={(value) => setActiveSheetIndex(Number(value))} value={String(activeSheetIndex)}>
                  <SelectTrigger className="w-[160px]" size="sm">
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
              <RibbonGroup label="Display">
                <div className="flex items-center gap-1.5 rounded-md border px-2 py-1">
                  <span className="text-muted-foreground text-[11px] font-medium">Canvas</span>
                  <Switch aria-label="Toggle experimental canvas renderer" checked={experimentalCanvas} onCheckedChange={setExperimentalCanvas} size="sm" />
                </div>
                <div className="flex items-center gap-1.5 rounded-md border px-2 py-1">
                  <span className="text-muted-foreground text-[11px] font-medium">Highlight</span>
                  <Switch aria-label="Toggle custom cell highlighting via getCellStyle" checked={highlightCells} onCheckedChange={setHighlightCells} size="sm" />
                </div>
                <div className="flex items-center gap-1.5 rounded-md border px-2 py-1">
                  <span className="text-muted-foreground text-[11px] font-medium">Read only</span>
                  <Switch aria-label="Toggle read only mode" checked={isReadOnly} onCheckedChange={setReadOnly} size="sm" />
                </div>
              </RibbonGroup>
            </>
          ) : null}
        </div>
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
  const [experimentalCanvas, setExperimentalCanvas] = React.useState(true);
  const [isReadOnly, setIsReadOnly] = React.useState(false);
  const [highlightCells, setHighlightCells] = React.useState(false);
  const dragDepthRef = React.useRef(0);

  const getCellStyle = React.useCallback<NonNullable<XlsxViewerProps["getCellStyle"]>>(
    ({ cell, isTableHeader }) => {
      if (!highlightCells || isTableHeader) {
        return undefined;
      }
      if (cell.row % 2 === 1) {
        return { backgroundColor: isDocumentDark ? "rgba(56, 189, 248, 0.16)" : "rgba(37, 99, 235, 0.08)" };
      }
      return undefined;
    },
    [highlightCells, isDocumentDark]
  );

  const controller = useXlsxViewerController(
    source?.type === "file"
      ? {
          allowResizeInReadOnly: true,
          file: source.file,
          fileName: source.fileName,
          readOnly: isReadOnly,
          readOnlyAboveBytes: AUTO_READ_ONLY_THRESHOLD_BYTES
        }
      : source?.type === "url"
        ? {
            allowResizeInReadOnly: true,
            src: source.src,
            fileName: source.fileName,
            readOnly: isReadOnly,
            readOnlyAboveBytes: AUTO_READ_ONLY_THRESHOLD_BYTES
          }
        : {
            allowResizeInReadOnly: true,
            readOnly: isReadOnly,
            readOnlyAboveBytes: AUTO_READ_ONLY_THRESHOLD_BYTES
          }
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
              highlightCells={highlightCells}
              isDocumentDark={isDocumentDark}
              onClear={handleClear}
              onLoadExampleUrl={handleLoadExampleUrl}
              onLoadUrl={handleLoadUrl}
              onOpenFile={() => fileInputRef.current?.click()}
              readOnly={isReadOnly}
              remoteUrl={remoteUrl}
              setExperimentalCanvas={setExperimentalCanvas}
              setHighlightCells={setHighlightCells}
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
                  getCellStyle={getCellStyle}
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
                        <PlaygroundIcon className="size-3" name="chevron-down" />
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
