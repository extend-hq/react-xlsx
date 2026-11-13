import * as React from "react";
import {
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  Baseline,
  Bold,
  Calculator,
  Check,
  ChevronLeft,
  ChevronRight,
  Code,
  DollarSign,
  Download,
  FileSpreadsheet,
  FolderOpen,
  Grid2x2,
  Italic,
  Link2,
  Minus,
  Moon,
  PaintBucket,
  Paintbrush,
  Percent,
  Plus,
  Redo2,
  RefreshCw,
  RotateCw,
  Sparkles,
  Strikethrough,
  Sun,
  TableCellsMerge,
  Tag,
  TextAlignCenter,
  TextAlignEnd,
  TextAlignStart,
  TextWrap,
  Trash2,
  Underline,
  Undo2,
  Upload,
  X,
} from "lucide-react";
import { useTheme } from "next-themes";
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
  type XlsxViewerProps,
} from "@extend-ai/react-xlsx";
import { FormulaEditorDemo } from "@/components/formula-editor-demo";
import { FormulaContextMenu } from "@/components/formula-context-menu";
import { WorksheetHeaderMenu } from "@/components/worksheet-header-menu";
import { InsertPivotButton, PivotPlaygroundProvider, PivotPlaygroundLayout, PivotRefreshAllButton, PivotTableRibbon } from "@/components/pivot-tables";
import type { DemoStylePreview } from "@/components/formula-demo-actions";
import { HIGHLIGHT_DEMO_RANGES, HIGHLIGHT_DEMO_URL, RangeHighlightDemo } from "@/components/range-highlight-demo";
import { RibbonButton, RibbonColorButton, RibbonDivider, RibbonGroup, RibbonMenuButton } from "@/components/ribbon";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Menu, MenuGroupLabel, MenuItem, MenuPopup, MenuSeparator, MenuTrigger } from "@/components/ui/menu";
import { Select, SelectItem, SelectPopup, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Slider } from "@/components/ui/slider";
import { Spinner } from "@/components/ui/spinner";
import { Switch } from "@/components/ui/switch";
import { Tabs, TabsList, TabsPanel, TabsTab } from "@/components/ui/tabs";
import { Tooltip, TooltipPopup, TooltipTrigger } from "@/components/ui/tooltip";
import { cn } from "@/lib/utils";
import { nativeDrawingClipboardAdapter } from "@/lib/native-clipboard";

const AUTO_READ_ONLY_THRESHOLD_BYTES = 5 * 1024 * 1024;
const PLAYGROUND_SAMPLE_URL = "/examples/crazy-chart-zoo-subtypes.xlsx";
const PIVOT_SAMPLE_URL = "/examples/pivot-tables.xlsx";

const RIBBON_TABS = ["Home", "Insert", "PivotTable", "Formulas", "Data", "Review", "View", "Developer"] as const;
type RibbonTab = (typeof RIBBON_TABS)[number];

const FONT_FAMILIES = ["Aptos", "Calibri", "Arial", "Georgia", "Times New Roman", "Courier New"] as const;
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72] as const;
const LOCALES = [
  { label: "English (US)", value: "en-US" },
  { label: "English (UK)", value: "en-GB" },
  { label: "French (France)", value: "fr-FR" },
  { label: "German (Germany)", value: "de-DE" },
  { label: "Japanese (Japan)", value: "ja-JP" },
];
const NUMBER_FORMATS = [
  { label: "General", sample: "abc", style: { numberFormat: { formatType: "general" } } },
  { label: "Number", sample: "1,234.00", style: { numberFormat: { formatType: "custom", formatString: "#,##0.00" } } },
  { label: "Currency", sample: "$1,234.00", style: { numberFormat: { formatType: "custom", formatString: "$#,##0.00" } } },
  { label: "Percentage", sample: "12.00%", style: { numberFormat: { formatType: "custom", formatString: "0.00%" } } },
  { label: "Short Date", sample: "1/31/2026", style: { numberFormat: { formatType: "custom", formatString: "m/d/yyyy" } } },
  { label: "Time", sample: "1:30 PM", style: { numberFormat: { formatType: "custom", formatString: "h:mm AM/PM" } } },
] satisfies Array<{ label: string; sample: string; style: XlsxCellStyleInput }>;
const CELL_STYLE_PRESETS = [
  { label: "Good", fill: "C6EFCE", color: "006100" },
  { label: "Neutral", fill: "FFEB9C", color: "9C5700" },
  { label: "Bad", fill: "FFC7CE", color: "9C0006" },
].map(({ color, fill, label }) => ({
  label,
  fill,
  color,
  style: {
    fill: { fillType: "solid", color: { colorType: "rgb", hex: fill } },
    font: { color: { colorType: "rgb", hex: color } },
  } satisfies XlsxCellStyleInput,
}));
const HEADING_STYLE: XlsxCellStyleInput = {
  border: { bottom: { style: "thick", color: { colorType: "rgb", hex: "4472C4" } } },
  font: { bold: true, color: { colorType: "rgb", hex: "44546A" }, size: 15 },
};

type ViewerSource =
  | { file: ArrayBuffer; fileName: string; type: "file" }
  | { fileName?: string; src: string; type: "url" }
  | null;

type PlaygroundSettings = {
  allowResizeInReadOnly: boolean;
  bandedRows: boolean;
  customHighlightStyle: boolean;
  documentDark: boolean;
  experimentalCanvas: boolean;
  locale: string;
  readOnly: boolean;
  useWorker: boolean;
};

type SettingsContextValue = PlaygroundSettings & {
  set: <K extends keyof PlaygroundSettings>(key: K, value: PlaygroundSettings[K]) => void;
};

const SettingsContext = React.createContext<SettingsContextValue | null>(null);

function useSettings() {
  const value = React.useContext(SettingsContext);
  if (!value) throw new Error("useSettings must be used inside <App>");
  return value;
}

function hexToStyleColor(hex: string): NonNullable<XlsxCellStyleInput["font"]>["color"] {
  return { colorType: "rgb", hex: hex.replace("#", "").toUpperCase() };
}

function isXlsxWorkerScriptResource(name: string) {
  return name.includes("xlsx-worker.js?worker_file") || /\/xlsx-worker(?:-[^/]+)?\.js(?:$|\?)/.test(name);
}

function useXlsxWorkerScriptLoadCount() {
  const [count, setCount] = React.useState(0);

  React.useEffect(() => {
    if (typeof performance === "undefined") return;
    const read = () => {
      const next = performance.getEntriesByType("resource").filter((entry) => isXlsxWorkerScriptResource(entry.name)).length;
      setCount((current) => (current === next ? current : next));
    };
    read();
    const id = window.setInterval(read, 1000);
    return () => window.clearInterval(id);
  }, []);

  return count;
}

const iconButtonClass =
  "flex items-center justify-center rounded-sm text-muted-foreground hover:bg-black/5 hover:text-foreground disabled:pointer-events-none disabled:opacity-40 dark:hover:bg-white/5";

/* -------------------------------------------------------------------------- */
/*                                  Title bar                                 */
/* -------------------------------------------------------------------------- */

function TitleBar({ onOpenFile }: { onOpenFile: () => void }) {
  const { canDownload, canExport, displayFileName, download, exportCsv, exportXlsx, sheets } = useXlsxViewer();
  const { readOnly } = useSettings();
  const { resolvedTheme, setTheme } = useTheme();
  const isDark = resolvedTheme === "dark";
  const titleButton = "h-7 text-[13px] font-normal sm:h-7";

  return (
    <header className="flex h-11 shrink-0 items-center gap-3 px-3">
      <div className="flex min-w-0 flex-1 items-center gap-2.5">
        <div className="flex size-7 shrink-0 items-center justify-center rounded-md bg-brand text-brand-foreground">
          <FileSpreadsheet className="size-4" />
        </div>
        <span className="truncate text-[13px] font-semibold">{sheets.length > 0 ? displayFileName : "Book1"}</span>
        {readOnly ? (
          <Badge className="rounded px-1.5 text-[10px]" variant="outline">
            Read-only
          </Badge>
        ) : null}
        <span className="hidden truncate text-xs text-muted-foreground md:inline">React XLSX Playground</span>
      </div>

      <div className="flex shrink-0 items-center gap-0.5">
        <Button className={titleButton} onClick={onOpenFile} size="sm" variant="ghost">
          <FolderOpen />
          Open
        </Button>
        <Menu>
          <MenuTrigger render={<Button className={titleButton} disabled={!canExport && !canDownload} size="sm" variant="ghost" />}>
            <Download />
            Download
          </MenuTrigger>
          <MenuPopup align="end" className="w-56">
            <MenuItem disabled={!canExport} onClick={exportXlsx}>
              Excel workbook (.xlsx)
            </MenuItem>
            <MenuItem disabled={!canExport} onClick={exportCsv}>
              CSV, current sheet (.csv)
            </MenuItem>
            <MenuSeparator />
            <MenuItem disabled={!canDownload} onClick={download}>
              Original file
            </MenuItem>
          </MenuPopup>
        </Menu>
        <Button className={titleButton} render={<a href="/vanilla" />} size="sm" variant="ghost">
          <Code />
          Vanilla JS
        </Button>
        <RibbonDivider />
        <Tooltip>
          <TooltipTrigger
            render={
              <Button
                aria-label={isDark ? "Switch to light mode" : "Switch to dark mode"}
                className="size-7 sm:size-7"
                onClick={() => setTheme(isDark ? "light" : "dark")}
                size="icon-sm"
                variant="ghost"
              />
            }
          >
            {isDark ? <Sun /> : <Moon />}
          </TooltipTrigger>
          <TooltipPopup side="bottom">{isDark ? "Light mode" : "Dark mode"}</TooltipPopup>
        </Tooltip>
      </div>
    </header>
  );
}

/* -------------------------------------------------------------------------- */
/*                                   Ribbon                                   */
/* -------------------------------------------------------------------------- */

function useStyleActions() {
  const { readOnly } = useSettings();
  const { activeCellAddress, selection } = useXlsxViewerSelection();
  const { readOnly: viewerReadOnly, setRangeStyle, setSelectedCellStyle } = useXlsxViewerEditing();
  const isReadOnly = readOnly || viewerReadOnly;
  const canEdit = Boolean(activeCellAddress) && !isReadOnly;

  const applyStyle = React.useCallback(
    (style: XlsxCellStyleInput) => {
      if (!canEdit) return;
      if (selection) setRangeStyle(selection, style);
      else setSelectedCellStyle(style);
    },
    [canEdit, selection, setRangeStyle, setSelectedCellStyle],
  );

  return { applyStyle, canEdit, hasSelection: Boolean(selection), isReadOnly };
}

function HomeRibbon() {
  const { applyStyle, canEdit, hasSelection, isReadOnly } = useStyleActions();
  const { canRedo, canUndo, mergeSelection, redo, undo, unmergeSelection } = useXlsxViewerEditing();
  const [fontFamily, setFontFamily] = React.useState<string>("Aptos");
  const [fontSize, setFontSize] = React.useState("11");
  const [fontColor, setFontColor] = React.useState("#C00000");
  const [fillColor, setFillColor] = React.useState("#FFFF00");
  const [borderColor, setBorderColor] = React.useState("#000000");
  const font = (style: NonNullable<XlsxCellStyleInput["font"]>) => applyStyle({ font: style });
  const align = (style: NonNullable<XlsxCellStyleInput["alignment"]>) => applyStyle({ alignment: style });
  const border = (style: NonNullable<XlsxCellStyleInput["border"]>) => applyStyle({ border: style });
  const thin = { style: "thin" as const, color: hexToStyleColor(borderColor) };

  return (
    <>
      <RibbonGroup>
        <RibbonButton disabled={!canUndo} label="Undo" onClick={undo} shortcut="⌘Z">
          <Undo2 />
        </RibbonButton>
        <RibbonButton disabled={!canRedo} label="Redo" onClick={redo} shortcut="⌘⇧Z">
          <Redo2 />
        </RibbonButton>
      </RibbonGroup>

      <RibbonGroup className="gap-1">
        <Select
          disabled={!canEdit}
          onValueChange={(value) => {
            if (!value) return;
            setFontFamily(value);
            font({ name: value });
          }}
          value={fontFamily}
        >
          <SelectTrigger aria-label="Font" className="w-36 min-w-0 text-[13px]" size="sm">
            <SelectValue />
          </SelectTrigger>
          <SelectPopup>
            {FONT_FAMILIES.map((family) => (
              <SelectItem key={family} style={{ fontFamily: family }} value={family}>
                {family}
              </SelectItem>
            ))}
          </SelectPopup>
        </Select>
        <Select
          disabled={!canEdit}
          onValueChange={(value) => {
            if (!value) return;
            setFontSize(value);
            font({ size: Number(value) });
          }}
          value={fontSize}
        >
          <SelectTrigger aria-label="Font size" className="w-16 min-w-0 text-[13px]" size="sm">
            <SelectValue />
          </SelectTrigger>
          <SelectPopup>
            {FONT_SIZES.map((size) => (
              <SelectItem key={size} value={String(size)}>
                {size}
              </SelectItem>
            ))}
          </SelectPopup>
        </Select>
      </RibbonGroup>

      <RibbonGroup>
        <RibbonButton disabled={!canEdit} label="Bold" onClick={() => font({ bold: true })} shortcut="⌘B">
          <Bold />
        </RibbonButton>
        <RibbonButton disabled={!canEdit} label="Italic" onClick={() => font({ italic: true })} shortcut="⌘I">
          <Italic />
        </RibbonButton>
        <RibbonButton disabled={!canEdit} label="Underline" onClick={() => font({ underline: "single" })} shortcut="⌘U">
          <Underline />
        </RibbonButton>
        <RibbonButton disabled={!canEdit} label="Strikethrough" onClick={() => font({ strikethrough: true })}>
          <Strikethrough />
        </RibbonButton>
        <Menu>
          <MenuTrigger render={<RibbonMenuButton disabled={!canEdit} icon={<Grid2x2 />} label="Borders" showLabel={false} />} />
          <MenuPopup align="start" className="w-52">
            <MenuItem onClick={() => border({ bottom: thin })}>Bottom border</MenuItem>
            <MenuItem onClick={() => border({ top: thin })}>Top border</MenuItem>
            <MenuItem onClick={() => border({ bottom: thin, left: thin, right: thin, top: thin })}>All borders</MenuItem>
            <MenuItem onClick={() => border({ bottom: { style: "thick", color: hexToStyleColor(borderColor) } })}>
              Thick bottom border
            </MenuItem>
            <MenuSeparator />
            <MenuItem
              onClick={() =>
                border({ bottom: { style: "none" }, left: { style: "none" }, right: { style: "none" }, top: { style: "none" } })
              }
            >
              No border
            </MenuItem>
            <MenuSeparator />
            <MenuGroupLabel>Line color</MenuGroupLabel>
            <label className="flex cursor-pointer items-center gap-2 rounded-sm px-2 py-1 text-sm hover:bg-accent">
              <input
                className="size-5 cursor-pointer rounded border-0 bg-transparent p-0"
                onChange={(event) => setBorderColor(event.target.value.toUpperCase())}
                type="color"
                value={borderColor}
              />
              <span className="font-mono text-xs">{borderColor}</span>
            </label>
          </MenuPopup>
        </Menu>
        <RibbonColorButton
          color={fillColor}
          disabled={!canEdit}
          icon={<PaintBucket />}
          label="Fill color"
          onApply={(color) => applyStyle({ fill: { fillType: "solid", color: hexToStyleColor(color) } })}
          onColorChange={setFillColor}
        />
        <RibbonColorButton
          color={fontColor}
          disabled={!canEdit}
          icon={<Baseline />}
          label="Font color"
          onApply={(color) => font({ color: hexToStyleColor(color) })}
          onColorChange={setFontColor}
        />
      </RibbonGroup>

      <RibbonGroup>
        <RibbonButton disabled={!canEdit} label="Align left" onClick={() => align({ horizontal: "left" })}>
          <TextAlignStart />
        </RibbonButton>
        <RibbonButton disabled={!canEdit} label="Center" onClick={() => align({ horizontal: "center" })}>
          <TextAlignCenter />
        </RibbonButton>
        <RibbonButton disabled={!canEdit} label="Align right" onClick={() => align({ horizontal: "right" })}>
          <TextAlignEnd />
        </RibbonButton>
        <RibbonButton disabled={!canEdit} label="Top align" onClick={() => align({ vertical: "top" })}>
          <AlignVerticalJustifyStart />
        </RibbonButton>
        <RibbonButton disabled={!canEdit} label="Middle align" onClick={() => align({ vertical: "center" })}>
          <AlignVerticalJustifyCenter />
        </RibbonButton>
        <RibbonButton disabled={!canEdit} label="Bottom align" onClick={() => align({ vertical: "bottom" })}>
          <AlignVerticalJustifyEnd />
        </RibbonButton>
        <RibbonButton disabled={!canEdit} label="Wrap text" onClick={() => align({ wrapText: true })}>
          <TextWrap />
        </RibbonButton>
        <RibbonButton disabled={!canEdit} label="Rotate text 45°" onClick={() => align({ rotation: 45 })}>
          <RotateCw />
        </RibbonButton>
        <Menu>
          <MenuTrigger
            render={<RibbonMenuButton disabled={!hasSelection || isReadOnly} icon={<TableCellsMerge />} label="Merge" />}
          />
          <MenuPopup align="start" className="w-44">
            <MenuItem onClick={() => mergeSelection("center")}>Merge &amp; Center</MenuItem>
            <MenuItem onClick={() => mergeSelection("across")}>Merge Across</MenuItem>
            <MenuItem onClick={() => mergeSelection()}>Merge Cells</MenuItem>
            <MenuSeparator />
            <MenuItem onClick={unmergeSelection}>Unmerge Cells</MenuItem>
          </MenuPopup>
        </Menu>
      </RibbonGroup>

      <RibbonGroup>
        <Menu>
          <MenuTrigger render={<RibbonMenuButton className="w-28 justify-between" disabled={!canEdit} label="General" />} />
          <MenuPopup align="start" className="w-56">
            {NUMBER_FORMATS.map((format) => (
              <MenuItem className="justify-between" key={format.label} onClick={() => applyStyle(format.style)}>
                {format.label}
                <span className="text-xs text-muted-foreground tabular-nums">{format.sample}</span>
              </MenuItem>
            ))}
          </MenuPopup>
        </Menu>
        <RibbonButton disabled={!canEdit} label="Currency format" onClick={() => applyStyle(NUMBER_FORMATS[2].style)}>
          <DollarSign />
        </RibbonButton>
        <RibbonButton disabled={!canEdit} label="Percent style" onClick={() => applyStyle(NUMBER_FORMATS[3].style)}>
          <Percent />
        </RibbonButton>
        <RibbonButton
          className="text-[11px] font-medium tabular-nums"
          disabled={!canEdit}
          label="Comma style"
          onClick={() => applyStyle({ numberFormat: { formatType: "custom", formatString: "#,##0" } })}
        >
          ,000
        </RibbonButton>
      </RibbonGroup>

      <RibbonGroup>
        <Menu>
          <MenuTrigger render={<RibbonMenuButton disabled={!canEdit} icon={<Paintbrush />} label="Cell Styles" />} />
          <MenuPopup align="start" className="w-64">
            <MenuGroupLabel>Good, Bad and Neutral</MenuGroupLabel>
            <div className="grid grid-cols-3 gap-1 px-1 pb-1">
              {CELL_STYLE_PRESETS.map((preset) => (
                <MenuItem
                  className="h-8 justify-center border border-black/10 text-xs data-highlighted:ring-2 data-highlighted:ring-brand/40"
                  key={preset.label}
                  onClick={() => applyStyle(preset.style)}
                  style={{ backgroundColor: `#${preset.fill}`, color: `#${preset.color}` }}
                >
                  {preset.label}
                </MenuItem>
              ))}
            </div>
            <MenuGroupLabel>Titles and Headings</MenuGroupLabel>
            <MenuItem className="mx-1 mb-1 font-semibold text-[#44546A]" onClick={() => applyStyle(HEADING_STYLE)}>
              <span className="border-b-2 border-[#4472C4]">Heading 1</span>
            </MenuItem>
          </MenuPopup>
        </Menu>
      </RibbonGroup>
    </>
  );
}

function InsertRibbon({
  onLoadSample,
  onLoadUrl,
  onOpenFile,
}: {
  onLoadSample: () => void;
  onLoadUrl: (url: string) => void;
  onOpenFile: () => void;
}) {
  const { sheets } = useXlsxViewer();
  const { addSheet, removeActiveSheet } = useXlsxViewerEditing();
  const { isReadOnly } = useStyleActions();
  const [url, setUrl] = React.useState("");

  return (
    <>
      <RibbonGroup><InsertPivotButton /></RibbonGroup>
      <RibbonGroup>
        <RibbonButton disabled={sheets.length === 0 || isReadOnly} label="New sheet" onClick={() => addSheet()} showLabel>
          <Plus />
        </RibbonButton>
        <RibbonButton disabled={sheets.length <= 1 || isReadOnly} label="Delete sheet" onClick={removeActiveSheet} showLabel>
          <Trash2 />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup>
        <RibbonButton label="Upload workbook" onClick={onOpenFile} showLabel>
          <Upload />
        </RibbonButton>
        <RibbonButton label="Sample workbook" onClick={onLoadSample} showLabel>
          <Sparkles />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup>
        <form
          className="flex items-center gap-1.5"
          onSubmit={(event) => {
            event.preventDefault();
            onLoadUrl(url);
          }}
        >
          <Input
            aria-label="Workbook URL"
            className="w-72 text-xs"
            onChange={(event) => setUrl(event.target.value)}
            placeholder="https://example.com/report.xlsx"
            size="sm"
            value={url}
          />
          <RibbonButton disabled={!url.trim()} label="Load from URL" showLabel type="submit">
            <Link2 />
          </RibbonButton>
        </form>
      </RibbonGroup>
    </>
  );
}

function FormulasRibbon() {
  const { canExport, recalculate } = useXlsxViewer();
  const { selection } = useXlsxViewerSelection();
  const { defineNamedRange } = useXlsxViewerEditing();
  const { isReadOnly } = useStyleActions();
  const [name, setName] = React.useState("");
  const define = () => {
    const trimmed = name.trim();
    if (!trimmed || !selection) return;
    defineNamedRange(trimmed, selection);
    setName("");
  };

  return (
    <>
      <RibbonGroup>
        <RibbonButton disabled={!canExport} label="Calculate now" onClick={recalculate} shortcut="F9" showLabel>
          <Calculator />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup>
        <form
          className="flex items-center gap-1.5"
          onSubmit={(event) => {
            event.preventDefault();
            define();
          }}
        >
          <Input
            aria-label="Name for the selected range"
            className="w-44 text-xs"
            onChange={(event) => setName(event.target.value)}
            placeholder="Name the selected range"
            size="sm"
            value={name}
          />
          <RibbonButton disabled={!selection || !name.trim() || isReadOnly} label="Define name" showLabel type="submit">
            <Tag />
          </RibbonButton>
        </form>
      </RibbonGroup>
    </>
  );
}

function DataRibbon({ onClear }: { onClear: () => void }) {
  const { canExport, recalculate, sheets } = useXlsxViewer();

  return (
    <>
      <RibbonGroup>
        <RibbonButton disabled={!canExport} label="Refresh all" onClick={recalculate} showLabel>
          <RefreshCw />
        </RibbonButton>
        <PivotRefreshAllButton />
        <RibbonButton disabled={sheets.length === 0} label="Close workbook" onClick={onClear} showLabel>
          <X />
        </RibbonButton>
      </RibbonGroup>
      <span className="px-2 text-xs text-muted-foreground">Sort and filter from any table header menu.</span>
    </>
  );
}

function ToggleRow({ checked, label, onCheckedChange }: { checked: boolean; label: string; onCheckedChange: (checked: boolean) => void }) {
  return (
    <Label className="h-7 cursor-pointer gap-2 rounded-md px-1.5 text-[13px] font-normal hover:bg-accent">
      <Switch checked={checked} className="data-checked:bg-brand" onCheckedChange={onCheckedChange} />
      {label}
    </Label>
  );
}

function ReviewRibbon({ onLoadHighlightSample }: { onLoadHighlightSample: (autoScroll: boolean) => void }) {
  const settings = useSettings();

  return (
    <>
      <RangeHighlightDemo
        customStyle={settings.customHighlightStyle}
        onCustomStyleChange={(value) => settings.set("customHighlightStyle", value)}
        onLoadSample={onLoadHighlightSample}
      />
    </>
  );
}

function ViewRibbon() {
  const settings = useSettings();
  const { activeSheetIndex, setActiveSheetIndex, sheets } = useXlsxViewer();
  const { canZoomIn, canZoomOut, setZoomScale, zoomIn, zoomOut } = useXlsxViewerZoom();
  const hasWorkbook = sheets.length > 0;

  return (
    <>
      <RibbonGroup>
        <RibbonButton disabled={!hasWorkbook || !canZoomOut} label="Zoom out" onClick={zoomOut}>
          <Minus />
        </RibbonButton>
        <RibbonButton disabled={!hasWorkbook} label="100%" onClick={() => setZoomScale(100)} showLabel />
        <RibbonButton disabled={!hasWorkbook || !canZoomIn} label="Zoom in" onClick={zoomIn}>
          <Plus />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup>
        <ToggleRow
          checked={settings.experimentalCanvas}
          label="Canvas renderer"
          onCheckedChange={(value) => settings.set("experimentalCanvas", value)}
        />
        <ToggleRow checked={settings.bandedRows} label="Banded rows" onCheckedChange={(value) => settings.set("bandedRows", value)} />
        <ToggleRow checked={settings.documentDark} label="Dark sheet" onCheckedChange={(value) => settings.set("documentDark", value)} />
      </RibbonGroup>
      <RibbonGroup>
        <Select items={LOCALES} onValueChange={(value) => value && settings.set("locale", value)} value={settings.locale}>
          <SelectTrigger aria-label="Locale" className="w-40 min-w-0 text-[13px]" size="sm">
            <SelectValue />
          </SelectTrigger>
          <SelectPopup>
            {LOCALES.map((locale) => (
              <SelectItem key={locale.value} value={locale.value}>
                {locale.label}
              </SelectItem>
            ))}
          </SelectPopup>
        </Select>
      </RibbonGroup>
      <RibbonGroup>
        <RibbonButton disabled={activeSheetIndex <= 0} label="Previous sheet" onClick={() => setActiveSheetIndex(activeSheetIndex - 1)}>
          <ChevronLeft />
        </RibbonButton>
        <RibbonButton
          disabled={activeSheetIndex >= sheets.length - 1}
          label="Next sheet"
          onClick={() => setActiveSheetIndex(activeSheetIndex + 1)}
        >
          <ChevronRight />
        </RibbonButton>
      </RibbonGroup>
    </>
  );
}

function useRuntimeStatus() {
  const { isLoadDeferred, isLoading, sheets } = useXlsxViewer();
  const { useWorker } = useSettings();
  const workerApiAvailable = typeof Worker !== "undefined";
  const label = isLoading
    ? "Loading"
    : isLoadDeferred
      ? "Deferred"
      : sheets.length > 0
        ? "Ready"
        : "Idle";
  const detail = !useWorker || !workerApiAvailable
    ? "Processing runs on the main thread."
    : "Workers are preferred, with automatic main-thread fallback.";
  const tone = sheets.length > 0 ? "bg-emerald-500" : "bg-neutral-400";
  return { detail, label, tone };
}

function DeveloperRibbon() {
  const settings = useSettings();
  const runtime = useRuntimeStatus();
  const scriptLoads = useXlsxWorkerScriptLoadCount();

  return (
    <>
      <RibbonGroup>
        <ToggleRow checked={settings.useWorker} label="Use web workers" onCheckedChange={(value) => settings.set("useWorker", value)} />
        <ToggleRow
          checked={settings.allowResizeInReadOnly}
          label="Resize when read-only"
          onCheckedChange={(value) => settings.set("allowResizeInReadOnly", value)}
        />
      </RibbonGroup>
      <RibbonGroup className="gap-4 px-3 text-xs">
        <span className="flex items-center gap-1.5">
          <span className={cn("size-1.5 rounded-full", runtime.tone)} />
          <span className="text-muted-foreground">Runtime</span>
          <span className="font-medium">{runtime.label}</span>
        </span>
        <span className="flex items-center gap-1.5">
          <span className="text-muted-foreground">Worker script loads</span>
          <span className="font-medium tabular-nums">{scriptLoads}</span>
        </span>
        <span className="text-muted-foreground">{runtime.detail}</span>
      </RibbonGroup>
    </>
  );
}

function Ribbon({
  onClear,
  onLoadPivotSample,
  onLoadHighlightSample,
  onLoadSample,
  onLoadUrl,
  onOpenFile,
}: {
  onClear: () => void;
  onLoadPivotSample: () => void;
  onLoadHighlightSample: (autoScroll: boolean) => void;
  onLoadSample: () => void;
  onLoadUrl: (url: string) => void;
  onOpenFile: () => void;
}) {
  const settings = useSettings();
  const [tab, setTab] = React.useState<RibbonTab>("Home");
  const panelClass = "flex items-center";

  return (
    <Tabs className="shrink-0 gap-0 px-2 pb-2" onValueChange={(value) => setTab(value as RibbonTab)} value={tab}>
      <div className="flex min-w-0 items-center gap-2">
        <div className="min-w-0 overflow-x-auto [scrollbar-width:none]">
          <TabsList
            className="h-8 gap-0.5 p-0 *:data-[slot=tab-indicator]:bg-brand *:data-[slot=tabs-tab]:hover:bg-black/5 dark:*:data-[slot=tabs-tab]:hover:bg-white/5"
            variant="underline"
          >
            {RIBBON_TABS.map((name) => (
              <TabsTab
                className="h-7 grow-0 rounded-sm px-2.5 text-[13px] font-normal text-foreground/75 hover:text-foreground data-active:font-semibold data-active:text-brand"
                key={name}
                value={name}
              >
                {name}
              </TabsTab>
            ))}
          </TabsList>
        </div>
        <div className="ms-auto flex shrink-0 items-center">
          <ToggleRow checked={settings.readOnly} label="Read only" onCheckedChange={(value) => settings.set("readOnly", value)} />
        </div>
      </div>
      <div className="mt-1 flex h-11 items-center overflow-x-auto rounded-lg border bg-background px-2 shadow-xs/5 [scrollbar-width:none]">
        <TabsPanel className={panelClass} value="Home">
          <HomeRibbon />
        </TabsPanel>
        <TabsPanel className={panelClass} value="Insert">
          <InsertRibbon onLoadSample={onLoadSample} onLoadUrl={onLoadUrl} onOpenFile={onOpenFile} />
        </TabsPanel>
        <TabsPanel className={panelClass} value="PivotTable">
          <PivotTableRibbon onLoadSample={onLoadPivotSample} />
        </TabsPanel>
        <TabsPanel className={panelClass} value="Formulas">
          <FormulasRibbon />
        </TabsPanel>
        <TabsPanel className={panelClass} value="Data">
          <DataRibbon onClear={onClear} />
        </TabsPanel>
        <TabsPanel className={panelClass} value="Review">
          <ReviewRibbon onLoadHighlightSample={onLoadHighlightSample} />
        </TabsPanel>
        <TabsPanel className={panelClass} value="View">
          <ViewRibbon />
        </TabsPanel>
        <TabsPanel className={panelClass} value="Developer">
          <DeveloperRibbon />
        </TabsPanel>
      </div>
    </Tabs>
  );
}

/* -------------------------------------------------------------------------- */
/*                                 Formula bar                                */
/* -------------------------------------------------------------------------- */

function FormulaBar() {
  const { readOnly } = useSettings();
  const { activeCell, getPivotTableAtCell } = useXlsxViewer();
  const { activeCellAddress, selectedRangeAddress } = useXlsxViewerSelection();
  const { readOnly: viewerReadOnly, selectedFormula, selectedFormulaTarget, setSelectedFormula } = useXlsxViewerEditing();
  const [draft, setDraft] = React.useState("");
  const [focused, setFocused] = React.useState(false);
  const initialRef = React.useRef("");
  const hasTarget = Boolean(activeCellAddress) || selectedFormulaTarget?.kind === "chartSeries";
  const isPivotCell = selectedFormulaTarget?.kind !== "chartSeries" && Boolean(activeCell && getPivotTableAtCell(activeCell));
  const disabled = !hasTarget || readOnly || viewerReadOnly || isPivotCell;
  const nameBox =
    selectedFormulaTarget?.kind === "chartSeries"
      ? `Series ${selectedFormulaTarget.seriesIndex + 1}`
      : (selectedRangeAddress?.split("!").pop() ?? activeCellAddress ?? "");

  React.useEffect(() => {
    if (!focused) setDraft(selectedFormula);
  }, [selectedFormula, activeCellAddress, focused, selectedFormulaTarget]);

  const commit = () => {
    if (!disabled && hasTarget && draft !== initialRef.current) setSelectedFormula(draft);
    setFocused(false);
  };
  const cancel = () => {
    setDraft(initialRef.current);
    setFocused(false);
  };

  return (
    <div className="flex h-8 shrink-0 items-center gap-1 border-y bg-background px-2">
      <input
        aria-label="Name box"
        className="h-6 w-28 shrink-0 rounded-sm border border-transparent px-1.5 font-mono text-xs outline-none hover:border-input focus:border-brand"
        readOnly
        value={nameBox}
      />
      <div className="h-5 w-px shrink-0 bg-border" />
      <div className="flex shrink-0 items-center">
        <button
          aria-label="Cancel"
          className={cn(iconButtonClass, "size-6")}
          disabled={!focused}
          onMouseDown={(event) => {
            event.preventDefault();
            cancel();
          }}
          type="button"
        >
          <X className="size-3.5" />
        </button>
        <button
          aria-label="Enter"
          className={cn(iconButtonClass, "size-6")}
          disabled={!focused}
          onMouseDown={(event) => {
            event.preventDefault();
            commit();
          }}
          type="button"
        >
          <Check className="size-3.5" />
        </button>
        <span className="flex h-6 w-7 items-center justify-center font-serif text-sm italic text-muted-foreground">fx</span>
      </div>
      <div className="h-5 w-px shrink-0 bg-border" />
      <input
        aria-label="Formula bar"
        className="h-full min-w-0 flex-1 bg-transparent px-2 font-mono text-[13px] outline-none disabled:cursor-not-allowed"
        disabled={disabled}
        title={isPivotCell ? "Change PivotTable values through its field settings or source data." : undefined}
        onBlur={() => focused && commit()}
        onChange={(event) => setDraft(event.target.value)}
        onFocus={() => {
          initialRef.current = draft;
          setFocused(true);
        }}
        onKeyDown={(event) => {
          if (event.key === "Enter") {
            event.preventDefault();
            commit();
            event.currentTarget.blur();
          } else if (event.key === "Escape") {
            event.preventDefault();
            cancel();
            event.currentTarget.blur();
          }
        }}
        value={draft}
      />
    </div>
  );
}

/* -------------------------------------------------------------------------- */
/*                             Sheet tabs + status                            */
/* -------------------------------------------------------------------------- */

function SheetTabs() {
  const { activeSheetIndex, setActiveSheetIndex, sheets } = useXlsxViewer();
  const { addSheet } = useXlsxViewerEditing();
  const { isReadOnly } = useStyleActions();
  const { thumbnails } = useXlsxViewerThumbnails({ includeHeaders: true, resolution: { maxHeight: 132, maxWidth: 200 } });

  return (
    <div className="flex h-8 shrink-0 items-stretch bg-chrome">
      <div className="flex items-center gap-0.5 px-1.5">
        <button
          aria-label="Previous sheet"
          className={cn(iconButtonClass, "size-6")}
          disabled={activeSheetIndex <= 0}
          onClick={() => setActiveSheetIndex(activeSheetIndex - 1)}
          type="button"
        >
          <ChevronLeft className="size-4" />
        </button>
        <button
          aria-label="Next sheet"
          className={cn(iconButtonClass, "size-6")}
          disabled={activeSheetIndex >= sheets.length - 1}
          onClick={() => setActiveSheetIndex(activeSheetIndex + 1)}
          type="button"
        >
          <ChevronRight className="size-4" />
        </button>
        <button
          aria-label="New sheet"
          className={cn(iconButtonClass, "ms-1 size-6 rounded-full")}
          disabled={sheets.length === 0 || isReadOnly}
          onClick={() => addSheet()}
          type="button"
        >
          <Plus className="size-4" />
        </button>
      </div>
      <div className="flex min-w-0 flex-1 items-stretch overflow-x-auto [scrollbar-width:none]" role="tablist">
        {sheets.map((sheet, index) => {
          const active = index === activeSheetIndex;
          return (
            <Tooltip key={sheet.name}>
              <TooltipTrigger
                delay={600}
                render={
                  <button
                    aria-selected={active}
                    className={cn(
                      "relative shrink-0 px-4 text-xs transition-colors",
                      active
                        ? "bg-background font-semibold text-brand shadow-[0_1px_2px_rgb(0_0_0/0.06)] after:absolute after:inset-x-3 after:bottom-1 after:h-0.5 after:rounded-full after:bg-brand"
                        : "text-foreground/75 hover:bg-black/5 hover:text-foreground dark:hover:bg-white/5",
                    )}
                    onClick={() => setActiveSheetIndex(index)}
                    role="tab"
                    type="button"
                  />
                }
              >
                {sheet.name}
              </TooltipTrigger>
              <TooltipPopup className="p-1" side="top">
                <SheetThumbnail name={sheet.name} thumbnail={thumbnails[index] ?? null} />
              </TooltipPopup>
            </Tooltip>
          );
        })}
      </div>
    </div>
  );
}

function SheetThumbnail({
  name,
  thumbnail,
}: {
  name: string;
  thumbnail: ReturnType<typeof useXlsxViewerThumbnails>["thumbnails"][number] | null;
}) {
  const canvasRef = React.useRef<HTMLCanvasElement | null>(null);

  React.useEffect(() => {
    thumbnail?.paint(canvasRef.current);
  }, [thumbnail]);

  return (
    <div className="flex w-[200px] flex-col gap-1.5 py-1">
      <div className="truncate text-[11px] font-medium">{name}</div>
      <div className="overflow-hidden rounded-sm border bg-white">
        {thumbnail ? (
          <canvas className="block h-auto w-full" height={thumbnail.height} ref={canvasRef} width={thumbnail.width} />
        ) : (
          <div className="flex h-[100px] items-center justify-center text-[11px] text-muted-foreground">No preview</div>
        )}
      </div>
    </div>
  );
}

function StatusBar() {
  const { isLoading, sheets } = useXlsxViewer();
  const { readOnly: viewerReadOnly } = useXlsxViewerEditing();
  const { selectedRangeAddress } = useXlsxViewerSelection();
  const { canZoomIn, canZoomOut, maxZoomScale, minZoomScale, resetZoom, setZoomScale, zoomIn, zoomOut, zoomScale } =
    useXlsxViewerZoom();
  const { readOnly } = useSettings();
  const runtime = useRuntimeStatus();
  const hasWorkbook = sheets.length > 0;
  const status = isLoading ? "Loading…" : readOnly || viewerReadOnly ? "Read-only" : "Ready";

  return (
    <footer className="flex h-7 shrink-0 items-center gap-4 border-t bg-chrome px-3 text-[11px] text-foreground/70">
      <span className="flex items-center gap-1.5">
        {isLoading ? <Spinner className="size-3" /> : null}
        {status}
      </span>
      {selectedRangeAddress ? <span className="hidden font-mono sm:inline">{selectedRangeAddress}</span> : null}
      <div className="flex-1" />
      <Tooltip>
        <TooltipTrigger render={<span className="hidden items-center gap-1.5 sm:flex" />}>
          <span className={cn("size-1.5 rounded-full", runtime.tone)} />
          {runtime.label}
        </TooltipTrigger>
        <TooltipPopup side="top">{runtime.detail}</TooltipPopup>
      </Tooltip>
      <div className="flex items-center gap-1">
        <button aria-label="Zoom out" className={cn(iconButtonClass, "size-5")} disabled={!hasWorkbook || !canZoomOut} onClick={zoomOut} type="button">
          <Minus className="size-3" />
        </button>
        <Slider
          aria-label="Zoom"
          className="w-24 **:data-[slot=slider-control]:min-w-0 **:data-[slot=slider-indicator]:bg-foreground/40 **:data-[slot=slider-thumb]:size-3 sm:**:data-[slot=slider-thumb]:size-3"
          disabled={!hasWorkbook}
          max={maxZoomScale}
          min={minZoomScale}
          onValueChange={(value) => setZoomScale(Array.isArray(value) ? value[0] : value)}
          value={Math.round(zoomScale)}
        />
        <button aria-label="Zoom in" className={cn(iconButtonClass, "size-5")} disabled={!hasWorkbook || !canZoomIn} onClick={zoomIn} type="button">
          <Plus className="size-3" />
        </button>
        <button
          aria-label="Reset zoom"
          className={cn(iconButtonClass, "h-5 w-10 justify-end px-1 tabular-nums text-foreground/70")}
          disabled={!hasWorkbook}
          onClick={resetZoom}
          type="button"
        >
          {Math.round(zoomScale)}%
        </button>
      </div>
    </footer>
  );
}

/* -------------------------------------------------------------------------- */
/*                                 Empty state                                */
/* -------------------------------------------------------------------------- */

function EmptyWorkbook({ onLoadSample, onOpenFile }: { onLoadSample: () => void; onOpenFile: () => void }) {
  return (
    <div className="flex h-full w-full items-center justify-center bg-background px-6">
      <div className="flex max-w-sm flex-col items-center text-center">
        <div className="mb-4 flex size-12 items-center justify-center rounded-xl bg-brand/10 text-brand">
          <FileSpreadsheet className="size-6" />
        </div>
        <h2 className="text-base font-semibold">Open a workbook</h2>
        <p className="mt-1 text-sm text-muted-foreground">Drop an .xlsx file anywhere, or pick one to get started.</p>
        <div className="mt-5 flex gap-2">
          <Button onClick={onOpenFile} size="sm">
            <Upload />
            Upload file
          </Button>
          <Button onClick={onLoadSample} size="sm" variant="outline">
            <Sparkles />
            Try the sample
          </Button>
        </div>
      </div>
    </div>
  );
}

/* -------------------------------------------------------------------------- */
/*                                     App                                    */
/* -------------------------------------------------------------------------- */

const DEFAULT_SETTINGS: PlaygroundSettings = {
  allowResizeInReadOnly: false,
  bandedRows: false,
  customHighlightStyle: false,
  documentDark: false,
  experimentalCanvas: true,
  locale: "en-US",
  readOnly: false,
  useWorker: true,
};

const SAMPLE_SOURCE: ViewerSource = { type: "url", src: PLAYGROUND_SAMPLE_URL, fileName: "crazy-chart-zoo-subtypes.xlsx" };

export function App() {
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [settings, setSettings] = React.useState(DEFAULT_SETTINGS);
  const [source, setSource] = React.useState<ViewerSource>(SAMPLE_SOURCE);
  const [isDragActive, setIsDragActive] = React.useState(false);
  const [stylePreview, setStylePreview] = React.useState<DemoStylePreview | null>(null);
  const initializedSourceRef = React.useRef<ViewerSource>(null);
  const pendingHighlightSampleRef = React.useRef<{ autoScroll: boolean } | null>(null);
  const dragDepthRef = React.useRef(0);

  const settingsValue = React.useMemo<SettingsContextValue>(
    () => ({ ...settings, set: (key, value) => setSettings((current) => ({ ...current, [key]: value })) }),
    [settings],
  );

  const getCellStyle = React.useCallback<NonNullable<XlsxViewerProps["getCellStyle"]>>(
    ({ cell, isTableHeader, sheetName }) => {
      const preview = stylePreview?.sheetName === sheetName ? stylePreview.cells.get(`${cell.row}:${cell.col}`) : undefined;
      if (preview) return preview;
      if (!settings.bandedRows || isTableHeader || cell.row % 2 !== 1) return undefined;
      return { backgroundColor: settings.documentDark ? "rgba(63, 178, 122, 0.14)" : "rgba(16, 124, 65, 0.06)" };
    },
    [settings.bandedRows, settings.documentDark, stylePreview],
  );

  const sharedOptions = {
    drawingClipboardAdapter: nativeDrawingClipboardAdapter,
    allowResizeInReadOnly: settings.allowResizeInReadOnly,
    locale: settings.locale,
    readOnly: settings.readOnly,
    readOnlyAboveBytes: AUTO_READ_ONLY_THRESHOLD_BYTES,
    useWorker: settings.useWorker,
  };
  const controller = useXlsxViewerController(
    source?.type === "file"
      ? { ...sharedOptions, file: source.file, fileName: source.fileName }
      : source?.type === "url"
        ? { ...sharedOptions, fileName: source.fileName, src: source.src }
        : sharedOptions,
  );

  React.useEffect(() => {
    if (!source || controller.isLoading) {
      initializedSourceRef.current = null;
      return;
    }
    if (!controller.activeTab || controller.error || initializedSourceRef.current === source) return;
    if (controller.src === PLAYGROUND_SAMPLE_URL) {
      const index = controller.sheets.findIndex(sheet => sheet.name === "Scatter Matrix");
      if (index >= 0 && controller.activeTab.sheetIndex !== index) {
        controller.setActiveSheetIndex(index);
        return;
      }
    }
    controller.setZoomScale(100);
    initializedSourceRef.current = source;
  }, [controller, source]);

  // Apply highlighted ranges after loading.
  React.useEffect(() => {
    const pending = pendingHighlightSampleRef.current;
    if (
      pending &&
      controller.src === HIGHLIGHT_DEMO_URL &&
      !controller.isLoading &&
      controller.sheets.some((sheet) => sheet.name === "Scatter Matrix") &&
      controller.highlightRanges(HIGHLIGHT_DEMO_RANGES, pending)
    ) {
      pendingHighlightSampleRef.current = null;
    }
  }, [controller.highlightRanges, controller.isLoading, controller.sheets, controller.src, source]);

  const loadWorkbookFile = React.useCallback(async (file: File) => {
    setSource({ file: await file.arrayBuffer(), fileName: file.name, type: "file" });
  }, []);
  const loadSample = React.useCallback(() => setSource(SAMPLE_SOURCE), []);
  const loadPivotSample = React.useCallback(() => setSource({ type: "url", src: PIVOT_SAMPLE_URL, fileName: "pivot-tables.xlsx" }), []);
  const loadUrl = React.useCallback((url: string) => {
    const trimmed = url.trim();
    if (trimmed) setSource({ src: trimmed, type: "url" });
  }, []);
  const loadHighlightSample = React.useCallback((autoScroll: boolean) => {
    pendingHighlightSampleRef.current = { autoScroll };
    setSource({ type: "url", src: HIGHLIGHT_DEMO_URL, fileName: "crazy-chart-zoo-subtypes.xlsx" });
  }, []);
  const openFilePicker = () => fileInputRef.current?.click();
  const isFileDrag = (event: React.DragEvent) => event.dataTransfer.types.includes("Files");

  return (
    <SettingsContext.Provider value={settingsValue}>
      <div
        className="relative flex h-dvh flex-col overflow-hidden bg-chrome text-foreground"
        onDragEnter={(event) => {
          if (!isFileDrag(event)) return;
          event.preventDefault();
          dragDepthRef.current += 1;
          setIsDragActive(true);
        }}
        onDragLeave={(event) => {
          if (!isFileDrag(event)) return;
          event.preventDefault();
          dragDepthRef.current = Math.max(0, dragDepthRef.current - 1);
          if (dragDepthRef.current === 0) setIsDragActive(false);
        }}
        onDragOver={(event) => {
          if (!isFileDrag(event)) return;
          event.preventDefault();
          event.dataTransfer.dropEffect = "copy";
        }}
        onDrop={(event) => {
          const file = event.dataTransfer.files[0];
          if (!file) return;
          event.preventDefault();
          dragDepthRef.current = 0;
          setIsDragActive(false);
          void loadWorkbookFile(file);
        }}
      >
        <input
          accept=".xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          className="hidden"
          onChange={(event) => {
            const file = event.target.files?.[0];
            event.target.value = "";
            if (file) void loadWorkbookFile(file);
          }}
          ref={fileInputRef}
          type="file"
        />

        <XlsxViewerProvider controller={controller} isDark={settings.documentDark}>
          <PivotPlaygroundProvider>
            <TitleBar onOpenFile={openFilePicker} />
            <Ribbon
              onClear={() => setSource(null)}
              onLoadPivotSample={loadPivotSample}
              onLoadHighlightSample={loadHighlightSample}
              onLoadSample={loadSample}
              onLoadUrl={loadUrl}
              onOpenFile={openFilePicker}
            />
            <FormulaBar />
            <PivotPlaygroundLayout>
              <XlsxViewer
                allowResizeInReadOnly={settings.allowResizeInReadOnly}
                className="h-full min-h-0 min-w-0 flex-1"
                contextMenuClassName="playground-context-menu"
                emptyState={<EmptyWorkbook onLoadSample={loadSample} onOpenFile={openFilePicker} />}
                experimentalCanvas={settings.experimentalCanvas}
                fileTooLargeState={<div className="grid h-full place-items-center text-sm font-medium">File is too large</div>}
                getCellStyle={getCellStyle}
                height="100%"
                highlightClassName={settings.customHighlightStyle ? "playground-range-highlight" : undefined}
                isDark={settings.documentDark}
                loadingState={
                  <div className="flex h-full w-full items-center justify-center text-muted-foreground">
                    <Spinner className="size-5" />
                  </div>
                }
                readOnly={settings.readOnly}
                renderCellEditorOverlay={(session) => <FormulaEditorDemo isDark={settings.documentDark} session={session} />}
                renderChartLoading={() => (
                  <div className="flex h-full w-full items-center justify-center text-muted-foreground">
                    <Spinner className="size-5" />
                  </div>
                )}
                contextMenuStyle={{ background: "transparent", border: 0, boxShadow: "none", minWidth: 0, padding: 0 }}
                renderContextMenu={menu => <FormulaContextMenu menu={menu} controller={controller} onPreview={setStylePreview} />}
                renderTableHeaderMenu={menu => <WorksheetHeaderMenu menu={menu} />}
                rounded={false}
                selectionColor="#107C41"
                showDefaultToolbar={false}
                showPivotTableControls={false}
              />
            </PivotPlaygroundLayout>
            <SheetTabs />
            <StatusBar />
          </PivotPlaygroundProvider>
        </XlsxViewerProvider>

        {isDragActive ? (
          <div className="pointer-events-none absolute inset-0 z-30 flex items-center justify-center bg-background/70 backdrop-blur-[2px]">
            <div className="flex flex-col items-center gap-3 rounded-xl border-2 border-dashed border-brand/60 bg-background px-10 py-8 text-center shadow-lg">
              <Upload className="size-6 text-brand" />
              <div className="text-sm font-medium">Drop to open workbook</div>
            </div>
          </div>
        ) : null}
      </div>
    </SettingsContext.Provider>
  );
}
