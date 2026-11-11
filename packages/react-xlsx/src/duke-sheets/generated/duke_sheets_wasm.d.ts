/* tslint:disable */
/* eslint-disable */

declare global {
    interface SymbolConstructor {
        readonly dispose: unique symbol;
    }
}

export interface JsRowCell {
    col: number;
    value: string;
    richText?: Array<{ text: string; font?: FontStyleInput | null }>;
    style?: any;
    mergeSpan?: JsMergeSpan;
    isMergedSecondary?: boolean;
    hyperlink?: any;
    comment?: any;
    formula?: string;
    image?: any;
}

export interface JsRow {
    index: number;
    cells: Array<JsRowCell>;
}

export interface JsMergeSpan {
    rowSpan: number;
    colSpan: number;
}

export interface JsRowsOptions {
    useFormattedValues?: boolean;
    useCalculatedValues?: boolean;
    includeStyles?: boolean;
    includeMergeInfo?: boolean;
    includeHyperlinks?: boolean;
    includeComments?: boolean;
    includeFormulas?: boolean;
    includeImages?: boolean;
    skipEmptyValues?: boolean;
    skipBlankValues?: boolean;
}

export interface ColorInput {
    colorType?: "auto" | "rgb" | "argb" | "theme" | "indexed";
    hex?: string;
    r?: number;
    g?: number;
    b?: number;
    a?: number;
    themeIndex?: number;
    tint?: number;
    paletteIndex?: number;
}

export interface FontStyleInput {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: "none" | "single" | "double" | "singleAccounting" | "doubleAccounting";
    strikethrough?: boolean;
    color?: ColorInput;
    verticalAlign?: "baseline" | "superscript" | "subscript";
    family?: number;
    charset?: number;
    scheme?: string;
}

export interface GradientStopInput {
    position: number;
    color: ColorInput;
}

export interface FillStyleInput {
    fillType?: "none" | "solid" | "pattern" | "gradient";
    color?: ColorInput;
    pattern?: string;
    foreground?: ColorInput;
    background?: ColorInput;
    gradientType?: "linear" | "path";
    angle?: number;
    stops?: GradientStopInput[];
}

export interface BorderEdgeInput {
    style?: "none" | "thin" | "medium" | "thick" | "dashed" | "dotted" | "double" | "hair" | "mediumDashed" | "dashDot" | "mediumDashDot" | "dashDotDot" | "mediumDashDotDot" | "slantDashDot";
    color?: ColorInput;
}

export interface BorderStyleInput {
    left?: BorderEdgeInput;
    right?: BorderEdgeInput;
    top?: BorderEdgeInput;
    bottom?: BorderEdgeInput;
    diagonal?: BorderEdgeInput;
    diagonalDirection?: "none" | "down" | "up" | "both";
}

export interface AlignmentInput {
    horizontal?: "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
    vertical?: "top" | "center" | "bottom" | "justify" | "distributed";
    wrapText?: boolean;
    shrinkToFit?: boolean;
    indent?: number;
    rotation?: number;
    readingOrder?: "contextDependent" | "leftToRight" | "rightToLeft";
}

export interface NumberFormatInput {
    formatType?: "general" | "builtin" | "custom";
    id?: number;
    formatString?: string;
}

export interface CellProtectionInput {
    locked?: boolean;
    hidden?: boolean;
}

export interface StyleInput {
    font?: FontStyleInput;
    fill?: FillStyleInput;
    border?: BorderStyleInput;
    alignment?: AlignmentInput;
    numberFormat?: NumberFormatInput;
    protection?: CellProtectionInput;
}

export class RowIterator implements IterableIterator<JsRow> {
    constructor(ws: Worksheet, opts?: JsRowsOptions, maxRow?: number);
    [Symbol.iterator](): RowIterator;
    next(): IteratorResult<JsRow>;
}

export interface DrawingMeta {
    name?: string;
    hidden: boolean;
    locked: boolean;
    printable: boolean;
    altText?: string;
    title?: string;
}

export interface DrawingMetaInput {
    name?: string;
    hidden?: boolean;
    locked?: boolean;
    printable?: boolean;
    altText?: string;
    title?: string;
}

export interface DrawingCellMarker {
    col: number;
    row: number;
    colOffsetEmu?: number;
    rowOffsetEmu?: number;
}

export type DrawingAnchor =
| { type: "twoCell"; from: DrawingCellMarker; to: DrawingCellMarker; editAs?: "twoCell" | "oneCell" | "absolute" }
| { type: "oneCell"; from: DrawingCellMarker; widthEmu: number; heightEmu: number }
| { type: "absolute"; xEmu: number; yEmu: number; widthEmu: number; heightEmu: number };

export interface DrawingChildTransform {
    xEmu?: number;
    yEmu?: number;
    cxEmu?: number;
    cyEmu?: number;
    rotation?: number;
    flipH?: boolean;
    flipV?: boolean;
}

export interface DrawingGroupTransform extends DrawingChildTransform {
    childXEmu?: number;
    childYEmu?: number;
    childCxEmu?: number;
    childCyEmu?: number;
}

export type DrawingPlacement =
| { anchor: DrawingAnchor; transform?: never }
| { anchor?: never; transform: DrawingChildTransform };

export type DrawingColor =
| { colorType: "auto" }
| { colorType: "rgb"; r: number; g: number; b: number }
| { colorType: "argb"; a: number; r: number; g: number; b: number }
| { colorType: "theme"; index: number; tint: number }
| { colorType: "indexed"; index: number };

export interface DrawingRunFont {
    bold?: boolean;
    italic?: boolean;
    size?: number;
    color?: DrawingColor;
    name?: string;
    underline?: "none" | "single" | "double" | "singleAccounting" | "doubleAccounting";
    strikethrough?: boolean;
    verticalAlign?: "baseline" | "superscript" | "subscript";
    family?: number;
    charset?: number;
    scheme?: string;
}

export interface DrawingText {
    runs: Array<{ text: string; font?: DrawingRunFont }>;
    horizontalAlignment?: "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
    verticalAlignment?: "top" | "center" | "bottom" | "justify" | "distributed";
}

export type FormControlKind =
| { kind: "button"; caption: DrawingText }
| { kind: "checkbox"; caption: DrawingText; state: "unchecked" | "checked" | "mixed"; cellLink?: string; no3D: boolean }
/**
 * `firstInGroup` reports whether this radio heads its group; writers
 * recompute it from group-box containment, so it is read-side
 * information. `"mixed"` never validates on write but can surface
 * when reading hostile files.
 */
| { kind: "optionButton"; caption: DrawingText; state: "unchecked" | "checked" | "mixed"; cellLink?: string; firstInGroup: boolean; no3D: boolean }
| { kind: "label"; caption: DrawingText }
| { kind: "groupBox"; caption: DrawingText; no3D: boolean }
| { kind: "listBox"; inputRange?: string; cellLink?: string; selection: "single" | "multi" | "extend"; selected: number[]; no3D: boolean }
| { kind: "dropdown"; inputRange?: string; cellLink?: string; selected?: number; lines: number; no3D: boolean }
| { kind: "scrollbar"; value: number; min: number; max: number; increment: number; page: number; horizontal: boolean; cellLink?: string }
| { kind: "spinner"; value: number; min: number; max: number; increment: number; cellLink?: string }
/** Unsupported legacy control, preserved for passthrough. */
| { kind: "unknown"; objectType: string; legacyObjectType?: number; caption: DrawingText };

export type FormControlKindInput =
| { kind: "button"; caption: DrawingText }
| { kind: "checkbox"; caption: DrawingText; state: "unchecked" | "checked" | "mixed"; cellLink?: string; no3D?: boolean }
/** `firstInGroup` is ignored on input; writers recompute it from group-box containment. */
| { kind: "optionButton"; caption: DrawingText; state: "unchecked" | "checked"; cellLink?: string; firstInGroup?: boolean; no3D?: boolean }
| { kind: "label"; caption: DrawingText }
| { kind: "groupBox"; caption: DrawingText; no3D?: boolean }
| { kind: "listBox"; inputRange?: string; cellLink?: string; selection: "single" | "multi" | "extend"; selected?: number[]; no3D?: boolean }
| { kind: "dropdown"; inputRange?: string; cellLink?: string; selected?: number; lines: number; no3D?: boolean }
| { kind: "scrollbar"; value: number; min: number; max: number; increment: number; page: number; horizontal?: boolean; cellLink?: string }
| { kind: "spinner"; value: number; min: number; max: number; increment: number; cellLink?: string }
| { kind: "unknown"; objectType: string; legacyObjectType?: number; caption?: DrawingText };

export interface FormControlPayload {
    kind: FormControlKind;
    macroName?: string;
    /**
     * Unmodeled VML ClientData child fragments preserved on any control
     * kind; opaque internal passthrough echoed back unchanged on write.
     */
    rawClientData: number[][];
    /**
     * Unmodeled XLSX formControlPr attributes preserved on any control
     * kind; opaque internal passthrough echoed back unchanged on write.
     */
    rawProperties: Array<[string, string]>;
    /** Original BIFF OBJ body for XLS passthrough of unknown controls. */
    rawObj?: number[];
}

export interface FormControlInputPayload {
    kind: FormControlKindInput;
    macroName?: string;
    /** Echo back the raw* fields unchanged when rewriting a control read from a file. */
    rawClientData?: Array<Uint8Array | number[]>;
    rawProperties?: Array<[string, string]>;
    rawObj?: Uint8Array | number[];
}

export interface DrawingImage {
    format: "png" | "jpeg" | "gif" | "bmp" | "emf" | "wmf" | "tiff" | "svg";
    mediaPath: string;
    svgMediaPath?: string;
    widthEmu: number;
    heightEmu: number;
    rotation?: number;
    flipH: boolean;
    flipV: boolean;
}

export interface DrawingImageInput extends Partial<Omit<DrawingImage, "format" | "widthEmu" | "heightEmu">> {
    format: DrawingImage["format"];
    widthEmu: number;
    heightEmu: number;
    data: Uint8Array | number[];
    svgData?: Uint8Array | number[];
}

export type DrawingShapeFill =
| { kind: "none" }
| { kind: "solid"; color: DrawingColor };

export interface DrawingShape {
    geometry: string;
    fill: DrawingShapeFill;
    line: { color?: DrawingColor; widthEmu?: number; dashStyle?: string; noFill: boolean };
    text?: DrawingText;
    rotation: number;
    flipH: boolean;
    flipV: boolean;
}

export interface DrawingShapeInput extends Partial<DrawingShape> {
    geometry?: string;
}

export interface DrawingComment {
    row: number;
    col: number;
    author: string;
    /** Plain text (runs concatenated). */
    text: string;
    /**
     * Rich runs; present on output when any run is formatted, and wins
     * over `text` on input when supplied.
     */
    richText?: DrawingText;
}

export type ChartDataReference =
| { refType: "formula"; formula: string }
| { refType: "numbers"; numbers: number[] }
| { refType: "strings"; strings: string[] };

export interface ChartShapeProperties {
    solidFillHex?: string;
    noFill: boolean;
    lineWidth?: number;
    lineColorHex?: string;
    lineNoFill: boolean;
    lineDashStyle?: string;
}

export interface ChartNumberFormat {
    formatCode: string;
    sourceLinked?: boolean;
}

export interface ChartDataLabels {
    showLegendKey?: boolean;
    showValue?: boolean;
    showCategoryName?: boolean;
    showSeriesName?: boolean;
    showPercent?: boolean;
    showBubbleSize?: boolean;
    separator?: string;
    position?: string;
    numberFormat?: ChartNumberFormat;
    showLeaderLines?: boolean;
}

export interface ChartDataSeries {
    name?: string;
    values: ChartDataReference;
    categories?: ChartDataReference;
    dataLabels?: ChartDataLabels;
    trendline?: { trendlineType: string; name?: string; order?: number; period?: number; forward?: number; backward?: number; intercept?: number; displayRSquared?: boolean; displayEquation?: boolean };
    errorBars?: { direction: string; barType: string; valueType: string; value?: number; noEndCap?: boolean };
    marker?: { symbol?: string; size?: number };
    dataPoints: Array<{ index: number; marker?: { symbol?: string; size?: number }; explosion?: number; shapeProperties?: ChartShapeProperties }>;
    smooth?: boolean;
    explosion?: number;
    invertIfNegative?: boolean;
    shapeProperties?: ChartShapeProperties;
}

export interface ChartAxis {
    title?: string;
    minimum?: number;
    maximum?: number;
    majorUnit?: number;
    minorUnit?: number;
    position: "bottom" | "top" | "left" | "right";
    numberFormat?: ChartNumberFormat;
    majorGridlines: boolean;
    minorGridlines: boolean;
    majorGridlinesShapeProperties?: ChartShapeProperties;
    minorGridlinesShapeProperties?: ChartShapeProperties;
    majorTickMark?: string;
    minorTickMark?: string;
    labelPosition?: string;
    delete?: boolean;
    crosses?: string;
    crossBetween?: string;
    shapeProperties?: ChartShapeProperties;
}

export interface Chart {
    chartType: string;
    title?: string;
    series: ChartDataSeries[];
    categoryAxis?: ChartAxis;
    valueAxis?: ChartAxis;
    legend?: { position: string; overlay: boolean };
    dataLabels?: ChartDataLabels;
    view3D?: { rotateX?: number; rotateY?: number; depthPercent?: number; heightPercent?: number; perspective?: number; rightAngleAxes?: boolean };
    dataTable?: { showHorizontalBorder?: boolean; showVerticalBorder?: boolean; showOutline?: boolean; showKeys?: boolean };
    displayBlanksAs?: "gap" | "span" | "zero";
    plotVisibleOnly?: boolean;
    layout?: { manualLayout?: { x?: number; y?: number; width?: number; height?: number } };
    shapeProperties?: ChartShapeProperties;
    is3D: boolean;
    varyColors?: boolean;
    gapWidth?: number;
    overlap?: number;
    firstSliceAngle?: number;
    holeSize?: number;
    bubbleScale?: number;
    showNegativeBubbles?: boolean;
    autoTitleDeleted?: boolean;
    roundedCorners?: boolean;
    showDlblsOverMax?: boolean;
    wireframe?: boolean;
    radarStyle?: string;
    typeGroups: Array<{ chartType: string; is3D: boolean; series: ChartDataSeries[]; dataLabels?: ChartDataLabels; varyColors?: boolean; gapWidth?: number; overlap?: number; firstSliceAngle?: number; holeSize?: number; bubbleScale?: number; showNegativeBubbles?: boolean; radarStyle?: string; wireframe?: boolean; axisIds: number[] }>;
    axes: Array<{ id: number; crossId: number; axis: ChartAxis }>;
}

export interface ChartSeriesInput {
    name?: string;
    values: ChartDataReference;
    categories?: ChartDataReference;
}

/** Exactly the chart fields accepted when authoring; all other Chart fields are read-only. */
export interface ChartInput {
    chartType: string;
    title?: string;
    series?: ChartSeriesInput[];
    is3D?: boolean;
    varyColors?: boolean;
    gapWidth?: number;
    overlap?: number;
}

export interface ChartExTitle {
    text?: string;
    position?: string;
    align?: string;
    overlay?: boolean;
    offset?: { top?: number; left?: number };
    shapeProperties?: ChartShapeProperties;
}

export interface ChartStyleReference {
    idx: number
    /** The colour override, as the XML it was read as. */
    color?: string
}

export interface ChartStyleEntry {
    lineReference: ChartStyleReference
    lineWidthScale?: number
    fillReference: ChartStyleReference
    effectReference: ChartStyleReference
    /** `major`, `minor` or `none`. */
    fontCollection: string
    fontColor?: string
    /** DrawingML kept as the XML it was read as. */
    shapeProperties?: string
    defaultRunProperties?: string
    bodyProperties?: string
    mods?: string
}

/**
 * A chart style part. `entries` is keyed by the element name the entry
 * belongs to (`chartArea`, `dataPoint`, ...); `raw` is set instead when
 * the part could not be modelled and is replayed as read.
 */
export interface ChartStyle {
    id?: number
    entries: Record<string, ChartStyleEntry>
    markerSymbol?: string
    markerSize?: number
    raw?: string
}

/** A chart colour style part. `raw` is set when it could not be modelled. */
export interface ChartColorStyle {
    method?: string
    id?: number
    colors: string[]
    variations: string[]
    raw?: string
}

export interface ChartExGridlines {
    shapeProperties?: ChartShapeProperties;
}

export interface ChartExDataLabelOverride {
    idx: number;
    position?: string;
    visibilitySeriesName?: boolean;
    visibilityCategoryName?: boolean;
    visibilityValue?: boolean;
    numberFormat?: ChartNumberFormat;
    separator?: string;
    shapeProperties?: ChartShapeProperties;
}

export interface ChartExColorPosition {
    positionType: 'extremeValue' | 'number' | 'percent';
    value?: number;
}

export interface ChartExSeriesLayoutProperties {
    parentLabelLayout?: string;
    regionLabelLayout?: string;
    visibility?: { connectorLines?: boolean; meanLine?: boolean; meanMarker?: boolean; nonoutliers?: boolean; outliers?: boolean };
    aggregation: boolean;
    binning?: { intervalClosed?: string; underflow?: string; overflow?: string; binSize?: number; binCount?: number };
    geography?: { projectionType?: string; viewedRegionType?: string; cultureLanguage?: string; cultureRegion?: string; attribution?: string };
    statistics?: { quartileMethod?: string };
    /** Absent when the cx:subtotals element is absent; [] when it is present but empty. */
    subtotals?: number[];
}

export interface ChartExSeries {
    layout: string;
    dataId: number;
    uniqueId?: string;
    hidden?: boolean;
    ownerIdx?: number;
    formatIdx?: number;
    text?: { formula?: string; value?: string };
    dataLabels?: { position?: string; visibilitySeriesName?: boolean; visibilityCategoryName?: boolean; visibilityValue?: boolean; numberFormat?: ChartNumberFormat; separator?: string; shapeProperties?: ChartShapeProperties; overrides: ChartExDataLabelOverride[]; hiddenLabels: number[] };
    dataPoints: Array<{ idx: number; shapeProperties?: ChartShapeProperties }>;
    layoutProperties?: ChartExSeriesLayoutProperties;
    axisIds: number[];
    valueColors: boolean;
    valueColorPositions?: { count?: number; min?: ChartExColorPosition; mid?: ChartExColorPosition; max?: ChartExColorPosition };
    shapeProperties?: ChartShapeProperties;
}

export interface ChartExAxis {
    id: number;
    hidden?: boolean;
    scaling: { scalingType: string; gapWidth?: number; min?: number; max?: number; majorUnit?: number; minorUnit?: number };
    title?: { text?: string; shapeProperties?: ChartShapeProperties };
    units?: { unit?: string };
    majorGridlines?: ChartExGridlines;
    minorGridlines?: ChartExGridlines;
    majorTickMarks?: string;
    minorTickMarks?: string;
    tickLabels: boolean;
    numberFormat?: ChartNumberFormat;
    shapeProperties?: ChartShapeProperties;
}

export interface ChartEx {
    layout: string;
    version?: string;
    featureList?: string;
    fallbackImg?: string;
    title?: ChartExTitle;
    data: Array<{ id: number; dimensions: Array<{ dimType: string; formula?: string; nfFormula?: string }> }>;
    plotArea: { plotSurface?: ChartShapeProperties; series: ChartExSeries[]; axes: ChartExAxis[]; shapeProperties?: ChartShapeProperties };
    legend?: { position?: string; align?: string; overlay?: boolean; offset?: { top?: number; left?: number }; shapeProperties?: ChartShapeProperties };
    shapeProperties?: ChartShapeProperties;
    formatOverrides: Array<{ idx: number; shapeProperties?: ChartShapeProperties }>;
    externalDataRelId?: string;
    externalDataAutoUpdate?: boolean;
    style?: ChartStyle;
    colorStyle?: ChartColorStyle;
}

export interface ChartExTitleInput {
    text?: string;
    position?: string;
    align?: string;
    overlay?: boolean;
}

/** Exactly the ChartEx fields accepted when authoring; all other ChartEx fields are read-only. */
export interface ChartExInput {
    layout: string;
    version?: string;
    featureList?: string;
    fallbackImg?: string;
    title?: ChartExTitleInput;
}

export interface RawDrawingMetadata {
    byteLength: number;
    relationships: Array<{ id: string; relType: string; target: string; external: boolean; hasPart: boolean }>;
}

/**
 * Resolved on-sheet placement in EMU: the anchor rectangle for
 * top-level drawings, the group-mapped (rotation/flip aware)
 * rectangle for group children.
 */
export type RectEmu = { xEmu: number; yEmu: number; widthEmu: number; heightEmu: number };

type DrawingNode = DrawingMeta & DrawingPlacement & { drawingPath: number[]; absoluteRectEmu: RectEmu };

export type ImageDrawing = DrawingNode & { kind: "image"; image: DrawingImage };
export type ChartDrawing = DrawingNode & { kind: "chart"; chart: Chart };
export type ChartExDrawing = DrawingNode & { kind: "chartEx"; chartEx: ChartEx };
export type FormControlDrawing = DrawingNode & { kind: "formControl"; formControl: FormControlPayload };
export type CommentDrawing = DrawingNode & { kind: "comment"; comment: DrawingComment };
export type ShapeDrawing = DrawingNode & { kind: "shape"; shape: DrawingShape };
export type GroupDrawing = DrawingNode & { kind: "group"; group: { groupTransform: DrawingGroupTransform; children: Drawing[] } };
export type RawDrawing = DrawingNode & { kind: "raw"; raw: RawDrawingMetadata };

export type Drawing = ImageDrawing | ChartDrawing | ChartExDrawing | FormControlDrawing | CommentDrawing | ShapeDrawing | GroupDrawing | RawDrawing;

type DrawingInputPlacement =
| { anchor: DrawingAnchor; transform?: never }
| { anchor?: never; transform: DrawingChildTransform };

export type DrawingInput = DrawingMetaInput & DrawingInputPlacement & (
| { kind: "image"; image: DrawingImageInput }
| { kind: "chart"; chart: ChartInput }
| { kind: "chartEx"; chartEx: ChartExInput }
| { kind: "formControl"; formControl: FormControlInputPayload }
| { kind: "comment"; comment: DrawingComment }
| { kind: "shape"; shape: DrawingShapeInput }
| { kind: "group"; group: { groupTransform?: DrawingGroupTransform; children?: DrawingInput[] } }
);

export interface FormControlInteractionResult {
    controlsChanged: number;
    linkedCellsChanged: number;
}

export interface Worksheet {
    iterateRows(opts?: JsRowsOptions): RowIterator;
    setCellStyle(address: string, style: StyleInput): void;
    setCellStyleAt(row: number, col: number, style: StyleInput): void;
    setRangeStyle(range: string, style: StyleInput): void;
    readonly drawings: Drawing[];
    readonly formControls: FormControlDrawing[];
    readonly formControlCount: number;
    readonly images: ImageDrawing[];
    readonly imageCount: number;
    readonly charts: ChartDrawing[];
    readonly chartCount: number;
    readonly chartsEx: ChartExDrawing[];
    readonly chartExCount: number;
    addDrawing(drawing: DrawingInput & { anchor: DrawingAnchor }): number;
    /** Drawing paths are positional; mutating the list invalidates previously returned paths. */
    insertDrawing(index: number, drawing: DrawingInput & { anchor: DrawingAnchor }): void;
    /** Drawing paths are positional; mutating the list invalidates previously returned paths. */
    setDrawing(path: number[], drawing: DrawingInput): void;
    /** Drawing paths are positional; mutating the list invalidates previously returned paths. */
    removeDrawing(path: number[]): void;
    /** Drawing paths are positional; mutating the list invalidates previously returned paths. */
    moveDrawing(from: number, to: number): void;
    /** Paths are positional; mutating the drawing list invalidates previously returned paths. */
    drawingImageData(path: number[]): Uint8Array;
    /** Paths are positional; mutating the drawing list invalidates previously returned paths. */
    drawingSvgData(path: number[]): Uint8Array | undefined;
    setFormControlCheckState(path: number[], state: "unchecked" | "checked" | "mixed"): FormControlInteractionResult;
}

export interface Workbook {
    /**
     * Resolve a drawing color to display RGB ("RRGGBB" hex) against
     * this workbook's theme palette; `auto` resolves to undefined.
     */
    resolveColor(color: DrawingColor): string | undefined;
}



export class CellValue {
    private constructor();
    free(): void;
    [Symbol.dispose](): void;
    asBoolean(): boolean | undefined;
    asError(): string | undefined;
    asNumber(): number | undefined;
    asText(): string | undefined;
    toJs(): any;
    toString(): string;
    readonly is_boolean: boolean;
    readonly is_empty: boolean;
    readonly is_error: boolean;
    readonly is_number: boolean;
    readonly is_text: boolean;
}

export class Workbook {
    free(): void;
    [Symbol.dispose](): void;
    addSheet(name: string): number;
    calculate(options?: any | null): any;
    defineName(name: string, refers_to: string): void;
    exportSnapshot(): Uint8Array;
    fontFamilies(): string[];
    /**
     * Load a workbook from bytes, auto-detecting the format (XLSX or XLS).
     */
    static fromBytes(data: Uint8Array): Workbook;
    static fromBytesForViewer(data: Uint8Array): Workbook;
    /**
     * Load a password-protected workbook from bytes, auto-detecting
     * the format (XLSX or XLS). `skipIntegrityCheck` skips the HMAC
     * integrity check on Agile-encrypted files (matches Office
     * behaviour); defaults to false.
     */
    static fromBytesWithPassword(data: Uint8Array, password: string, skip_integrity_check?: boolean | null): Workbook;
    static fromCsvString(csv: string): Workbook;
    static fromSnapshot(data: Uint8Array): Workbook;
    getNamedRange(name: string): string | undefined;
    getSheet(index: number): Worksheet;
    getSheetByName(name: string): Worksheet;
    static loadCsvString(csv: string): Workbook;
    constructor();
    removeSheet(index: number): void;
    /**
     * Save the first sheet as a CSV string, with form-control state
     * synchronized into linked cells in the output.
     */
    saveCsvString(): string;
    /**
     * Save the workbook as encrypted XLS bytes after synchronizing
     * form-control state into linked cells. `profile` selects
     * the FilePass variant; `null` defaults to RC4 CryptoAPI 128.
     * Valid values: `"rc4-cryptoapi"`, `"rc4-legacy"`, `"xor"`.
     * `keyBits` controls RC4 CryptoAPI key size (40 or 128). XOR is
     * not certified to interoperate with modern Excel.
     */
    saveXlsBytesEncrypted(password: string, profile?: string | null, key_bits?: number | null): Uint8Array;
    /**
     * Save XLSB bytes with form-control state synchronized into linked cells,
     * replacing existing values and formulas in the output.
     */
    saveXlsbBytes(): Uint8Array;
    /**
     * Save XLSX bytes with form-control state synchronized into linked cells,
     * replacing existing values and formulas in the output.
     */
    saveXlsxBytes(): Uint8Array;
    /**
     * Save the workbook as encrypted XLSX bytes after synchronizing
     * form-control state into linked cells. `profile` selects
     * the encryption variant; passing `null`/`undefined` uses the
     * Agile-256 default. Valid values: `"agile"`, `"standard"`.
     * `keyBits` and `spinCount` override the defaults where the
     * profile supports them.
     */
    saveXlsxBytesEncrypted(password: string, profile?: string | null, key_bits?: number | null, spin_count?: number | null): Uint8Array;
    setWorkbookProtection(protection: any): void;
    sheetIndex(name: string): number | undefined;
    /**
     * Project all form-control state into linked cells.
     */
    syncFormControls(): number;
    /**
     * Drive controls from formula-backed linked cells.
     */
    syncFormControlsFromLinkedCells(): number;
    takeViewerArchive(): any;
    readonly activeSheet: number;
    readonly chartsheetCount: number;
    readonly chartsheets: any;
    readonly isEmpty: boolean;
    readonly namedRanges: any;
    readonly settings: any;
    readonly sheetCount: number;
    readonly sheetNames: string[];
    readonly sheetOrder: any;
    /**
     * The workbook theme's 12 clrScheme colors as `RRGGBB` hex, in
     * theme-index order (background 1, text 1, background 2, text 2,
     * accent 1-6, hyperlink, followed hyperlink). The Office default
     * palette when the file carries no theme.
     */
    readonly themePalette: string[];
    readonly totalSheetCount: number;
    readonly workbookProtection: any;
}

export class Worksheet {
    private constructor();
    free(): void;
    [Symbol.dispose](): void;
    getCalculatedValue(address: string): CellValue;
    getCalculatedValueAt(row: number, col: number): CellValue;
    getCell(address: string): CellValue;
    getCellAt(row: number, col: number): CellValue;
    getCellStyle(address: string): any;
    getCellStyleAt(row: number, col: number): any;
    getColumnOutlineLevel(col: number): number;
    getColumnWidth(col: number): number | undefined;
    getComment(address: string): any;
    getCommentAt(row: number, col: number): any;
    getFormattedValue(address: string): string;
    getFormattedValueAt(row: number, col: number): string;
    getFormulaAt(row: number, col: number): string | undefined;
    getHyperlink(address: string): any;
    getHyperlinkAt(row: number, col: number): any;
    getImageAt(row: number, col: number): any;
    getMergeSpan(row: number, col: number): any;
    getRichTextAt(row: number, col: number): any;
    getRowHeight(row: number): number | undefined;
    getRowOutlineLevel(row: number): number;
    getSpillSource(row: number, col: number): any;
    getTableByName(name: string): any;
    getValueSnapshotAt(row: number, col: number): Uint8Array;
    hasComment(address: string): boolean;
    hasCommentAt(row: number, col: number): boolean;
    isColumnCollapsed(col: number): boolean;
    isColumnHidden(col: number): boolean;
    isMergedSecondary(row: number, col: number): boolean;
    isRowCollapsed(row: number): boolean;
    isRowHidden(row: number): boolean;
    isSpillSource(row: number, col: number): boolean;
    isSpillTarget(row: number, col: number): boolean;
    mergeCells(range_str: string): void;
    restoreValueSnapshotAt(row: number, col: number, bytes: Uint8Array): void;
    setCell(address: string, value: any): void;
    setCellStyle(address: string, style: any): void;
    setCellStyleAt(row: number, col: number, style: any): void;
    setColumnWidth(col: number, width: number): void;
    setFormula(address: string, formula: string): void;
    setProtectedRanges(ranges: any): void;
    setProtection(protection: any): void;
    setRangeStyle(range_str: string, style: any): void;
    setRowHeight(row: number, height: number): void;
    unmergeCells(range_str: string): boolean;
    usedRange(): any;
    readonly autoFilter: any;
    readonly cellCount: number;
    readonly colBreaks: any;
    readonly commentAuthors: string[];
    readonly commentCount: number;
    readonly comments: any;
    readonly conditionalFormatCount: number;
    readonly conditionalFormats: any;
    readonly dataValidationCount: number;
    readonly dataValidations: any;
    readonly date1904: boolean;
    readonly defaultColumnWidth: number;
    readonly defaultRowHeight: number;
    readonly formulaCells: any;
    /**
     * Get the number of formula cells in this worksheet.
     */
    readonly formulaCount: number;
    readonly freezePanes: any;
    readonly hyperlinkCount: number;
    readonly hyperlinks: any;
    readonly isEmpty: boolean;
    readonly isSelected: boolean;
    readonly mergedRegions: any;
    readonly name: string;
    readonly pageSetup: any;
    readonly printArea: string | undefined;
    readonly protectedRanges: any;
    readonly protection: any;
    readonly repeatCols: Uint32Array | undefined;
    readonly repeatRows: Uint32Array | undefined;
    readonly rowBreaks: any;
    readonly selections: any;
    readonly splitPanes: any;
    readonly tabColor: any;
    readonly tableCount: number;
    readonly tables: any;
    readonly visibility: string;
    readonly zoomScale: number | undefined;
}

export function decodeArchive(data: Uint8Array): any;

export function init(): void;

export function inspectArchive(data: Uint8Array): any;

export function parseXmlTree(xml: string): string;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly workbook_syncFormControls: (a: number) => number;
    readonly workbook_syncFormControlsFromLinkedCells: (a: number) => number;
    readonly worksheet_addDrawing: (a: number, b: any) => [number, number, number];
    readonly worksheet_chartCount: (a: number) => [number, number, number];
    readonly worksheet_chartExCount: (a: number) => [number, number, number];
    readonly worksheet_charts: (a: number) => [number, number, number];
    readonly worksheet_chartsEx: (a: number) => [number, number, number];
    readonly worksheet_drawingImageData: (a: number, b: any) => [number, number, number, number];
    readonly worksheet_drawingSvgData: (a: number, b: any) => [number, number, number, number];
    readonly worksheet_drawings: (a: number) => [number, number, number];
    readonly worksheet_formControlCount: (a: number) => [number, number, number];
    readonly worksheet_formControls: (a: number) => [number, number, number];
    readonly worksheet_imageCount: (a: number) => [number, number, number];
    readonly worksheet_images: (a: number) => [number, number, number];
    readonly worksheet_insertDrawing: (a: number, b: number, c: any) => [number, number];
    readonly worksheet_moveDrawing: (a: number, b: number, c: number) => [number, number];
    readonly worksheet_removeDrawing: (a: number, b: any) => [number, number];
    readonly worksheet_setDrawing: (a: number, b: any, c: any) => [number, number];
    readonly worksheet_setFormControlCheckState: (a: number, b: any, c: number, d: number) => [number, number, number];
    readonly decodeArchive: (a: number, b: number) => [number, number, number];
    readonly inspectArchive: (a: number, b: number) => [number, number, number];
    readonly parseXmlTree: (a: number, b: number) => [number, number, number, number];
    readonly workbook_activeSheet: (a: number) => [number, number, number];
    readonly workbook_chartsheetCount: (a: number) => [number, number, number];
    readonly workbook_chartsheets: (a: number) => [number, number, number];
    readonly workbook_fontFamilies: (a: number) => [number, number];
    readonly workbook_fromBytesForViewer: (a: number, b: number) => [number, number, number];
    readonly workbook_isEmpty: (a: number) => [number, number, number];
    readonly workbook_namedRanges: (a: number) => [number, number, number];
    readonly workbook_resolveColor: (a: number, b: any) => [number, number, number, number];
    readonly workbook_settings: (a: number) => [number, number, number];
    readonly workbook_sheetIndex: (a: number, b: number, c: number) => [number, number, number];
    readonly workbook_sheetOrder: (a: number) => [number, number, number];
    readonly workbook_takeViewerArchive: (a: number) => [number, number, number];
    readonly workbook_themePalette: (a: number) => [number, number];
    readonly workbook_totalSheetCount: (a: number) => [number, number, number];
    readonly workbook_workbookProtection: (a: number) => [number, number, number];
    readonly worksheet_autoFilter: (a: number) => [number, number, number];
    readonly worksheet_cellCount: (a: number) => [number, number, number];
    readonly worksheet_colBreaks: (a: number) => [number, number, number];
    readonly worksheet_commentAuthors: (a: number) => [number, number, number, number];
    readonly worksheet_commentCount: (a: number) => [number, number, number];
    readonly worksheet_comments: (a: number) => [number, number, number];
    readonly worksheet_conditionalFormatCount: (a: number) => [number, number, number];
    readonly worksheet_conditionalFormats: (a: number) => [number, number, number];
    readonly worksheet_dataValidationCount: (a: number) => [number, number, number];
    readonly worksheet_dataValidations: (a: number) => [number, number, number];
    readonly worksheet_date1904: (a: number) => [number, number, number];
    readonly worksheet_defaultColumnWidth: (a: number) => [number, number, number];
    readonly worksheet_defaultRowHeight: (a: number) => [number, number, number];
    readonly worksheet_formulaCells: (a: number) => [number, number, number];
    readonly worksheet_formulaCount: (a: number) => [number, number, number];
    readonly worksheet_freezePanes: (a: number) => [number, number, number];
    readonly worksheet_getCellStyle: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getCellStyleAt: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getColumnOutlineLevel: (a: number, b: number) => [number, number, number];
    readonly worksheet_getComment: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getCommentAt: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getFormattedValue: (a: number, b: number, c: number) => [number, number, number, number];
    readonly worksheet_getFormattedValueAt: (a: number, b: number, c: number) => [number, number, number, number];
    readonly worksheet_getFormulaAt: (a: number, b: number, c: number) => [number, number, number, number];
    readonly worksheet_getHyperlink: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getHyperlinkAt: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getMergeSpan: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getRichTextAt: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getRowOutlineLevel: (a: number, b: number) => [number, number, number];
    readonly worksheet_getRowsBatch: (a: number, b: number, c: number, d: any) => [number, number, number];
    readonly worksheet_getSpillSource: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getTableByName: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getValueSnapshotAt: (a: number, b: number, c: number) => [number, number, number, number];
    readonly worksheet_hasComment: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_hasCommentAt: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_hyperlinkCount: (a: number) => [number, number, number];
    readonly worksheet_hyperlinks: (a: number) => [number, number, number];
    readonly worksheet_isColumnCollapsed: (a: number, b: number) => [number, number, number];
    readonly worksheet_isColumnHidden: (a: number, b: number) => [number, number, number];
    readonly worksheet_isEmpty: (a: number) => [number, number, number];
    readonly worksheet_isMergedSecondary: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_isRowCollapsed: (a: number, b: number) => [number, number, number];
    readonly worksheet_isRowHidden: (a: number, b: number) => [number, number, number];
    readonly worksheet_isSelected: (a: number) => [number, number, number];
    readonly worksheet_isSpillSource: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_isSpillTarget: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_mergedRegions: (a: number) => [number, number, number];
    readonly worksheet_pageSetup: (a: number) => [number, number, number];
    readonly worksheet_printArea: (a: number) => [number, number, number, number];
    readonly worksheet_protectedRanges: (a: number) => [number, number, number];
    readonly worksheet_protection: (a: number) => [number, number, number];
    readonly worksheet_repeatCols: (a: number) => [number, number, number, number];
    readonly worksheet_repeatRows: (a: number) => [number, number, number, number];
    readonly worksheet_restoreValueSnapshotAt: (a: number, b: number, c: number, d: number, e: number) => [number, number];
    readonly worksheet_rowBreaks: (a: number) => [number, number, number];
    readonly worksheet_selections: (a: number) => [number, number, number];
    readonly worksheet_splitPanes: (a: number) => [number, number, number];
    readonly worksheet_tabColor: (a: number) => [number, number, number];
    readonly worksheet_tableCount: (a: number) => [number, number, number];
    readonly worksheet_tables: (a: number) => [number, number, number];
    readonly worksheet_visibility: (a: number) => [number, number, number, number];
    readonly worksheet_zoomScale: (a: number) => [number, number, number];
    readonly __wbg_cellvalue_free: (a: number, b: number) => void;
    readonly __wbg_workbook_free: (a: number, b: number) => void;
    readonly __wbg_worksheet_free: (a: number, b: number) => void;
    readonly cellvalue_asBoolean: (a: number) => number;
    readonly cellvalue_asError: (a: number) => [number, number];
    readonly cellvalue_asNumber: (a: number) => [number, number];
    readonly cellvalue_asText: (a: number) => [number, number];
    readonly cellvalue_is_boolean: (a: number) => number;
    readonly cellvalue_is_empty: (a: number) => number;
    readonly cellvalue_is_error: (a: number) => number;
    readonly cellvalue_is_number: (a: number) => number;
    readonly cellvalue_is_text: (a: number) => number;
    readonly cellvalue_toJs: (a: number) => any;
    readonly cellvalue_toString: (a: number) => [number, number];
    readonly init: () => void;
    readonly workbook_addSheet: (a: number, b: number, c: number) => [number, number, number];
    readonly workbook_calculate: (a: number, b: number) => [number, number, number];
    readonly workbook_defineName: (a: number, b: number, c: number, d: number, e: number) => [number, number];
    readonly workbook_exportSnapshot: (a: number) => [number, number, number, number];
    readonly workbook_fromBytes: (a: number, b: number) => [number, number, number];
    readonly workbook_fromBytesWithPassword: (a: number, b: number, c: number, d: number, e: number) => [number, number, number];
    readonly workbook_fromCsvString: (a: number, b: number) => [number, number, number];
    readonly workbook_fromSnapshot: (a: number, b: number) => [number, number, number];
    readonly workbook_getNamedRange: (a: number, b: number, c: number) => [number, number, number, number];
    readonly workbook_getSheet: (a: number, b: number) => [number, number, number];
    readonly workbook_getSheetByName: (a: number, b: number, c: number) => [number, number, number];
    readonly workbook_new: () => number;
    readonly workbook_removeSheet: (a: number, b: number) => [number, number];
    readonly workbook_saveCsvString: (a: number) => [number, number, number, number];
    readonly workbook_saveXlsBytesEncrypted: (a: number, b: number, c: number, d: number, e: number, f: number) => [number, number, number, number];
    readonly workbook_saveXlsbBytes: (a: number) => [number, number, number, number];
    readonly workbook_saveXlsxBytes: (a: number) => [number, number, number, number];
    readonly workbook_saveXlsxBytesEncrypted: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number, number];
    readonly workbook_setWorkbookProtection: (a: number, b: any) => [number, number];
    readonly workbook_sheetCount: (a: number) => [number, number, number];
    readonly workbook_sheetNames: (a: number) => [number, number, number, number];
    readonly worksheet_getCalculatedValue: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getCalculatedValueAt: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getCell: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getCellAt: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getColumnWidth: (a: number, b: number) => [number, number, number, number];
    readonly worksheet_getImageAt: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_getRowHeight: (a: number, b: number) => [number, number, number, number];
    readonly worksheet_mergeCells: (a: number, b: number, c: number) => [number, number];
    readonly worksheet_name: (a: number) => [number, number, number, number];
    readonly worksheet_setCell: (a: number, b: number, c: number, d: any) => [number, number];
    readonly worksheet_setCellStyle: (a: number, b: number, c: number, d: any) => [number, number];
    readonly worksheet_setCellStyleAt: (a: number, b: number, c: number, d: any) => [number, number];
    readonly worksheet_setColumnWidth: (a: number, b: number, c: number) => [number, number];
    readonly worksheet_setFormula: (a: number, b: number, c: number, d: number, e: number) => [number, number];
    readonly worksheet_setProtectedRanges: (a: number, b: any) => [number, number];
    readonly worksheet_setProtection: (a: number, b: any) => [number, number];
    readonly worksheet_setRangeStyle: (a: number, b: number, c: number, d: any) => [number, number];
    readonly worksheet_setRowHeight: (a: number, b: number, c: number) => [number, number];
    readonly worksheet_unmergeCells: (a: number, b: number, c: number) => [number, number, number];
    readonly worksheet_usedRange: (a: number) => [number, number, number];
    readonly workbook_loadCsvString: (a: number, b: number) => [number, number, number];
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
    readonly __wbindgen_exn_store: (a: number) => void;
    readonly __externref_table_alloc: () => number;
    readonly __wbindgen_externrefs: WebAssembly.Table;
    readonly __wbindgen_free: (a: number, b: number, c: number) => void;
    readonly __externref_table_dealloc: (a: number) => void;
    readonly __externref_drop_slice: (a: number, b: number) => void;
    readonly __wbindgen_start: () => void;
}

export type SyncInitInput = BufferSource | WebAssembly.Module;

/**
 * Instantiates the given `module`, which can either be bytes or
 * a precompiled `WebAssembly.Module`.
 *
 * @param {{ module: SyncInitInput }} module - Passing `SyncInitInput` directly is deprecated.
 *
 * @returns {InitOutput}
 */
export function initSync(module: { module: SyncInitInput } | SyncInitInput): InitOutput;

/**
 * If `module_or_path` is {RequestInfo} or {URL}, makes a request and
 * for everything else, calls `WebAssembly.instantiate` directly.
 *
 * @param {{ module_or_path: InitInput | Promise<InitInput> }} module_or_path - Passing `InitInput` directly is deprecated.
 *
 * @returns {Promise<InitOutput>}
 */
export default function __wbg_init (module_or_path?: { module_or_path: InitInput | Promise<InitInput> } | InitInput | Promise<InitInput>): Promise<InitOutput>;
