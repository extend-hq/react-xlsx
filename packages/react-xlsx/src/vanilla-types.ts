import type { Properties } from "csstype";
import type {
  UseXlsxViewerControllerOptions,
  XlsxActiveCellOverlayRenderProps,
  XlsxCellAddress,
  XlsxChart,
  XlsxContextMenuAction,
  XlsxContextMenuActionId,
  XlsxContextMenuContext,
  XlsxFormControl,
  XlsxFormControlActionEvent,
  XlsxFormControlChangeEvent,
  XlsxFormControlState,
  XlsxImage,
  XlsxImageRect,
  XlsxViewerController
} from "./types";

export type XlsxVanillaContent = Node | string | number | null;
export type XlsxVanillaStyle = Properties<string | number>;

export interface XlsxVanillaChartLoadingRenderProps {
  chart: XlsxChart;
  rect: XlsxImageRect;
  stage?: "metadata" | "renderer" | "map";
}

export interface XlsxVanillaContextMenuRenderProps {
  context: XlsxContextMenuContext;
  actions: XlsxContextMenuAction[];
  runAction: (actionId: XlsxContextMenuActionId) => Promise<boolean>;
  closeMenu: () => void;
  position: { x: number; y: number };
  isDark: boolean;
  busy: boolean;
  error: string | null;
}

export interface XlsxVanillaFileTooLargeRenderProps {
  displayFileName: string;
  fileSizeBytes: number;
  maxFileSizeBytes: number;
}

export interface XlsxVanillaImageRenderProps {
  image: XlsxImage;
  rect: XlsxImageRect;
  style: XlsxVanillaStyle;
}

export interface XlsxVanillaFormControlRenderProps {
  activate: () => void;
  checked: boolean;
  control: XlsxFormControl;
  disabled: boolean;
  items: string[];
  label: string;
  mixed: boolean;
  readOnly: boolean;
  rect: XlsxImageRect;
  setSelected: (selected: number | number[] | undefined) => boolean;
  setState: (state: XlsxFormControlState) => boolean;
  setValue: (value: number) => boolean;
  stopPropagation: (event: Event) => void;
  style: XlsxVanillaStyle;
  zoomFactor: number;
}

export interface XlsxVanillaCellStyleContext {
  cell: XlsxCellAddress;
  hasChartHighlight: boolean;
  hasConditionalFormat: boolean;
  hasHyperlink: boolean;
  hasValidation: boolean;
  isMerged: boolean;
  isTableHeader: boolean;
  resolvedStyle: XlsxVanillaStyle;
  sheetName: string;
  value: string;
  workbookSheetIndex: number;
}

export interface XlsxVanillaViewerOptions extends UseXlsxViewerControllerOptions {
  highlightedRanges?: string | readonly string[] | null;
  autoScrollToHighlightedRanges?: boolean;
  highlightClassName?: string;
  className?: string;
  emptyState?: XlsxVanillaContent;
  enableCanvasSelectionAnimation?: boolean;
  enableGestureZoom?: boolean;
  experimentalCanvas?: boolean;
  errorState?: XlsxVanillaContent | ((error: Error) => XlsxVanillaContent);
  fileTooLargeState?: XlsxVanillaContent | ((props: XlsxVanillaFileTooLargeRenderProps) => XlsxVanillaContent);
  getCellStyle?: (context: XlsxVanillaCellStyleContext) => XlsxVanillaStyle | null | undefined;
  headerBackgroundColor?: string;
  headerTextColor?: string;
  height?: XlsxVanillaStyle["height"];
  isDark?: boolean;
  loadingState?: XlsxVanillaContent;
  onFormControlAction?: (event: XlsxFormControlActionEvent) => void;
  onFormControlChange?: (event: XlsxFormControlChangeEvent) => void;
  renderChartLoading?: (props: XlsxVanillaChartLoadingRenderProps) => XlsxVanillaContent;
  renderContextMenu?: (props: XlsxVanillaContextMenuRenderProps) => XlsxVanillaContent;
  enableContextMenu?: boolean;
  contextMenuClassName?: string;
  contextMenuStyle?: XlsxVanillaStyle & { [property: `--${string}`]: string | number | undefined };
  renderFormControl?: (props: XlsxVanillaFormControlRenderProps) => XlsxVanillaContent;
  renderImage?: (props: XlsxVanillaImageRenderProps) => XlsxVanillaContent;
  rounded?: boolean;
  selectionColor?: string;
  selectionFillColor?: string;
  selectionHeaderColor?: string;
  renderActiveCellOverlay?: (props: XlsxActiveCellOverlayRenderProps) => XlsxVanillaContent;
  showImages?: boolean;
  showDefaultToolbar?: boolean;
  toolbar?: XlsxVanillaContent | ((controller: XlsxViewerController) => XlsxVanillaContent);
}

export interface XlsxVanillaViewer {
  readonly element: HTMLElement;
  getController(): XlsxViewerController | null;
  reload(): void;
  setOptions(options: Partial<XlsxVanillaViewerOptions>): void;
  subscribe(listener: (controller: XlsxViewerController) => void): () => void;
  destroy(): void;
}
