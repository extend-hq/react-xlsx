import * as React from "react";
import { createPortal } from "react-dom";
import { __createContentNode, __createContentRoot, __createControlledViewer } from "@extend-ai/react-xlsx/internal";
import { useXlsxViewerController } from "./controller";
import { ViewerAppearanceContext, ViewerContext } from "./XlsxViewer";
import type { XlsxViewerController, XlsxViewerProps } from "./types";

type Slot = { host: HTMLElement; node: React.ReactNode; wasConnected: boolean };

class SlotStore {
  private readonly slots = new Map<string, Slot>();
  private readonly listeners = new Set<() => void>();
  private observer: MutationObserver | null = null;
  private revision = 0;
  private scheduled = false;
  private destroyed = false;

  readonly subscribe = (listener: () => void) => {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  };

  readonly getSnapshot = () => this.revision;
  readonly getServerSnapshot = () => 0;

  observe(element: HTMLElement) {
    this.destroyed = false;
    const target = element.ownerDocument.documentElement;
    this.observer = new MutationObserver(() => this.sweep());
    this.observer.observe(target, { childList: true, subtree: true });
  }

  render(key: string, node: React.ReactNode, owner: Document) {
    let slot = this.slots.get(key);
    if (!slot) {
      const host = owner.createElement("span");
      host.style.display = "contents";
      slot = { host, node, wasConnected: false };
      this.slots.set(key, slot);
    } else {
      slot.node = node;
    }
    this.notify();
    return __createContentNode(slot.host) as React.ReactNode;
  }

  portals() {
    return Array.from(this.slots, ([key, slot]) => createPortal(slot.node, slot.host, key));
  }

  destroy() {
    this.destroyed = true;
    this.observer?.disconnect();
    this.observer = null;
    this.slots.clear();
    this.listeners.clear();
  }

  private sweep() {
    let changed = false;
    for (const [key, slot] of this.slots) {
      if (slot.host.isConnected) {
        slot.wasConnected = true;
      } else if (slot.wasConnected) {
        this.slots.delete(key);
        changed = true;
      }
    }
    if (changed) this.notify();
  }

  private notify() {
    if (this.destroyed || this.scheduled) return;
    this.scheduled = true;
    queueMicrotask(() => {
      this.scheduled = false;
      if (this.destroyed) return;
      this.revision += 1;
      this.listeners.forEach((listener) => listener());
    });
  }
}

function RuntimeContent({ content }: { content: React.ReactNode }) {
  const hostRef = React.useRef<HTMLSpanElement>(null);
  const rootRef = React.useRef<ReturnType<typeof __createContentRoot> | null>(null);

  React.useLayoutEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const root = __createContentRoot(host);
    rootRef.current = root;
    return () => {
      rootRef.current = null;
      root.destroy();
    };
  }, []);
  React.useLayoutEffect(() => { rootRef.current?.update(content); }, [content]);

  return <span ref={hostRef} style={{ display: "contents" }} />;
}

function renderOptions(
  props: XlsxViewerProps,
  controller: XlsxViewerController,
  slots: SlotStore,
  owner: Document
): XlsxViewerProps {
  const slot = (key: string, node: React.ReactNode) => slots.render(key, node, owner);
  const runtime = (node: React.ReactNode) => <RuntimeContent content={node} />;
  const {
    emptyState, errorState, fileTooLargeState, loadingComponent, loadingState,
    renderActiveCellOverlay, renderChartLoading, renderContextMenu, renderFormControl,
    renderImage, renderImageSelection, renderScroller, renderTableHeaderMenu, toolbar,
    ...other
  } = props;

  return {
    ...other,
    controller,
    emptyState: emptyState === undefined ? undefined : slot("empty", emptyState),
    loadingComponent: loadingComponent === undefined ? undefined : slot("loading-component", loadingComponent) as React.ReactElement,
    loadingState: loadingState === undefined ? undefined : slot("loading", loadingState),
    errorState: errorState === undefined ? undefined : typeof errorState === "function"
      ? (error) => slot("error", errorState(error))
      : slot("error", errorState),
    fileTooLargeState: fileTooLargeState === undefined ? undefined : typeof fileTooLargeState === "function"
      ? ({ defaultNode, ...rest }) => slot("too-large", fileTooLargeState({ ...rest, defaultNode: runtime(defaultNode) }))
      : slot("too-large", fileTooLargeState),
    toolbar: toolbar === undefined ? undefined : typeof toolbar === "function"
      ? (current) => slot("toolbar", toolbar(current))
      : slot("toolbar", toolbar),
    renderActiveCellOverlay: renderActiveCellOverlay
      ? (value) => slot(`active:${value.cell.row}:${value.cell.col}`, renderActiveCellOverlay(value))
      : undefined,
    renderChartLoading: renderChartLoading
      ? ({ defaultNode, ...value }) => slot(`chart:${value.chart.id}`, renderChartLoading({ ...value, defaultNode: runtime(defaultNode) }))
      : undefined,
    renderContextMenu: renderContextMenu
      ? ({ defaultNode, ...value }) => slot("context-menu", renderContextMenu({ ...value, defaultNode: runtime(defaultNode) }))
      : undefined,
    renderFormControl: renderFormControl
      ? (value) => slot(`control:${value.control.id}`, renderFormControl(value))
      : undefined,
    renderImage: renderImage
      ? ({ defaultNode, ...value }) => slot(`image:${value.image.id}`, renderImage({ ...value, defaultNode: runtime(defaultNode) }))
      : undefined,
    renderImageSelection: renderImageSelection
      ? ({ defaultNode, ...value }) => slot(`image-selection:${value.image.id}`, renderImageSelection({ ...value, defaultNode: runtime(defaultNode) }))
      : undefined,
    renderScroller: renderScroller
      ? ({ children, ...value }) => slot("scroller", renderScroller({ ...value, children: runtime(children) }))
      : undefined,
    renderTableHeaderMenu: renderTableHeaderMenu
      ? (value) => slot(`table:${value.table.name}:${value.cell.row}:${value.cell.col}`, renderTableHeaderMenu(value))
      : undefined
  };
}

function BoundViewer({ props, controller }: { props: XlsxViewerProps; controller: XlsxViewerController }) {
  const hostRef = React.useRef<HTMLDivElement>(null);
  const viewerRef = React.useRef<ReturnType<typeof __createControlledViewer> | null>(null);
  const slots = React.useMemo(() => new SlotStore(), []);
  React.useSyncExternalStore(slots.subscribe, slots.getSnapshot, slots.getServerSnapshot);

  React.useLayoutEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    slots.observe(host);
    const viewer = __createControlledViewer(host);
    viewerRef.current = viewer;
    return () => {
      viewerRef.current = null;
      viewer.destroy();
      slots.destroy();
    };
  }, [slots]);

  React.useLayoutEffect(() => {
    const host = hostRef.current;
    if (host) viewerRef.current?.update(renderOptions(props, controller, slots, host.ownerDocument));
  }, [controller, props, slots]);

  return (
    <ViewerAppearanceContext.Provider value={{
      headerBackgroundColor: props.headerBackgroundColor,
      headerTextColor: props.headerTextColor,
      isDark: props.isDark ?? false
    }}>
      <ViewerContext.Provider value={controller}>
        <div ref={hostRef} style={{ display: "contents" }} />
        {slots.portals()}
      </ViewerContext.Provider>
    </ViewerAppearanceContext.Provider>
  );
}

function UncontrolledViewer(props: XlsxViewerProps) {
  const controller = useXlsxViewerController(props);
  return <BoundViewer props={props} controller={controller} />;
}

export function XlsxViewer(props: XlsxViewerProps) {
  const context = React.useContext(ViewerContext);
  const controller = props.controller ?? context;
  return controller
    ? <BoundViewer props={props} controller={controller} />
    : <UncontrolledViewer {...props} />;
}
