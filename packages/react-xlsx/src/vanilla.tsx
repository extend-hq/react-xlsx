import * as React from "react";
import { createRoot } from "react-dom/client";
import { XlsxViewer as InternalViewer } from "./XlsxViewer";
import { useXlsxViewerController } from "./controller";
import type { XlsxViewerController, XlsxViewerProps } from "./types";
import type { XlsxVanillaContent, XlsxVanillaViewer, XlsxVanillaViewerOptions } from "./vanilla-types";

export { XlsxFileSizeLimitExceededError } from "./controller";
export { parseXlsxRange } from "./range-reference";
export { initWasm, setWasmSource, type XlsxWasmSource } from "./wasm";
export type * from "./vanilla-types";

function DomContent({ node }: { node: Node }) {
  const hostRef = React.useRef<HTMLSpanElement>(null);

  React.useLayoutEffect(() => {
    const host = hostRef.current;
    if (!host) return;

    host.replaceChildren(node);
    return () => {
      if (node.parentNode === host) host.removeChild(node);
    };
  }, [node]);

  return <span ref={hostRef} style={{ display: "contents" }} />;
}

function toReactContent(content: XlsxVanillaContent | undefined): React.ReactNode {
  return content !== null && typeof content === "object" ? <DomContent node={content} /> : content;
}

function toReactOptions(options: XlsxVanillaViewerOptions): XlsxViewerProps {
  const {
    emptyState,
    errorState,
    fileTooLargeState,
    loadingState,
    renderActiveCellOverlay,
    renderChartLoading,
    renderContextMenu,
    renderFormControl,
    renderImage,
    toolbar,
    ...viewerOptions
  } = options;

  return {
    ...viewerOptions,
    emptyState: toReactContent(emptyState),
    errorState: typeof errorState === "function"
      ? (error) => toReactContent(errorState(error))
      : toReactContent(errorState),
    fileTooLargeState: typeof fileTooLargeState === "function"
      ? ({ defaultNode: _defaultNode, ...props }) => toReactContent(fileTooLargeState(props))
      : toReactContent(fileTooLargeState),
    loadingState: toReactContent(loadingState),
    renderActiveCellOverlay: renderActiveCellOverlay
      ? (props) => toReactContent(renderActiveCellOverlay(props))
      : undefined,
    renderChartLoading: renderChartLoading
      ? ({ defaultNode: _defaultNode, ...props }) => toReactContent(renderChartLoading(props))
      : undefined,
    renderContextMenu: renderContextMenu
      ? ({ defaultNode: _defaultNode, ...props }) => toReactContent(renderContextMenu(props))
      : undefined,
    renderFormControl: renderFormControl
      ? (props) => toReactContent(renderFormControl({ ...props, stopPropagation: (event) => event.stopPropagation() }))
      : undefined,
    renderImage: renderImage
      ? ({ defaultNode: _defaultNode, ...props }) => toReactContent(renderImage(props))
      : undefined,
    toolbar: typeof toolbar === "function"
      ? (controller) => toReactContent(toolbar(controller))
      : toReactContent(toolbar)
  };
}

function VanillaViewerRoot({
  generation,
  options,
  publish
}: {
  generation: number;
  options: XlsxVanillaViewerOptions;
  publish: (controller: XlsxViewerController, generation: number) => void;
}) {
  const controller = useXlsxViewerController(options);
  const viewerOptions = React.useMemo(() => toReactOptions(options), [options]);

  React.useLayoutEffect(() => publish(controller, generation), [controller, generation, publish]);

  return <InternalViewer {...viewerOptions} controller={controller} />;
}

function mergeOptions(
  current: XlsxVanillaViewerOptions,
  next: Partial<XlsxVanillaViewerOptions>
): XlsxVanillaViewerOptions {
  if (next.file !== undefined && next.src !== undefined) {
    throw new TypeError("Provide either file or src, not both.");
  }

  const merged = { ...current, ...next };
  if (next.file !== undefined) merged.src = undefined;
  if (next.src !== undefined) merged.file = undefined;
  return merged;
}

function createViewerElement(container: Element | DocumentFragment) {
  const owner = container.ownerDocument;
  const element = owner.createElement("div");
  element.style.display = "contents";
  container.appendChild(element);
  return element;
}

function createViewer(
  container: Element | DocumentFragment,
  initialOptions: XlsxVanillaViewerOptions = {}
): XlsxVanillaViewer {
  let options = mergeOptions({}, initialOptions);
  const element = createViewerElement(container);
  const root = createRoot(element);
  const listeners = new Set<(controller: XlsxViewerController) => void>();
  let controller: XlsxViewerController | null = null;
  let destroyed = false;
  let notificationPending = false;
  let generation = 0;

  const publish = (next: XlsxViewerController, publishedGeneration: number) => {
    if (destroyed || publishedGeneration !== generation) return;
    controller = next;
    if (notificationPending) return;
    notificationPending = true;
    queueMicrotask(() => {
      notificationPending = false;
      if (destroyed || !controller) return;
      const current = controller;
      for (const listener of listeners) listener(current);
    });
  };

  const render = () => root.render(
    <VanillaViewerRoot key={generation} generation={generation} options={options} publish={publish} />
  );
  const assertActive = () => {
    if (destroyed) throw new Error("The viewer has been destroyed.");
  };

  render();

  return {
    element,
    getController: () => controller,
    reload() {
      assertActive();
      generation += 1;
      controller = null;
      render();
    },
    setOptions(next) {
      assertActive();
      const updated = mergeOptions(options, next);
      if (updated.file !== options.file || updated.src !== options.src) {
        generation += 1;
        controller = null;
      }
      options = updated;
      render();
    },
    subscribe(listener) {
      assertActive();
      listeners.add(listener);
      if (controller) listener(controller);
      return () => listeners.delete(listener);
    },
    destroy() {
      if (destroyed) return;
      destroyed = true;
      listeners.clear();
      controller = null;
      root.unmount();
      element.remove();
    }
  };
}

export class XlsxViewer implements XlsxVanillaViewer {
  readonly element: HTMLElement;
  private readonly viewer: XlsxVanillaViewer;

  constructor(container: Element | DocumentFragment, options: XlsxVanillaViewerOptions = {}) {
    this.viewer = createViewer(container, options);
    this.element = this.viewer.element;
  }

  getController() { return this.viewer.getController(); }
  setOptions(options: Partial<XlsxVanillaViewerOptions>) { this.viewer.setOptions(options); }
  reload() { this.viewer.reload(); }
  subscribe(listener: (controller: XlsxViewerController) => void) { return this.viewer.subscribe(listener); }
  destroy() { this.viewer.destroy(); }
}

export function createXlsxViewer(
  container: Element | DocumentFragment,
  options: XlsxVanillaViewerOptions = {}
): XlsxViewer {
  return new XlsxViewer(container, options);
}

export function __createControlledViewer(container: Element | DocumentFragment) {
  const element = createViewerElement(container);
  const root = createRoot(element);
  return {
    update(props: unknown) { root.render(<InternalViewer {...props as XlsxViewerProps} />); },
    destroy() { root.unmount(); element.remove(); }
  };
}

export function __createContentNode(node: Node): unknown {
  return toReactContent(node);
}

export function __createContentRoot(container: Element) {
  const root = createRoot(container);
  return {
    update(content: unknown) { root.render(content as React.ReactNode); },
    destroy() { root.unmount(); }
  };
}
