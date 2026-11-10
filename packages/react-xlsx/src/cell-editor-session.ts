import * as React from "react";
import type { XlsxCellAddress, XlsxCellEditorContext, XlsxCellEditorSession } from "./types";

export function useCellEditorSession({
  enabled, cell, sheetIndex, sheetName, originalValue, value, inputRef,
  isEditing, onChange, onCommit, onCancel
}: {
  enabled: boolean;
  cell: XlsxCellAddress | null;
  sheetIndex: number;
  sheetName: string;
  originalValue: string;
  value: string;
  inputRef: React.RefObject<HTMLTextAreaElement | null>;
  isEditing: () => boolean;
  onChange: (value: string) => void;
  onCommit: (value: string) => void;
  onCancel: () => void;
}) {
  const state = React.useMemo(() => enabled && cell ? {
    controller: new AbortController(),
    context: { cell: { ...cell }, sheetIndex, sheetName, originalValue } satisfies XlsxCellEditorContext,
    beforePreview: null as string | null
  } : null, [enabled, cell, sheetIndex]);
  const latest = React.useRef({ state, value, isEditing, onChange, onCommit, onCancel });
  latest.current = { state, value, isEditing, onChange, onCommit, onCancel };
  const [, refresh] = React.useReducer((revision: number) => revision + 1, 0);

  React.useLayoutEffect(() => () => {
    state?.controller.abort();
    if (state?.beforePreview != null && latest.current.state === null && latest.current.isEditing()) {
      latest.current.onCancel();
    }
  }, [state]);

  const commands = React.useMemo(() => {
    if (!state) return null;
    const current = () => latest.current.state === state && !state.controller.signal.aborted && latest.current.isEditing();
    const setValue = (next: string) => {
      if (!current()) return false;
      latest.current.value = next;
      latest.current.onChange(next);
      return true;
    };
    const focus = () => {
      if (!current() || !inputRef.current) return false;
      inputRef.current.focus({ preventScroll: true });
      return true;
    };
    return {
      getValue: () => current() ? latest.current.value : null,
      setValue,
      focus,
      preview(next: string) {
        if (!current()) return false;
        state.beforePreview ??= latest.current.value;
        setValue(next);
        refresh();
        return true;
      },
      discardPreview() {
        if (!current() || state.beforePreview === null) return false;
        const previous = state.beforePreview;
        state.beforePreview = null;
        setValue(previous);
        refresh();
        focus();
        return true;
      },
      commit(next?: string) {
        if (!current()) return false;
        state.controller.abort();
        latest.current.onCommit(next ?? latest.current.value);
        return true;
      },
      cancel() {
        if (!current()) return false;
        state.controller.abort();
        latest.current.onCancel();
        return true;
      }
    };
  }, [state, inputRef]);
  const hasPreview = React.useCallback(() => latest.current.state?.beforePreview != null, []);
  const session: XlsxCellEditorSession | null = state && commands ? {
    ...commands,
    context: state.context,
    signal: state.controller.signal,
    value,
    isPreview: state.beforePreview !== null
  } : null;
  return { session, hasPreview };
}
