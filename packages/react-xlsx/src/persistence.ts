import type { Workbook } from "@dukelib/sheets-wasm";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

export interface MutationNotificationState {
  baselineRevision: number | null;
  pendingRevision: number | null;
}

export interface MutationNotificationTransition {
  notificationRevision: number | null;
  state: MutationNotificationState;
}

export function createMutationNotificationState(): MutationNotificationState {
  return {
    baselineRevision: null,
    pendingRevision: null
  };
}

/**
 * Baselines controller refreshes during workbook loading and emits only later
 * revisions. Mutations made while chart assets hydrate are queued and emitted
 * once hydration completes so an early edit cannot be lost.
 */
export function advanceMutationNotificationState(
  current: MutationNotificationState,
  input: {
    hasWorkbook: boolean;
    isChartsLoading: boolean;
    isLoading: boolean;
    revision: number;
  }
): MutationNotificationTransition {
  if (input.isLoading || !input.hasWorkbook) {
    return {
      notificationRevision: null,
      state: createMutationNotificationState()
    };
  }

  if (current.baselineRevision === null) {
    return {
      notificationRevision: null,
      state: {
        baselineRevision: input.revision,
        pendingRevision: null
      }
    };
  }

  if (current.baselineRevision !== input.revision) {
    if (input.isChartsLoading) {
      return {
        notificationRevision: null,
        state: {
          baselineRevision: input.revision,
          pendingRevision: input.revision
        }
      };
    }

    return {
      notificationRevision: input.revision,
      state: {
        baselineRevision: input.revision,
        pendingRevision: null
      }
    };
  }

  if (!input.isChartsLoading && current.pendingRevision !== null) {
    return {
      notificationRevision: current.pendingRevision,
      state: {
        baselineRevision: current.baselineRevision,
        pendingRevision: null
      }
    };
  }

  return {
    notificationRevision: null,
    state: current
  };
}

export function cloneBytes(bytes: Uint8Array): Uint8Array {
  const nextBytes = new Uint8Array(bytes.byteLength);
  nextBytes.set(bytes);
  return nextBytes;
}

function sanitizeSavedWorkbookBytes(bytes: Uint8Array): Uint8Array {
  try {
    const archive = unzipSync(bytes);
    const stylesEntry = archive["xl/styles.xml"];
    if (stylesEntry) {
      const stylesXml = strFromU8(stylesEntry)
        .replace(/&amp;quot;/g, "&quot;")
        .replace(/&amp;apos;/g, "&apos;");
      archive["xl/styles.xml"] = strToU8(stylesXml);
    }

    return zipSync(archive, { level: 6 });
  } catch {
    return cloneBytes(bytes);
  }
}

/** The single XLSX serialization path used by history, download, and persistence. */
export function createSavedWorkbookBytes(
  targetWorkbook: Workbook,
  mergeAssets: (savedBytes: Uint8Array) => Uint8Array
): Uint8Array {
  const sanitizedBytes = sanitizeSavedWorkbookBytes(targetWorkbook.saveXlsxBytes());
  return mergeAssets(sanitizedBytes);
}
