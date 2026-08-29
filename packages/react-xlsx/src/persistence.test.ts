import assert from "node:assert/strict";
import test from "node:test";
import type { Workbook } from "@dukelib/sheets-wasm";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import {
  advanceMutationNotificationState,
  createMutationNotificationState,
  createSavedWorkbookBytes,
  type MutationNotificationState
} from "./persistence.ts";

test("mutation notifications baseline hydration and emit later revisions once", () => {
  let state = createMutationNotificationState();

  let transition = advanceMutationNotificationState(state, {
    hasWorkbook: false,
    isChartsLoading: false,
    isLoading: true,
    revision: 0
  });
  assert.equal(transition.notificationRevision, null);
  state = transition.state;

  transition = advanceMutationNotificationState(state, {
    hasWorkbook: true,
    isChartsLoading: true,
    isLoading: false,
    revision: 0
  });
  assert.equal(transition.notificationRevision, null);
  assert.deepEqual(transition.state, {
    baselineRevision: 0,
    pendingRevision: null
  });
  state = transition.state;

  transition = advanceMutationNotificationState(state, {
    hasWorkbook: true,
    isChartsLoading: true,
    isLoading: false,
    revision: 1
  });
  assert.equal(transition.notificationRevision, null);
  assert.deepEqual(transition.state, {
    baselineRevision: 1,
    pendingRevision: 1
  });
  state = transition.state;

  transition = advanceMutationNotificationState(state, {
    hasWorkbook: true,
    isChartsLoading: false,
    isLoading: false,
    revision: 1
  });
  assert.equal(transition.notificationRevision, 1);
  state = transition.state;

  transition = advanceMutationNotificationState(state, {
    hasWorkbook: true,
    isChartsLoading: false,
    isLoading: false,
    revision: 1
  });
  assert.equal(transition.notificationRevision, null);
  state = transition.state;

  transition = advanceMutationNotificationState(state, {
    hasWorkbook: true,
    isChartsLoading: false,
    isLoading: false,
    revision: 2
  });
  assert.equal(transition.notificationRevision, 2);
});

test("mutation notifications reset and re-baseline when the workbook changes", () => {
  let state: MutationNotificationState = {
    baselineRevision: 4,
    pendingRevision: null
  };

  let transition = advanceMutationNotificationState(state, {
    hasWorkbook: true,
    isChartsLoading: false,
    isLoading: true,
    revision: 4
  });
  assert.equal(transition.notificationRevision, null);
  assert.deepEqual(transition.state, createMutationNotificationState());
  state = transition.state;

  transition = advanceMutationNotificationState(state, {
    hasWorkbook: true,
    isChartsLoading: false,
    isLoading: false,
    revision: 0
  });
  assert.equal(transition.notificationRevision, null);
  assert.deepEqual(transition.state, {
    baselineRevision: 0,
    pendingRevision: null
  });
});

test("saved workbook bytes use the sanitizing persistence pipeline", () => {
  const sourceBytes = zipSync({
    "[Content_Types].xml": strToU8("<Types />"),
    "xl/styles.xml": strToU8("<style>&amp;quot;quoted&amp;quot; &amp;apos;value&amp;apos;</style>"),
    "xl/worksheets/sheet1.xml": strToU8("<worksheet />")
  });
  let saveCalls = 0;
  const workbook = {
    saveXlsxBytes() {
      saveCalls += 1;
      return sourceBytes;
    }
  } as unknown as Workbook;

  let mergeCalls = 0;
  const savedBytes = createSavedWorkbookBytes(workbook, (sanitizedBytes) => {
    mergeCalls += 1;
    return new Uint8Array(sanitizedBytes);
  });
  const archive = unzipSync(savedBytes);

  assert.equal(saveCalls, 1);
  assert.equal(mergeCalls, 1);
  assert.notEqual(savedBytes, sourceBytes);
  assert.equal(
    strFromU8(archive["xl/styles.xml"]),
    "<style>&quot;quoted&quot; &apos;value&apos;</style>"
  );
  assert.equal(strFromU8(archive["xl/worksheets/sheet1.xml"]), "<worksheet />");
  assert.equal(
    strFromU8(unzipSync(sourceBytes)["xl/styles.xml"]),
    "<style>&amp;quot;quoted&amp;quot; &amp;apos;value&amp;apos;</style>"
  );
});
