# Worksheet-backed pivot tables

Pivot operations run in the patched DukeSheets Rust engine. React and vanilla controllers expose the same workbook operations. The playground provides a PivotTable ribbon and field panel; native saved report cells render in both canvas and DOM modes.

## Editing workflow

Select a source range with unique, nonempty column headers, choose **Insert → PivotTable**, choose its destination, and place fields in Rows, Columns, Values, or Filters. Changes remain a draft until **Create PivotTable** or **Apply Changes**. The panel also supports field dragging, keyboard reorder buttons, grouping, member selection, label/value filters, calculation settings, and report layout.

Editing an existing report uses its saved cache. **Refresh** reads the source again. **Refresh all** stages all changes together. Source edits do not silently refresh a pivot. Creating, changing, refreshing, removing, and drilling through a pivot participates in undo/redo. Pivot output cells reject direct value/formula edits, paste, fill, merge, and ordinary table sorting; formatting remains editable.

Mutation work is staged before committing it to the React controller. Invalid ranges, overlapping reports, occupied destinations, protected sheets, failed calculations, and unsupported configurations report an error while retaining the original workbook. Refresh can resize a report, but cannot overwrite unrelated occupied cells.

Refresh-all recalculates formulas between reports. If one pivot uses another pivot's results as its source, refresh that chain in source order; automatic dependency ordering is not implemented.

## Public API

`useXlsxViewerPivots()` exposes these methods. They also exist on `useXlsxViewer()` and the vanilla viewer controller.

| Member | Behavior |
| --- | --- |
| `pivotTables` | Workbook-wide pivot metadata, effective field members, output range, chart bindings, and unsupported-feature reasons. |
| `getPivotSourceFields(source)` | Read field names and typed member keys from a worksheet range, table, or supported static name. |
| `getPivotTableAtCell(cell, sheetName?)` | Find the pivot containing a worksheet cell. |
| `pivotFieldListVisible` / `setPivotFieldListVisible(visible)` | Read or change field-list visibility for built-in or custom controls. |
| `upsertPivotTable(input)` | Create or configure a pivot; returns its metadata. |
| `refreshPivotTable(id)` | Rebuild the pivot cache and report from source data. |
| `refreshAllPivotTables()` | Refresh all reports as one undoable transaction. |
| `removePivotTable(id)` | Remove the report and detach its PivotCharts. |
| `pivotDrillDown(id, row, col)` | Read the source records behind a result cell, using zero-based worksheet coordinates. |
| `showPivotDetails(id, row, col)` | Create a details worksheet containing those records. |
| `togglePivotGroup(id, toggle)` | Expand or collapse a group from the report metadata. |
| `createPivotChart(id, options?)` | Create a native chart bound to the pivot; refresh keeps its data synchronized. |

Mutating methods throw on failure or when editing is disabled. Input and metadata types are exported from both package entry points, including `XlsxPivotTable`, `XlsxPivotTableInput`, `XlsxPivotSource`, `XlsxPivotAxisField`, `XlsxPivotValueField`, and `XlsxPivotFilter`. Field indexes are zero-based source indexes. Filter/base-item keys come from the returned member metadata; captions are not interchangeable with typed keys.

Calculated fields use `sourceName` to identify their formula independently from the display caption. Date hierarchy levels can share a source field; their effective names and grouping units identify the level. Values can appear on rows or columns. `preserveFormatting` defaults to true. `refreshOnLoad` requests automatic refresh when the viewer opens the workbook; a failed automatic refresh exposes `refreshError` and retains saved results.

## Custom field panes and controls

The built-in pane is hidden by default. Set `showPivotTableControls={true}` to enable the Excel-style field pane when a pivot is selected. Replace that pane with `renderPivotTableControls`, or leave it hidden and render your own pane elsewhere. The render context includes the editor, workbook controller, appearance, and `defaultNode` for composing the built-in pane with other content.

`useXlsxPivotTableEditor()` is a headless hook used inside `XlsxViewerProvider`. It exposes the selected draft, open/dirty/stale states, errors, details, and actions for source inspection, field placement and reorder, editing, apply/discard, refresh, and removal. `edit()` accepts an updater for any supported pivot setting. It also exposes `viewer` for chart, drill-through, export, and other workbook actions. Supply the returned editor through `pivotTableEditor` so a custom toolbar and pane share the same draft and validation state.

`pivotFieldListVisible` and `setPivotFieldListVisible()` are shared by the controller, `useXlsxViewerPivots()`, the editor, and the context menu's checked Show/Hide Field List action. Selecting a pivot automatically opens the field list by default. Clicking outside hides it temporarily; returning to a pivot restores it. Explicitly hiding or closing the pane keeps it hidden until Show Field List is chosen again. This preference lasts for the controller's lifetime, including sheet and workbook changes, and works independently of the built-in-pane option. A custom editor can opt out of automatic opening with `useXlsxPivotTableEditor({ initialOpen: false })`.

```tsx
function PivotWorkspace() {
  const editor = useXlsxPivotTableEditor();
  return <>
    <button onClick={editor.create}>Insert PivotTable</button>
    <XlsxViewer
      pivotTableEditor={editor}
      showPivotTableControls
      renderPivotTableControls={({ editor, defaultNode }) => (
        <MyFieldPane editor={editor} fallback={defaultNode} />
      )}
    />
  </>;
}
```

`DefaultXlsxPivotTableFieldList` can also be rendered independently with an editor, `className`, `style`, and `isDark`. The playground keeps the built-in pane disabled and renders a separate resizable COSS pane from the public hook. Vanilla viewers accept `renderPivotTableControls` returning DOM content and receive the same editor actions and workbook controller.

### Header menus and selection guides

`renderTableHeaderMenu` customizes both table and pivot header triggers and their popovers. Its `kind` identifies the header type. Pivot headers also provide `pivot`, containing the field selector, typed members, filters, sort state, read-only/error state, controlled open state, and actions for sorting by label or measure, member selection, label/value/top filters, and clearing filters. `renderPivotTableHeaderMenu` optionally overrides only pivot headers.

Forward `triggerProps` to the custom trigger and `pivot.popupProps` to the custom popup so focus, keyboard handling, and outside clicks continue to work. Nested popup portals should use the popup element as their container. `defaultNode` is available to retain the built-in controls. The playground demonstrates this shared callback with COSS triggers, table menus, and pivot filter popovers.

Pivot field guides appear while the field list is visible and a pivot cell is active. They outline the row fields, column fields, and data regions above the worksheet contents, while preserving the active-cell border. `pivotTableHighlightColor` customizes their color. Header buttons remain available with the field list hidden and are clipped at frozen-pane boundaries.

## Compatibility boundary

This implementation targets worksheet-backed XLSX pivots. Native Data Model/OLAP/external connections require additional engines or services and are preserved without editing. Reports with unsupported native slicer/timeline dependencies also remain read-only so changes cannot invalidate their cache indexes. The playground's visual member/date filters edit pivot filters; they do not create standalone Excel slicer or timeline objects.

The engine supports sum, count, count numbers, average, minimum, maximum, product, variance and standard-deviation aggregates; row/column/report fields; multiple values; compact/outline/tabular layouts; group expansion; totals and custom subtotals; sorting; member/label/value/top filters; numeric/manual/date grouping; calculated fields/items; and the exposed Show Values As modes. Some combinations are rejected where Excel also imposes restrictions. Calculated expressions support arithmetic, parentheses, field/item references, and SUM, PRODUCT, AVERAGE, MIN, MAX, ABS, ROUND, and SQRT. Unsupported functions, comparisons, and conditional expressions are rejected before changing the report. Grouped fields must be filtered using their grouped members; combining a raw-source filter with grouping is rejected because Excel refresh would produce different results.

Pivot styles use approximate palettes for viewer rendering. Original saved cell formatting and unsupported package data are preserved when possible, but exact Excel visual parity is not established. Unrecognized features are surfaced in `unsupportedFeatures` rather than silently flattened into an editable report. Saved pivot results can still be viewed.

Exports contain actual pivot definitions, cache definitions/records, relationships, and PivotChart associations. Excel interoperability requires validating both an initial open and a refresh: a workbook that merely displays saved result cells does not establish editable-pivot compatibility.
