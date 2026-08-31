## Connector Picker, Snapshot, and Ref Des Rename

### Purpose

This subsystem wires the library (`05_library_core.md`) and the connector
editor (`06_connector_editor.md`) into the Creator's actual workflow: Add
Connector picks a library part (or defines a new one) and freezes its
definition into `_Snapshot`; Manage Library browses, edits, and deletes
library entries; Remove Connector drops a placed instance; and ref des
rename rewrites every chart reference and rejects collisions.

### `_Snapshot`'s three fixed regions

`_Snapshot` (very hidden, code name `shSnapshot`) holds three regions on
one physical sheet, all sized in `build/layout.py`:

| Region | Rows | Columns |
|---|---|---|
| Connectors | 2-201 (header row 1) | Same 11 columns as the library's `Connectors` sheet |
| Pins | 211-2210 (header row 210) | Same 7 columns as the library's `Pins` sheet |
| Photo shapes | n/a | Arbitrary pixel positions via `modLibrary.EmbedConnectorPhoto`'s grid, overlapping the data rows cosmetically - harmless, since the sheet is never displayed and shapes are a separate object model from cells |

One sheet rather than three because `_Snapshot` is meant to be a single
frozen unit copied out of `HarnessCreator.xlsm` conceptually, not a set of
independent tables - and because `modLibrary`'s bounded-window functions
(`05_library_core.md`) already support multiple tables coexisting on one
sheet by construction. 200 connectors and 2000 pin rows is generous
headroom for what one harness actually uses (a handful of connectors)
while staying a small, fixed cost on a sheet nobody scrolls through by
hand.

### `SnapshotConnector`'s idempotency

`modSnapshot.SnapshotConnector` (`src/vba/modSnapshot.bas:13-69`) is frozen
once per distinct ConnectorID, not per ref des instance: it checks
`FindConnectorRow` against the Connectors region first and returns `True`
immediately if that ID is already present, before reading anything from the
library. Two ref des instances of the same library part (`J1` and `J2` both
`DTM-04P`) share one snapshot row and one pin block - the second `Add
Connector` call is a no-op against `_Snapshot`, even though it still
allocates its own ref des and its own `Connectors`-sheet instance row via
`modConnectors.AddConnectorInstance`.

#### Photo-cache lookup order

`SnapshotConnector` (`src/vba/modSnapshot.bas:39-62`, calling `modLibrary.
CachePhotoPath` and `modLibrary.ExportShapeToFile`) checks in this order:

1. The `.jpg` cache (`CachePhotoPath(..., "jpg")`) - the cache the
   connector editor's Save always writes on a successful save, so this is
   the reliable, plain-file path for any connector saved through the
   editor.
2. The `.png` cache - the extension used before the `.jpg` cache existed;
   kept so an older connector still finds its photo.
3. `modLibrary.ExportShapeToFile` against the library's embedded `Shape` -
   the clipboard-dependent Shape-Copy/Chart-Paste mechanism
   (`05_library_core.md`), unreliable for VBA-triggered operations on this
   machine, used only as a last resort when neither cache file exists.

`tests/test_snapshot.py`'s
`test_snapshot_connector_prefers_the_jpg_cache_over_reexporting_the_shape`
seeds only the `.jpg` cache directly (no clipboard) and asserts
`SnapshotConnector` finds and uses it rather than falling through.

### The `SelectionChange`-caches-prior-value technique

A plain `Worksheet_Change` handler only ever sees a cell's *new* value -
there is no built-in "previous value" to compare against. `shConnectors.evt`
(`src/vba/sheets/shConnectors.evt`) caches the ref des a cell held
*immediately before* it was edited, in `Worksheet_SelectionChange`
(`mLastRefDes`, `mLastRefDesRow`), then compares against that cache when
`Worksheet_Change` fires. This is the only reliable way to detect a rename
on a plain cell edit without a custom UI: by the time `Worksheet_Change`
runs, the cell already holds the new value, and the row's identity (having
moved, if compaction from a delete happened concurrently) is tracked by
row number captured at selection time.

The actual rename decision lives in `modConnectors.ApplyConnectorEdit`
(`src/vba/modConnectors.bas:256-286`), not in the sheet module itself, and
it does two things on every `Connectors`-sheet edit, not just a rename:
before any rename logic runs at all, it calls
`modChart.RefreshChartRowsForConnector` for every edited row
(`src/vba/modConnectors.bas:262-267`) - this is what keeps a chart row's
From/To Pin dropdown in sync when, say, a connector's Pin Count is edited
directly on the `Connectors` sheet rather than through the editor. Only
after that does it check whether this specific edit is a rename at all
(single cell, column A, matching the cached row, a real value change), and
if so delegates to `modConnectors.RenameRefDes`. `ApplyConnectorEdit`
returns a `modContract` result (`RENAMED` with the new ref des as payload,
`RENAME_REJECTED` with the reverted value as payload, or `NO_RENAME` for
anything else), and the adapter (`Worksheet_Change`) only ever branches on
that outcome - reverting the cell on `RENAME_REJECTED`, updating its own
cache on `RENAMED`. This is a departure worth noting from the strict layer
split: `modConnectors` is a layer 0 module
(`docs/design/01_architecture_overview.md`) that here returns a
layer-1-shaped `modContract` envelope, and the adapter calls it directly
rather than through a layer 1 action module -
`docs/design/04_enforcement.md`'s allowlist explicitly permits this
(`"shConnectors": {"modConnectors.CONN_FIRST_ROW", "modConnectors.
ApplyConnectorEdit"}`, `tests/test_layering.py:40`) because renaming is
tightly coupled to the sheet-event lifecycle (prior-value caching) in a way
that doesn't cleanly decompose into a separate action module without
duplicating that state.

`RenameRefDes` itself (`src/vba/modConnectors.bas:97-130`) rejects a
collision by counting matches: because the sheet edit has already happened
by the time this runs, the renamed row already carries the new ref des, so
exactly one match across the whole `Connectors` sheet is the non-colliding
case - more than one means a different row already used that value.

### The picker/browser split

Three forms, one function each button click delegates to:

- **`frmConnectorPicker`** ("Add Connector"): pick an existing library part,
  or define a new one. `cmdAdd_Click` calls `modPickerActions.AddFromLibrary`
  (`src/vba/modPickerActions.bas:8-29`) - one transaction that allocates a
  ref des (`modConnectors.AddConnectorInstance`) and freezes the definition
  (`modSnapshot.SnapshotConnector`).
- **`frmManageLibrary`** ("Manage Library"): edit, delete, import, and
  export library connectors (import/export are covered in
  `08_import_export.md`). `cmdDelete_Click` calls `modManageActions.
  DeleteFromLibrary`.
- **`frmRemoveConnector`** ("Remove Connector"): drops one placed instance
  from the current harness, presenting every placed instance as a list
  (`modConnectors.InstanceIndex`) for a consistent picker experience with
  Add Connector and Manage Library. `cmdRemove_Click` calls
  `modConnectorActions.RemoveInstance`.

`modLibrary.ConnectorIndex` and `modConnectors.InstanceIndex`
(`src/vba/modConnectors.bas:135-161`) supply both forms' and the picker's
list-box display strings from a single tested source, per `docs/design/
02_layering_rules.md`'s user-visible-text rule, rather than any form
assembling `"<ID> - <Name>"` text inline.

`docs/design/04_enforcement.md` names the test that holds this invariant:
`test_every_click_handler_delegates` fails if any `cmdXxx_Click` on
`frmConnectorPicker`, `frmManageLibrary`, `frmRemoveConnector`, or
`frmConnectorEditor` does domain work without calling into a layer 1
action module (`modPickerActions`, `modManageActions`, `modEditorActions`,
or - for the ref-des-rename exception above - the allowlisted
`modConnectors.ApplyConnectorEdit`).

**`frmManageLibrary.cmdEdit_Click` is exempted from this check entirely** -
`NON_DELEGATING_HANDLERS` (`tests/test_layering.py:45`) lists it alongside
`cmdCancel_Click`/`cmdClose_Click` because its job is to hand off to
`frmConnectorEditor`, not to make its own domain decision. It reads the
selected connector's fields directly (`modLibrary.ReadConnector`), seeds
the `_Edit` scratch sheet directly (`modPinEditor.LoadScratchPins`), then
loads and shows `frmConnectorEditor` before unloading itself
(`src/vba/forms/frmManageLibrary.evt:39-60`) - both direct layer 0 calls
are allow-listed for this adapter specifically
(`"frmManageLibrary": {"modLibrary.ConnectorIndex", "modLibrary.
ReadConnector", "modLibrary.LIB_ROW_CAP", "modPinEditor.LoadScratchPins",
"modLibrary.FindConnectorRow"}`, `tests/test_layering.py:30-34`), because
handing off pre-loaded state to another form is not itself a transaction
with a pass/fail outcome to wrap in a `modContract` result. `Load`
(not `Show`) is called first so `frmConnectorEditor.LoadForEdit` can
populate the form before it becomes visible - showing the editor first,
then unloading `frmManageLibrary` only afterward, matters: unloading
`frmManageLibrary` before the editor is shown discards everything
`LoadForEdit` had just written.

### `cmdNew_Click`: defining and adding a connector in one flow

`frmConnectorPicker.cmdNew_Click` (`src/vba/forms/frmConnectorPicker.
evt:55-79`) handles defining a brand-new connector and adding it to the
current harness as a single flow. It clears
`modConnectorUI.LastSavedConnectorID` (`src/vba/modConnectorUI.bas:12`) - a
standard module variable, not a property on the form itself - then shows
`frmConnectorEditor` modally. `frmConnectorEditor.cmdSave_Click` sets that
variable on a successful save. Only after `Show` returns, meaning the
editor has already unloaded itself, does `cmdNew_Click` read the variable
back and call `AddFromLibrary` if a connector was actually saved.

This is `docs/design/03_handler_lifecycle.md`'s documented worked
exception: reading `LastSavedConnectorID` *after* the editor form has
unloaded is safe only because the value lives in a standard module. A
property on the form instance would have been reset (or the read would
risk re-triggering `UserForm_Initialize` on the predeclared instance)
by the time `cmdNew_Click` resumes after `Show` returns.
