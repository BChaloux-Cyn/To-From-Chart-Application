# Phase 2a: Connector Library Core - Design

Date: 2026-08-26
Status: Reflects merged code as of Phase 2e

## Purpose

Every connector a student defines lives in one shared, macro-free file,
`dist/ConnectorLibrary.xlsx`, so a part is defined once and reused across
any number of harnesses. `src/vba/modLibrary.bas` is the schema-agnostic
reader/writer that the Creator uses against that file - and, per this doc's
"bounded window" section below, against two other row ranges that share the
same three-table shape without being dedicated whole sheets.

`modLibrary.bas` is a layer 0 primitives module in the layering strategy
`docs/design/01_architecture_overview.md` describes: it operates only on
`Worksheet`, `Range`, and scalar arguments, with no knowledge of forms or
dialogs. That split was added after this subsystem was built; this doc
covers the schema and storage conventions that sit underneath it and are
unaffected by it.

## The three-table schema

`build/library_layout.py:11` defines the sheet order - `Connectors`,
`Pins`, `Photos` - built by `build_library_sheets` and used unmodified by
`dist/ConnectorLibrary.xlsx`.

`Connectors` (`build/library_layout.py:13-16`, columns mirrored in
`src/vba/modLibrary.bas:7-18` as `LIB_COL_ID` through `LIB_COL_ORIGIN`):

| Col | Field |
|---|---|
| 1 | ConnectorID |
| 2 | Name |
| 3 | Manufacturer |
| 4 | PartNumber |
| 5 | Type |
| 6 | PinCount |
| 7 | Notes |
| 8 | PhotoShapeName |
| 9 | CreatedUtc |
| 10 | ModifiedUtc |
| 11 | Origin |

`Pins` (`build/library_layout.py:19`, columns mirrored in
`src/vba/modLibrary.bas:30-37` as `PIN_COL_CONNID` through
`PIN_COL_LABELY`): ConnectorID, PinNumber, PinLabel, NormX, NormY, LabelX,
LabelY.

`Photos` is unstructured - no header row of fields, just embedded picture
shapes placed in a grid (see Photo grid, below).

The same three-sheet layout backs three different physical locations: the
standalone `ConnectorLibrary.xlsx` file, a library file a student imports
(2d), and a fixed-size region on the Creator's own `_Snapshot` sheet (2c).
One reader and one writer serve all three because every `modLibrary`
function takes its target range as an explicit window rather than assuming
a dedicated sheet - see below.

## The bounded-window convention

Every CRUD function in `modLibrary.bas` takes `nFirstRow` and `nLastRow` as
explicit parameters (e.g. `WriteConnector(wsConn, nFirstRow, nLastRow,
vFields)` at `src/vba/modLibrary.bas:77`) rather than scanning "to the
sheet's last used row." This is what lets the same functions serve a
dedicated whole sheet - `ConnectorLibrary.xlsx`, where the window is
generous (`LIB_ROW_CAP = 100000`, `src/vba/modLibrary.bas:23`) - and a
fixed-size region on `_Snapshot`, which holds the Connectors block and the
Pins block on the *same physical sheet* in overlapping row ranges. Scanning
or shifting "to the sheet's bottom" from inside the Connectors block would
read into, or destructively shift, the Pins block below it.

`DeleteConnector` (`src/vba/modLibrary.bas:147-169`) and
`DeletePinsForConnector` (`src/vba/modLibrary.bas:275-304`) both honor this
by never using `Range.Delete Shift:=xlUp`, which would pull rows from below
`nLastRow` upward and corrupt whatever else shares the sheet. Instead,
`DeleteConnector` swaps the last used row's data into the deleted row's
slot and clears the now-duplicate tail row; `DeletePinsForConnector`
compacts in a single pass, copying every non-matching row down to a write
cursor and clearing the leftover tail. Both stay strictly inside
`[nFirstRow, nLastRow]`.

### `LastUsedRowInWindow` and its bug

`src/vba/modLibrary.bas:44-60` is `Public` (not `Private`) because
`modPinEditor` (2b) reuses it for the same bounded-window-safe delete it
needs for "Delete Pin."

The obvious idiom for "last used row in a bounded window" is
`Cells(nLastRow, col).End(xlUp).Row` - a direct adaptation of the
whole-sheet `Cells(Rows.Count, col).End(xlUp)` idiom. That whole-sheet
version works only because `Cells(Rows.Count, col)` is virtually guaranteed
empty. `Cells(nLastRow, col)` has no such guarantee: once a window fills up,
or a caller probes a small window right at existing data, the starting cell
is occupied and `End(xlUp)` walks *up* through the contiguous non-blank run,
overshooting past the real data - in the worst case into a header row above
`nFirstRow`. `LastUsedRowInWindow` checks whether `nLastRow` itself is
occupied first and short-circuits to `nLastRow` in that case, since no
function in this module ever looks past `nLastRow` anyway. `tests/test_
library_connectors.py` and `tests/test_library_pins.py` both include a
boundary test for this (`test_write_respects_the_row_window`, `test_delete_
stays_inside_its_row_window` and their `Pins` counterparts).

## The field-order-array convention

No VBA `Type` crosses the `Application.Run` boundary pytest uses to call
these functions directly - a custom `Type` cannot be marshaled across that
COM boundary. Every record-shaped function instead takes or returns a plain
array in field order: `vFields` for `WriteConnector`/`WritePin` is an
11- or 7-element array matching the column order above, and `ReadConnector`
/`ReadPinsForConnector` return the same shape back. `src/vba/
modLibrary.bas:81` (`WriteConnector`) and `:208` (`WritePin`) both guard
this with a field-count check (`UBound(vFields) - LBound(vFields) + 1 <>
LIB_FIELD_COUNT`) before touching the sheet, rejecting a malformed array
rather than writing a partial row.

## Photo grid and cache path

`Photos` holds one embedded picture shape per connector, named
`PHOTO_<ConnectorID>` and laid out in a fixed grid
(`PHOTO_GRID_COLUMNS = 4`, `PHOTO_GRID_CELL_WIDTH/HEIGHT = 120`,
`PHOTO_GRID_MARGIN = 8`, `src/vba/modLibrary.bas:39-42`).
`EmbedConnectorPhoto` (`src/vba/modLibrary.bas:312-344`) places the next
shape at `(Shapes.Count Mod 4, Shapes.Count \ 4)` scaled by the cell size,
replacing any existing shape for the same ConnectorID first.

`CachePhotoPath` (`src/vba/modLibrary.bas:346-356`) builds
`<folder>\Photos\<ConnectorID>.<ext>`, creating the `Photos\` subfolder if
missing. Its extension defaults to `"png"` but takes an optional override -
`frmConnectorEditor`'s editor-preview cache passes `"jpg"`
(`tests/test_library_photos.py:89-94`) because `LoadPicture`'s legacy OLE
loader unreliably rejects valid PNGs (see the 2b doc), while `modSnapshot`'s
harness-chart cache uses the `"png"` default. This parameter did not exist
in the original 2a plan; it was added once 2b needed a second, differently-
constrained cache consumer.

### A behavior added after the original plan: keep the existing photo on an empty path

The original plan had `EmbedConnectorPhoto` return `""` whenever
`sImagePath` did not exist on disk. Manual verification of 2c
(`docs/superpowers/plans/phase-2-manual-verification.md`) found this broke
editing an existing connector's fields or pins *without* picking a new
photo: the editor's re-export of the already-embedded photo
(`ExportShapeToFile`, below) is itself unreliable, so `mPhotoPath` could
come back blank, which made `EmbedConnectorPhoto` - and so the entire save -
fail outright even though the connector already had a perfectly good photo
on file.

`EmbedConnectorPhoto` now distinguishes "no path supplied, but a photo is
already embedded for this ID" from "no path supplied, and none exists yet."
The former keeps the existing shape and returns its name; only the latter
returns `""` (`src/vba/modLibrary.bas:319-331`, tested by
`test_embed_photo_with_no_new_path_keeps_the_existing_one` in `tests/
test_library_photos.py:26-42`).

## `ConnectorIndex`: the browsable list, rendered here

`ConnectorIndex(wsConn)` (`src/vba/modLibrary.bas:119-145`) was not part of
the original 2a plan - it was added so the picker and manage-library forms
(2c) have a single tested source for their list-box display strings rather
than assembling `"<ID> - <Name>"` text in a form. It returns a 2D array,
one row per connector, column 1 the display string and column 2 the bare
ConnectorID that string resolves to; an empty sheet returns `Empty`, never
a zero-length array (`test_connector_index_of_an_empty_sheet_is_empty`,
`tests/test_library_connectors.py:88-89`). Unlike the CRUD functions above,
it scans the whole sheet (`Cells(Rows.Count, ...).End(xlUp)`,
`src/vba/modLibrary.bas:123`) rather than taking a row window, because it
is only ever called against the dedicated whole-sheet library, never
against a shared sheet like `_Snapshot`.

## `ExportShapeToFile`: exporting a shape from a very-hidden sheet

Also absent from the original 2a plan. Excel has no direct "export a Shape
to an image file" call; `ExportShapeToFile`
(`src/vba/modLibrary.bas:358-394`) uses the standard workaround of pasting
the shape into a throwaway `ChartObject` on its own sheet and exporting the
chart. Two complications this function handles that a naive Copy/Paste
would not:

- `Chart.Paste`'s clipboard target must be the sheet's `ActiveSheet`, but a
  very hidden sheet (`_Edit`, `_Snapshot`) can never become active. The
  function temporarily sets the host sheet visible and activates it for the
  duration of the export, restoring its prior visibility afterward
  (`src/vba/modLibrary.bas:370-380`, `:391`).
- A `Shape` reference obtained while its sheet was still very hidden copies
  as empty even after the sheet is unhidden and activated; the shape must
  be re-fetched by name once the sheet is active (`src/vba/
  modLibrary.bas:378-380`).
- `Chart.Paste` can silently paste nothing - no error, no other indication -
  leaving a blank exported image, since `Chart.Export` writes a file either
  way. `ExportShapeToFile` checks `cht.Chart.Shapes.Count` before trusting
  the paste and reports failure via its `Boolean` return rather than
  trusting that a file landed on disk (`src/vba/modLibrary.bas:386-388`,
  `:393`).

Manual verification (2b/2c) found `Chart.Paste` can silently paste nothing
in this environment regardless of automation versus a real interactive
session - the same conclusion the 2d doc reaches independently for
`CopyConnectorPhoto`'s sibling clipboard mechanism. `tests/
test_library_photos.py:97-121` tests this structurally (asserting the
source contains the `Chart.Shapes.Count` check and the hidden-sheet
activation dance) rather than asserting a live Copy/Paste actually lands,
since a live assertion would not be any more reliable in CI than it was by
hand.

## Summary of deviations from the 2a plan

| Plan said | Code does | Why |
|---|---|---|
| `EmbedConnectorPhoto` returns `""` for any missing `sImagePath` | Returns the existing shape name if one is already embedded for that ConnectorID | An empty path during an edit-without-new-photo save must not fail the whole save (found in 2c manual verification) |
| `CachePhotoPath(sWorkbookFolder, sConnectorID) As String` | Takes an optional third `sExtension` parameter, default `"png"` | 2b's editor-preview cache needs `"jpg"`; `modSnapshot`'s cache keeps `"png"` |
| No `ConnectorIndex` function | Added, whole-sheet scan, returns display-string/ID pairs | Single tested source for picker/manage-library list rendering (2c), keeping display text out of forms per `docs/design/02_layering_rules.md`'s user-visible-text rule |
| No `ExportShapeToFile` function | Added, Shape-Copy/Chart-Paste/Chart-Export with hidden-sheet handling | Needed to re-export an already-embedded photo from `_Edit`/`_Snapshot` when a save doesn't supply a new photo path |
