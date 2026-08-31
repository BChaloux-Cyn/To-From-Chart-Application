# Phase 2b: Connector Editor - Design

Date: 2026-08-26
Status: Reflects merged code as of Phase 2e

## Purpose

The connector editor is where a student defines or edits one connector:
its fields, its photo, and click-to-place pin markers with independently
movable anchor and marker positions. It is a UserForm, `frmConnectorEditor`,
opened from Manage Library (2c).

## Why the form is built in code, not a static `.frm`

`pytest` drives Excel over COM, which cannot open a Visual Basic IDE
project file directly, and a `.frm`/`.frx` pair is binary-ish and painful to
diff. `build/form_layout.py` instead drives the VBIDE `Designer` object
model at build time - `excel_com.add_userform` (`build/excel_com.py`) adds a
`VBEXT_CT_MSFORM` component and returns its `Designer` surface, and
`form_layout.build_connector_editor_form` adds every control (`FIELD_
CONTROLS`, a list of `(progid, name, left, top, width, height, extra_
props)` tuples) programmatically. There is no source file for the form's
visual layout other than this Python module - the same "build from text
source" principle the sheets and VBA modules follow.

## The three-way split

The UI/logic-separation work (`docs/design/01_architecture_overview.md`)
restructured this subsystem's original single-form design into three
pieces. What was originally one form's code-behind is now:

- **`modPinEditor.bas` (layer 0 primitives).** Pure sheet operations on the
  `_Edit` scratch sheet and normalized-coordinate math, with no knowledge of
  forms: `PlacePin`, `RemovePin`, `MoveAnchor`, `MoveMarker`,
  `SnapLabelToPin`, `NeedsLeaderLine`, `MarkerSitsOnAnchor`,
  `FitAspectRatio`, plus two functions added after the original 2b plan -
  `MarkerTopLeft`/`NormFromMarker` (`src/vba/modPinEditor.bas:189-209`,
  exact inverses of each other) and `PinGeometry`
  (`src/vba/modPinEditor.bas:153-164`), all pixel/normalized-coordinate
  conversions the adapter and `clsPinMarker` need but that stay pure enough
  to unit test directly.
- **`modEditorActions.bas` (layer 1 actions).** Every decision a click or
  button represents: the photo guard (`CanLoadPhoto`), the cache source
  decision (`PhotoSourceForEdit`), and the click/place/delete/pin-number
  actions and queries (`PhotoClickAction`, `DeletePinRequest`,
  `SnapLabelRequest`, `ClearAllPins`, `PinListItems`, `NextPinNumber`,
  `SaveFromEditor`), all callable and tested via `Application.Run`
  (`tests/test_editor_actions.py`).
- **`frmConnectorEditor.evt` (layer 2 adapter).** The thin adapter: reads
  controls, calls one `modEditorActions` function, writes controls back.
  This is the only piece of the subsystem manually verified rather than
  automatically tested, because pytest cannot simulate a mouse drag over a
  modal UserForm.

`docs/design/04_enforcement.md`'s `test_adapters_call_only_permitted_
layer0_members` allows `frmConnectorEditor` a short, explicitly-reasoned
list of direct layer 0 calls that have no real transaction to lift:
`FitAspectRatio`, `MarkerTopLeft`, `ClearScratchPins`, `SlugifyConnectorID`,
`ExportShapeToFile` (`tests/test_layering.py:23-27`). `clsPinMarker` is
allowed `NormFromMarker` and `MoveMarker` for the same reason - "the drag
handler converts pixels to a normalized point and stores it. Both are layer
0 primitives with their own tests; there is no transaction here to lift"
(`tests/test_layering.py:35-38`).

## The `_Edit` scratch sheet

In-progress pin edits for whichever connector is currently open live on a
dedicated very-hidden sheet, `_Edit` (code name `shEdit`), rather than an
in-memory structure - the exact same `Worksheet` + `(nFirstRow, nLastRow)`
window pattern `modLibrary` established in 2a, with its own constants
`SCRATCH_FIRST_ROW = 2`, `SCRATCH_LAST_ROW = 2000`
(`src/vba/modPinEditor.bas:4-5`). Cancel (`cmdCancel_Click`,
`src/vba/forms/frmConnectorEditor.evt:275-278`) means calling
`modPinEditor.ClearScratchPins` and never copying `_Edit`'s rows into the
library - there is no separate "discard" code path because the scratch
sheet is the only place an in-progress pin exists at all. `RebuildPinList`
(`src/vba/forms/frmConnectorEditor.evt:95-108`) rebuilds both the list box
and the on-screen markers by re-reading `_Edit` on every call rather than
keeping a parallel in-memory list, which is what let the original design's
`mListPinNumbers`/`mNextPinNumber` counters - which could desync from the
sheet - be retired (`tests/test_editor_actions.py:63-73`,
`test_pin_list_items_survives_a_deleted_middle_pin`).

## The anchor-versus-marker distinction

Every pin carries two independent normalized positions: the anchor
(`NormX`/`NormY`, the cavity on the connector face) and the marker
(`LabelX`/`LabelY`, where the numbered circle is drawn). A fresh placement
sets them identical (`modPinEditor.PlacePin`,
`src/vba/modPinEditor.bas:53-63`). Three gestures touch them differently:

| Gesture | What moves | Handler |
|---|---|---|
| Drag a marker | Marker only; anchor untouched | `clsPinMarker.mLabel_MouseUp` -> `modPinEditor.MoveMarker` |
| Select a pin, click the image (Place Pins off) | Anchor; marker follows only if it was still sitting on the anchor | `imgPhoto_MouseUp` -> `modEditorActions.PhotoClickAction` -> `modPinEditor.MoveAnchor` |
| Snap Label to Pin | Marker snaps back onto the anchor | `cmdSnapLabel_Click` -> `modEditorActions.SnapLabelRequest` -> `modPinEditor.SnapLabelToPin` |

`MoveAnchor` (`src/vba/modPinEditor.bas:91-111`) decides whether the marker
travels with the anchor by checking `MarkerSitsOnAnchor` against the *old*
anchor position before overwriting it - a marker within
`PREVIEW_LEADER_THRESHOLD` (0.01 normalized units,
`src/vba/modPinEditor.bas:12`) of the anchor is considered "on" it. This
threshold is the editor's own live-preview tolerance only, distinct from
the spec's render-time leader rule (based on the rendered oval's actual
pixel radius against final photo size) - that geometry exists only once a
harness is rendered, which this editor never does.

## `clsPinMarker`: one `WithEvents` instance per marker

A VBA form cannot statically declare `WithEvents` for N runtime-created
controls - `WithEvents` only works on a variable declared at compile time,
and the number of markers isn't known until pins are placed. `clsPinMarker`
(`src/vba/clsPinMarker.cls`) is a small wrapper class holding one
`WithEvents mLabel As MSForms.Label`; `AddMarkerControl`
(`src/vba/forms/frmConnectorEditor.evt:170-200`) creates one `clsPinMarker`
instance per placed pin and keeps them in a form-level `Collection`
(`mMarkers`), keyed by pin number. Drag is `MouseDown` (record the grab
offset), `MouseMove` while the button is held (reposition the label),
`MouseUp` (convert the label's final pixel position back to a normalized
point via `NormFromMarker` and commit it with `modPinEditor.MoveMarker`).

## The deliberate limitation: `mConnectorID` is fixed at photo-load time

`mConnectorID` is computed once, in `cmdLoadPhoto_Click`
(`src/vba/forms/frmConnectorEditor.evt:121-138`), from whatever Name and
Part Number are filled in at the moment a photo is loaded
(`modLibrary.SlugifyConnectorID`). Editing Name or Part Number afterward
does not recompute it, so pins already placed keep referencing the original
ID. `modEditorActions.CanLoadPhoto` (`src/vba/modEditorActions.bas:105-111`)
guards the *order* of this - Load Photo is refused with a
`MISSING_NAME_OR_PART` outcome until both fields are non-blank
(`tests/test_editor_actions.py:100-110`) - but does not solve recomputation
after the fact. This mirrors the spec's own precedent for a similar rough
edge (static callout markers discarded on manual harness edits): a
documented limitation given the natural fill-in order, not a bug.

## Photo caching, backfill, and the JPG-only restriction

The original 2b plan's photo picker offered PNG, JPG, and BMP via
`LoadPicture`. The shipped `modEditorActions.PhotoFileFilter`
(`src/vba/modEditorActions.bas:8-10`) offers JPG only:
`LoadPicture`'s legacy OLE loader raises error 481 on some valid PNGs on
certain Windows/Office configurations, even though `Shapes.AddPicture`
(used by `modLibrary.EmbedConnectorPhoto`) handles the same file fine.
Restricting the picker to what `LoadPicture` can actually open is the fix
(`tests/test_editor_actions.py:14-19`).

Editing an existing connector needs the photo back for the preview without
a clipboard round trip on every open. `modEditorActions.PhotoSourceForEdit`
(`src/vba/modEditorActions.bas:117-128`) prefers an on-disk JPG cache
(`modLibrary.CachePhotoPath(..., "jpg")`); when no cache exists yet, it
returns a `NEEDS_BACKFILL` outcome carrying the path the cache *should*
live at, and `LoadExistingPhoto`
(`src/vba/forms/frmConnectorEditor.evt:51-69`) does a one-time backfill
attempt by exporting the embedded `Shape` via `modLibrary.ExportShapeToFile`
- the same clipboard-dependent Shape-Copy/Chart-Paste mechanism documented
in the 2a doc, known to be unreliable for VBA-triggered operations on this
machine. A successful Save always refreshes the cache going forward
(`modEditorActions.PhotoCacheRefreshTarget`,
`src/vba/modEditorActions.bas:20-31`, called from `cmdSave_Click`), so the
backfill path only matters for a connector saved before caching existed.

## Save: ID collision guard and the result envelope

`modEditorActions.SaveFromEditor`
(`src/vba/modEditorActions.bas:218-247`) is the whole save transaction bar
the workbook open/close and the post-save photo-cache copy, which stay in
the adapter per `docs/design/02_layering_rules.md`. It was not in the
original 2b plan's `SaveConnector` signature: the original had no
collision check at all. The shipped version takes both the connector's
current (possibly just-edited) ID and its `sOriginalID` (blank for a new
connector), and rejects the save with `ID_COLLISION` only when the
candidate ID already names a *different* row than the one this session
opened - re-saving the connector currently being edited is never flagged as
a collision against itself (`tests/test_editor_actions.py:286-300`). A save
with no photo path at all fails loudly with `SAVE_FAILED`
(`test_saving_with_no_photo_fails_loudly_rather_than_silently`,
`tests/test_editor_actions.py:303-307`) - the pre-refactor build silently
did nothing in this case, leaving a student with no indication why Save had
no effect.

Every `modEditorActions` action returns `modContract`'s three-element
result envelope (`docs/design/02_layering_rules.md#the-result-envelope`),
and the adapter never assembles its own message text - `MsgBox
modMessages.MessageFor(vResult), modMessages.MessageStyleFor(vResult)`
appears wherever a non-silent outcome needs reporting
(`src/vba/forms/frmConnectorEditor.evt:127-130`, `:258-260`).

## `NextPinNumber`: gap-filling instead of a running counter

The original plan's form tracked `mNextPinNumber`, a simple incrementing
counter. The shipped `modEditorActions.NextPinNumber`
(`src/vba/modEditorActions.bas:80-100`) instead derives the lowest unused
pin number in `1..nPinCount` from `PinListItems` on every placement, so a
deleted pin's number becomes available for reuse rather than a connector's
pins running past its declared `PinCount`
(`tests/test_editor_actions.py:82-94`,
`test_a_deleted_pin_number_is_reused_by_the_next_placement`). This removes
another piece of form-local state that could previously desync from the
sheet.

## Summary of deviations from the 2b plan

| Plan said | Code does | Why |
|---|---|---|
| One form's code-behind does everything | Split into `modPinEditor` (layer 0), `modEditorActions` (layer 1), `frmConnectorEditor.evt` (layer 2) | UI/logic-separation work, so every decision is reachable via `Application.Run` |
| `mListPinNumbers`, `mNextPinNumber` form-level counters | Retired; `PinListItems`/`NextPinNumber` derive state from `_Edit` on every call | Counters could desync from the sheet; a single source of truth cannot |
| Photo picker offers PNG/JPG/BMP | JPG only (`PhotoFileFilter`) | `LoadPicture` error 481 on some valid PNGs; `PhotoFileFilter` only offers formats `LoadPicture` can actually open |
| `SaveConnector` has no ID-collision check | `SaveFromEditor` compares candidate ID against `sOriginalID`, rejecting collisions with a *different* connector only | An edited Name/Part Number can change the derived ID mid-edit; re-saving yourself must not self-collide |
| A save with no photo silently does nothing | Fails loudly with `SAVE_FAILED` | Silent no-op gave a student no indication why Save had no effect |
| No photo cache / backfill flow | `PhotoSourceForEdit` + one-time `ExportShapeToFile` backfill for connectors saved before caching existed | Avoids a clipboard round trip on every edit-open; falls back for older data |
| No `MarkerTopLeft`/`NormFromMarker`/`PinGeometry` | Added to `modPinEditor` | Pixel/normalized-coordinate conversions needed by the adapter and `clsPinMarker`, kept as pure, independently-tested layer 0 functions |
