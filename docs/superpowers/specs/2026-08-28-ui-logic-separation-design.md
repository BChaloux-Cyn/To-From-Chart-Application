# UI and Logic Separation - Design

Date: 2026-08-28
Status: Approved for planning

## Purpose

Make the behaviour currently locked inside UserForm and worksheet event
handlers reachable from the pytest suite, so that a green run implies the
application works rather than implying the source text still contains
certain strings.

## The problem

683 of roughly 1690 lines of VBA live in `.evt` files. `Application.Run`
reaches Public procedures in standard modules only; a form's or a sheet's
document module is unreachable by design. Those handlers also call `MsgBox`,
`InputBox`, `GetOpenFilename` and `.Show`, each of which blocks a headless
automation run regardless of reachability.

Because the behaviour cannot be called, `tests/test_connector_editor_wiring.py`
and `tests/test_manage_library_transfer_wiring.py` assert on VBA source text
instead. Those assertions are tautological. `test_place_pins_is_capped_at_pin_count`
asserts that the string `lstPins.ListCount >= nPinCount` appears in the
module; renaming a variable fails a correct implementation, and inverting the
comparison passes a broken one.

The decision-making is what needs to move. The control manipulation can stay
where it is.

## Architecture: three layers

| Layer | Modules | May reference | May not reference |
|---|---|---|---|
| 0 - primitives | `modUtil`, `modState`, `modLibrary`, `modChart`, `modConnectors`, `modSnapshot`, `modLibraryTransfer`, `modPinEditor` | `Worksheet`, `Range`, scalars | anything above |
| 1 - actions | `modEditorActions`, `modPickerActions`, `modManageActions`, `modContract`, `modMessages` | layer 0, each other, `Worksheet`, `Range`, `Workbook`, scalars | `MSForms.*`, `MsgBox`, `InputBox`, `GetOpenFilename`, `GetSaveAsFilename`, `.Show`, `Workbooks.Open` |
| 2 - adapters | `*.evt` for forms and sheets, `modConnectorUI` launchers | layer 1 only | layer 0 |

### Layer rules

An adapter handler does three things: read controls, call one layer 1
function, write controls. It may branch on the returned outcome to choose
what to display. It may not branch on domain state to decide what the
operation is.

A file path is an argument to a layer 1 function, never something the
function goes and asks the user for. Filesystem reads and writes are
permitted in layer 1; dialogs are not.

Workbook lifecycle stays in layer 2. The handler performs
`Workbooks.Open(ThisWorkbook.Path & "\ConnectorLibrary.xlsx")` and passes the
resulting `Worksheet` objects down, matching how `modLibrary` and
`modSnapshot` already take sheets as parameters. This is what lets a test
substitute its own `library_wb` fixture.

### What stays in an adapter

Control property assignment, dialog invocation, workbook open and close,
`Unload`, and event-lifecycle state such as `shConnectors.mLastRefDes`.

### Handler lifecycle discipline

An action must not have its inputs destroyed while it runs. Two properties of
VBA make that safe, and both need protecting.

VBA is single threaded. While a handler executes, no other form event can
fire, including `UserForm_QueryClose` from the system Close button. The form
therefore cannot unload underneath a running action. The single hole in that
guarantee is `DoEvents`, which re-enters the message pump and allows a queued
Close click to fire `QueryClose` mid-action - which in `frmConnectorPicker`
and `frmManageLibrary` closes `mLibrary` and invalidates the very `Worksheet`
objects the action is holding.

Arguments are fully evaluated before the call. `Array(Trim$(txtName.Text), ...)`
is materialized into a Variant array before the action receives it, and
Strings and Variants are copied, so an action holds values rather than live
references into the form.

The layering protects the action's side for free: layer 1 cannot reference
`MSForms.*` and cannot call `.Show` or `Unload`, so an action structurally
cannot unload the form that called it. Ordering within the handler is not
covered by that, so three rules apply:

1. **Capture then act.** A handler reads every control value and form-level
   variable it needs into locals before the action call. Nothing after the
   call touches the form.
2. **`Unload Me` is the final statement in its branch.** No line after it may
   reference a control or a form-level variable.
3. **No `DoEvents`** in layer 1, and none in an adapter between capturing
   state and unloading.

All three hold in the current code. They are stated here because nothing
enforces them, and the failure mode is silent data loss rather than an error.

`frmConnectorPicker.cmdNew_Click` reads `modConnectorUI.LastSavedConnectorID`
after `frmConnectorEditor` has already unloaded itself. That works only
because the value lives in a standard module; form state would have been
gone. `src/vba/modConnectorUI.bas:8` records that this was the fix for
exactly this failure. The mechanism is kept unchanged.

## The result contract

### Actions and queries

Layer 1 holds two kinds of function, and only one of them is bound by the
envelope contract.

An **action** either mutates state or renders a pass/fail judgement. It
returns the three element result described below. `SaveFromEditor`,
`PhotoClickAction`, `DeletePinRequest`, `CanLoadPhoto`, `AddFromLibrary`,
`DeleteFromLibrary`, `ExportToWorkbook`, `ImportAllFromWorkbook`,
`AttachReplacementPhoto`, `PhotoSourceForEdit`, `ApplyHarnessEdit` and
`ApplyConnectorEdit` are actions.

A **query** is a pure lookup with no failure mode worth naming. It returns a
bare typed value. `PinListItems`, `TypeListItems`, `PhotoFileFilter`,
`PhotoCacheRefreshTarget`, `MarkerControlName`, `MessageFor` and
`MessageStyleFor` are queries. A query that finds nothing returns `Empty` or
an empty string, never a failure envelope.

Layer 0 primitives are bound by neither convention. `FitAspectRatio` and
`PinGeometry` keep their existing bare `Array` returns, as do the two new
coordinate conversions.

Tests call actions through the `run_action` helper, which validates the
envelope, and queries through the plain `run` helper.

### The envelope

Every layer 1 action returns a three element result. The contract is
enforced structurally, not by convention, because the codebase already runs
two incompatible array conventions side by side:
`modLibrary.ReadPinsForConnector` returns `ReDim vResult(1 To n, 1 To 7)`,
which re-bases from 1 to 0 when it crosses `Application.Run`, while
`modPinEditor.PinGeometry` returns an `Array(...)`, which does not. A test
asserting `result[1]` can therefore pass while a form reading `vResult(1)`
receives a different element. A comment at
`src/vba/forms/frmConnectorPicker.evt:31` already documents this hazard.

### One constructor, one set of accessors

```vba
' modContract.bas
Public Function Success(ByVal sOutcome As String, Optional ByVal vPayload As Variant) As Variant
Public Function Failure(ByVal sOutcome As String, Optional ByVal vPayload As Variant) As Variant
Public Function Ok(vResult As Variant) As Boolean
Public Function Outcome(vResult As Variant) As String
Public Function Payload(vResult As Variant) As Variant
Public Function TableRowCount(vPayload As Variant) As Long
Public Function PayloadKind(ByVal sOutcome As String) As String
Public Function OutcomeCodes() As Variant
```

Action functions never build a result by hand; they return
`modContract.Success("PLACED", nPinNumber)`. Adapters never index one by
hand; they call `modContract.Ok(vResult)`. The literals `(0)`, `(1)` and
`(2)` appear exactly once each in the codebase, inside `modContract`, which
is tested directly.

`Success` and `Failure` build results with the `Array` function, which is
zero based for both an in-process VBA caller and a COM caller, so
`vResult(0)` in VBA and `result[0]` in pytest are the same element. No
`Option Base` directive may be added anywhere; it would break this property.

### Payload kinds

`PayloadKind` declares the payload type for each outcome as one of `NONE`,
`STRING`, `LONG`, `DOUBLE` or `TABLE`. `OutcomeCodes` returns the full
registry. `Success` and `Failure` validate the payload against the declared
kind and raise on mismatch, and raise on an outcome code absent from the
registry.

This converts the worst class of UI failure - an adapter calling
`CStr(vResult(2))` on what turned out to be an array - into a loud failure at
the point of construction.

### Table shape

Any table crossing the seam - a `TABLE` payload on an action, or the bare
return of a query such as `PinListItems` - is either `Empty` or a two
dimensional array with `LBound` of 1 on both dimensions. It is never a zero
length array. Adapters call `modContract.TableRowCount` rather than `LBound`
or `UBound` directly, which removes the "iterate a zero row array" crash
class.

This matches what `modLibrary.ReadPinsForConnector` already does, so existing
callers need no change.

### User-visible text

All text a student reads is produced in layer 1.

```vba
' modMessages.bas
Public Function MessageFor(vResult As Variant) As String
Public Function MessageStyleFor(vResult As Variant) As Long   ' vbExclamation | vbInformation
```

A handler becomes
`MsgBox modMessages.MessageFor(vResult), modMessages.MessageStyleFor(vResult)`.
No formatting logic remains in any `.evt` file, and a test asserts the
literal string including any interpolated connector ID.

The same principle applies to lists. `PinListItems` and `ConnectorIndex`
return render-ready display strings, so `J1 - Deutsch DTM 4-way` is asserted
in a test rather than assembled in a form.

Layer 1 therefore knows English wording. That is a string table, not a UI
dependency; it references no MSForms type.

### Parameter types at the seam

A layer 1 parameter declares the type the control actually supplies, so the
parsing and validation sit inside the tested function where the form's real
input reaches them.

| Control source | Real type | Failure if a test passes something else |
|---|---|---|
| `txtPinCount.Text` | `String`, possibly empty, `" 4 "`, `"4.5"` | passing `4` skips the `IsNumeric` guard entirely |
| `imgPhoto_MouseUp` X, Y | `Single` | Python floats are `Double`; values diverge around the 8th decimal |
| `lstPins.ListIndex` | `Long`, `-1` when nothing is selected | passing `0` never exercises the no-selection path |
| cell `.Value` | `Variant`, usually `Double` | passing `1` skips a `CLng` coercion |

`PhotoClickAction` accordingly declares `sPinCountText As String`, not
`nPinCount As Long`.

Because `X / imgPhoto.Width` is `Single` arithmetic widened to `Double`,
tests compare normalized coordinates with `pytest.approx`, and any new
comparison on normalized coordinates uses a tolerance.
`modPinEditor.MarkerSitsOnAnchor` already does this and is the pattern to
follow.

## Extraction inventory

### modEditorActions (new)

From `frmConnectorEditor` and `clsPinMarker`.

| Function | Kind | Replaces | Returns |
|---|---|---|---|
| `PhotoClickAction(wsScratch, sConnectorID, bPlaceMode, nSelectedPin, sPinCountText, dNormX, dNormY)` | action | all of `imgPhoto_MouseUp` | `PLACED`, `MOVED_ANCHOR`, `BAD_PIN_COUNT`, `PIN_LIMIT_REACHED`, `NO_OP` |
| `SaveFromEditor(wsLibConn, wsLibPins, wsLibPhotos, wsScratch, sConnectorID, sOriginalID, vFields, sPhotoPath, sNowUtc)` | action | the collision check and save in `cmdSave_Click` | `SAVED`, `ID_COLLISION`, `SAVE_FAILED` |
| `PhotoSourceForEdit(sWorkbookPath, sConnectorID)` | action | `LoadExistingPhoto`'s cache and backfill decision | `CACHE_READY`, `NEEDS_BACKFILL`, both with the cache path as a `STRING` payload |
| `CanLoadPhoto(sName, sPartNumber)` | action | the `cmdLoadPhoto_Click` guard | `OK`, `MISSING_NAME_OR_PART` |
| `DeletePinRequest(wsScratch, sConnectorID, nPinNumber)` | action | `cmdDeletePin_Click` | `DELETED`, `NOT_FOUND` |
| `PhotoCacheRefreshTarget(sWorkbookPath, sConnectorID, sPhotoPath)` | query | the FileCopy-onto-itself guard | destination path, or an empty string when no copy is needed |
| `PinListItems(wsScratch, sConnectorID)` | query | `lstPins` plus `mListPinNumbers` | table of display, pin number, labelX, labelY |
| `TypeListItems(wsLists)` | query | the `UserForm_Initialize` combo population | table of type names |
| `PhotoFileFilter()` | query | the `GetOpenFilename` filter string | string |
| `MarkerControlName(nPinNumber)` | query | the `lblMarker` naming convention | string |

`PhotoClickAction` derives both the placed count and the next pin number from
the `_Edit` sheet rather than accepting them as arguments. That is what
retires `mNextPinNumber` alongside `mListPinNumbers` and makes the scratch
sheet the single source of truth.

`TypeListItems` and `PhotoFileFilter` exist purely to move a hard-won fact
out of an unreachable handler and into an assertable return value: the
`_Lists` sheet must resolve against `ThisWorkbook`, and `LoadPicture` rejects
valid PNGs with error 481 on some Office builds.

### modPinEditor (existing, layer 0)

Gains the two pixel and normalized coordinate conversions currently inline in
`AddMarkerControl`, `RepositionMarkerControl` and
`clsPinMarker.mLabel_MouseUp`:

```vba
MarkerTopLeft(dNormX, dNormY, dPhotoLeft, dPhotoTop, dPhotoW, dPhotoH, dMarkerW, dMarkerH) ' -> Array(left, top)
NormFromMarker(dLeft, dTop, dMarkerW, dMarkerH, dPhotoLeft, dPhotoTop, dPhotoW, dPhotoH)   ' -> Array(normX, normY)
```

They are exact inverses, so a round-trip test pins down the half-width
centering that has caused defects before.

### modLibrary (existing, layer 0)

Gains `ConnectorIndex(wsConn)`, returning display string and ConnectorID per
row. This is where `modConnectorUI.RefreshConnectorList` goes; it is a
primitive read of the Connectors sheet and belongs beside `ReadConnector`
rather than in either form's action module.

### modPickerActions (new)

`AddFromLibrary(wsSnapshot, wsLibConn, wsLibPins, wsLibPhotos, sConnectorID)`
performs `ReadConnector`, `AddConnectorInstance` and `SnapshotConnector` as
one transaction. Outcomes `ADDED` with the ref des as payload, `NOT_FOUND`,
`ADD_FAILED`. Both `cmdAdd_Click` and `cmdNew_Click` call it, which collapses
the duplication currently present between them.

### modManageActions (new)

| Function | Outcomes |
|---|---|
| `DeleteFromLibrary(wsLibConn, wsLibPins, wsLibPhotos, sWorkbookPath, sConnectorID)` | `DELETED` - the three sheet deletes plus the orphaned cache file removal |
| `ExportToWorkbook(wsLibConn, wsLibPins, wsLibPhotos, destWb, sConnectorID)` | `EXPORTED`, `EXPORT_FAILED`; `BuildExportSheets` moves inside |
| `ImportAllFromWorkbook(srcWb, wsLibConn, wsLibPins, wsLibPhotos)` | `IMPORTED`, `TABLE` payload of destination ID and photo-ok flag per row |
| `AttachReplacementPhoto(wsLibPhotos, sDestID, sPath)` | `ATTACHED`, `FAILED` |

`ImportAllFromWorkbook` is the largest single gain: 50 lines of loop,
ID-collision renaming and photo-failure detection become one call, leaving
only the replacement-photo file picker in the handler, invoked for the rows
whose flag came back false.

### Sheet modules

`Range` is not a UI type and marshals through `Application.Run`, so these
extract wholesale.

`modChart.ApplyHarnessEdit(wsHarness, rTarget)` takes the bulk versus
per-cell dispatch, row clamping, validation rebuilds and title block units.
The handler keeps only the `EnableEvents` guard and the error trap.

`modConnectors.ApplyConnectorEdit(rTarget, sPriorRefDes, nPriorRow)` returns
`RENAMED`, `RENAME_REJECTED` with the value to restore as payload, or
`NO_RENAME`. `mLastRefDes` and `mLastRefDesRow` stay in the sheet module as
genuine event-lifecycle state and are passed in.

### Deletions

- `modConnectorUI.RefreshConnectorList`, replaced by
  `modLibrary.ConnectorIndex`. `modConnectorUI` retains only the `ShowXxx`
  launchers and `LastSavedConnectorID`, all correctly layer 2.
- `modState.IsTestMode`, which is dead code: defined and built into the
  `_State` sheet, never read.
- The grep assertions listed under Testing below.

### Explicitly out of scope

`modConnectors.AddConnectorInstance` resolves its target sheet through
`ThisWorkbook` rather than accepting a `Worksheet`. Changing that would be
more testable but ripples through passing tests for no gain to this goal.

## Data flow example: Save

```vba
Private Sub cmdSave_Click()
    Dim lib As Workbook
    Set lib = Workbooks.Open(ThisWorkbook.Path & "\ConnectorLibrary.xlsx")

    Dim vResult As Variant
    vResult = modEditorActions.SaveFromEditor( _
        lib.Worksheets("Connectors"), lib.Worksheets("Pins"), lib.Worksheets("Photos"), _
        ThisWorkbook.Worksheets("_Edit"), mConnectorID, mOriginalConnectorID, _
        Array(Trim$(txtName.Text), Trim$(txtManufacturer.Text), Trim$(txtPartNumber.Text), _
              cboType.Text, Trim$(txtPinCount.Text), Trim$(txtNotes.Text)), _
        mPhotoPath, Format$(Now, "yyyy-mm-ddThh:mm:ssZ"))

    lib.Close SaveChanges:=modContract.Ok(vResult)
    If Not modContract.Ok(vResult) Then
        MsgBox modMessages.MessageFor(vResult), modMessages.MessageStyleFor(vResult)
        Exit Sub
    End If

    Dim sCopyTo As String
    sCopyTo = modEditorActions.PhotoCacheRefreshTarget(ThisWorkbook.Path, mConnectorID, mPhotoPath)
    If Len(sCopyTo) > 0 Then FileCopy mPhotoPath, sCopyTo

    modConnectorUI.LastSavedConnectorID = mConnectorID
    Unload Me
End Sub
```

The ID collision rule, the timestamp, the FileCopy guard and the message text
are all now reachable from pytest. What remains is a `Workbooks.Open`, a
`MsgBox`, a `FileCopy` and an `Unload`.

## Error handling

Expected failure is an outcome code, never a VBA error. `ID_COLLISION`,
`BAD_PIN_COUNT` and `EXPORT_FAILED` are ordinary return values.

Unexpected failure propagates. Layer 1 gets no blanket `On Error Resume
Next`; an escaping error surfaces as a `com_error` in tests and reaches the
handler's existing trap in production. This matters because most defects
recorded in `docs/superpowers/plans/phase-2-manual-verification.md` were
silent: clicking to place a pin did nothing when `mConnectorID` was blank,
and a duplicate part number silently overwrote another connector's row. A
function obliged to return an outcome cannot fail silently; the worst it can
do is return an outcome the handler does not recognise, which is a visible
message rather than nothing.

The narrow existing `On Error Resume Next` uses that probe for a Shape's
existence stay. They test for absence rather than swallowing failure.

## Testing

### Behavioural replacements

| Retired assertion | Replaced by |
|---|---|
| `test_place_pins_is_capped_at_pin_count` | `PhotoClickAction` returns `PIN_LIMIT_REACHED` once the placed count reaches `sPinCountText` |
| `test_save_rejects_a_part_number_that_collides...` | `SaveFromEditor` returns `ID_COLLISION`; a same-ID re-save returns `SAVED` |
| `test_load_photo_requires_name_and_part_number_first` | `CanLoadPhoto` over blank/blank, name only, part only, both |
| `test_load_existing_photo_prefers_the_disk_cache...` | `PhotoSourceForEdit` reports no backfill with a cache file present, backfill without |
| `test_save_refreshes_the_photo_cache_via_plain_file_copy` | `PhotoCacheRefreshTarget` returns empty when source is already the cache, dest path otherwise |
| `test_form_tracks_pin_numbers_independently_of_list_position` | `PinListItems` returns pin 3 at row 2 after pin 2 is deleted; `mListPinNumbers` no longer exists to desync |
| `test_type_combo_is_populated_from_this_workbooks_lists_sheet` | `TypeListItems` |
| `test_load_photo_filter_excludes_png` | `PhotoFileFilter` |
| `test_pin_marker_calls_move_marker_on_drop` | `NormFromMarker` and `MarkerTopLeft` round trip |
| the three export, import and replacement-prompt greps | `ExportToWorkbook`, `ImportAllFromWorkbook` asserting a false photo flag for an unextractable photo, `AttachReplacementPhoto` |

### Structural lints

The remaining eight grep assertions are all of one kind - the adapter is
wired to the right thing - so five parametrized tests enforce the layering
and lifecycle rules generically and cover any form added later without
modification.

```python
test_evt_files_call_no_layer0_module
test_action_modules_open_no_dialogs
test_every_click_handler_delegates
test_no_doevents_anywhere
test_nothing_follows_unload_me      # no control or form-variable reference after Unload Me in a Sub
```

The last two enforce the handler lifecycle discipline above. Both are
implementable with the `module_source` helper the existing wiring tests
already use.

These are still source-text tests. They are not the per-case wiring greps
being deleted: those assert on logic, which is what makes them tautological.
These assert on architecture, for which source text is the correct
instrument.

### Retained assertions

`test_photo_fit_box_is_a_fixed_constant` and
`test_pin_marker_uses_a_small_fixed_badge_size` stay. Neither is replaced by
anything. They guard against passing `imgPhoto.Width` as the fit box, an
argument mistake no `FitAspectRatio` test can catch, and against a marker
reverting to the default control size. Both concern irreducibly
presentational code.

### Envelope validation

Every action call in the suite passes through one helper, so no test can
pass against a malformed return.

```python
def run_action(wb, macro, *args) -> Result:
    raw = run(wb, macro, *args)
    assert isinstance(raw, (tuple, list)) and len(raw) == 3, f"{macro} returned {raw!r}"
    ok, outcome, payload = raw
    assert isinstance(ok, bool) and isinstance(outcome, str) and outcome
    assert outcome in KNOWN_OUTCOMES
    return Result(ok, outcome, payload)
```

`KNOWN_OUTCOMES` is read once from `modContract.OutcomeCodes()` rather than
mirrored in Python, so it cannot drift.

### Test files

`test_contract.py`, `test_messages.py`, `test_editor_actions.py`,
`test_picker_actions.py`, `test_manage_actions.py`, `test_marker_geometry.py`,
`test_sheet_event_logic.py`, `test_layering.py`. Roughly 55 to 65 tests
replacing 20, reusing the existing `wb`, `library_wb` and `tmp_path`
fixtures.

### Order of work

Strict TDD, one function at a time: write the behavioural test against the
new signature so it fails because the function does not exist, extract the
logic unchanged, watch it pass, then delete the grep it retires. There are
no behavioural tests over this code today, so each extraction has no safety
net beyond the test that motivates it.

## Limitations

Nothing here drives a UserForm. `Application.Run` cannot click a button, and
a modal `.Show` blocks headless automation. Every decision becomes testable;
what stays manual is that controls are wired to those decisions and render
correctly.

The phase therefore ends with a manual verification pass in the style of
`docs/superpowers/plans/phase-2-manual-verification.md`, but a much shorter
one, covering wiring and appearance rather than re-checking every rule.

## Build changes

Five module names are added to `VBA_MODULES` in `build/build.py`:
`modContract.bas`, `modMessages.bas`, `modEditorActions.bas`,
`modPickerActions.bas`, `modManageActions.bas`. VBA resolves standard module
references at compile time, so import order does not matter.

## Considered and rejected

**Screen automation (pywinauto, AutoIt, AutoHotkey).** MSForms 2.0 controls
are windowless; they have no HWND of their own, and
`AccessibleObjectFromWindow` cannot retrieve an `IDispatch` pointer from a
UserForm's window class. The Windows 8 interfaces that let a windowless
ActiveX control expose itself to UI Automation require the control to
implement them, which MSForms 2.0 never did. The remaining fallback is
generating mouse events at hardcoded screen coordinates, which is more
brittle than the grep tests being deleted, breaks on DPI and layout changes,
and requires an interactive desktop session.

**A TestMode flag gating the dialogs.** `modState.IsTestMode` already exists
unused and is the seed of this approach. It does not help: `Application.Run`
still cannot reach a form's private handler, so the logic remains uncallable
whether or not its dialogs are suppressed.

**Folding the extracted logic into the existing domain modules.** Avoids new
files, but `modLibrary` is already 362 lines and would clear 500, mixing
cell-level primitives with user-intent transactions in one file. That
altitude mixing is a large part of what makes `modLibrary` hard to read now.

**A single `modCommands` module above the primitives.** The cleanest rule to
state, but one module holding every transaction from all three forms lands
near 400 lines with no internal cohesion.

## Deferred

**`Me.Hide` in place of `Unload Me`.** The canonical VBA way to return a
value from a modal form: the instance and its state survive `Hide`, the modal
`.Show` returns to the caller, and the caller reads properties off the
still-live form before unloading it. That would allow deleting the
`modConnectorUI.LastSavedConnectorID` global entirely. Not done in this
phase: this form's Load, Show and Unload lifecycle has already produced three
separate defects recorded in the source comments, the global demonstrably
works, and changing it during a testability refactor combines two risky
changes.

**A `CallByName` form driver.** VBA's `CallByName` can invoke a control event
handler by name if the handler is declared `Public`, and `Load` instantiates
a form without showing it. A standard-module shim iterating `VBA.UserForms`
would give pytest a reachable path to `cmdSave_Click` or
`imgPhoto_MouseUp(1, 0, 45.0, 60.0)`, closing most of the manual gap above.
It only becomes usable once the dialogs are out of the handlers, which this
design achieves, so it is a natural follow-on. Two things need verifying
first: whether `CallByName` reaches a control event handler as opposed to an
ordinary Public Sub on the form, and whether a loaded, unshown form survives
across separate `Application.Run` calls.
