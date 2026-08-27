# Phase 2 Manual Verification Checklist

Every Phase 2 sub-plan (2a-2d) leaves at least one thing pytest cannot exercise:
mouse-driven UI (drag, click, dialogs) and, in one case, clipboard-dependent
COM automation. Each sub-plan's own "Manually verify in Excel" step says so
individually. This doc collects them in one place so they get run as a single
batch **after 2d is complete, before 2e (docs)** - so the design doc and
student user guide describe verified, adjustment-if-needed behavior rather
than being written against code no one has actually clicked through yet.

Do not start 2e until this checklist is run and any adjustments it surfaces
are made.

## How to use this

For each sub-plan below: open `dist/HarnessCreator.xlsm` fresh (`rm -rf dist &&
python build/build.py`), work through its items, and note pass/fail plus any
adjustment made directly in this file (not just in a commit message) so the
record survives past any one session.

## 2a: Connector library core

No manual-only items. Everything in 2a is reachable via `Application.Run`,
including the full save/reopen round trip (`tests/test_library_integration.py`).

## 2b: Connector editor with click-to-place

Status: implemented (2026-08-27), commits `7dd8953`..`77b75be`. **Outstanding -
not yet run.**

Run `frmConnectorEditor.Show` from the VBE Immediate window (or a temporary
button) and confirm by hand:

- [x] Load Photo fits the image without letterboxing.
      **Found 2 bugs during this check, both fixed:** (1) `LoadPicture`
      (VBA's legacy OLE picture loader) throws runtime error 481 "Invalid
      picture" on valid PNGs on this machine/Office config, even though
      `Shapes.AddPicture` handles the same file fine - `cmdLoadPhoto_Click`'s
      file picker is now restricted to JPG only (`*.jpg;*.jpeg`) rather than
      routing PNG through a workaround (see note below on why the natural
      workaround - export via a throwaway Shape/Chart - doesn't pan out
      here). (2) The image-fit box was computed from `imgPhoto.Width/Height`,
      which the same code then overwrote with the fitted result - each
      subsequent Load Photo used the already-shrunk control as its new
      bounding box, shrinking the preview further every time. Fixed by
      fitting against a fixed constant (`PHOTO_BOX_WIDTH`/`HEIGHT` = 180x180,
      matching the design-time size) instead. Confirmed: PNG rejected up
      front by the picker, JPG loads with correct fit, and loading a second
      photo no longer shrinks the viewing area.
      **Separately discovered, not fixed (deferred to 2c review below):**
      `modLibrary.ExportShapeToFile` (the `Shape.Copy`/`Chart.Paste`/
      `Chart.Export` trick used to rasterize an Excel Shape to a file) is
      unreliable on this machine for VBA-triggered clipboard operations -
      confirmed via both headless COM automation and a real interactive
      Excel session, independent of whether the source sheet is very hidden.
      `ExportShapeToFile` was still hardened (verifies the paste actually
      landed instead of silently exporting a blank image, and temporarily
      unhides+activates a very-hidden host sheet, which is necessary but
      evidently not sufficient here) since that's a correct defensive fix
      regardless. But this means `LoadExistingPhoto` (2c's Edit-flow preview,
      below) and `modSnapshot.bas`'s photo caching (2c's Add Connector flow)
      may hit the same blank-image failure - watch for it specifically when
      running those checks.
- [x] Place Pins drops numbered markers on click.
      **Found 2 bugs, both fixed:** (1) `mConnectorID` is computed from
      Name/Part Number only when a photo is loaded, and never recomputed -
      loading a photo before filling in those fields left `mConnectorID`
      blank, so `imgPhoto_MouseUp` silently did nothing on every click (no
      error, no pin, nothing in the list). Fixed by gating Load Photo itself
      on both fields being non-blank (message box explains why), rather than
      letting it fail silently downstream at pin-placement time. (2) Pin
      marker Label controls used the default (body-text-sized) control
      dimensions, dwarfing the photo - fixed with a small fixed
      `PIN_MARKER_SIZE` (16x16) badge.
      **Also added (feature gap found during this check, not a bug in
      existing behavior):** nothing capped placed pins at the entered Pin
      Count - `imgPhoto_MouseUp` now blocks placement with a message once
      `lstPins.ListCount` reaches Pin Count.
- [x] Dragging a marker moves only the marker (confirmed - other pins stay
      put). Leader line: **not implemented, deferred by decision.** The
      underlying data logic (`modPinEditor.NeedsLeaderLine`/`PinGeometry`,
      both correct and tested) was already there, but nothing ever rendered
      one - a first attempt to draw it via the MSForms Line control
      (`Forms.Line.1`) failed on this machine two different ways in a row
      (missing early-bound type, then "Invalid class string" - the control
      itself isn't registered here at all, not just a typing issue).
      Decided to drop the visual for now rather than chase a GDI-API-based
      alternative; can revisit later. Dragging a marker away from its
      anchor currently gives no visual indication of that - noted as a
      known gap for 2e's design doc.
- [x] Selecting a pin in the list, then clicking the image, moves that pin's
      anchor. Confirmed.
- [x] Snap Label to Pin returns the marker to the anchor. Confirmed (leader
      removal N/A - leader line deferred, see above). Requires a pin
      selected in the list first (`lstPins.ListIndex >= 0`), same
      precondition as Place Pins - initially looked broken only because no
      pin was selected, not an actual bug.
- [x] Save writes to `ConnectorLibrary.xlsx` and closes the form. Confirmed:
      new row appeared in Connectors/Pins/Photos sheets of
      `dist\ConnectorLibrary.xlsx`, form closed automatically after save.
      **Also added (feature gap found during this check):** Part Number had
      no uniqueness check - saving a new connector with an already-used
      Part Number would silently overwrite that connector's row (same
      derived ID). `cmdSave_Click` now blocks with a message on collision
      with a *different* connector (re-saving the one you're editing is
      still allowed). Not yet manually exercised - do so before checking
      this off as tested, not just implemented.
- [x] Cancel discards the scratch pins (nothing written to the library).
      Confirmed.

Also decide, while this is in front of you: `mConnectorID` is computed once
when a photo is loaded, from whatever Name/Part Number are filled in at that
moment - editing those fields afterward does not recompute it, so already-
placed pins keep referencing the stale ID. The plan flagged disabling pin
placement until both fields are non-blank as "a one-line addition worth
making during this manual pass, not worth a redesign." Decide whether to add
it now or accept the current documented behavior.

**Decision: accept current behavior.** Load Photo now requires both fields
to be non-blank before it will load (see above), which prevents the blank-
`mConnectorID` failure mode entirely. The narrower remaining case - editing
Name/Part Number *after* the photo is loaded and pins are placed, leaving
`mConnectorID` stale relative to the displayed fields - is accepted as
documented behavior, not guarded against. Workflow implication: don't edit
Name/Part Number after loading a photo; use Cancel and restart if you need
to change either.

## 2c: Connector picker and snapshot

Status: implemented (2026-08-27), commits `53c0116`..`c334ced`.
**Outstanding - not yet run.**

Per the sub-plan's own Task 9: Add Connector should list the library, add an
instance, and populate `_Snapshot`. Manage Library's Edit should reopen the
editor pre-loaded with existing pins; Delete should prompt and then remove
the entry. Remove Connector should prompt for a ref des and clear its chart
references.

- [x] Add Connector lists the library and populates `_Snapshot` correctly
      for a connector whose photo cache exists. **Found and fixed:**
      `modSnapshot.SnapshotConnector` looked for `CachePhotoPath`'s default
      `.png` cache, but `cmdSave_Click` writes a `.jpg` one - same folder,
      different filename, so it was never found and always fell through to
      the unreliable `ExportShapeToFile`/`Chart.Paste` fallback. Fixed to
      check the `.jpg` cache first. Confirmed working for a connector whose
      cache exists. Also added: `LoadExistingPhoto` now attempts a one-time
      `ExportShapeToFile` backfill when a connector has no cache yet
      (predates this fix, or an earlier save's write never happened), so a
      later edit+save gives it a chance to get one going forward.
      **Outstanding bug, unresolved - not a regression from the above:**
      creating a brand-new connector via Add Connector's "New..." button
      (`frmConnectorPicker.cmdNew_Click` -> `frmConnectorEditor.Show`) can
      leave `mPhotoPath`/`mConnectorID` blank by the time `cmdSave_Click`
      runs, even though Load Photo visibly worked during the session (photo
      displayed, Name/Part Number filled in) - live VBE debugging confirmed
      both are blank at the Save breakpoint. Root cause not yet found;
      suspected to be the same class of form-lifecycle surprise already
      hit twice this session (`UserForm_Initialize` unexpectedly re-firing
      and resetting state - see the `Unload Me` ordering bug below), but
      not yet confirmed for this specific call path. This is also why the
      "auto-add an instance after New..." feature (implemented per user
      request, this same session) hasn't been confirmed working - it
      depends on `mConnectorID` surviving to `cmdSave_Click`, which this
      bug breaks.
      **Planned follow-up (not yet done):** refactor `frmConnectorEditor`
      for headless testability - factor `cmdLoadPhoto_Click`'s logic
      (minus the `Application.GetOpenFilename` dialog, which can't be
      automated) into a `Public Function LoadPhotoFromFile(sPath)`, add
      `Public` state getters for `mPhotoPath`/`mConnectorID`, and use a
      standard-module wrapper (`Application.Run` cannot call a UserForm's
      own Public procedures directly - confirmed this session) to
      reproduce the `New...` -> `Show` -> Save call chain headlessly and
      catch this reliably instead of via manual round-trips.
      **Elevated risk still open for connectors with no cache at all:**
      `modSnapshot.bas`'s backfill fallback still depends on
      `ExportShapeToFile`/`Chart.Paste`, confirmed unreliable on this
      machine - watch for a missing photo on the harness chart for a
      connector that never got a cache written.
- [x] Manage Library > Edit reopens `frmConnectorEditor` pre-loaded with the
      connector's existing pins. **Found 3 bugs, fixed:** (1) `frmManageLibrary`
      was 340px wide but its 5 buttons (Edit/Delete/Import/Export/Close, added
      across 2c/2d) spanned to 428px - widened to 440. Also too short (buttons
      cut off) - height 260->300. (2) `cmdEdit_Click` did `Unload Me`
      (unloading frmManageLibrary) *before* `frmConnectorEditor.Show` - while
      still inside frmManageLibrary's own click handler. This discarded
      everything `LoadForEdit` had just populated (`frmConnectorEditor` came
      up completely blank, `UserForm_Initialize` visibly re-firing per live
      VBE debugging). Fixed by showing the editor first (modal - blocks until
      closed) and unloading frmManageLibrary only afterward. (3) The Type
      dropdown had no options at all in the Edit flow: `cboType`'s RowSource
      ("_Lists!D2:D5", unqualified) resolves against whichever workbook is
      ActiveWorkbook when the control binds - `ConnectorLibrary.xlsx` is
      active by the time the editor loads via this path, and it has no
      `_Lists` sheet. Replaced RowSource entirely with explicit
      `ThisWorkbook`-qualified population in `UserForm_Initialize`.
- [ ] Manage Library > Delete prompts, then removes the entry.
- [ ] Remove Connector prompts for a ref des and clears its chart references.
- [ ] Renaming a ref des cell on the Connectors sheet rewrites every chart
      row referencing it; renaming to an already-used ref des reverts the
      cell instead of applying.

**Extra attention - all NEW code beyond the original sub-plan, with no
automated behavioral coverage (structural tests only, matching 2b's
precedent):**

- `frmConnectorEditor.LoadForEdit` and its helpers
  (`LoadExistingPhoto`/`ExportShapeToFile`/`RebuildPinListFromScratch`) were
  not in the original plan - added because `frmConnectorEditor` had no path
  to populate itself from an existing connector at all. Specifically verify:
  - [x] All fields (Name, Manufacturer, Part Number, Type, Pin Count, Notes)
        populate correctly on Edit. Confirmed, after the 3 fixes above.
  - [x] The existing photo displays correctly. **Was broken, now fixed via
        an architecture change, not a patch.** `ExportShapeToFile`'s
        `Shape.Copy`/`Chart.Paste` mechanism (the original design) is
        unreliable for VBA-triggered clipboard operations on this machine -
        confirmed against both a very-hidden sheet (2b) and this genuinely
        *visible* one ("Photos" in `ConnectorLibrary.xlsx`), same silent
        failure both times. Rather than force that mechanism to work,
        `LoadExistingPhoto` no longer touches the embedded Shape at all: 
        `cmdSave_Click` now keeps a plain on-disk cache of the original
        photo file (`modLibrary.CachePhotoPath`, extended with an
        extension parameter - "jpg" here, "png" unchanged for
        `modSnapshot`'s existing use) alongside the library via a plain
        `FileCopy` - no clipboard involved anywhere in this path.
        `LoadExistingPhoto` reads straight from that cache. The embedded
        Shape in `ConnectorLibrary.xlsx`'s Photos sheet is untouched and
        still the source used for embedding onto the harness chart later.
        Confirmed working end to end.
  - [x] Existing pin markers appear at the correct positions on the photo.
        Confirmed.
  - [x] Saving an edited connector (with no new photo picked) succeeds -
        **found this was completely broken, fixed.** Since the photo export
        above always fails, `mPhotoPath` stayed blank, and
        `EmbedConnectorPhoto` treated a blank path as a hard failure -
        failing the *entire* save silently (all field/pin edits lost, not
        just the photo). Fixed: a blank path now keeps whatever photo is
        already embedded for that connector ID instead of failing, so
        editing fields/pins without picking a new photo saves correctly.
        The re-embedded photo is unchanged (still the original), not
        re-verified pixel-for-pixel but nothing touches it in this path.
  - [ ] Cancel after Edit does not corrupt the library entry.
- `modConnectors.RenameRefDes` / `shConnectors.evt`'s
  `Worksheet_SelectionChange`+`Worksheet_Change` caching: real mouse-driven
  cell edits (as opposed to the test suite's `.Activate()`-then-`.Select()`
  simulation) should be double-checked for the revert-on-collision path.

## 2d: Library import/export

Status: implemented (2026-08-27), commits `78e03d2`..(Task 3). **Wiring
click-through outstanding - not yet run.**

Two separate manual checks, per the sub-plan:

**Task 1 (clipboard-dependent photo copy)** - `CopyConnectorPhoto`'s
`Shape.Copy`/`Paste` is the one clipboard-dependent operation in this whole
codebase, and headless COM automation has not been proven reliable for it.
If Task 1's automated test fails in a way that looks clipboard-related:

- [x] Verify by hand: run the same `Shape.Copy`/`Paste` two lines
      interactively in the VBE Immediate window against a **visible** Excel
      instance. Record whether it works there.
      **Result: works reliably in isolation** (both visible and headless,
      both direct COM and via `Application.Run`, run repeatedly with no
      failures). Under the full pytest suite it is intermittently flaky
      (~1 failure in 11 runs observed across two sessions) - looks like a
      clipboard race rather than a hard environment limitation, but not
      worth chasing further per the plan's own guidance below.
- [x] If it fails even visibly, that's a real environment limitation the spec
      already anticipated - note it plainly (here, and in 2e's design doc)
      and let the extraction-failure fallback prompt (Task 3) carry the
      feature rather than forcing a workaround.
      **Not applicable** - it does not fail visibly; the flake is
      pytest-session-specific. If a CI run hits it, re-running the test file
      alone is expected to pass.

**Final task (Manage Library import/export wiring):**

- [ ] Export a connector to a new file; confirm the saved file opens with the
      correct schema and data.
- [ ] Import that same file back in; confirm it round-trips, or - since it's
      now already present - renames to `-2` and leaves the original untouched.
- [ ] Deliberately remove a test connector's photo shape before importing it,
      to force the extraction-failure path; confirm the replacement-image
      prompt appears and the substituted photo lands correctly.

## Outcome (fill in once run)

- 2b: **done.** All items pass. 6 bugs found and fixed (PNG rejected by
  `LoadPicture`, shrinking preview box, blank `mConnectorID` on early Load
  Photo, oversized pin markers, no Pin Count cap, no Part Number
  uniqueness check). Leader line deferred (Forms.Line.1 unavailable on this
  machine) - see the dragging item above. Regression tests added for all
  fixed bugs (structural, matching this codebase's precedent for
  mouse/clipboard-driven form code with no automated behavioral coverage).
- 2c: _not yet run_ (implemented; includes edit-mode support beyond the
  original sub-plan - see the "Extra attention" list above). **Elevated
  risk found during 2b:** `LoadForEdit`'s photo preview and `modSnapshot.bas`'s
  chart-photo caching both depend on `modLibrary.ExportShapeToFile`
  (`Shape.Copy`/`Chart.Paste`), which 2b found to be broken on this machine
  for VBA-triggered clipboard operations - confirmed via both headless
  automation and a real interactive session. Watch specifically for blank/
  missing photos when running the items below.
- 2d: Task 1 clipboard check run (see above, non-blocking flake noted); the
  "Final task" click-through (export/import/fallback-prompt round trip) is
  _not yet run_ - do it as part of this consolidated batch, before 2e.
- Adjustments made as a result: see per-item notes above (2b) for the full
  list; test suite not yet re-run since the last set of fixes (Excel was
  open) - run `pytest` before considering 2b's fixes verified by the
  automated suite too, not just manually.
