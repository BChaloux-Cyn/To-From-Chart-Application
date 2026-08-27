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

- [ ] Load Photo fits the image without letterboxing.
- [ ] Place Pins drops numbered markers on click.
- [ ] Dragging a marker moves only the marker, and shows a leader line once
      pulled away from its anchor.
- [ ] Selecting a pin in the list, then clicking the image, moves that pin's
      anchor.
- [ ] Snap Label to Pin removes the leader (marker returns to the anchor).
- [ ] Save writes to `ConnectorLibrary.xlsx` and closes the form.
- [ ] Cancel discards the scratch pins (nothing written to the library).

Also decide, while this is in front of you: `mConnectorID` is computed once
when a photo is loaded, from whatever Name/Part Number are filled in at that
moment - editing those fields afterward does not recompute it, so already-
placed pins keep referencing the stale ID. The plan flagged disabling pin
placement until both fields are non-blank as "a one-line addition worth
making during this manual pass, not worth a redesign." Decide whether to add
it now or accept the current documented behavior.

## 2c: Connector picker and snapshot

Status: implemented (2026-08-27), commits `53c0116`..`c334ced`.
**Outstanding - not yet run.**

Per the sub-plan's own Task 9: Add Connector should list the library, add an
instance, and populate `_Snapshot`. Manage Library's Edit should reopen the
editor pre-loaded with existing pins; Delete should prompt and then remove
the entry. Remove Connector should prompt for a ref des and clear its chart
references.

- [ ] Add Connector lists the library and populates `_Snapshot` correctly.
- [ ] Manage Library > Edit reopens `frmConnectorEditor` pre-loaded with the
      connector's existing pins.
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
  - [ ] All fields (Name, Manufacturer, Part Number, Type, Pin Count, Notes)
        populate correctly on Edit.
  - [ ] The existing photo displays correctly (this round-trips through a
        temp-file export of the embedded Shape via a throwaway ChartObject -
        `ExportShapeToFile` - which has no automated coverage at all).
  - [ ] Existing pin markers appear at the correct positions on the photo.
  - [ ] Saving an edited connector (with no new photo picked) succeeds and
        the re-embedded photo in `ConnectorLibrary.xlsx` still looks correct
        - this is the scenario `ExportShapeToFile` exists to support.
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

- 2b: _not yet run_
- 2c: _not yet run_ (implemented; includes edit-mode support beyond the
  original sub-plan - see the "Extra attention" list above)
- 2d: Task 1 clipboard check run (see above, non-blocking flake noted); the
  "Final task" click-through (export/import/fallback-prompt round trip) is
  _not yet run_ - do it as part of this consolidated batch, before 2e.
- Adjustments made as a result: _none yet_
