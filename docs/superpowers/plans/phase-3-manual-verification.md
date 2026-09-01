# Phase 3 Manual Verification Checklist

Every Phase 3 sub-plan (3a-3e) leaves at least one thing pytest cannot exercise: a
mouse-driven dialog, clipboard-dependent COM automation, or a visual/print result
that a numeric assertion on `PageSetup`/`Shape` properties cannot fully confirm.
This doc collects them in one place so they get run as a single batch **after 3e is
complete, before 3f (docs)** - matching `docs/superpowers/plans/phase-2-manual-verification.md`'s
precedent exactly, so the retrospective design docs and student user guide update
describe verified, adjustment-if-needed behavior rather than being written against
code no one has actually clicked through or printed.

Do not start 3f until this checklist is run and any adjustments it surfaces are made.

## How to use this

**During execution of 3a-3e:** if a task's own plan text does not already name a
manual-only item its work introduces, the executor adds one here, under the
matching sub-plan's section, before moving on - the same way each Phase 2
sub-plan's own "Manually verify in Excel" note fed into
`phase-2-manual-verification.md`. Do not defer discovering these items to the
batch pass at the end; only running them is deferred, not noticing them.

**When running the batch pass:** open `dist/HarnessCreator.xlsm` fresh (`rm -rf
dist && python build/build.py`), work through each sub-plan's items in order
(3a through 3e - later ones depend on earlier ones having produced a real saved
file to load), and record pass/fail plus any adjustment made directly in this
file (not just in a commit message), so the record survives past any one session.

## 3a: Harness workbook shell and Save

Status: not yet run.

- [ ] Click **Save Harness** on a harness with no `HarnessPath` set yet; confirm it
      falls through to Save Harness As (the `GetSaveAsFilename` dialog appears) -
      `modHarnessUI.SaveHarness`'s no-path branch has no automated coverage,
      since `Application.Run` cannot drive a dialog.
- [ ] Save to a real path, then click **Save Harness** again with no changes;
      confirm no dialog appears the second time (the path is now in `_State`) and
      the file is overwritten correctly.
- [ ] Open the saved `.xlsx` in an ordinary (non-headless, interactive) Excel
      session - not just via COM automation in a test - and confirm no "Enable
      Content" prompt appears and no VBA project is present (`Alt+F11` shows no
      project, or is entirely absent from the ribbon).
- [ ] Confirm the `_Snapshot` sheet's photo copy (`modHarnessBuild.CopySnapshot`'s
      `Shape.Copy`/`Paste`, 3a Task 4) is not the clipboard-flakiness case
      `phase-2-manual-verification.md`'s 2d section already documented for a
      different code path - run the full test suite two or three times in a row
      and note whether `test_copy_snapshot_round_trips_connectors_pins_and_photos`
      or the Task 7 integration test ever fails intermittently.

## 3b: Connector page rendering

Status: not yet run.

- [ ] Open a saved harness with at least one connector that has a leader line
      (a marker deliberately pulled off its anchor during the connector editor's
      Place Pins step) and confirm, by looking at the printed/print-previewed
      page, that the leader line renders where expected and looks like a leader
      line (thin, connecting the pulled marker to the correct cavity) rather than
      just being numerically present per `test_leader_line_drawn_only_when_marker_is_pulled_off_anchor`.
- [ ] Confirm a connector page's photo is legible at its rendered size (300pt max
      box) and not stretched or distorted - `FitAspectRatio`'s math is tested, but
      "looks right on the page" is not.
- [ ] Confirm oval callouts do not visually overlap the photo's cavities they are
      not meant to label, for a connector with pins close together on the photo.
- [ ] Confirm a connector with no photo cache (the "elevated risk" case
      `phase-2-manual-verification.md`'s 2c section documents) still produces a
      usable page - no photo, but the pin table and metadata are present and the
      page does not look broken or empty in a way that would confuse a student.

## 3c: Live pin-table formulas

Status: not yet run.

- [ ] Open a saved harness in Excel (not headless COM), hand-edit a wire's Color,
      AWG, Length, Signal, or Termination on the `Harness` sheet, and confirm the
      corresponding connector page's pin table updates immediately with no
      "Enable Content" prompt and no visible recalculation delay or error -
      `test_write_live_formulas_fills_every_row` and the round-trip tests confirm
      this programmatically; this item confirms it also holds for a real,
      interactively-typed edit rather than a COM-driven `.Value =` write.
- [ ] Confirm a formula-column cell (e.g., `Wire To`) does not show a raw `#N/A`
      or other visible error for any pin, wired or not - only a blank cell or the
      expected text.

## 3d: Page setup

Status: not yet run.

- [ ] Print-preview (`Ctrl+F2` in Excel, not just read `PageSetup` properties) the
      `Harness` sheet of a saved harness with more wires than fit on one printed
      page; confirm the chart header row repeats on every printed page and the
      layout is legible in landscape.
- [ ] Print-preview a connector page for a connector with only one or two pins;
      confirm the photo is not clipped by the print area's `CONN_PAGE_MIN_PRINT_ROW`
      floor (3d Task 2) - if it is clipped, the constant needs raising, not the
      test loosening.
- [ ] Print-preview a connector page for a connector with many pins (enough to
      overflow past `CONN_PAGE_MIN_PRINT_ROW`); confirm the pin table is not cut
      off mid-row at the page boundary in a way that makes it unreadable.
- [ ] Confirm the footer (harness number, revision, `Page X of Y`) actually
      appears in the print preview / exported-to-PDF-by-hand output, not just in
      `PageSetup.CenterFooter`'s text value.

## 3e: Harness load

Status: not yet run.

- [ ] Click **Open Harness**, pick a real saved harness file via the file dialog,
      and confirm the Creator's `Harness`, `Connectors`, and connector pages
      (if any were open before) are fully replaced with the loaded file's state -
      `modHarnessUI.OpenHarness`'s dialog and `Workbooks.Open`/`.Close` lifecycle
      has no automated coverage.
- [ ] Deliberately pick a non-harness `.xlsx` file (e.g., a blank workbook) via
      Open Harness; confirm the failure message appears and the Creator's
      existing state is left untouched, not partially overwritten.
- [ ] Load a harness, then immediately try adding a new wire row and a new
      connector; confirm the From/To Conn dropdown and dependent Pin dropdowns
      work correctly against the reconstructed `Connectors` sheet, not just that
      `Validation.Formula1` holds the right text (`test_load_harness_reconstructs_creator_state`
      checks the formula text; this item checks the dropdown actually behaves
      correctly when clicked).
- [ ] Load a harness that was saved from a *different* Creator session (rebuild
      `dist/` between saving and loading, or use two separate `HarnessCreator.xlsm`
      copies) to rule out any accidental in-memory-state dependency the automated
      round-trip test (which loads within the same test run) would not catch.

## Outcome (fill in once run)

- 3a: not yet run.
- 3b: not yet run.
- 3c: not yet run.
- 3d: not yet run.
- 3e: not yet run.
- Adjustments made as a result: none yet.
