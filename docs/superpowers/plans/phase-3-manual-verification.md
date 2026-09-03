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

Status: PASS (all 4 items).

- [x] PASS - Click **Save Harness** on a harness with no `HarnessPath` set yet; confirm it
      falls through to Save Harness As (the `GetSaveAsFilename` dialog appears) -
      `modHarnessUI.SaveHarness`'s no-path branch has no automated coverage,
      since `Application.Run` cannot drive a dialog.
- [x] PASS - Save to a real path, then click **Save Harness** again with no changes;
      confirm no dialog appears the second time (the path is now in `_State`) and
      the file is overwritten correctly.
- [x] PASS - Open the saved `.xlsx` in an ordinary (non-headless, interactive) Excel
      session - not just via COM automation in a test - and confirm no "Enable
      Content" prompt appears and no VBA project is present (`Alt+F11` shows no
      project, or is entirely absent from the ribbon).
- [x] PASS - Confirm the `_Snapshot` sheet's photo copy (`modHarnessBuild.CopySnapshot`'s
      `Shape.Copy`/`Paste`, 3a Task 4) is not the clipboard-flakiness case
      `phase-2-manual-verification.md`'s 2d section already documented for a
      different code path - run the full test suite two or three times in a row
      and note whether `test_copy_snapshot_round_trips_connectors_pins_and_photos`
      or the Task 7 integration test ever fails intermittently.
      (Ran `test_copy_snapshot_round_trips_connectors_pins_and_photos` and
      `test_full_harness_round_trips_through_a_saved_file` 3x each: 6/6 passed,
      no flakiness observed.)

## 3b: Connector page rendering

Status: PASS (all 4 items). Bug found and fixed during batch pass (see Outcome).

- [x] PASS - Open a saved harness with at least one connector that has a leader line
      (a marker deliberately pulled off its anchor during the connector editor's
      Place Pins step) and confirm, by looking at the printed/print-previewed
      page, that the leader line renders where expected and looks like a leader
      line (thin, connecting the pulled marker to the correct cavity) rather than
      just being numerically present per `test_leader_line_drawn_only_when_marker_is_pulled_off_anchor`.
      (Note: the connector editor's own in-form visual feedback while dragging is
      still the deferred item from `phase-2-manual-verification.md` 2b - no line
      follows the cursor there. That is separate from this item, which is the
      printed connector page's leader line via `modConnectorPage.PlaceLeaderLines`
      / `Shapes.AddLine` - a different, implemented mechanism. Confirmed rendering
      correctly.)
- [x] PASS - Confirm a connector page's photo is legible at its rendered size (300pt max
      box) and not stretched or distorted - `FitAspectRatio`'s math is tested, but
      "looks right on the page" is not.
- [x] PASS - Confirm oval callouts do not visually overlap the photo's cavities they are
      not meant to label, for a connector with pins close together on the photo.
- [x] PASS - Confirm a connector with no photo cache (the "elevated risk" case
      `phase-2-manual-verification.md`'s 2c section documents) still produces a
      usable page - no photo, but the pin table and metadata are present and the
      page does not look broken or empty in a way that would confuse a student.

## 3c: Live pin-table formulas

Status: PASS (both items). Bug found and fixed during this pass (see Outcome).

- [x] PASS - Open a saved harness in Excel (not headless COM), hand-edit a wire's Color,
      AWG, Length, Signal, or Termination on the `Harness` sheet, and confirm the
      corresponding connector page's pin table updates immediately with no
      "Enable Content" prompt and no visible recalculation delay or error -
      `test_write_live_formulas_fills_every_row` and the round-trip tests confirm
      this programmatically; this item confirms it also holds for a real,
      interactively-typed edit rather than a COM-driven `.Value =` write.
      (Also confirmed the Color/AWG/Termination columns show working dropdowns
      matching the Creator's lists, per the `CopyChartValidation` fix below.)
- [x] PASS - Confirm a formula-column cell (e.g., `Wire To`) does not show a raw `#N/A`
      or other visible error for any pin, wired or not - only a blank cell or the
      expected text.

## 3d: Page setup

Status: PASS (all 8 items). Manual pass surfaced several real gaps in the
print design (not just automated-vs-interactive drift); see Outcome for the
full list of adjustments.

- [x] PASS - Print-preview (`Ctrl+F2` in Excel, not just read `PageSetup` properties) the
      `Harness` sheet of a saved harness with more wires than fit on one printed
      page; confirm the chart header row **and the title block** (rows 1-6) repeat
      on every printed page, the layout is legible in landscape, and the chart
      prints noticeably larger now that margins are narrower.
- [x] PASS - Print-preview a connector page for a connector with only one or two pins;
      confirm the photo is not clipped by the print area's `CONN_PAGE_MIN_PRINT_ROW`
      floor (3d Task 2), the page is now landscape, and the photo/table render
      larger with the narrower margins - if clipped, the constant needs raising,
      not the test loosening.
- [x] PASS - Print-preview a connector page for a connector with many pins (enough to
      overflow past `CONN_PAGE_MIN_PRINT_ROW`); confirm the pin table is not cut
      off mid-row at the page boundary in a way that makes it unreadable.
- [x] PASS - Confirm the footer (harness number, revision, `Page X of Y`) actually
      appears in the print preview / exported-to-PDF-by-hand output, not just in
      `PageSetup.CenterFooter`'s text value.
- [x] PASS - Confirm each connector page's new compact header row (harness number,
      revision, ref des, connector ID) is legible and doesn't overlap the photo
      or pin table.
- [x] PASS - Open the saved harness's `Harness` sheet and confirm the title-block gray
      fill now extends to cover a long Harness Name and a long Description,
      instead of stopping at one narrow column, the labels are present, and
      each gray box has a black border. Confirmed in both the Creator's own
      sheet and a freshly-saved harness file.
- [x] PASS - Use File > Print > Print Entire Workbook to print the full drawing set
      (chart + every connector page) in one job from a saved harness; confirm
      it's legible end to end. This is the documented workaround for print
      scope, not an automatic default - see Outcome for why.
- [x] PASS - Re-verify 3b's photo-legibility, callout-overlap, and leader-line items
      still hold now that connector pages are landscape instead of portrait -
      those were verified under the old portrait layout.

## 3e: Harness load

Status: PASS (all 4 items).

- [x] PASS - Click **Open Harness**, pick a real saved harness file via the file dialog,
      and confirm the Creator's `Harness`, `Connectors`, and connector pages
      (if any were open before) are fully replaced with the loaded file's state -
      `modHarnessUI.OpenHarness`'s dialog and `Workbooks.Open`/`.Close` lifecycle
      has no automated coverage.
- [x] PASS - Deliberately pick a non-harness `.xlsx` file (e.g., a blank workbook) via
      Open Harness; confirm the failure message appears and the Creator's
      existing state is left untouched, not partially overwritten.
- [x] PASS - Load a harness, then immediately try adding a new wire row and a new
      connector; confirm the From/To Conn dropdown and dependent Pin dropdowns
      work correctly against the reconstructed `Connectors` sheet, not just that
      `Validation.Formula1` holds the right text (`test_load_harness_reconstructs_creator_state`
      checks the formula text; this item checks the dropdown actually behaves
      correctly when clicked).
- [x] PASS - Load a harness that was saved from a *different* Creator session (rebuild
      `dist/` between saving and loading, or use two separate `HarnessCreator.xlsm`
      copies) to rule out any accidental in-memory-state dependency the automated
      round-trip test (which loads within the same test run) would not catch.

## Outcome (fill in once run)

- 3a: PASS, all 4 items.
- 3b: PASS, all 4 items. Found the pin-number callout text bug during the batch
  pass (see Adjustments).
- 3c: PASS, both items. Found the missing dropdown-validation bug during the
  batch pass (see Adjustments).
- 3d: PASS, all 8 items. Manual pass surfaced print-design gaps well beyond
  the original checklist items; fixes implemented and re-verified (see
  Adjustments).
- 3e: PASS, all 4 items. No adjustments needed.
- Adjustments made as a result:
  - Connector pages switched from portrait to landscape
    (`modPageSetup.ApplyConnectorPageSetup`), per explicit preference after
    manual print testing - not a bug, a design change. This needs 3b's
    already-PASSed photo/callout/leader-line items re-verified under the new
    orientation.
  - Both `ApplyHarnessPageSetup` and `ApplyConnectorPageSetup` now set
    narrow print margins (0.25in sides, 0.75in top/bottom, 0.3in
    header/footer) via a shared `ApplyNarrowMargins` helper, so the
    `FitToPagesWide/Tall=1` auto-scale prints larger content on both the
    chart and each connector page, per explicit request.
  - `ApplyHarnessPageSetup`'s `PrintTitleRows` widened from just the chart
    header row (`$6:$6`) to the full title-block range (`$1:$6`), so the
    harness name/number/revision/date reprint on every page when the chart
    overflows, per explicit request.
  - New `modConnectorPage.WritePageTitleBlock`, called from
    `modHarnessBuild.BuildConnectorPages`: each connector page now carries
    its own compact header (harness number, revision, ref des, connector
    ID) in a merged `A1:I1`, so the harness/connector context is visible on
    every page of the drawing set, not just the `Harness` sheet.
  - Title-block gray-fill overflow bug: `Harness Name`, `Harness Number`,
    `Revision`, `Student`, `Class/Project`, `Date`, and `Description` value
    cells were each a single narrow chart-grid column, so the gray fill (and
    the text) stopped short for anything longer than that column - most
    visibly `Description`. Fixed by merging each value cell with enough of
    its free neighboring cells to hold realistic text (`Description` gets
    `B4:F4`; the rest get one extra column), in both the Creator's own build
    (`build/layout.py`'s `build_harness`) and the saved harness's from-scratch
    build (`modHarnessBuild.CopyTitleBlock`). `Length Units` (`H4`) is a
    short controlled value and was left unmerged.
  - That merge exposed a second, previously-latent bug it's important to
    call out separately: `modChart.NewHarness` cleared the title block via
    `Names(...).RefersToRange.ClearContents`, and Excel refuses
    `ClearContents` on a range that covers only part of a merged cell - even
    the merge's own top-left anchor cell. Confirmed via a standalone COM
    script reproducing `"We can't do that to a merged cell."` before the fix.
    New Harness and Load Harness (which calls `NewHarness` first) would
    silently fail partway through and leave stale state (wrong length units,
    dirty flag not cleared, connector list/check results not cleared) with no
    visible error, since the failure was swallowed by an existing `On Error
    GoTo CleanUp`. Fixed by clearing `.MergeArea` instead of the bare range,
    which is a no-op-safe superset for non-merged cells too.
  - Print scope: printing/print-previewing shows only the active sheet by
    default; the desired combined "chart + every connector page" print job
    would need a `Workbook_BeforePrint` VBA hook, which conflicts with the
    saved harness's macro-free design (verified by 3a). Decision: document
    File > Print > Print Entire Workbook in the user guide as the way to
    print the full drawing set, rather than automating it.
  - Tests added: `test_apply_harness_page_setup` and
    `test_apply_connector_page_setup` extended for orientation/margins/
    `PrintTitleRows`; `test_copy_title_block_widens_long_value_cells_with_a_merge`;
    `test_write_page_title_block_shows_harness_and_connector_metadata`;
    `test_build_connector_pages_creates_one_sheet_per_instance` extended for
    the new header. Full suite (438 tests) passes; `dist/` rebuilt fresh.
  - Second manual pass on the above found two more real gaps in the saved
    harness's title block specifically: (1) `modHarnessBuild.CopyTitleBlock`
    only ever copied the title-block *values* (`B2,E2,H2,...`), never wrote
    the labels (`Harness Name`, `Student`, `Description`, etc.) at all - the
    saved file's title block was unlabeled gray boxes. (2) the saved
    harness's `Harness` sheet had no column widths set anywhere (unlike the
    Creator's own chart sheet, which gets `CHART_COLUMN_WIDTHS` from
    `build/layout.py`), so even the merged value cells sat at Excel's
    default width and a real Harness Name/Description still clipped. Fixed
    by writing the labels in `CopyTitleBlock` (bold, matching
    `build/layout.py`'s `TITLE_BLOCK` text), giving the saved harness its own
    copy of chart column widths (mirrored as a VBA constant, since the
    saved file is built from a blank workbook rather than a template), and
    widening columns A and G (in both `build/layout.py` and the VBA mirror)
    past their original chart-only sizing so `Harness Name`/`Description`
    and `Length Units` labels fit. Also added a black border around every
    title-block gray value box (both the Creator's own build and the saved
    harness's), per explicit request. Added
    `test_harness_title_block_labels_are_bordered_and_columns_fit_them`
    (Creator sheet), `test_copy_title_block_borders_and_widens_label_columns`,
    and extended `test_copy_title_block_round_trips_every_field` for the
    labels. Full suite (440 tests) passes; `dist/` rebuilt fresh.
  - `modHarnessBuild.CopyChartValidation` (called from `modHarnessActions.SaveHarness`):
    the saved harness's `Harness` sheet had no dropdown validation on Color, AWG,
    or Termination - `CopyChartRows` only ever copied cell values, not the
    Creator's chart validation. Fixed by copying the Creator's `_Lists` sheet
    into a new very-hidden `_Lists` sheet in the saved workbook, defining
    `ListColor`/`ListAWG`/`ListTermination` named ranges over it, and applying
    list validation to the Color, AWG, From Term, and To Term columns. From/To
    Conn and From/To Pin are intentionally excluded (per-harness-dynamic, not a
    static list). Added `test_lists_sheet_is_very_hidden`,
    `test_copy_chart_validation_matches_the_creators_dropdowns`, and
    `test_copy_chart_validation_lists_resolve_to_the_creators_values` to
    `tests/test_harness_build.py`. Full suite (436 tests) passes; `dist/`
    rebuilt fresh. Manually confirmed the dropdowns appear and list the
    correct values in an interactive Excel session.
  - `modConnectorPage.PlaceCallouts` set the oval callout's fill to white but
    never set the pin-number text's font color, so it rendered in the
    TextFrame2 default color (white), invisible against the white fill. Also,
    the text was left-aligned/top-anchored with the shape's default margins
    (which are large relative to the 14pt oval), so once color was fixed the
    number rendered off-center overlapping the oval's edge, and was not bold
    per additional feedback. Fixed by explicitly setting
    `Font.Fill.ForeColor.RGB = RGB(0,0,0)`, `Font.Bold = msoTrue`,
    `ParagraphFormat.Alignment = msoAlignCenter`, `VerticalAnchor =
    msoAnchorMiddle`, and zeroing all four `TextFrame2` margins. Added
    `test_place_callouts_draws_black_text_on_the_white_oval` and
    `test_place_callouts_text_is_bold_and_centered_in_the_oval` to
    `tests/test_connector_page.py` (both reproduced the bug before the fix).
    Full suite (433 tests) passes after the fix; `dist/` rebuilt fresh.
