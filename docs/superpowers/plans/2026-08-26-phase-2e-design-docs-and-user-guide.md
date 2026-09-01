# Phase 2e: Design Documentation and Student User Guide Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Document what Phase 2 actually built - a technical design doc per subsystem, verified against the real code rather than the plans that preceded it, plus a student-facing guide covering the workflow that exists today.

**Architecture:** These are retrospective docs, not specs written ahead of implementation. Each technical doc is produced by reading the merged code (not the plan documents) and describing what is actually there - file paths, function signatures, the row-window and sheet-as-storage conventions that emerged across 2a-2d - so drift between plan and reality gets caught and corrected, not silently copied forward. The student guide is scoped strictly to what a student can actually do in the Creator as of the end of Phase 2; it says nothing about Save Harness, Export PDF, or Check Drawing, all of which are Phase 3-4 and do not exist yet.

**Spec:** `docs/superpowers/specs/2026-08-26-harness-creator-design.md`

**Existing documentation to build on, not duplicate:** `docs/design/` (`00_purpose_scope.md` through `04_enforcement.md`) already documents the layer 0/1/2 split, the `modContract` result envelope, handler-lifecycle discipline, and the automated checks in `tests/test_layering.py` - all added by the UI/logic-separation work after this plan was written. Every per-subsystem doc in Task 2 should link into that structure for those cross-cutting conventions (see each step's guidance below) rather than re-explaining them, and should describe the subsystem in terms of where its code now sits in that layering (which layer 0 primitives, which layer 1 action module, which `.evt` adapters) instead of the pre-refactor form/module split this plan was originally scoped against.

**Depends on:** 2a, 2b, 2c, and 2d **implemented and merged**, not merely planned, *and* the UI/logic-separation plan (`docs/superpowers/plans/2026-08-28-ui-logic-separation.md`) merged as well - it restructured 2a-2d's code into the layer 0/1/2 split `docs/design/` now documents, so a doc written against the pre-separation structure would already be describing dead code. Do not start this plan's tasks until all of these are green on `master` - a design doc written against plans rather than code will describe intentions, not behavior, and the whole point of writing this last is to avoid that.

**Part of Phase 2** (2a-2d cover the implementation; this is 2e, the last sub-plan).

## Global Constraints

- Every claim in a technical doc must be verifiable against a specific file and, where it names behavior, a specific test. A claim with no such anchor is a guess, not documentation.
- The student guide describes only what is built and working as of this plan's execution. Anything from Phase 3+ (save, load, export, Check Drawing) is out of scope for it entirely - not "coming soon," just absent, so the guide never goes stale by promising something that doesn't exist yet.
- No emojis, per the user's global preference.

## File Structure

| File | Responsibility |
|---|---|
| `docs/superpowers/specs/2026-08-26-phase-2a-library-core-design.md` | What 2a actually built: schema, bounded-window reader/writer, why |
| `docs/superpowers/specs/2026-08-26-phase-2b-connector-editor-design.md` | What 2b actually built: form-in-code, click-to-place, the `_Edit` scratch pattern |
| `docs/superpowers/specs/2026-08-26-phase-2c-picker-and-snapshot-design.md` | What 2c actually built: `_Snapshot`'s regions, ref des rename, the picker/browser split |
| `docs/superpowers/specs/2026-08-26-phase-2d-import-export-design.md` | What 2d actually built: collision-safe import, the clipboard-photo-copy risk and its fallback |
| `docs/student-guide.md` | Plain-language guide to the Creator as it stands after Phase 2 |
| `README.md` | Modified: Status section updated to reflect Phase 2 |
| `docs/design/00_purpose_scope.md` through `04_enforcement.md` | **Not modified by this plan** - existing cross-cutting layering reference, linked from each Task 2 doc rather than duplicated |

---

### Task 1: Verify what actually shipped

**Files:**
- Read-only: everything under `build/`, `src/vba/`, `tests/`, `docs/design/`.

**Interfaces:**
- Consumes: the merged state of 2a-2d and the UI/logic-separation plan.
- Produces: a working list (not a file - just this task's own output, carried into Tasks 2-3) of every function, sheet, and form actually present, to write the remaining tasks against.

- [ ] **Step 1: Read the current layering reference first**

Read `docs/design/00_purpose_scope.md` through `04_enforcement.md` before touching the sub-plans. They already describe the layer 0/1/2 split, the `modContract` envelope, and handler-lifecycle rules that now sit on top of everything 2a-2d built. Treat this as the current architecture, not optional background - it tells you which module each piece of 2a-2d logic actually lives in today.

- [ ] **Step 2: Diff plan against reality**

For each of 2a, 2b, 2c, 2d: read the actual current `src/vba/*.bas`, `src/vba/*.cls`, `src/vba/forms/*.evt`, `src/vba/sheets/*.evt`, and `build/*.py`, and compare against that sub-plan's document. Note every place execution deviated from the plan - a renamed function, a different bound, a task that got merged or split differently, *and* every place logic that sub-plan describes as living in a form or sheet module has since moved into a layer 1 action module (`modEditorActions`, `modPickerActions`, `modManageActions`) per the UI/logic-separation plan. These deviations, not the plan text, are what Tasks 2-3 must describe.

- [ ] **Step 3: Run the full suite as a sanity check**

Run:

```bash
python -m pytest -v
```

Expected: all passed. If it is not green, stop - fix or triage that first. A design doc written against a red suite is documenting a broken state as if it were finished.

- [ ] **Step 4: Commit nothing yet**

This task produces no file changes - it is the reading pass that grounds Tasks 2 and 3. Proceed directly to Task 2.

---

### Task 2: Per-subsystem technical design docs

**Files:**
- Create: `docs/superpowers/specs/2026-08-26-phase-2a-library-core-design.md`
- Create: `docs/superpowers/specs/2026-08-26-phase-2b-connector-editor-design.md`
- Create: `docs/superpowers/specs/2026-08-26-phase-2c-picker-and-snapshot-design.md`
- Create: `docs/superpowers/specs/2026-08-26-phase-2d-import-export-design.md`

**Interfaces:**
- Consumes: Task 1's reading pass.
- Produces: four design docs, one per subsystem.

Each doc follows the same shape as the project's existing spec (`docs/superpowers/specs/2026-08-26-harness-creator-design.md`) - prose sections with tables where the source is tabular, not a copy of the implementation plan's task list. A reader who never saw the plan should be able to understand the subsystem's shape, its key decisions, and where to find its code. For the layer 0/1/2 split, the result envelope, and handler-lifecycle rules, link to the relevant `docs/design/*.md` file instead of re-explaining them - these docs describe how each subsystem uses that architecture, not the architecture itself.

- [ ] **Step 1: Write the 2a doc**

Create `docs/superpowers/specs/2026-08-26-phase-2a-library-core-design.md` covering: the three-table schema and where each table lives (`dist/ConnectorLibrary.xlsx`'s `Connectors`/`Pins`/`Photos` sheets); the bounded-window `(nFirstRow, nLastRow)` convention in `modLibrary.bas` and why it exists (reused unchanged by `_Snapshot` in 2c, which shares a sheet with another table - a whole-sheet `Cells(Rows.Count,...).End(xlUp)` idiom would corrupt that); `LastUsedRowInWindow`'s specific bug (an occupied `nLastRow` overshoots) and why it is `Public`; the field-order-array convention for every record type and why no VBA `Type` crosses `Application.Run`; the photo grid and cache-path scheme. Cite file:line for each claim. `modLibrary.bas` is a layer 0 primitives module per `docs/design/01_architecture_overview.md` - note that in passing, but this doc's focus is the schema and storage conventions themselves, which sit below and are unaffected by the layering added on top.

- [ ] **Step 2: Write the 2b doc**

Create `docs/superpowers/specs/2026-08-26-phase-2b-connector-editor-design.md` covering: why UserForms are built via the VBIDE `Designer` object model in `build/form_layout.py` rather than a static `.frm` file; the `_Edit` scratch sheet and why in-progress pin edits live there instead of an in-memory structure; the current three-way split across `modPinEditor.bas` (layer 0 primitives), `modEditorActions.bas` (layer 1 - the photo guard, cache source, click/place/delete/pin-number actions and queries, all tested via `Application.Run`) and `frmConnectorEditor.evt` (layer 2 - the thin adapter that reads controls, calls one `modEditorActions` function, and writes controls; this is the only piece manually verified rather than tested) - see `docs/design/01_architecture_overview.md` for what "layer" means here; the anchor-versus-marker distinction and exactly which fields each gesture touches; the `clsPinMarker` `WithEvents`-per-instance pattern and why a form cannot statically `WithEvents` N runtime controls; the deliberate limitation that `mConnectorID` is fixed at photo-load time.

- [ ] **Step 3: Write the 2c doc**

Create `docs/superpowers/specs/2026-08-26-phase-2c-picker-and-snapshot-design.md` covering: `_Snapshot`'s three fixed regions and their row bounds, and why one sheet rather than three; `SnapshotConnector`'s idempotency (frozen once per ConnectorID, not per ref des instance); the `SelectionChange`-caches-prior-value technique `RenameRefDes` depends on and why a plain `Worksheet_Change` cannot detect a rename on its own; the picker/browser split (Add Connector versus Manage Library) and why each button now calls exactly one `modPickerActions` or `modManageActions` function (layer 1) rather than doing the work itself - `docs/design/04_enforcement.md` names the test that holds this invariant (`test_every_click_handler_delegates`).

- [ ] **Step 4: Write the 2d doc**

Create `docs/superpowers/specs/2026-08-26-phase-2d-import-export-design.md` covering: `ImportConnector`'s collision-safe rename and Origin-field stamping, and how that has since been extended into the Keep/Overwrite conflict prompt and delete-cascade behavior driven by `modManageActions`/`modLibraryTransfer` (layer 1) and reported through `modContract`'s `IMPORTED`/`EXPORTED` outcome codes and `modMessages`' text for them (see `docs/design/02_layering_rules.md`); the `Shape.Copy`/`Paste` clipboard mechanism `CopyConnectorPhoto` uses and why it is the one part of this codebase with an honestly uncertain reliability profile in headless automation - `tests/test_snapshot.py`'s own comments and `modSnapshot.SnapshotConnector`'s fallback-of-last-resort ordering document the same conclusion for the sibling photo-export path, worth cross-referencing; **the actual outcome of Task 1's Step 6 manual verification** - state plainly whether headless clipboard photo-copy worked in this environment or not, since that determines whether the extraction-failure fallback is the common path or a rare one in practice.

- [ ] **Step 5: Cross-check every doc against Task 1's deviation notes**

For each doc, confirm it describes what Task 1 found actually shipped, not what the corresponding plan said would be built. Where they differ, the doc must reflect reality.

- [ ] **Step 6: Commit**

```bash
git add docs/superpowers/specs/2026-08-26-phase-2a-library-core-design.md \
        docs/superpowers/specs/2026-08-26-phase-2b-connector-editor-design.md \
        docs/superpowers/specs/2026-08-26-phase-2c-picker-and-snapshot-design.md \
        docs/superpowers/specs/2026-08-26-phase-2d-import-export-design.md
git commit -m "docs: add retrospective design docs for phase 2's four subsystems"
```

---

### Task 3: Student user guide

**Files:**
- Create: `docs/student-guide.md`
- Modify: `README.md`

**Interfaces:**
- Consumes: Task 1's reading pass.
- Produces: a guide a student can follow start to finish using only what exists today.

- [ ] **Step 1: Write the guide**

Create `docs/student-guide.md` with these sections, each describing an actual, currently-working action - open `HarnessCreator.xlsm`, walk each flow by hand while writing this, and describe exactly what happens:

- **What this tool is** - one paragraph, adapted from the spec's Purpose section, written for a student rather than a developer.
- **Adding a connector to your harness** - Home > Add Connector, picking an existing library part versus New... to define one that doesn't exist yet.
- **Defining a new connector** - the connector editor: fields, Load Photo, Place Pins (click to drop a numbered marker), dragging a marker versus reselecting a pin and clicking its real position, Snap Label to Pin, Save.
- **Filling in the to-from chart** - the Harness sheet's columns, how From Conn drives the From Pin dropdown, what each column means.
- **Managing your connector library** - Home > Manage Library: editing an existing connector, deleting one, and (state plainly, based on Task 1's findings) importing a connector another student shared with you, exporting one to share back.
- **Renaming a connector** - editing its ref des on the Connectors sheet updates the chart automatically; two connectors can't share a ref des.
- **Starting over** - New Harness.

Do not include a "Saving your work" or "Printing / exporting a PDF" section - that functionality does not exist yet. If a student would reasonably ask "how do I save this," that question is answered by Phase 3, not this guide.

- [ ] **Step 2: Update the README's Status section**

In `README.md`, replace the "Status" section's final paragraph to reflect Phase 2:

```markdown
## Status

Phase 1 complete: build system, Creator shell, to-from chart with dependent
pin dropdowns. Phase 2 complete: connector library file and reader/writer,
the connector editor with click-to-place pin markers, the connector picker,
per-harness connector snapshots, ref des rename, and library import/export.
See `docs/student-guide.md` for how to use what's built so far, the
`docs/superpowers/specs/` design docs for each subsystem's decisions, and
`docs/design/` for the layering rules the code follows day to day. Phases 3
and 4 (harness save/load with rendered connector pages, validation and
export) are specified but not yet built.
```

- [ ] **Step 3: Verify the guide against the running application**

Open `dist/HarnessCreator.xlsm` and follow the guide section by section, exactly as written, doing nothing it doesn't say to do. Every step must work as described. Fix the guide, not the workbook, for any mismatch that isn't a real bug - if it is a real bug, stop and report it rather than documenting around it.

- [ ] **Step 4: Commit**

```bash
git add docs/student-guide.md README.md
git commit -m "docs: add the student user guide for the phase 2 workflow"
```

---

## Self-Review

**Spec coverage.** Your "Both" and "written last" answers are both satisfied: per-subsystem technical docs (Task 2) plus the student user guide (Task 3), sequenced after 2a-2d's implementation via Task 1's mandatory reading pass rather than authored against the plans.

**Why this plan has no failing-test/passing-test steps like 2a-2d.** There is no code here to test - the deliverable is prose, and its correctness criterion is different: verified against the running application (Task 3 Step 3) and cross-checked against actual shipped code rather than plan text (Task 2 Step 5), not asserted by pytest. Task 1's "run the full suite" step is the closest equivalent - a sanity check that there is something real to document before documenting it.

**A real risk this plan depends on, not created by it.** Task 2 Step 4 explicitly requires stating 2d's clipboard-copy outcome honestly, whichever way it went. If it turned out unreliable, the design doc and this plan's confidence in "library import and export" as a finished feature both need to say so plainly - a design doc that asserts a feature works when Task 1's own verification found otherwise would be a worse outcome than 2d shipping with a known, documented limitation.

**No placeholders.** Every doc section above is scoped to specific, concrete content to verify and write, not "document the API" or "write user docs" left unspecified.
