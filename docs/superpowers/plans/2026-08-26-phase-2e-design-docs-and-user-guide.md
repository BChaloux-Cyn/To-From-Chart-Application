# Phase 2e: Design Documentation and Student User Guide Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Document what Phase 2 actually built - a technical design doc per subsystem, verified against the real code rather than the plans that preceded it, plus a student-facing guide covering the workflow that exists today.

**Architecture:** These are retrospective docs, not specs written ahead of implementation. Each technical doc is produced by reading the merged code (not the plan documents) and describing what is actually there - file paths, function signatures, the row-window and sheet-as-storage conventions that emerged across 2a-2d - so drift between plan and reality gets caught and corrected, not silently copied forward. The student guide is scoped strictly to what a student can actually do in the Creator as of the end of Phase 2; it says nothing about Save Harness, Export PDF, or Check Drawing, all of which are Phase 3-4 and do not exist yet.

**Spec:** `docs/superpowers/specs/2026-08-26-harness-creator-design.md`

**Depends on:** 2a, 2b, 2c, and 2d **implemented and merged**, not merely planned. Do not start this plan's tasks until all four are green on `master` - a design doc written against plans rather than code will describe intentions, not behavior, and the whole point of writing this last is to avoid that.

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

---

### Task 1: Verify what actually shipped

**Files:**
- Read-only: everything under `build/`, `src/vba/`, `tests/`.

**Interfaces:**
- Consumes: the merged state of 2a-2d.
- Produces: a working list (not a file - just this task's own output, carried into Tasks 2-3) of every function, sheet, and form actually present, to write the remaining tasks against.

- [ ] **Step 1: Diff plan against reality**

For each of 2a, 2b, 2c, 2d: read the actual current `src/vba/*.bas`, `src/vba/*.cls`, `src/vba/forms/*.evt`, `src/vba/sheets/*.evt`, and `build/*.py`, and compare against that sub-plan's document. Note every place execution deviated from the plan (a renamed function, a different bound, a task that got merged or split differently) - these deviations, not the plan text, are what Tasks 2-3 must describe.

- [ ] **Step 2: Run the full suite as a sanity check**

Run:

```bash
python -m pytest -v
```

Expected: all passed. If it is not green, stop - fix or triage that first. A design doc written against a red suite is documenting a broken state as if it were finished.

- [ ] **Step 3: Commit nothing yet**

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

Each doc follows the same shape as the project's existing spec (`docs/superpowers/specs/2026-08-26-harness-creator-design.md`) - prose sections with tables where the source is tabular, not a copy of the implementation plan's task list. A reader who never saw the plan should be able to understand the subsystem's shape, its key decisions, and where to find its code.

- [ ] **Step 1: Write the 2a doc**

Create `docs/superpowers/specs/2026-08-26-phase-2a-library-core-design.md` covering: the three-table schema and where each table lives (`dist/ConnectorLibrary.xlsx`'s `Connectors`/`Pins`/`Photos` sheets); the bounded-window `(nFirstRow, nLastRow)` convention in `modLibrary.bas` and why it exists (reused unchanged by `_Snapshot` in 2c, which shares a sheet with another table - a whole-sheet `Cells(Rows.Count,...).End(xlUp)` idiom would corrupt that); `LastUsedRowInWindow`'s specific bug (an occupied `nLastRow` overshoots) and why it is `Public`; the field-order-array convention for every record type and why no VBA `Type` crosses `Application.Run`; the photo grid and cache-path scheme. Cite file:line for each claim.

- [ ] **Step 2: Write the 2b doc**

Create `docs/superpowers/specs/2026-08-26-phase-2b-connector-editor-design.md` covering: why UserForms are built via the VBIDE `Designer` object model in `build/form_layout.py` rather than a static `.frm` file; the `_Edit` scratch sheet and why in-progress pin edits live there instead of an in-memory structure; the `modPinEditor.bas` / `frmConnectorEditor.evt` split (all logic testable via `Application.Run`, all UI wiring untestable and manually verified); the anchor-versus-marker distinction and exactly which fields each gesture touches; the `clsPinMarker` `WithEvents`-per-instance pattern and why a form cannot statically `WithEvents` N runtime controls; the deliberate limitation that `mConnectorID` is fixed at photo-load time.

- [ ] **Step 3: Write the 2c doc**

Create `docs/superpowers/specs/2026-08-26-phase-2c-picker-and-snapshot-design.md` covering: `_Snapshot`'s three fixed regions and their row bounds, and why one sheet rather than three; `SnapshotConnector`'s idempotency (frozen once per ConnectorID, not per ref des instance); the `SelectionChange`-caches-prior-value technique `RenameRefDes` depends on and why a plain `Worksheet_Change` cannot detect a rename on its own; the picker/browser split (Add Connector versus Manage Library) and why each button calls exactly one already-tested function.

- [ ] **Step 4: Write the 2d doc**

Create `docs/superpowers/specs/2026-08-26-phase-2d-import-export-design.md` covering: `ImportConnector`'s collision-safe rename and Origin-field stamping; the `Shape.Copy`/`Paste` clipboard mechanism `CopyConnectorPhoto` uses and why it is the one part of this codebase with an honestly uncertain reliability profile in headless automation; **the actual outcome of Task 1's Step 6 manual verification** - state plainly whether headless clipboard photo-copy worked in this environment or not, since that determines whether the extraction-failure fallback is the common path or a rare one in practice.

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
See `docs/student-guide.md` for how to use what's built so far, and the
`docs/superpowers/specs/` design docs for how it's built. Phases 3 and 4
(harness save/load with rendered connector pages, validation and export)
are specified but not yet built.
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
