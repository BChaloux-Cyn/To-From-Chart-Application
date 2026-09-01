# Phase 3f: Design Documentation and Student User Guide Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Document what Phase 3 actually built - a technical design doc per subsystem, verified against the real code rather than the plans that preceded it, plus a student user guide update covering Save, Save As, and Open Harness now that they exist.

**Architecture:** Retrospective docs, following `docs/superpowers/plans/2026-08-26-phase-2e-design-docs-and-user-guide.md`'s precedent exactly: each technical doc is produced by reading the merged code (not 3a-3e's plan documents) and describing what is actually there, so drift between plan and reality is caught rather than copied forward. The student guide addition is scoped strictly to what a student can do as of the end of Phase 3 - it says nothing about Export PDF, Save Archive Copy, or Check Drawing, all of which remain Phase 4 and do not exist yet.

**Spec:** `docs/superpowers/specs/2026-08-26-harness-creator-design.md`

**Existing documentation to build on, not duplicate:** `docs/design/` (the layer 0/1/2 split, the `modContract` envelope, handler-lifecycle rules) and the four Phase 2 subsystem docs in `docs/superpowers/specs/2026-08-26-phase-2*-design.md`. Every doc in Task 2 links into `docs/design/` for cross-cutting conventions rather than re-explaining them, and references the Phase 2 library-core doc where Phase 3 reuses its schema (`_Snapshot`'s row-window convention, the `LBound`-relative payload indexing) rather than re-describing it.

**Depends on:** 3a, 3b, 3c, 3d, and 3e **implemented and merged**, and `phase-3-manual-verification.md` **run, with any adjustments it surfaced already made**. Do not start this plan's tasks until all of that is true and the suite is green on `master` - a design doc written against unverified behavior describes intentions, not behavior, which is the same risk 2e's own header calls out.

**Part of Phase 3** (see 3a's header for the full breakdown). This is the last sub-plan.

## Global Constraints

- Every claim in a technical doc must be verifiable against a specific file and, where it names behavior, a specific test. A claim with no such anchor is a guess, not documentation.
- The student guide addition describes only what is built and working as of this plan's execution, confirmed against `phase-3-manual-verification.md`'s outcome, not merely against passing pytest. Anything from Phase 4 (Export PDF, Save Archive Copy, Check Drawing) stays absent, not "coming soon."
- No emojis, per the user's global preference.

## File Structure

| File | Responsibility |
|---|---|
| `docs/superpowers/specs/2026-09-01-phase-3a-harness-save-shell-design.md` | What 3a actually built: the `Harness`/`_Snapshot` shell, the live join-key columns, Save/Save As |
| `docs/superpowers/specs/2026-09-01-phase-3b-connector-page-rendering-design.md` | What 3b actually built: photo placement, oval callouts, leader lines, the metadata cell |
| `docs/superpowers/specs/2026-09-01-phase-3c-live-pin-tables-design.md` | What 3c actually built: the join-key lookup formulas and their From/To asymmetry |
| `docs/superpowers/specs/2026-09-01-phase-3d-page-setup-design.md` | What 3d actually built: the differentiated Harness/connector-page print setup |
| `docs/superpowers/specs/2026-09-01-phase-3e-harness-load-design.md` | What 3e actually built: reconstructing `Connectors` from `CONN_<RefDes>` sheet names, the load-time `EnableEvents` discipline |
| `docs/student-guide.md` | Modified: adds Save, Save As, and Open Harness sections |
| `README.md` | Modified: Status section updated to reflect Phase 3 |

---

### Task 1: Verify what actually shipped

**Files:**
- Read-only: everything under `build/`, `src/vba/`, `tests/`, `docs/design/`, and `phase-3-manual-verification.md`'s filled-in Outcome section.

**Interfaces:**
- Consumes: the merged state of 3a-3e and the completed manual-verification pass.
- Produces: a working list (not a file) of every function, sheet, and formula actually present, to write Tasks 2-3 against.

- [ ] **Step 1: Read `phase-3-manual-verification.md`'s Outcome section first**

Any adjustment recorded there (a constant changed, a leader-line rendering fix, a page-setup constant raised) supersedes the corresponding sub-plan's original text - the doc must describe what shipped after those adjustments, not the plan's first draft.

- [ ] **Step 2: Diff plan against reality**

For each of 3a, 3b, 3c, 3d, 3e: read the actual current `src/vba/modHarness*.bas`, `src/vba/modConnectorPage.bas`, `src/vba/modPageSetup.bas`, and `build/*.py`, and compare against that sub-plan's document. Note every place execution deviated - a renamed function, a different constant value, a task that got merged or split differently.

- [ ] **Step 3: Run the full suite as a sanity check**

Run:

```bash
python -m pytest -v
```

Expected: all passed. If it is not green, stop - fix or triage that first.

- [ ] **Step 4: Commit nothing yet**

This task produces no file changes - proceed directly to Task 2.

---

### Task 2: Per-subsystem technical design docs

**Files:**
- Create the five files listed in File Structure above.

**Interfaces:**
- Consumes: Task 1's reading pass.
- Produces: five design docs, one per sub-plan.

Each doc follows the same shape as the Phase 2 subsystem docs and the project spec - prose with tables where the source is tabular, not a copy of the implementation plan's task list.

- [ ] **Step 1: Write the 3a doc**

Cover: why the saved `Harness` sheet keeps the Creator's full 1000-row range instead of trimming to content; the two hidden join-key columns as live formulas (`=IF(A7="","",A7&"|"&B7)`) rather than static text, and why that specifically is what lets a hand edit survive a round trip; why no connector-instance table is saved directly, deferring that reconstruction to 3e; the `BuildHarnessSheets` fresh-workbook contract mirrored from `modManageActions.ExportToWorkbook`'s existing precedent. Cite file:line for each claim, and link to `docs/design/02_layering_rules.md` (or wherever the layer 0/1/2 split is documented) for the `modHarnessBuild`/`modHarnessActions`/`modHarnessUI` split rather than re-explaining it.

- [ ] **Step 2: Write the 3b doc**

Cover: why `modConnectorPage.bas` reuses `modPinEditor`'s plain-`Double` geometry functions (`FitAspectRatio`, `MarkerTopLeft`, `MarkerSitsOnAnchor`) but not its worksheet-level ones (`FindPinRow`'s hardcoded `_Edit` row window, incompatible with `_Snapshot`'s window); the photo-cache-over-clipboard decision for placing a connector page's photo; the metadata cell's exact address and why it must stay a plain hidden column rather than very-hidden; the leader-line rendering itself as the first time this codebase draws one, referencing `phase-2-manual-verification.md`'s 2b section where it was deferred from the live editor preview.

- [ ] **Step 3: Write the 3c doc**

Cover: the join-key lookup formula pattern (`LookupFormula`/`WireToFormula`), the nested-`IFERROR`-over-`MATCH` structure and why a single concatenated text key is used instead of a multi-criteria array formula; the specific asymmetry between direction-independent columns (Signal/Color/AWG/Length) and direction-dependent ones (Termination, Wire To); why blank rather than an error is the unwired-pin result, and how that connects to the spec's stated validation philosophy even though Check Drawing itself is still Phase 4.

- [ ] **Step 4: Write the 3d doc**

Cover: the differentiated page setup decision (landscape + repeating titles on `Harness` only) and its rationale as confirmed during 3d's discussion; why print area is computed from actual content (`LastUsedChartRow`, `LastUsedTableRow`) rather than the sheets' full fixed allotments; `CONN_PAGE_MIN_PRINT_ROW`'s purpose as a floor against a short pin table clipping a tall photo, and whatever value the manual-verification pass (Task 1) confirmed or adjusted it to.

- [ ] **Step 5: Write the 3e doc**

Cover: why `Connectors` must be reconstructed at all (no instance table is saved, per 3a/3b) and exactly how (`CONN_<RefDes>` sheet-name enumeration plus the metadata cell); the `LBound`-relative payload-indexing hazard from `docs/superpowers/plans/2026-08-28-ui-logic-separation-design.md` and how this plan's in-process call to `ReadConnector` is affected by it; the `EnableEvents` guard discipline across every bulk write into a live Creator sheet, and why `_Snapshot` needs none; the deliberate asymmetry between clearing stale photo shapes and leaving stale `_Snapshot` row data alone.

- [ ] **Step 6: Cross-check every doc against Task 1's deviation notes**

For each doc, confirm it describes what Task 1 found actually shipped (including any manual-verification adjustment), not what the corresponding plan said would be built.

- [ ] **Step 7: Commit**

```bash
git add docs/superpowers/specs/2026-09-01-phase-3a-harness-save-shell-design.md \
        docs/superpowers/specs/2026-09-01-phase-3b-connector-page-rendering-design.md \
        docs/superpowers/specs/2026-09-01-phase-3c-live-pin-tables-design.md \
        docs/superpowers/specs/2026-09-01-phase-3d-page-setup-design.md \
        docs/superpowers/specs/2026-09-01-phase-3e-harness-load-design.md
git commit -m "docs: add retrospective design docs for phase 3's five subsystems"
```

---

### Task 3: Student user guide update

**Files:**
- Modify: `docs/student-guide.md`
- Modify: `README.md`

**Interfaces:**
- Consumes: Task 1's reading pass.
- Produces: an updated guide covering Save, Save As, and Open Harness alongside the existing Phase 2 sections.

- [ ] **Step 1: Add the new sections**

In `docs/student-guide.md`, insert after the existing "Filling in the to-from chart" section (before "Managing your connector library"):

- **Saving your harness** - Home > Save Harness writes a plain `.xlsx` file with your chart, a page per connector showing its photo and numbered pins, and a table of wire data beside each photo. Save Harness As lets you choose where. Once saved, Save Harness writes to the same file again with no prompt.
- **Opening a saved harness** - Home > Open Harness replaces everything currently in the Creator with the chosen file's contents - your chart, your connectors, everything. State plainly (based on Task 1's findings) what happens if you pick a file that is not a saved harness.
- **What the saved file looks like** - the chart is still there and still editable by hand in plain Excel, with no macro warning; each connector's page updates automatically if you correct a wire's color, gauge, length, signal, or termination, but dragging a pin number's position on the printed page does not survive being saved again from the Creator.

Do not add a "Printing / exporting a PDF" or "Checking your work" section - Export PDF, Save Archive Copy, and Check Drawing do not exist yet (Phase 4).

- [ ] **Step 2: Update the README's Status section**

In `README.md`, replace the Status section's final paragraph:

```markdown
## Status

Phase 1 complete: build system, Creator shell, to-from chart with dependent
pin dropdowns. Phase 2 complete: connector library, editor, picker, snapshots,
ref des rename, library import/export. Phase 3 complete: harness Save/Save As/
Open, generated connector pages with photo callouts and live pin tables, and
page setup baked in at save time. See `docs/student-guide.md` for how to use
what's built so far, the `docs/superpowers/specs/` design docs for each
subsystem's decisions, and `docs/design/` for the layering rules the code
follows day to day. Phase 4 (Check Drawing, PDF export, archive copy) is
specified but not yet built.
```

- [ ] **Step 3: Verify the guide against the running application**

Open `dist/HarnessCreator.xlsm` and follow the new sections exactly as written - save a harness, open it back up, hand-edit a wire and confirm the connector page updates. Fix the guide, not the workbook, for any mismatch that is not a real bug; if it is a real bug, stop and report it rather than documenting around it.

- [ ] **Step 4: Commit**

```bash
git add docs/student-guide.md README.md
git commit -m "docs: add the student user guide sections for the phase 3 workflow"
```

---

## Self-Review

**Spec coverage.** Mirrors 2e's own answer: per-subsystem technical docs (Task 2) plus a student guide update (Task 3), sequenced strictly after 3a-3e's implementation and the manual-verification pass via Task 1's mandatory reading pass, rather than authored against plan text.

**Why this plan has no failing-test/passing-test steps like 3a-3e.** Same reasoning as 2e: the deliverable is prose, verified against the running application (Task 3 Step 3) and cross-checked against actual shipped code (Task 2 Step 6), not asserted by pytest.

**Why this plan explicitly depends on `phase-3-manual-verification.md` being run, not just on 3a-3e being merged.** Unlike 2e (which depended on 2a-2d plus the separately-scheduled UI/logic-separation plan), Phase 3's manual-verification doc is designed to accumulate adjustments across all five sub-plans before this plan starts - a design doc written before that pass runs risks describing a `CONN_PAGE_MIN_PRINT_ROW` value, a leader-line rendering detail, or a page-setup choice that the manual pass changes out from under it.

**No placeholders.** Every doc section above is scoped to specific, concrete content to verify and write, not "document the API" or "write user docs" left unspecified.
