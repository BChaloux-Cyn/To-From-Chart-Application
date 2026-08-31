# Phase 2d: Library Import and Export - Design

Date: 2026-08-26
Status: Reflects merged code as of Phase 2e

## Purpose

Moves one connector's full definition - record, pins, and photo - between
two library-schema workbooks, so a connector one student defines can be
merged into another's library.

## Collision-safe rename became overwrite-with-confirmation

The original 2d plan had `modLibraryTransfer.ImportConnector` call
`modLibrary.UniqueConnectorID` and silently rename a colliding import to
`<ID>-2`, keeping the original untouched. The shipped
`ImportConnector` (`src/vba/modLibraryTransfer.bas:45-72`) does the
opposite: it writes under the source's own ConnectorID unconditionally,
overwriting whatever local row already has that ID, and replacing the
local pin set entirely (`modLibrary.DeletePinsForConnector` before
rewriting, so an overwrite with fewer pins doesn't leave the previous,
larger set's extra rows behind).

The collision decision moved up a layer, to the UI. `frmManageLibrary.
cmdImport_Click` (`src/vba/forms/frmManageLibrary.evt:121-169`) checks
`modLibrary.FindConnectorRow` for each connector in the source file before
ever calling `ImportOneConnector`, and on a collision asks the student
directly: *"A connector with Part Number '\<ID\>' already exists in the
library. Overwrite it with the imported version?"* Answering No counts the
connector as **Kept**; answering Yes proceeds to import-as-overwrite. This
is why `ImportConnector` itself no longer needs to rename anything - by the
time it runs, a collision unambiguously means overwrite, decided by the
person who owns both libraries, not an automatic renaming scheme that could
leave two similarly-named entries a student has to sort out later.
`Origin` is always stamped with the source file's name
(`modLibraryTransfer.bas:53`), so an overwritten connector's provenance is
still visible afterward.

## Overwrite's ripple effect: cascading instance removal

`modManageActions.ImportOneConnector`
(`src/vba/modManageActions.bas:53-79`) wraps `ImportConnector` with a check
the original plan didn't need, because the original plan never overwrote
in place: if an overwritten connector's pin count changes, any harness
instance already placed against that ConnectorID would hold stale pin
references. `ImportOneConnector` compares the old and new `PinCount` and,
if they differ, calls `modConnectors.RemoveInstancesOfConnectorType` (the
same cascade `modManageActions.DeleteFromLibrary` uses for an outright
library deletion, 2c doc) to remove every affected chart instance, folding
the removed ref designators into a `CONNECTOR_IMPORTED_CASCADED` outcome so
the UI can report exactly what was removed.

## Reporting through `modContract` and `modMessages`

Every user-facing outcome of an import or export - `EXPORTED`,
`EXPORT_FAILED`, `CONNECTOR_IMPORTED`, `CONNECTOR_IMPORTED_CASCADED`,
`CONNECTOR_DELETED`, `CONNECTOR_DELETED_CASCADED`, `LIBRARY_EXPORTED` - is a
`modContract` result built in `modManageActions`
(`src/vba/modManageActions.bas`), with display text produced by
`modMessages.MessageFor`/`ImportSummaryMessage`
(`src/vba/modMessages.bas`), matching
`docs/design/02_layering_rules.md`'s user-visible-text rule. A whole-import
run reports a single summary - imported/kept/overwritten counts plus any
cascaded removals - via `modMessages.ImportSummaryMessage`
(`src/vba/forms/frmManageLibrary.evt:164-168`), rather than one message box
per connector in the source file.

## `Export Library`: exporting every connector at once

Not present in the original 2d plan at all, which scoped export to one
connector selected in the list (its own self-review explicitly named
multi-connector export as deliberately out of scope: "the spec's own
phrasing is singular"). The shipped `frmManageLibrary` adds a fourth
button, `cmdExportLibrary_Click`, calling
`modManageActions.ExportLibraryToWorkbook`
(`src/vba/modManageActions.bas:94-111`), which iterates
`modLibrary.ConnectorIndex` and calls the same per-connector
`modLibraryTransfer.ExportConnector` the single-connector path uses,
reporting a `LIBRARY_EXPORTED` outcome carrying the count actually
exported.

## `CopyConnectorPhoto`'s clipboard mechanism and its reliability

`modLibraryTransfer.CopyConnectorPhoto`
(`src/vba/modLibraryTransfer.bas:4-15`) is the one place in this codebase
that moves an *already-embedded* picture shape between two workbooks: VBA
has no direct `Shape.Export`-to-another-sheet call, so the standard
technique is `Shape.Copy` on the source followed by `Worksheet.Paste` on
the destination, which goes through the Windows clipboard exactly as the
spec's "Photo cache" section anticipated could fail ("If extraction fails,
the editor prompts for the image file rather than failing outright"). It
returns `False` on any error rather than raising, via a bare `On Error
GoTo Failed`.

**The actual outcome of manual verification** (`docs/superpowers/plans/
phase-2-manual-verification.md`, 2d Task 1): `CopyConnectorPhoto`'s
`Shape.Copy`/`Paste` **works reliably** in this environment - confirmed
both visibly and headlessly, both via direct COM and through
`Application.Run`, run repeatedly with no failures in isolation. Under the
full pytest suite it showed one intermittent flake in eleven runs across
two sessions, consistent with a clipboard race rather than a hard
environment limitation, and the plan's own guidance (verify by hand,
don't chase further if it's not a hard failure) was followed rather than
building a workaround. The later, more thorough
`docs/superpowers/plans/2026-08-28-ui-separation-manual-verification.md`
pass (2026-08-31, 35 checks, all passing) exercised the export-then-import
round trip (#22-23) and the deliberate extraction-failure fallback prompt
(#24) with no clipboard failure observed. **This makes `CopyConnectorPhoto`
the common, working path for photo transfer in practice, not a rare
fallback** - a materially different reliability profile from its sibling
mechanism, `modLibrary.ExportShapeToFile` (2a doc), which uses a different
technique (`Shape.Copy` into a throwaway `ChartObject`, then
`Chart.Export`, needed specifically to rasterize a shape on a *very hidden*
sheet to a file) and was confirmed unreliable on this machine regardless of
automation versus a real interactive session. The two are easy to conflate
- both are Shape-Copy-based and both exist because of the same underlying
gap (no direct "export a Shape" API) - but they are separate code paths
with separate, independently-verified reliability outcomes.

`ImportedPhotoOk` (`src/vba/modManageActions.bas:81-90`) does not trust
`CopyConnectorPhoto`'s own return value for the fallback-prompt decision;
it checks whether the shape it should have produced actually exists in the
destination `Photos` sheet. `cmdImport_Click` prompts for a replacement
image file only when that check fails
(`src/vba/forms/frmManageLibrary.evt:183-194`), matching the spec's
"prompts for the image file rather than failing outright" line verbatim.

## Summary of deviations from the 2d plan

| Plan said | Code does | Why |
|---|---|---|
| `ImportConnector` renames a colliding ID via `UniqueConnectorID`, keeping the original untouched | Overwrites the colliding row in place under the source's own ID; renaming was dropped entirely | Collision handling moved to the UI as an explicit Keep/Overwrite prompt - a deliberate choice a student makes, not a silent auto-rename that leaves two similarly-named entries to sort out later |
| No cascade behavior on import | An overwrite that changes `PinCount` removes every stale harness instance of that ConnectorID (`CONNECTOR_IMPORTED_CASCADED`) | An overwrite-in-place, unlike a rename, can invalidate pin references already placed on the chart |
| Export is single-connector only, by design (explicitly out of scope in the plan's self-review) | `Export Library` added, exporting every library connector into one shared workbook | Natural companion to whole-library import; reuses the same per-connector `ExportConnector` |
| No explicit statement of the clipboard mechanism's real-world reliability | `CopyConnectorPhoto` confirmed reliable (rare pytest-session flake only); documented as the common path, distinct from the separately-confirmed-unreliable `ExportShapeToFile` | Required by this doc's own writing task, and needed to correctly scope the "known limitation" the spec anticipated - it applies to `ExportShapeToFile`'s use cases (2b/2c preview and snapshot backfill), not to import/export |
