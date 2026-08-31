## Library Import and Export

### Purpose

Moves one connector's full definition - record, pins, and photo - between
two library-schema workbooks, so a connector one student defines can be
merged into another's library.

### Import collisions are resolved by the UI, as overwrite-with-confirmation

`modLibraryTransfer.ImportConnector` (`src/vba/modLibraryTransfer.
bas:45-72`) writes under the source's own ConnectorID unconditionally,
overwriting whatever local row already has that ID, and replacing the
local pin set entirely (`modLibrary.DeletePinsForConnector` before
rewriting, so an overwrite with fewer pins doesn't leave the previous,
larger set's extra rows behind).

The collision decision itself lives in the UI, not in `ImportConnector`.
`frmManageLibrary.cmdImport_Click` (`src/vba/forms/frmManageLibrary.
evt:121-169`) checks `modLibrary.FindConnectorRow` for each connector in
the source file before ever calling `ImportOneConnector`, and on a
collision asks the student directly: *"A connector with Part Number
'\<ID\>' already exists in the library. Overwrite it with the imported
version?"* Answering No counts the connector as **Kept**; answering Yes
proceeds to import-as-overwrite. By the time `ImportConnector` runs, a
collision unambiguously means overwrite, decided by the person who owns
both libraries, rather than an automatic renaming scheme that would leave
two similarly-named entries to sort out later. `Origin` is always stamped
with the source file's name (`modLibraryTransfer.bas:53`), so an
overwritten connector's provenance is still visible afterward.

### Overwrite's ripple effect: cascading instance removal

`modManageActions.ImportOneConnector`
(`src/vba/modManageActions.bas:53-79`) wraps `ImportConnector` with a pin-
count check: if an overwritten connector's pin count changes, any harness
instance already placed against that ConnectorID would hold stale pin
references. `ImportOneConnector` compares the old and new `PinCount` and,
if they differ, calls `modConnectors.RemoveInstancesOfConnectorType` (the
same cascade `modManageActions.DeleteFromLibrary` uses for an outright
library deletion, `07_picker_and_snapshot.md`) to remove every affected
chart instance, folding the removed ref designators into a
`CONNECTOR_IMPORTED_CASCADED` outcome so the UI can report exactly what
was removed.

### Reporting through `modContract` and `modMessages`

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

### `Export Library`: exporting every connector at once

`frmManageLibrary`'s `cmdExportLibrary_Click` calls `modManageActions.
ExportLibraryToWorkbook` (`src/vba/modManageActions.bas:94-111`), which
iterates `modLibrary.ConnectorIndex` and calls the same per-connector
`modLibraryTransfer.ExportConnector` the single-connector export path
uses, reporting a `LIBRARY_EXPORTED` outcome carrying the count actually
exported. It exists as the natural companion to whole-library import: a
student sharing their whole library does not have to export each
connector one at a time.

### `CopyConnectorPhoto`'s clipboard mechanism and its reliability

`modLibraryTransfer.CopyConnectorPhoto`
(`src/vba/modLibraryTransfer.bas:4-15`) is the one place in this codebase
that moves an *already-embedded* picture shape between two workbooks: VBA
has no direct `Shape.Export`-to-another-sheet call, so the standard
technique is `Shape.Copy` on the source followed by `Worksheet.Paste` on
the destination, which goes through the Windows clipboard. Because that
copy can fail, the design accounts for extraction failure explicitly
rather than assuming it always succeeds: `CopyConnectorPhoto` returns
`False` on any error rather than raising, via a bare `On Error GoTo
Failed`.

**The actual outcome of manual verification:** `CopyConnectorPhoto`'s
`Shape.Copy`/`Paste` **works reliably** in this environment - confirmed
both visibly and headlessly, both via direct COM and through
`Application.Run`, run repeatedly with no failures in isolation. Under the
full pytest suite it showed one intermittent flake in eleven runs across
two sessions, consistent with a clipboard race rather than a hard
environment limitation, and the same conclusion held up under a second,
later, more thorough verification pass (35 checks, all passing) that
exercised the export-then-import round trip and the deliberate
extraction-failure fallback prompt with no clipboard failure observed.
**This makes `CopyConnectorPhoto` the common, working path for photo
transfer in practice, not a rare fallback** - a materially different
reliability profile from its sibling mechanism, `modLibrary.
ExportShapeToFile` (`05_library_core.md`), which uses a different
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
(`src/vba/forms/frmManageLibrary.evt:183-194`), so a failed photo copy
never silently leaves a connector without a photo.
