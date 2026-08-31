# Wire Harness Documentation Creator

An Excel-based tool for students learning electrical documentation. Fill in a
to-from chart, attach photographs of the connectors, number the positions on
those photos, and get a print-ready harness drawing.

## Artifacts

| File | What it is |
|---|---|
| `dist/HarnessCreator.xlsm` | The editor. Build artifact - never edit by hand. |
| `dist/ConnectorLibrary.xlsx` | The reusable connector library. (Phase 2.) |
| `<harness>.xlsx` | A saved drawing. Macro-free and print-ready. (Phase 3.) |

## Developer setup

Requires Windows and desktop Excel.

    python -m pip install -r requirements.txt

Enable Excel's programmatic VBA access once, or the build cannot inject code:
File > Options > Trust Center > Trust Center Settings > Macro Settings >
"Trust access to the VBA project object model". Per-user and reversible; it is
not required on student machines.

Verify:

    python build/build.py --check

## Build and test

    python build/build.py
    python -m pytest -v

The `.xlsm` is generated from `src/vba/*.bas` and `build/layout.py`. To change
behaviour, change the source and rebuild. Edits made directly in the workbook
are lost on the next build.

## Layout

| Path | Responsibility |
|---|---|
| `build/excel_com.py` | COM lifecycle, VBA injection, prerequisite check |
| `build/layout.py` | Sheet layout constants and builders |
| `build/build.py` | Build orchestration and CLI |
| `src/vba/mod{Contract,Messages}.bas` | Result envelope and user-visible text |
| `src/vba/mod*Actions.bas` | User-intent transactions; the only thing forms may call |
| `src/vba/` | Layer 0 primitives; `sheets/*.evt` and `forms/*.evt` are UI adapters |
| `tests/` | pytest suites driving Excel COM |
| `docs/design/` | Current-state design documentation: layering strategy plus every subsystem |
| `docs/superpowers/specs/` | Original design spec |
| `docs/superpowers/plans/` | Implementation plans |

## Status

Phase 1 complete: build system, Creator shell, to-from chart with dependent
pin dropdowns. Phase 2 complete: connector library file and reader/writer,
the connector editor with click-to-place pin markers, the connector picker,
per-harness connector snapshots, ref des rename, and library import/export.
See `docs/user-guide/user-guide.md` for how to use what's built so far, and
`docs/design/` for how the Creator is built - each subsystem's decisions
alongside the layering rules the code follows day to day. Phases 3 and 4
(harness save/load with rendered connector pages, validation and export)
are specified but not yet built.
