## Purpose and Scope

This directory is the comprehensive, current-state design documentation for
the Wire Harness Creator: what each part of the codebase does, how it is
built, and why it works the way it does. It covers the connector library
and every subsystem built on it, and the cross-cutting layering strategy
that constrains all of them. Where a claim here is testable, it is anchored
to a specific file, line, and test - not just described in prose.

### Contents

| Doc | Covers |
|---|---|
| [`01_architecture_overview.md`](01_architecture_overview.md) | Why the layer 0/1/2 split exists and what each layer may or may not do |
| [`02_layering_rules.md`](02_layering_rules.md) | The rules for each layer in detail, and the `modContract` result envelope |
| [`03_handler_lifecycle.md`](03_handler_lifecycle.md) | Ordering rules that keep an adapter from losing data mid-action |
| [`04_enforcement.md`](04_enforcement.md) | How `tests/test_layering.py` checks the rules above automatically |
| [`05_library_core.md`](05_library_core.md) | The connector library's three-table schema, the bounded-window storage convention, and photo caching |
| [`06_connector_editor.md`](06_connector_editor.md) | The connector editor: click-to-place pins, anchor vs. marker, the `clsPinMarker` drag mechanism |
| [`07_picker_and_snapshot.md`](07_picker_and_snapshot.md) | Add Connector, Manage Library, Remove Connector, `_Snapshot`, and ref des rename |
| [`08_import_export.md`](08_import_export.md) | Library import/export between students, and the clipboard-dependent photo copy |

This directory does not cover the build system itself (`build/` and the
sheet layout constants in `build/layout.py`; see the top-level `README.md`
for those) or how to use the built application (`docs/user-guide/
user-guide.md`'s job).
