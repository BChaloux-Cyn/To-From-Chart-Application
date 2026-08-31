## Architecture Overview

### The problem

`Application.Run` — the entry point pytest uses to drive Excel over COM —
can only reach `Public` procedures in standard modules. It cannot reach a
form's or a sheet's document module. Those handlers also call `MsgBox`,
`InputBox`, `GetOpenFilename`, and `.Show`, each of which blocks a headless
test run regardless of reachability.

Before this layering existed, behaviour that lived only in `.evt` files
could not be called from a test at all, so tests asserted on VBA source
text instead (e.g. checking that a string like `lstPins.ListCount >=
nPinCount` appeared in a module). Those assertions are tautological:
renaming a variable fails a correct implementation, and inverting the
comparison passes a broken one.

The fix is to move decision-making out of forms and sheets into plain
standard modules that `Application.Run` can call directly, and confine
control manipulation and dialogs to a thin adapter layer.

### The three layers

| Layer | Modules | May call | May not call |
|---|---|---|---|
| 0 — primitives | `modUtil`, `modState`, `modLibrary`, `modChart`, `modConnectors`, `modSnapshot`, `modLibraryTransfer`, `modPinEditor` | `Worksheet`, `Range`, scalars | anything above |
| 1 — actions | `modEditorActions`, `modPickerActions`, `modManageActions`, `modContract`, `modMessages` | layer 0, each other, `Worksheet`, `Range`, `Workbook`, scalars | `MSForms.*`, `MsgBox`, `InputBox`, `GetOpenFilename`, `GetSaveAsFilename`, `.Show`, `Workbooks.Open` |
| 2 — adapters | `*.evt` for forms and sheets, `modConnectorUI` launchers | layer 1 only | layer 0 |

Layer 0 holds primitive operations on worksheets and ranges with no
knowledge of the UI. Layer 1 holds the actual decisions a user's click
represents, expressed as plain functions with typed inputs and a
predictable return shape. Layer 2 is disposable glue: read a control, call
one layer 1 function, write a control.

Because every decision lives in layer 1 and layer 1 has no UI dependency,
`Application.Run` can call it directly and a test can assert on the return
value instead of on source text.

See [`02_layering_rules.md`](02_layering_rules.md) for what each layer may
reference in detail, [`03_handler_lifecycle.md`](03_handler_lifecycle.md)
for the ordering rules that keep an adapter from losing data mid-action, and
[`04_enforcement.md`](04_enforcement.md) for how `tests/test_layering.py`
checks all of this automatically.
