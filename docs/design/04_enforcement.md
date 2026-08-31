## Enforcement

The rules in this document are not just convention — `tests/test_layering.py`
checks them against the built workbook's VBA source on every test run, by
inspecting each module's `CodeModule` over COM.

| Check | What it catches |
|---|---|
| `test_action_modules_open_no_dialogs` | A layer 1 module (`modContract`, `modMessages`, `modEditorActions`, `modPickerActions`, `modManageActions`) referencing `MsgBox`, `InputBox`, `GetOpenFilename`, `GetSaveAsFilename`, `.Show`, `Unload`, `Workbooks.Open`, `DoEvents`, or `MSForms` |
| `test_adapters_call_only_permitted_layer0_members` | An adapter (`.evt` or `clsPinMarker`) calling a layer 0 member directly instead of going through an action module. A short, named allowlist (`ALLOWED_LAYER0_IN_ADAPTERS` in the test file) covers the few primitives with no real transaction to lift, each with a one-line reason in the test's comments |
| `test_no_doevents_anywhere` | `DoEvents` anywhere in the codebase — see [`03_handler_lifecycle.md`](03_handler_lifecycle.md) for why this is unconditionally forbidden, not just in layer 1 |
| `test_nothing_follows_unload_me` | Any adapter `Sub` that reads a form-level variable (`m[A-Z]...`) or a control (`txt`/`lst`/`cbo`/`cmd`/`img`/`tgl` followed by `.`) after `Unload Me` |
| `test_every_click_handler_delegates` | A `cmdXxx_Click` handler on a form that does domain work without calling into `modEditorActions`, `modPickerActions`, or `modManageActions`. `NON_DELEGATING_HANDLERS` in the test file lists the handlers that legitimately do nothing but unload or hand off to another form |
| `test_no_option_base_directive` | Any module declaring `Option Base 1`, which would flip the zero-basing the [result envelope](02_layering_rules.md#the-result-envelope) depends on |

### Adding new code

When adding a new adapter call into layer 0, or a new form handler, running
`pytest tests/test_layering.py -v` will fail loudly and name the exact
module and reference that violates a rule, before it becomes an untestable
behaviour buried in a `.evt` file. If a violation is a deliberate, narrow
exception (as the entries in `ALLOWED_LAYER0_IN_ADAPTERS` are), add it to
the relevant allowlist in the test file with a comment explaining why no
action-module transaction applies — do not weaken the check itself.
