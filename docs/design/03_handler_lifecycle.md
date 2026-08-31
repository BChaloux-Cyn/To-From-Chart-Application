## Handler Lifecycle Discipline

An action must not have its inputs destroyed while it runs. Two properties
of VBA make that safe by default, and both need protecting by convention
rather than the compiler.

### Why this matters

VBA is single-threaded. While a handler executes, no other form event can
fire, including `UserForm_QueryClose` from the system Close button — so a
form cannot unload itself out from under a running action. The one hole in
that guarantee is `DoEvents`, which re-enters the message pump and lets a
queued Close click fire `QueryClose` mid-action. In `frmConnectorPicker` and
`frmManageLibrary`, that would close `mLibrary` and invalidate the very
`Worksheet` objects the action holds.

Arguments are fully evaluated before a call — `Array(Trim$(txtName.Text), ...)` 
is materialized into a `Variant` array before the callee receives it,
and `String`/`Variant` values are copied — so an action holds values rather
than live references into the form.

The layering rules in [`02_layering_rules.md`](02_layering_rules.md)
protect the callee's side for free: layer 1 cannot reference `MSForms.*` and
cannot call `.Show` or `Unload`, so an action can never unload the form that
called it. Ordering within the handler itself is not covered by that, so
three rules apply there:

### The three rules

1. **Capture then act.** A handler reads every control value and
   form-level variable it needs into locals *before* the action call.
   Nothing after the call touches the form.
2. **`Unload Me` is the final statement in its branch.** No line after it
   may reference a control or a form-level variable.
3. **No `DoEvents`** in layer 1, and none in an adapter between capturing
   state and unloading.

### A worked exception

`frmConnectorPicker.cmdNew_Click` reads
`modConnectorUI.LastSavedConnectorID` *after* `frmConnectorEditor` has
already unloaded itself. That's safe only because the value lives in a
standard module — form state would have been gone by then.
`src/vba/modConnectorUI.bas:8` records that this was the original fix for
exactly this failure, and the mechanism is kept unchanged rather than
"cleaned up" into form state.
