## Layering Rules

### What an adapter (`.evt`) may do

An adapter handler does three things: read controls, call one layer 1
function, write controls. It may branch on the returned outcome to choose
what to display (e.g. which message to show). It may not branch on domain
state to decide what the operation is — that decision belongs to the layer
1 function it calls.

What stays in an adapter: control property assignment, dialog invocation
(`MsgBox`, `GetOpenFilename`, ...), workbook open and close, `Unload`, and
event-lifecycle state such as `shConnectors.mLastRefDes`.

A file path is an argument to a layer 1 function, never something the
function goes and asks the user for. Filesystem reads and writes are
permitted in layer 1; dialogs are not.

Workbook lifecycle stays in layer 2: the handler performs
`Workbooks.Open(...)` and passes the resulting `Worksheet` objects down to
the layer 1 action, matching how `modLibrary` and `modSnapshot` already take
sheets as parameters. This is what lets a test substitute its own
`library_wb` fixture instead of a real second workbook.

### What layer 1 (actions/queries) may do

Layer 1 holds two kinds of function:

- An **action** mutates state or renders a pass/fail judgement. It returns
  the three-element result envelope described below (e.g. `SaveFromEditor`,
  `PhotoClickAction`, `DeletePinRequest`, `CanLoadPhoto`).
- A **query** is a pure lookup with no failure mode worth naming. It returns
  a bare typed value, never a result envelope (e.g. `PinListItems`,
  `TypeListItems`, `PhotoFileFilter`). A query that finds nothing returns
  `Empty` or an empty string.

Layer 1 may reference layer 0, other layer 1 modules, and `Worksheet`,
`Range`, `Workbook`, or scalar types. It may never reference `MSForms.*`,
call `MsgBox`, `InputBox`, `GetOpenFilename`, `GetSaveAsFilename`, `.Show`,
or `Workbooks.Open`. Those are all UI or lifecycle concerns that belong in
an adapter.

A layer 1 parameter declares the type the control actually supplies (e.g.
`sPinCountText As String`, not `nPinCount As Long`), so parsing and
validation happen inside the tested function rather than in the adapter
before the call.

### What layer 0 (primitives) may do

Layer 0 modules operate on `Worksheet`, `Range`, and scalars only. They have
no knowledge of layer 1 or layer 2 and must not reference either.

### The result envelope

Every layer 1 **action** returns a three-element result built by
`modContract.Success(sOutcome, vPayload)` or `modContract.Failure(sOutcome,
vPayload)`. Adapters never index a result by hand; they call
`modContract.Ok(vResult)`, `modContract.Outcome(vResult)`, and
`modContract.Payload(vResult)`. The literals `(0)`, `(1)`, `(2)` for a
result's shape appear only inside `modContract` itself.

This exists because the codebase has two array conventions that disagree
about 0- vs. 1-basing across the `Application.Run` boundary — see the
[design spec](../superpowers/specs/2026-08-28-ui-logic-separation-design.md#the-envelope)
for the concrete example. `modContract.Success`/`Failure` build with `Array`,
which is zero-based on both the VBA and the pytest side, so `vResult(0)` and
`result[0]` are always the same element, provided no module ever declares
`Option Base 1`.

`modContract.PayloadKind` declares the payload type (`NONE`, `STRING`,
`LONG`, `DOUBLE`, `TABLE`) for every outcome code, and `Success`/`Failure`
raise if the payload doesn't match — turning a bad cast at the UI boundary
into a loud failure at construction time instead.

A `TABLE` payload, or the bare return of a table-shaped query, is either
`Empty` or a two-dimensional array with `LBound` of 1 on both dimensions —
never a zero-length array. Adapters call `modContract.TableRowCount` rather
than indexing `LBound`/`UBound` directly.

### User-visible text

All text a student reads is produced in layer 1, by `modMessages.MessageFor`
and `modMessages.MessageStyleFor`. A handler becomes:

```vba
MsgBox modMessages.MessageFor(vResult), modMessages.MessageStyleFor(vResult)
```

No string formatting or message assembly remains in any `.evt` file. The
same applies to list display strings (e.g. `"J1 - Deutsch DTM 4-way"`
returned by `PinListItems`/`ConnectorIndex`) — they are asserted in a test,
not assembled in a form.
