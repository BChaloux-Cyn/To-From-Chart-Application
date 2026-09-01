# Phase 3c: Live Pin-Table Formulas Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fill in the six formula columns 3b's pin-table skeleton reserved (`Wire To`, `Signal`, `Color`, `AWG`, `Termination`, `Length`) with INDEX/MATCH formulas against 3a's join-key columns, so a student who hand-edits a saved harness's chart sees every connector page's pin table update with no macro involved.

**Architecture:** Each pin-table row already knows its own pin number (3b wrote it as a static value) and its page's Ref Des (baked into the formula text at render time, since a page never changes which connector it renders without a full Save re-render). A formula looks up the chart row whose `FromConn|FromPin` join key equals `"<RefDes>|<PinNumber>"`; if none matches, it falls back to `ToConn|ToPin`. `Signal`/`Color`/`AWG`/`Length` read the same physical chart column regardless of which side matched (a wire has exactly one signal, color, gauge, and length); `Termination` and `Wire To` read the From-side or To-side column depending on which match succeeded, matching the spec's "written from this connector's point of view" instruction precisely.

**Tech Stack:** VBA (Excel 16.0 COM automation), Python 3.13/pywin32/pytest.

**Spec:** `docs/superpowers/specs/2026-08-26-harness-creator-design.md`

**Depends on:** 3a and 3b, implemented and merged - this plan writes into the exact columns 3a's `Harness` sheet and 3b's pin-table skeleton established.

**Part of Phase 3** (see 3a's header for the full breakdown). This is 3c.

## Global Constraints

- INDEX/MATCH only, never XLOOKUP - the spec requires the saved file to open correctly on Excel 2016.
- Formulas must be exact-match (`MATCH(..., 0)`), never approximate - a wire's join key is an exact string, not a sorted range.
- **Column map, fixed by 3a and 3b, restated here since every formula in this plan references it by letter:**

  | Sheet | Column | Meaning |
  |---|---|---|
  | `Harness` | A | From Conn |
  | `Harness` | B | From Pin |
  | `Harness` | C | From Term |
  | `Harness` | D | Signal |
  | `Harness` | E | Color |
  | `Harness` | F | AWG |
  | `Harness` | G | Length |
  | `Harness` | H | To Term |
  | `Harness` | I | To Conn |
  | `Harness` | J | To Pin |
  | `Harness` | L (hidden) | `FromConn\|FromPin` join key |
  | `Harness` | M (hidden) | `ToConn\|ToPin` join key |
  | `CONN_<RefDes>` | J | Pin (static) |
  | `CONN_<RefDes>` | K | Label (static) |
  | `CONN_<RefDes>` | L | Wire To (this plan) |
  | `CONN_<RefDes>` | M | Signal (this plan) |
  | `CONN_<RefDes>` | N | Color (this plan) |
  | `CONN_<RefDes>` | O | AWG (this plan) |
  | `CONN_<RefDes>` | P | Termination (this plan) |
  | `CONN_<RefDes>` | Q | Length (this plan) |

- The chart's row range on both sheets is fixed at 7-1006 (`modHarnessBuild.SAVED_CHART_FIRST_ROW`/`LAST_ROW`) - every formula's `MATCH` range is exactly `$7:$1006`, matching 3a's decision to keep the full range rather than trim to content.

## File Structure

| File | Responsibility |
|---|---|
| `src/vba/modConnectorPage.bas` | Modified: adds the formula-text builders and `WriteLiveFormulas`. |
| `tests/test_connector_page.py` | Additions: exact formula text for each builder. |
| `tests/test_harness_save_integration.py` | Modified: the round-trip test evaluates pin-table cells (not just formula text) against real chart data, including a From-match row, a To-match row, and an unwired pin. |

---

### Task 1: Formula-text builders

**Files:**
- Modify: `src/vba/modConnectorPage.bas`
- Test: `tests/test_connector_page.py` (additions)

**Interfaces:**
- Consumes: nothing new.
- Produces:
  - VBA `modConnectorPage.LookupFormula(ByVal sFromCol As String, ByVal sToCol As String, ByVal sRefDes As String, ByVal nTableRow As Long) As String` - the shared two-level `IFERROR(INDEX(...,MATCH(...From...)), IFERROR(INDEX(...,MATCH(...To...)), ""))` pattern, used for `Signal`/`Color`/`AWG`/`Length` (where `sFromCol = sToCol`) and `Termination` (where they differ).
  - VBA `modConnectorPage.WireToFormula(ByVal sRefDes As String, ByVal nTableRow As Long) As String` - the compound `"<RefDes>-<Pin>"` lookup for the opposite endpoint.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_connector_page.py`:

```python
def test_lookup_formula_for_a_direction_independent_column(wb):
    formula = run(wb, "modConnectorPage.LookupFormula", "D", "D", "J1", 2)
    assert formula == (
        '=IFERROR(INDEX(Harness!$D$7:$D$1006,MATCH("J1|"&$J2,Harness!$L$7:$L$1006,0)),'
        'IFERROR(INDEX(Harness!$D$7:$D$1006,MATCH("J1|"&$J2,Harness!$M$7:$M$1006,0)),""))'
    )


def test_lookup_formula_for_termination_uses_different_from_and_to_columns(wb):
    formula = run(wb, "modConnectorPage.LookupFormula", "C", "H", "J1", 3)
    assert formula == (
        '=IFERROR(INDEX(Harness!$C$7:$C$1006,MATCH("J1|"&$J3,Harness!$L$7:$L$1006,0)),'
        'IFERROR(INDEX(Harness!$H$7:$H$1006,MATCH("J1|"&$J3,Harness!$M$7:$M$1006,0)),""))'
    )


def test_wire_to_formula_combines_ref_des_and_pin(wb):
    formula = run(wb, "modConnectorPage.WireToFormula", "J1", 2)
    assert formula == (
        '=IFERROR(INDEX(Harness!$I$7:$I$1006,MATCH("J1|"&$J2,Harness!$L$7:$L$1006,0))&"-"&'
        'INDEX(Harness!$J$7:$J$1006,MATCH("J1|"&$J2,Harness!$L$7:$L$1006,0)),'
        'IFERROR(INDEX(Harness!$A$7:$A$1006,MATCH("J1|"&$J2,Harness!$M$7:$M$1006,0))&"-"&'
        'INDEX(Harness!$B$7:$B$1006,MATCH("J1|"&$J2,Harness!$M$7:$M$1006,0)),""))'
    )
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_connector_page.py -v -k "lookup_formula or wire_to_formula"
```

Expected: FAIL - neither function exists.

- [ ] **Step 3: Write the functions**

Append to `src/vba/modConnectorPage.bas`:

```vb
Private Function KeyExpr(ByVal sRefDes As String, ByVal nTableRow As Long) As String
    KeyExpr = """" & sRefDes & "|""&$J" & CStr(nTableRow)
End Function

Public Function LookupFormula(ByVal sFromCol As String, ByVal sToCol As String, _
                              ByVal sRefDes As String, ByVal nTableRow As Long) As String
    Dim sKey As String
    sKey = KeyExpr(sRefDes, nTableRow)
    LookupFormula = "=IFERROR(INDEX(Harness!$" & sFromCol & "$7:$" & sFromCol & "$1006," & _
        "MATCH(" & sKey & ",Harness!$L$7:$L$1006,0))," & _
        "IFERROR(INDEX(Harness!$" & sToCol & "$7:$" & sToCol & "$1006," & _
        "MATCH(" & sKey & ",Harness!$M$7:$M$1006,0)),""""))"
End Function

Public Function WireToFormula(ByVal sRefDes As String, ByVal nTableRow As Long) As String
    Dim sKey As String
    sKey = KeyExpr(sRefDes, nTableRow)
    WireToFormula = "=IFERROR(INDEX(Harness!$I$7:$I$1006,MATCH(" & sKey & ",Harness!$L$7:$L$1006,0))&""-""&" & _
        "INDEX(Harness!$J$7:$J$1006,MATCH(" & sKey & ",Harness!$L$7:$L$1006,0))," & _
        "IFERROR(INDEX(Harness!$A$7:$A$1006,MATCH(" & sKey & ",Harness!$M$7:$M$1006,0))&""-""&" & _
        "INDEX(Harness!$B$7:$B$1006,MATCH(" & sKey & ",Harness!$M$7:$M$1006,0)),""""))"
End Function
```

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_connector_page.py -v -k "lookup_formula or wire_to_formula"
```

Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add src/vba/modConnectorPage.bas tests/test_connector_page.py
git commit -m "feat: add the pin-table lookup formula builders"
```

---

### Task 2: Write the formulas into every pin-table row

**Files:**
- Modify: `src/vba/modConnectorPage.bas`
- Modify: `src/vba/modHarnessBuild.bas`
- Test: `tests/test_connector_page.py` (additions)

**Interfaces:**
- Consumes: `LookupFormula`/`WireToFormula` (Task 1), `CONN_TABLE_FIRST_ROW`/`CONN_TABLE_FIRST_COL` (3b).
- Produces: VBA `modConnectorPage.WriteLiveFormulas(wsPage As Worksheet, ByVal sRefDes As String, vPins As Variant)` - writes columns L-Q (Wire To, Signal, Color, AWG, Termination, Length) for every row `WriteTableSkeleton` (3b) populated, in the same row order.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_connector_page.py`:

```python
def test_write_live_formulas_fills_every_row(wb, app):
    dest = app.Workbooks.Add()
    try:
        ws = dest.Worksheets(1)
        pins = (
            ("DTM-04P", 1, "+12V", 0.1, 0.1, 0.1, 0.1),
            ("DTM-04P", 2, "GND", 0.9, 0.1, 0.9, 0.1),
        )
        run(wb, "modConnectorPage.WriteTableSkeleton", ws, pins)
        run(wb, "modConnectorPage.WriteLiveFormulas", ws, "J1", pins)

        assert ws.Cells(2, 12).Formula == run(wb, "modConnectorPage.WireToFormula", "J1", 2)
        assert ws.Cells(2, 13).Formula == run(wb, "modConnectorPage.LookupFormula", "D", "D", "J1", 2)
        assert ws.Cells(2, 14).Formula == run(wb, "modConnectorPage.LookupFormula", "E", "E", "J1", 2)
        assert ws.Cells(2, 15).Formula == run(wb, "modConnectorPage.LookupFormula", "F", "F", "J1", 2)
        assert ws.Cells(2, 16).Formula == run(wb, "modConnectorPage.LookupFormula", "C", "H", "J1", 2)
        assert ws.Cells(2, 17).Formula == run(wb, "modConnectorPage.LookupFormula", "G", "G", "J1", 2)
        assert ws.Cells(3, 12).Formula == run(wb, "modConnectorPage.WireToFormula", "J1", 3)
    finally:
        dest.Close(SaveChanges=False)
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
python -m pytest tests/test_connector_page.py -v -k write_live_formulas
```

Expected: FAIL - `WriteLiveFormulas` does not exist.

- [ ] **Step 3: Write the function**

Append to `src/vba/modConnectorPage.bas`:

```vb
Public Sub WriteLiveFormulas(wsPage As Worksheet, ByVal sRefDes As String, vPins As Variant)
    Dim i As Long, r As Long

    If IsEmpty(vPins) Then Exit Sub

    For i = LBound(vPins, 1) To UBound(vPins, 1)
        r = CONN_TABLE_FIRST_ROW + (i - LBound(vPins, 1))

        wsPage.Cells(r, CONN_TABLE_FIRST_COL + 2).Formula = WireToFormula(sRefDes, r)
        wsPage.Cells(r, CONN_TABLE_FIRST_COL + 3).Formula = LookupFormula("D", "D", sRefDes, r)
        wsPage.Cells(r, CONN_TABLE_FIRST_COL + 4).Formula = LookupFormula("E", "E", sRefDes, r)
        wsPage.Cells(r, CONN_TABLE_FIRST_COL + 5).Formula = LookupFormula("F", "F", sRefDes, r)
        wsPage.Cells(r, CONN_TABLE_FIRST_COL + 6).Formula = LookupFormula("C", "H", sRefDes, r)
        wsPage.Cells(r, CONN_TABLE_FIRST_COL + 7).Formula = LookupFormula("G", "G", sRefDes, r)
    Next i
End Sub
```

- [ ] **Step 4: Wire it into BuildConnectorPages**

In `src/vba/modHarnessBuild.bas`'s `BuildConnectorPages`, add one line after `modConnectorPage.WriteTableSkeleton wsPage, vPins`:

```vb
        modConnectorPage.WriteTableSkeleton wsPage, vPins
        modConnectorPage.WriteLiveFormulas wsPage, sRefDes, vPins
```

- [ ] **Step 5: Run the tests to verify they pass**

Run:

```bash
python -m pytest tests/test_connector_page.py -v -k write_live_formulas
```

Expected: 1 passed.

- [ ] **Step 6: Commit**

```bash
git add src/vba/modConnectorPage.bas src/vba/modHarnessBuild.bas tests/test_connector_page.py
git commit -m "feat: write live pin-table formulas into every connector page"
```

---

### Task 3: Prove the formulas evaluate correctly

**Files:**
- Modify: `tests/test_harness_save_integration.py`

**Interfaces:**
- Consumes: everything from Tasks 1-2, plus the existing `test_full_harness_round_trips_through_a_saved_file` fixture (3a/3b).

This is where the plan stops trusting formula *text* and checks formula *results* - the spec's own testing section calls for exactly this ("assert that pin-table formulas evaluate to the expected wire data").

- [ ] **Step 1: Extend the fixture with a second connector and a two-sided wire**

In `tests/test_harness_save_integration.py`, extend the existing test's `Harness`/`Connectors` setup so the wire has a real endpoint on both sides and a second, unwired pin exists to prove the blank case:

```python
    wsConn.Cells(3, 1).Value = "J2"
    wsConn.Cells(3, 2).Value = "DTM-04P"
    wsConn.Cells(3, 3).Value = "Deutsch DTM 4-way"
    wsConn.Cells(3, 5).Value = "Connector"
    wsConn.Cells(3, 6).Value = 4

    run(wb, "modLibrary.WritePin", wsSnap, 211, 2210, ("DTM-04P", 2, "GND", 0.9, 0.1, 0.9, 0.1))
```

(Insert alongside the existing `J1`/pin-1 setup from 3a/3b, before `dest = app.Workbooks.Add()`. The existing chart row already wires `J1` pin 1 to `J2` pin 1; `J1` and `J2` both now have a defined pin 2 that no chart row references.)

- [ ] **Step 2: Add evaluated-value assertions**

Add to the `reopened` block, after the existing `CONN_J1` assertions:

```python
        j1 = reopened.Worksheets("CONN_J1")
        assert j1.Cells(2, 13).Value == "12V_SW"   # Signal, matched as From
        assert j1.Cells(2, 14).Value == "Red"       # Color
        assert j1.Cells(2, 15).Value == "18"        # AWG
        assert j1.Cells(2, 17).Value == 24           # Length
        assert j1.Cells(2, 12).Value == "J2-1"      # Wire To, matched as From
        assert j1.Cells(3, 12).Value == ""          # pin 2, unwired: blank, not an error

        j2 = reopened.Worksheets("CONN_J2")
        assert j2.Cells(2, 13).Value == "12V_SW"    # same wire, matched as To this time
        assert j2.Cells(2, 12).Value == "J1-1"      # Wire To points back at J1
        assert j2.Cells(2, 16).Value == j1.Cells(2, 16).Value or True  # Termination differs by side; see step 3
```

Replace that last placeholder line with the real assertion once the fixture's From Term / To Term values are known (Step 3 below fills them in explicitly, so this line should not survive into the committed test as written here).

- [ ] **Step 3: Give the fixture explicit, distinct From/To terminations and assert both sides**

Ensure the existing chart row (row 7) sets both termination columns distinctly, e.g.:

```python
    wsHarness.Cells(7, 3).Value = "Crimp Pin"   # From Term
    wsHarness.Cells(7, 8).Value = "Ring Terminal"  # To Term
```

Then replace Step 2's placeholder line with:

```python
        assert j1.Cells(2, 16).Value == "Crimp Pin"       # From side sees From Term
        assert j2.Cells(2, 16).Value == "Ring Terminal"   # To side sees To Term
```

- [ ] **Step 4: Run the test**

Run:

```bash
python -m pytest tests/test_harness_save_integration.py -v
```

Expected: 1 passed. This is the test most likely to expose a formula-logic mistake (the plan's own formula text was never executed by Excel before now) - if it fails, fix `LookupFormula`/`WireToFormula` in `modConnectorPage.bas` per Tasks 1-2 rather than adjusting the expected values to match a wrong result.

- [ ] **Step 5: Run the whole suite from a clean build**

Run:

```bash
rm -rf dist
python -m pytest -v
```

Expected: everything passes.

- [ ] **Step 6: Commit**

```bash
git add tests/test_harness_save_integration.py
git commit -m "test: prove connector-page pin tables evaluate correctly for both wire directions"
```

---

### Task 4: Prove a hand edit updates the pin table with no macro

**Files:**
- Modify: `tests/test_harness_save_integration.py`

**Interfaces:**
- Consumes: the saved, reopened file from Task 3's test.

This is the behavior the spec calls out by name: "a student who opens a saved harness and corrects a length, color, gauge, termination, or signal sees the connector pages update... with no macros involved." The saved file has no VBA project at all (asserted already by `test_full_harness_round_trips_through_a_saved_file`), so if this passes, it is Excel's own formula recalculation doing the work, not anything this codebase runs.

- [ ] **Step 1: Add the hand-edit assertion**

Append inside the `reopened` block, after Task 3's assertions:

```python
        wsHarnessReopened = reopened.Worksheets("Harness")
        wsHarnessReopened.Cells(7, 5).Value = "Blue"  # hand-edit Color on the open, macro-free file
        assert j1.Cells(2, 14).Value == "Blue"
        assert j2.Cells(2, 14).Value == "Blue"
```

- [ ] **Step 2: Run the test**

Run:

```bash
python -m pytest tests/test_harness_save_integration.py -v
```

Expected: 1 passed.

- [ ] **Step 3: Commit**

```bash
git add tests/test_harness_save_integration.py
git commit -m "test: prove a hand edit to a saved harness updates its connector pages live"
```

---

## Self-Review

**Spec coverage for this sub-plan.** "Pin tables as live formulas" and "the table is written from this connector's point of view" (Wire To is the opposite endpoint, Termination is this connector's own terminal) are both directly implemented and directly tested (Task 3), including the specific asymmetry the spec calls out - the same wire's Termination differs by which side is rendering it, while Signal/Color/AWG/Length do not. Task 4 covers "a student who opens a saved harness and corrects... sees the connector pages update... with no macros involved" word for word.

**Why `MATCH(..., 0)` against a hidden text column rather than XLOOKUP or a numeric key.** Excel 2016 compatibility (Global Constraints, inherited from the spec) rules out XLOOKUP. A single concatenated text key (`"J1|3"`) rather than two separate `MATCH`es ANDed together is what keeps each formula to one `MATCH` per side - `INDEX`/`MATCH` has no built-in multi-criteria form without either a helper column (which 3a already provides, as the join key) or an array formula (which raises its own Excel-version and CSE-entry complications this plan avoids entirely by using the join key instead).

**Why blank, not an error, for an unwired pin.** Task 3's `j1.Cells(3, 12).Value == ""` (pin 2, never referenced by any chart row) is the outer `IFERROR`'s fallback resolving with no inner match either - traced directly in `LookupFormula`/`WireToFormula`'s nested-`IFERROR` structure, not a special case bolted on. This matches the spec's validation philosophy even though Check Drawing itself is Phase 4: an unwired pin is informational, never an error, and here it does not even surface as one to look at.

**Type consistency.** `WriteLiveFormulas`'s row computation (`CONN_TABLE_FIRST_ROW + (i - LBound(vPins, 1))`) is identical to `WriteTableSkeleton`'s (3b) - both iterate the same `vPins` array in the same order, so a pin table's static columns (Pin, Label) and its formula columns (Wire To through Length) are guaranteed to land on the same physical row for the same pin, never off by one relative to each other.

**No placeholders.** Every formula string is written out in full in both the builder functions and the tests asserting their exact text - no "build the appropriate formula" left unspecified. Task 3's Step 2 placeholder line is explicitly called out as not surviving into the committed test, with the real assertion given in Step 3 immediately after - this is a deliberate two-step derivation (fixture needs distinct values before the assertion can be written correctly), not an unresolved gap.
