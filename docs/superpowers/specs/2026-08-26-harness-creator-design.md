# Wire Harness Documentation Creator - Design

Date: 2026-08-26
Status: Approved for planning

## Purpose

High school students struggle to produce electrical documentation. This tool
gives them a guided way to build a wire harness drawing by filling in a
to-from chart and attaching photographs of the connectors they are actually
holding, with each position numbered on the photo. The output is a
print-ready drawing they can turn in as a PDF.

## Architecture

Three artifacts, with a strict ownership boundary between them.

### 1. The Creator - `HarnessCreator.xlsm`

A single macro-enabled application workbook. It is an editor, not a drawing.
It is the only place a harness is authored and the only place a connector is
defined or modified. Students never fill it in and save it as their work;
they use it to produce harness files.

### 2. The Connector Library - `ConnectorLibrary.xlsx`

A macro-free workbook that lives beside the Creator and accumulates every
connector a student has defined. A connector is defined once and reused in
any number of harnesses. Libraries can be exported and imported so a
connector defined by one student can be merged into another's library.

### 3. Harness files - `<name>.xlsx`

Plain macro-free workbooks produced by the Creator. Each is a complete,
print-ready drawing that opens on any machine with no macro warning. It
carries a frozen snapshot of every connector it uses, so it never depends on
the library being present.

### Ownership rules

- Connector definitions are created and edited only in the Creator.
- Harness files are authored in the Creator, but remain directly editable
  and directly printable in Excel without it.
- A harness's connector snapshot is frozen at the moment the connector is
  added. Later library edits never propagate into an existing harness. To
  pick up a corrected connector, the student removes and re-adds it.

## Data flow

```
Connector Editor --writes--> ConnectorLibrary.xlsx
                                    |
                                    | student picks a connector
                                    v
                          Creator (in-memory harness)
                                    |
                       Save --------+-------- Load
                        |                       ^
                        v                       |
                  harness.xlsx  -----------------
                        |
                        +--> Ctrl+P (student, no macros)
                        +--> Export PDF button (Creator, auto-named)
```

The visible to-from chart inside a harness file is the canonical wire data.
The Creator reads it back verbatim on load, so a student's hand edits to the
chart survive a round trip. The hidden sheet in a harness file holds only the
frozen connector snapshot and file metadata, never a duplicate of the wire
data.

## Creator workbook structure

| Sheet | Visibility | Purpose |
|---|---|---|
| `Home` | visible | Instructions and command buttons; shows the loaded harness path and unsaved-changes state |
| `Harness` | visible | Title block and the to-from chart - the editing surface |
| `Connectors` | visible | Connector instances used by the loaded harness |
| `Check` | visible | Output of Check Drawing |
| `_Snapshot` | very hidden | Frozen connector definitions, pins, and photos for the loaded harness |
| `_Lists` | very hidden | Pick lists and their named ranges |
| `_State` | very hidden | Current file path, dirty flag, build version, test-mode flag |

### Home commands

New Harness, Open Harness, Save Harness, Save Harness As, Export PDF,
Save Archive Copy, Add Connector, Remove Connector, Manage Library,
Check Drawing.

There is no harness index. Open Harness is a file dialog.

### Title block fields

Harness Name, Harness Number, Revision, Student Name, Class or Project,
Date, Length Units (`in` or `mm`), Description.

Length Units is a workbook-level setting. Changing it rewrites the Length
column header to `Length (in)` or `Length (mm)`. It does not convert existing
values.

### Connector instances and reference designators

Two identifiers do different jobs. **ConnectorID** names a part type in the
library and is shared by every harness that uses that part. **Ref Des** names
one physical connector in one harness, and is what appears in the From Conn
and To Conn dropdowns, titles the connector page, and labels the photo on the
drawing.

Because the ref des is the instance key, a harness can use the same library
part more than once. Two `DTM-04P` bodies become `J1` and `J2`, each with its
own page, its own photo, and its own pin usage, both resolving to a single
library definition.

A harness may have any number of connectors, including one. Nothing in the
design assumes two endpoints of different kinds; a single connector with
flying leads is a valid harness whose far ends are `Tail`, `Stud`, or
`Splice` entries.

The `Connectors` sheet holds one row per instance: Ref Des, ConnectorID,
Name, Part Number, Type, Pin Count. Adding a connector auto-assigns the next
free ref des using a prefix chosen by the library entry's Type:

| Type | Prefix | Example |
|---|---|---|
| `Connector` | `J` | `J1`, `J2` |
| `Stud` | `ST` | `ST1` |
| `Splice` | `SP` | `SP1` |
| `Tail` | `TL` | `TL1` |

The assigned ref des is editable. Renaming one rewrites every reference to it
in the chart. Ref designators must be unique within a harness; a rename that
collides is rejected.

### To-from chart columns

Columns are ordered to read along the wire itself: the starting endpoint and
its terminal, then the wire's own properties, then the far terminal and the
ending endpoint. Notes is the only column that falls outside that sequence.

| Col | Header | Entry |
|---|---|---|
| A | From Conn | Dropdown of ref designators on the `Connectors` sheet |
| B | From Pin | Dependent dropdown, rebuilt by VBA when column A changes |
| C | From Term | Dropdown from `_Lists` |
| D | Signal | Free text |
| E | Color | Dropdown from `_Lists` |
| F | AWG | Dropdown from `_Lists` |
| G | Length (unit) | Number greater than zero |
| H | To Term | Dropdown from `_Lists` |
| I | To Conn | Dropdown of ref designators |
| J | To Pin | Dependent dropdown, rebuilt by VBA when column I changes |
| K | Notes | Free text |

There is no wire ID column. A wire is identified by its row position and by
its From/To endpoints.

Row count is open-ended. The chart grows as rows are used.

### Seeded pick lists

Editable on `_Lists`.

- Color: Black, White, Red, Green, Blue, Yellow, Orange, Brown, Violet,
  Gray, Pink, Tan, Light Blue, Light Green, Other
- AWG: 24, 22, 20, 18, 16, 14, 12, 10, 8
- Termination: Crimp Pin, Crimp Socket, Ring Terminal, Spade Terminal,
  Butt Splice, Ferrule, Solder Cup, Quick Disconnect, Bare Tinned, None

## Connector library schema

The same three-sheet schema is used by the library workbook, by an exported
library file, and by the `_Snapshot` sheet inside a harness. One reader and
one writer serve all three.

`Connectors`, header in row 1:

| Field | Notes |
|---|---|
| ConnectorID | Unique slug, uppercase, non-alphanumerics replaced with `-` |
| Name | Free text, required |
| Manufacturer | Free text |
| PartNumber | Free text |
| Type | `Connector`, `Stud`, `Splice`, or `Tail` |
| PinCount | Integer, 1 or greater |
| Notes | Free text |
| PhotoShapeName | Name of the picture on the `Photos` sheet |
| CreatedUtc | ISO 8601 |
| ModifiedUtc | ISO 8601 |
| Origin | `Local`, or the filename of the library it was imported from |

`Pins`:

| Field | Notes |
|---|---|
| ConnectorID | Foreign key |
| PinNumber | Integer, 1 to PinCount |
| PinLabel | Free text, may be blank |
| NormX | Anchor point, 0.0 to 1.0, fraction of photo width |
| NormY | Anchor point, 0.0 to 1.0, fraction of photo height |
| LabelX | Marker center, 0.0 to 1.0, fraction of photo width |
| LabelY | Marker center, 0.0 to 1.0, fraction of photo height |

Every pin carries two positions: the anchor point, which is the cavity on the
connector face, and the marker position, which is where the numbered circle
is drawn. When a pin is first placed the two are identical and the marker
sits directly on the point. When the marker is pulled away far enough that it
no longer covers its anchor, a leader line is drawn between them. Both pairs
are normalized, so the relationship survives any change of display or print
scale.

`Photos`: one picture per connector, laid out in a grid, each shape named
`PHOTO_<ConnectorID>`.

ConnectorID is derived from PartNumber when present, otherwise from Name. On
collision, a numeric suffix is appended.

### Non-connector endpoints

Because From and To are dropdown-only, endpoints that are not connectors are
modeled as single-position library entries with Type `Stud`, `Splice`, or
`Tail`. A ring lug on a battery stud is a one-pin `Stud` entry. This keeps
every endpoint inside the validated dropdown model.

### Photo cache

Photos are embedded in the library workbook so a library is a single portable
file. A working copy of each photo is also cached as
`Photos/<ConnectorID>.png` beside the library, and that cache is what the
editor loads. The cache is written when a connector is defined, and
regenerated by extracting the embedded picture when a connector arrives from
an imported library. If extraction fails, the editor prompts for the image
file rather than failing outright.

## Connector editor

Opened from Manage Library. A UserForm constructed in code at build time.

Fields: Name, Manufacturer, Part Number, Type, Pin Count, Notes.
Controls: Load Photo, an Image control, a pin list, Place Pins toggle,
Delete Pin, Clear Pins, Snap Label to Pin, Save, Cancel.

### Click-to-place

On photo load, the picture is loaded with `LoadPicture` and its aspect ratio
read from the returned `StdPicture` HIMETRIC dimensions. The Image control is
sized to that aspect ratio inside a fixed bounding box and set to stretch, so
the displayed image exactly fills the control and there is no letterboxing to
compensate for.

With Place Pins active, each `MouseUp` on the Image control drops the next
pin number at that point, recorded as `NormX = X / Control.Width` and
`NormY = Y / Control.Height`, with `LabelX` and `LabelY` set to match.
Markers are numbered Label controls added at runtime.

### Moving a marker versus moving a pin

The two gestures do different things, which is why both exist.

- **Dragging a marker** moves the label only. The anchor point stays on the
  cavity and a leader line appears once the marker clears it. This is how a
  student pulls numbers off a crowded connector face into the margin.
- **Selecting a pin in the list and clicking the image** moves the anchor
  point - the student is correcting where the pin actually is. If the marker
  was still sitting on its anchor it travels with it; if it had been pulled
  away, it stays put and the leader re-aims.
- **Snap Label to Pin** returns the selected pin's marker to its anchor,
  removing the leader.

Dragging is the primary gesture because it is what a student will try
unprompted. Runtime-created markers get their mouse events through a small
`WithEvents` wrapper class held in a collection, one instance per marker.
Drag is `MouseDown` to record the grab offset, `MouseMove` while the button
is held to reposition, `MouseUp` to commit.

Because positions are stored normalized, they survive any later change to
display size or print scale.

## Harness rendering

Saving a harness writes a new workbook containing:

- `Harness` - title block and the to-from chart, plus two hidden helper
  columns holding `FromConn|FromPin` and `ToConn|ToPin` join keys.
- One `CONN_<RefDes>` sheet per connector instance.
- `_Snapshot` - very hidden, the frozen connector schema sheets.

### Connector page layout

The photo is placed at a fixed anchor, scaled to a fixed maximum width with
height following its aspect ratio. For each pin, an oval shape carrying the
pin number is centered at
`photo.Left + LabelX * photo.Width`, `photo.Top + LabelY * photo.Height`,
with white fill and black border.

When the distance between the marker center and the anchor point exceeds the
marker's radius - that is, when the marker no longer covers its own anchor -
a thin leader line is drawn from the marker's edge to
`photo.Left + NormX * photo.Width`, `photo.Top + NormY * photo.Height`.
Markers left sitting on their anchors get no leader.

Beside the photo, a pin table with columns Pin, Label, Wire To, Signal,
Color, AWG, Termination, Length. Every cell in that table is an INDEX/MATCH
formula against the hidden join-key columns on the `Harness` sheet, matching
first as a From endpoint and falling back to a To endpoint. INDEX/MATCH is
used rather than XLOOKUP so the file works on Excel 2016 and later.

The table is written from this connector's point of view. Wire To is the
opposite endpoint, and Termination is the terminal on this connector's end -
From Term when the row matched as a From endpoint, To Term when it matched as
a To endpoint.

Because the pin tables are live formulas, a student who opens a saved harness
and corrects a length, color, gauge, termination, or signal sees the
connector pages update with no macros involved. Callout markers are static
shapes; a student who drags one by hand will have that change discarded the
next time the Creator re-renders the file. This is documented behavior, not
an error condition.

### Page setup

Baked into every sheet at save time: print area, page breaks, print titles
repeating the chart header row, fit-to-width scaling, landscape orientation
for the chart, and a footer carrying harness number, revision, and page
numbering. This is what makes Ctrl+P produce a correct PDF without macros.

## Save, load, and export

- **Save** builds the harness workbook from the Creator's current state and
  writes `.xlsx`. Save As prompts for a path.
- **Load** opens the chosen `.xlsx`, copies `_Snapshot` into the Creator,
  reads the title block and every used chart row, closes the file, then runs
  validation and reports findings on `Check`.
- **Export PDF** requires a saved harness; if there are unsaved changes it
  prompts to save first. It exports the saved file to
  `<HarnessNumber>_Rev<Revision>.pdf`, falling back to `<HarnessName>` when
  the number is blank.
- **Save Archive Copy** writes `<HarnessNumber>_Rev<Revision>_<yyyymmdd>.xlsx`
  to a chosen folder.

Unsaved-change tracking lives in `_State` and is set by any edit to the
`Harness` or `Connectors` sheets.

## Validation - Check Drawing

Reported as errors:

- Missing From Conn, From Pin, From Term, Color, AWG, Length, To Term,
  To Conn, or To Pin on a used row
- An endpoint referencing a connector not present in the snapshot
- A pin number outside 1 to PinCount
- From endpoint identical to To endpoint
- Length not numeric or not greater than zero

Reported as information, not errors:

- Pins defined on a connector but not wired

Output is a table on `Check`: row number, severity, message. The same
validation runs automatically after Load.

### Validation is advisory

Check Drawing never blocks anything. Saving a harness, exporting a PDF, and
writing an archive copy all succeed with errors outstanding. "Error" means
the tool believes something is wrong and is worth a second look, not that the
student is forbidden from producing their drawing.

This is deliberate. A student may be documenting a harness that is unusual,
half-finished, or correct in a way the tool does not model, and the tool
being wrong must never stop them from turning in work. Anything the tool
cannot be certain about is reported as information rather than as an error.

## Build system

The `.xlsm` is a build artifact, never hand-edited. All source is text.

```
build/          build.py and COM helpers; produces dist/
src/vba/        *.bas module source; UserForms constructed in code
tests/          pytest suites driving Excel COM
dist/           HarnessCreator.xlsm, ConnectorLibrary.xlsx
docs/           this spec, plus the student user guide
```

`build.py` drives Excel COM to create the workbooks, apply formatting, data
validation, named ranges, and command buttons, import the `.bas` modules, and
construct the UserForms, then saves as `.xlsm`.

**Prerequisite**: Excel's "Trust access to the VBA project object model" must
be enabled (File, Options, Trust Center, Trust Center Settings, Macro
Settings). `build.py --check` verifies this and prints instructions if it is
off. This is a per-user, reversible setting the developer enables; it is not
required on student machines.

Python dependencies: `pywin32`, `pytest`.

## Testing

pytest drives Excel COM against the built artifact.

- **Structural**: expected sheets exist with correct visibility; headers
  match; validation lists and named ranges resolve; every command button has
  a resolvable `OnAction`.
- **VBA units**: pure functions called directly through `Application.Run` -
  ID slugification, normalized-coordinate math, join-key construction, unit
  header text.
- **Ref designators**: adding connectors of each type assigns the expected
  prefix and next free number; adding the same library part twice yields two
  distinct instances resolving to one definition; renaming a ref des rewrites
  every chart reference; a colliding rename is rejected.
- **Library round trip**: write connectors and pins, save, reopen, assert
  equality; import a second library and assert merge and collision behavior.
- **Render**: build a harness with two connectors and several wires, save,
  reopen, assert the `CONN_` sheets exist, that oval count equals pin count,
  that oval positions match the stored marker coordinates within tolerance,
  and that pin-table formulas evaluate to the expected wire data. Assert a
  leader line is drawn for a pin whose marker was offset from its anchor, and
  that no leader exists for a pin whose marker sits on its anchor.
- **Validation**: seed each error class and assert it appears on `Check`.
  Assert that Save, Export PDF, and Save Archive Copy all still succeed while
  errors are outstanding.
- **Export**: export a PDF and assert the file exists and is non-trivial.

To keep logic testable, no logic module raises UI. `MsgBox` and dialogs are
confined to UI modules, and `_State` carries a test-mode flag that suppresses
them.

Every feature ships with tests. Every bug fix ships with a test that would
have caught it.

## Phasing

1. Build system, test harness, Creator shell, `Harness` sheet with the
   working to-from chart, pick lists, and dependent pin dropdowns.
2. Connector library file, library reader/writer, connector editor with
   click-to-place, connector picker, library import and export, snapshot
   embedding.
3. Harness save and load, generated connector pages with rendered callouts
   and live pin tables, page setup.
4. Check Drawing, PDF export, archive copy, student user guide.
5. Deferred, may never be built: an auto-drawn harness diagram showing
   connectors as boxes and wires as lines between pins.

## Decisions taken, with rationale

| Decision | Rationale |
|---|---|
| Harness files are macro-free `.xlsx` | Students should not have to click Enable Content on their own work, and old drawings should not carry old code. Baked-in page setup gives correct Ctrl+P PDF output without macros. |
| Connector snapshot frozen | A submitted drawing must never change after the fact. Simplest rule to explain and to implement. |
| Connectors editable only in the Creator | Prevents per-file drift of a shared definition. |
| Visible chart is canonical | Avoids a duplicate copy of wire data and makes hand edits safe on round trip. |
| No wire ID column | Chosen by the user. Rows are identified by position and endpoints. |
| Pin tables as live formulas | Lets a saved harness stay useful and self-correcting without macros. |
| Normalized pin coordinates | Callouts survive resizing and print scaling. |
| Separate anchor and marker positions per pin | Lets a student pull numbers off a crowded connector face without lying about where the pin is. Costs one extra coordinate pair and stays invisible until a marker is dragged. |
| Build from text source | An `.xlsm` is opaque to review and diffing; a build script keeps the real source readable and lets the work be tested automatically. |

## Out of scope

- Auto-drawn harness diagram (phase 5, deferred)
- Duplicate-endpoint checking. Two wires landing on the same connector pin is
  not reported. The rule is only correct for `Connector` endpoints - studs
  and splices are multi-drop by nature - and even for a connector a double
  crimp is sometimes the intended design. If this is revisited it ships as
  information, never as an error, because the tool cannot tell a mistake from
  a deliberate choice here.
- Wire bundle or branch modeling, and overall harness length calculations
- Bill of materials generation
- Any non-Windows or Excel-for-web support
- Multi-user or shared-write access to a single library file
