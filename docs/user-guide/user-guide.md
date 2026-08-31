# Wire Harness Creator - Student User Guide

This guide covers what the Creator can do today. It does not cover saving,
loading, printing, or exporting a harness file - that functionality is not
built yet. For now, everything you do lives inside `HarnessCreator.xlsm`
for the current session.

## What this tool is

The Wire Harness Creator gives you a guided way to document a wire harness:
you fill in a to-from chart describing every wire, and attach a photograph
of each connector you're actually holding, with each pin numbered on the
photo. `HarnessCreator.xlsm` is the editor - it is not itself a drawing, and
you never turn it in as your work. You use it to define connectors and fill
in the chart; a later phase of this tool will turn that into a print-ready
harness file.

## Adding a connector to your harness

From the `Home` sheet, click **Add Connector**. A picker opens listing
every connector already defined in your connector library, shown as
`ConnectorID - Name`.

- Select one and click **Add** to place an instance of it on your harness.
  The Creator assigns the next free reference designator for that
  connector's type (`J1`, `J2`, ... for a `Connector`; `ST1` for a `Stud`;
  `SP1` for a `Splice`; `TL1` for a `Tail`) and adds a row for it on the
  `Connectors` sheet.
- If the part you need isn't in the library yet, click **New...** instead.
  This opens the connector editor (below) so you can define it. Once you
  save, the Creator automatically adds it to your harness the same way Add
  would have - you don't need to reopen the picker and add it a second
  time.

Adding the same library part twice (two physical `DTM-04P` connectors in
one harness, say) is normal - each Add gives you a new reference designator
(`J1`, then `J2`) that both point back to the same connector definition.

## Defining a new connector

The connector editor opens from **Add Connector > New...**, or from
**Manage Library > Edit** on an existing part. It has these fields and
controls:

- **Name, Manufacturer, Part Number, Type, Pin Count, Notes** - Type is
  one of `Connector`, `Stud`, `Splice`, or `Tail`; this determines the
  reference-designator prefix new instances get. Name and Part Number must
  both be filled in before you can load a photo.
- **Load Photo** - opens a file picker restricted to JPG files. Once
  loaded, the photo is fit to the preview box without stretching or
  letterboxing.
- **Place Pins** (toggle) - while this is on, each click on the photo drops
  the next numbered pin marker at that spot. Placement stops once you've
  placed as many pins as Pin Count says the connector has.
- **Moving a pin you've already placed** works two different ways,
  depending on what you're trying to fix:
  - **Drag a marker's numbered badge** to pull it away from a crowded spot
    on the photo - this only moves the number, not the pin's actual
    location.
  - **Turn Place Pins off, select the pin in the list, then click its real
    position on the photo** - this corrects where the pin actually is. If
    you hadn't already dragged its badge away, the badge follows; if you
    had, the badge stays where you put it.
- **Delete Pin** - removes the pin selected in the list. Its number becomes
  available for the next pin you place.
- **Clear Pins** - removes every placed pin at once.
- **Snap Label to Pin** - moves the selected pin's badge back onto its
  actual position.
- **Save** - writes the connector, its pins, and its photo to your
  connector library. **Save requires a photo to be loaded** - saving
  without one fails with a message rather than silently doing nothing.
- **Cancel** - discards everything you did in this session; nothing is
  written to the library.

Fill in Name and Part Number, then load your photo, before you start
placing pins - editing Name or Part Number after loading a photo does not
update which connector your placed pins belong to.

## Filling in the to-from chart

The `Harness` sheet's chart is where you describe each wire, one row per
wire:

| Column | What it means |
|---|---|
| From Conn | Which connector this wire starts at - a dropdown of every reference designator on the `Connectors` sheet |
| From Pin | Which pin on that connector - rebuilt automatically once you pick From Conn |
| From Term | The termination type at that end (crimp pin, ring terminal, etc.) |
| Signal | Free text - what this wire carries |
| Color | Wire color |
| AWG | Wire gauge |
| Length (in) | Wire length in the units shown in the title block; must be a number greater than zero |
| To Term | Termination type at the far end |
| To Conn | Which connector this wire ends at |
| To Pin | Which pin on that connector - rebuilt automatically once you pick To Conn |
| Notes | Free text |

From Conn drives From Pin: until you pick a connector for a row, that
row's From Pin has no valid choices. Picking a connector rebuilds the From
Pin dropdown to exactly that connector's pin numbers. To Conn and To Pin
work the same way, independently.

The **Length Units** field in the title block (`in` or `mm`) changes the
Length column's header text to match. It does not convert numbers you've
already entered.

## Managing your connector library

**Home > Manage Library** opens a browser listing every connector in your
library.

- **Edit** reopens the connector editor pre-loaded with that connector's
  fields, photo, and pins, so you can correct or extend a definition you
  already made.
- **Delete** removes a connector from the library entirely, after you
  confirm. If you've already placed an instance of that connector on your
  current harness, deleting it also removes those instances from the
  `Connectors` sheet and clears the chart cells that referenced them - you
  won't be left with a wire pointing at a connector that no longer exists.
- **Import...** brings in every connector from another student's shared
  library file. If an incoming connector's Part Number already matches one
  in your library, you're asked whether to keep yours or overwrite it with
  the imported one - nothing is merged automatically. If overwriting
  changes that connector's pin count and you already have an instance of
  it on your chart, that instance is removed from the chart the same way
  Delete removes one. If a connector's photo can't be carried over
  automatically, you're prompted to choose a replacement image for it.
- **Export Connector** saves the currently selected connector - its
  record, pins, and photo - to a new `.xlsx` file another student can
  import.
- **Export Library** does the same for every connector in your library at
  once, into a single shared file.

## Renaming a connector

A connector instance's reference designator (the `J1`, `ST1`, etc. on the
`Connectors` sheet) can be renamed by editing it directly on that sheet.
Doing so updates every row in the chart that referenced the old name. Two
connector instances can't share a reference designator - if you rename one
to a value already in use, the cell reverts to what it was.

## Starting over

**Home > New Harness** clears the chart, the connectors list, and the
title block, and resets everything to a blank harness.
