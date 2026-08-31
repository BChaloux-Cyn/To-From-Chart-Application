# UI Separation - Manual Verification

Date run:
Build:

| # | Check | Expected | Outcome |
|---|---|---|---|
| 1 | Home > Add Connector | Picker opens, list shows "ID - Name" for every library connector | |
| 2 | Pick one, Add | Connector row appears on Connectors, _Snapshot gains the definition | |
| 3 | Add the same one again | Second ref des allocated (J1 then J2) | |
| 4 | Picker > New | Editor opens with the Type dropdown populated | |
| 5 | Editor: fill Name and Part Number, Load Photo | File picker offers JPG only; photo appears fitted, not stretched | |
| 6 | Load Photo again | Photo does not shrink further than the first load | |
| 7 | Editor: Load Photo with Name blank | "Enter Name and Part Number before loading a photo." | |
| 8 | Place Pins on, click the photo four times with Pin Count 4 | Four small numbered badges appear where clicked | |
| 9 | Click a fifth time | "All 4 pins have been placed." | |
| 10 | Clear Pin Count, click the photo | "Enter a valid Pin Count before placing pins." | |
| 11 | Drag a badge away from its pin | Badge moves and stays where dropped | |
| 12 | Select a pin, Place Pins off, click elsewhere | Anchor moves; badge follows only if it was still on the anchor | |
| 13 | Select the middle pin, Delete Pin | That badge and list row disappear; remaining numbers unchanged | |
| 14 | Place another pin after that deletion | New pin reuses the deleted number (the lowest unused number, never past the connector's pin count) | |
| 15 | Snap Label with a pin selected | Badge jumps back onto its pin | |
| 16 | Clear Pins | All badges and list rows disappear | |
| 17 | Save with all fields and a photo | Editor closes; connector appears in Manage Library | |
| 18 | Save a second connector using an existing Part Number | "Part Number already exists in the library (...)" | |
| 19 | **Save without ever loading a photo** | "Could not save ... Load a photo before saving." (was silent before) | |
| 20 | Manage Library > Edit an existing connector | Fields populate; photo preview appears; pins show as badges | |
| 21 | Manage Library > Delete (connector not placed on the chart) | Confirmation, then "Deleted <ID>." and the list refreshes | |
| 22 | Manage Library > Export Connector | Save dialog, then "Exported <ID>." | |
| 23 | Manage Library > Import that file into another library | "Import complete. Imported: 1. Kept: 0. Overwritten: 0." | |
| 24 | Import a file whose photo cannot be extracted | Replacement-photo prompt appears for that connector only | |
| 25 | Rename a ref des on Connectors to an unused value | Rename sticks; chart references follow | |
| 26 | Rename a ref des onto an existing one | Cell reverts to the previous value | |
| 27 | Paste a large block over the chart | Pin dropdowns rebuild; pasted values are not cleared | |
| 28 | Change the units field in the title block | Length units update on the chart | |
| 29 | Add an instance of a connector to the chart, then Manage Library > Delete that same connector | Confirmation, then "Deleted from the library. Removed 1 connector instance(s) from the chart: <RefDes>."; the chart row and its cell references are gone | |
| 30 | Home > Remove Connector | Picker opens listing every placed instance as "<RefDes> - <Name>" | |
| 31 | Pick one, Remove | That row disappears from Connectors and its chart cell references clear | |
| 32 | Manage Library > Export Library | Save dialog, then "Exported <N> connector(s)." | |
| 33 | Import a whole-library file where one connector's Part Number already exists locally | For each conflict: "A connector with Part Number '<ID>' already exists in the library. Overwrite it with the imported version?" | |
| 34 | Answer No (Keep) on that prompt | Local connector is untouched; final summary counts it under Kept | |
| 35 | Re-import and answer Yes (Overwrite), with the incoming connector's pin count different from the local one, while an instance of it is placed on the chart | Local connector's data is replaced; "Removed N connector instance(s) from the chart: ..." appears in the final summary and those chart rows are gone | |
