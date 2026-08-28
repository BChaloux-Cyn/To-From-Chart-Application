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
| 14 | Place another pin after that deletion | New pin takes the next unused number, not the deleted one | |
| 15 | Snap Label with a pin selected | Badge jumps back onto its pin | |
| 16 | Clear Pins | All badges and list rows disappear | |
| 17 | Save with all fields and a photo | Editor closes; connector appears in Manage Library | |
| 18 | Save a second connector using an existing Part Number | "Part Number already exists in the library (...)" | |
| 19 | **Save without ever loading a photo** | "Could not save ... Load a photo before saving." (was silent before) | |
| 20 | Manage Library > Edit an existing connector | Fields populate; photo preview appears; pins show as badges | |
| 21 | Manage Library > Delete | Confirmation, then the row disappears and the list refreshes | |
| 22 | Manage Library > Export | Save dialog, then "Exported <ID>." | |
| 23 | Manage Library > Import that file into another library | "Import complete. 1 connector(s) imported." | |
| 24 | Import a file whose photo cannot be extracted | Replacement-photo prompt appears for that connector only | |
| 25 | Rename a ref des on Connectors to an unused value | Rename sticks; chart references follow | |
| 26 | Rename a ref des onto an existing one | Cell reverts to the previous value | |
| 27 | Paste a large block over the chart | Pin dropdowns rebuild; pasted values are not cleared | |
| 28 | Change the units field in the title block | Length units update on the chart | |
