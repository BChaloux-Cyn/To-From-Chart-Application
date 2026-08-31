## Purpose and Scope

This document describes the layering strategy used in the VBA source under
`src/vba/`: why it exists, what each layer may and may not do, and how it is
enforced.

It covers the connector library, editor, picker, and manage-library
workbook code (`src/vba/mod*Actions.bas`, `modContract`, `modMessages`, and
the layer 0 primitive modules), and the `.evt` adapters that wire them to
forms and worksheet events.

It does not cover the build system (`build/`) or the sheet layout constants
in `build/layout.py`; see the top-level `README.md` for those.

The full design rationale and the task-by-task extraction this layering was
built from live in `docs/superpowers/specs/2026-08-28-ui-logic-separation-design.md`
and `docs/superpowers/plans/2026-08-28-ui-logic-separation.md`. This document
is a shorter reference for day-to-day work; consult the spec for the "why"
behind an edge case not covered here.
