import pytest

EXPECTED = [
    ("Home", "shHome", -1),
    ("Harness", "shHarness", -1),
    ("Connectors", "shConnectors", -1),
    ("Check", "shCheck", -1),
    ("_Snapshot", "shSnapshot", 2),
    ("_Edit", "shEdit", 2),
    ("_Lists", "shLists", 2),
    ("_State", "shState", 2),
]


def test_sheet_count(wb):
    assert wb.Worksheets.Count == len(EXPECTED)


@pytest.mark.parametrize("tab,codename,visibility", EXPECTED)
def test_sheet_exists_with_codename_and_visibility(wb, tab, codename, visibility):
    sheet = wb.Worksheets(tab)
    assert sheet.CodeName == codename
    assert sheet.Visible == visibility


def test_sheet_order_matches_spec(wb):
    actual = [wb.Worksheets(i + 1).Name for i in range(wb.Worksheets.Count)]
    assert actual == [tab for tab, _, _ in EXPECTED]
