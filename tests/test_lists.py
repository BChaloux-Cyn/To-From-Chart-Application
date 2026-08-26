import pytest

COLORS = [
    "Black", "White", "Red", "Green", "Blue", "Yellow", "Orange",
    "Brown", "Violet", "Gray", "Pink", "Tan", "Light Blue",
    "Light Green", "Other",
]
AWGS = ["24", "22", "20", "18", "16", "14", "12", "10", "8"]
TERMINATIONS = [
    "Crimp Pin", "Crimp Socket", "Ring Terminal", "Spade Terminal",
    "Butt Splice", "Ferrule", "Solder Cup", "Quick Disconnect",
    "Bare Tinned", "None",
]


def column_values(wb, column: int, count: int):
    sheet = wb.Worksheets("_Lists")
    return [str(sheet.Cells(row + 2, column).Value) for row in range(count)]


def test_color_list_seeded(wb):
    assert column_values(wb, 1, len(COLORS)) == COLORS


def test_awg_list_seeded(wb):
    assert column_values(wb, 2, len(AWGS)) == AWGS


def test_termination_list_seeded(wb):
    assert column_values(wb, 3, len(TERMINATIONS)) == TERMINATIONS


@pytest.mark.parametrize(
    "name,expected_count",
    [
        ("ListColor", len(COLORS)),
        ("ListAWG", len(AWGS)),
        ("ListTermination", len(TERMINATIONS)),
    ],
)
def test_list_name_resolves_to_the_right_height(wb, name, expected_count):
    assert wb.Names(name).RefersToRange.Rows.Count == expected_count


def test_refdes_name_survives_an_empty_connector_sheet(wb):
    # No connectors defined yet; the name must still resolve rather than error.
    assert wb.Names("ListRefDes").RefersToRange.Rows.Count == 1
