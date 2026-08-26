import pytest

from tests.conftest import run


@pytest.mark.parametrize(
    "part_number,name,expected",
    [
        ("DTM06-4S", "Deutsch DTM 4-way", "DTM06-4S"),
        ("dtm06-4s", "Deutsch DTM 4-way", "DTM06-4S"),
        # Spec's global constraint: "non-alphanumerics replaced with -" -
        # spaces become hyphens, not stripped, matching the DTM case below.
        ("", "Chassis Ground Stud", "CHASSIS-GROUND-STUD"),
        ("DTM 06/4-S!", "", "DTM-06-4-S-"),
        ("  ", "  ", ""),
    ],
)
def test_slugify(wb, part_number, name, expected):
    assert run(wb, "modLibrary.SlugifyConnectorID", part_number, name) == expected


def test_unique_id_passes_through_when_no_collision(wb, library_wb):
    ws = library_wb.Worksheets("Connectors")
    result = run(wb, "modLibrary.UniqueConnectorID", ws, 2, 100000, "DTM-04P")
    assert result == "DTM-04P"


def test_unique_id_appends_a_numeric_suffix_on_collision(wb, library_wb):
    ws = library_wb.Worksheets("Connectors")
    fields = ("DTM-04P", "A", "", "", "Connector", 4, "", "", "", "", "Local")
    run(wb, "modLibrary.WriteConnector", ws, 2, 100000, fields)

    result = run(wb, "modLibrary.UniqueConnectorID", ws, 2, 100000, "DTM-04P")
    assert result == "DTM-04P-2"


def test_unique_id_skips_past_multiple_collisions(wb, library_wb):
    ws = library_wb.Worksheets("Connectors")
    for connector_id in ("DTM-04P", "DTM-04P-2", "DTM-04P-3"):
        fields = (connector_id, "A", "", "", "Connector", 4, "", "", "", "", "Local")
        run(wb, "modLibrary.WriteConnector", ws, 2, 100000, fields)

    result = run(wb, "modLibrary.UniqueConnectorID", ws, 2, 100000, "DTM-04P")
    assert result == "DTM-04P-4"
