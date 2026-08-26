import pytest

CHART_HEADERS = [
    "From Conn", "From Pin", "From Term", "Signal", "Color", "AWG",
    "Length (in)", "To Term", "To Conn", "To Pin", "Notes",
]

TB_NAMES = [
    "TB_Name", "TB_Number", "TB_Rev", "TB_Student",
    "TB_Class", "TB_Date", "TB_Desc", "TB_Units",
]


@pytest.mark.parametrize("index,header", list(enumerate(CHART_HEADERS, start=1)))
def test_chart_header_text_and_order(wb, index, header):
    assert wb.Worksheets("Harness").Cells(6, index).Value == header


def test_chart_has_exactly_eleven_columns(wb):
    assert wb.Worksheets("Harness").Cells(6, 12).Value is None


@pytest.mark.parametrize("name", TB_NAMES)
def test_title_block_name_resolves_to_the_harness_sheet(wb, name):
    target = wb.Names(name).RefersToRange
    assert target.Worksheet.Name == "Harness"
    assert target.Cells.Count == 1


def test_units_default_to_inches(wb):
    assert wb.Names("TB_Units").RefersToRange.Value == "in"


def test_units_cell_offers_both_options(wb):
    assert wb.Names("TB_Units").RefersToRange.Validation.Formula1 == "in,mm"
