import pytest

XL_VALIDATE_LIST = 3
XL_VALIDATE_DECIMAL = 2

# (column index, validation type, Formula1)
EXPECTED = [
    (1, XL_VALIDATE_LIST, "=ListRefDes"),
    (3, XL_VALIDATE_LIST, "=ListTermination"),
    (5, XL_VALIDATE_LIST, "=ListColor"),
    (6, XL_VALIDATE_LIST, "=ListAWG"),
    (7, XL_VALIDATE_DECIMAL, "0"),
    (8, XL_VALIDATE_LIST, "=ListTermination"),
    (9, XL_VALIDATE_LIST, "=ListRefDes"),
]


@pytest.mark.parametrize("column,vtype,formula", EXPECTED)
def test_first_data_row_validation(wb, column, vtype, formula):
    cell = wb.Worksheets("Harness").Cells(7, column)
    assert cell.Validation.Type == vtype
    assert cell.Validation.Formula1 == formula


@pytest.mark.parametrize("column,vtype,formula", EXPECTED)
def test_last_data_row_validation(wb, column, vtype, formula):
    cell = wb.Worksheets("Harness").Cells(1006, column)
    assert cell.Validation.Type == vtype
    assert cell.Validation.Formula1 == formula


@pytest.mark.parametrize("column", [2, 10])
def test_pin_columns_start_without_validation(wb, column):
    cell = wb.Worksheets("Harness").Cells(7, column)
    with pytest.raises(Exception):
        _ = cell.Validation.Type


@pytest.mark.parametrize("column", [4, 11])
def test_free_text_columns_have_no_validation(wb, column):
    cell = wb.Worksheets("Harness").Cells(7, column)
    with pytest.raises(Exception):
        _ = cell.Validation.Type
