from tests.conftest import run


def test_artifact_is_produced(artifact):
    assert artifact.exists()
    assert artifact.suffix == ".xlsm"


def test_vba_module_is_present_and_callable(wb):
    assert run(wb, "modUtil.BuildStamp") == "0.1.0"


def test_join_key_normalizes_case_and_whitespace(wb):
    assert run(wb, "modUtil.JoinKey", " j1 ", 3) == "J1|3"


XL_CONTINUOUS = 1
XL_EDGE_LEFT = 7


def test_harness_title_block_labels_are_bordered_and_columns_fit_them(wb):
    ws = wb.Worksheets("Harness")

    assert ws.Range("A2").Value == "Harness Name"
    assert ws.Range("A4").Value == "Description"
    assert ws.Range("G4").Value == "Length Units"

    borders = ws.Range("B2").MergeArea.Borders
    assert borders(XL_EDGE_LEFT).LineStyle == XL_CONTINUOUS

    assert ws.Columns(1).ColumnWidth >= 14  # "Harness Name" / "Description" fit
    assert ws.Columns(7).ColumnWidth >= 13  # "Length Units" fits
