import pytest
import pywintypes

from tests.conftest import run


def test_success_builds_a_three_element_result(wb):
    assert run(wb, "modContract.Success", "SAVED", "DTM-04P") == (True, "SAVED", "DTM-04P")


def test_failure_builds_a_three_element_result(wb):
    assert run(wb, "modContract.Failure", "ID_COLLISION", "DTM-04P") == (False, "ID_COLLISION", "DTM-04P")


def test_a_none_kind_outcome_carries_an_empty_payload(wb):
    assert run(wb, "modContract.Success", "OK") == (True, "OK", None)


def test_accessors_read_the_slots_by_name(wb):
    result = run(wb, "modContract.Success", "PLACED", 3)
    assert run(wb, "modContract.Ok", result) is True
    assert run(wb, "modContract.Outcome", result) == "PLACED"
    assert run(wb, "modContract.Payload", result) == 3


def test_an_unknown_outcome_code_raises(wb):
    with pytest.raises(pywintypes.com_error):
        run(wb, "modContract.Success", "NOT_A_REAL_CODE", "x")


def test_a_payload_of_the_wrong_kind_raises(wb):
    # SAVED declares STRING; 42 arrives as a numeric Variant.
    with pytest.raises(pywintypes.com_error):
        run(wb, "modContract.Success", "SAVED", 42)


def test_every_declared_code_has_a_payload_kind(wb):
    for code in run(wb, "modContract.OutcomeCodes"):
        assert run(wb, "modContract.PayloadKind", code) != ""


def test_table_row_count_handles_empty_and_populated(wb):
    assert run(wb, "modContract.TableRowCount", None) == 0
    assert run(wb, "modContract.TableRowCount", ((1, 2), (3, 4))) == 2
