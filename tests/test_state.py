from tests.conftest import run


def test_seeded_keys_are_readable(wb):
    assert run(wb, "modState.GetState", "LengthUnits") == "in"
    assert run(wb, "modState.GetState", "HarnessPath") == ""


def test_unknown_key_returns_empty(wb):
    assert run(wb, "modState.GetState", "NoSuchKey") == ""


def test_set_state_round_trips(wb):
    run(wb, "modState.SetState", "HarnessPath", r"C:\temp\x.xlsx")
    assert run(wb, "modState.GetState", "HarnessPath") == r"C:\temp\x.xlsx"


def test_set_state_creates_a_missing_key(wb):
    run(wb, "modState.SetState", "BrandNewKey", "value")
    assert run(wb, "modState.GetState", "BrandNewKey") == "value"


def test_dirty_flag_starts_clear_and_toggles(wb):
    assert run(wb, "modState.IsDirty") is False
    run(wb, "modState.MarkDirty")
    assert run(wb, "modState.IsDirty") is True
    run(wb, "modState.ClearDirty")
    assert run(wb, "modState.IsDirty") is False
