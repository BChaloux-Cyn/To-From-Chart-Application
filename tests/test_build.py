from tests.conftest import run


def test_artifact_is_produced(artifact):
    assert artifact.exists()
    assert artifact.suffix == ".xlsm"


def test_vba_module_is_present_and_callable(wb):
    assert run(wb, "modUtil.BuildStamp") == "0.1.0"


def test_join_key_normalizes_case_and_whitespace(wb):
    assert run(wb, "modUtil.JoinKey", " j1 ", 3) == "J1|3"
