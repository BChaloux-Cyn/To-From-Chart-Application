from tests.conftest import run


def message_for(wb, ok, outcome, payload=None):
    builder = "modContract.Success" if ok else "modContract.Failure"
    result = run(wb, builder, outcome, payload) if payload is not None else run(wb, builder, outcome)
    return run(wb, "modMessages.MessageFor", result)


def test_id_collision_names_the_offending_id(wb):
    assert message_for(wb, False, "ID_COLLISION", "DTM-04P") == (
        "Part Number already exists in the library (DTM-04P). "
        "Choose a different Part Number."
    )


def test_save_failed_tells_the_student_what_to_do(wb):
    # The current build fails this save silently. The message is the fix.
    assert message_for(wb, False, "SAVE_FAILED", "DTM-04P") == (
        "Could not save DTM-04P. Load a photo before saving."
    )


def test_pin_limit_reached_names_the_cap(wb):
    assert message_for(wb, False, "PIN_LIMIT_REACHED", 4) == "All 4 pins have been placed."


def test_bad_pin_count_needs_no_payload(wb):
    assert message_for(wb, False, "BAD_PIN_COUNT") == (
        "Enter a valid Pin Count before placing pins."
    )


def test_missing_name_or_part(wb):
    assert message_for(wb, False, "MISSING_NAME_OR_PART") == (
        "Enter Name and Part Number before loading a photo."
    )


def test_export_success_and_failure(wb):
    assert message_for(wb, True, "EXPORTED", "DTM-04P") == "Exported DTM-04P."
    assert message_for(wb, False, "EXPORT_FAILED", "DTM-04P") == "Could not export DTM-04P."


def test_library_exported_names_the_count(wb):
    assert message_for(wb, True, "LIBRARY_EXPORTED", 3) == "Exported 3 connector(s)."


def test_connector_deleted_names_the_id(wb):
    assert message_for(wb, True, "CONNECTOR_DELETED", "DTM-04P") == "Deleted DTM-04P."


def test_connector_deleted_cascaded_lists_the_removed_instances(wb):
    assert message_for(wb, True, "CONNECTOR_DELETED_CASCADED", ("J1", "J2")) == (
        "Deleted from the library. "
        "Removed 2 connector instance(s) from the chart: J1, J2."
    )


def test_instance_not_found_names_the_ref_des(wb):
    assert message_for(wb, False, "INSTANCE_NOT_FOUND", "J99") == (
        "No connector instance 'J99' found."
    )


def test_silent_outcomes_produce_no_message(wb):
    for outcome in ("PLACED", "MOVED_ANCHOR", "NO_OP", "OK", "NO_RENAME"):
        payload = 1 if outcome in ("PLACED", "MOVED_ANCHOR") else None
        assert message_for(wb, True, outcome, payload) == ""


def test_style_is_information_on_success_and_exclamation_on_failure(wb):
    ok = run(wb, "modContract.Success", "EXPORTED", "DTM-04P")
    bad = run(wb, "modContract.Failure", "EXPORT_FAILED", "DTM-04P")
    assert run(wb, "modMessages.MessageStyleFor", ok) == 64      # vbInformation
    assert run(wb, "modMessages.MessageStyleFor", bad) == 48     # vbExclamation
