from __future__ import annotations

from typing import Any

from office_export.worker import execute_request


class FakeOptions:
    SaveNormalPrompt = True
    ConfirmConversions = True
    UpdateLinksAtOpen = True


class FakeApplication:
    Version = "16.0"
    Build = "12345"
    ActivePrinter = "Test Printer"
    Options = FakeOptions()

    def __init__(self) -> None:
        self.quit_called = False

    def Quit(self) -> None:
        self.quit_called = True


def test_worker_quits_only_a_process_proven_to_be_new() -> None:
    application = FakeApplication()
    snapshots = iter([set(), {101}])
    initialized: list[bool] = []
    uninitialized: list[bool] = []
    events: list[dict[str, Any]] = []
    result = execute_request(
        {"action": "capabilities", "application": "word"},
        create_application=lambda progid: application,
        co_initialize=lambda: initialized.append(True),
        co_uninitialize=lambda: uninitialized.append(True),
        process_snapshot=lambda name: next(snapshots),
        application_pid=lambda app: None,
        emit_event=events.append,
    )
    assert result["application"]["process_id"] == 101
    assert result["application"]["process_owned"] is True
    assert application.quit_called is True
    assert initialized == [True]
    assert uninitialized == [True]
    assert events[0] == {"event": "office_process", "application": "word", "pid": 101, "owned": True}
    assert any(event["event"] == "operation_result" for event in events)
    assert events[-1] == {"event": "worker_stage", "stage": "com_uninitialize_finished"}


def test_worker_does_not_quit_preexisting_application() -> None:
    application = FakeApplication()
    snapshots = iter([{101}, {101}])
    result = execute_request(
        {"action": "capabilities", "application": "word"},
        create_application=lambda progid: application,
        co_initialize=lambda: None,
        co_uninitialize=lambda: None,
        process_snapshot=lambda name: next(snapshots),
        application_pid=lambda app: 101,
        emit_event=lambda event: None,
    )
    assert result["application"]["process_owned"] is False
    assert application.quit_called is False
