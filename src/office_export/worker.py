from __future__ import annotations

import ctypes
import json
import os
import sys
import threading
from ctypes import wintypes
from pathlib import Path
from typing import Any

from office_export.office_adapters import (
    AdapterFailure,
    export_excel_charts,
    export_excel_pdf,
    export_powerpoint_images,
    export_powerpoint_pdf,
    export_word_pdf,
    inspect_excel,
    inspect_powerpoint,
    inspect_word,
)

APPLICATIONS = {
    "word": {"progid": "Word.Application", "process": "WINWORD.EXE", "display": "Microsoft Word"},
    "excel": {"progid": "Excel.Application", "process": "EXCEL.EXE", "display": "Microsoft Excel"},
    "powerpoint": {
        "progid": "PowerPoint.Application",
        "process": "POWERPNT.EXE",
        "display": "Microsoft PowerPoint",
    },
}


def main() -> int:
    try:
        request = json.load(sys.stdin)
        if not isinstance(request, dict):
            raise AdapterFailure("worker_invalid_request", "Worker request must be a JSON object.", category="usage")
        result = execute_request(request)
        sys.stdout.write(json.dumps({"ok": True, **result}, ensure_ascii=False))
        return 0
    except AdapterFailure as exc:
        payload: dict[str, Any] = {
            "ok": False,
            "error": {"code": exc.code, "message": exc.message, "category": exc.category},
        }
        if exc.details:
            payload["error"]["details"] = exc.details
        sys.stdout.write(json.dumps(payload, ensure_ascii=False))
        return 1
    except Exception as exc:
        failure = _unexpected_failure(exc)
        sys.stdout.write(json.dumps({"ok": False, "error": failure}, ensure_ascii=False))
        return 1


def execute_request(
    request: dict[str, Any],
    *,
    create_application: Any | None = None,
    co_initialize: Any | None = None,
    co_uninitialize: Any | None = None,
    process_snapshot: Any | None = None,
    application_pid: Any | None = None,
    emit_event: Any | None = None,
) -> dict[str, Any]:
    application_name = request.get("application")
    if application_name not in APPLICATIONS:
        raise AdapterFailure("worker_invalid_application", "Worker application must be word, excel, or powerpoint.")
    action = request.get("action")
    if action not in {"capabilities", "inspect", "export_pdf", "export_images", "export_charts"}:
        raise AdapterFailure("worker_invalid_action", "Worker action is not supported.", category="usage")

    using_real_com = create_application is None
    if create_application is None or co_initialize is None or co_uninitialize is None:
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise AdapterFailure(
                "office_automation_unavailable",
                "Windows Office automation support is unavailable. Reinstall office-export on Windows.",
                category="capability",
            ) from exc
        create_application = create_application or (lambda progid: win32com.client.DispatchEx(progid))
        co_initialize = co_initialize or pythoncom.CoInitialize
        co_uninitialize = co_uninitialize or pythoncom.CoUninitialize
    process_snapshot = process_snapshot or _snapshot_processes
    application_pid = application_pid or _application_pid
    emit_event = emit_event or _emit_event

    details = APPLICATIONS[application_name]
    application = None
    document = None
    owned = False
    pid: int | None = None
    cleanup_warnings: list[dict[str, str]] = []
    co_initialize()
    try:
        before = process_snapshot(details["process"])
        try:
            application = create_application(details["progid"])
        except Exception as exc:
            if _is_class_not_registered(exc):
                raise AdapterFailure(
                    f"{application_name}_not_installed",
                    f"{details['display']} is not installed or its COM registration is unavailable.",
                    category="capability",
                ) from exc
            raise AdapterFailure(
                f"{application_name}_start_failed",
                f"Could not start {details['display']} ({type(exc).__name__}).",
                category="capability",
                details=_com_details(exc),
            ) from exc
        pid = application_pid(application)
        after = process_snapshot(details["process"])
        created_processes = sorted(after - before)
        if pid is None and len(created_processes) == 1:
            pid = created_processes[0]
        owned = pid is not None and pid in created_processes
        emit_event({"event": "office_process", "application": application_name, "pid": pid, "owned": owned})
        _configure_application(application_name, application)
        emit_event({"event": "worker_stage", "stage": "application_configured"})
        application_info = _application_info(application_name, application, pid, owned)
        if action == "capabilities":
            response = {"application": application_info, "capabilities": _capabilities(application_name)}
            emit_event({"event": "operation_result", "response": {"ok": True, **response}})
            return response

        source = _source_path(request)
        options = request.get("options") if isinstance(request.get("options"), dict) else {}
        try:
            document = _open_document(application_name, application, source, options)
        except Exception as exc:
            raise _open_failure(application_name, source, exc) from exc
        emit_event({"event": "worker_stage", "stage": "document_opened"})

        if action == "inspect":
            if application_name == "word":
                payload = inspect_word(document)
            elif application_name == "excel":
                payload = inspect_excel(document, application)
            else:
                payload = inspect_powerpoint(document)
        elif action == "export_pdf":
            output = _output_file(request)
            if application_name == "word":
                payload = export_word_pdf(document, output, options)
            elif application_name == "excel":
                if options.get("refresh_data"):
                    document.RefreshAll()
                    try:
                        application.CalculateUntilAsyncQueriesDone()
                    except Exception:
                        pass
                payload = export_excel_pdf(document, application, output, options)
            else:
                payload = export_powerpoint_pdf(document, output, options)
        elif action == "export_images" and application_name == "powerpoint":
            payload = export_powerpoint_images(document, _output_directory(request), options)
        elif action == "export_charts" and application_name == "excel":
            payload = export_excel_charts(document, _output_directory(request), options)
        else:
            raise AdapterFailure(
                "unsupported_office_operation",
                f"{details['display']} does not support worker action '{action}'.",
                category="usage",
            )
        emit_event({"event": "worker_stage", "stage": "operation_complete"})
        response = {"application": application_info, "result": payload, "warnings": cleanup_warnings}
        emit_event({"event": "operation_result", "response": {"ok": True, **response}})
        return response
    finally:
        cleanup_watchdog = None
        if using_real_com:
            cleanup_watchdog = threading.Timer(5.0, lambda: os._exit(90))
            cleanup_watchdog.daemon = True
            cleanup_watchdog.start()
        if document is not None:
            emit_event({"event": "worker_stage", "stage": "document_close_started"})
            try:
                _close_document(application_name, document)
            except Exception as exc:
                cleanup_warnings.append(
                    {
                        "code": "document_close_failed",
                        "message": f"Office document cleanup failed ({type(exc).__name__}: {exc}).",
                    }
                )
            emit_event({"event": "worker_stage", "stage": "document_close_finished"})
        document = None
        if application is not None and owned:
            emit_event({"event": "worker_stage", "stage": "application_quit_started"})
            try:
                application.Quit()
            except Exception as exc:
                cleanup_warnings.append(
                    {
                        "code": "application_quit_failed",
                        "message": f"Office application cleanup failed ({type(exc).__name__}: {exc}).",
                    }
                )
            emit_event({"event": "worker_stage", "stage": "application_quit_finished"})
        application = None
        emit_event({"event": "worker_stage", "stage": "com_uninitialize_started"})
        co_uninitialize()
        emit_event({"event": "worker_stage", "stage": "com_uninitialize_finished"})
        if using_real_com and owned and pid is not None and not _wait_for_process_exit(pid, 1500):
            emit_event({"event": "worker_stage", "stage": "owned_process_termination_started"})
            if _terminate_owned_process(pid):
                cleanup_warnings.append(
                    {
                        "code": "owned_office_process_forced",
                        "message": (
                            "The isolated Office process remained after graceful cleanup and was stopped. "
                            "No preexisting Office process was affected."
                        ),
                    }
                )
            emit_event({"event": "worker_stage", "stage": "owned_process_termination_finished"})
        if cleanup_watchdog is not None:
            cleanup_watchdog.cancel()


def _configure_application(application_name: str, application: Any) -> None:
    _set_if_possible(application, "AutomationSecurity", 3)
    if application_name == "word":
        _set_if_possible(application, "Visible", False)
        _set_if_possible(application, "DisplayAlerts", 0)
        _set_if_possible(application.Options, "SaveNormalPrompt", False)
        _set_if_possible(application.Options, "ConfirmConversions", False)
        _set_if_possible(application.Options, "UpdateLinksAtOpen", False)
    elif application_name == "excel":
        _set_if_possible(application, "Visible", False)
        _set_if_possible(application, "DisplayAlerts", False)
        _set_if_possible(application, "AskToUpdateLinks", False)
        _set_if_possible(application, "EnableEvents", False)
        _set_if_possible(application, "ScreenUpdating", False)
    else:
        _set_if_possible(application, "DisplayAlerts", 1)


def _open_document(application_name: str, application: Any, source: Path, options: dict[str, Any]) -> Any:
    if application_name == "word":
        return application.Documents.Open(
            FileName=str(source),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Revert=False,
            PasswordDocument="",
            WritePasswordDocument="",
            NoEncodingDialog=True,
        )
    if application_name == "excel":
        return application.Workbooks.Open(
            Filename=str(source),
            UpdateLinks=3 if options.get("update_links") else 0,
            ReadOnly=True,
            Password="",
            WriteResPassword="",
            IgnoreReadOnlyRecommended=True,
            Notify=False,
            AddToMru=False,
            CorruptLoad=0,
        )
    return application.Presentations.Open(FileName=str(source), ReadOnly=True, Untitled=False, WithWindow=False)


def _close_document(application_name: str, document: Any) -> None:
    if application_name == "excel":
        document.Close(SaveChanges=False)
    elif application_name == "word":
        document.Close(SaveChanges=0)
    else:
        document.Close()


def _application_info(application_name: str, application: Any, pid: int | None, owned: bool) -> dict[str, Any]:
    return {
        "id": application_name,
        "name": APPLICATIONS[application_name]["display"],
        "version": _string_or_none(lambda: application.Version),
        "build": _string_or_none(lambda: application.Build),
        "process_id": pid,
        "process_owned": owned,
        "active_printer": _active_printer(application_name, application),
    }


def _active_printer(application_name: str, application: Any) -> str | None:
    if application_name not in {"word", "excel"}:
        return None
    return _string_or_none(lambda: application.ActivePrinter)


def _capabilities(application_name: str) -> dict[str, Any]:
    common = {"inspect": True, "pdf": True, "macros_disabled": True, "read_only": True}
    if application_name == "powerpoint":
        common["native_images"] = ["png", "jpeg"]
        common["output_types"] = [
            "slides",
            "notes",
            "outline",
            "handout1",
            "handout2",
            "handout3",
            "handout4",
            "handout6",
            "handout9",
        ]
    if application_name == "excel":
        common["native_chart_images"] = ["png"]
    return common


def _source_path(request: dict[str, Any]) -> Path:
    source = Path(str(request.get("source", "")))
    if not source.is_file():
        raise AdapterFailure("source_not_found", f"Source file does not exist: {source}", category="usage")
    return source.resolve()


def _output_file(request: dict[str, Any]) -> Path:
    output = Path(str(request.get("output_file", "")))
    if not output.parent.is_dir():
        raise AdapterFailure("worker_output_directory_missing", "Worker output directory does not exist.")
    return output.resolve()


def _output_directory(request: dict[str, Any]) -> Path:
    output = Path(str(request.get("output_dir", "")))
    if not output.is_dir():
        raise AdapterFailure("worker_output_directory_missing", "Worker output directory does not exist.")
    return output.resolve()


def _open_failure(application_name: str, source: Path, exc: Exception) -> AdapterFailure:
    text = _exception_text(exc).lower()
    if any(token in text for token in ("password", "protected", "permission", "rights management", "sensitivity")):
        return AdapterFailure(
            "source_protected",
            f"{APPLICATIONS[application_name]['display']} could not open the protected source: {source}",
            category="security",
            details=_com_details(exc),
        )
    return AdapterFailure(
        f"{application_name}_open_failed",
        f"{APPLICATIONS[application_name]['display']} could not open '{source}' ({type(exc).__name__}).",
        details=_com_details(exc),
    )


def _unexpected_failure(exc: Exception) -> dict[str, Any]:
    if isinstance(exc, AdapterFailure):
        payload: dict[str, Any] = {"code": exc.code, "message": exc.message, "category": exc.category}
        if exc.details:
            payload["details"] = exc.details
        return payload
    text = _exception_text(exc).lower()
    category = "security" if any(token in text for token in ("password", "protected", "permission")) else "conversion"
    code = "source_protected" if category == "security" else "office_com_error"
    return {
        "code": code,
        "message": f"Microsoft Office automation failed ({type(exc).__name__}).",
        "category": category,
        "details": _com_details(exc),
    }


def _com_details(exc: Exception) -> dict[str, Any]:
    details: dict[str, Any] = {"exception_type": type(exc).__name__}
    hresult = getattr(exc, "hresult", None)
    if isinstance(hresult, int):
        details["hresult"] = hresult
        details["hresult_hex"] = f"0x{hresult & 0xFFFFFFFF:08X}"
    excepinfo = getattr(exc, "excepinfo", None)
    if excepinfo and len(excepinfo) > 2 and excepinfo[2]:
        details["com_description"] = str(excepinfo[2])[:500]
    return details


def _exception_text(exc: Exception) -> str:
    values = [str(exc)]
    values.extend(str(value) for value in getattr(exc, "args", ()))
    excepinfo = getattr(exc, "excepinfo", None)
    if excepinfo:
        values.extend(str(value) for value in excepinfo if value)
    return " ".join(values)


def _is_class_not_registered(exc: Exception) -> bool:
    text = _exception_text(exc).lower()
    return "class not registered" in text or "-2147221164" in text or "0x80040154" in text


def _set_if_possible(target: Any, name: str, value: Any) -> None:
    try:
        setattr(target, name, value)
    except Exception:
        pass


def _string_or_none(getter: Any) -> str | None:
    try:
        value = getter()
    except Exception:
        return None
    text = str(value).strip() if value is not None else ""
    return text or None


def _application_pid(application: Any) -> int | None:
    hwnd = None
    for name in ("Hwnd", "HWND"):
        try:
            hwnd = int(getattr(application, name))
            break
        except Exception:
            continue
    if not hwnd:
        return None
    pid = wintypes.DWORD()
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    user32.GetWindowThreadProcessId(wintypes.HWND(hwnd), ctypes.byref(pid))
    return int(pid.value) or None


def _wait_for_process_exit(pid: int, timeout_ms: int) -> bool:
    if sys.platform != "win32" or pid <= 0:
        return True
    synchronize = 0x00100000
    wait_object_0 = 0x00000000
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
    kernel32.WaitForSingleObject.restype = wintypes.DWORD
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL
    handle = kernel32.OpenProcess(synchronize, False, pid)
    if not handle:
        return True
    try:
        return int(kernel32.WaitForSingleObject(handle, timeout_ms)) == wait_object_0
    finally:
        kernel32.CloseHandle(handle)


def _terminate_owned_process(pid: int) -> bool:
    if sys.platform != "win32" or pid <= 0:
        return False
    process_terminate = 0x0001
    synchronize = 0x00100000
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
    kernel32.TerminateProcess.restype = wintypes.BOOL
    kernel32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
    kernel32.WaitForSingleObject.restype = wintypes.DWORD
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL
    handle = kernel32.OpenProcess(process_terminate | synchronize, False, pid)
    if not handle:
        return False
    try:
        if not kernel32.TerminateProcess(handle, 1):
            return False
        kernel32.WaitForSingleObject(handle, 2000)
        return True
    finally:
        kernel32.CloseHandle(handle)


def _emit_event(payload: dict[str, Any]) -> None:
    sys.stderr.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stderr.flush()


class _ProcessEntry32W(ctypes.Structure):
    _fields_ = [
        ("dwSize", wintypes.DWORD),
        ("cntUsage", wintypes.DWORD),
        ("th32ProcessID", wintypes.DWORD),
        ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
        ("th32ModuleID", wintypes.DWORD),
        ("cntThreads", wintypes.DWORD),
        ("th32ParentProcessID", wintypes.DWORD),
        ("pcPriClassBase", ctypes.c_long),
        ("dwFlags", wintypes.DWORD),
        ("szExeFile", wintypes.WCHAR * 260),
    ]


def _snapshot_processes(image_name: str) -> set[int]:
    if sys.platform != "win32":
        return set()
    invalid_handle = ctypes.c_void_p(-1).value
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateToolhelp32Snapshot.argtypes = [wintypes.DWORD, wintypes.DWORD]
    kernel32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    kernel32.Process32FirstW.argtypes = [wintypes.HANDLE, ctypes.POINTER(_ProcessEntry32W)]
    kernel32.Process32FirstW.restype = wintypes.BOOL
    kernel32.Process32NextW.argtypes = [wintypes.HANDLE, ctypes.POINTER(_ProcessEntry32W)]
    kernel32.Process32NextW.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL
    snapshot = kernel32.CreateToolhelp32Snapshot(0x00000002, 0)
    if snapshot == invalid_handle:
        return set()
    entry = _ProcessEntry32W()
    entry.dwSize = ctypes.sizeof(_ProcessEntry32W)
    found: set[int] = set()
    try:
        success = kernel32.Process32FirstW(snapshot, ctypes.byref(entry))
        while success:
            if entry.szExeFile.casefold() == image_name.casefold():
                found.add(int(entry.th32ProcessID))
            success = kernel32.Process32NextW(snapshot, ctypes.byref(entry))
    finally:
        kernel32.CloseHandle(snapshot)
    return found


if __name__ == "__main__":
    exit_code = main()
    sys.stdout.flush()
    sys.stderr.flush()
    os._exit(exit_code)
