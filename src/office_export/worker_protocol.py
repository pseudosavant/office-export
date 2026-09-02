from __future__ import annotations

import ctypes
import json
import subprocess
import sys
from pathlib import Path
from typing import Any

from office_export.errors import (
    CapabilityError,
    ConversionError,
    OfficeExportError,
    SecurityError,
    UsageError,
    WorkerTimeoutError,
)


def run_worker(request: dict[str, Any], timeout: float) -> dict[str, Any]:
    command = [sys.executable, "-m", "office_export.worker"]
    creation_flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
    process = subprocess.Popen(
        command,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        creationflags=creation_flags,
    )
    serialized = json.dumps(request, ensure_ascii=False)
    try:
        stdout, stderr = process.communicate(serialized, timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        partial_stderr = _text(exc.stderr)
        process.kill()
        _, trailing_stderr = process.communicate()
        diagnostics = trailing_stderr or partial_stderr
        events = _events(diagnostics)
        owned_pids = _owned_process_ids(diagnostics)
        terminated: list[int] = []
        for pid in owned_pids:
            if _terminate_process(pid):
                terminated.append(pid)
        recovered = _recover_checkpoint(events, terminated, reason="timeout_during_cleanup")
        if recovered is not None:
            return recovered
        raise WorkerTimeoutError(
            f"Microsoft Office did not finish within {timeout:g} seconds.",
            details={
                "timeout_seconds": timeout,
                "terminated_owned_processes": terminated,
                "worker_events": events,
            },
        ) from exc

    if not stdout.strip():
        events = _events(stderr)
        terminated: list[int] = []
        if process.returncode == 90:
            for pid in _owned_process_ids(stderr):
                if _terminate_process(pid):
                    terminated.append(pid)
        recovered = _recover_checkpoint(events, terminated, reason="cleanup_watchdog")
        if recovered is not None:
            return recovered
        raise ConversionError(
            "worker_no_response",
            "The Office worker exited without a structured response.",
            details={"exit_code": process.returncode, "diagnostics": _diagnostic_text(stderr)},
        )
    try:
        response = json.loads(stdout)
    except json.JSONDecodeError as exc:
        raise ConversionError(
            "worker_invalid_response",
            "The Office worker returned an invalid response.",
            details={"exit_code": process.returncode, "diagnostics": _diagnostic_text(stderr)},
        ) from exc
    if not isinstance(response, dict):
        raise ConversionError("worker_invalid_response", "The Office worker response was not an object.")
    if not response.get("ok"):
        _raise_worker_error(response.get("error"))
    if stderr.strip():
        response["worker_events"] = _events(stderr)
    return response


def _raise_worker_error(raw: Any) -> None:
    error = raw if isinstance(raw, dict) else {}
    code = str(error.get("code") or "office_worker_failed")
    message = str(error.get("message") or "The Office worker failed.")
    details = error.get("details") if isinstance(error.get("details"), dict) else None
    category = error.get("category")
    exception: OfficeExportError
    if category == "usage":
        exception = UsageError(message, code=code, details=details)
    elif category == "capability":
        exception = CapabilityError(code, message, details=details)
    elif category == "security":
        exception = SecurityError(code, message, details=details)
    else:
        exception = ConversionError(code, message, details=details)
    raise exception


def _events(stderr: str) -> list[dict[str, Any]]:
    events: list[dict[str, Any]] = []
    for line in stderr.splitlines():
        try:
            value = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            events.append(value)
    return events


def _owned_process_ids(stderr: str) -> list[int]:
    return [
        int(event["pid"])
        for event in _events(stderr)
        if event.get("event") == "office_process" and event.get("owned") is True and isinstance(event.get("pid"), int)
    ]


def _recover_checkpoint(
    events: list[dict[str, Any]],
    terminated: list[int],
    *,
    reason: str,
) -> dict[str, Any] | None:
    checkpoints = [event.get("response") for event in events if event.get("event") == "operation_result"]
    if not checkpoints or not isinstance(checkpoints[-1], dict) or checkpoints[-1].get("ok") is not True:
        return None
    stages = {event.get("stage") for event in events if event.get("event") == "worker_stage"}
    if "operation_complete" not in stages and checkpoints[-1].get("result") is not None:
        return None
    response = dict(checkpoints[-1])
    warnings = list(response.get("warnings") or [])
    warnings.append(
        {
            "code": "worker_cleanup_forced",
            "message": "The Office operation completed, but COM cleanup exceeded its grace period. "
            "The isolated worker and any proven owned Office process were stopped.",
            "reason": reason,
            "terminated_owned_processes": terminated,
        }
    )
    response["warnings"] = warnings
    response["worker_events"] = events
    return response


def _terminate_process(pid: int) -> bool:
    if sys.platform != "win32" or pid <= 0:
        return False
    process_terminate = 0x0001
    synchronize = 0x00100000
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.argtypes = [ctypes.c_ulong, ctypes.c_int, ctypes.c_ulong]
    kernel32.OpenProcess.restype = ctypes.c_void_p
    kernel32.TerminateProcess.argtypes = [ctypes.c_void_p, ctypes.c_uint]
    kernel32.TerminateProcess.restype = ctypes.c_int
    kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
    kernel32.CloseHandle.restype = ctypes.c_int
    handle = kernel32.OpenProcess(process_terminate | synchronize, False, pid)
    if not handle:
        return False
    try:
        return bool(kernel32.TerminateProcess(handle, 1))
    finally:
        kernel32.CloseHandle(handle)


def _diagnostic_text(value: str, limit: int = 4000) -> str:
    clean = "\n".join(line for line in value.splitlines() if not line.lstrip().startswith("{"))
    return clean[-limit:]


def _text(value: str | bytes | None) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return value


def worker_module_path() -> Path:
    return Path(__file__).with_name("worker.py")
