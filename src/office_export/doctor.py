from __future__ import annotations

import os
import platform
import struct
import sys
import tempfile
from pathlib import Path
from typing import Any

from office_export import __version__
from office_export.results import runtime_versions
from office_export.worker_protocol import run_worker


def run_doctor(*, timeout: float = 30.0) -> dict[str, Any]:
    result: dict[str, Any] = {
        "platform": {
            "system": platform.system(),
            "release": platform.release(),
            "version": platform.version(),
            "python": platform.python_version(),
            "architecture": platform.machine(),
        },
        "tool_version": __version__,
        "dependencies": runtime_versions(),
        "temporary_directory": _temporary_directory_check(),
        "office_export": {"available": sys.platform == "win32", "applications": {}},
        "pdf_rasterization": {
            "available": runtime_versions()["pdfium"] is not None and runtime_versions()["pillow"] is not None,
            "office_required": False,
        },
        "printing": _printing_info(),
        "warnings": [],
    }
    if sys.platform != "win32":
        result["office_export"]["reason"] = "Office export requires Windows."
        return result

    for application in ("word", "excel", "powerpoint"):
        executable = _office_executable(application)
        item: dict[str, Any] = {
            "registered_executable": str(executable) if executable else None,
            "bitness": _pe_bitness(executable) if executable else None,
        }
        try:
            response = run_worker({"action": "capabilities", "application": application}, timeout)
            item.update(
                {
                    "available": True,
                    "application": response["application"],
                    "capabilities": response["capabilities"],
                }
            )
        except Exception as exc:
            context = getattr(exc, "context", None)
            item.update(
                {
                    "available": False,
                    "error": context.to_dict() if context else {"code": "probe_failed", "message": str(exc)},
                }
            )
        result["office_export"]["applications"][application] = item
    result["office_export"]["available"] = any(
        item.get("available") for item in result["office_export"]["applications"].values()
    )
    if not result["printing"].get("printers"):
        result["warnings"].append(
            {
                "code": "no_printer_detected",
                "message": "No installed printer was detected. Word and Excel pagination may fail.",
            }
        )
    return result


def _temporary_directory_check() -> dict[str, Any]:
    try:
        with tempfile.TemporaryDirectory(prefix="office-export-doctor-") as directory:
            probe = Path(directory) / "write-test"
            probe.write_bytes(b"ok")
            return {"writable": probe.read_bytes() == b"ok", "path": str(Path(directory).parent)}
    except OSError as exc:
        return {"writable": False, "error": f"{type(exc).__name__}: {exc}"}


def _printing_info() -> dict[str, Any]:
    if sys.platform != "win32":
        return {"available": False, "spooler": None, "printers": [], "microsoft_print_to_pdf": False}
    spooler: str | None = None
    printers: list[str] = []
    try:
        import win32serviceutil

        status = win32serviceutil.QueryServiceStatus("Spooler")[1]
        spooler = {1: "stopped", 2: "start_pending", 3: "stop_pending", 4: "running"}.get(status, str(status))
    except Exception:
        spooler = "unknown"
    try:
        import win32print

        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        discovered = win32print.EnumPrinters(flags, None, 2)
        printers = sorted(
            {
                str(item.get("pPrinterName")) if isinstance(item, dict) else str(item[2])
                for item in discovered
                if (item.get("pPrinterName") if isinstance(item, dict) else item[2])
            },
            key=str.casefold,
        )
    except Exception:
        pass
    return {
        "available": spooler == "running",
        "spooler": spooler,
        "printers": printers,
        "microsoft_print_to_pdf": any(name.casefold() == "microsoft print to pdf" for name in printers),
    }


def _office_executable(application: str) -> Path | None:
    names = {"word": "WINWORD.EXE", "excel": "EXCEL.EXE", "powerpoint": "POWERPNT.EXE"}
    try:
        import winreg

        key_path = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{names[application]}"
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
                try:
                    with winreg.OpenKey(hive, key_path, 0, winreg.KEY_READ | view) as key:
                        value, _ = winreg.QueryValueEx(key, "")
                        path = Path(os.path.expandvars(value))
                        if path.is_file():
                            return path
                except OSError:
                    continue
    except ImportError:
        return None
    return None


def _pe_bitness(path: Path) -> str | None:
    try:
        with path.open("rb") as stream:
            if stream.read(2) != b"MZ":
                return None
            stream.seek(0x3C)
            offset = struct.unpack("<I", stream.read(4))[0]
            stream.seek(offset + 4)
            machine = struct.unpack("<H", stream.read(2))[0]
    except (OSError, struct.error):
        return None
    return {0x014C: "32-bit", 0x8664: "64-bit", 0xAA64: "ARM64"}.get(machine, f"machine-0x{machine:04X}")
