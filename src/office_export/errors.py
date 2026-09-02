from __future__ import annotations

from dataclasses import dataclass
from typing import Any

EXIT_OK = 0
EXIT_USAGE = 2
EXIT_INPUT = 3
EXIT_CAPABILITY = 4
EXIT_SECURITY = 5
EXIT_CONVERSION = 6
EXIT_TIMEOUT = 7
EXIT_INTERNAL = 8


@dataclass(slots=True)
class ErrorContext:
    code: str
    message: str
    exit_code: int
    details: dict[str, Any] | None = None

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {"code": self.code, "message": self.message}
        if self.details:
            payload["details"] = self.details
        return payload


class OfficeExportError(Exception):
    def __init__(
        self,
        code: str,
        message: str,
        *,
        exit_code: int,
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self.context = ErrorContext(code=code, message=message, exit_code=exit_code, details=details)


class UsageError(OfficeExportError):
    def __init__(self, message: str, *, code: str = "usage_error", details: dict[str, Any] | None = None) -> None:
        super().__init__(code, message, exit_code=EXIT_USAGE, details=details)


class InputError(OfficeExportError):
    def __init__(self, code: str, message: str, *, details: dict[str, Any] | None = None) -> None:
        super().__init__(code, message, exit_code=EXIT_INPUT, details=details)


class CapabilityError(OfficeExportError):
    def __init__(self, code: str, message: str, *, details: dict[str, Any] | None = None) -> None:
        super().__init__(code, message, exit_code=EXIT_CAPABILITY, details=details)


class SecurityError(OfficeExportError):
    def __init__(self, code: str, message: str, *, details: dict[str, Any] | None = None) -> None:
        super().__init__(code, message, exit_code=EXIT_SECURITY, details=details)


class ConversionError(OfficeExportError):
    def __init__(self, code: str, message: str, *, details: dict[str, Any] | None = None) -> None:
        super().__init__(code, message, exit_code=EXIT_CONVERSION, details=details)


class WorkerTimeoutError(OfficeExportError):
    def __init__(self, message: str, *, details: dict[str, Any] | None = None) -> None:
        super().__init__("office_timeout", message, exit_code=EXIT_TIMEOUT, details=details)
