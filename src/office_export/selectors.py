from __future__ import annotations

from collections.abc import Iterable

from office_export.errors import UsageError


def parse_number_selection(
    value: str | None,
    *,
    label: str,
    count: int | None = None,
    default_all: bool = True,
) -> list[int] | None:
    if value is None:
        if count is not None and default_all:
            return list(range(1, count + 1))
        return None
    normalized = value.strip().lower()
    if normalized == "all":
        if count is None:
            return None
        return list(range(1, count + 1))
    if not normalized:
        raise _syntax_error(label)

    selected: list[int] = []
    seen: set[int] = set()
    for raw_part in normalized.split(","):
        part = raw_part.strip()
        if not part:
            raise _syntax_error(label)
        if "-" in part:
            bounds = part.split("-")
            if len(bounds) != 2 or not all(bound.isdigit() for bound in bounds):
                raise _syntax_error(label)
            start, end = (int(bound) for bound in bounds)
            if start > end:
                raise UsageError(f"--{label} range '{part}' is descending. Use {end}-{start} instead.")
            numbers: Iterable[int] = range(start, end + 1)
        elif part.isdigit():
            numbers = (int(part),)
        else:
            raise _syntax_error(label)
        for number in numbers:
            if number < 1:
                raise UsageError(f"--{label} uses one-based numbers. Zero is not valid.")
            if count is not None and number > count:
                raise UsageError(f"--{label} contains {number}, but the source has {count} item(s).")
            if number in seen:
                raise UsageError(f"--{label} selects {number} more than once.")
            selected.append(number)
            seen.add(number)
    return sorted(selected)


def contiguous_ranges(numbers: list[int]) -> list[tuple[int, int]]:
    if not numbers:
        return []
    values = sorted(set(numbers))
    ranges: list[tuple[int, int]] = []
    start = previous = values[0]
    for number in values[1:]:
        if number == previous + 1:
            previous = number
            continue
        ranges.append((start, previous))
        start = previous = number
    ranges.append((start, previous))
    return ranges


def is_contiguous(numbers: list[int]) -> bool:
    return len(contiguous_ranges(numbers)) <= 1


def parse_sheet_selector(value: str) -> int | str:
    normalized = value.strip()
    if not normalized:
        raise UsageError("--sheet cannot be empty.")
    if normalized.isdigit():
        number = int(normalized)
        if number < 1:
            raise UsageError("--sheet uses one-based indices. Zero is not valid.")
        return number
    return normalized


def parse_chart_selector(value: str) -> tuple[str, str]:
    sheet, separator, name = value.partition("!")
    if not separator or not sheet.strip() or not name.strip():
        raise UsageError("--chart must use SHEET!NAME syntax, such as Summary!Revenue Chart.")
    return sheet.strip(), name.strip()


def _syntax_error(label: str) -> UsageError:
    return UsageError(f"--{label} must use a comma-separated one-based list such as 1,3-5.")
