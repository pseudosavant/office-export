from __future__ import annotations

import pytest

from office_export.errors import UsageError
from office_export.selectors import (
    contiguous_ranges,
    is_contiguous,
    parse_chart_selector,
    parse_number_selection,
    parse_sheet_selector,
)


def test_number_selection_parses_sorts_and_defaults() -> None:
    assert parse_number_selection(None, label="pages", count=3) == [1, 2, 3]
    assert parse_number_selection("all", label="pages", count=2) == [1, 2]
    assert parse_number_selection("4,1-2", label="pages", count=4) == [1, 2, 4]


@pytest.mark.parametrize("value", ["", "1,,2", "two", "3-1", "0", "1,1", "1-2,2", "9"])
def test_number_selection_rejects_invalid_values(value: str) -> None:
    with pytest.raises(UsageError):
        parse_number_selection(value, label="slides", count=4)


def test_contiguous_ranges() -> None:
    assert contiguous_ranges([1, 2, 4, 5, 8]) == [(1, 2), (4, 5), (8, 8)]
    assert is_contiguous([3, 4, 5]) is True
    assert is_contiguous([3, 5]) is False


def test_excel_selectors() -> None:
    assert parse_sheet_selector(" 2 ") == 2
    assert parse_sheet_selector(" Summary ") == "Summary"
    assert parse_chart_selector("Summary!Revenue Chart") == ("Summary", "Revenue Chart")
    with pytest.raises(UsageError):
        parse_chart_selector("Revenue Chart")
