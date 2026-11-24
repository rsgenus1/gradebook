from __future__ import annotations

import pytest

from gradebook.rut import (
    compute_check_digit,
    format_rut,
    has_check_digit,
    is_valid_rut,
    normalize_rut,
    split_rut,
)


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("12.345.678-5", "123456785"),
        ("9.876.543-k", "9876543K"),
        (" 12345678-5 ", "123456785"),
        ("12345678K", "12345678K"),
        ("12 345 678 5", "123456785"),
        ("  9876543-k  ", "9876543K"),
        ("", ""),
        (None, ""),
        (12345678, "12345678"),
    ],
)
def test_normalize_rut(raw: object, expected: str) -> None:
    assert normalize_rut(raw) == expected


def test_normalize_is_idempotent() -> None:
    once = normalize_rut("9.876.543-k")
    assert normalize_rut(once) == once


@pytest.mark.parametrize(
    ("body", "digit"),
    [
        ("12345678", "5"),
        ("9876543", "3"),
        ("11111111", "1"),
        ("76086428", "5"),
        ("8", "6"),
    ],
)
def test_compute_check_digit(body: str, digit: str) -> None:
    assert compute_check_digit(body) == digit


def test_compute_check_digit_rejects_non_numeric() -> None:
    with pytest.raises(ValueError):
        compute_check_digit("12K")


@pytest.mark.parametrize(("normalized", "expected"), [
    ("123456785", True),
    ("98765433", True),
    ("111111111", True),
    ("123456789", False),
    ("12345678K", False),
    ("", True),          # vacío: no hay dígito que validar
    ("12345678", True),  # incompleto, no inválido
])
def test_is_valid_rut(normalized: str, expected: bool) -> None:
    assert is_valid_rut(normalized) is expected


def test_split_and_has_check_digit() -> None:
    assert split_rut("123456785") == ("12345678", "5")
    assert split_rut("9876543K") == ("9876543", "K")
    assert split_rut("12345678") == ("12345678", "")
    assert has_check_digit("123456785")
    assert not has_check_digit("12345678")
    assert not has_check_digit("")


def test_format_rut() -> None:
    assert format_rut("123456785") == "12.345.678-5"
    assert format_rut("9876543K") == "9.876.543-K"
    assert format_rut("") == ""
