"""Normalización y validación (módulo 11) de RUT chilenos.

La normalización es siempre silenciosa: no bloquea datos. La validación del
dígito verificador solo se usa para generar avisos en la UI.
"""

from __future__ import annotations

import re

# Caracteres que separan miles o el dígito verificador en formatos habituales.
_SEPARATORS = re.compile(r"[\s.\-_/\u00a0\u2013\u2014]")

_DV_FACTORS = (2, 3, 4, 5, 6, 7)


def normalize_rut(raw: object) -> str:
    """Devuelve el RUT sin puntos, guiones ni espacios, y en mayúsculas.

    >>> normalize_rut("12.345.678-5")
    '123456785'
    >>> normalize_rut("9.876.543-k")
    '9876543K'
    >>> normalize_rut(" 12345678-5 ")
    '123456785'
    """
    if raw is None:
        return ""
    text = _SEPARATORS.sub("", str(raw))
    return text.strip().upper()


def compute_check_digit(body: str) -> str:
    """Calcula el dígito verificador de un cuerpo numérico con módulo 11."""
    if not body or not body.isdigit():
        raise ValueError(f"Cuerpo de RUT inválido: {body!r}")
    total = 0
    for index, digit in enumerate(reversed(body)):
        total += int(digit) * _DV_FACTORS[index % len(_DV_FACTORS)]
    remainder = 11 - (total % 11)
    if remainder == 11:
        return "0"
    if remainder == 10:
        return "K"
    return str(remainder)


def split_rut(normalized: str) -> tuple[str, str]:
    """Separa cuerpo y dígito verificador.

    Un RUT sin dígito verificador devuelve ``(rut, "")``: el dato simplemete
    está incompleto, no es inválido.
    """
    if not normalized:
        return "", ""
    last = normalized[-1]
    if last == "K":
        return normalized[:-1], last
    if last.isdigit() and len(normalized) >= 9:
        return normalized[:-1], last
    return normalized, ""


def has_check_digit(normalized: str) -> bool:
    body, dv = split_rut(normalized)
    return bool(body) and bool(dv)


def is_valid_rut(normalized: str) -> bool:
    """Valida el dígito verificador. Un RUT incompleto no se marca inválido."""
    body, dv = split_rut(normalized)
    if not dv:
        return True
    if not body.isdigit():
        return False
    return compute_check_digit(body) == dv


def format_rut(normalized: str) -> str:
    """Formatea un RUT normalizado con puntos y guion, para mostrar al usuario."""
    body, dv = split_rut(normalized)
    if not body:
        return normalized
    chunks: list[str] = []
    while len(body) > 3:
        chunks.insert(0, body[-3:])
        body = body[:-3]
    chunks.insert(0, body)
    grouped = ".".join(chunks)
    return f"{grouped}-{dv}" if dv else grouped
