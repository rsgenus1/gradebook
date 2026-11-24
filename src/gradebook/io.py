"""Lectura de la planilla de entrada (.xlsx / .csv) en memoria.

Todo se lee como texto para no perder ceros a la izquierda ni la ``K`` de un
RUT. La API no acepta rutas: recibe ``bytes`` o un file-like.
"""

from __future__ import annotations

import io
import re
import unicodedata
from collections.abc import Sequence
from typing import Any, BinaryIO

import fastexcel
import polars as pl

from .constants import COL_NAME, COL_RUT
from .models import ColumnGuess, InputFormatError, Student, StudentTable, Warning, WarningCode
from .rut import format_rut, is_valid_rut, normalize_rut

Source = bytes | bytearray | memoryview | BinaryIO

_XLSX_MAGIC = b"PK\x03\x04"
_XLS_MAGIC = b"\xd0\xcf\x11\xe0"

_HEADER_ROW = 1  # la primera fila es el encabezado
_FIRST_DATA_ROW = 2

_NAME_HINTS = (
    "nombre",
    "nombres",
    "nombre completo",
    "alumno",
    "alumna",
    "estudiante",
    "name",
    "student",
)
_RUT_HINTS = (
    "rut",
    "run",
    "cedula",
    "cedula identidad",
    "documento",
    "dni",
    "id",
    "identificacion",
)

_RUT_VALUE_RE = re.compile(r"^\d{6,9}[0-9K]?$")
_NAME_VALUE_RE = re.compile(r"^[^\W\d_][\w .,'’-]*$", re.UNICODE)


# --- Entrada pública -------------------------------------------------------


def detect_columns(source: Source) -> ColumnGuess:
    """Adivina qué columnas son nombre y RUT mirando encabezados y valores."""
    columns, rows = _read_table(source)
    return _guess_columns(columns, rows)


def read_students(
    source: Source,
    *,
    name_col: str | None = None,
    rut_col: str | None = None,
) -> StudentTable:
    """Lee la planilla y devuelve estudiantes normalizados más avisos.

    ``source`` puede ser ``bytes`` o un objeto con ``.read()``. Nunca se
    escriben datos personales a disco ni a logs.
    """
    columns, rows = _read_table(source)
    guess = _guess_columns(columns, rows)

    name_col = name_col or guess.name_col
    rut_col = rut_col or guess.rut_col

    missing = [
        label
        for label, column in (("nombre", name_col), ("RUT", rut_col))
        if not column
    ]
    if missing:
        available = ", ".join(columns) or "(sin columnas)"
        raise InputFormatError(
            "No pude identificar la columna de "
            + " ni ".join(missing)
            + f". Columnas disponibles: {available}."
        )
    assert name_col is not None and rut_col is not None
    if name_col not in columns:
        raise InputFormatError(f"La columna {name_col!r} no existe en la planilla.")
    if rut_col not in columns:
        raise InputFormatError(f"La columna {rut_col!r} no existe en la planilla.")
    if name_col == rut_col:
        raise InputFormatError("La columna de nombre y la de RUT no pueden ser la misma.")

    students: list[Student] = []
    warnings: list[Warning] = []
    first_row_by_rut: dict[str, int] = {}

    for row_number, row in rows:
        name = " ".join(row[name_col].split())
        rut = normalize_rut(row[rut_col])

        if not name:
            warnings.append(
                Warning(
                    WarningCode.NAME_EMPTY,
                    "El nombre está vacío en esta fila.",
                    row_number,
                )
            )
        if not rut:
            warnings.append(
                Warning(WarningCode.RUT_EMPTY, "El RUT está vacío en esta fila.", row_number)
            )
        elif not is_valid_rut(rut):
            warnings.append(
                Warning(
                    WarningCode.RUT_INVALID,
                    f"El RUT {format_rut(rut)} tiene un dígito verificador inválido "
                    "(se guarda igual).",
                    row_number,
                )
            )

        if rut:
            previous = first_row_by_rut.get(rut)
            if previous is not None:
                warnings.append(
                    Warning(
                        WarningCode.RUT_DUPLICATE,
                        f"El RUT {format_rut(rut)} está repetido "
                        f"(también aparece en la fila {previous}).",
                        row_number,
                    )
                )
            else:
                first_row_by_rut[rut] = row_number

        students.append(Student(name=name, rut=rut))

    if not students:
        raise InputFormatError("La planilla no contiene estudiantes.")

    return StudentTable(
        students=tuple(students),
        warnings=tuple(warnings),
        columns=columns,
        name_col=name_col,
        rut_col=rut_col,
    )


# --- Lectura cruda ---------------------------------------------------------


def _to_bytes(source: Source) -> bytes:
    if isinstance(source, bytes):
        return source
    if isinstance(source, (bytearray, memoryview)):
        return bytes(source)
    if hasattr(source, "read"):
        data = source.read()
        if isinstance(data, str):
            return data.encode("utf-8")
        if isinstance(data, (bytes, bytearray, memoryview)):
            return bytes(data)
        raise TypeError("El file-like debe entregar bytes")
    raise TypeError(
        "La entrada debe ser bytes o un file-like (no una ruta): "
        "lee el archivo en la capa que corresponda."
    )


def _read_table(source: Source) -> tuple[tuple[str, ...], list[tuple[int, dict[str, str]]]]:
    """Devuelve (encabezados, [(fila_excel, {columna: texto})])."""
    data = _to_bytes(source)
    if not data:
        raise InputFormatError("El archivo está vacío.")

    if data.startswith(_XLSX_MAGIC):
        frame = _read_xlsx(data)
    elif data.startswith(_XLS_MAGIC):
        raise InputFormatError(
            "El formato .xls antiguo no está soportado. Guarda la planilla como .xlsx o .csv."
        )
    else:
        frame = _read_csv(data)

    columns = _clean_headers(frame.columns)
    rows: list[tuple[int, dict[str, str]]] = []
    for index, values in enumerate(frame.iter_rows()):
        row = {
            column: _clean_text(value)
            for column, value in zip(columns, values, strict=False)
        }
        if any(row.values()):
            rows.append((index + _FIRST_DATA_ROW, row))
    return columns, rows


def _read_xlsx(data: bytes) -> pl.DataFrame:
    try:
        workbook = fastexcel.read_excel(io.BytesIO(data))
        if not workbook.sheet_names:
            raise InputFormatError("El archivo no tiene hojas.")
        sheet = workbook.load_sheet(0, dtypes="string", whitespace_as_null=True)
        return sheet.to_polars()
    except InputFormatError:
        raise
    except Exception as exc:  # noqa: BLE001 - se traduce a error de dominio
        raise InputFormatError(f"No pude leer el archivo Excel: {exc}") from exc


def _read_csv(data: bytes) -> pl.DataFrame:
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "latin-1"):
        try:
            text = data.decode(encoding)
        except UnicodeDecodeError:
            continue
        for separator in _candidate_separators(text):
            try:
                frame = pl.read_csv(
                    io.StringIO(text),
                    infer_schema=False,
                    separator=separator,
                    truncate_ragged_lines=True,
                    ignore_errors=True,
                )
            except Exception as exc:  # noqa: BLE001 - probamos otra separación
                last_error = exc
                continue
            if frame.height:
                return frame
    detail = f": {last_error}" if last_error else ""
    raise InputFormatError(f"No pude interpretar la planilla como CSV{detail}")


def _candidate_separators(text: str) -> tuple[str, ...]:
    sample = next((line for line in text.splitlines() if line.strip()), "")
    counts = {sep: sample.count(sep) for sep in (",", ";", "\t", "|")}
    ordered = sorted(counts, key=lambda sep: counts[sep], reverse=True)
    return tuple(sep for sep in ordered if counts[sep] > 0) or (",",)


def _clean_headers(columns: Sequence[str]) -> tuple[str, ...]:
    cleaned: list[str] = []
    seen: dict[str, int] = {}
    for index, column in enumerate(columns):
        name = " ".join(str(column).split()) if column is not None else ""
        if not name:
            name = f"Columna {index + 1}"
        key = name.casefold()
        if key in seen:
            seen[key] += 1
            name = f"{name} ({seen[key]})"
        else:
            seen[key] = 1
        cleaned.append(name)
    return tuple(cleaned)


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


# --- Detección de columnas -------------------------------------------------


def _ascii_fold(text: str) -> str:
    decomposed = unicodedata.normalize("NFKD", text)
    without_accents = "".join(ch for ch in decomposed if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", "", without_accents.casefold())


def _header_score(header: str, hints: Sequence[str]) -> int:
    normalized = _ascii_fold(header)
    if not normalized:
        return 0
    for hint in hints:
        key = _ascii_fold(hint)
        if normalized == key:
            return 3
    for hint in hints:
        key = _ascii_fold(hint)
        if normalized.startswith(key) or normalized.endswith(key):
            return 2
    for hint in hints:
        key = _ascii_fold(hint)
        if key and key in normalized:
            return 1
    return 0


def _column_values(rows: Sequence[tuple[int, dict[str, str]]], column: str) -> list[str]:
    values = [row[column].strip() for _, row in rows]
    return [value for value in values if value][:30]


def _looks_like_ruts(values: Sequence[str]) -> bool:
    if len(values) < 2:
        return False
    matches = sum(1 for value in values if _RUT_VALUE_RE.match(normalize_rut(value)))
    return matches / len(values) >= 0.6


def _looks_like_names(values: Sequence[str]) -> bool:
    if len(values) < 2:
        return False
    matches = sum(
        1
        for value in values
        if _NAME_VALUE_RE.match(value) and not _RUT_VALUE_RE.match(normalize_rut(value))
    )
    return matches / len(values) >= 0.6


def _guess_columns(
    columns: Sequence[str],
    rows: Sequence[tuple[int, dict[str, str]]],
) -> ColumnGuess:
    notes: dict[str, str] = {}
    scores: dict[str, int] = {column: _header_score(column, _NAME_HINTS) for column in columns}
    rut_scores: dict[str, int] = {column: _header_score(column, _RUT_HINTS) for column in columns}

    rut_col = max(rut_scores, key=lambda column: rut_scores[column], default=None)
    if rut_col is None or rut_scores[rut_col] == 0:
        rut_col = next(
            (
                column
                for column in columns
                if _looks_like_ruts(_column_values(rows, column))
            ),
            None,
        )
        if rut_col is not None:
            notes[rut_col] = "detectada por sus valores"
    else:
        notes[rut_col] = "detectada por el encabezado"
        # Si el encabezado es ambiguo, confirmamos con los valores.
        if rut_scores[rut_col] < 3 and not _looks_like_ruts(_column_values(rows, rut_col)):
            rut_col = next(
                (
                    column
                    for column in columns
                    if column != rut_col and _looks_like_ruts(_column_values(rows, column))
                ),
                rut_col,
            )

    name_col = max(scores, key=lambda column: scores[column], default=None)
    if name_col is None or scores[name_col] == 0 or name_col == rut_col:
        name_col = next(
            (
                column
                for column in columns
                if column != rut_col and _looks_like_names(_column_values(rows, column))
            ),
            None,
        )
        if name_col is not None:
            notes[name_col] = "detectada por sus valores"
    else:
        notes[name_col] = "detectada por el encabezado"

    return ColumnGuess(
        columns=tuple(columns),
        name_col=name_col,
        rut_col=rut_col,
        notes=notes,
    )


__all__ = ["detect_columns", "read_students", "COL_NAME", "COL_RUT"]
