"""Modelos de dominio del motor y excepciones propias.

Todo lo que está aquí es Python puro: no depende de Streamlit ni de ninguna
otra interfaz.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field

# --- Errores de dominio ----------------------------------------------------


class GradebookError(Exception):
    """Error base del motor."""


class InputFormatError(GradebookError):
    """La planilla de entrada no se puede leer o no tiene los datos mínimos."""


class InvalidSchemeError(GradebookError):
    """La configuración de ponderaciones no es válida."""


# --- Avisos (datos, no excepciones) ---------------------------------------


class WarningCode:
    """Códigos estables para que la UI decida cómo presentar cada aviso."""

    NAME_EMPTY = "nombre_vacio"
    RUT_EMPTY = "rut_vacio"
    RUT_DUPLICATE = "rut_duplicado"
    RUT_INVALID = "rut_invalido"


@dataclass(frozen=True, slots=True)
class Warning:
    """Aviso no bloqueante asociado a una fila de la planilla."""

    code: str
    message: str
    row: int | None = None

    def __str__(self) -> str:
        if self.row is None:
            return self.message
        return f"Fila {self.row}: {self.message}"


# --- Estudiantes -----------------------------------------------------------


@dataclass(frozen=True, slots=True)
class Student:
    name: str
    rut: str  # normalizado: sin puntos, sin guion, mayúsculas


@dataclass(frozen=True, slots=True)
class StudentTable:
    """Estudiantes normalizados más los avisos detectados al leer la planilla."""

    students: tuple[Student, ...]
    warnings: tuple[Warning, ...] = ()
    columns: tuple[str, ...] = ()
    name_col: str | None = None
    rut_col: str | None = None

    def __len__(self) -> int:
        return len(self.students)

    def __iter__(self) -> Iterable[Student]:
        return iter(self.students)

    def warnings_by_code(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for warning in self.warnings:
            counts[warning.code] = counts.get(warning.code, 0) + 1
        return counts


@dataclass(frozen=True, slots=True)
class ColumnGuess:
    """Resultado de autodetectar las columnas de nombre y RUT."""

    columns: tuple[str, ...]
    name_col: str | None = None
    rut_col: str | None = None
    notes: Mapping[str, str] = field(default_factory=dict)
