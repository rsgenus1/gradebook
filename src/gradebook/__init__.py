"""API pública del motor de planillas de notas.

Uso típico:

    from gradebook import read_students, PRESETS, build_workbook

    students = read_students(open("entrada.xlsx", "rb").read())
    xlsx = build_workbook(students, PRESETS["3_evaluaciones"])

El motor es Python puro: no importa Streamlit ni ningún framework de UI.
"""

from __future__ import annotations

from .grading import (
    PRESET_LABELS,
    PRESETS,
    Evaluation,
    GradingScheme,
    accumulative_average,
    final_from_scheme,
    round_grade,
    weight_fraction,
    weighted_final,
)
from .io import detect_columns, read_students
from .models import (
    ColumnGuess,
    GradebookError,
    InputFormatError,
    InvalidSchemeError,
    Student,
    StudentTable,
    Warning,
    WarningCode,
)
from .rut import compute_check_digit, format_rut, is_valid_rut, normalize_rut
from .workbook import build_workbook

__version__ = "0.1.0"

__all__ = [
    "PRESET_LABELS",
    "PRESETS",
    "ColumnGuess",
    "Evaluation",
    "GradebookError",
    "GradingScheme",
    "InputFormatError",
    "InvalidSchemeError",
    "Student",
    "StudentTable",
    "Warning",
    "WarningCode",
    "accumulative_average",
    "build_workbook",
    "compute_check_digit",
    "detect_columns",
    "final_from_scheme",
    "format_rut",
    "is_valid_rut",
    "normalize_rut",
    "read_students",
    "round_grade",
    "weight_fraction",
    "weighted_final",
]
