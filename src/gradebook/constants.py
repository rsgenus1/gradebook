"""Constantes compartidas por el motor.

Nombres de hojas, encabezados, escalas de notas y colores viven aquí para que
no haya textos "mágicos" repetidos por el código.
"""

from __future__ import annotations

from decimal import Decimal

# --- Escala de notas -------------------------------------------------------

MIN_GRADE = Decimal("1.0")
MAX_GRADE = Decimal("7.0")
PASSING_GRADE = Decimal("4.0")
GRADE_STEP = Decimal("0.1")

# --- Hojas, tabla y columnas ----------------------------------------------

SHEET_GRADES = "Grades"
SHEET_STATS = "Stats"
SHEET_FINAL = "Final"
TABLE_NAME = "tblGrades"

COL_NAME = "Nombre"
COL_RUT = "RUT"
COL_ACCUMULATIVE_GRADE = "Nota Acumulativa"
COL_FINAL = "Promedio Final"

ACCUMULATIVE_PREFIX = "Acumulativa"
EVALUATION_PREFIX = "Evaluación"
DEFAULT_N_CONTROLS = 6
MAX_N_CONTROLS = 30

# --- Validación de notas ---------------------------------------------------

GRADE_ERROR_TITLE = "Nota inválida"
GRADE_ERROR_MESSAGE = (
    f"Ingresa una nota entre {MIN_GRADE} y {MAX_GRADE}, "
    "o deja la celda vacía si el estudiante no tiene nota."
)

# --- Formato numérico ------------------------------------------------------

FMT_GRADE = "0.0"
FMT_PERCENT = "0%"
FMT_PERCENT_ONE_DECIMAL = "0.0%"
FMT_INTEGER = "0"
FMT_TEXT = "@"

# --- Colores (ARGB/hex tal como los entiende xlsxwriter) -------------------

COLOR_HEADER_BG = "#1F3864"          # columnas de entrada
COLOR_HEADER_FONT = "#FFFFFF"
COLOR_CALC_HEADER_BG = "#375623"     # columnas calculadas
COLOR_WEIGHT_BG = "#D9E2F3"
COLOR_WEIGHT_FONT = "#1F3864"
COLOR_RED_FONT = "#C00000"
COLOR_TITLE = "#1F3864"
COLOR_SUBTITLE = "#595959"

COLOR_APPROVED = "#2E7D32"
COLOR_APPROVED_LIGHT = "#66BB6A"
COLOR_APPROVED_DARK = "#1B5E20"
COLOR_FAILED_ORANGE = "#EF6C00"
COLOR_FAILED_RED = "#C62828"
COLOR_FAILED_DARK = "#8E1B1B"

TABLE_STYLE = "Table Style Medium 2"
