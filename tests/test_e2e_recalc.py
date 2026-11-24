"""Cálculo end-to-end: generar, llenar, recalcular con LibreOffice y comparar.

Si LibreOffice no está disponible (o no puede arrancar en un entorno
restringido), los tests se saltan en vez de fallar.
"""

from __future__ import annotations

import shutil
import statistics
import subprocess
from pathlib import Path

import pytest
from make_samples import fill_workbook
from openpyxl import load_workbook

from gradebook import PRESETS, Student, StudentTable
from gradebook.grading import accumulative_average, final_from_scheme
from gradebook.rut import compute_check_digit

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")
SCHEME = PRESETS["3_evaluaciones"]

pytestmark = pytest.mark.skipif(SOFFICE is None, reason="LibreOffice no está instalado")


def make_students(count: int = 5) -> StudentTable:
    students = []
    for index in range(1, count + 1):
        body = str(10_000_000 + index)
        students.append(Student(f"Estudiante {index}", body + compute_check_digit(body)))
    return StudentTable(tuple(students))


def recalc(source: Path, outdir: Path) -> Path:
    outdir.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [
            SOFFICE,
            "--headless",
            "--norestore",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(outdir),
            str(source),
        ],
        capture_output=True,
        text=True,
        timeout=240,
    )
    target = outdir / source.name
    if result.returncode != 0 or not target.exists():
        pytest.skip(f"LibreOffice no pudo recalcular el archivo: {result.stderr or result.stdout}")
    return target


def value_or_none(raw: object) -> float | None:
    if raw in (None, ""):
        return None
    return float(raw)


def assert_close(actual: float | None, expected: float | None) -> None:
    if expected is None:
        assert actual is None, f"se esperaba vacío y llegó {actual!r}"
    else:
        assert actual == pytest.approx(expected)


def expected_stats(values: list[float | None]) -> dict[str, float | None]:
    numbers = [value for value in values if value is not None]
    if not numbers:
        return {"max": None, "min": None, "count": 0, "failed": None, "pct": None,
                "median": None, "stdev": None}
    failed = sum(1 for value in numbers if value < 4.0)
    return {
        "max": max(numbers),
        "min": min(numbers),
        "count": len(numbers),
        "failed": failed,
        "pct": failed / len(numbers),
        "median": statistics.median(numbers),
        "stdev": statistics.stdev(numbers) if len(numbers) >= 2 else None,
    }


def assert_no_errors(workbook) -> None:
    errors = []
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("#"):
                    errors.append(f"{sheet.title}!{cell.coordinate}={cell.value}")
    assert errors == []


def test_final_grades_and_all_statistics_match_python(tmp_path: Path) -> None:
    students = make_students(5)
    grades = [
        {"evaluations": [5.5, 4.0, 6.0], "controls": [3.55, 4.0, 6.95, None, 1.0, 7.0]},
        {"evaluations": [4.0, 4.0, 3.95], "controls": [5.05, 3.9, None, None, None, None]},
        {"evaluations": [6.0, 7.0, 5.05], "controls": [None] * 6},
        {"evaluations": [None, None, None], "controls": [None] * 6},
        {"evaluations": [1.0, 1.0, 1.0], "controls": [1.0] * 6},
    ]
    source = tmp_path / "curso.xlsx"
    fill_workbook(students, SCHEME, grades, source)

    workbook = load_workbook(recalc(source, tmp_path / "recalc"), data_only=True)
    assert_no_errors(workbook)
    sheet = workbook["Grades"]

    accum_values: list[float | None] = []
    final_values: list[float | None] = []
    for offset, grade in enumerate(grades):
        row = 3 + offset
        expected_accum = accumulative_average(grade["controls"])
        expected_final = final_from_scheme(
            SCHEME, evaluations=grade["evaluations"], controls=grade["controls"]
        )
        assert_close(
            value_or_none(sheet.cell(row=row, column=12).value),
            None if expected_accum is None else float(expected_accum),
        )
        assert_close(
            value_or_none(sheet.cell(row=row, column=13).value),
            None if expected_final is None else float(expected_final),
        )
        accum_values.append(None if expected_accum is None else float(expected_accum))
        final_values.append(None if expected_final is None else float(expected_final))

    stats = workbook["Final"]
    evaluation_values = [
        [grade["evaluations"][index] for grade in grades] for index in range(3)
    ]
    column_values = [*evaluation_values, accum_values, final_values]

    for column_offset, values in enumerate(column_values):
        column = 2 + column_offset  # B..F
        expected = expected_stats(values)
        assert_close(value_or_none(stats.cell(row=5, column=column).value), expected["max"])
        assert_close(value_or_none(stats.cell(row=6, column=column).value), expected["min"])
        assert stats.cell(row=7, column=column).value == expected["count"]
        assert stats.cell(row=8, column=column).value == expected["failed"]
        assert_close(value_or_none(stats.cell(row=9, column=column).value), expected["pct"])
        assert_close(value_or_none(stats.cell(row=10, column=column).value), expected["median"])
        assert_close(value_or_none(stats.cell(row=11, column=column).value), expected["stdev"])


def test_evaluation_without_students_does_not_break(tmp_path: Path) -> None:
    students = make_students(3)
    grades = [
        {"evaluations": [5.0, None, 5.0], "controls": [5.0] * 6},
        {"evaluations": [6.0, None, 6.0], "controls": [6.0] * 6},
        {"evaluations": [4.0, None, 4.0], "controls": [4.0] * 6},
    ]
    source = tmp_path / "sin_asistentes.xlsx"
    fill_workbook(students, SCHEME, grades, source)

    workbook = load_workbook(recalc(source, tmp_path / "recalc"), data_only=True)
    assert_no_errors(workbook)
    stats = workbook["Final"]

    assert value_or_none(stats["C5"].value) in (None, 0.0)  # MAX de columna vacía
    assert value_or_none(stats["C6"].value) is None          # MIN
    assert stats["C7"].value == 0                            # asistentes
    assert stats["C8"].value == 0                            # reprobados
    assert value_or_none(stats["C9"].value) is None          # división por cero
    assert value_or_none(stats["C11"].value) is None         # desviación con 0 datos


def test_single_student_has_no_division_errors(tmp_path: Path) -> None:
    students = make_students(1)
    grades = [{"evaluations": [5.5, 6.0, 4.5], "controls": [5.0, 4.0, 6.0, 5.0, 4.5, 6.0]}]
    source = tmp_path / "un_estudiante.xlsx"
    fill_workbook(students, SCHEME, grades, source)

    workbook = load_workbook(recalc(source, tmp_path / "recalc"), data_only=True)
    assert_no_errors(workbook)
    stats = workbook["Final"]

    expected_final = final_from_scheme(
        SCHEME, evaluations=grades[0]["evaluations"], controls=grades[0]["controls"]
    )
    assert value_or_none(workbook["Grades"]["M3"].value) == pytest.approx(
        float(expected_final)
    )
    assert value_or_none(stats["F11"].value) is None  # STDEV.S con un solo dato
