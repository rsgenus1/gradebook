"""Genera planillas de ejemplo (entrada y salida) con datos sintéticos.

Los RUT son inventados: el dígito verificador se calcula con módulo 11. Nunca
se usan datos reales.

Uso:

    python tests/make_samples.py                 # escribe en tests/samples/
    python tests/make_samples.py --outdir otro/
"""

from __future__ import annotations

import argparse
import io
import random
from collections.abc import Sequence
from pathlib import Path

import openpyxl
import xlsxwriter

from gradebook import (
    PRESETS,
    Evaluation,
    GradingScheme,
    StudentTable,
    build_workbook,
    read_students,
)
from gradebook.rut import compute_check_digit

FIRST_NAMES = [
    "Ana", "José", "María", "Nicolás", "Sofía", "Matías", "Valentina", "Benjamín",
    "Florencia", "Ignacio", "Catalina", "Sebastián", "Antonia", "Diego", "Josefa",
    "Tomás", "Camila", "Vicente", "Martina", "Joaquín", "Fernanda", "Cristóbal",
    "Isidora", "Maximiliano", "Trinidad", "Felipe", "Amanda", "Agustín", "Emilia",
    "Renato", "Paz", "Bruno", "Maite", "Constanza", "Bárbara", "Andrés",
]

LAST_NAMES = [
    "Ñúñez", "Peña", "Muñoz", "Rojas", "González", "Cárcamo", "Ibáñez", "O'Higgins",
    "Valdés", "Riquelme", "Bustamante", "Araya", "Sepúlveda", "Cifuentes", "Jara",
    "Leiva", "Paredes", "Vergara", "Tapia", "Molina",
]

# Notas de borde pedidas en AGENTS.md.
EDGE_GRADES = [3.55, 3.95, 5.05, 6.95, 1.0, 7.0, 3.9, 4.0]

GRADE_POOL = [
    1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 3.9, 4.0, 4.5, 5.0, 5.5, 6.0, 6.5, 6.95, 7.0
]


# --- Datos de entrada ------------------------------------------------------


def make_rut(body: int, style: int = 0) -> str:
    """RUT sintético con dígito verificador correcto y formato variable."""
    text = str(body)
    digit = compute_check_digit(text)
    grouped = f"{body:,}".replace(",", ".")
    variants = (
        f"{grouped}-{digit}",
        f"{text}-{digit}",
        f"{text}{digit}",
        f" {grouped.lower()}-{digit.lower()} ",
        f"  {text}-{digit}",
    )
    return variants[style % len(variants)]


def synthetic_rows(count: int = 35, seed: int = 7) -> list[dict[str, str]]:
    rng = random.Random(seed)
    bodies = rng.sample(range(9_000_000, 30_000_000), count)
    rows: list[dict[str, str]] = []
    for index, body in enumerate(bodies):
        first = FIRST_NAMES[index % len(FIRST_NAMES)]
        last = LAST_NAMES[(index * 3) % len(LAST_NAMES)]
        rows.append(
            {
                "Nombre Completo": f"{first} {last}",
                "RUT": make_rut(body, index),
                "Curso": "8° Básico B",
            }
        )
    return rows


def rows_with_warnings() -> list[dict[str, str]]:
    """Incluye un RUT duplicado, uno vacío y uno con dígito inválido."""
    return [
        {"Nombre Completo": "Ana Ñúñez", "RUT": "12.345.678-5"},
        {"Nombre Completo": "Beto Soto", "RUT": "12.345.678-5"},
        {"Nombre Completo": "Carla Díaz", "RUT": ""},
        {"Nombre Completo": "Diego Rojas", "RUT": "11.111.111-1"},
        {"Nombre Completo": "", "RUT": "20.000.000-1"},
    ]


def write_input_xlsx(rows: Sequence[dict[str, str]], path: Path) -> None:
    workbook = xlsxwriter.Workbook(str(path))
    worksheet = workbook.add_worksheet("Curso")
    headers = list(rows[0].keys()) if rows else ["Nombre Completo", "RUT"]
    bold = workbook.add_format({"bold": True})
    worksheet.write_row(0, 0, headers, bold)
    for index, row in enumerate(rows, start=1):
        for column, header in enumerate(headers):
            value = row.get(header, "")
            worksheet.write_string(index, column, value)
    workbook.close()


# --- Notas sintéticas ------------------------------------------------------


def build_grades(
    scheme: GradingScheme,
    count: int,
    seed: int = 11,
    blank_rate: float = 0.15,
    blank_evaluation: int | None = None,
) -> list[dict[str, list[float | None]]]:
    """Genera notas por estudiante con blancos, casos borde y faltantes."""
    rng = random.Random(seed)
    n_evaluations = len(scheme.evaluations)
    n_controls = len(scheme.control_names)

    grades: list[dict[str, list[float | None]]] = []
    for index in range(count):
        evaluations: list[float | None] = []
        for position in range(n_evaluations):
            if blank_evaluation is not None and position == blank_evaluation or rng.random() < blank_rate:
                evaluations.append(None)
            else:
                evaluations.append(rng.choice(GRADE_POOL))

        controls: list[float | None] = [
            None if rng.random() < blank_rate else rng.choice(GRADE_POOL)
            for _ in range(n_controls)
        ]

        if index == 0 and n_evaluations >= 3:
            evaluations = [3.55, 3.95, 5.05, *evaluations[3:]]
            controls = [
                6.95, 1.0, 7.0, 3.9, 4.0, None, *controls[6:]
            ][:n_controls]
        if index == 1 and n_evaluations >= 2:
            evaluations[1] = None  # evaluación faltante → cuenta 1.0 en el final
        if index == 2 and n_controls:
            controls = [None] * n_controls  # sin controles → acumulativa vacía
        if index == count - 1:
            evaluations = [None] * n_evaluations
            controls = [None] * n_controls

        grades.append({"evaluations": evaluations, "controls": controls})
    return grades


def fill_workbook(
    students: StudentTable,
    scheme: GradingScheme,
    grades: Sequence[dict[str, list[float | None]]],
    path: Path,
) -> None:
    """Escribe notas de ejemplo dentro del workbook generado por el motor."""
    workbook = openpyxl.load_workbook(io.BytesIO(build_workbook(students, scheme)))
    sheet = workbook["Grades"]

    header_row = 2  # fila de encabezados (1-based)
    column_by_header = {
        sheet.cell(row=header_row, column=column).value: column
        for column in range(1, sheet.max_column + 1)
    }

    for offset, grade in enumerate(grades):
        row = header_row + 1 + offset
        for evaluation, value in zip(scheme.evaluations, grade["evaluations"], strict=True):
            if value is not None:
                sheet.cell(row=row, column=column_by_header[evaluation.name]).value = value
        for name, value in zip(scheme.control_names, grade["controls"], strict=True):
            if value is not None:
                sheet.cell(row=row, column=column_by_header[name]).value = value

    workbook.save(path)


# --- Escenarios ------------------------------------------------------------


def custom_scheme() -> GradingScheme:
    return GradingScheme(
        (Evaluation("Proyecto", "40"), Evaluation("Examen final", "60")),
        controls_weight="0",
        n_controls=0,
    )


def generate(outdir: Path) -> list[Path]:
    outdir.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []

    main_rows = synthetic_rows(35)
    input_path = outdir / "entrada_curso.xlsx"
    write_input_xlsx(main_rows, input_path)
    written.append(input_path)

    warnings_path = outdir / "entrada_avisos.xlsx"
    write_input_xlsx(rows_with_warnings(), warnings_path)
    written.append(warnings_path)

    main_students = read_students(input_path.read_bytes())
    single_student = StudentTable(main_students.students[:1])
    scenarios: list[tuple[str, StudentTable, GradingScheme, int, int]] = [
        ("curso_3eval.xlsx", main_students, PRESETS["3_evaluaciones"], 11, 0),
        ("curso_2eval.xlsx", main_students, PRESETS["2_evaluaciones"], 21, 0),
        ("curso_personalizado.xlsx", main_students, custom_scheme(), 31, 0),
        ("curso_1estudiante.xlsx", single_student, PRESETS["3_evaluaciones"], 41, 0),
        ("curso_sin_asistentes.xlsx", main_students, PRESETS["3_evaluaciones"], 51, 1),
    ]

    for filename, students, scheme, seed, blank_evaluation in scenarios:
        grades = build_grades(
            scheme,
            len(students.students),
            seed=seed,
            blank_evaluation=blank_evaluation,
        )
        path = outdir / filename
        fill_workbook(students, scheme, grades, path)
        written.append(path)

    return written


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--outdir",
        type=Path,
        default=Path(__file__).resolve().parent / "samples",
        help="carpeta de salida (por defecto tests/samples)",
    )
    args = parser.parse_args(argv)
    for path in generate(args.outdir):
        print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
