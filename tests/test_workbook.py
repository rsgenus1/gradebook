from __future__ import annotations

import io

import pytest
from openpyxl import load_workbook
from openpyxl.chart import BarChart, PieChart
from openpyxl.worksheet.worksheet import Worksheet

from gradebook import (
    PRESETS,
    Evaluation,
    GradingScheme,
    InputFormatError,
    InvalidSchemeError,
    Student,
    StudentTable,
    build_workbook,
)

STUDENTS = StudentTable(
    (
        Student("Ana Ñúñez", "123456785"),
        Student("Beto Soto", "98765433"),
        Student("Carla Díaz", "111111111"),
    )
)
SCHEME = PRESETS["3_evaluaciones"]
HEADER_ROW = 2
FIRST_DATA_ROW = 3


def build(students: StudentTable = STUDENTS, scheme: GradingScheme = SCHEME) -> bytes:
    return build_workbook(students, scheme)


def load(data: bytes, *, values: bool = False):
    return load_workbook(io.BytesIO(data), data_only=values)


def header_column(sheet: Worksheet, header: str) -> int:
    for column in range(1, sheet.max_column + 1):
        if sheet.cell(row=HEADER_ROW, column=column).value == header:
            return column
    raise AssertionError(f"No se encontró la columna {header!r}")


def test_sheets_and_order() -> None:
    workbook = load(build())
    assert workbook.sheetnames == ["Grades", "Stats", "Final"]


def test_grades_headers_in_order() -> None:
    workbook = load(build())
    headers = [
        workbook["Grades"].cell(row=HEADER_ROW, column=column).value
        for column in range(1, workbook["Grades"].max_column + 1)
    ]
    assert headers == [
        "Nombre",
        "RUT",
        "Evaluación 1",
        "Evaluación 2",
        "Evaluación 3",
        "Acumulativa 1",
        "Acumulativa 2",
        "Acumulativa 3",
        "Acumulativa 4",
        "Acumulativa 5",
        "Acumulativa 6",
        "Nota Acumulativa",
        "Promedio Final",
    ]


def test_rut_is_stored_as_text() -> None:
    sheet = load(build())["Grades"]
    column = header_column(sheet, "RUT")
    cell = sheet.cell(row=FIRST_DATA_ROW, column=column)
    assert cell.value == "123456785"
    assert cell.data_type == "s"
    assert cell.number_format == "@"


def test_table_covers_all_students_and_headers() -> None:
    sheet = load(build())["Grades"]
    table = sheet.tables["tblGrades"]
    assert table.ref == f"A{HEADER_ROW}:M{HEADER_ROW + len(STUDENTS.students)}"
    assert table.tableStyleInfo.showRowStripes is True


def test_data_validation_allows_blank_and_1_to_7() -> None:
    sheet = load(build())["Grades"]
    validations = sheet.data_validations.dataValidation
    assert len(validations) == 1
    validation = validations[0]
    assert validation.type == "decimal"
    assert validation.operator in (None, "between")
    assert validation.formula1 == "1.0"
    assert validation.formula2 == "7.0"
    assert str(validation.sqref) == f"C{FIRST_DATA_ROW}:K{HEADER_ROW + len(STUDENTS.students)}"


def test_weights_live_in_cells_and_use_percent_format() -> None:
    sheet = load(build())["Grades"]
    assert sheet["A1"].value == "Pesos"
    for header, expected in [
        ("Evaluación 1", 0.2),
        ("Evaluación 2", 0.25),
        ("Evaluación 3", 0.25),
        ("Nota Acumulativa", 0.3),
    ]:
        cell = sheet.cell(row=1, column=header_column(sheet, header))
        assert cell.value == pytest.approx(expected)
        assert "%" in cell.number_format


def test_weight_check_cell_shows_ok() -> None:
    data = build()
    formula = load(data)["Grades"]["M1"].value
    cached = load(data, values=True)["Grades"]["M1"].value
    assert "OK 100%" in formula
    assert cached == "OK 100%"


def test_final_formula_references_weight_cells_and_min_grade() -> None:
    sheet = load(build())["Grades"]
    column = header_column(sheet, "Promedio Final")
    formula = sheet.cell(row=FIRST_DATA_ROW, column=column).value
    assert formula.startswith("=IF(COUNT(")
    assert "C$1" in formula and "L$1" in formula
    assert "IF(ISNUMBER(" in formula
    assert "ROUND(" in formula
    assert ",1))" in formula


def test_accumulative_formula_averages_only_present_controls() -> None:
    sheet = load(build())["Grades"]
    column = header_column(sheet, "Nota Acumulativa")
    formula = sheet.cell(row=FIRST_DATA_ROW, column=column).value
    assert "COUNT(F3:K3)=0" in formula
    assert "AVERAGE(F3:K3)" in formula
    assert "ROUND(" in formula


def test_calculated_columns_have_different_header_color_and_comment() -> None:
    sheet = load(build())["Grades"]
    input_header = sheet.cell(row=HEADER_ROW, column=header_column(sheet, "Evaluación 1"))
    accum_header = sheet.cell(row=HEADER_ROW, column=header_column(sheet, "Nota Acumulativa"))
    final_header = sheet.cell(row=HEADER_ROW, column=header_column(sheet, "Promedio Final"))

    assert input_header.fill.start_color.rgb != accum_header.fill.start_color.rgb
    assert accum_header.fill.start_color.rgb == final_header.fill.start_color.rgb
    for header_cell in (accum_header, final_header):
        comment = header_cell.comment
        assert comment is not None and "calculada" in comment.text.lower()


def test_conditional_formatting_for_low_grades() -> None:
    sheet = load(build())["Grades"]
    rules = [
        rule
        for conditional_format in sheet.conditional_formatting
        for rule in conditional_format.rules
    ]
    formulas = [str(rule.formula[0]) for rule in rules if rule.type == "expression"]
    assert any("ROUND" in formula and "<4" in formula for formula in formulas)


def test_final_sheet_has_summary_table() -> None:
    sheet = load(build())["Final"]
    labels = [sheet.cell(row=row, column=1).value for row in range(5, 12)]
    assert labels == [
        "Máximo",
        "Mínimo",
        "Asistentes",
        "Reprobados",
        "% Reprobados",
        "Mediana",
        "Desviación estándar",
    ]

    assert sheet["B5"].value.startswith("=IF(COUNT(tblGrades[Evaluación 1])")
    assert "MAX(tblGrades[Evaluación 1])" in sheet["B5"].value
    assert sheet["B7"].value == "=COUNT(tblGrades[Evaluación 1])"
    assert sheet["B8"].value == '=COUNTIF(tblGrades[Evaluación 1],"<4")'
    assert sheet["B9"].value == '=IFERROR(B8/B7,"")'
    assert "STDEV.S(tblGrades[Evaluación 1])" in sheet["B11"].value

    # Auxiliares de los gráficos del promedio final.
    assert "COUNTIFS" in sheet["B15"].value
    assert "COUNTIF" in sheet["B24"].value
    assert "COUNTIFS" in sheet["B26"].value


def test_stats_sheet_has_a_bar_and_pie_per_evaluation() -> None:
    sheet = load(build())["Stats"]

    # Una sección por evaluación: barras + torta cada una.
    assert len(sheet._charts) == 6
    for index in range(0, 6, 2):
        assert isinstance(sheet._charts[index], BarChart)
        assert isinstance(sheet._charts[index + 1], PieChart)

    # Auxiliares de cada sección, apuntando a su propia evaluación.
    assert "tblGrades[Evaluación 1]" in sheet["B6"].value
    assert "COUNTIFS" in sheet["B6"].value
    assert "tblGrades[Evaluación 1]" in sheet["E6"].value
    assert "tblGrades[Evaluación 2]" in sheet["B28"].value
    assert "tblGrades[Evaluación 3]" in sheet["B50"].value

    # Una página por evaluación; se ajusta solo el ancho para que la leyenda
    # de la torta no caiga en una página extra.
    assert len(sheet.row_breaks.brk) == 2
    page_props = sheet.sheet_properties.pageSetUpPr
    assert page_props is not None and page_props.fitToPage is True
    # xlsxwriter omite los atributos que quedan en su valor por defecto (1).
    assert sheet.page_setup.fitToWidth in (None, 1)
    assert sheet.page_setup.fitToHeight == 0


def test_final_sheet_has_two_charts() -> None:
    charts = load(build())["Final"]._charts
    assert len(charts) == 2
    assert isinstance(charts[0], BarChart)
    assert isinstance(charts[1], PieChart)


def test_freeze_panes_and_print_setup() -> None:
    sheet = load(build())["Grades"]
    assert sheet.freeze_panes == "C3"
    assert sheet.page_setup.orientation == "landscape"
    assert sheet.sheet_properties.pageSetUpPr.fitToPage is True
    assert sheet.print_title_rows == "$1:$2"


def test_custom_scheme_without_controls() -> None:
    scheme = GradingScheme(
        (Evaluation("Proyecto", "40"), Evaluation("Examen", "60")), "0", n_controls=0
    )
    sheet = load(build(scheme=scheme))["Grades"]
    headers = [
        sheet.cell(row=HEADER_ROW, column=column).value
        for column in range(1, sheet.max_column + 1)
    ]
    assert headers == ["Nombre", "RUT", "Proyecto", "Examen", "Promedio Final"]
    assert "Nota Acumulativa" not in headers
    assert sheet["E3"].value.startswith("=IF(COUNT(C3,D3)=0")


def test_single_student_works() -> None:
    sheet = load(build(students=StudentTable((Student("Ana", "123456785"),))))["Grades"]
    assert sheet["A3"].value == "Ana"
    assert sheet.max_row == 3


def test_empty_course_is_rejected() -> None:
    with pytest.raises(InputFormatError):
        build(students=StudentTable(()))


def test_invalid_scheme_is_rejected() -> None:
    scheme = GradingScheme((Evaluation("Uno", "90"),), "0", n_controls=0)
    with pytest.raises(InvalidSchemeError, match="100"):
        build(scheme=scheme)
