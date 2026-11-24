"""Construcción del workbook Excel con las hojas ``Grades`` y ``Stats``.

El archivo solo contiene fórmulas vivas: la app no calcula notas. Todo lo que
depende del esquema de ponderación se ubica en celdas visibles para que el
profesor pueda editarlo en Excel.
"""

from __future__ import annotations

import io
from dataclasses import dataclass
from decimal import Decimal

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

from . import constants as C
from .grading import GradingScheme, weight_fraction
from .models import InputFormatError, StudentTable

# Filas base de la hoja Grades (0-based).
_WEIGHT_ROW = 0
_HEADER_ROW = 1
_FIRST_DATA_ROW = 2

_GRADE_DECIMALS = abs(C.GRADE_STEP.as_tuple().exponent)

# Filas base de la hoja Stats (0-based).
_STATS_TITLE_ROW = 0
_STATS_SUBTITLE_ROW = 1
_STATS_HEADER_ROW = 3
_STATS_FIRST_ROW = 4

_AUX_DIST_TITLE_ROW = 12
_AUX_DIST_HEADER_ROW = 13
_AUX_DIST_FIRST_ROW = 14

_AUX_PIE_TITLE_ROW = 21
_AUX_PIE_HEADER_ROW = 22
_AUX_PIE_FIRST_ROW = 23

_CHART_ANCHOR_ROW = 27
_CHART_WIDTH = 460
_CHART_HEIGHT = 300

_STAT_ROWS: tuple[tuple[str, str, str], ...] = (
    ("max", "Máximo", "grade"),
    ("min", "Mínimo", "grade"),
    ("count", "Asistentes", "integer"),
    ("failed", "Reprobados", "integer"),
    ("failed_pct", "% Reprobados", "percent"),
    ("median", "Mediana", "grade"),
    ("stdev", "Desviación estándar", "stdev"),
)


@dataclass(frozen=True)
class _Layout:
    """Ubicación de cada columna en la hoja Grades."""

    headers: tuple[str, ...]
    evaluation_cols: tuple[int, ...]
    control_cols: tuple[int, ...]
    accumulative_col: int | None
    final_col: int

    @property
    def input_cols(self) -> tuple[int, ...]:
        return self.evaluation_cols + self.control_cols

    @property
    def component_cols(self) -> tuple[int, ...]:
        if self.accumulative_col is None:
            return self.evaluation_cols
        return self.evaluation_cols + (self.accumulative_col,)

    @property
    def grade_cols(self) -> tuple[int, ...]:
        return self.component_cols + (self.final_col,)

    @property
    def last_col(self) -> int:
        return len(self.headers) - 1


def _build_layout(scheme: GradingScheme) -> _Layout:
    headers: list[str] = [C.COL_NAME, C.COL_RUT]
    column = 2
    evaluation_cols: list[int] = []
    for evaluation in scheme.evaluations:
        headers.append(evaluation.name)
        evaluation_cols.append(column)
        column += 1

    control_cols: list[int] = []
    accumulative_col: int | None = None
    if scheme.has_controls:
        for name in scheme.control_names:
            headers.append(name)
            control_cols.append(column)
            column += 1
        headers.append(C.COL_ACCUMULATIVE_GRADE)
        accumulative_col = column
        column += 1

    headers.append(C.COL_FINAL)
    final_col = column
    return _Layout(
        headers=tuple(headers),
        evaluation_cols=tuple(evaluation_cols),
        control_cols=tuple(control_cols),
        accumulative_col=accumulative_col,
        final_col=final_col,
    )


# --- Formatos --------------------------------------------------------------


class _Formats:
    """Todos los formatos del workbook, creados una sola vez."""

    def __init__(self, workbook: xlsxwriter.Workbook) -> None:
        base = {"font_name": "Calibri", "font_size": 11}

        self.name = workbook.add_format({**base, "align": "left", "valign": "vcenter", "indent": 1})
        self.rut = workbook.add_format(
            {**base, "num_format": C.FMT_TEXT, "align": "center", "valign": "vcenter"}
        )
        self.grade = workbook.add_format({**base, "num_format": C.FMT_GRADE, "align": "center"})
        self.calc_grade = workbook.add_format(
            {**base, "num_format": C.FMT_GRADE, "align": "center", "italic": True}
        )

        self.header = workbook.add_format(
            {
                **base,
                "bold": True,
                "font_color": C.COLOR_HEADER_FONT,
                "bg_color": C.COLOR_HEADER_BG,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
            }
        )
        self.header_calc = workbook.add_format(
            {
                **base,
                "bold": True,
                "font_color": C.COLOR_HEADER_FONT,
                "bg_color": C.COLOR_CALC_HEADER_BG,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
            }
        )
        self.weight_label = workbook.add_format(
            {
                **base,
                "bold": True,
                "font_color": C.COLOR_WEIGHT_FONT,
                "bg_color": C.COLOR_WEIGHT_BG,
                "align": "left",
                "indent": 1,
            }
        )
        self.group_label = workbook.add_format(
            {
                **base,
                "bold": True,
                "font_color": C.COLOR_WEIGHT_FONT,
                "bg_color": C.COLOR_WEIGHT_BG,
                "align": "center",
            }
        )
        self.weight = workbook.add_format(
            {
                **base,
                "bold": True,
                "num_format": C.FMT_PERCENT,
                "font_color": C.COLOR_WEIGHT_FONT,
                "bg_color": C.COLOR_WEIGHT_BG,
                "align": "center",
            }
        )
        self.check_ok = workbook.add_format(
            {
                **base,
                "bold": True,
                "font_color": C.COLOR_APPROVED,
                "bg_color": C.COLOR_WEIGHT_BG,
                "align": "center",
            }
        )
        self.check_bad = workbook.add_format(
            {**base, "bold": True, "font_color": C.COLOR_RED_FONT, "bg_color": "#FCE4E4"}
        )
        self.red = workbook.add_format({**base, "font_color": C.COLOR_RED_FONT})

        self.title = workbook.add_format(
            {**base, "bold": True, "font_size": 15, "font_color": C.COLOR_TITLE}
        )
        self.subtitle = workbook.add_format(
            {**base, "italic": True, "font_color": C.COLOR_SUBTITLE}
        )
        self.section = workbook.add_format(
            {**base, "bold": True, "font_size": 11, "font_color": C.COLOR_TITLE}
        )
        self.stat_label = workbook.add_format(
            {
                **base,
                "bold": True,
                "align": "left",
                "indent": 1,
                "bg_color": "#F2F2F2",
                "border": 1,
                "border_color": "#BFBFBF",
            }
        )
        self.stat_integer = workbook.add_format(
            {**base, "num_format": C.FMT_INTEGER, "align": "center", "border": 1, "border_color": "#BFBFBF"}
        )
        self.stat_grade = workbook.add_format(
            {**base, "num_format": C.FMT_GRADE, "align": "center", "border": 1, "border_color": "#BFBFBF"}
        )
        self.stat_stdev = workbook.add_format(
            {**base, "num_format": "0.00", "align": "center", "border": 1, "border_color": "#BFBFBF"}
        )
        self.stat_percent = workbook.add_format(
            {
                **base,
                "num_format": C.FMT_PERCENT_ONE_DECIMAL,
                "align": "center",
                "border": 1,
                "border_color": "#BFBFBF",
            }
        )
        self.aux_header = workbook.add_format(
            {
                **base,
                "bold": True,
                "bg_color": C.COLOR_WEIGHT_BG,
                "font_color": C.COLOR_WEIGHT_FONT,
                "align": "center",
            }
        )
        self.aux_label = workbook.add_format({**base, "align": "left", "indent": 1})
        self.aux_value = workbook.add_format({**base, "num_format": C.FMT_INTEGER, "align": "center"})


# --- Utilidades ------------------------------------------------------------


def _cell(row: int, column: int, *, absolute_row: bool = False) -> str:
    letter = xl_col_to_name(column)
    return f"{letter}${row + 1}" if absolute_row else f"{letter}{row + 1}"


def _literal(value: Decimal) -> str:
    """Decimal → texto que Excel entiende como número (sin notación científica)."""
    text = format(value.normalize(), "f")
    return text.removesuffix(".0") if text.endswith(".0") else text


def _escape_table_header(name: str) -> str:
    special = set("#[]'@")
    return "".join(f"'{char}" if char in special else char for char in name)


def _table_ref(header: str) -> str:
    return f"{C.TABLE_NAME}[{_escape_table_header(header)}]"


def _final_formula(row: int, layout: _Layout) -> str:
    cells = [(column, _cell(row, column)) for column in layout.component_cols]
    weight_terms = " + ".join(
        f"{xl_col_to_name(column)}${_WEIGHT_ROW + 1}*"
        f"IF(ISNUMBER({cell}),{cell},{_literal(C.MIN_GRADE)})"
        for column, cell in cells
    )
    count_args = ",".join(cell for _, cell in cells)
    return f'=IF(COUNT({count_args})=0,"",ROUND({weight_terms},{_GRADE_DECIMALS}))'


def _accumulative_formula(row: int, layout: _Layout) -> str:
    first = xl_col_to_name(layout.control_cols[0])
    last = xl_col_to_name(layout.control_cols[-1])
    row_number = row + 1
    return (
        f'=IF(COUNT({first}{row_number}:{last}{row_number})=0,"",'
        f"ROUND(AVERAGE({first}{row_number}:{last}{row_number}),{_GRADE_DECIMALS}))"
    )


# --- API interna -----------------------------------------------------------


def build_workbook(students: StudentTable, scheme: GradingScheme) -> bytes:
    """Genera el ``.xlsx`` (hojas Grades y Stats) y lo devuelve como bytes."""
    scheme.validate()
    if not students.students:
        raise InputFormatError("Se necesita al menos un estudiante para generar la planilla.")

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(
        output,
        {"in_memory": True, "use_future_functions": True},
    )
    workbook.set_properties(
        {
            "title": "Planilla de notas",
            "subject": "Notas y estadísticas del curso",
            "comments": "Generado con organizador de planilla de notas",
        }
    )
    formats = _Formats(workbook)
    layout = _build_layout(scheme)

    _write_grades(workbook, formats, layout, students, scheme)
    _write_stats(workbook, formats, layout, scheme)

    workbook.close()
    return output.getvalue()


def _write_grades(
    workbook: xlsxwriter.Workbook,
    fmt: _Formats,
    layout: _Layout,
    students: StudentTable,
    scheme: GradingScheme,
) -> None:
    sheet = workbook.add_worksheet(C.SHEET_GRADES)
    sheet.hide_gridlines(2)
    sheet.activate()

    last_data_row = _FIRST_DATA_ROW + len(students.students) - 1

    _setup_grades_columns(sheet, fmt, layout, students)
    _write_weight_row(sheet, fmt, layout, scheme)
    _write_grades_table(sheet, fmt, layout, students)

    # Notas de entrada: decimal entre 1 y 7, vacío permitido.
    first_input_col = layout.input_cols[0]
    last_input_col = layout.input_cols[-1]
    sheet.data_validation(
        _FIRST_DATA_ROW,
        first_input_col,
        last_data_row,
        last_input_col,
        {
            "validate": "decimal",
            "criteria": "between",
            "minimum": float(C.MIN_GRADE),
            "maximum": float(C.MAX_GRADE),
            "ignore_blank": True,
            "input_title": "Nota",
            "input_message": (
                f"Escala {C.MIN_GRADE} – {C.MAX_GRADE}. "
                "Deja vacío si el estudiante no tiene nota."
            ),
            "error_title": C.GRADE_ERROR_TITLE,
            "error_message": C.GRADE_ERROR_MESSAGE,
        },
    )

    # Notas bajo 4.0 en rojo (se compara el valor ya redondeado a 1 decimal,
    # para que el color coincida con lo que se muestra en pantalla).
    for column in layout.grade_cols:
        letter = xl_col_to_name(column)
        sheet.conditional_format(
            _FIRST_DATA_ROW,
            column,
            last_data_row,
            column,
            {
                "type": "formula",
                "criteria": (
                    f"=AND(ISNUMBER({letter}{_FIRST_DATA_ROW + 1}),"
                    f"ROUND({letter}{_FIRST_DATA_ROW + 1},{_GRADE_DECIMALS})<"
                    f"{_literal(C.PASSING_GRADE)})"
                ),
                "format": fmt.red,
            },
        )

    sheet.freeze_panes(_FIRST_DATA_ROW, 2)
    sheet.set_row(_WEIGHT_ROW, 20)
    sheet.set_row(_HEADER_ROW, 34)
    sheet.set_zoom(90)

    sheet.set_landscape()
    sheet.set_paper(9)  # A4
    # Una sola página (LibreOffice escala hacia arriba si solo se limita el
    # ancho, y una planilla angosta terminaría repartida en dos páginas).
    sheet.fit_to_pages(1, 1)
    sheet.repeat_rows(_WEIGHT_ROW, _HEADER_ROW)
    sheet.center_horizontally()
    sheet.set_margins(0.25, 0.25, 0.5, 0.5)
    sheet.set_footer("&C&P de &N")


def _setup_grades_columns(
    sheet: xlsxwriter.worksheet.Worksheet,
    fmt: _Formats,
    layout: _Layout,
    students: StudentTable,
) -> None:
    longest_name = max(
        (len(student.name) for student in students.students),
        default=len(C.COL_NAME),
    )
    sheet.set_column(0, 0, min(max(longest_name + 3, 18), 45), fmt.name)
    sheet.set_column(1, 1, 16, fmt.rut)

    input_width = max(len(layout.headers[col]) for col in layout.input_cols) + 4
    sheet.set_column(
        layout.input_cols[0], layout.input_cols[-1], max(input_width, 12), fmt.grade
    )

    calc_start = (
        layout.accumulative_col if layout.accumulative_col is not None else layout.final_col
    )
    calc_width = max(
        len(layout.headers[col]) for col in layout.grade_cols if col >= calc_start
    )
    sheet.set_column(calc_start, layout.final_col, max(calc_width + 4, 12), fmt.calc_grade)


def _write_weight_row(
    sheet: xlsxwriter.worksheet.Worksheet,
    fmt: _Formats,
    layout: _Layout,
    scheme: GradingScheme,
) -> None:
    sheet.write_string(_WEIGHT_ROW, 0, "Pesos", fmt.weight_label)
    for column, weight in zip(layout.component_cols, scheme.component_weights, strict=True):
        sheet.write_number(_WEIGHT_ROW, column, float(weight_fraction(weight)), fmt.weight)

    if layout.control_cols:
        first_control = layout.control_cols[0]
        last_control = layout.control_cols[-1]
        if first_control == last_control:
            sheet.write_string(_WEIGHT_ROW, first_control, "Controles", fmt.group_label)
        else:
            sheet.merge_range(
                _WEIGHT_ROW, first_control, _WEIGHT_ROW, last_control, "Controles", fmt.group_label
            )

    first = xl_col_to_name(layout.component_cols[0])
    last = xl_col_to_name(layout.component_cols[-1])
    weights = f"{first}${_WEIGHT_ROW + 1}:{last}${_WEIGHT_ROW + 1}"
    formula = (
        f'=IF(ABS(SUM({weights})-1)<0.0001,"OK 100%","Suma = "&TEXT(SUM({weights}),"0.0%"))'
    )
    sheet.write_formula(_WEIGHT_ROW, layout.final_col, formula, fmt.check_ok, "OK 100%")

    sheet.conditional_format(
        _WEIGHT_ROW,
        layout.final_col,
        _WEIGHT_ROW,
        layout.final_col,
        {
            "type": "formula",
            "criteria": f"=ABS(SUM({weights})-1)>0.0001",
            "format": fmt.check_bad,
        },
    )


def _write_grades_table(
    sheet: xlsxwriter.worksheet.Worksheet,
    fmt: _Formats,
    layout: _Layout,
    students: StudentTable,
) -> None:
    last_data_row = _FIRST_DATA_ROW + len(students.students) - 1

    for offset, student in enumerate(students.students):
        row = _FIRST_DATA_ROW + offset
        sheet.write_string(row, 0, student.name, fmt.name)
        sheet.write_string(row, 1, student.rut, fmt.rut)
        if layout.accumulative_col is not None:
            sheet.write_formula(
                row,
                layout.accumulative_col,
                _accumulative_formula(row, layout),
                fmt.calc_grade,
                "",
            )
        sheet.write_formula(row, layout.final_col, _final_formula(row, layout), fmt.calc_grade, "")

    calculated_headers = {C.COL_ACCUMULATIVE_GRADE, C.COL_FINAL}
    columns = [
        {
            "header": header,
            "header_format": fmt.header_calc if header in calculated_headers else fmt.header,
        }
        for header in layout.headers
    ]
    sheet.add_table(
        _HEADER_ROW,
        0,
        last_data_row,
        layout.last_col,
        {
            "name": C.TABLE_NAME,
            "style": C.TABLE_STYLE,
            "banded_rows": True,
            "banded_columns": False,
            "first_column": False,
            "autofilter": True,
            "columns": columns,
        },
    )

    for column in (*_calculated_cols(layout),):
        sheet.write_comment(
            _HEADER_ROW,
            column,
            "Columna calculada: no editar. "
            "Usa las columnas de la izquierda para ingresar las notas.",
            {"x_scale": 2.0, "y_scale": 0.8},
        )


def _calculated_cols(layout: _Layout) -> tuple[int, ...]:
    columns = [layout.final_col]
    if layout.accumulative_col is not None:
        columns.insert(0, layout.accumulative_col)
    return tuple(columns)


# --- Hoja Stats ------------------------------------------------------------


def _write_stats(
    workbook: xlsxwriter.Workbook,
    fmt: _Formats,
    layout: _Layout,
    scheme: GradingScheme,
) -> None:
    sheet = workbook.add_worksheet(C.SHEET_STATS)
    sheet.hide_gridlines(2)

    stat_columns: list[tuple[str, int]] = [
        (evaluation.name, column)
        for evaluation, column in zip(scheme.evaluations, layout.evaluation_cols, strict=True)
    ]
    if layout.accumulative_col is not None:
        stat_columns.append((C.COL_ACCUMULATIVE_GRADE, layout.accumulative_col))
    stat_columns.append((C.COL_FINAL, layout.final_col))

    last_value_col = len(stat_columns)
    # El subtítulo necesita algo de ancho también en configuraciones angostas.
    merge_col = max(last_value_col, 7)

    sheet.set_column(0, 0, 26)
    sheet.set_column(1, max(last_value_col, 7), 15)

    sheet.merge_range(
        _STATS_TITLE_ROW,
        0,
        _STATS_TITLE_ROW,
        merge_col,
        "Estadísticas del curso",
        fmt.title,
    )
    sheet.merge_range(
        _STATS_SUBTITLE_ROW,
        0,
        _STATS_SUBTITLE_ROW,
        merge_col,
        "Fórmulas vivas: se actualizan solas al editar la hoja "
        f"{C.SHEET_GRADES}. Las celdas en blanco cuentan como inasistencia.",
        fmt.subtitle,
    )

    sheet.write_string(_STATS_HEADER_ROW, 0, "Estadística", fmt.header)
    for offset, (header, column) in enumerate(stat_columns, start=1):
        is_calculated = column in (layout.accumulative_col, layout.final_col)
        sheet.write_string(
            _STATS_HEADER_ROW,
            offset,
            header,
            fmt.header_calc if is_calculated else fmt.header,
        )

    assistants_row = _STATS_FIRST_ROW + 2
    failed_row = _STATS_FIRST_ROW + 3
    passing = _literal(C.PASSING_GRADE)

    for offset, (header, _column) in enumerate(stat_columns, start=1):
        reference = _table_ref(header)
        letter = xl_col_to_name(offset)
        assistants_cell = f"{letter}{assistants_row + 1}"
        failed_cell = f"{letter}{failed_row + 1}"

        for index, (key, label, kind) in enumerate(_STAT_ROWS):
            row = _STATS_FIRST_ROW + index
            if offset == 1:
                sheet.write_string(row, 0, label, fmt.stat_label)
            formula = _stat_formula(key, reference, assistants_cell, failed_cell, passing)
            sheet.write_formula(row, offset, formula, _stat_format(fmt, kind), "")

    _write_distribution_aux(sheet, fmt)
    _write_approval_aux(sheet, fmt)

    _add_bar_chart(workbook, sheet)
    _add_pie_chart(workbook, sheet)

    sheet.set_landscape()
    sheet.set_paper(9)
    sheet.fit_to_pages(1, 1)
    sheet.set_margins(0.3, 0.3, 0.5, 0.5)


def _stat_formula(
    key: str,
    reference: str,
    assistants_cell: str,
    failed_cell: str,
    passing: str,
) -> str:
    return {
        "max": f'=IF(COUNT({reference})=0,"",MAX({reference}))',
        "min": f'=IF(COUNT({reference})=0,"",MIN({reference}))',
        "count": f"=COUNT({reference})",
        "failed": f'=COUNTIF({reference},"<{passing}")',
        "failed_pct": f'=IFERROR({failed_cell}/{assistants_cell},"")',
        "median": f'=IFERROR(MEDIAN({reference}),"")',
        "stdev": f'=IFERROR(STDEV.S({reference}),"")',
    }[key]


def _stat_format(fmt: _Formats, kind: str) -> xlsxwriter.format.Format:
    return {
        "grade": fmt.stat_grade,
        "integer": fmt.stat_integer,
        "percent": fmt.stat_percent,
        "stdev": fmt.stat_stdev,
    }[kind]


def _write_distribution_aux(
    sheet: xlsxwriter.worksheet.Worksheet,
    fmt: _Formats,
) -> None:
    final_ref = _table_ref(C.COL_FINAL)
    sheet.write_string(_AUX_DIST_TITLE_ROW, 0, "Distribución de Promedio Final", fmt.section)
    sheet.write_string(_AUX_DIST_HEADER_ROW, 0, "Intervalo", fmt.aux_header)
    sheet.write_string(_AUX_DIST_HEADER_ROW, 1, "Estudiantes", fmt.aux_header)

    for index, low in enumerate(range(1, 7)):
        row = _AUX_DIST_FIRST_ROW + index
        label = f"[{low},{low + 1}]" if low == 6 else f"[{low},{low + 1}["
        if low == 6:
            formula = f'=COUNTIFS({final_ref},">=6",{final_ref},"<=7")'
        else:
            formula = f'=COUNTIFS({final_ref},">={low}",{final_ref},"<{low + 1}")'
        sheet.write_string(row, 0, label, fmt.aux_label)
        sheet.write_formula(row, 1, formula, fmt.aux_value, "")


def _write_approval_aux(
    sheet: xlsxwriter.worksheet.Worksheet,
    fmt: _Formats,
) -> None:
    final_ref = _table_ref(C.COL_FINAL)
    passing = _literal(C.PASSING_GRADE)
    rows = (
        ("Aprobado (≥ 4.0)", f'=COUNTIF({final_ref},">={passing}")'),
        ("Reprobado < 3.0", f'=COUNTIF({final_ref},"<3")'),
        ("Reprobado 3.0 – 3.9", f'=COUNTIFS({final_ref},">=3",{final_ref},"<{passing}")'),
    )
    sheet.write_string(_AUX_PIE_TITLE_ROW, 0, "Aprobación del curso", fmt.section)
    sheet.write_string(_AUX_PIE_HEADER_ROW, 0, "Categoría", fmt.aux_header)
    sheet.write_string(_AUX_PIE_HEADER_ROW, 1, "Estudiantes", fmt.aux_header)
    for index, (label, formula) in enumerate(rows):
        row = _AUX_PIE_FIRST_ROW + index
        sheet.write_string(row, 0, label, fmt.aux_label)
        sheet.write_formula(row, 1, formula, fmt.aux_value, "")


def _add_bar_chart(
    workbook: xlsxwriter.Workbook,
    sheet: xlsxwriter.worksheet.Worksheet,
) -> None:
    chart = workbook.add_chart({"type": "column"})
    first = _AUX_DIST_FIRST_ROW
    last = _AUX_DIST_FIRST_ROW + 5
    chart.add_series(
        {
            "name": "Estudiantes",
            "categories": [C.SHEET_STATS, first, 0, last, 0],
            "values": [C.SHEET_STATS, first, 1, last, 1],
            "data_labels": {
                "value": True,
                "category": False,
                "series_name": False,
                "legend_key": False,
                "font": {"size": 10, "bold": True},
            },
            "points": [
                {"fill": {"color": C.COLOR_FAILED_DARK}},
                {"fill": {"color": C.COLOR_FAILED_RED}},
                {"fill": {"color": C.COLOR_FAILED_ORANGE}},
                {"fill": {"color": C.COLOR_APPROVED_LIGHT}},
                {"fill": {"color": C.COLOR_APPROVED}},
                {"fill": {"color": C.COLOR_APPROVED_DARK}},
            ],
        }
    )
    chart.set_title({"name": "Distribución de Promedio Final", "name_font": {"size": 13, "bold": True}})
    chart.set_x_axis(
        {
            "name": "Nota",
            "name_font": {"size": 10},
            "num_font": {"size": 9},
            "major_tick_mark": "none",
        }
    )
    chart.set_y_axis(
        {
            "name": "Cantidad de estudiantes",
            "name_font": {"size": 10},
            "num_font": {"size": 9},
            "major_gridlines": {"visible": False},
            "min": 0,
            "major_unit": 1,
            "major_tick_mark": "none",
        }
    )
    chart.set_legend({"none": True})
    chart.set_chartarea({"border": {"none": True}})
    chart.set_plotarea({"border": {"none": True}})
    chart.set_size({"width": _CHART_WIDTH, "height": _CHART_HEIGHT})
    sheet.insert_chart(f"A{_CHART_ANCHOR_ROW + 1}", chart, {"x_offset": 6, "y_offset": 6})


def _add_pie_chart(
    workbook: xlsxwriter.Workbook,
    sheet: xlsxwriter.worksheet.Worksheet,
) -> None:
    chart = workbook.add_chart({"type": "pie"})
    first = _AUX_PIE_FIRST_ROW
    last = _AUX_PIE_FIRST_ROW + 2
    chart.add_series(
        {
            "name": "Aprobación",
            "categories": [C.SHEET_STATS, first, 0, last, 0],
            "values": [C.SHEET_STATS, first, 1, last, 1],
            "data_labels": {
                "percentage": True,
                "category": False,
                "series_name": False,
                "value": False,
                "legend_key": False,
                "font": {"size": 11, "bold": True},
            },
            "points": [
                {"fill": {"color": C.COLOR_APPROVED}},
                {"fill": {"color": C.COLOR_FAILED_DARK}},
                {"fill": {"color": C.COLOR_FAILED_ORANGE}},
            ],
        }
    )
    chart.set_title({"name": "Aprobación del curso", "name_font": {"size": 13, "bold": True}})
    chart.set_legend({"position": "right", "font": {"size": 9}})
    chart.set_chartarea({"border": {"none": True}})
    chart.set_plotarea({"border": {"none": True}})
    chart.set_size({"width": _CHART_WIDTH, "height": _CHART_HEIGHT})
    sheet.insert_chart(f"F{_CHART_ANCHOR_ROW + 1}", chart, {"x_offset": 6, "y_offset": 6})


__all__ = ["build_workbook"]
