"""Adaptador Streamlit del motor ``gradebook``.

Aquí solo hay widgets y llamadas al motor. Ninguna regla de negocio: la
validación de ponderaciones, la normalización de RUT y las fórmulas viven en
``gradebook``.
"""

from __future__ import annotations

import sys
from decimal import Decimal
from pathlib import Path

import pandas as pd
import streamlit as st

# Permite `streamlit run ui/streamlit_app.py` desde el repositorio sin instalar.
_SRC = Path(__file__).resolve().parents[1] / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from gradebook import (  # noqa: E402
    PRESET_LABELS,
    PRESETS,
    ColumnGuess,
    Evaluation,
    GradingScheme,
    InputFormatError,
    InvalidSchemeError,
    StudentTable,
    WarningCode,
    build_workbook,
    detect_columns,
    format_rut,
    read_students,
)

_PRESET_KEYS = list(PRESETS)
_CUSTOM = "personalizada"

_WARNING_LABELS = {
    WarningCode.NAME_EMPTY: "Nombre vacío",
    WarningCode.RUT_EMPTY: "RUT vacío",
    WarningCode.RUT_DUPLICATE: "RUT duplicado",
    WarningCode.RUT_INVALID: "Dígito verificador inválido",
}


@st.cache_data(show_spinner=False, max_entries=8)
def _guess(data: bytes) -> ColumnGuess:
    return detect_columns(data)


@st.cache_data(show_spinner=False, max_entries=8)
def _students(data: bytes, name_col: str, rut_col: str) -> StudentTable:
    return read_students(data, name_col=name_col, rut_col=rut_col)


def _index_of(options: list[str], value: str | None) -> int:
    return options.index(value) if value in options else 0


def _render_warnings(table: StudentTable) -> None:
    if not table.warnings:
        st.success("Sin avisos: todos los datos se ven bien.")
        return

    counts = table.warnings_by_code()
    resumen = " · ".join(
        f"{_WARNING_LABELS.get(code, code)}: {count}" for code, count in sorted(counts.items())
    )
    st.warning(f"Revisa antes de descargar — {resumen}.")
    with st.expander("Ver detalle de los avisos"):
        for warning in table.warnings:
            st.write(f"- Fila {warning.row}: {warning.message}" if warning.row else f"- {warning.message}")


def _render_preview(table: StudentTable) -> None:
    records = [
        {
            "Fila": index + 2,
            "Nombre": student.name,
            "RUT normalizado": student.rut,
            "RUT visible": format_rut(student.rut),
        }
        for index, student in enumerate(table.students)
    ]
    st.dataframe(records, hide_index=True, use_container_width=True, height=320)


def _weight_editor(key: str) -> list[Evaluation]:
    st.caption("Cada fila es una evaluación regular. Puedes agregar o quitar filas.")
    default = pd.DataFrame(
        [
            {"Evaluación": "Evaluación 1", "Peso %": 50.0},
            {"Evaluación": "Evaluación 2", "Peso %": 50.0},
        ]
    )
    edited = st.data_editor(
        default,
        key=key,
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        column_config={
            "Evaluación": st.column_config.TextColumn("Evaluación", required=True),
            "Peso %": st.column_config.NumberColumn(
                "Peso %", min_value=0.0, max_value=100.0, step=0.5, format="%.1f"
            ),
        },
    )
    evaluations: list[Evaluation] = []
    for _, record in edited.iterrows():
        raw_name = record.get("Evaluación")
        name = "" if pd.isna(raw_name) else str(raw_name).strip()
        if not name:
            continue
        raw_weight = record.get("Peso %")
        weight = Decimal("0") if pd.isna(raw_weight) else Decimal(str(raw_weight))
        evaluations.append(Evaluation(name, weight))
    return evaluations


def _configuration() -> tuple[GradingScheme | None, str]:
    """Devuelve el esquema elegido y un mensaje cuando todavía no es válido."""
    choice = st.radio(
        "Configuración de ponderación",
        [*_PRESET_KEYS, _CUSTOM],
        format_func=lambda key: PRESET_LABELS.get(key, "Personalizada"),
        horizontal=True,
    )

    if choice != _CUSTOM:
        preset = PRESETS[choice]
        n_controls = st.number_input(
            "Cantidad de controles",
            min_value=1,
            max_value=30,
            value=preset.n_controls,
            step=1,
        )
        return (
            GradingScheme(preset.evaluations, preset.controls_weight, int(n_controls)),
            "",
        )

    controls_weight = Decimal(
        str(
            st.number_input(
                "Peso de la acumulativa (0 % para omitirla)",
                min_value=0.0,
                max_value=100.0,
                value=30.0,
                step=0.5,
            )
        )
    )
    evaluations = _weight_editor(key="pesos_personalizados")

    if controls_weight > 0:
        n_controls = int(
            st.number_input("Cantidad de controles", min_value=1, max_value=30, value=6, step=1)
        )
    else:
        n_controls = 0

    scheme = GradingScheme(tuple(evaluations), controls_weight, n_controls)
    try:
        scheme.validate()
    except InvalidSchemeError as exc:
        return scheme, str(exc)
    return scheme, ""


def main() -> None:
    st.set_page_config(page_title="Planilla de notas", page_icon="📋", layout="wide")
    st.title("Planilla de notas")
    st.write(
        "Sube la lista del curso (nombre y RUT). Descargarás un Excel con la hoja "
        "**Grades** para llenar las notas, la hoja **Stats** (barras y torta por "
        "cada evaluación) y la hoja **Final** (resumen y promedio final). Las "
        "estadísticas se actualizan solas."
    )

    uploaded = st.file_uploader("Planilla de entrada", type=["xlsx", "xls", "csv"])
    if uploaded is None:
        st.info(
            "Acepta `.xlsx`, `.xls` o `.csv` con al menos una columna de nombre y una de RUT."
        )
        return

    data = uploaded.getvalue()
    try:
        guess = _guess(data)
    except InputFormatError as exc:
        st.error(str(exc))
        return

    columns = list(guess.columns)
    left, right = st.columns(2)
    name_col = left.selectbox("Columna de nombre", columns, index=_index_of(columns, guess.name_col))
    rut_col = right.selectbox("Columna de RUT", columns, index=_index_of(columns, guess.rut_col))

    if name_col == rut_col:
        st.error("La columna de nombre y la de RUT no pueden ser la misma.")
        return

    try:
        table = _students(data, name_col, rut_col)
    except InputFormatError as exc:
        st.error(str(exc))
        return

    st.subheader(f"Vista previa ({len(table.students)} estudiantes)")
    _render_preview(table)
    _render_warnings(table)

    st.subheader("Ponderaciones")
    scheme, problem = _configuration()

    if scheme is not None:
        valid = not problem
        total = scheme.total_weight
        if valid:
            st.success(f"Los pesos suman {format(total, 'f')} %.")
        else:
            st.error(
                f"Los pesos no están listos: {problem}. "
                f"Suma actual: {format(total, 'f')} %."
            )

        if valid:
            try:
                workbook = build_workbook(table, scheme)
            except (InputFormatError, InvalidSchemeError) as exc:
                st.error(str(exc))
            else:
                st.download_button(
                    "Descargar planilla de notas (.xlsx)",
                    data=workbook,
                    file_name="planilla_de_notas.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

    st.caption(
        "Privacidad: los nombres y RUT se procesan en memoria y no se guardan en disco "
        "ni se envían a servicios externos."
    )


main()
