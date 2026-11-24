# Organizador de Planilla de Notas

Genera un archivo Excel de notas a partir de una lista de estudiantes
(nombre + RUT). El `.xlsx` incluye dos hojas:

- **`Grades`**: planilla para que el profesor ingrese las notas.
- **`Stats`**: estadísticas y gráficos que se recalculan solos.

Las notas **no se calculan en la app**: el archivo trae fórmulas vivas de
Excel, así que se actualiza mientras el profesor escribe notas.

## Instalación

```bash
uv sync                      # motor + tests
uv sync --extra ui --extra dev  # además Streamlit, pytest y ruff
```

## Uso

UI web:

```bash
uv run streamlit run ui/streamlit_app.py
```

CLI:

```bash
uv run python -m gradebook entrada.xlsx --preset 3_evaluaciones -o salida.xlsx
```

## Estructura

```
src/gradebook/   motor en Python puro (sin Streamlit)
ui/              adaptador Streamlit
tests/           datos sintéticos y tests
TASK.md          plan de trabajo y estado
```

Reglas de negocio, formatos y decisiones de diseño: ver `AGENTS.md` y
`TASK.md`.
