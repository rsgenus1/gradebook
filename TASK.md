# TASK — Organizador de Planilla de Notas

Plan de trabajo del proyecto. Cada casilla se marca al completarse, en el mismo
commit que la implementa.

## Contexto

App Streamlit que recibe una lista de estudiantes (nombre + RUT) y devuelve un
`.xlsx` con dos hojas:

- **`Grades`**: planilla que llena el profesor (fórmulas vivas de Excel).
- **`Stats`**: estadísticas y gráficos que se recalculan solos al llenar `Grades`.

El motor (`src/gradebook/`) es Python puro, sin Streamlit; la UI (`ui/`) es un
adaptador delgado. Ver `AGENTS.md` para el detalle de las reglas.

## Tareas

### 1. Andamiaje del proyecto
- [x] `pyproject.toml` con dependencias (polars, fastexcel, xlsxwriter) y extras `ui` / `dev`.
- [x] Estructura `src/gradebook/`, `ui/`, `tests/`.
- [x] `ruff` y `pytest` configurados; entorno instalado con `uv`.
- [x] `TASK.md` y `README.md` iniciales.

### 2. Modelos y errores del dominio
- [x] `models.py`: `Student`, `StudentTable`, `Warning`, `ColumnGuess`.
- [x] Excepciones propias: `GradebookError`, `InputFormatError`, `InvalidSchemeError`.
- [x] Constantes compartidas (notas, nombres de hojas, tabla).

### 3. Normalización de RUT
- [x] `normalize_rut()`: sin puntos, guion ni espacios, en mayúsculas.
- [x] Validación módulo 11 (solo como aviso, nunca bloqueo).
- [x] Detección de duplicados y de RUT vacíos.

### 4. Ponderaciones y redondeo
- [x] `GradingScheme` inmutable con evaluaciones + peso de la acumulativa.
- [x] `PRESETS` para 2 y 3 evaluaciones (35/35/30 y 20/25/25/30).
- [x] Validación de suma 100 % con `Decimal` (sin errores de flotante).
- [x] `round_grade()` half-up; promedio ponderado con `Decimal` de punta a punta.
- [x] Regla de componente faltante = 1.0 solo en el promedio final.

### 5. Lectura de la planilla de entrada
- [x] `read_students()` y `detect_columns()` sobre `.xlsx` / `.csv` en memoria.
- [x] Todas las columnas leídas como texto (RUT con ceros y `K` intactos).
- [x] Autodetección de columnas de nombre y RUT por encabezado.

### 6. Hoja `Grades`
- [x] Tabla `tblGrades` con filas alternadas y encabezados legibles.
- [x] Pesos en celdas encima de cada columna ponderada + verificación `OK 100%`.
- [x] Columnas calculadas `Nota Acumulativa` y `Promedio Final` con fórmulas.
- [x] Validación de datos 1–7 (permite vacío), formato `0.0`, rojo si `< 4.0`.
- [x] Filas alternadas, congelar paneles, anchos, impresión horizontal a 1 página.

### 7. Hoja `Stats`
- [x] Máximo, mínimo, asistentes, reprobados, % reprobados, mediana y desv. estándar.
- [x] Todo con fórmulas que leen `Grades`, con `IFERROR` donde corresponde.
- [x] Sin errores visibles con cursos vacíos o de un estudiante.

### 8. Gráficos
- [x] Tabla auxiliar con intervalos `[1,2[` … `[6,7]` y `COUNTIFS`.
- [x] Gráfico de barras por intervalos (sin leyenda redundante).
- [x] Gráfico de torta: aprobado / reprobado `< 3.0` / reprobado `3.0–3.9`.
- [x] Títulos, etiquetas, colores coherentes, sin cuadrículas recargadas.

### 9. API pública y CLI
- [x] `gradebook/__init__.py` exporta la API mínima.
- [x] CLI `python -m gradebook entrada.xlsx --preset 3_evaluaciones -o salida.xlsx`.

### 10. UI Streamlit
- [x] Subida, elección de columnas, vista previa normalizada y avisos.
- [x] Selección de preset / personalizada con suma en vivo y bloqueo de descarga.
- [x] Descarga del `.xlsx` en memoria; sin persistir datos personales.

### 11. Datos de ejemplo y tests
- [ ] `tests/make_samples.py` con curso sintético (~35), tildes/ñ, blancos y casos borde.
- [ ] `test_rut.py`, `test_grading.py`, `test_workbook.py`, `test_engine_independence.py`.
- [ ] Escenarios: 1 estudiante, evaluación sin asistentes, cada preset.

### 12. Verificación end-to-end
- [ ] Recalcular con LibreOffice y comparar `Promedio Final` y estadísticas.
- [ ] Comprobar ausencia de `#DIV/0!`, `#NUM!` y `#VALUE!`.

### 13. Revisión visual
- [ ] Exportar a PDF y revisar las imágenes por preset.
- [ ] Corregir problemas de legibilidad, superposición y colores.

### 14. Cierre
- [ ] README con uso, arquitectura y limitaciones.
- [ ] Listar al usuario las decisiones de la sección 10 de `AGENTS.md`.
