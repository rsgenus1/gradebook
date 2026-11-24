# Organizador de Planilla de Notas

Genera un archivo Excel de notas a partir de la lista de un curso
(nombre + RUT). El `.xlsx` trae tres hojas:

- **`Grades`**: planilla para que el profesor ingrese las notas.
- **`Stats`**: un gráfico de barras y uno de torta por **cada evaluación
  regular**.
- **`Final`**: tabla resumen con todas las estadísticas y los gráficos del
  `Promedio Final`.

Todo se recalcula solo.

La app **no calcula notas**: escribe **fórmulas vivas de Excel**, así que el
archivo se actualiza mientras el profesor escribe. También normaliza los RUT,
avisa de posibles errores (sin bloquear) y arma la estructura ponderada del
curso.

## Instalación

Requiere Python 3.11+ y, opcionalmente, LibreOffice para los tests de
recálculo.

```bash
uv sync                          # motor + dependencias
uv sync --extra ui --extra dev   # + Streamlit, pytest, ruff, openpyxl
```

El motor (`gradebook`) se instala y funciona **sin Streamlit**.

## Uso

### UI web

```bash
uv run streamlit run ui/streamlit_app.py
```

1. Sube la planilla (`.xlsx`, `.xls` o `.csv`).
2. Elige las columnas de nombre y RUT; revisa la vista previa normalizada y
   los avisos.
3. Elige la ponderación (2 evaluaciones, 3 evaluaciones o personalizada).
4. Descarga el `.xlsx` generado.

### CLI

```bash
uv run python -m gradebook entrada.xlsx --preset 3_evaluaciones -o salida.xlsx

# ponderación personalizada sin controles
uv run python -m gradebook entrada.csv \
  --evaluations "Proyecto:40,Examen final:60" --controls-weight 0 -o salida.xlsx

# columnas explícitas y 8 controles
uv run python -m gradebook entrada.xlsx --name-col "Nombre Completo" \
  --rut-col "RUT" --controls 8 -o salida.xlsx
```

### Formato de entrada

Cualquier `.xlsx` o `.xls` (primera hoja), o un `.csv` (separador `,` `;`
tabulación o `|`) con una columna de nombre y una de RUT. Todo se lee como **texto**, para no
perder ceros a la izquierda ni la `K`. Las columnas se autodetectan por
encabezado (`nombre`, `alumno`, `rut`, `run`, `documento`, …) o, si no hay
pista, por el contenido.

## Cálculo

| Regla | Comportamiento |
|---|---|
| Escala | notas de 1.0 a 7.0, un decimal |
| Redondeo | half-up (`3.55 → 3.6`, `3.95 → 4.0`), nunca bancario |
| `Nota Acumulativa` | promedio de los controles **ingresados**, redondeado a 1 decimal |
| `Promedio Final` | suma ponderada, redondeada al final |
| Componente faltante | cuenta como **1.0**, solo dentro de `Promedio Final` |
| Todo en blanco | `Promedio Final` queda vacío (inasistente total), no 1.0 |
| Aprobación | nota `≥ 4.0` |

Los pesos se escriben en la fila superior de `Grades` y las fórmulas los
referencian; si se editan y dejan de sumar 100 %, la celda de control muestra
la suma en rojo.

`Stats` muestra, por cada evaluación regular, su distribución por intervalos
(`[1,2[`, …, `[6,7]`) y su aprobación (aprobados, reprobados bajo 3.0 y entre
3.0 y 3.9). Cada evaluación ocupa su propia página al imprimir.

`Final` concentra la tabla con máximo, mínimo, asistentes, reprobados,
% reprobados, mediana y desviación estándar **muestral** (`STDEV.S`) de cada
evaluación, de la acumulativa y del promedio final, más los dos gráficos de
`Promedio Final`.

Ambas hojas usan referencias estructuradas (`tblGrades[...]`), de modo que
incluyen filas nuevas si se agregan estudiantes dentro de la tabla.

## Arquitectura

```
src/gradebook/          motor Python puro — nunca importa Streamlit
  models.py             Student, StudentTable, Warning, ColumnGuess, errores
  rut.py                normalización y módulo 11
  grading.py            GradingScheme, presets, Decimal y redondeo half-up
  io.py                 lectura de .xlsx / .csv en memoria
  workbook.py           hojas Grades, Stats y Final con xlsxwriter
  cli.py / __main__.py  CLI mínimo
ui/streamlit_app.py     adaptador delgado (solo widgets y llamadas al motor)
tests/                  muestras sintéticas y tests
```

La UI no valida pesos, no normaliza RUT ni arma fórmulas: solo llama a la API
del motor. Los avisos se devuelven como datos (`Warning`), no como excepciones.

## Tests

```bash
uv run pytest
uv run ruff check .
```

- `test_rut.py`, `test_grading.py`: normalización, redondeo half-up, suma 100 %
  con `Decimal`, componentes faltantes.
- `test_io.py`: lectura de `.xlsx`/`.csv`, texto, autodetección y avisos.
- `test_workbook.py`: hojas, encabezados, RUT como texto, validación de datos,
  pesos referenciados, formatos, gráficos por evaluación y saltos de página.
- `test_engine_independence.py`: el motor funciona con `streamlit` bloqueado.
- `test_e2e_recalc.py`: llena notas, recalcula con LibreOffice y compara el
  promedio final y **todas** las estadísticas contra Python
  (`Decimal`, `statistics`). Se salta si LibreOffice no puede arrancar.

### Muestras y revisión visual

```bash
uv run python tests/make_samples.py          # tests/samples/*.xlsx
soffice --headless --convert-to pdf --outdir tests/samples/pdf \
  tests/samples/curso_3eval.xlsx
pdftoppm -png -r 110 tests/samples/pdf/curso_3eval.pdf tests/samples/pdf/page
```

`tests/samples/` está en `.gitignore`. Las muestras se llenan modificando el
XML de la hoja (sin pasar por openpyxl), para que los gráficos conserven el
formato exacto que genera xlsxwriter.

## Decisiones por defecto

1. **Un solo archivo** `.xlsx` con dos hojas (`Grades`, `Stats`); no dos
   archivos con referencias externas.
2. **Redondeo intermedio**: `Nota Acumulativa` se redondea a 1 decimal antes de
   ponderar; el `Promedio Final` se redondea solo al final.
3. **Faltantes**: 1.0 únicamente dentro del `Promedio Final`; si no hay ningún
   componente, el final queda vacío; los controles en blanco se ignoran al
   promediar la acumulativa.
4. **Desviación estándar muestral** (`STDEV.S`) con `IFERROR` si hay menos de 2
   datos.
5. **Gráficos por evaluación regular**: barras + torta para cada una; al
   final, resumen y gráficos del `Promedio Final` en la hoja `Final`. Se
   mantiene en el mismo archivo (no en un workbook aparte) porque un archivo
   separado necesitaría referencias externas a `Grades` o duplicar la planilla,
   y las referencias se rompen al mover o renombrar archivos.
6. **Aprobación desde 4.0**.
7. **Intervalos** `[1,2[`, `[2,3[`, …, `[6,7]` con `COUNTIFS`.
8. **Impresión**: `Grades` y `Final` caben en una página (horizontal, A4);
   `Stats` usa una página por evaluación con saltos manuales. Se ajusta solo el
   ancho de `Stats` para que la leyenda de la torta no caiga en otra página.
9. **RUT sin dígito verificador** no se marca inválido: solo se avisa cuando el
   dígito existe y no corresponde.
10. **Errores bloqueantes** solo para entrada ilegible o ponderaciones que no
    suman 100 %; el resto son avisos.

## Limitaciones conocidas

- Lee solo la primera hoja de un libro de Excel.
- El `.xls` heredado se lee con calamine; los `.xls` con contraseña o
  corruptos se rechazan con un error claro.
- `xlsxwriter` no calcula fórmulas: los visores que no recalculan ven celdas
  vacías hasta abrir el archivo en Excel o LibreOffice (ambos recalculan).
- Los gráficos usan etiquetas de datos estándar de Excel; si se edita el
  archivo con openpyxl, la librería reserializa los gráficos y LibreOffice
  puede mostrar etiquetas más largas.
