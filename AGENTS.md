# AGENTS.md

Guía para agentes de código que trabajen en este repositorio.

## 1. Propósito

App **Streamlit** que toma una lista de estudiantes (nombre + RUT) y genera un archivo Excel de notas, ya formateado y legible, con dos hojas:

- **`Grades`**: planilla principal donde el profesor ingresa las notas.
- **`Stats`**: estadísticas y gráficos que se actualizan solos a medida que se llenan notas en `Grades`.

La app **no calcula notas por su cuenta**: genera el archivo con fórmulas de Excel vivas. El profesor trabaja en Excel después de descargarlo.

## 2. Stack

- Python 3.11+
- `streamlit` (UI): **solo** en el adaptador de UI, nunca en el motor (ver sección 2)
- `polars` + `fastexcel` (lectura de la planilla de entrada)
- `xlsxwriter` (escritura del Excel: fórmulas, tablas, formato condicional, validación de datos, gráficos nativos)
- `pytest` y `ruff` para tests y estilo
- Gestión de dependencias en `pyproject.toml` (usar `uv` si está disponible)

Estructura sugerida:

```
src/gradebook/              # MOTOR: Python puro, sin imports de streamlit
  models.py                 # dataclasses: Student, GradingScheme, avisos, excepciones
  rut.py                    # normalización (y validación opcional) de RUT
  grading.py                # presets de ponderación, validación, redondeo
  io.py                     # lectura de la planilla de entrada (bytes / file-like)
  workbook.py               # construcción del Excel (Grades + Stats)
  __init__.py               # API pública del motor (ver abajo)
ui/
  streamlit_app.py          # ADAPTADOR: solo widgets y llamadas al motor
tests/
  make_samples.py           # genera planillas de ejemplo (entrada y salida)
  test_rut.py
  test_grading.py
  test_workbook.py
  test_engine_independence.py
  samples/                  # salidas generadas (ignorar en git)
```

### Separación motor / UI

El motor debe poder usarse desde cualquier interfaz (otra UI web, una API HTTP, un CLI, un script) sin tocar su código. Reglas:

- **Prohibido importar `streamlit` (o cualquier framework de UI/web) dentro de `src/gradebook/`.** Tampoco `st.session_state`, `st.cache_*`, ni nada que dependa del entorno de Streamlit.
- El motor recibe y devuelve **datos y bytes**, no objetos de UI:
  - entrada: `bytes` o file-like (no rutas, no `UploadedFile`);
  - salida: `bytes` del `.xlsx`.
- API pública mínima, exportada desde `gradebook/__init__.py`:

```python
read_students(source, *, name_col=None, rut_col=None) -> StudentTable   # incluye avisos
detect_columns(source) -> ColumnGuess
normalize_rut(raw: str) -> str
GradingScheme            # dataclass inmutable; .validate()
PRESETS                  # {"2_evaluaciones": ..., "3_evaluaciones": ...}
build_workbook(students: StudentTable, scheme: GradingScheme) -> bytes
```

- Los errores de dominio son excepciones propias (`InvalidSchemeError`, `InputFormatError`). Los avisos (RUT duplicado, dígito verificador inválido, etc.) se devuelven **como datos** (lista de objetos con código y mensaje) y la UI decide cómo mostrarlos.
- La UI no contiene reglas de negocio: no valida la suma 100%, no normaliza RUT, no arma fórmulas. Solo llama al motor y presenta el resultado.
- `streamlit` va en un extra opcional de `pyproject.toml` (`pip install .[ui]`). El motor debe instalarse y funcionar sin él.
- Opcional pero recomendado: un CLI mínimo (`python -m gradebook entrada.xlsx --preset 3_evaluaciones -o salida.xlsx`) como prueba viva de que el motor no depende de la UI.
- Un test (`test_engine_independence.py`) debe fallar si el motor importa `streamlit`: por ejemplo, bloquear `streamlit` en `sys.modules` y comprobar que `import gradebook` y `build_workbook` siguen funcionando.

## 3. Entrada

Planilla (`.xlsx` o `.csv`) con al menos nombre y RUT.

- Leer **todas las columnas como texto** (un RUT sin dígito verificador puede inferirse como número y perder ceros o la `K`).
- Dejar que el usuario elija qué columna es nombre y cuál es RUT (`selectbox`), autodetectando por encabezado (`nombre`, `rut`, `run`, `id`, etc.).
- Mostrar vista previa con el RUT ya normalizado antes de generar.

### Normalización del RUT

Resultado: `str` sin puntos, sin guion, sin espacios y en mayúsculas.

| Entrada          | Salida      |
|------------------|-------------|
| `12.345.678-5`   | `123456785` |
| `9.876.543-k`    | `9876543K`  |
| ` 12345678-5 `   | `123456785` |

- En el Excel la columna RUT se escribe como **texto** (formato `@`), nunca como número.
- Advertir (sin bloquear) en la UI si hay RUT vacíos, duplicados o con dígito verificador inválido (módulo 11). Es solo un aviso.

## 4. Ponderaciones

El usuario elige una configuración en la UI:

| Preset                   | Evaluaciones regulares | Controles (acumulativa) |
|--------------------------|------------------------|--------------------------|
| 2 evaluaciones (default) | 35% + 35%              | 30%                      |
| 3 evaluaciones           | 20% + 25% + 25%        | 30%                      |
| Personalizada            | N evaluaciones, nombre y peso libres | peso libre (puede ser 0%) |

Reglas:

- Los pesos de **toda** configuración deben sumar exactamente 100%. En personalizada, mostrar la suma en vivo y **deshabilitar la descarga** mientras no sea 100%.
- Comparar sin errores de flotante (usar `Decimal` o enteros en décimas de punto porcentual, no `sum(...) == 100` con floats).
- Cantidad de controles: por defecto 6 (`Acumulativa 1` … `Acumulativa 6`), editable por el usuario. Si el peso de la acumulativa es 0%, omitir esas columnas.
- Notas en escala **1 a 7**.

## 5. Redondeo (importante)

Las notas se aproximan a **1 decimal, half-up**: `3.55 → 3.6`, `3.89 → 3.9`, `5.05 → 5.1`, `3.95 → 4.0`. **No es redondeo bancario.**

- En las fórmulas de Excel usar `ROUND(x, 1)`, que ya se comporta como half-up para positivos.
- En Python **nunca** usar `round()` (es bancario y además sufre errores de flotante). Para tests o valores esperados usar:

```python
from decimal import Decimal, ROUND_HALF_UP

def round_grade(x) -> Decimal:
    return Decimal(str(x)).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
```

- Si se calcula un promedio ponderado en Python, hacerlo con `Decimal` de punta a punta, no con floats.

## 6. Workbook `Grades`

Tabla de Excel (`add_table`, nombre `tblGrades`) con una fila por estudiante.

Columnas, en orden:

1. `Nombre`
2. `RUT` (texto)
3. `Evaluación 1`, `Evaluación 2`, … (entrada)
4. `Acumulativa 1` … `Acumulativa N` (entrada)
5. `Nota Acumulativa` (**calculada**): promedio de los controles ingresados, redondeado a 1 decimal; vacío si no hay ninguno.
6. `Promedio Final` (**calculado**): suma ponderada, redondeada a 1 decimal.

Pesos:

- Escribirlos en celdas visibles **justo encima del encabezado** de cada columna ponderada (formato `%`). Las fórmulas deben **referenciar esas celdas**, no llevar los pesos escritos a mano.
- Incluir una celda de verificación: muestra `OK 100%` o `Suma = xx%` (en rojo, con formato condicional) si los pesos editados dejan de sumar 100%.

Reglas de cálculo (valores por defecto; ver sección 10):

- Celda de nota en blanco = inasistente / sin nota. **Nunca** escribir 0 ni 1.0 en las celdas de entrada ni en `Nota Acumulativa`: los blancos quedan blancos (así las estadísticas de cada columna cuentan bien a los asistentes).
- **Componente faltante = nota mínima (1.0)**, pero **solo dentro de la fórmula de `Promedio Final`**. Los componentes son cada evaluación regular y `Nota Acumulativa`.
- Si el estudiante no tiene **ningún** componente con nota (todo en blanco), `Promedio Final` queda vacío (`""`): es inasistente total, no un 1.0.
- Forma de la fórmula (ejemplo con 2 evaluaciones; la nota mínima sale de una constante `MIN_GRADE = 1.0` del código, no escrita a mano en varios lugares):

```
=IF(COUNT(E3,F3,AO3)=0,"",
   ROUND(E$2*IF(ISNUMBER(E3),E3,1) + F$2*IF(ISNUMBER(F3),F3,1) + AO$2*IF(ISNUMBER(AO3),AO3,1), 1))
```

- `Nota Acumulativa` = promedio de los controles **ingresados** (los controles en blanco se ignoran), redondeado a 1 decimal; si no hay ninguno, queda `""` y el final la trata como componente faltante (1.0).

Entrada de datos:

- Validación de datos en las celdas de notas: decimal entre 1 y 7, permitiendo vacío, con mensaje de error claro.
- Formato numérico `0.0` y alineación centrada.
- Formato condicional: notas `< 4.0` en rojo.
- Columnas calculadas con encabezado de otro color y texto "calculado" en el comentario o subtítulo, para que se distingan de las de entrada.

### Legibilidad

- **Filas alternadas por color** (zebra). Preferir el estilo de tabla con `banded rows` (se mantiene al filtrar y ordenar). Si se usa formato condicional, que sea `=MOD(ROW(),2)=0` aplicado a todo el rango de la tabla.
- Fila de encabezado y las dos primeras columnas (`Nombre`, `RUT`) congeladas.
- Anchos de columna ajustados al contenido, sin texto cortado ni `####`.
- Autofiltro activo, bordes finos, fuente única (Calibri 11).
- Configurar impresión: horizontal, ajustar a 1 página de ancho, repetir encabezado.

## 7. Workbook `Stats`

Todas las celdas son **fórmulas** que leen `Grades`. Preferir referencias estructuradas a la tabla (`tblGrades[Promedio Final]`) para que las estadísticas incluyan filas nuevas; si alguna herramienta de verificación no las soporta, usar rangos fijos con margen y documentarlo.

Columnas de estadísticas (una por cada una de estas): cada **Evaluación regular**, **Nota Acumulativa** y **Promedio Final**.

Filas:

| Estadística       | Definición |
|-------------------|------------|
| Máximo            | `MAX` |
| Mínimo            | `MIN` |
| Asistentes        | `COUNT` (celdas con número; en blanco = inasistente) |
| Reprobados        | cantidad con nota `< 4.0` |
| % Reprobados      | Reprobados / Asistentes |
| Mediana           | `MEDIAN` |
| Desviación estándar | ver sección 10 |

- Envolver con `IFERROR(..., "")` donde pueda haber división por cero o muestras de menos de 2 datos.
- `% Reprobados` con formato porcentaje (`0.0%`).
- En la columna `Promedio Final`, "asistente" es quien tiene al menos un componente con nota; quien tenga componentes faltantes entra con 1.0 en ellos (ver sección 6) y puede aparecer como reprobado.

### Gráficos (sobre `Promedio Final`)

Cada gráfico se alimenta de una pequeña tabla auxiliar con fórmulas, visible en la hoja.

1. **Barras** por intervalos: `[1,2[`, `[2,3[`, `[3,4[`, `[4,5[`, `[5,6[`, `[6,7]`.
   - Conteo con `COUNTIFS(rango, ">=lo", rango, "<hi")`; el último intervalo es cerrado: `"<=7"`.
2. **Torta** con tres categorías: `Aprobado (≥ 4.0)`, `Reprobado < 3.0`, `Reprobado 3.0 – 3.9`.

Estética: título y títulos de ejes, etiquetas de datos, sin leyenda redundante en el gráfico de barras, colores consistentes (verde = aprobado, tonos rojo/naranja = reprobado), sin cuadrícula recargada, sin superposición con las tablas. Ocultar líneas de cuadrícula de la hoja.

## 8. UI Streamlit

Vive en `ui/streamlit_app.py` y es un adaptador delgado: solo widgets y llamadas a la API del motor (sección 2). Flujo lineal:

1. Subir planilla.
2. Elegir columnas de nombre y RUT; ver vista previa normalizada y avisos.
3. Elegir configuración de ponderación (+ cantidad de controles). Personalizada con `st.data_editor`.
4. Botón de descarga del `.xlsx` (todo en memoria con `BytesIO`).

**Privacidad:** nombres y RUT son datos personales. No escribirlos en logs, no persistirlos en disco ni en `st.session_state` más allá de la sesión, y no enviarlos a servicios externos.

## 9. Gotchas de xlsxwriter

- Las fórmulas se escriben **siempre en inglés y con coma** (`=IF(A1>3,1,0)`), sin importar el idioma de Excel; Excel las muestra localizadas (`SI`, `PROMEDIO`, `;`).
- Para funciones nuevas (`STDEV.S`, etc.) usar `Workbook(..., {"use_future_functions": True})`.
- `xlsxwriter` no calcula: los visores que no recalculan (vistas previas, lectores de datos) muestran el valor cacheado. Pasar `value=""` en `write_formula` para las fórmulas que inicialmente deben verse vacías, en vez del `0` por defecto. Excel y LibreOffice recalculan al abrir.
- Un `""` devuelto por una fórmula es texto: `MAX`, `MIN`, `MEDIAN`, `COUNT` y `COUNTIF` con criterios numéricos lo ignoran, que es lo deseado.
- Las tablas no admiten celdas combinadas ni encabezados repetidos.

## 10. Decisiones por confirmar

Implementar con estos valores por defecto, dejarlos como constantes en `grading.py` / `workbook.py` y fáciles de cambiar. **Al terminar, listarlos al usuario.**

1. **"2 workbooks"**: por defecto, un solo archivo `.xlsx` con dos hojas (`Grades`, `Stats`). Dos archivos separados exigirían referencias externas (`[Grades.xlsx]Grades!…`) que se rompen al mover o renombrar archivos y piden confirmar vínculos al abrir.
2. **Redondeo intermedio**: `Nota Acumulativa` se redondea a 1 decimal antes de ponderar; el `Promedio Final` se redondea al final.
3. **Componentes faltantes**: ya definido por el usuario: se asume la nota mínima (1.0) en el `Promedio Final`. Quedan como valores por defecto dos detalles que no se discutieron: (a) con todos los componentes en blanco el final queda vacío y (b) los controles en blanco se ignoran al promediar la acumulativa, en vez de contar como 1.0.
4. **Desviación estándar**: muestral (`STDEV.S`), con `IFERROR` si hay menos de 2 datos.
5. **Gráficos**: solo sobre `Promedio Final`.
6. **Aprobación**: nota `≥ 4.0`.

## 11. Pruebas y revisión

### Datos de ejemplo (`tests/make_samples.py`)

Generar con datos sintéticos (nunca RUT reales; calcular el dígito verificador con módulo 11):

- Curso de ~35 estudiantes con nombres con tildes y `ñ`, y RUT en formatos mezclados (`12.345.678-5`, `12345678-5`, minúsculas, con espacios, con `K`).
- Notas aleatorias con semilla fija, con **blancos** (inasistentes) en distintas columnas.
- Casos borde: `3.55`, `3.95`, `5.05`, `6.95`, `1.0`, `7.0`, `3.9` vs `4.0`.
- Escenarios extra: curso de 1 estudiante, curso sin ningún asistente en una evaluación, y una por cada preset (2 eval, 3 eval, personalizada).
- Estudiantes con componentes faltantes: una evaluación en blanco (el final debe usar 1.0 en ese componente), sin ningún control (acumulativa en blanco, final con 1.0 en ese componente) y con todo en blanco (final vacío).

### Tests automatizados

- `test_rut.py`: tabla de normalización, `K` minúscula, espacios, entrada vacía.
- `test_grading.py`: redondeo half-up con los casos borde, suma de pesos 100% (incluido caso flotante como 33.3 + 33.3 + 33.4), rechazo cuando no suma 100%, y la regla de componente faltante = 1.0 (valor esperado calculado en Python con `Decimal`).
- `test_engine_independence.py`: el motor se importa y genera un workbook con `streamlit` bloqueado; además, ningún archivo de `src/gradebook/` contiene `import streamlit`.
- `test_workbook.py`: abrir la salida con `openpyxl` y comprobar hojas, encabezados, RUT como texto, validación de datos, pesos referenciados desde celdas y existencia de los dos gráficos.
- **Cálculo end-to-end**: llenar notas con valores conocidos, recalcular con LibreOffice y comparar `Promedio Final` y todas las estadísticas contra valores esperados calculados en Python (`Decimal` para notas, `statistics` para mediana y desviación estándar).

Recalcular con LibreOffice:

```bash
soffice --headless --convert-to xlsx --outdir tests/samples/recalc tests/samples/curso_3eval.xlsx
# luego leer con openpyxl.load_workbook(..., data_only=True)
```

(En Arch: `libreoffice-fresh` y `poppler`.)

### Revisión estética obligatoria

Antes de dar el trabajo por terminado:

```bash
soffice --headless --convert-to pdf --outdir tests/samples/pdf tests/samples/curso_3eval.xlsx
pdftoppm -png -r 110 tests/samples/pdf/curso_3eval.pdf tests/samples/pdf/page
```

**Abrir las imágenes y mirarlas** (no basta con que el test pase). Revisar para cada preset:

- [ ] Filas alternadas con buen contraste, sin cortar texto, sin `####`.
- [ ] Columnas de entrada distinguibles de las calculadas; notas `< 4.0` en rojo.
- [ ] Pesos visibles encima de las columnas y cuadre del 100%.
- [ ] `Stats` ordenada, sin elementos superpuestos, con porcentajes legibles.
- [ ] Gráficos con título, etiquetas, colores coherentes y los intervalos exactos pedidos.
- [ ] Casos borde: sin asistentes y curso de 1 estudiante no muestran errores (`#DIV/0!`, `#NUM!`).
- [ ] Impresión: cabe en 1 página de ancho.

Iterar hasta que todo se vea bien; si algo no se puede verificar visualmente, decirlo explícitamente.

## 12. Convenciones

- Cero lógica de negocio en la UI; todo vive en el motor, con funciones puras y type hints.
- `ruff check` y `pytest` deben pasar antes de cada commit.
- Constantes para nombres de hojas, encabezados y colores; sin textos "mágicos" repetidos.
- Cambios pequeños y commits descriptivos.

## 13. Definition of Done

- La app genera `Grades` + `Stats` para los tres tipos de configuración.
- RUT normalizado y almacenado como texto.
- Motor independiente de Streamlit (`import gradebook` funciona sin `streamlit` instalado) y UI como adaptador delgado.
- Redondeo half-up verificado con los casos borde.
- Componente faltante cuenta como 1.0 solo en `Promedio Final`, verificado contra valores esperados.
- Estadísticas y gráficos se actualizan al ingresar notas en Excel (verificado con recálculo en LibreOffice).
- Revisión visual completada con imágenes generadas.
- Tests y linter pasan.
- Decisiones de la sección 10 informadas al usuario.
