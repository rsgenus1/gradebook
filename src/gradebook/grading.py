"""Ponderaciones, presets y redondeo de notas.

Regla clave: los pesos se manejan como ``Decimal`` en puntos porcentuales
(35, 20.5, …). Comparar con ``Decimal`` evita los errores de coma flotante de
``sum(...) == 100``.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

from .constants import (
    ACCUMULATIVE_PREFIX,
    COL_ACCUMULATIVE_GRADE,
    COL_FINAL,
    COL_NAME,
    COL_RUT,
    DEFAULT_N_CONTROLS,
    EVALUATION_PREFIX,
    GRADE_STEP,
    MAX_N_CONTROLS,
    MIN_GRADE,
)
from .models import InvalidSchemeError

Number = int | float | str | Decimal

HUNDRED = Decimal(100)


def to_decimal(value: Number) -> Decimal:
    """Convierte a ``Decimal`` pasando por ``str`` (nunca por el float binario)."""
    if isinstance(value, Decimal):
        return value
    if isinstance(value, bool):
        raise TypeError("Los booleanos no representan notas ni pesos")
    return Decimal(str(value))


def round_grade(value: Number) -> Decimal:
    """Redondea a 1 decimal con half-up (3.55 → 3.6), no bancario."""
    return to_decimal(value).quantize(GRADE_STEP, rounding=ROUND_HALF_UP)


def weight_fraction(weight: Number) -> Decimal:
    """Peso en puntos porcentuales → fracción que entiende Excel (35 → 0.35)."""
    return to_decimal(weight) / HUNDRED


def _normalize_name(name: str) -> str:
    return " ".join(name.split()).casefold()


@dataclass(frozen=True, slots=True)
class Evaluation:
    """Una evaluación regular con su peso en puntos porcentuales."""

    name: str
    weight: Decimal

    def __post_init__(self) -> None:
        object.__setattr__(self, "name", str(self.name))
        object.__setattr__(self, "weight", to_decimal(self.weight))


@dataclass(frozen=True, slots=True)
class GradingScheme:
    """Configuración inmutable de ponderaciones de un curso.

    ``evaluations`` son las notas regulares; ``controls_weight`` es el peso
    agregado de los controles (la ``Nota Acumulativa``). La suma total debe ser
    exactamente 100.
    """

    evaluations: tuple[Evaluation, ...]
    controls_weight: Decimal = Decimal(0)
    n_controls: int = DEFAULT_N_CONTROLS

    def __post_init__(self) -> None:
        object.__setattr__(self, "evaluations", tuple(self.evaluations))
        object.__setattr__(self, "controls_weight", to_decimal(self.controls_weight))
        object.__setattr__(self, "n_controls", int(self.n_controls))

    # --- Derivados --------------------------------------------------------

    @property
    def total_weight(self) -> Decimal:
        return sum((ev.weight for ev in self.evaluations), Decimal(0)) + self.controls_weight

    @property
    def has_controls(self) -> bool:
        """Con peso 0 % no se generan columnas de controles ni acumulativa."""
        return self.controls_weight > 0

    @property
    def control_names(self) -> tuple[str, ...]:
        if not self.has_controls:
            return ()
        return tuple(f"{ACCUMULATIVE_PREFIX} {i}" for i in range(1, self.n_controls + 1))

    @property
    def component_names(self) -> tuple[str, ...]:
        """Componentes del promedio final, en el orden en que se ponderan."""
        names = [ev.name for ev in self.evaluations]
        if self.has_controls:
            names.append(COL_ACCUMULATIVE_GRADE)
        return tuple(names)

    @property
    def component_weights(self) -> tuple[Decimal, ...]:
        weights = [ev.weight for ev in self.evaluations]
        if self.has_controls:
            weights.append(self.controls_weight)
        return tuple(weights)

    # --- Validación -------------------------------------------------------

    def validate(self) -> None:
        """Lanza ``InvalidSchemeError`` si la configuración no es utilizable."""
        problems: list[str] = []

        if not self.evaluations:
            problems.append("debe haber al menos una evaluación")

        control_names = {_normalize_name(n) for n in self.control_names}
        reserved = {
            _normalize_name(COL_NAME),
            _normalize_name(COL_RUT),
            _normalize_name(COL_ACCUMULATIVE_GRADE),
            _normalize_name(COL_FINAL),
        } | control_names

        seen: dict[str, str] = {}
        for evaluation in self.evaluations:
            name = " ".join(evaluation.name.split())
            if not name:
                problems.append("las evaluaciones deben tener nombre")
                continue
            if len(name) > 200:
                problems.append(f"el nombre {name[:20]!r}… es demasiado largo")
            key = _normalize_name(name)
            if key in seen:
                problems.append(f"el nombre {name!r} está repetido")
            if key in reserved:
                problems.append(f"el nombre {name!r} choca con una columna de la planilla")
            seen[key] = name
            if evaluation.weight < 0:
                problems.append(f"el peso de {name!r} no puede ser negativo")

        if self.controls_weight < 0:
            problems.append("el peso de la acumulativa no puede ser negativo")
        if self.controls_weight > HUNDRED:
            problems.append("el peso de la acumulativa no puede superar 100 %")
        if self.has_controls and self.n_controls < 1:
            problems.append("debe haber al menos un control si la acumulativa pesa")
        if self.has_controls and self.n_controls > MAX_N_CONTROLS:
            problems.append(f"el máximo es {MAX_N_CONTROLS} controles")
        if self.total_weight != HUNDRED:
            problems.append(
                f"los pesos suman {self.total_weight.normalize()} % y deben sumar 100 %"
            )

        if problems:
            raise InvalidSchemeError("; ".join(problems))


def _evaluations(*weights: str) -> tuple[Evaluation, ...]:
    return tuple(
        Evaluation(f"{EVALUATION_PREFIX} {index}", Decimal(weight))
        for index, weight in enumerate(weights, start=1)
    )


PRESETS: dict[str, GradingScheme] = {
    "2_evaluaciones": GradingScheme(_evaluations("35", "35"), Decimal("30")),
    "3_evaluaciones": GradingScheme(_evaluations("20", "25", "25"), Decimal("30")),
}

PRESET_LABELS: dict[str, str] = {
    "2_evaluaciones": "2 evaluaciones (35 % + 35 % + acumulativa 30 %)",
    "3_evaluaciones": "3 evaluaciones (20 % + 25 % + 25 % + acumulativa 30 %)",
}

for _scheme in PRESETS.values():
    _scheme.validate()


def accumulative_average(controls: Sequence[Number | None]) -> Decimal | None:
    """Promedio de los controles ingresados, redondeado a 1 decimal.

    Los controles en blanco se ignoran. Si no hay ninguno, devuelve ``None``
    (y el promedio final lo trata como componente faltante).
    """
    present = [to_decimal(value) for value in controls if value is not None]
    if not present:
        return None
    return round_grade(sum(present, Decimal(0)) / len(present))


def weighted_final(
    components: Sequence[Number | None],
    scheme: GradingScheme,
) -> Decimal | None:
    """Promedio ponderado final, redondeado a 1 decimal.

    Un componente en blanco cuenta como ``MIN_GRADE`` solo aquí. Si *todos*
    están en blanco, el resultado es ``None`` (inasistente total), no 1.0.
    """
    weights = scheme.component_weights
    if len(components) != len(weights):
        raise ValueError(
            f"Se esperaban {len(weights)} componentes y llegaron {len(components)}"
        )
    if all(value is None for value in components):
        return None
    total = Decimal(0)
    for value, weight in zip(components, weights, strict=True):
        grade = MIN_GRADE if value is None else to_decimal(value)
        total += grade * weight_fraction(weight)
    return round_grade(total)


def final_from_scheme(
    scheme: GradingScheme,
    *,
    evaluations: Sequence[Number | None],
    controls: Sequence[Number | None] = (),
) -> Decimal | None:
    """Atajo útil para tests y para el CLI: calcula el final en Python puro."""
    accumulative = accumulative_average(controls) if scheme.has_controls else None
    components = [*evaluations]
    if scheme.has_controls:
        components.append(accumulative)
    return weighted_final(components, scheme)
