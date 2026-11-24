from __future__ import annotations

from decimal import Decimal

import pytest

from gradebook.grading import (
    PRESETS,
    Evaluation,
    GradingScheme,
    accumulative_average,
    final_from_scheme,
    round_grade,
    weight_fraction,
    weighted_final,
)
from gradebook.models import InvalidSchemeError


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("3.55", "3.6"),
        ("3.89", "3.9"),
        ("5.05", "5.1"),
        ("3.95", "4.0"),
        ("6.95", "7.0"),
        ("6.94", "6.9"),
        ("1.0", "1.0"),
        ("7.0", "7.0"),
        (3.45, "3.5"),
        (Decimal("4.45"), "4.5"),
        ("1.15", "1.2"),  # el float 1.15 se redondearía a 1.1
    ],
)
def test_round_grade_half_up(value: object, expected: str) -> None:
    assert str(round_grade(value)) == expected


def test_round_grade_never_bankers() -> None:
    # Python round() usa banca: round(2.25, 1) == 2.2. Aquí siempre sube.
    assert round_grade("2.25") == Decimal("2.3")
    assert round_grade("2.45") == Decimal("2.5")


def test_weight_fraction() -> None:
    assert weight_fraction("35") == Decimal("0.35")
    assert weight_fraction(Decimal("20.5")) == Decimal("0.205")


@pytest.mark.parametrize("preset", PRESETS.values())
def test_presets_are_valid_and_sum_100(preset: GradingScheme) -> None:
    preset.validate()
    assert preset.total_weight == Decimal("100")


def test_presets_expected_shapes() -> None:
    two = PRESETS["2_evaluaciones"]
    assert [e.weight for e in two.evaluations] == [Decimal("35"), Decimal("35")]
    assert two.controls_weight == Decimal("30")
    assert two.component_names == ("Evaluación 1", "Evaluación 2", "Nota Acumulativa")

    three = PRESETS["3_evaluaciones"]
    assert [e.weight for e in three.evaluations] == [
        Decimal("20"),
        Decimal("25"),
        Decimal("25"),
    ]
    assert three.controls_weight == Decimal("30")


def test_decimal_sum_avoids_float_error() -> None:
    scheme = GradingScheme(
        (Evaluation("Uno", "33.3"), Evaluation("Dos", "33.3"), Evaluation("Tres", "33.4")),
        controls_weight="0",
        n_controls=0,
    )
    scheme.validate()  # con floats la suma daría 100.00000000000001

    assert scheme.total_weight == Decimal("100")


def test_rejects_when_weights_do_not_sum_100() -> None:
    scheme = GradingScheme(
        (Evaluation("Uno", "33.33"), Evaluation("Dos", "33.33"), Evaluation("Tres", "33.33")),
        controls_weight="0",
        n_controls=0,
    )
    with pytest.raises(InvalidSchemeError, match="100"):
        scheme.validate()


def test_rejects_weights_over_100_and_negatives() -> None:
    with pytest.raises(InvalidSchemeError):
        GradingScheme((Evaluation("Uno", "120"),), "0", 0).validate()
    with pytest.raises(InvalidSchemeError):
        GradingScheme((Evaluation("Uno", "50"), Evaluation("Dos", "60")), "-10", 0).validate()


def test_rejects_empty_or_invalid_structure() -> None:
    with pytest.raises(InvalidSchemeError):
        GradingScheme((), "0", 0).validate()
    with pytest.raises(InvalidSchemeError):
        GradingScheme((Evaluation("Uno", "50"), Evaluation("uno", "50")), "0", 0).validate()
    with pytest.raises(InvalidSchemeError):
        GradingScheme((Evaluation("RUT", "100"),), "0", 0).validate()
    with pytest.raises(InvalidSchemeError):
        GradingScheme((Evaluation("Uno", "70"),), "30", 0).validate()


def test_controls_with_zero_weight_hide_columns() -> None:
    scheme = GradingScheme((Evaluation("Uno", "50"), Evaluation("Dos", "50")), "0", 6)
    scheme.validate()
    assert not scheme.has_controls
    assert scheme.control_names == ()
    assert scheme.component_names == ("Uno", "Dos")


def test_accumulative_average_ignores_blanks() -> None:
    assert accumulative_average(["3.55", "4.0", None]) == Decimal("3.8")
    assert accumulative_average([None, None]) is None
    assert accumulative_average([]) is None
    assert accumulative_average(["6.95"]) == Decimal("7.0")


def test_missing_component_counts_as_min_grade_only_in_final() -> None:
    scheme = PRESETS["3_evaluaciones"]
    # 5.5*0.2 + 1.0*0.25 (faltante) + 6.0*0.25 + 4.5*0.3
    final = final_from_scheme(
        scheme,
        evaluations=["5.5", None, "6.0"],
        controls=["4.0", "4.0", "5.0", "5.0", "4.5", "4.0"],
    )
    expected = round_grade(
        Decimal("5.5") * Decimal("0.2")
        + Decimal("1.0") * Decimal("0.25")
        + Decimal("6.0") * Decimal("0.25")
        + Decimal("4.5") * Decimal("0.3")
    )
    assert final == expected


def test_accumulative_blank_counts_as_min_grade_in_final() -> None:
    scheme = PRESETS["3_evaluaciones"]
    final = final_from_scheme(scheme, evaluations=["4.0", "4.0", "4.0"], controls=[])
    assert final == round_grade(Decimal("4.0") * Decimal("0.7") + Decimal("1.0") * Decimal("0.3"))


def test_all_components_blank_gives_none() -> None:
    scheme = PRESETS["3_evaluaciones"]
    assert final_from_scheme(scheme, evaluations=[None, None, None], controls=[None] * 6) is None
    assert weighted_final([None, None, None, None], scheme) is None


def test_weighted_final_rounds_only_at_the_end() -> None:
    # 3.45 * 0.5 + 3.45 * 0.5 = 3.45 → 3.5 (no se redondea cada término a 3.4...)
    scheme = GradingScheme((Evaluation("Uno", "50"), Evaluation("Dos", "50")), "0", 0)
    assert weighted_final(["3.45", "3.45"], scheme) == Decimal("3.5")
