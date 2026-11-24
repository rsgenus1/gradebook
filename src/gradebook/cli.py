"""CLI mínimo: prueba viva de que el motor no depende de la UI.

    python -m gradebook entrada.xlsx --preset 3_evaluaciones -o salida.xlsx
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .grading import PRESETS, Evaluation, GradingScheme
from .io import read_students
from .models import GradebookError
from .workbook import build_workbook


def _parse_evaluations(value: str) -> tuple[Evaluation, ...]:
    """Convierte ``"Prueba 1:40,Prueba 2:60"`` en evaluaciones ponderadas."""
    evaluations: list[Evaluation] = []
    for chunk in value.split(","):
        name, separator, weight = chunk.rpartition(":")
        if not separator or not name.strip():
            raise argparse.ArgumentTypeError(
                f"Formato esperado 'Nombre:peso', recibí {chunk!r}"
            )
        evaluations.append(Evaluation(name.strip(), weight.strip()))
    if not evaluations:
        raise argparse.ArgumentTypeError("Debes indicar al menos una evaluación")
    return tuple(evaluations)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="gradebook",
        description="Genera un Excel de notas con hojas Grades y Stats.",
    )
    parser.add_argument("input", type=Path, help="planilla de entrada (.xlsx o .csv)")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("notas.xlsx"),
        help="archivo de salida (por defecto notas.xlsx)",
    )
    parser.add_argument(
        "--preset",
        choices=sorted(PRESETS),
        default="2_evaluaciones",
        help="configuración de ponderación predefinida",
    )
    parser.add_argument("--name-col", help="columna con el nombre (autodetectada si se omite)")
    parser.add_argument("--rut-col", help="columna con el RUT (autodetectada si se omite)")
    parser.add_argument(
        "--controls",
        type=int,
        default=None,
        help="cantidad de controles (por defecto 6, o el del preset)",
    )
    parser.add_argument(
        "--evaluations",
        type=_parse_evaluations,
        default=None,
        help="ponderación personalizada, p. ej. 'Prueba 1:40,Examen:60'",
    )
    parser.add_argument(
        "--controls-weight",
        default=None,
        help="peso de la acumulativa en la ponderación personalizada (p. ej. 30)",
    )
    return parser


def _scheme_from_args(args: argparse.Namespace) -> GradingScheme:
    if args.evaluations is None:
        scheme = PRESETS[args.preset]
        if args.controls is not None:
            scheme = GradingScheme(
                scheme.evaluations,
                scheme.controls_weight,
                n_controls=args.controls,
            )
        return scheme

    controls_weight = args.controls_weight if args.controls_weight is not None else "0"
    return GradingScheme(
        args.evaluations,
        controls_weight,
        n_controls=args.controls or 6,
    )


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        source = args.input.read_bytes()
        students = read_students(source, name_col=args.name_col, rut_col=args.rut_col)
        scheme = _scheme_from_args(args)
        scheme.validate()
        workbook_bytes = build_workbook(students, scheme)
    except GradebookError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except OSError as exc:
        print(f"Error leyendo {args.input}: {exc}", file=sys.stderr)
        return 1

    args.output.write_bytes(workbook_bytes)

    # Privacidad: nunca se imprimen nombres ni RUT, solo conteos.
    print(f"Listo: {len(students.students)} estudiantes → {args.output}")
    if students.warnings:
        summary = ", ".join(
            f"{count} × {code}" for code, count in sorted(students.warnings_by_code().items())
        )
        print(f"Avisos: {summary} (ver vista previa en la UI para el detalle).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
