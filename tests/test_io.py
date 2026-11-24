from __future__ import annotations

import io
from pathlib import Path

import pytest
import xlsxwriter

from gradebook import InputFormatError, WarningCode, detect_columns, read_students


def make_xlsx(rows: list[list[object]], headers: list[str]) -> bytes:
    buffer = io.BytesIO()
    workbook = xlsxwriter.Workbook(buffer, {"in_memory": True})
    sheet = workbook.add_worksheet("Lista")
    sheet.write_row(0, 0, headers)
    for row_index, row in enumerate(rows, start=1):
        for column_index, value in enumerate(row):
            if isinstance(value, (int, float)):
                sheet.write_number(row_index, column_index, value)
            else:
                sheet.write_string(row_index, column_index, str(value))
    workbook.close()
    return buffer.getvalue()


def test_reads_xlsx_with_accents_and_mixed_ruts() -> None:
    data = make_xlsx(
        [
            ["Ana Ñúñez", "12.345.678-5", "8A"],
            ["José Peña", "12345678-k", "8A"],
            ["María Soto", " 9876543-3 ", "8A"],
        ],
        ["Nombre Completo", "RUT", "Curso"],
    )
    table = read_students(data)
    assert table.name_col == "Nombre Completo"
    assert table.rut_col == "RUT"
    assert [student.name for student in table.students] == [
        "Ana Ñúñez",
        "José Peña",
        "María Soto",
    ]
    assert [student.rut for student in table.students] == [
        "123456785",
        "12345678K",
        "98765433",
    ]


def test_xlsx_numbers_are_read_as_text() -> None:
    data = make_xlsx([["Ana", 123456785]], ["Nombre", "RUT"])
    table = read_students(data)
    assert table.students[0].rut == "123456785"


def test_reads_csv_with_bom_and_semicolon() -> None:
    raw = "\ufeffNombre;RUT\nAna;12.345.678-5\nBeto;9876543-3\n".encode("utf-8")
    table = read_students(raw)
    assert [student.rut for student in table.students] == ["123456785", "98765433"]
    assert table.name_col == "Nombre"


def test_reads_file_like() -> None:
    raw = b"Nombre,RUT\nAna,123456785\n"
    table = read_students(io.BytesIO(raw))
    assert len(table.students) == 1


def test_empty_rows_are_skipped() -> None:
    raw = b"Nombre,RUT\nAna,123456785\n,\nBeto,98765433\n"
    table = read_students(raw)
    assert len(table.students) == 2


def test_detects_columns_by_values_when_headers_are_odd() -> None:
    raw = b"Documento,Persona\n12345678-5,Ana\n9876543-3,Beto\n11111111-1,Carla\n"
    guess = detect_columns(raw)
    assert guess.rut_col == "Documento"
    assert guess.name_col == "Persona"


def test_warnings_for_empty_duplicate_and_invalid_rut() -> None:
    raw = (
        b"Nombre,RUT\n"
        b"Ana,12.345.678-5\n"
        b"Beto,12.345.678-5\n"
        b"Carla,\n"
        b"Diego,99.999.999-9\n"
        b",20202020-2\n"
    )
    table = read_students(raw)
    codes = {warning.code for warning in table.warnings}
    assert codes == {
        WarningCode.RUT_DUPLICATE,
        WarningCode.RUT_EMPTY,
        WarningCode.RUT_INVALID,
        WarningCode.NAME_EMPTY,
    }
    duplicate = next(
        warning for warning in table.warnings if warning.code == WarningCode.RUT_DUPLICATE
    )
    assert duplicate.row == 3  # segunda aparición (fila 2 = primera Ana)
    counts = table.warnings_by_code()
    assert counts[WarningCode.RUT_DUPLICATE] == 1


def test_raises_on_empty_input() -> None:
    with pytest.raises(InputFormatError):
        read_students(b"")


def test_reads_legacy_xls() -> None:
    fixture = Path(__file__).resolve().parent / "fixtures" / "entrada.xls"
    table = read_students(fixture.read_bytes())
    assert table.name_col == "Nombre"
    assert table.rut_col == "RUT"
    assert [student.name for student in table.students] == [
        "Ana Ñúñez",
        "Beto Soto",
        "Carla Díaz",
    ]
    assert [student.rut for student in table.students] == [
        "123456785",
        "98765433",
        "111111111",
    ]


def test_raises_on_unreadable_ole_file() -> None:
    with pytest.raises(InputFormatError, match="Excel"):
        read_students(b"\xd0\xcf\x11\xe0" + b"esto no es un libro")


def test_raises_when_columns_are_missing() -> None:
    with pytest.raises(InputFormatError, match="nombre"):
        read_students(b"Algo,Otro\na,b\n")


def test_raises_when_same_column_is_chosen() -> None:
    with pytest.raises(InputFormatError, match="misma"):
        read_students(b"Nombre,RUT\nAna,123456785\n", name_col="Nombre", rut_col="Nombre")


def test_rejects_paths() -> None:
    with pytest.raises(TypeError, match="ruta"):
        read_students("planilla.xlsx")  # type: ignore[arg-type]
