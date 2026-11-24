"""El motor debe funcionar sin Streamlit instalado o importable."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

SRC = Path(__file__).resolve().parents[1] / "src"


def test_no_streamlit_imports_in_engine_source() -> None:
    import re

    pattern = re.compile(r"^\s*(?:import|from)\s+streamlit\b", re.MULTILINE)
    offenders = []
    for path in (SRC / "gradebook").rglob("*.py"):
        if pattern.search(path.read_text(encoding="utf-8")):
            offenders.append(str(path))
    assert offenders == []


def test_engine_works_with_streamlit_blocked() -> None:
    code = """
import builtins

_real_import = builtins.__import__


def _guard(name, *args, **kwargs):
    if name.split(".")[0] == "streamlit":
        raise ImportError("streamlit está bloqueado en este test")
    return _real_import(name, *args, **kwargs)


builtins.__import__ = _guard

import gradebook
from gradebook import PRESETS, Student, StudentTable, build_workbook

students = StudentTable((Student("Ana", "123456785"),))
data = build_workbook(students, PRESETS["2_evaluaciones"])
assert data.startswith(b"PK\\x03\\x04")
assert "streamlit" not in __import__("sys").modules
print("ok")
"""
    result = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        text=True,
        timeout=120,
        env={**os.environ, "PYTHONPATH": str(SRC)},
    )
    assert result.returncode == 0, result.stderr
    assert result.stdout.strip() == "ok"
