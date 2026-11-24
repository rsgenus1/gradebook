from __future__ import annotations

from pathlib import Path

import pytest

pytest.importorskip("streamlit")


def test_app_renders_without_uploaded_file() -> None:
    from streamlit.testing.v1 import AppTest

    app_path = Path(__file__).resolve().parents[1] / "ui" / "streamlit_app.py"
    app = AppTest.from_file(str(app_path), default_timeout=60)
    app.run()
    assert list(app.exception) == []
    assert any("xlsx" in message.value for message in app.info)
