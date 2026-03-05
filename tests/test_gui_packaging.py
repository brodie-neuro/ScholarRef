from __future__ import annotations

import logging
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import scholarref_gui
import scholarref_runtime


def test_gui_smoke_test_mode() -> None:
    rc = scholarref_gui.main(["--smoke-test"])
    assert rc == 0
    assert scholarref_runtime.resource_path("logo", "logo.png").exists()
    assert scholarref_runtime.resource_path("logo", "logo-removebg-preview (1).png").exists()


def test_configure_logging_falls_back_when_primary_log_is_blocked(monkeypatch, tmp_path: Path) -> None:
    root = logging.getLogger()
    original_handlers = list(root.handlers)
    for handler in original_handlers:
        root.removeHandler(handler)

    real_handler = scholarref_runtime.RotatingFileHandler
    calls = {"count": 0}

    def fake_handler(*args, **kwargs):
        calls["count"] += 1
        if calls["count"] == 1:
            raise PermissionError("blocked")
        return real_handler(*args, **kwargs)

    monkeypatch.setattr(scholarref_runtime, "RotatingFileHandler", fake_handler)
    monkeypatch.setattr(scholarref_runtime, "log_file_path", lambda: tmp_path / "blocked" / "scholarref.log")
    monkeypatch.setattr(scholarref_runtime.tempfile, "gettempdir", lambda: str(tmp_path / "fallback"))

    path = scholarref_runtime.configure_logging()
    assert str(path).startswith(str(tmp_path / "fallback"))
    assert path.exists()

    for handler in list(root.handlers):
        if getattr(handler, "_scholarref_handler", False):
            root.removeHandler(handler)
            handler.close()
    for handler in original_handlers:
        root.addHandler(handler)
