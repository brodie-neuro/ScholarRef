#!/usr/bin/env python3
"""Runtime helpers for packaged and source-based ScholarRef runs."""

from __future__ import annotations

import importlib.metadata
import logging
import os
import platform
import re
import sys
import tempfile
from logging.handlers import RotatingFileHandler
from pathlib import Path


APP_NAME = "ScholarRef"


def _version_from_package_metadata() -> str | None:
    try:
        return importlib.metadata.version("scholarref")
    except importlib.metadata.PackageNotFoundError:
        return None


def _version_from_pyproject() -> str | None:
    pyproject = Path(__file__).resolve().with_name("pyproject.toml")
    if not pyproject.exists():
        return None

    try:
        text = pyproject.read_text(encoding="utf-8")
    except OSError:
        return None

    match = re.search(r'(?m)^version\s*=\s*"([^"]+)"\s*$', text)
    if match:
        return match.group(1)
    return None


def resolve_app_version() -> str:
    return _version_from_package_metadata() or _version_from_pyproject() or "1.0.0"


APP_VERSION = resolve_app_version()


def bundle_root() -> Path:
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return Path(base)
    return Path(__file__).resolve().parent


def resource_path(*parts: str) -> Path:
    return bundle_root().joinpath(*parts)


def app_data_dir() -> Path:
    if os.name == "nt":
        root = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    else:
        root = Path(os.environ.get("XDG_STATE_HOME", Path.home() / ".local" / "state"))
    path = root / APP_NAME
    try:
        path.mkdir(parents=True, exist_ok=True)
        return path
    except OSError:
        fallback = Path(tempfile.gettempdir()) / APP_NAME
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback


def log_dir() -> Path:
    path = app_data_dir() / "logs"
    path.mkdir(parents=True, exist_ok=True)
    return path


def log_file_path() -> Path:
    return log_dir() / "scholarref.log"


def _fallback_log_file_path() -> Path:
    path = Path(tempfile.gettempdir()) / APP_NAME / "logs" / "scholarref.log"
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def configure_logging() -> Path:
    path = log_file_path()
    root = logging.getLogger()
    if any(getattr(h, "_scholarref_handler", False) for h in root.handlers):
        return path

    root.setLevel(logging.INFO)
    try:
        handler = RotatingFileHandler(
            path,
            maxBytes=1_000_000,
            backupCount=3,
            encoding="utf-8",
        )
    except OSError:
        path = _fallback_log_file_path()
        handler = RotatingFileHandler(
            path,
            maxBytes=1_000_000,
            backupCount=3,
            encoding="utf-8",
        )
    handler._scholarref_handler = True  # type: ignore[attr-defined]
    handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")
    )
    root.addHandler(handler)
    logging.captureWarnings(True)
    logging.getLogger(__name__).info("Logging started")
    return path


def debug_info() -> str:
    return "\n".join(
        [
            f"App: {APP_NAME}",
            f"Version: {APP_VERSION}",
            f"Python: {sys.version.split()[0]}",
            f"Platform: {platform.platform()}",
            f"Executable: {sys.executable}",
            f"Bundle root: {bundle_root()}",
            f"Log file: {log_file_path()}",
        ]
    )
