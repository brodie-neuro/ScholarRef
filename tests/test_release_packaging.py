from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import scholarref_runtime


def _read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def test_runtime_version_matches_pyproject() -> None:
    pyproject = _read_text(ROOT / "pyproject.toml")
    assert f'version = "{scholarref_runtime.APP_VERSION}"' in pyproject
    assert "py-modules" in pyproject


def test_installer_is_per_user_and_smoke_testable() -> None:
    installer = _read_text(ROOT / "installer" / "ScholarRef.iss")
    smoke_test = _read_text(ROOT / "scripts" / "test_windows_installer.ps1")
    clean_venv = _read_text(ROOT / "scripts" / "test_clean_venv.ps1")

    assert "DefaultDirName={localappdata}\\Programs\\{#MyAppName}" in installer
    assert "PrivilegesRequired=lowest" in installer
    assert "SetupIconFile=..\\logo\\scholarref-mark.ico" in installer
    assert "VersionInfoProductName={#MyAppName}" in installer
    assert "VersionInfoDescription={#MyAppName} Installer" in installer
    assert "/VERYSILENT" in smoke_test
    assert "--smoke-test" in smoke_test
    assert '"-m", "venv"' in clean_venv


def test_workflow_builds_installer_and_release_metadata() -> None:
    workflow = _read_text(ROOT / ".github" / "workflows" / "windows-package.yml")
    trust_doc = _read_text(ROOT / "docs" / "windows-trust-and-signing.md")
    readme = _read_text(ROOT / "README.md")
    spec = _read_text(ROOT / "ScholarRef.spec")
    release_cfg = _read_text(ROOT / ".github" / "release.yml")
    issue_cfg = _read_text(ROOT / ".github" / "ISSUE_TEMPLATE" / "config.yml")

    assert "build_windows_installer.ps1" in workflow
    assert "test_windows_installer.ps1" in workflow
    assert "SHA256SUMS.txt" in workflow
    assert "release-manifest.json" in workflow
    assert "unsigned freeware" in workflow.lower()
    assert "GITHUB_STEP_SUMMARY" in workflow
    assert "unsigned freeware" in trust_doc.lower()
    assert "docs/windows-trust-and-signing.md" in readme
    assert "unknown publisher" in readme.lower()
    assert "scholarref-mark.ico" in spec
    assert "pyproject.toml" in spec
    assert "Windows Packaging" in release_cfg
    assert "Latest Windows Release" in issue_cfg
