# Windows Release Checklist

## CI gate

- `python -m pytest -q` passes.
- `python -m py_compile scholarref.py scholarref_gui.py scholarref_runtime.py convert_to_plosone.py verify_reference_integrity.py` passes.
- `powershell -ExecutionPolicy Bypass -File .\scripts\build_windows_installer.ps1 -SkipTests` succeeds.
- `powershell -ExecutionPolicy Bypass -File .\scripts\test_windows_installer.ps1` succeeds.
- `dist\SHA256SUMS.txt` and `dist\release-manifest.json` are produced.

## Clean machine gate

- Test on a clean Windows 11 VM or Windows Sandbox.
- Install using `ScholarRef-setup-<version>.exe`, not just the unpacked app folder.
- Confirm the app opens with no import, path, or asset errors.
- Convert one APA -> Vancouver `.docx`.
- Convert one Vancouver -> APA `.docx`.
- Confirm the output files open in Microsoft Word.
- Confirm a log file exists under `%LOCALAPPDATA%\ScholarRef\logs\scholarref.log`.
- Use `Copy Debug Info` and verify the clipboard includes version, platform, executable, and log path.
- Uninstall and confirm the application is removed cleanly.

## Release assets

- Publish `ScholarRef-windows-x64.zip`.
- Publish `ScholarRef-setup-<version>.exe`.
- Publish `SHA256SUMS.txt`.
- Publish `release-manifest.json`.
- Include the exact version and supported Windows versions in the release notes.

## Unsigned release gate

- State clearly that the Windows build is unsigned.
- State clearly that Windows may show `Unknown publisher`.
- Tell users to download only from GitHub Releases.
- Tell users to verify `SHA256SUMS.txt`.
- Ask bug reporters for `Copy Debug Info`, logs, and any SmartScreen or antivirus warning text.

See [windows-trust-and-signing.md](windows-trust-and-signing.md) for the unsigned freeware trust model.
