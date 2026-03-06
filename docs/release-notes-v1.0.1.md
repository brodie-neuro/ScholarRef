# ScholarRef v1.0.1

Windows patch release for the public `.exe` build.

## Downloads

- `ScholarRef-setup-1.0.1.exe`
- `ScholarRef-windows-x64.zip`
- `SHA256SUMS.txt`
- `release-manifest.json`

## Fixes

- Removed the accidental public packaging dependency on a private local `convert_to_plosone.py` file.
- GitHub package builds now install and test cleanly from the published repository alone.
- `verify_reference_integrity.py` now defaults to `references-only`, which is the supported public verification mode.

## Notes

- This is an unsigned freeware release.
- Download only from GitHub Releases.
- Verify the hashes in `SHA256SUMS.txt` before sharing or mirroring the files.

## Included in this release

- Windows 10 x64 support
- Windows 11 x64 support
- GUI `.exe` app
- Standard installer
