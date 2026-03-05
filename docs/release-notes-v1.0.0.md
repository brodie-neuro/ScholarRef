# ScholarRef v1.0.0

First Windows `.exe` release of ScholarRef.

## Downloads

- `ScholarRef-setup-1.0.0.exe`
- `ScholarRef-windows-x64.zip`
- `SHA256SUMS.txt`
- `release-manifest.json`

## Supported targets

- Windows 10 x64
- Windows 11 x64

## Important

- This is an unsigned freeware release.
- Windows may show `Unknown publisher`.
- Download only from the official GitHub Releases page.
- Verify the hashes in `SHA256SUMS.txt` before sharing or mirroring the files.

## Included in this release

- Desktop Windows installer and portable packaged app
- GUI logging and `Copy Debug Info`
- Automated Windows packaging workflow
- Hybrid reference normalization for messy mixed-format bibliographies
- Reference-boundary fixes to avoid eating appendices or trailing sections
- Duplicate-reference collapsing for exact repeated entries
- Case-insensitive reference-header detection
- Automatic inference of untitled reference lists when the bibliography block is structurally obvious

## Known limitations

- The Windows build is unsigned.
- SmartScreen or antivirus products may warn on some machines.
- Clean `.docx` input is still preferred where possible, even though hybrid fallback handling is much stronger now.

## Bug reports

If something fails, include:

- ScholarRef version
- Windows version
- whether SmartScreen or antivirus blocked the file
- `Copy Debug Info` output
- relevant lines from `%LOCALAPPDATA%\ScholarRef\logs\scholarref.log`
