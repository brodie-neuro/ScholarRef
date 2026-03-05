# Windows Trust And Signing

## Current ScholarRef policy

ScholarRef is currently distributed as unsigned freeware.

That means:

- the installer and packaged `ScholarRef.exe` do not carry a commercial code-signing certificate
- Windows may show `Unknown publisher`
- SmartScreen or antivirus products may add extra friction on some machines

This is expected for the current release model.

## What users should trust instead

Because the build is unsigned, trust comes from release provenance and file integrity:

1. download only from the official GitHub Releases page
2. verify the SHA-256 hash from `SHA256SUMS.txt`
3. confirm the release manifest matches the published files
4. use `Copy Debug Info` and logs if something fails

ScholarRef already supports that workflow:

- GitHub Actions builds the Windows artifacts
- GitHub Releases publishes the installer, zip, checksums, and release manifest
- `SHA256SUMS.txt` and `release-manifest.json` let users verify the files
- the GUI writes logs to `%LOCALAPPDATA%\ScholarRef\logs\scholarref.log`
- the GUI exposes `Copy Debug Info` for issue reports

## What this does not solve

Checksums prove integrity.
They do not remove SmartScreen warnings.

Without code signing, Windows reputation remains weaker than for signed commercial software.

That is the tradeoff for a zero-cost freeware distribution path.

## Release-note policy for unsigned builds

Every public Windows release should state all of the following clearly:

- the build is unsigned
- the installer may show `Unknown publisher`
- the official download location is GitHub Releases
- the matching SHA-256 values are published in `SHA256SUMS.txt`
- supported targets are Windows 10 x64 and Windows 11 x64

## What to ask from bug reporters

If a user reports install or launch problems, ask for:

- app version
- Windows version
- whether SmartScreen or antivirus blocked the file
- `Copy Debug Info` output
- relevant lines from `%LOCALAPPDATA%\ScholarRef\logs\scholarref.log`

## Future option

If you ever decide to buy a code-signing certificate later, the build scripts already have optional signing hooks.
They are not required for the current unsigned freeware workflow.
