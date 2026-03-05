# Contributing to ScholarRef

Thanks for your interest in improving ScholarRef.

## Before opening a PR

1. Open an issue first (`Bug Report` or `Suggestion`) so scope is agreed.
2. Keep changes focused and small.
3. Add/update tests or reproducible validation notes where relevant.

## Development checks

Run:

```bash
python -m py_compile scholarref.py reference_converter.py verify_reference_integrity.py scholarref_gui.py
```

If you changed conversion logic, run a quick conversion smoke test:

```bash
python scholarref.py --mode a2v --input "sample.docx" --output "sample_out.docx"
```

## Pull request guidance

- Describe the exact behavior change.
- Include before/after examples for citation/reference transformations.
- Note any edge cases and known limitations.

## Licensing note

This project uses the `ANM-1.0` license (attribution required, no modification redistribution without permission).
By submitting a contribution, you agree it may be incorporated into ScholarRef under that license.
