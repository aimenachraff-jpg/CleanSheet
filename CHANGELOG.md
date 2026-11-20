# Changelog

## 1.2.0 - reliability, GUI, and product features

### Critical data-integrity fixes
- CSV files now decode through a chain (UTF-8-BOM, UTF-8, UTF-16 BOM, Windows-1256, ISO-8859-6, then Windows-1252) so Arabic CSVs exported from Windows are read correctly instead of being garbled (mojibake).
- CSV dialect sniffing failures fall back to comma-separated instead of crashing on tiny or unusual files.
- The GUI now has a country/locale selector and passes it to the cleaner, so phone numbers use the correct country code (+213 / +212 / +216 / +20 / +33) instead of always defaulting to Algeria (+213).
- GUI cleaning runs on a background thread with a progress indicator; the window stays responsive on large files.
- The tool refuses to overwrite an existing `*_cleaned.xlsx` unless you pass `--force`.

### Reliability & testing (88 tests total)
- New tests for CSV encoding edge cases, GUI locale/threading, original-file-untouched (sha256), output reproducibility, config round-trip, malformed files, and a 5,000-row large file.
- All tests run in CI on every push.

### Product features
- Batch processing: clean multiple files in one CLI invocation (per-file continue-on-error).
- `--dry-run`: report what would change without writing any output file.
- `--export-log`: write a `*_audit.csv` change log next to the cleaned output.
- `dedup_keys` config: customize which columns define a duplicate (default: whole row).
- Reusable cleaning configurations round-trip through `--write-config` / `--config`.

### GUI improvements
- Spreadsheet preview of the first rows before you clean.
- Modern, professional interface with a consistent accent color.
- Optional drag-and-drop via `tkinterdnd2` (never a hard dependency).
- Clearer error messages in the CLI (no raw tracebacks).

### Documentation
- New `PRIVACY.md` (offline, local-only; no telemetry, cloud, or accounts), `SECURITY.md`, `ROADMAP.md`, and a landing-page specification in `docs/landing-page-spec.md`.
- README updated in English, Arabic, and French.

### Under the hood
- Package version bumped to 1.2.0; install with `pip install cleansheet==1.2.0` or `pip install --upgrade cleansheet`.

## 1.1.0 - locale packs and pip install

- New `clean_sheet_locales.py`: country packs for Algeria (dz), Morocco (ma), Tunisia (tn), Egypt (eg), France (fr), and a `generic` pack that only checks number plausibility.
- New `--locale` flag on the command line. Phone prefixes, national length, currency words (MAD/درهم, TND/د.ت, EGP/جنيه...) and the leading-zero convention now come from the pack. A `--config` file still wins over the pack.
- Tunisian numbers (8 digits, no leading zero) no longer get a false "leading 0 missing" warning.
- Install from PyPI: `pip install cleansheet`, then `cleansheet file.xlsx` or `cleansheet-gui`.
- 15 new tests (57 total).

## 1.0.0 - first public release

- Rule-based cleaner for Excel/CSV: extra spaces, casing, exact duplicates, mixed price and date formats, number words, Arabic-Indic digits, two-level headers, total rows.
- Finds the table anywhere on the sheet; guesses column types and shows the guess to confirm.
- Output workbook with five sheets: Clean, Review, Change log, Before-After, Summary. Yellow = assumption to check, orange = unreadable or missing, original kept.
- Tkinter GUI (`clean_sheet_gui.py`, `start_CleanSheet.bat`) and standalone Windows exe on the Releases page.
- 42 tests.
