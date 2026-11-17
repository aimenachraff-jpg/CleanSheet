# Changelog

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
