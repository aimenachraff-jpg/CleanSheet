# CleanSheet — Project Update

**Product:** CleanSheet — offline-first Excel/CSV cleaner (rule-based, no AI)
**Repo:** https://github.com/aimenachraff-jpg/CleanSheet (local: C:\Users\stream\Desktop\EXLcleaner)
**Current version:** 1.1.0 (pyproject.toml)
**Current phase:** Phase A critical data-integrity fixes COMPLETE; moving to Phase B reliability/testing

---

## Baseline findings (established, verified)

- Repository is a git working tree, branch `main`, working tree CLEAN (nothing uncommitted).
- Branch `main` is 10 commits ahead of `origin/master` and 1 behind; remote is `aimenachraff-jpg/CleanSheet`.
- Python 3.10.11; openpyxl 3.1.5; tkinter available. Only runtime dependency: openpyxl.
- Three modules: `clean_sheet.py` (~605 lines CLI cleaner), `clean_sheet_locales.py` (6 locale packs), `clean_sheet_gui.py` (~193 lines Tkinter GUI).
- Tests: 57 tests across `test_clean_sheet.py` (17), `test_locale_packs.py` (15), `test_table_detection.py` (9), `test_multilevel.py` (11), `test_gui_logic.py` (5). **All 57 pass.** (CI runs the same 5 on ubuntu-latest / Python 3.12.)
- Docs: README.md, README.ar.md, README.fr.md, CONTRIBUTING.md, CHANGELOG.md, one issue template (`file_broke_it.yml`), GitHub Actions `tests.yml`.
- Packaging: pip-installable (`cleansheet`, `cleansheet-gui` entry points), PyInstaller `build_exe.bat` + `CleanSheet.spec`, unsigned Windows exe on Releases.
- Sample files: `sample_messy_arabic.xlsx`, `sample_messy_moroccan.xlsx`, `sample_messy_sales.xlsx` (+ cleaned twins).
- `private/`, `dist/`, `build/`, `cleaned_output/`, `_release/`, `_shots/` are gitignored and empty.

## Completed tasks
- Repository inspection (structure, source, tests, docs, packaging, CI, git history)
- Baseline test execution: all 57 tests pass (green baseline)
- Phase 1 technical/security audit complete — documented in AUDIT.md (3 HIGH, 5 MEDIUM, 3 LOW findings)
- IMPLEMENTATION_PLAN.md written (all 5 phases, task-level)
- Phase A critical data-integrity fixes COMPLETE:
  - A1 decode_text() chain: utf-8-sig → utf-8 → utf-16(BOM) → windows-1256 → iso-8859-6 → cp1252 (fixes H1 Arabic CSV mojibake)
  - A2 CSV Sniffer try/except → csv.excel fallback (fixes M2 sniff crash)
  - A3 GUI locale selector + locale passed through run_job → cs.run (fixes H2 silent +213 phone corruption)
  - A4 background-thread GUI cleaning + indeterminate progress + status label (fixes H3 GUI freeze)
  - A5 refuse to overwrite existing *_cleaned.xlsx unless --force (fixes M4 silent overwrite)

## In progress
- Phase B reliability & testing expansion

## Next steps
- Phase B: expanded test suite (encoding edge cases already in test_csv_encoding.py; add ambiguous duplicates, dates/decimals/currencies/phones/missing values, malformed/large files, original-untouched, reproducibility)
- Phase C GUI improvements (modern look, spreadsheet preview, review-surfacing, drag-and-drop optional)
- Phase D features (batch, reusable config, dedup keys, dry-run, audit log)
- Phase E docs/distribution (PRIVACY.md, SECURITY.md, ROADMAP.md, README updates, landing-page spec)

## Known limitations (pre-existing, documented in README)
- Reads the first table on the chosen sheet only
- No scans/PDFs, no formulas/dashboards, no merging files with different layouts
- Does not invent values; unreadable stays unreadable (orange)
- Tested on synthetic files

## Architectural decisions to preserve
- Rule-based only (no AI, no telemetry, no cloud, no accounts) — core differentiator
- Input file never modified; every change logged; ambiguous values flagged, never guessed
- openpyxl as the only dependency; standard library otherwise
- Deterministic behavior as default

## Important audit findings so far (see AUDIT.md for full detail)
- HIGH: Arabic CSVs encoded Windows-1256/ISO-8859-6 mojibake (cp1252 fallback only)
- HIGH: GUI always uses `dz` locale → non-Algerian phone numbers silently get wrong country code
- HIGH: GUI runs cleaning synchronously on the Tkinter thread → freezes, no progress
- MEDIUM: output workbook drops original metadata/styles/.xlsm macros with no warning
- MEDIUM: CSV sniffer can raise; no UTF-16/ISO-8859-6; `merge_locale` mutates global MONTHS
