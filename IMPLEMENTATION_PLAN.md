# CleanSheet — Implementation Plan

> Status: proposed. Review with the human partner before executing.
> Execution method: native (I implement each task in this session, then one final review).

**Goal:** Take CleanSheet v1.1.0 to a reliable, polished, privacy-first spreadsheet cleaner, fixing the 3 HIGH data-integrity findings first, then expanding tests, GUI, features, and docs.

**Architecture:** Keep the existing 3-module structure (`clean_sheet.py` CLI/core, `clean_sheet_locales.py` packs, `clean_sheet_gui.py` Tkinter GUI). No restructuring. All changes are additive or targeted bug fixes inside those files plus new test files. GUI cleaning moves to a background thread with results marshaled back to Tkinter via `after()`/`queue.Queue`.

**Tech stack:** Python 3.8+, openpyxl (only dependency), standard library, Tkinter. No new runtime dependencies (repo rule). Tests use unittest only.

**Spec:** AUDIT.md (findings H1-H3, M1-M5, L1-L3), PROJECT_UPDATE.md, and the user's 5-phase objectives.

---

## Global constraints (apply to every task)

- Rule-based only. No AI, no telemetry, no cloud upload, no accounts, no paid service.
- Input file is never modified; every change logged; ambiguous values flagged (verify/fail/missing), never silently guessed.
- openpyxl is the only dependency; standard library for everything else.
- `requires-python >= 3.8`; keep code 3.8-compatible.
- Deterministic behavior as the default.
- Preserve existing behavior and test suite (57 tests must stay green).
- Never commit credentials/API keys; never send customer spreadsheets anywhere.
- Data integrity → Reliability → Usability → Distribution priority order.

## Review focus (failure modes each task's tests must pin)

1. Arabic CSV in Windows-1256 / ISO-8859-6 / UTF-16 decodes to correct text.
2. GUI phone parse uses the selected locale country code, never hard-coded `+213`.
3. GUI stays responsive and shows progress while cleaning a multi-thousand-row file.
4. Original input file bytes are unchanged after every run (hash before/after).
5. Existing `_cleaned.xlsx` is not silently overwritten (fail or suffix).
6. Ambiguous date `03/04/2026` → verify, not silently decided.
7. Ambiguous price `2,500` → verify.
8. Exact duplicate removed (first kept); reused ID flagged; different-ID similar row flagged.
9. Malformed/tiny CSV and unusual delimiter do not crash with a raw traceback.
10. Large file (~5000 rows) completes without memory blow-up and GUI reports progress.

---

## Phase A — Critical data-integrity bug fixes (do first)

### Task A1: Robust CSV decode chain (H1, M2)

**Files:**
- Modify: `clean_sheet.py` — add `decode_text(raw: bytes) -> str` and use it in `_read_grid`
- Test: create `test_csv_encoding.py`

**Interfaces:**
- `decode_text(raw: bytes) -> str` — tries utf-8-sig, utf-8, utf-16 (BOM), windows-1256, iso-8859-6, cp1252 last; raises a clear error only if none decode.
- `_read_grid` calls `decode_text(raw)` instead of the 2-codec try/except.

**Steps:**
- [ ] Write failing tests: Windows-1256 Arabic CSV, ISO-8859-6, UTF-16, UTF-8-BOM, cp1252 Latin, undecodable garbage → clear error.
- [ ] Run tests → FAIL (garbage/mojibake or UnicodeDecodeError).
- [ ] Implement `decode_text`; wire into `_read_grid`.
- [ ] Run tests → PASS; run full 57-test suite → green.

### Task A2: CSV sniffer fallback (M2)

**Files:**
- Modify: `clean_sheet.py` — `_read_grid` sniff block
- Test: `test_csv_encoding.py`

**Interfaces:**
- `csv.Sniffer().sniff(...)` wrapped in try/except → fall back to `csv.excel` (comma). Tiny/unusual CSV no longer raises.

**Steps:**
- [ ] Write failing test: 2-row CSV with a delimiter Sniffer chokes on → returns rows, no exception.
- [ ] Run → FAIL with exception; implement try/except fallback; run → PASS.

### Task A3: GUI locale selector + pass locale to cleaner (H2)

**Files:**
- Modify: `clean_sheet_gui.py` — `App.__init__` (locale combobox), `run_job` signature + body, `clean` call site
- Test: `test_gui_locale.py` (headless)

**Interfaces:**
- `run_job(path, columns, date_order="dmy", out_dir=None, also_csv=False, sheet=None, header_row=None, levels=None, skip_totals=True, locale="dz")` → passes `locale` into `cs.run(..., locale=locale)`.
- GUI shows a locale combobox (`clean_sheet_locales.DESCRIPTIONS` labels) defaulting to `dz`; chosen value flows to `run_job`.
- `App` stores `self.locale_var`.

**Steps:**
- [ ] Write headless test: `run_job(..., locale="fr")` → phone `0612345678` parsed as `+33612345678`; locale `"dz"` → `+213612345678`.
- [ ] Run → FAIL (locale ignored / always dz); implement combobox + `locale` param; run → PASS.
- [ ] Manually verify GUI opens with locale combobox (display permitting).

### Task A4: Background-thread GUI cleaning + progress (H3)

**Files:**
- Modify: `clean_sheet_gui.py` — `App.clean`, add worker thread + `after` polling + progress label, disable/enable buttons
- Test: `test_gui_thread.py` (headless: verify run_job invoked on thread and result delivered via queue/after)

**Interfaces:**
- `App.clean()` spawns `threading.Thread(target=self._run_cleaning)`; result placed on `queue.Queue`; `self.after(100, self._poll_result)` delivers it to the UI thread.
- A "Cleaning…" indeterminate progress label + Clean button disabled while running; re-enabled on completion/error.
- Errors surfaced via `messagebox` from the UI thread (not from the worker).

**Steps:**
- [ ] Write headless test that monkeypatches `run_job` to sleep and confirms `clean()` returns quickly (non-blocking) and eventually delivers `self.result`.
- [ ] Run → FAIL (blocks); implement thread + queue polling; run → PASS.
- [ ] Run full suite → green.

### Task A5: Never silently overwrite existing output (M4)

**Files:**
- Modify: `clean_sheet.py` — `run()` output-target logic
- Test: `test_clean_sheet.py` (add to Workflow) or new test

**Interfaces:**
- `run(..., out_dir, ...)` — if target `*_cleaned.xlsx` already exists, raise `SystemExit` with a clear message ("output already exists; choose another --out-dir") unless a new `--force` flag is passed.

**Steps:**
- [ ] Write failing test: pre-existing output file → `cs.run` raises SystemExit / refuses; with `--force` overwrites.
- [ ] Run → FAIL; implement; run → PASS.

---

## Phase B — Reliability & testing

### Task B1: Expand the test suite (Phase 2 requirements)

**Files:**
- Create: `test_clean_sheet_extended.py` (or extend existing files)
- No source change

**Interfaces:** Tests call public functions (`parse_price`, `parse_date`, `parse_phone`, `clean_table`, `cs.run`, `decode_text`).

**Steps:**
- [ ] Add tests for: Arabic/French/English inputs; encoding edge cases (UTF-8, UTF-8-BOM, Windows-1252, Windows-1256, ISO-8859-6, UTF-16); ambiguous duplicates; dates/decimal separators/currencies/phone numbers/missing values; malformed and large files; original-file-untouched (hash before/after); output reproducibility (same input → same output); config round-trip (`--write-config` then `--config` reproduces).
- [ ] Run extended suite → all green; combined with 57 baseline → green.
- [ ] Commit tests.

### Task B2: Realistic synthetic datasets

**Files:**
- Create: `samples/sample_messy_*.xlsx` expansions or new `samples/` directory with fictional Arabic/Moroccan/French/English files exercising phones, prices, dates, duplicates, missing values.
- No source change

**Steps:**
- [ ] Generate 3-5 fictional sample files; verify each cleans without crash and produces expected sheets.
- [ ] Commit samples.

---

## Phase C — GUI improvements (Phase 3)

### Task C1: Modern, professional interface

**Files:**
- Modify: `clean_sheet_gui.py` — ttk styles, padding, colors, layout polish (keep Tkinter architecture)

**Interfaces:** No API change; visual only. `App.__init__` restyled.

**Steps:**
- [ ] Apply a cleaner ttk style (consistent font, button padding, section spacing, hint color, result label styling).
- [ ] Run GUI smoke (display permitting) + full test suite → green.

### Task C2: Spreadsheet preview before cleaning

**Files:**
- Modify: `clean_sheet_gui.py` — after inspect_file, show a small read-only grid preview of first N rows/detected headers

**Interfaces:** `inspect_file` already returns headers/samples; add preview widget populated from samples + first rows.

**Steps:**
- [ ] Add a read-only preview area under the column-type table showing first ~5 rows of detected values.
- [ ] Run suite → green.

### Task C3: Review uncertain modifications before applying

**Files:**
- Modify: `clean_sheet_gui.py` — after `run_job`, before finishing, surface count of verify/fail/missing and offer "Open Review sheet" / re-run option

**Interfaces:** `run_job` result already includes `verify`, `unresolved`, `missing`. GUI result label already reports them; make the Review sheet opening more prominent.

**Steps:**
- [ ] Prominent "Open Review sheet" action when verify/unresolved/missing > 0; clearer result wording.
- [ ] Run suite → green.

### Task C4: Drag-and-drop file support

**Files:**
- Modify: `clean_sheet_gui.py` — optional `tkinterdnd2` drop target guarded by try/import; graceful fallback to file picker
- No new hard dependency (tkinterdnd2 not added to requirements.txt)

**Interfaces:** `App` registers a drop target if `tkinterdnd2` is importable; dropped path → `load_file`.

**Steps:**
- [ ] Attempt `import tkinterdnd2`; if available, wire drop; otherwise file picker only (no crash, no hard dep).
- [ ] Run suite → green (headless, tkinterdnd2 absent → fallback path tested implicitly).

### Task C5: Better error messages (M5)

**Files:**
- Modify: `clean_sheet.py` — `main()` wraps `run()` in try/except with friendly one-line error + exit 1; `clean_sheet_gui.py` — already shows messagebox, keep
- Test: CLI error path

**Steps:**
- [ ] Implement friendly CLI errors; run a bad-input scenario → clean message, exit 1, no traceback.

---

## Phase D — Product features (Phase 4, implement where justified)

### Task D1: Batch processing of multiple files

**Files:**
- Modify: `clean_sheet.py` — `main` accepts multiple inputs or a `--batch` file list; loop `run()` per file, aggregate summary

**Interfaces:** `main(argv)` handles multiple `input` positional args; returns per-file results; overall exit 0 if all ok.

**Steps:**
- [ ] Implement multi-input CLI; test two files in one invocation → two outputs, one summary.

### Task D2: Reusable cleaning configurations (fixes L2)

**Files:**
- Modify: `clean_sheet_gui.py` — `run_job` writes full effective config (columns, date_order, locale, sheet, header_row, levels, skip_totals, also_csv); add "Load config" button
- Modify: `clean_sheet.py` — `run` accepts config; `--config` round-trip already works

**Interfaces:** `run_job` writes `{columns, date_order, locale, sheet, header_row, levels, skip_totals, also_csv}` to `*_used_config.json`.

**Steps:**
- [ ] Persist full config; add GUI "Load saved config" to prefill column types + locale.
- [ ] Test config round-trip reproduces identical output.

### Task D3: Customizable duplicate detection

**Files:**
- Modify: `clean_sheet.py` — `clean_table` / `run` gains `--dedup-keys` (columns to consider) option; default = exact full-row

**Interfaces:** `clean_table(headers, rows, cfg, dedup_keys=None)`; cfg key `dedup_keys`.

**Steps:**
- [ ] Implement optional key-column dedup; test: rows differing in a non-key column are NOT removed when dedup keys configured.

### Task D4: Dry-run mode

**Files:**
- Modify: `clean_sheet.py` — `run(..., dry_run=False)` returns cells/log without writing output files

**Interfaces:** `run` returns the same dict but skips `write_workbook`/report when `dry_run=True`; CLI `--dry-run` prints summary only.

**Steps:**
- [ ] Implement `--dry-run`; test: no output files written, report dict still returned.

### Task D5: Improved reports / before/after / exportable audit logs

**Files:**
- Modify: `clean_sheet.py` — `report_md` enhancements + optional CSV/JSON audit log export (`--export-log`)

**Interfaces:** `run(..., export_log=False)` writes `*_audit.csv`/`*_audit.json` of the change log.

**Steps:**
- [ ] Enrich markdown report; add audit-log export; test file created and parseable.

---

## Phase E — Documentation & distribution (Phase 5)

### Task E1: README & docs polish

**Files:**
- Modify: `README.md`, `README.ar.md`, `README.fr.md`, `CONTRIBUTING.md`
- Add: `PRIVACY.md` (no data leaves machine, no temp files, offline), `SECURITY.md`, `ROADMAP.md`

**Steps:**
- [ ] Update README with CLI flags, locale table, limitations, install from source + exe.
- [ ] Add privacy/security/roadmap docs.
- [ ] Commit.

### Task E2: Packaging & release process notes

**Files:**
- Add: `docs/release.md` (or section in CONTRIBUTING) — how to cut a release, build the exe, PyPI upload, tagging; GitHub Actions test badge already present

**Steps:**
- [ ] Document release steps; verify `pip install .` and `python -m PyInstaller` build from clean checkout (do not ship signed exe — note SmartScreen).

### Task E3: Landing-page specification

**Files:**
- Add: `docs/landing-page-spec.md` — structure, copy, no pricing/telemetry/cloud/account

**Steps:**
- [ ] Write spec only (no implementation of a paid service).

---

## Definition of done (per phase)

- Phase A: H1/H2/H3/M4 fixed; full suite green; no regressions in 57 baseline tests.
- Phase B: extended suite green; samples verified; original-file-untouched and reproducibility proven by tests.
- Phase C: GUI restyled + preview + progress + locale + review-surfaced; headless GUI tests green.
- Phase D: batch, reusable config, dedup keys, dry-run, audit log each tested.
- Phase E: README/AR/FR/CONTRIBUTING updated; PRIVACY/SECURITY/ROADMAP/landing-spec present.

---

## Notes

- Do not push to GitHub without authorization.
- Commit after each task with a clean conventional-commit message (`fix:`, `feat:`, `docs:`, `test:`).
- Before declaring a task done, run the affected test file and the full 57-test baseline.
