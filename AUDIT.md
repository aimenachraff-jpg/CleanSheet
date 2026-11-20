# CleanSheet — Technical & Security Audit (Phase 1)

**Product:** CleanSheet v1.1.0 — offline-first Excel/CSV cleaner
**Scope:** `clean_sheet.py`, `clean_sheet_gui.py`, `clean_sheet_locales.py`, tests, packaging, docs
**Baseline:** all 57 tests pass; working tree clean; branch `main`
**Priority order applied:** Data integrity → Reliability → Usability → Distribution

---

## How to read this audit

Each finding lists: severity, the symptom, the evidence (code location + reproduction logic), affected files, and recommended fix. Severity reflects impact on the product's core promise: local, deterministic, never-guess spreadsheet cleaning for Arabic/French/English North-African data.

---

## Findings

### HIGH — H1: Arabic CSV files encoded Windows-1256 / ISO-8859-6 are garbled (mojibake)

**Symptom:** An Arabic CSV saved by Windows (common in the MENA region) is read with corrupted Arabic characters; cleaning then fails or corrupts customer data.

**Evidence:** `clean_sheet.py:_read_grid()` (line ~344-346):

```python
try: text = raw.decode("utf-8-sig")
except UnicodeDecodeError: text = raw.decode("cp1252")
```

Only two codecs are tried: UTF-8 (with BOM) and Windows-1252. Windows-1256 (Arabic) and ISO-8859-6 are NOT attempted, so an Arabic Windows-1256 file decodes as cp1252 and every Arabic byte maps to the wrong Latin/Arabic glyph. This directly breaks the product's main differentiator (Arabic data support) and can silently corrupt customer spreadsheets.

**Affected files:** `clean_sheet.py` (`_read_grid`)

**Recommended fix:** Replace the two-codec fallback with an ordered decode chain: `utf-8-sig` → `utf-8` → `utf-16` (detect BOM) → `windows-1256` → `iso-8859-6` → `windows-1252` (last resort), trying each and accepting the first that decodes without error AND produces plausible text. Do not add a new dependency (keep the no-new-deps rule); implement a small `decode_text(bytes)` helper in `clean_sheet.py` with a test per codec.

---

### HIGH — H2: GUI silently applies the wrong country code to phone numbers (always `dz` locale)

**Symptom:** A French/Moroccan/Tunisian/Egyptian user cleaning phones via the GUI gets Algerian `+213` prefixes with status "ok" (no flag). Silent data corruption for every non-Algerian GUI user.

**Evidence:** `clean_sheet_gui.py:run_job()` (line ~44) calls `cs.run(path, cfg_path, out, sheet, header_row, also_csv, None, levels, skip_totals)` — `run()` defaults to `locale="dz"` (clean_sheet.py:575). The GUI exposes no locale selector, so the effective locale is always Algeria regardless of the user's country. `parse_phone("0612345678", dz)` → `+213612345678` "ok"; a French `06` number becomes a wrong Algerian number without any verify flag.

**Affected files:** `clean_sheet_gui.py` (`run_job`, `App`)

**Recommended fix:** Add a locale picker to the GUI (matching `DESCRIPTIONS` in `clean_sheet_locales.py`) and pass the chosen locale through `run_job` → `cs.run(..., locale=...)`. Default to `dz` only when nothing is chosen (documented). This restores the CLI's `--locale` guarantee in the GUI.

---

### HIGH — H3: GUI freezes during cleaning; no progress indication (no threading)

**Symptom:** The Tkinter window stops responding while `cs.run()` executes on the main thread. On large files the user cannot tell whether the app is working or hung; no progress bar, no cancel, no "Cleaning…" state.

**Evidence:** `clean_sheet_gui.py:App.clean()` (line ~172) calls `run_job(...)` synchronously and blocks the Tkinter event loop until it returns. `run_job` → `cs.run()` reads the whole grid, parses every cell, and writes a new workbook — all on the UI thread.

**Affected files:** `clean_sheet_gui.py` (`App.clean`)

**Recommended fix:** Run `run_job` on a background thread (or `concurrent.futures.ThreadPoolExecutor`) and marshal results back to the Tkinter thread with `after()` / `queue`. Add a determinate/indeterminate progress label and a "Cleaning…" state; disable the Clean button while running; keep the CLI synchronous (unchanged). Do not block the main thread.

---

### MEDIUM — M1: Output workbook drops original workbook properties, styles, number formats, and .xlsm macros with no warning

**Symptom:** The cleaned `.xlsx` is a brand-new openpyxl workbook. Original document properties (title/author/created), cell styles, column widths, number formats, and VBA macros (`.xlsm`) are lost. openpyxl silently strips macros on save — the user is not told.

**Evidence:** `clean_sheet.py:write_workbook()` builds `openpyxl.Workbook()` from scratch (line ~503). `_read_grid` opens `.xlsm` but `load_workbook` defaults `keep_vba=False`, and nothing warns about macro loss. `list_sheets` accepts `.xlsm`.

**Affected files:** `clean_sheet.py` (`write_workbook`, `_read_grid`, `list_sheets`)

**Recommended fix:** (a) Refuse or clearly warn when the input is `.xlsm` (macros cannot be preserved by openpyxl) — do not silently drop them. (b) Copy a small set of safe workbook properties (title, creator, description) where available into the output Summary. (c) Document in README/Limitations that output is a clean workbook (values + cleaning annotations), not a formatting clone. This preserves the "transparent cleaning operations" differentiator.

---

### MEDIUM — M2: CSV dialect sniffing can crash; no UTF-16 support

**Symptom:** A small/unusual CSV (e.g. few rows, uncommon delimiter) makes `csv.Sniffer().sniff()` raise, propagating as a raw exception instead of a graceful error. UTF-16 encoded CSVs are not supported.

**Evidence:** `clean_sheet.py:_read_grid()` line ~347: `csv.Sniffer().sniff(text[:4096], delimiters=",;\t")` has no fallback if sniffing fails. Decode chain only covers utf-8-sig and cp1252 (see H1).

**Affected files:** `clean_sheet.py` (`_read_grid`)

**Recommended fix:** Wrap sniffing in try/except and fall back to `csv.excel` (comma). Add UTF-16 (BOM detection) to the decode chain (H1 fix covers this).

---

### MEDIUM — M3: `merge_locale()` mutates the global `MONTHS` table

**Symptom:** `merge_locale` updates the module-global `MONTHS` dict (clean_sheet.py:296). Month names accumulate across runs/locales in a long-running process (e.g. GUI used repeatedly). Additive-only so low risk, but impure global state makes behavior order-dependent and harder to test/reason about.

**Evidence:** `clean_sheet.py:merge_locale()` (line ~290-297) — `MONTHS.update({ar_norm(k): v ...})`.

**Affected files:** `clean_sheet.py` (`merge_locale`, `MONTHS`)

**Recommended fix:** Stop mutating the global. Build the effective month table per-call inside `resolve_config` / pass it through `cfg`, or keep a module-level frozen base and merge into a local copy used by `parse_date`. Low-risk refactor; add a test that two successive `merge_locale` calls with different locales don't leak month names.

---

### MEDIUM — M4: Existing `_cleaned.xlsx` output is silently overwritten; no unique naming

**Symptom:** Re-running a clean on the same input/out-dir replaces the previous `_cleaned.xlsx` and `_report.md` without warning. A user who forgets they already cleaned loses the earlier output.

**Evidence:** `clean_sheet.py:run()` (line ~582-583) only refuses `target == src`; nothing guards an existing output file.

**Affected files:** `clean_sheet.py` (`run`)

**Recommended fix:** If the target already exists, either (a) fail with a clear message asking the user to pick another out-dir, or (b) auto-suffix (`_cleaned_1.xlsx`). Option (a) is safer and matches "never silently discard." CLI flag `--force` could override. Add a test.

---

### MEDIUM — M5: CLI crashes with a raw traceback on unexpected errors

**Symptom:** An unexpected exception (e.g. unreadable file, permission error) dumps a Python traceback and a non-zero exit instead of a friendly, actionable message. Poor for nontechnical business users.

**Evidence:** `clean_sheet.py:main()` (line ~593-602) only expects `SystemExit`-style failures from `cs.run()`; other exceptions propagate.

**Affected files:** `clean_sheet.py` (`main`)

**Recommended fix:** Wrap `run()` in `main()` with try/except → print a one-line error + hint, exit code 1, no stack trace unless `--verbose`. Keep `SystemExit` messages as-is.

---

### LOW — L1: `detect_header_row` scans only the first 60 rows

**Symptom:** A table whose header starts below row 60 is not found ("No header row found" or wrong detection).

**Evidence:** `clean_sheet.py:detect_header_row()` (line ~357-369) — `scan = grid[:60]`.

**Affected files:** `clean_sheet.py` (`detect_header_row`)

**Recommended fix:** Increase the scan window (e.g. 200 rows) or scan until the first data-like row. Minor; tables in real business files rarely start that deep. Add test.

---

### LOW — L2: GUI-written `_used_config.json` is incomplete (drops locale and other effective settings)

**Symptom:** The config the GUI saves next to the output contains only `columns` and `date_order`. Reusing it loses the locale choice and other settings, so re-runs behave differently.

**Evidence:** `clean_sheet_gui.py:run_job()` (line ~42-43) writes `{"columns": columns, "date_order": date_order}`.

**Affected files:** `clean_sheet_gui.py` (`run_job`)

**Recommended fix:** Persist the full effective config (columns, date_order, locale, also_csv, sheet, header_row, levels, skip_totals) so a saved config reproduces the run. Ties into Phase 4 "reusable cleaning configurations."

---

### LOW — L3: ID column value type is inconsistent (int vs str) depending on input cell type

**Symptom:** An ID cell `5` read as int stays int; read as string stays string. Duplicate/REUSE detection compares `repr()` so values still group, but downstream consumers of the Clean sheet get inconsistent ID types.

**Evidence:** `clean_sheet.py:clean_table()` id branch (line ~448-451) — the second list comprehension condition is convoluted and type-dependent.

**Affected files:** `clean_sheet.py` (`clean_table`)

**Recommended fix:** Normalize IDs to strings consistently (or ints when all digits) at parse time, before logging/dedup. Cosmetic; low risk.

---

## Severity summary

| Severity | Count | Findings |
|----------|-------|----------|
| HIGH | 3 | H1 Arabic CSV encoding, H2 GUI locale phone corruption, H3 GUI freeze/no progress |
| MEDIUM | 5 | M1 workbook properties/macros lost, M2 CSV sniff/UTF-16, M3 global MONTHS mutation, M4 silent output overwrite, M5 CLI raw tracebacks |
| LOW | 3 | L1 60-row header scan, L2 incomplete GUI config, L3 ID type inconsistency |

## Immediate risk to data integrity

H1 and H2 can silently alter customer data (garbled Arabic text; wrong country phone code). H2 is especially dangerous because it returns status "ok" with no flag, so it passes the "never silently guess" principle only on paper. These two must be fixed before any polish work.

## Recommended order of fixes

1. H1 — Arabic/Windows-1256/UTF-16 CSV decoding (data integrity)
2. H2 — GUI locale selector + pass locale to `cs.run` (data integrity)
3. H3 — background-thread cleaning + progress in GUI (reliability/usability)
4. M4 — don't silently overwrite existing output (data integrity)
5. M1 — .xlsm warning / properties carryover (transparency)
6. M2 — CSV sniffer fallback + UTF-16 (reliability)
7. M3, M5, L1, L2, L3 — hardening polish

## What is NOT a finding (intentional design, keep)

- Input file never modified; output written to a new file (confirmed by test + code)
- Ambiguous values flagged verify/fail/missing, never silently guessed (core principle)
- openpyxl as the only dependency (CONTRIBUTING rule #2)
- Exact duplicate removal with first-kept-wins, logged
- Formulas read as cached values; formulas without cached results warned (not a bug, documented limitation)

## Files to create next

- `PROJECT_UPDATE.md` (created) — living status record
- Prioritized implementation plan (markdown)
- Expanded test suite (encoding edge cases, ambiguous duplicates, malformed/large files, original-untouched, reproducibility)
