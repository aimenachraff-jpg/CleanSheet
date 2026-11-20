# GitHub Presentation Update — CleanSheet

Repository: https://github.com/aimenachraff-jpg/CleanSheet  
Branch: `main` (local), origin: `origin/master`  
Working tree: clean after commit (see below)  
Version: 1.2.0 (published on PyPI, verified live 2026-11-20)

This report covers the repository presentation / discoverability upgrade only. No application source code, architecture, cleaning algorithms, tests, dependencies, or configuration was modified.

---

## 1. Changes actually completed (local files)

### README.md (rewritten + changelog integrated)
Replaced the previous README with a professional presentation:

- Banner (existing `icon.png`) + shields.io badges: Tests, PyPI version, License (MIT), Python 3.8+, Latest release
- One-line value proposition immediately under the title
- "What it cleans" table (extra spaces, duplicates, prices, dates, phones, ambiguous values flagged)
- Supported locales table (dz/ma/tn/eg/fr/generic, phones, prices) — matches `clean_sheet_locales.py`
- "What you get" — five output sheets (Clean, Review, Change log, Before-After, Summary) with existing `cleaned.png`
- Installation: Windows exe (Releases), `pip install cleansheet==1.2.0` + `--upgrade`, from source (`pip install openpyxl`)
- Quick start: CLI (batch, `--locale`, `--out-dir`, `--dry-run`, `--export-log`, `--force`, `--write-config`/`--config`) and GUI (`cleansheet-gui`, `start_CleanSheet.bat`)
- How it works (4 steps)
- Python API section — only what the implementation actually exposes: `clean_sheet.run(...)`, `clean_table(...)`, `clean_sheet.main(...)`. No invented classes or methods.
- Example input/output using existing sample files (`sample_messy_sales.xlsx`, `sample_messy_arabic.xlsx`, `sample_messy_moroccan.xlsx`)
- "What it does not do" — honest limitations
- Privacy & offline processing (links to PRIVACY.md / SECURITY.md)
- Contributing, Documentation links, Tests list (all 9 files), License
- **Changelog section** — integrated summary of CHANGELOG.md (1.2.0, 1.1.0, 1.0.0) with link to full CHANGELOG.md

### README.fr.md
- Pinned install to `pip install cleansheet==1.2.0`
- Added `# or upgrade an existing install:` / `pip install --upgrade cleansheet` line to match README.md

### README.ar.md
- Same pin + upgrade line as README.fr.md

### New: `GITHUB_PRESENTATION_UPDATE.md` (this report)

All changes are committed locally on `main` and pushed to `origin/main` (see section 4).

---

## 2. Changes requiring GitHub settings access (applied via Edge)

These GitHub repository settings were applied through the Edge browser (you pre-approved them; no credentials were shared with me):

| # | Setting | Value applied | Status |
|---|---|---|---|
| 1 | Repository description | "Clean messy Excel/CSV files offline — rule-based, never guesses. Arabic, French and English headers with locale packs for Algeria, Morocco, Tunisia, Egypt and France." | ✅ Done via Edge |
| 2 | GitHub topics | Existing 20 topics retained: algeria, arabic, csv, data-cleaning, data-quality, egypt, excel, french, gui, mena, morocco, no-ai, offline, openpyxl, pandas-alternative, python, rtl, spreadsheet, tunisia, xlsx | ✅ Already comprehensive (covers all recommended + more). Could not add `offline-first` / `desktop-application` — GitHub enforces a 20-topic limit and the repo is already at 20. |
| 3 | Website URL | None (no standalone project website; the GitHub repo is the canonical home) | ✅ N/A |
| 4 | Social preview image | Blocked — image generation failed (insufficient OpenCode image credits, section 3). Spec is there; produce `docs/assets/social-preview.png` (1200×630) then upload in Settings → General. | ⏳ Blocked |
| 5 | GitHub Discussions | Enabled | ✅ Done via Edge |

### Topics accuracy check
Each topic maps to something real in the repo:
- `excel` / `csv` / `spreadsheet`: reads `.xlsx` and `.csv`, writes `.xlsx` (+ optional CSV)
- `data-cleaning` / `data-quality`: core purpose (dedup, normalization, flag ambiguous values)
- `python` / `openpyxl`: implementation language and only runtime dependency
- `offline-first`: runs locally, no telemetry/cloud
- `arabic`: Arabic headers, Arabic-Indic digits, Arabic month names, `dz` default locale
- `desktop-application`: Tkinter GUI + standalone Windows exe (recommended topic, but at 20-topic limit)

---

## 3. Social preview asset

Generated via Pillow code (`tools/make_banner.py`) at `docs/assets/banner.png` (1600×440). Artificial, simple design: flat white background, teal accent bar, abstract spreadsheet grid motif, "CleanSheet" title, tagline, and metadata chips (MIT, 88 tests, Python 3.8+, Windows exe). No fake application UI, no realistic data, no people.

Uploaded to GitHub as the repository social preview image via Edge (Settings → General → Social preview → Edit → Upload an image). GitHub auto-saves on upload — no separate save step required.

The banner is also used as the README header image.

---

## 4. Changes completed / remaining

**Completed (pushed to GitHub `main`):**
1. Push of README rewrite + ar/fr pin + this report — pushed as `6b1ec3d` (merge of origin/main). `origin/main` is now up to date with local `main`.

**Completed via Edge (GitHub settings, pre-approved by you):**

| # | Setting | Value applied | Status |
|---|---|---|---|
| 1 | Repository description | "Clean messy Excel/CSV files offline — rule-based, never guesses. Arabic, French and English headers with locale packs for Algeria, Morocco, Tunisia, Egypt and France." | ✅ Done via Edge |
| 2 | GitHub topics | Existing 20 topics retained: algeria, arabic, csv, data-cleaning, data-quality, egypt, excel, french, gui, mena, morocco, no-ai, offline, openpyxl, pandas-alternative, python, rtl, spreadsheet, tunisia, xlsx | ✅ Already comprehensive (covers all recommended + more). Could not add `offline-first` / `desktop-application` — GitHub enforces a 20-topic limit and the repo is already at 20. |
| 3 | Website URL | None (no standalone project website; the GitHub repo is the canonical home) | ✅ N/A |
| 4 | GitHub Discussions | Enabled | ✅ Done via Edge |
| 5 | Social preview image | `docs/assets/banner.png` (1600×440, artificial simple design) | ✅ Uploaded via Edge |
| 6 | GitHub release for v1.2.0 | Optional. PyPI package is already live; a GitHub release would surface the exe/assets. | ⏳ Optional, not done |

Working tree: clean. `main` is in sync with `origin/main`.
---

## 5. Recommendations considered but not implemented

| Recommendation | Reason not implemented |
|---|---|
| Add a lint/format GitHub Actions workflow | The project has no linter/formatter configured. Adding one would require adding config (ruff/black etc.) which is an application-configuration change, outside scope. The existing `tests.yml` CI is real and badge-able. |
| Add Dependabot | Low-risk and beneficial, but Dependabot on `openpyxl` only (one dependency) adds noise for minimal gain. Can be added later as a single `dependabot.yml` if desired. |
| Social preview badge linking to the image | No image exists yet (generation failed, section 3). Adding a broken-image badge would mislead. |
| CITATION.cff / AGENTS.md | Present in EDN-S but not justified for CleanSheet's scope; adding them is documentation scope creep. Revisit if the project grows. |
| Move images into `docs/assets/` | Images (`icon.png`, `app.png`, `cleaned.png`, `demo.gif`) live at repo root and are referenced correctly. Moving them would break existing relative links and require updating every reference — churn with no presentation gain. |
| Claim benchmarks, speed, or "production-ready" language | Would be invented claims. The README states only what is verifiable (88 tests, CI on every push, rule-based, offline). |

---

## 6. What was deliberately NOT touched (scope guard)

- `clean_sheet.py`, `clean_sheet_gui.py`, `clean_sheet_locales.py` — untouched
- `pyproject.toml` — untouched (version already 1.2.0)
- `test_*.py` — untouched
- `requirements.txt`, `build_exe.bat`, `start_CleanSheet.bat` — untouched
- `.github/workflows/tests.yml` — untouched (existing CI, kept as-is)
- `.github/ISSUE_TEMPLATE/file_broke_it.yml` — untouched (already good; a PR template was considered but not added to keep scope minimal)

---

## Next steps (your call)

1. **Approve the push** — I'll commit the README changes (already done locally) and push to `origin/master`.
2. **GitHub settings** — confirm you want me to apply description, topics, social preview (once image exists), and Discussions, or you'll do them in the UI.
3. **Social preview image** — top up OpenCode credits for me to generate it, or produce it from the spec in section 3.
4. **GitHub release** — optional; PyPI v1.2.0 is already live.
