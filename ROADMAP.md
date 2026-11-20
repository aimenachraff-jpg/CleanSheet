# Roadmap — CleanSheet

CleanSheet is a rule-based, offline-first Excel/CSV cleaner. The roadmap keeps the
product focused on its differentiators: local processing, privacy, Arabic/French/English
(North African) formats, transparent cleaning operations, and human review of uncertain
values. It does **not** plan a paid service, telemetry, cloud upload, or accounts.

## Completed (v1.1.0 → current)

- Locale packs for Algeria (dz), Morocco (ma), Tunisia (tn), Egypt (eg), France (fr),
  plus a `generic` pack; `--locale` flag on the CLI.
- Pip-installable package (`cleansheet`, `cleansheet-gui` entry points) on PyPI.
- Tkinter GUI with column-type preview and confirmation before cleaning.
- Output workbook: Summary, Clean, Review, Change log, Before-After, (optionally)
  Skipped rows. Yellow = assumption to verify, orange = unreadable/missing, original
  kept.
- 88 tests covering cleaning rules, table/header detection, GUI logic, locale packs,
  CSV encoding edge cases, GUI locale/threading, and reliability.

## In progress / next

- Documentation package: PRIVACY.md, SECURITY.md, ROADMAP.md (this file), README
  refresh in English / Arabic / French, landing-page specification.
- Windows executable packaging guidance and release-process notes.
- Verify the full test suite passes in CI on every push.

## Under consideration (not committed)

- GUI spreadsheet preview before cleaning (preview of first rows added; a full
  interactive grid preview is a candidate).
- Batch processing from the CLI (multiple input files in one invocation).
- Dry-run mode (`--dry-run`) and exportable audit log (`--export-log`).
- Customizable duplicate detection via a `dedup_keys` config setting.
- Optional drag-and-drop in the GUI via `tkinterdnd2` (never a hard dependency).

## Not on the roadmap (by design)

- AI or machine-learning components in the cleaning workflow.
- Telemetry, cloud sync, user accounts, or any paid tier.
- Scans/PDFs, dashboards, or merging files with different layouts (out of scope).

## Versioning

Releases use semantic versioning. Breaking changes to cleaning behavior are avoided;
new flags and locale packs are additive. Check the CHANGELOG for each release.
