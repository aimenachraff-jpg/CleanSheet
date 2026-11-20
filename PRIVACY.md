# Privacy — CleanSheet

CleanSheet is an **offline-first, local-only** spreadsheet cleaner. Nothing about your
data ever leaves your computer, and no account, cloud upload, or telemetry is required
to use it.

## What CleanSheet does with your data

- **Reads** the `.xlsx` or `.csv` file you point it at, on your machine only.
- **Writes** a cleaned copy (`*_cleaned.xlsx`), a markdown report, and (optionally) a CSV
  copy and an audit-log CSV next to your file — all written locally.
- **Never modifies** your original file. Every run writes to a new file; CleanSheet
  refuses to overwrite an existing output unless you explicitly pass `--force`.
- **Keeps values it cannot read.** Ambiguous or unreadable cells are flagged for your
  review (yellow/orange) and left as-is. Nothing is silently guessed, filled in, or
  discarded.

## What CleanSheet does NOT do

- No telemetry, analytics, or usage reporting of any kind.
- No cloud upload, sync, or remote processing.
- No account, sign-in, or user tracking.
- No AI model is consulted, and no customer spreadsheet is ever sent to an external
  service. CleanSheet is entirely rule-based (openpyxl + Python standard library).
- No temporary files are created that contain your data. Output is written directly to
  the location you choose.

## Sample data only in the repository

The sample files (`sample_messy_*.xlsx`) and all tests use fictional names, phones,
and orders. Do not paste a real client file into a GitHub issue; anonymize it first
(keep the shapes and formats, replace the values). See the "My file broke it" issue
template.

## Your responsibilities

CleanSheet gives you a change log and a Review sheet so you can verify every change
before using the cleaned file. Review flagged (yellow/orange) cells, especially
ambiguous dates, prices, and phone numbers, before sending the cleaned file to anyone
or relying on it for decisions.

MIT-licensed; provided "as is" without warranty. See LICENSE.
