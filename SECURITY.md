# Security — CleanSheet

## Dependency policy

CleanSheet intentionally has **one runtime dependency**: `openpyxl` (for reading and
writing Excel files). Everything else is Python standard library (`argparse`, `csv`,
`json`, `re`, `unicodedata`, `tkinter` for the GUI). There are no transitive
dependencies to audit beyond openpyxl's own, and no API keys, secrets, or credentials
are used or stored.

The optional Windows `.exe` is built with PyInstaller (a build-time tool only; not a
runtime dependency of the Python package).

## Data handling

- The input file is never modified; output is written to a new file.
- CleanSheet does not send customer spreadsheets or any business data to external
  services, including AI services. The cleaning workflow is fully functional without
  any network access.
- Formulas are read as their last-saved cached values. Cells whose formula has no
  cached result read as empty and are flagged with a warning; macros (`.xlsm`) are not
  preserved by openpyxl and the tool warns rather than silently dropping them.

## Supported input types

`.xlsx`, `.xlsm` (macros not preserved), `.csv`, `.tsv`, `.txt`. CSV files are decoded
trying UTF-8 (with BOM), UTF-16 (BOM), Windows-1256, ISO-8859-6, then Windows-1252 as a
last resort, so Arabic CSVs exported from Windows are read correctly instead of being
mojibake'd.

## Reporting security concerns

If you believe you have found a security issue, open a GitHub issue (public) describing
the problem. Do not attach real client data; use a fictional sample that reproduces the
issue. There is no bug-bounty program and no security contact beyond the public issue
tracker.

MIT-licensed; provided "as is" without warranty. See LICENSE.
